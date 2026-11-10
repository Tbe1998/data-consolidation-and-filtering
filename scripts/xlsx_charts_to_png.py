#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 Excel 里的内嵌图表按指定图幅重绘成 PNG（需要 matplotlib）。

用途：Excel/Origin 里的曲线图要排进报告，但直接截图分辨率不够、比例不统一。
这里读 chart XML 里的缓存数据自行绘制，图幅、字体、图例位置、坐标轴范围可控。

图幅约定（工程报告常用，可改）：
    时程图  186 x 56 mm   →  A4 通栏一条（220 dpi 约 1610 x 484 px）
    测斜图   90 x 210 mm  →  A4 半栏通高（220 dpi 约 779 x 1818 px）

样式约定（2026-11 起）：
    - 所有曲线实线；9 色互不相近（Okabe-Ito 扩展）；每条曲线专属空心标记，
      黑白打印也能对上日期；
    - 纵轴范围按数据动态计算：1.5×极值起、向上取整到 nice 步长，落在 1.5~2.5 倍；
      有负值时下限同步外扩，绝不裁掉曲线；
    - 同组（同一工作表）曲线共用纵轴范围；量程不足组内最大跨度 40% 的图
      （如轴力后几张）保留自身范围，避免被压成直线——这是原模板既定规则。

用法：
    python3 scripts/xlsx_charts_to_png.py 曲线图.xlsx out/png
    python3 scripts/xlsx_charts_to_png.py 曲线图.xlsx out/png --dpi 300 \
        --time-sheets '图1*' --time-size 186x56 --incl-size 90x210
"""

import argparse
import datetime
import json
import fnmatch
import math
import os
import re
import zipfile

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt                      # noqa: E402
import matplotlib.dates as mdates                    # noqa: E402
from matplotlib.ticker import FuncFormatter          # noqa: E402
from matplotlib import font_manager as fm            # noqa: E402

# 9 种色相/明度都拉开的颜色（色盲友好，黑白打印靠标记形状区分）
COLORS = ["#0072B2", "#E69F00", "#009E73", "#CC79A7", "#56B4E9",
          "#D55E00", "#7B5EA7", "#8C564B", "#000000"]
MARKERS = ["o", "s", "^", "D", "v", "P", "X", "*", "h"]


# ---------------------------------------------------------------- 解析

def load_sheets(path):
    z = zipfile.ZipFile(path)
    wb = z.read("xl/workbook.xml").decode()
    rels = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"',
                           z.read("xl/_rels/workbook.xml.rels").decode()))
    sheets = re.findall(r'<sheet name="([^"]+)" sheetId="\d+"[^>]*r:id="(rId\d+)"', wb)
    out = []
    for name, rid in sheets:
        p = "xl/" + rels[rid].lstrip("/").replace("xl/", "")
        x = z.read(p).decode()
        d = re.findall(r'<drawing r:id="(rId\d+)"', x)
        charts = []
        if d:
            dr = "xl/worksheets/_rels/" + p.split("/")[-1] + ".rels"
            m = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', z.read(dr).decode()))
            tgt = "xl/" + m[d[0]].replace("../", "")
            dx = z.read(tgt).decode()
            cids = re.findall(r'<c:chart[^>]*r:id="(rId\d+)"', dx)
            m2 = dict(re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"',
                                 z.read(tgt.replace("/drawings/", "/drawings/_rels/") + ".rels").decode()))
            charts = ["xl/charts/" + m2[c].split("/")[-1] for c in cids]
        out.append((name, [z.read(c).decode() for c in charts]))
    return out


def cache_vals(block):
    d = {}
    for m in re.finditer(r'<c:pt idx="(\d+)"[^>]*>\s*<c:v>([^<]*)</c:v>', block):
        d[int(m.group(1))] = m.group(2)
    if not d:
        return []
    return [d.get(i) for i in range(max(d) + 1)]


def natkey(name):
    """测点自然排序：DL2 排在 DL10 前面。"""
    return [int(t) if t.isdigit() else str(t).lower()
            for t in re.split(r"(\d+)", str(name))]


def series_of(s):
    out = []
    for m in re.finditer(r"<c:ser>.*?</c:ser>", s, re.S):
        b = m.group(0)
        tx = re.search(r"<c:tx>.*?</c:tx>", b, re.S)
        name = re.findall(r"<c:v>([^<]*)</c:v>", tx.group(0))[-1] if tx else ""
        xb = re.search(r"<c:xVal>.*?</c:xVal>", b, re.S)
        yb = re.search(r"<c:yVal>.*?</c:yVal>", b, re.S)
        out.append(dict(name=name,
                        x=cache_vals(xb.group(0)) if xb else [],
                        y=cache_vals(yb.group(0)) if yb else []))
    return out


def axes_of(s):
    out = []
    for m in re.finditer(r"<c:valAx>.*?</c:valAx>", s, re.S):
        b = m.group(0)

        def g(pat):
            mm = re.search(pat, b)
            return mm.group(1) if mm else None

        title = re.search(r"<c:title>.*?</c:title>", b, re.S)
        out.append(dict(
            pos=g(r'<c:axPos val="(\w)"'),
            orient=g(r'<c:orientation val="(\w+)"'),
            mn=g(r'<c:min val="([-\d.]+)"'), mx=g(r'<c:max val="([-\d.]+)"'),
            mu=g(r'<c:majorUnit val="([\d.]+)"'),
            title=(re.findall(r"<a:t>([^<]*)</a:t>", title.group(0))[0] if title else None)))
    return out


def rect(s, tag="plotArea"):
    if tag == "plotArea":
        m = re.search(r"<c:plotArea>\s*<c:layout>\s*<c:manualLayout>(.*?)</c:manualLayout>",
                      s, re.S)
    else:
        m = re.search(r"<c:legend>.*?<c:manualLayout>(.*?)</c:manualLayout>", s, re.S)
    if not m:
        return None
    d = dict(re.findall(r'<c:(x|y|w|h) val="([\d.]+)"', m.group(1)))
    if len(d) < 4:
        return None
    return tuple(float(d[k]) for k in "xywh")


def legend_pos(s):
    m = re.search(r'<c:legend>.*?<c:legendPos val="(\w)"', s, re.S)
    return m.group(1) if m else "b"


def is_inclinometer(s, axes):
    if any(a["orient"] == "maxMin" for a in axes):
        return True
    return any((a["title"] or "").startswith("深度") for a in axes)


# ---------------------------------------------------------------- 轴范围

def nice_num(x):
    """把 x 收成 1/2/5×10^n 的“好看”数。"""
    if x <= 0:
        return 1.0
    exp = math.floor(math.log10(x))
    base = 10.0 ** exp
    f = x / base
    for m in (1, 2, 5, 10):
        if f <= m + 1e-12:
            return m * base
    return 10 * base


def data_range(vmin, vmax):
    """按 1.5×极值定上下限，步长取 nice 数，保证 1.5~2.5 倍且不裁曲线。"""
    hi_raw = 1.5 * vmax if vmax > 0 else 0.0
    lo_raw = -1.5 * abs(vmin) if vmin < 0 else 0.0
    step = nice_num((hi_raw - lo_raw) / 8.0)
    hi = math.ceil(hi_raw / step - 1e-9) * step if hi_raw else 0.0
    lo = math.floor(lo_raw / step + 1e-9) * step if lo_raw else 0.0
    if hi == lo:                    # 全 0 的图
        hi, lo = step, 0.0
    return lo, hi, step


def fmt_step(step):
    if step >= 1:
        return FuncFormatter(lambda v, _: "{:,.0f}".format(v))
    if step >= 0.1:
        return FuncFormatter(lambda v, _: "{:,.1f}".format(v))
    return FuncFormatter(lambda v, _: "{:,.2f}".format(v))


def frame(ax, grid="y"):
    for sp in ("top", "right"):
        ax.spines[sp].set_visible(False if grid == "y" else ax.spines[sp].get_visible())
    for sp in ("left", "bottom"):
        ax.spines[sp].set_color("#C9C9C9")
        ax.spines[sp].set_linewidth(0.8)
    ax.grid(axis=grid, color="#E9E9E9", linewidth=0.8)
    ax.set_axisbelow(True)
    ax.tick_params(length=0, labelsize=6.5, colors="#404040")


def setup_font(extra=None):
    """注册系统中文字体；找不到时图内中文会变方块，所以尽量多扫几个目录。"""
    cands = list(extra or []) + [
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf",
        "C:/Windows/Fonts/msyh.ttc",
        "C:/Windows/Fonts/simhei.ttf",
        "/System/Library/Fonts/PingFang.ttc",
    ]
    import glob
    for pat in (os.path.expanduser("~/.fonts/*.tt[fc]"),
                os.path.expanduser("~/.fonts/*.otf"),
                os.path.expanduser("~/.local/share/fonts/*"),
                "/usr/share/fonts/**/*CJK*",
                "/usr/share/fonts/**/*NotoSansSC*",
                "/usr/share/fonts/**/*wqy*"):
        cands.extend(glob.glob(pat, recursive=True))
    for p in cands:
        if os.path.exists(p) and "unifont" not in p.lower():
            try:
                fm.fontManager.addfont(p)
            except Exception:
                pass
    names = [f.name for f in fm.fontManager.ttflist
             if any(k in f.name for k in ("Zen Hei", "CJK", "YaHei",
                                          "Noto Sans SC", "Source Han", "SimHei"))]
    names = list(dict.fromkeys(names))
    plt.rcParams.update({"font.sans-serif": names + ["DejaVu Sans"],
                         "axes.unicode_minus": False, "figure.facecolor": "white"})
    return names


def mm(v):
    return v / 25.4


# ---------------------------------------------------------------- 绘图

def render(src, outdir, dpi=220, time_size=(186, 56), incl_size=(90, 210),
           time_sheets=None, only=None, debug=False, incl_subdir="incl"):
    """time_sheets=None 时按轴自动判别时程/测斜；给通配符则按工作表名强制指定。"""
    os.makedirs(outdir, exist_ok=True)
    incl_dir = os.path.join(outdir, incl_subdir) if incl_subdir else outdir
    os.makedirs(incl_dir, exist_ok=True)
    index = []
    groups_out = []
    for name, charts in load_sheets(src):
        if name == "说明" or not charts:
            continue
        parsed = []
        for s in charts:
            title = (re.findall(r"<a:t>([^<]*)</a:t>",
                                re.search(r"<c:title>.*?</c:title>", s, re.S).group(0))[0]
                     if re.search(r"<c:title>.*?</c:title>", s, re.S) else name)
            parsed.append(dict(raw=s, title=title, ser=series_of(s), axs=axes_of(s)))
        if time_sheets:
            is_time = any(fnmatch.fnmatch(name, p) for p in time_sheets)
        else:                                  # 自动：全组都是深度剖面才算测斜
            is_time = not all(is_inclinometer(c["raw"], c["axs"]) for c in parsed)

        # 时程图：先按测点前缀分类（DL/JS/…），类内自然排序（1、2、3…10），每图 ≤5 点
        counts = [len(c["ser"]) for c in parsed]
        if is_time and counts and min(counts) > 1:
            allser = [sp for c in parsed for sp in c["ser"]]
            tpl = parsed[0]
            base = re.sub(r"\s*\d+-\d+\s*$", "", tpl["title"]).strip()
            classes, order = {}, []
            for sp in allser:
                m = re.match(r"[A-Za-z]+", sp["name"])
                cls = m.group(0).upper() if m else ""
                if cls not in classes:
                    classes[cls] = []
                    order.append(cls)
                classes[cls].append(sp)
            newp = []
            for cls in order:
                sps = sorted(classes[cls], key=lambda sp: natkey(sp["name"]))
                for i in range(0, len(sps), 5):
                    chunk = sps[i:i + 5]
                    t = "%s %s~%s" % (base, chunk[0]["name"], chunk[-1]["name"])
                    newp.append(dict(raw=tpl["raw"], title=t, ser=chunk, axs=tpl["axs"]))
            parsed = newp

        # 同组纵轴统一：跨度 ≥ 组内最大跨度 40% 的图共用一套范围
        ymin = [min(float(v) for sp in c["ser"] for v in sp["y"]
                    if v not in (None, "")) for c in parsed if c["ser"]]
        ymax = [max(float(v) for sp in c["ser"] for v in sp["y"]
                    if v not in (None, "")) for c in parsed if c["ser"]]
        spans = [b - a for a, b in zip(ymin, ymax)] if ymin else [0]
        max_span = max(spans) if spans else 0
        if ymin:
            g_lo, g_hi, g_step = data_range(min(ymin), max(ymax))
        else:
            g_lo, g_hi, g_step = 0.0, 1.0, 0.2

        for k, c in enumerate(parsed, 1):
            if only and not re.search(only, c["title"]):
                continue
            ser, axs = c["ser"], c["axs"]
            incl = is_inclinometer(c["raw"], axs) or not is_time
            pr = rect(c["raw"], "plotArea")
            lg = rect(c["raw"], "legend")
            lp = legend_pos(c["raw"])
            w, h = incl_size if incl else time_size
            fig = plt.figure(figsize=(mm(w), mm(h)))
            if pr:
                pr = (pr[0], 1 - pr[1] - pr[3], pr[2], pr[3])
            if incl and pr and pr[0] < 0.10:
                pr = (0.10, pr[1], pr[0] + pr[2] - 0.10, pr[3])
            ax = fig.add_axes(pr if pr else (0.10, 0.24, 0.84, 0.56))

            if incl:
                _draw_incl(ax, fig, ser, axs, c["title"], pr, lg, lp)
            else:
                ys = [float(v) for sp in ser for v in sp["y"] if v not in (None, "")]
                own = (min(ys), max(ys)) if ys else (0, 1)
                span = own[1] - own[0]
                mx = max(abs(own[0]), abs(own[1])) or 1.0
                # 同组共用轴范围；但量程太小、或共用后曲线占比 <1/2.6 的图，
                # 保留自身范围（原模板的“量程保护”推广到新规则）
                if span >= 0.4 * max_span:
                    lo, hi, step = g_lo, g_hi, g_step
                else:
                    lo, hi, step = data_range(*own)
                if ys and (min(ys) < lo or max(ys) > hi):      # 兜底，绝不裁
                    lo, hi, step = data_range(min(min(ys), lo), max(max(ys), hi))
                    print("!! %s 轴范围不足，已扩展为 %g~%g" % (c["title"], lo, hi))
                if debug and ys:
                    mx = max(abs(min(ys)), abs(max(ys))) or 1
                    print("  %-42s data %8.2f~%8.2f  axis %8.2f~%8.2f  step %7.2f  ratio %.2f"
                          % (c["title"][-24:], min(ys), max(ys), lo, hi, step,
                             (hi - lo) / mx))
                _draw_time(ax, ser, axs, c["title"], lp, lo, hi, step)
                if ys:
                    assert min(ys) >= lo and max(ys) <= hi, c["title"] + " 有数据越界"

            safe = re.sub(r"[/\\]", "-", c["title"])
            if len(parsed) > 1 and "~" not in safe:
                safe += " %d-%d" % (k, len(parsed))
            if incl and "全过程" not in safe:
                safe += " 全过程"
            target = incl_dir if incl else outdir
            fn = safe + ".png"
            fig.savefig(os.path.join(target, fn), dpi=dpi, facecolor="white")
            plt.close(fig)
            rel = os.path.join(incl_subdir, fn) if incl and incl_subdir else fn
            index.append((len(index) + 1, name, c["title"], rel, w, h))
            if is_time:
                base_d = datetime.date(1899, 12, 30)
                xs = [sp["x"] for sp in ser if sp["x"]]
                dates = [str(base_d + datetime.timedelta(days=int(float(v)))) for v in xs[0]] if xs else []
                groups_out.append(dict(sheet=name, title=c["title"], file=fn,
                                       points=[sp["name"] for sp in ser],
                                       dates=dates,
                                       series={sp["name"]: sp["y"] for sp in ser}))

    with open(os.path.join(outdir, "_index.txt"), "w", encoding="utf-8") as fh:
        fh.write("序号\t工作表\t图名\t文件\t图幅mm\n")
        for i, sh, t, fn, w, h in index:
            fh.write("%02d\t%s\t%s\t%s\t%dx%d\n" % (i, sh, t, fn, w, h))
    with open(os.path.join(outdir, "_groups.json"), "w", encoding="utf-8") as fh:
        json.dump(groups_out, fh, ensure_ascii=False)
    return index


def _draw_time(ax, ser, axs, title, lp, lo, hi, step):
    base = datetime.date(1899, 12, 30)
    allx = []
    for i, sp in enumerate(ser):
        xs, ys = [], []
        for xv, yv in zip(sp["x"], sp["y"]):
            if xv in (None, "") or yv in (None, ""):
                continue
            xs.append(base + datetime.timedelta(days=float(xv)))
            ys.append(float(yv))
        if not xs:
            continue
        allx += xs
        col = COLORS[i % len(COLORS)]
        ax.plot(xs, ys, color=col, lw=1.5, ls="-",
                marker=MARKERS[i % len(MARKERS)], ms=3.8,
                mfc="white", mec=col, mew=0.8, label=sp["name"], zorder=3,
                markevery=max(1, len(xs) // 8))
    if not allx:
        return
    if lo < 0 < hi:
        ax.axhline(0, color="#C9C9C9", lw=0.8, zorder=1)
    frame(ax, "y")
    ax.set_ylim(lo, hi)
    ax.set_yticks([lo + i * step for i in range(int(round((hi - lo) / step)) + 1)])
    ax.yaxis.set_major_formatter(fmt_step(step))
    ax.set_xlim(min(allx), max(allx))
    span = (max(allx) - min(allx)).days
    ax.xaxis.set_major_locator(mdates.AutoDateLocator(
        minticks=3, maxticks=7, interval_multiples=True))
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter("%Y/%m" if span > 300 else "%Y/%m/%d"))
    ax.set_ylabel((axs[1]["title"] if len(axs) > 1 else None) or "累计变化量", fontsize=7,
                  color="#595959", labelpad=6)
    ncol = min(5, max(1, len(ser)))
    if lp == "r":
        ax.legend(loc="center left", bbox_to_anchor=(1.01, 0.5), frameon=False,
                  fontsize=6.3, handlelength=2.0, labelspacing=0.6,
                  handletextpad=0.5, borderaxespad=0)
    else:
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.10), ncol=ncol,
                  frameon=False, fontsize=6.3, handlelength=2.2,
                  columnspacing=1.2, handletextpad=0.6, borderaxespad=0)
    ax.set_title(title, loc="center", fontsize=8.6, fontweight="bold",
                 color="#1F1F1F", pad=6)


def _draw_incl(ax, fig, ser, axs, title, pr, lg, lp):
    for i, sp in enumerate(ser):
        xs, ys = [], []
        for xv, yv in zip(sp["x"], sp["y"]):
            if xv in (None, "") or yv in (None, ""):
                continue
            xs.append(float(xv))
            ys.append(float(yv))
        if not xs:
            continue
        col = COLORS[i % len(COLORS)]
        ax.plot(xs, ys, color=col, lw=1.4, ls="-",
                marker=MARKERS[i % len(MARKERS)], ms=3.2,
                mfc="white", mec=col, mew=0.7, label=sp["name"], zorder=3,
                markevery=max(1, len(xs) // 24))
    frame(ax, "both")
    ax.tick_params(labelsize=6.0, pad=2)
    if len(axs) >= 2 and axs[0]["mn"] and axs[0]["mx"]:
        xa, ya = axs[0], axs[1]
        xlo, xhi = float(xa["mn"]), float(xa["mx"])
        xst = float(xa["mu"]) if xa["mu"] else (xhi - xlo) / 5
        ax.set_xlim(xlo, xhi)
        ax.set_xticks([xlo + i * xst for i in range(int(round((xhi - xlo) / xst)) + 1)])
        ax.xaxis.set_major_formatter(fmt_step(xst))
        dlo, dhi = float(ya["mn"]), float(ya["mx"])
        dst = float(ya["mu"]) if ya["mu"] else 2
        ax.set_ylim(dhi, dlo)
        ax.set_yticks([dlo + i * dst for i in range(int(round((dhi - dlo) / dst)) + 1)])
        ax.yaxis.set_major_formatter(fmt_step(dst))
        ax.set_xlabel(xa["title"] or "位移（mm）", fontsize=6.8, color="#595959", labelpad=3)
        fig.text(0.028, (pr[1] + pr[3] / 2) if pr else 0.5,
                 ya["title"] or "深度（m）", rotation=90, va="center", ha="center",
                 fontsize=6.4, color="#595959")
    ax.xaxis.tick_top()
    ax.xaxis.set_label_position("top")
    if lp == "r" and lg:
        ax.legend(loc="lower right", bbox_to_anchor=(1.0, 0.0), ncol=1, frameon=True,
                  facecolor="white", edgecolor="none", fontsize=6.0,
                  handlelength=1.6, labelspacing=0.4, handletextpad=0.4,
                  borderaxespad=0.6, framealpha=1.0)
    else:
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.06), ncol=2,
                  frameon=False, fontsize=6.0, handlelength=1.6,
                  labelspacing=0.35, columnspacing=0.9,
                  handletextpad=0.4, borderaxespad=0)
    ax.set_title(title, loc="left", fontsize=7.6, fontweight="bold",
                 color="#1F1F1F", pad=14)


def _size(t):
    w, h = t.lower().split("x")
    return (float(w), float(h))


def main(argv=None):
    ap = argparse.ArgumentParser(description="Excel 内嵌图表 → PNG")
    ap.add_argument("src")
    ap.add_argument("outdir")
    ap.add_argument("--dpi", type=int, default=220)
    ap.add_argument("--time-size", default="186x56", help="时程图图幅 mm，默认 186x56")
    ap.add_argument("--incl-size", default="90x210", help="测斜图图幅 mm，默认 90x210")
    ap.add_argument("--time-sheets", default="auto",
                    help="按工作表名通配判断时程图（逗号分隔）；auto=按坐标轴自动判别（推荐）")
    ap.add_argument("--incl-subdir", default="incl",
                    help="测斜 PNG 子目录名（相对 outdir，默认 incl）；传空串则与平铺")
    ap.add_argument("--only", default=None, help="只导出图名匹配该正则的图")
    ap.add_argument("--font", action="append", default=[], help="追加中文字体路径")
    ap.add_argument("--debug", action="store_true", help="打印每张图的轴范围与倍数")
    a = ap.parse_args(argv)

    import logging
    logging.getLogger("matplotlib.font_manager").setLevel(logging.ERROR)
    fonts = setup_font(a.font)
    pats = None if a.time_sheets.strip().lower() == "auto" else \
        tuple(p.strip() for p in a.time_sheets.split(",") if p.strip())
    idx = render(a.src, a.outdir, a.dpi, _size(a.time_size), _size(a.incl_size),
                 pats, a.only, a.debug, a.incl_subdir)
    n_incl = sum(1 for r in idx if "全过程" in r[3])
    n_time = len(idx) - n_incl
    print("导出 %d 张 → %s（时程 %d / 测斜 %d，测斜在 %s/）"
          % (len(idx), a.outdir, n_time, n_incl, a.incl_subdir or "."))
    print("中文字体: %s" % ("、".join(fonts[:3]) if fonts else "未找到（图内中文可能成方块）"))
    for i, sh, t, fn, w, h in idx[:5]:
        print("  %02d %s (%gx%gmm)" % (i, fn, w, h))
    gj = os.path.join(a.outdir, "_groups.json")
    try:
        ng = len(json.load(open(gj, encoding="utf-8")))
    except Exception:
        ng = 0
    if ng == 0 and n_incl:
        print("!! 只有测斜图、没有时程图：可以继续 make_delivery（筛选表只含测斜块）")
        return 0
    if ng == 0:
        print("!! 没有识别出任何时程图（_groups.json 为空）；"
              "若确有测点-日期曲线，用 --time-sheets '图1*' 指定工作表")
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
