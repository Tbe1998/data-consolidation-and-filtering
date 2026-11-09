#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 Excel 里的内嵌图表按指定图幅重绘成 PNG（需要 matplotlib）。

用途：Excel/Origin 里的曲线图要排进报告，但直接截图分辨率不够、比例不统一。
这里读 chart XML 里的缓存数据自行绘制，图幅、字体、图例位置可控。

图幅约定（工程报告常用，可改）：
    时程图  186 x 56 mm   →  A4 通栏一条（220 dpi 约 1610 x 484 px）
    测斜图   90 x 210 mm  →  A4 半栏通高（220 dpi 约 779 x 1818 px）

用法：
    python3 scripts/xlsx_charts_to_png.py 曲线图.xlsx out/png
    python3 scripts/xlsx_charts_to_png.py 曲线图.xlsx out/png --dpi 300 \
        --time-sheets '图1*' --time-size 186x56 --incl-size 90x210
"""

import argparse
import datetime
import fnmatch
import os
import re
import zipfile

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt                      # noqa: E402
import matplotlib.dates as mdates                    # noqa: E402
from matplotlib.ticker import FuncFormatter          # noqa: E402
from matplotlib import font_manager as fm            # noqa: E402

MK = {"circle": "o", "square": "s", "triangle": "^", "diamond": "D",
      "x": "x", "plus": "P", "star": "*", "dash": "_"}


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


def series_of(s):
    out = []
    for m in re.finditer(r"<c:ser>.*?</c:ser>", s, re.S):
        b = m.group(0)
        tx = re.search(r"<c:tx>.*?</c:tx>", b, re.S)
        name = re.findall(r"<c:v>([^<]*)</c:v>", tx.group(0))[-1] if tx else ""
        xb = re.search(r"<c:xVal>.*?</c:xVal>", b, re.S)
        yb = re.search(r"<c:yVal>.*?</c:yVal>", b, re.S)
        col = re.search(r'<a:srgbClr val="([0-9A-Fa-f]{6})"/>', b)
        mk = re.search(r'<c:symbol val="(\w+)"/>', b)
        out.append(dict(
            name=name,
            x=cache_vals(xb.group(0)) if xb else [],
            y=cache_vals(yb.group(0)) if yb else [],
            color="#" + col.group(1) if col else "#2E5C8A",
            marker=mk.group(1) if mk else "circle"))
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
    """测斜图的判据：深度轴 maxMin（深度向下增）或轴位置在左且纵轴标题含深度。"""
    if any(a["orient"] == "maxMin" for a in axes):
        return True
    return any((a["title"] or "").startswith("深度") for a in axes)


# ---------------------------------------------------------------- 绘图

def fnum(step):
    if step >= 1:
        return FuncFormatter(lambda v, _: "{:,.0f}".format(v))
    if step >= 0.1:
        return FuncFormatter(lambda v, _: "{:,.1f}".format(v))
    return FuncFormatter(lambda v, _: "{:,.2f}".format(v))


def ticks(lo, hi, st, cap=26):
    n = int(round((hi - lo) / st)) if st else 8
    n = max(1, min(n, cap))
    return [round(lo + i * (hi - lo) / n, 6) for i in range(n + 1)]


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
    cands = list(extra or []) + [
        "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "C:/Windows/Fonts/msyh.ttc",
    ]
    for p in cands:
        if os.path.exists(p):
            try:
                fm.fontManager.addfont(p)
            except Exception:
                pass
    names = [f.name for f in fm.fontManager.ttflist
             if "Zen Hei" in f.name or "CJK" in f.name or "YaHei" in f.name]
    plt.rcParams.update({"font.sans-serif": names + ["DejaVu Sans"],
                         "axes.unicode_minus": False, "figure.facecolor": "white"})


def mm(v):
    return v / 25.4


def render(src, outdir, dpi=220, time_size=(186, 56), incl_size=(90, 210),
           time_sheets=("图1*",), only=None):
    os.makedirs(outdir, exist_ok=True)
    index = []
    for name, charts in load_sheets(src):
        if name == "说明" or not charts:
            continue
        for k, s in enumerate(charts, 1):
            title = (re.findall(r"<a:t>([^<]*)</a:t>",
                                re.search(r"<c:title>.*?</c:title>", s, re.S).group(0))[0]
                     if re.search(r"<c:title>.*?</c:title>", s, re.S) else name)
            if only and not re.search(only, title):
                continue
            ser = series_of(s)
            axs = axes_of(s)
            incl = is_inclinometer(s, axs) or not any(
                fnmatch.fnmatch(name, p) for p in time_sheets)
            pr = rect(s, "plotArea")
            lg = rect(s, "legend")
            lp = legend_pos(s)
            w, h = incl_size if incl else time_size
            fig = plt.figure(figsize=(mm(w), mm(h)))

            if pr:
                pr = (pr[0], 1 - pr[1] - pr[3], pr[2], pr[3])   # y 轴翻到 matplotlib 坐标
            if incl and pr and pr[0] < 0.10:                     # 左侧留给深度轴标题
                pr = (0.10, pr[1], pr[0] + pr[2] - 0.10, pr[3])
            ax = fig.add_axes(pr if pr else (0.10, 0.24, 0.84, 0.56))

            if incl:
                _draw_incl(ax, fig, ser, axs, title, pr, lg, lp)
            else:
                _draw_time(ax, ser, axs, title, lp)

            safe = re.sub(r"[/\\]", "-", title)
            fn = "{}{}.png".format(safe, " %d-%d" % (k, len(charts)) if len(charts) > 1 else "")
            fig.savefig(os.path.join(outdir, fn), dpi=dpi, facecolor="white")
            plt.close(fig)
            index.append((len(index) + 1, name, title, fn, w, h))

    with open(os.path.join(outdir, "_index.txt"), "w", encoding="utf-8") as fh:
        fh.write("序号\t工作表\t图名\t文件\t图幅mm\n")
        for i, sh, t, fn, w, h in index:
            fh.write("%02d\t%s\t%s\t%s\t%dx%d\n" % (i, sh, t, fn, w, h))
    return index


def _draw_time(ax, ser, axs, title, lp):
    base = datetime.date(1899, 12, 30)
    allx = []
    for sp in ser:
        xs, ys = [], []
        for xv, yv in zip(sp["x"], sp["y"]):
            if xv in (None, "") or yv in (None, ""):
                continue
            xs.append(base + datetime.timedelta(days=float(xv)))
            ys.append(float(yv))
        if not xs:
            continue
        allx += xs
        ax.plot(xs, ys, color=sp["color"], lw=1.5,
                marker=MK.get(sp["marker"], "o"), ms=3.2, mfc=sp["color"],
                mec="white", mew=0.6, label=sp["name"], zorder=3,
                markevery=max(1, len(xs) // 10))
    if not allx:
        return
    ax.axhline(0, color="#C9C9C9", lw=0.8, zorder=1)
    frame(ax, "y")
    ya = axs[1] if len(axs) > 1 else axs[0]
    if ya["mn"] and ya["mx"]:
        lo, hi = float(ya["mn"]), float(ya["mx"])
        st = float(ya["mu"]) if ya["mu"] else (hi - lo) / 5
        ax.set_ylim(lo, hi)
        ax.set_yticks(ticks(lo, hi, st))
        ax.yaxis.set_major_formatter(fnum(st))
    ax.set_xlim(min(allx), max(allx))
    span = (max(allx) - min(allx)).days
    ax.xaxis.set_major_locator(mdates.AutoDateLocator(
        minticks=3, maxticks=7, interval_multiples=True))
    ax.xaxis.set_major_formatter(
        mdates.DateFormatter("%Y/%m" if span > 300 else "%Y/%m/%d"))
    ax.set_ylabel(ya["title"] or "累计变化量（mm）", fontsize=7, color="#595959", labelpad=6)
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
    for sp in ser:
        xs, ys = [], []
        for xv, yv in zip(sp["x"], sp["y"]):
            if xv in (None, "") or yv in (None, ""):
                continue
            xs.append(float(xv))
            ys.append(float(yv))
        if not xs:
            continue
        ax.plot(xs, ys, color=sp["color"], lw=1.4, marker="o", ms=2.6,
                mfc=sp["color"], mec="white", mew=0.4, label=sp["name"], zorder=3,
                markevery=max(1, len(xs) // 24))
    frame(ax, "both")
    ax.tick_params(labelsize=6.0, pad=2)
    if len(axs) >= 2 and axs[0]["mn"] and axs[0]["mx"]:
        xa, ya = axs[0], axs[1]
        xlo, xhi = float(xa["mn"]), float(xa["mx"])
        xst = float(xa["mu"]) if xa["mu"] else (xhi - xlo) / 5
        ax.set_xlim(xlo, xhi)
        ax.set_xticks(ticks(xlo, xhi, xst))
        ax.xaxis.set_major_formatter(fnum(xst))
        dlo, dhi = float(ya["mn"]), float(ya["mx"])
        dst = float(ya["mu"]) if ya["mu"] else 2
        ax.set_ylim(dhi, dlo)                       # 深度向下增大
        ax.set_yticks(ticks(dlo, dhi, dst))
        ax.yaxis.set_major_formatter(fnum(dst))
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
    ap.add_argument("--time-sheets", default="图1*", help="按工作表名通配判断时程图，可逗号分隔")
    ap.add_argument("--only", default=None, help="只导出图名匹配该正则的图")
    ap.add_argument("--font", action="append", default=[], help="追加中文字体路径")
    a = ap.parse_args(argv)

    setup_font(a.font)
    pats = [p.strip() for p in a.time_sheets.split(",") if p.strip()]
    idx = render(a.src, a.outdir, a.dpi, _size(a.time_size), _size(a.incl_size),
                 tuple(pats), a.only)
    print("导出 %d 张 → %s" % (len(idx), a.outdir))
    for i, sh, t, fn, w, h in idx[:5]:
        print("  %02d %s (%gx%gmm)" % (i, fn, w, h))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
