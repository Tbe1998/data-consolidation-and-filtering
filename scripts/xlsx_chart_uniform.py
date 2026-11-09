#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""批量统一 Excel 内嵌图表的版式（零第三方依赖）。

原理：.xlsx 是 zip，图表是 xl/charts/chartN.xml。选定每组里用户已经调好的一张
「基准图」，把它的绘图区几何、图例块、坐标轴刻度和标题格式复制到同组其余图表，
数据（numRef / numCache）一个字节都不动。

用法：
    # 1) 看结构，生成配置草稿
    python3 scripts/xlsx_chart_uniform.py --init 曲线图.xlsx > chart.json

    # 2) 改 chart.json 的 groups（哪几张表一组、谁是基准、要不要改图名）

    # 3) 套版式（原文件不动）
    python3 scripts/xlsx_chart_uniform.py 曲线图.xlsx 曲线图_统一版式.xlsx --config chart.json
"""

import argparse
import datetime
import json
import posixpath
import re
import sys
import zipfile

UNIT_RE = r"[（(]\s*(mm|cm|m|kN|kPa|MPa|kN·m|‰|%)\s*[)）]\s*$"
UNIT_INLINE_RE = r"[（(]\s*(mm|cm|m|kN|kPa|MPa|kN·m|‰|%)\s*[)）]"
INDEX_RUN_RE = r"<a:r>(?:(?!</a:r>).)*?<a:t>\s*\d+\s*/\s*\d+\s*</a:t></a:r>"
INDEX_IN_TEXT_RE = r"\s+\d+\s*/\s*\d+\s*(?=\[)"


# ---------------------------------------------------------------- zip / rels

def _target(base, tgt):
    if tgt.startswith("/"):
        return tgt.lstrip("/")
    return posixpath.normpath(posixpath.join(base, tgt))


def load(path):
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n) for n in z.namelist() if not n.endswith("/")}


def rels_of(xml, base):
    """{rId: 'xl/...'}；base 是 rels 所属目录，如 'xl/worksheets'。"""
    out = {}
    for rid, tgt in re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', xml):
        out[rid] = _target(base, tgt)
    return out


def sheet_paths(xml):
    """{sheet_name: 'xl/worksheets/sheetN.xml'}，按工作簿顺序。"""
    wb = xml["xl/workbook.xml"].decode("utf-8")
    rels = rels_of(xml["xl/_rels/workbook.xml.rels"].decode("utf-8"), "xl")
    out = {}
    for name, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
        p = rels.get(rid)
        if p and p in xml:
            out[name] = p
    return out


def sheets_of(xml):
    """[(sheet_name, chart_path...)]，按工作簿顺序。"""
    wb = xml["xl/workbook.xml"].decode("utf-8")
    rels = rels_of(xml["xl/_rels/workbook.xml.rels"].decode("utf-8"), "xl")
    out = []
    for name, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"', wb):
        sheet_path = rels.get(rid)
        if not sheet_path or sheet_path not in xml:
            out.append((name, []))
            continue
        sx = xml[sheet_path].decode("utf-8")
        d = re.search(r'<drawing r:id="(rId\d+)"', sx)
        charts = []
        if d:
            drels_path = posixpath.join(posixpath.dirname(sheet_path), "_rels",
                                        posixpath.basename(sheet_path) + ".rels")
            drels = rels_of(xml[drels_path].decode("utf-8"), posixpath.dirname(sheet_path)) \
                if drels_path in xml else {}
            drawing = drels.get(d.group(1))
            if drawing and drawing in xml:
                dx = xml[drawing].decode("utf-8")
                cids = re.findall(r'<c:chart[^>]*r:id="(rId\d+)"', dx)
                crp = posixpath.join(posixpath.dirname(drawing), "_rels",
                                     posixpath.basename(drawing) + ".rels")
                crels = rels_of(xml[crp].decode("utf-8"), posixpath.dirname(drawing)) \
                    if crp in xml else {}
                charts = [crels[c] for c in cids if c in crels]
        out.append((name, charts))
    return out


# ---------------------------------------------------------------- xml 小工具

# ---------------------------------------------------------------- 表头日期

def _col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def _serial_to_iso(v):
    try:
        n = float(v)
    except (TypeError, ValueError):
        return None
    if not (20000 <= n <= 80000):           # 不在 1954~2119 年之间，多半不是日期
        return None
    return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(n))).isoformat()


def _row_values(xml, sheet_name, row, from_col, count):
    """读某工作表某一行的单元格值（零依赖：直接解 sheet XML + sharedStrings）。"""
    p = sheet_paths(xml).get(sheet_name)
    if not p:
        return []
    sx = xml[p].decode("utf-8")
    shared = re.findall(r"<si>(.*?)</si>", xml.get("xl/sharedStrings.xml", b"").decode("utf-8"), re.S)
    shared = ["".join(re.findall(r"<t[^>]*>([^<]*)</t>", si)) for si in shared]
    cells = {}
    for m in re.finditer(r'<c r="([A-Z]+)%d"((?:\s[^>]*)?)(?:/>|>(.*?)</c>)' % row, sx, re.S):
        cells[_col_index(m.group(1))] = (m.group(2) or "", m.group(3) or "")
    out = []
    for c in range(from_col, from_col + count):
        attrs, body = cells.get(c, ("", ""))
        v = re.search(r"<v>([^<]*)</v>", body)
        if 't="s"' in attrs and v:
            val = shared[int(v.group(1))] if int(v.group(1)) < len(shared) else None
        elif 't="inlineStr"' in attrs:
            val = "".join(re.findall(r"<t[^>]*>([^<]*)</t>", body))
        elif v:
            val = _serial_to_iso(v.group(1)) or v.group(1)
        else:
            val = None
        if val:
            val = str(val).strip()
            for fmt in ("%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日", "%Y-%m-%d"):
                try:
                    val = datetime.datetime.strptime(val, fmt).date().isoformat()
                    break
                except ValueError:
                    pass
        out.append(val)
    return out


def _apply_legend_dates(s, sheet_name, vals, row=2, from_col=2):
    """把系列名从序列号换成表头日期，并改成引用表头单元格（改表头即可自动同步）。"""
    vals = list(vals)
    blocks = list(re.finditer(r"<c:tx><c:strRef>(.*?)</c:strRef></c:tx>", s, re.S))
    out, last = [], 0
    for k, b in enumerate(blocks):
        if k >= len(vals) or not vals[k]:
            break
        inner = b.group(1)
        col = _col_letter(from_col + k)
        inner = re.sub(r"<c:f>.*?</c:f>",
                       "<c:f>'%s'!$%s$%d</c:f>" % (sheet_name.replace("'", "''"),
                                                   col, row), inner, flags=re.S)
        if re.search(r"<c:v>[^<]*</c:v>", inner):
            inner = re.sub(r"<c:v>[^<]*</c:v>",
                           "<c:v>%s</c:v>" % escape(vals[k]), inner, count=1)
        out.append(s[last:b.start()])
        out.append("<c:tx><c:strRef>" + inner + "</c:strRef></c:tx>")
        last = b.end()
    out.append(s[last:])
    return "".join(out)


def block(s, tag):
    m = re.search(r"<c:%s(?:\s[^>]*)?>.*?</c:%s>" % (tag, tag), s, re.S)
    return None if not m else (m.start(), m.end(), m.group(0))


def axes(s):
    return [(m.start(), m.end(), m.group(0))
            for m in re.finditer(r"<c:valAx>.*?</c:valAx>", s, re.S)]


def title_text(s):
    b = block(s, "title")
    if not b:
        return None
    t = re.findall(r"<a:t>([^<]*)</a:t>", b[2])
    return "".join(t).strip() if t else None


def set_title_text(s, text):
    """保留基准标题的字体/字号/位置，只换文字。多 run 的标题合并成一个 run。"""
    b = block(s, "title")
    if not b:
        return s
    blk = b[2]
    runs = list(re.finditer(r"<a:r>(.*?)</a:r>", blk, re.S))
    if not runs:
        return s
    first = runs[0]
    t_run = re.search(r"<a:t>[^<]*</a:t>", first.group(0))
    if not t_run:
        return s
    blk2 = blk[:first.start()] + first.group(0).replace(
        t_run.group(0), "<a:t>%s</a:t>" % escape(text), 1) + blk[first.end():]
    for r in reversed(runs[1:]):
        blk2 = blk2[:r.start()] + blk2[r.end():]
    return s[:b[0]] + blk2 + s[b[1]:]


def escape(t):
    return (t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def clean_title_text(t, cfg, renames=None):
    if t is None:
        return None
    for a, b in (renames or {}).items():
        t = t.replace(a, b)
    if cfg.get("strip_index", True):
        t = re.sub(r"\s*\d+\s*/\s*\d+\s*$", "", t).strip()
    if cfg.get("strip_units", True):
        t = re.sub(UNIT_RE, "", t.strip()).strip()
    return t


def drop_index_runs(s, cfg):
    if not cfg.get("strip_index", True):
        return s
    b = block(s, "title")
    if not b:
        return s
    new = re.sub(INDEX_RUN_RE, "", b[2])
    return s[:b[0]] + new + s[b[1]:]


def series_values(s):
    """[(series_index, x_min, x_max, y_min, y_max)]，只读缓存值。"""
    out = []
    for i, m in enumerate(re.finditer(r"<c:ser>.*?</c:ser>", s, re.S)):
        b = m.group(0)

        def rng(tag):
            v = []
            for mm in re.finditer(r"<c:%s>.*?</c:%s>" % (tag, tag), b, re.S):
                v += [float(x) for x in re.findall(r"<c:v>(-?[\d.eE+]+)</c:v>", mm.group(0))]
            return (min(v), max(v)) if v else (None, None)

        xlo, xhi = rng("xVal")
        ylo, yhi = rng("yVal")
        out.append((i, xlo, xhi, ylo, yhi))
    return out


def axis_range(s, ax_pos):
    """该轴对应的数据范围：x 轴看 xVal，y 轴看 yVal。"""
    lo_i, hi_i = (1, 2) if ax_pos in ("b", "t", "x") else (3, 4)
    rows = series_values(s)
    los = [r[lo_i] for r in rows if r[lo_i] is not None]
    his = [r[hi_i] for r in rows if r[hi_i] is not None]
    if not los or not his:
        return (None, None)
    return (min(los), max(his))


# ---------------------------------------------------------------- 核心

def uniform(src, out, cfg):
    xml = load(src)
    structure = sheets_of(xml)
    sheet_index = {}
    for name, charts in structure:
        sheet_index[name] = charts
    owner = {}
    for name, charts in structure:
        for p in charts:
            owner[p] = name

    report = {"groups": [], "untouched": [], "missing": []}
    for gi, g in enumerate(cfg.get("groups", [])):
        names = g.get("sheets") or []
        targets = []
        for n in names:
            if n not in sheet_index:
                report["missing"].append(n)
                continue
            targets += sheet_index[n]
        if not targets:
            continue
        ref_i = int(g.get("ref", 0))
        if ref_i >= len(targets):
            ref_i = 0
        ref_path = targets[ref_i]
        ref_s = xml[ref_path].decode("utf-8")
        ref_axes = axes(ref_s)
        ref_leg = block(ref_s, "legend")
        ref_layout = re.search(r"<c:plotArea>\s*(?:<c:layout>.*?</c:layout>|<c:layout/>)",
                               ref_s, re.S)
        ref_title = block(ref_s, "title")
        renames = g.get("rename") or {}
        ratio = float(g.get("axis_min_span_ratio", cfg.get("axis_min_span_ratio", 0.4)))
        kept = []

        for ti, path in enumerate(targets):
            s = xml[path].decode("utf-8")
            ld = g.get("legend_dates")
            if ld:
                vals = _row_values(xml, owner.get(path, ""), int(ld.get("row", 2)),
                                   int(ld.get("from_col", 2)), int(ld.get("count", 6)))
                if vals:
                    s = _apply_legend_dates(s, owner.get(path, ""), vals,
                                            int(ld.get("row", 2)), int(ld.get("from_col", 2)))
            if ti == ref_i:
                s = drop_index_runs(s, cfg)
                new_t = clean_title_text(title_text(s), cfg, renames)
                if new_t:
                    s = set_title_text(s, new_t)
                xml[path] = s.encode("utf-8")
                continue

            # 1) 绘图区几何
            if ref_layout:
                tl = re.search(r"<c:plotArea>\s*(?:<c:layout>.*?</c:layout>|<c:layout/>)",
                               s, re.S)
                if tl:
                    s = s[:tl.start()] + ref_layout.group(0) + s[tl.end():]

            # 2) 图例（去掉 legendEntry，避免把基准图的系列名带过来）
            if ref_leg:
                leg = re.sub(r"<c:legendEntry>.*?</c:legendEntry>", "", ref_leg[2], flags=re.S)
                tb = block(s, "legend")
                if tb:
                    s = s[:tb[0]] + leg + s[tb[1]:]

            # 3) 坐标轴
            for k in range(len(ref_axes)):
                tgt_axes = axes(s)          # 每替换一个轴，后面的偏移量都会变，必须重取
                if k >= len(tgt_axes):
                    break
                ts, te, tblk = tgt_axes[k]
                rblk = ref_axes[k][2]
                newax = rblk

                # 轴 ID 必须用本图的，否则 Excel 认为轴和系列断链
                rid = re.search(r'<c:axId val="(\d+)"/>', tblk)
                cid = re.search(r'<c:crossAx val="(\d+)"/>', tblk)
                nid = re.search(r'<c:axId val="(\d+)"/>', newax)
                ncid = re.search(r'<c:crossAx val="(\d+)"/>', newax)
                if rid and nid:
                    newax = newax[:nid.start(1)] + rid.group(1) + newax[nid.end(1):]
                if cid and ncid:
                    newax = newax[:ncid.start(1)] + cid.group(1) + newax[ncid.end(1):]

                # 轴标题文字用本图的，格式用基准的
                old_tt = re.findall(r"<a:t>([^<]*)</a:t>", tblk)
                if old_tt:
                    def _sub(m, _t=old_tt[0]):
                        return "<a:t>%s</a:t>" % _t
                    newax = re.sub(r"<a:t>[^<]*</a:t>", _sub, newax, count=1)

                # 刻度：基准步长 + 本图数据外扩；量级差太多就保留本图原刻度
                mn = re.search(r'<c:min val="([-\d.]+)"', newax)
                mx = re.search(r'<c:max val="([-\d.]+)"', newax)
                mu = re.search(r'<c:majorUnit val="([\d.]+)"', newax)
                omn = re.search(r'<c:min val="([-\d.]+)"', tblk)
                omx = re.search(r'<c:max val="([-\d.]+)"', tblk)
                keep_scale = False
                if mn and mx and omn and omx:
                    ref_span = float(mx.group(1)) - float(mn.group(1))
                    old_span = float(omx.group(1)) - float(omn.group(1))
                    if ref_span > 0 and old_span < ratio * ref_span:
                        keep_scale = True
                if mn and mx and not keep_scale:
                    pos = re.search(r'<c:axPos val="(\w)"/>', tblk)
                    which = "x" if (pos and pos.group(1) in "bt") else "y"
                    dlo, dhi = axis_range(s, which)
                    if dlo is not None:
                        lo, hi = float(mn.group(1)), float(mx.group(1))
                        step = float(mu.group(1)) if mu else (hi - lo) / 5
                        while dlo < lo:
                            lo -= step
                        while dhi > hi:
                            hi += step
                        newax = re.sub(r'<c:min val="[-\d.]+"',
                                       '<c:min val="%g"' % lo, newax, count=1)
                        newax = re.sub(r'<c:max val="[-\d.]+"',
                                       '<c:max val="%g"' % hi, newax, count=1)
                if keep_scale:
                    kept.append(path)
                else:
                    s = s[:ts] + newax + s[te:]

            # 4) 图名：格式抄基准，文字用本图的（去单位、去 n/m）
            s = drop_index_runs(s, cfg)
            new_t = clean_title_text(title_text(s), cfg, renames)
            if ref_title and new_t:
                if not block(s, "title"):
                    s += ""  # 没有标题就保持没有
                else:
                    s = set_title_text(s, new_t)
            xml[path] = s.encode("utf-8")

        report["groups"].append({
            "name": g.get("name", "group%d" % gi),
            "sheets": names,
            "charts": len(targets),
            "ref": ref_path,
            "kept_scale": kept,
        })

    # 5) 说明页 / 清单页的文字（sharedStrings）同步去单位、去序号
    cs = cfg.get("clean_strings")
    if cs:
        xml["xl/sharedStrings.xml"] = _clean_shared(
            xml.get("xl/sharedStrings.xml", b""), cs).encode("utf-8")

    # 6) 写回（保留除 chart/sharedStrings 外的所有条目）
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for n, d in xml.items():
            zo.writestr(n, d)

    # 7) 数据零改动校验：逐图比对 numCache
    report["verify"] = verify(src, out)
    return report


def _clean_shared(raw, cs):
    s = raw.decode("utf-8")
    marks = cs.get("contains") or []
    if isinstance(marks, str):
        marks = [marks]

    def one(m):
        t = m.group(1)
        if marks and not all(x in t for x in marks):
            return m.group(0)
        if cs.get("strip_index", True):
            t = re.sub(INDEX_IN_TEXT_RE, "", t)
        if cs.get("strip_units", True):
            t = re.sub(UNIT_INLINE_RE, "", t)
        for a, b in (cs.get("rename") or {}).items():
            t = t.replace(a, b)
        t = re.sub(r"\s*\[", " [", t)
        t = re.sub(r"\s{2,}", " ", t).strip()
        return "<t>%s</t>" % escape(t)

    return re.sub(r"<t[^>]*>([^<]*)</t>", one, s)


def _numcache(xml_bytes):
    s = xml_bytes.decode("utf-8")
    return re.findall(r"<c:numCache>.*?</c:numCache>", s, re.S)


def verify(a, b):
    xa, xb = load(a), load(b)
    ok = bad = 0
    for n in xa:
        if n.startswith("xl/charts/chart") and n.endswith(".xml") and n in xb:
            if _numcache(xa[n]) == _numcache(xb[n]):
                ok += 1
            else:
                bad += 1
    return {"charts": ok, "data_changed": bad}


def init_config(path):
    xml = load(path)
    data = {"groups": [], "strip_units": True, "strip_index": True,
            "axis_min_span_ratio": 0.4,
            "clean_strings": {"contains": ["[数据页"], "strip_index": True,
                              "strip_units": True}}
    for name, charts in sheets_of(xml):
        if not charts or name in ("说明", "Sheet1"):
            continue
        data["groups"].append({"name": name, "sheets": [name], "ref": 0,
                               "charts": len(charts)})
    return data


def main(argv=None):
    ap = argparse.ArgumentParser(description="批量统一 Excel 内嵌图表版式")
    ap.add_argument("src", nargs="?")
    ap.add_argument("out", nargs="?")
    ap.add_argument("--config", "-c")
    ap.add_argument("--init", metavar="XLSX", help="打印结构草稿（JSON）后退出")
    ap.add_argument("--json", action="store_true", help="报告输出 JSON")
    a = ap.parse_args(argv)

    if a.init:
        print(json.dumps(init_config(a.init), ensure_ascii=False, indent=2))
        return 0
    if not (a.src and a.out):
        ap.error("需要 SRC 和 OUT（或用 --init）")
    cfg = {}
    if a.config:
        with open(a.config, encoding="utf-8") as f:
            cfg = json.load(f)
    rep = uniform(a.src, a.out, cfg)

    if a.json:
        print(json.dumps(rep, ensure_ascii=False, indent=2))
    else:
        for g in rep["groups"]:
            note = " 保留自刻度 %d 张" % len(g["kept_scale"]) if g["kept_scale"] else ""
            print("[%s] %d 张，基准 %s%s" % (g["name"], g["charts"], g["ref"], note))
        if rep["missing"]:
            print("!! 找不到工作表: %s" % "、".join(rep["missing"]))
        print("数据校验: %d 张图缓存值一致，%d 张有变化" %
              (rep["verify"]["charts"], rep["verify"]["data_changed"]))
        print("OUT", a.out)
    return 0 if not rep["verify"]["data_changed"] else 2


if __name__ == "__main__":
    sys.exit(main())
