#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""交付前验收：核对「曲线图 + 三张表 + PNG 包」是否真的齐全、数量自洽。

用法：
    python3 scripts/verify_delivery.py --dir deliverables --title "110二期X3#井"
    python3 scripts/verify_delivery.py --dir out --title 测试项目 \
        --groups out/png/_groups.json --expect-incl 10 --strict

检查项：
    1. 5 个交付件存在且能被打开（PDF 页数、xlsx 工作表、zip 条目）；
    2. PDF 页数 = 1 + ceil(时程图数/4) + 测斜孔数；
    3. PNG 包里时程 + 测斜张数与 _索引.csv 一致；
    4. 筛选后汇总表的嵌图数 = 时程图数 + 测斜孔数；
    5. 各点最大累计值表行数 = 测点数（+ 孔数）；
    6. _groups.json 每组测点 ≤5、图名无单位/序号、日期不像 Excel 序列号残留。

退出码：0 全过；2 有 FAIL（--strict 时 WARN 也算 FAIL）；1 参数/文件缺失。
"""
import argparse
import csv
import io
import json
import math
import os
import re
import sys
import zipfile

try:
    import openpyxl
except Exception:                                     # pragma: no cover
    openpyxl = None

RESULTS = []


def add(level, name, detail=""):
    RESULTS.append((level, name, detail))


def ok(name, detail=""):
    add("PASS", name, detail)


def warn(name, detail=""):
    add("WARN", name, detail)


def fail(name, detail=""):
    add("FAIL", name, detail)


def pdf_pages(path):
    try:
        from PIL import Image
        with Image.open(path) as im:
            return getattr(im, "n_frames", 1)
    except Exception:
        pass
    data = open(path, "rb").read()
    return len(re.findall(rb"/Type\s*/Page[^s]", data))


def detect_items(rows):
    """与 make_delivery 相同判据：第 2 行是「点号 + 日期」。"""
    if len(rows) < 3:
        return False
    return (str(rows[1][0]).strip() == "点号"
            and re.match(r"\d{4}-\d{2}-\d{2}", str(rows[1][1])))


def main(argv=None):
    ap = argparse.ArgumentParser(description="交付件验收")
    ap.add_argument("--dir", required=True, help="交付目录")
    ap.add_argument("--title", required=True, help="项目名（与文件名前缀一致）")
    ap.add_argument("--groups", default=None, help="_groups.json 路径（可选，缺省在目录内找）")
    ap.add_argument("--expect-incl", type=int, default=None, help="预期测斜孔数，用于强校验")
    ap.add_argument("--strict", action="store_true", help="WARN 也按 FAIL 处理")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)

    if openpyxl is None:
        print("缺 openpyxl，无法验收 xlsx。先跑 scripts/preflight.py --route C")
        return 1

    d, t = a.dir, a.title
    files = {
        "汇总": os.path.join(d, "%s_数据汇总表.xlsx" % t),
        "筛选": os.path.join(d, "%s_筛选后数据汇总表.xlsx" % t),
        "最大": os.path.join(d, "%s_各点最大累计值汇总表.xlsx" % t),
        "PDF": os.path.join(d, "%s_曲线图.pdf" % t),
        "ZIP": os.path.join(d, "%s_曲线图_PNG.zip" % t),
    }
    for k, p in files.items():
        if os.path.exists(p) and os.path.getsize(p) > 0:
            ok("%s存在" % k, "%s (%.0f KB)" % (os.path.basename(p), os.path.getsize(p) / 1024.0))
        else:
            fail("%s存在" % k, "缺失或为空: %s" % p)
    if any(level == "FAIL" for level, _, _ in RESULTS):
        return _report(a, 2)

    wb_sum = openpyxl.load_workbook(files["汇总"], data_only=True)
    wb_max = openpyxl.load_workbook(files["最大"], data_only=True)
    wb_sel = openpyxl.load_workbook(files["筛选"], data_only=True)

    # ---- 汇总表：识别测项宽表与测斜矩阵
    data = {ws.title: [list(r) for r in ws.iter_rows(values_only=True)]
            for ws in wb_sum.worksheets}
    items = {k: v for k, v in data.items() if not k.startswith("测斜") and detect_items(v)}
    holes = {k: v for k, v in data.items() if k.startswith("测斜")}
    n_points = sum(max(0, len(v) - 2) for v in items.values())
    ok("汇总表结构", "%d 个测项 / %d 个测点 / %d 个测斜孔 / %d 个工作表"
       % (len(items), n_points, len(holes), len(data)))
    if not items:
        fail("测项识别", "没有任何工作表满足「第2行=点号+日期」，汇总表结构不符")

    # ---- 最大累计值表
    ws = wb_max.worksheets[0]
    n_max = max(0, ws.max_row - 1)
    if n_max == n_points:
        ok("最大累计值行数", "%d = 测点数" % n_max)
    else:
        warn("最大累计值行数", "%d 行 vs 汇总表 %d 点（有隐藏/空行时可能不等）" % (n_max, n_points))
    if len(wb_max.worksheets) > 1:
        n_hole_max = max(0, wb_max.worksheets[1].max_row - 1)
        if n_hole_max == len(holes):
            ok("测斜孔最大表行数", "%d = 孔数" % n_hole_max)
        else:
            fail("测斜孔最大表行数", "%d 行 vs 汇总表 %d 孔" % (n_hole_max, len(holes)))

    # ---- groups
    gp = a.groups or os.path.join(os.path.dirname(files["PDF"]), "_groups.json")
    groups = []
    if os.path.exists(gp):
        groups = json.load(open(gp, encoding="utf-8"))
        ok("_groups.json", "%d 张时程图" % len(groups))
    else:
        for cand in (os.path.join(d, "png", "_groups.json"),
                     os.path.join(os.path.dirname(os.path.abspath(d)), "png", "_groups.json"),
                     os.path.join(os.path.dirname(files["PDF"]), "png", "_groups.json")):
            if os.path.exists(cand):
                groups = json.load(open(cand, encoding="utf-8"))
                gp = cand
                ok("_groups.json", "%d 张时程图（自动找到 %s）" % (len(groups), cand))
                break
        else:
            warn("_groups.json", "未找到，按 ZIP 索引反推时程图数")

    for g in groups:
        pts = g.get("points") or []
        if len(pts) > 5:
            fail("分块测点数", "%s 有 %d 个测点（应 ≤5）" % (g.get("title"), len(pts)))
        if re.search(r"[（(](mm|kN|‰|°)[）)]\s*$", g.get("title", "")):
            fail("图名单位", "%s 结尾带单位" % g.get("title"))
        if re.search(r"\d+\s*/\s*\d+\s*$", g.get("title", "")):
            fail("图名序号", "%s 结尾带 n/m 序号" % g.get("title"))
        bad_dates = [x for x in (g.get("dates") or []) if str(x)[:4] in ("1899", "1900")]
        if bad_dates:
            warn("日期序列", "%s 含 %s（x 值可能不是 Excel 日期序列）"
                 % (g.get("title"), bad_dates[:3]))

    # ---- ZIP
    z = zipfile.ZipFile(files["ZIP"])
    names = [n for n in z.namelist() if n.lower().endswith(".png")]
    idx_name = next((n for n in z.namelist() if n.endswith("_索引.csv")), None)
    n_time = n_incl = None
    if idx_name:
        txt = z.read(idx_name).decode("utf-8-sig")
        rows = list(csv.DictReader(io.StringIO(txt)))
        n_time = sum(1 for r in rows if "时程" in (r.get("类型") or ""))
        n_incl = sum(1 for r in rows if "测斜" in (r.get("类型") or ""))
        if len(rows) == len(names):
            ok("ZIP 索引", "%d 条 = %d 张 PNG（时程 %d / 测斜 %d）"
               % (len(rows), len(names), n_time, n_incl))
        else:
            fail("ZIP 索引", "%d 条 vs %d 张 PNG" % (len(rows), len(names)))
    else:
        warn("ZIP 索引", "没有 _索引.csv")
    if a.expect_incl is not None and n_incl is not None and n_incl != a.expect_incl:
        fail("测斜图数", "%d 张 vs 预期 %d 孔" % (n_incl, a.expect_incl))

    # ---- PDF 页数
    pages = pdf_pages(files["PDF"])
    if n_time is None:
        n_time = len(groups)
    if n_incl is None:
        n_incl = len(holes)
    expect_pages = 1 + int(math.ceil(n_time / 4.0)) + n_incl if n_time else 1 + n_incl
    if pages == expect_pages:
        ok("PDF 页数", "%d = 1 + ceil(%d/4) + %d" % (pages, n_time, n_incl))
    else:
        fail("PDF 页数", "%d 页 vs 预期 %d（时程 %d / 测斜 %d）"
             % (pages, expect_pages, n_time, n_incl))

    # ---- 筛选后汇总表：嵌图数
    n_img = 0
    for ws in wb_sel.worksheets:
        n_img += len(getattr(ws, "_images", []))
    expect_img = (len(groups) if groups else (n_time or 0)) + (n_incl or 0)
    if expect_img and n_img == expect_img:
        ok("筛选表嵌图", "%d 张 = 时程 %d + 测斜 %d"
           % (n_img, len(groups) if groups else (n_time or 0), n_incl or 0))
    elif expect_img:
        fail("筛选表嵌图", "%d 张 vs 预期 %d" % (n_img, expect_img))

    # ---- 单位提示
    known = [k for k in items if not any(u in k for u in ("沉降", "位移", "水位", "倾斜", "轴力", "测斜"))]
    if known:
        warn("单位需人工确认", "、".join(known[:5]))

    return _report(a, 2 if any(r[0] == "FAIL" for r in RESULTS) else
                   (2 if a.strict and any(r[0] == "WARN" for r in RESULTS) else 0))


def _report(a, code):
    fails = sum(1 for r in RESULTS if r[0] == "FAIL")
    warns = sum(1 for r in RESULTS if r[0] == "WARN")
    if a.json:
        print(json.dumps({"ok": code == 0, "fails": fails, "warns": warns,
                          "items": [{"level": l, "name": n, "detail": d} for l, n, d in RESULTS]},
                         ensure_ascii=False, indent=2))
    else:
        for level, name, detail in RESULTS:
            print("%-4s %-14s %s" % (level, name, detail))
        print("\n结果: %s（FAIL %d / WARN %d）" % ("通过" if code == 0 else "未通过", fails, warns))
    return code


if __name__ == "__main__":
    raise SystemExit(main())
