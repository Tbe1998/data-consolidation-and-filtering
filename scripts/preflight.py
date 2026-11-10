#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""执行前环境自检。任何 agent 在跑本技能脚本前都应先跑这一条。

用法：
    python3 scripts/preflight.py                    # 查全部
    python3 scripts/preflight.py --route A          # 长表→宽表（零依赖）
    python3 scripts/preflight.py --route B          # 列表筛选 / 监控（零依赖）
    python3 scripts/preflight.py --route C          # 曲线图四件套（需 openpyxl/Pillow/matplotlib）
    python3 scripts/preflight.py --route C --json

退出码：0 = 可以开工；1 = 有阻塞项（按输出的 fix 命令处理后重跑）。
"""
import argparse
import glob
import importlib.util
import json
import os
import platform
import py_compile
import sys
import tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

SCRIPTS = {
    "A": ["long_to_wide.py", "pick_representatives.py"],
    "B": ["long_to_wide.py", "pick_representatives.py"],
    "C": ["xlsx_chart_uniform.py", "xlsx_charts_to_png.py",
          "pick_profile_curves.py", "make_delivery.py", "verify_delivery.py"],
}

FONT_CANDIDATES = [
    "/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc",
    "/usr/share/fonts/truetype/wqy/wqy-microhei.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf",
    "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
    "C:/Windows/Fonts/msyh.ttc",
    "C:/Windows/Fonts/simhei.ttf",
    "/System/Library/Fonts/PingFang.ttc",
]
FONT_GLOBS = [
    os.path.expanduser("~/.fonts/*.tt[fc]"),
    os.path.expanduser("~/.fonts/*.otf"),
    os.path.expanduser("~/.local/share/fonts/*"),
    "/usr/share/fonts/**/*CJK*",
    "/usr/share/fonts/**/*NotoSansSC*",
    "/usr/share/fonts/**/*wqy*",
]


def find_fonts():
    found = [p for p in FONT_CANDIDATES if os.path.exists(p)]
    for pat in FONT_GLOBS:
        found.extend(glob.glob(pat, recursive=True))
    out = []
    for p in found:
        if p not in out and os.path.isfile(p):
            out.append(p)
    return out


def has_mod(name):
    try:
        return importlib.util.find_spec(name) is not None
    except Exception:
        return False


def main(argv=None):
    ap = argparse.ArgumentParser(description="data-consolidation-and-filtering 环境自检")
    ap.add_argument("--route", choices=["A", "B", "C"], default="C",
                    help="A=长表→宽表 B=列表筛选/监控 C=曲线图四件套")
    ap.add_argument("--summary", default=None, help="（路线 C）检查汇总工作簿结构是否可被脚本识别")
    ap.add_argument("--workbook", default=None, help="（路线 C）检查原始曲线工作簿是否含内嵌图表")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)

    blockers, warnings, fixes = [], [], []
    if sys.version_info < (3, 8):
        blockers.append("Python %s < 3.8" % platform.python_version())
        fixes.append("装 Python 3.8+（建议 3.11）后重跑")

    # 1) 脚本完整性与语法
    missing_scripts, bad_scripts = [], []
    for name in SCRIPTS[a.route]:
        p = os.path.join(HERE, name)
        if not os.path.exists(p):
            missing_scripts.append(name)
            continue
        try:
            py_compile.compile(p, doraise=True,
                               cfile=os.path.join(tempfile.gettempdir(), name + ".pyc"))
        except Exception as e:
            bad_scripts.append("%s: %s" % (name, e))
    if missing_scripts:
        blockers.append("缺少脚本: " + ", ".join(missing_scripts))
        fixes.append("重新克隆仓库，确认 scripts/ 与 SKILL.md 同级")
    if bad_scripts:
        blockers.append("脚本语法错误: " + "; ".join(bad_scripts))
        fixes.append("重新拉取仓库，不要手改脚本")

    # 2) 依赖
    pkgs = {"openpyxl": has_mod("openpyxl"), "PIL": has_mod("PIL"),
            "matplotlib": has_mod("matplotlib")}
    if a.route == "C":
        if not pkgs["openpyxl"]:
            blockers.append("缺 openpyxl（写 xlsx / 生成四件套必需）")
        if not pkgs["PIL"]:
            blockers.append("缺 Pillow（PDF 拼版与封面必需）")
        if not pkgs["matplotlib"]:
            blockers.append("缺 matplotlib（把内嵌图表重绘成 PNG 必需）")
    else:
        for k in ("openpyxl", "PIL", "matplotlib"):
            if not pkgs[k]:
                warnings.append("未装 %s（本路线可不装；要读 xlsx 时再补）" % k)
    if any(not v for v in pkgs.values()):
        fixes.append(
            "python3 -m venv .venv && .venv/bin/pip install -i "
            "https://pypi.tuna.tsinghua.edu.cn/simple openpyxl pillow matplotlib"
            "   # 然后用 .venv/bin/python 跑脚本")

    # 3) 中文字体（matplotlib 出图必需；缺了图内中文会成豆腐块）
    fonts = find_fonts()
    if a.route == "C" and not fonts:
        warnings.append("未找到中文字体：图内中文可能显示为方块")
        fixes.append(
            "放一个中文字体到 ~/.fonts/ 后 fc-cache -f，或用 --font 指定 .ttf/.ttc 路径；"
            "也可只用英文图内文字（Excel 表内保持中文）")

    # 4) 写权限
    try:
        fd, p = tempfile.mkstemp(prefix="dcf_")
        os.write(fd, b"ok")
        os.close(fd)
        os.unlink(p)
    except Exception as e:
        blockers.append("临时目录不可写: %s" % e)
        fixes.append("换一个有写权限的 TMPDIR 再跑")

    # 5) 输入文件结构（可选，但强烈建议）
    inputs = {}
    if a.summary:
        if not os.path.exists(a.summary):
            blockers.append("汇总工作簿不存在: %s" % a.summary)
        elif not pkgs["openpyxl"]:
            blockers.append("装了 openpyxl 才能检查 --summary")
        else:
            import re
            import openpyxl
            try:
                wb = openpyxl.load_workbook(a.summary, data_only=True)
                items, points, holes = [], 0, []
                for ws in wb.worksheets:
                    rows = [list(r) for r in ws.iter_rows(values_only=True)]
                    if ws.title.startswith("测斜"):
                        holes.append(ws.title[2:])
                    elif (len(rows) > 2 and str(rows[1][0]).strip() == "点号"
                          and re.match(r"\d{4}-\d{2}-\d{2}", str(rows[1][1]))):
                        items.append(ws.title)
                        points += max(0, len(rows) - 2)
                inputs["summary"] = {"items": items, "points": points, "holes": holes}
                if not items:
                    blockers.append("汇总表没有可识别的测项宽表（第2行需 A=点号, B=YYYY-MM-DD）")
                    fixes.append("按 references/agent-runbook.md §3 调整表头后重跑")
                if not holes:
                    warnings.append("汇总表没有以「测斜」开头的工作表：测斜曲线/孔表将为空")
                if not items and not holes:
                    fixes.append("确认 --summary 指向的是数据汇总表，不是原始曲线工作簿")
            except Exception as e:
                blockers.append("汇总工作簿打不开: %s" % e)
    if a.workbook:
        if not os.path.exists(a.workbook):
            blockers.append("原始曲线工作簿不存在: %s" % a.workbook)
        else:
            try:
                import zipfile
                with zipfile.ZipFile(a.workbook) as z:
                    charts = [n for n in z.namelist()
                              if n.startswith("xl/charts/") and n.endswith(".xml")]
                inputs["workbook"] = {"charts": len(charts)}
                if not charts:
                    blockers.append("原始工作簿没有内嵌图表（xl/charts/*.xml 为 0）：请改用路线 A")
                    fixes.append("要出曲线图必须提供 Excel 内嵌图表工作簿")
            except Exception as e:
                blockers.append("原始工作簿不是有效 xlsx: %s" % e)

    ok = not blockers
    report = {
        "ok": ok,
        "route": a.route,
        "python": platform.python_version(),
        "skill_dir": ROOT,
        "scripts_dir": HERE,
        "packages": pkgs,
        "fonts": fonts[:5],
        "inputs": inputs,
        "blockers": blockers,
        "warnings": warnings,
        "fix": list(dict.fromkeys(fixes)),
    }
    if a.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print("preflight route=%s python=%s skill=%s" % (a.route, report["python"], ROOT))
        print("依赖: " + " ".join("%s=%s" % (k, "ok" if v else "MISSING") for k, v in pkgs.items()))
        print("中文字体: " + (", ".join(os.path.basename(f) for f in fonts[:3]) if fonts else "未找到"))
        for key, val in inputs.items():
            print("输入 %s: %s" % (key, val))
        for w in warnings:
            print("WARN  " + w)
        for b in blockers:
            print("BLOCK " + b)
        if not ok:
            print("按以下命令修复后重跑：")
            for f in report["fix"]:
                print("  " + f)
        else:
            print("OK 可以开工")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
