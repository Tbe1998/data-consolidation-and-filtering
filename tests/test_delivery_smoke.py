#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""端到端冒烟：最小输入 → make_delivery → verify_delivery 必须通过。

覆盖最容易踩的两件事：
  1. 不给 --incl 时自动找 <charts>/incl，且孔号匹配不依赖文件名模板；
  2. 交付件数量自洽（PDF 页数、ZIP 张数、嵌图数、最大表行数）。
"""
import json
import re
import math
import os
import subprocess
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAKE = os.path.join(ROOT, "scripts", "make_delivery.py")
VERIFY = os.path.join(ROOT, "scripts", "verify_delivery.py")

try:
    import openpyxl
    from openpyxl.drawing.image import Image as XLImage  # noqa: F401
    from PIL import Image
    HAVE_DEPS = True
except Exception:                                      # pragma: no cover
    HAVE_DEPS = False


def build_fixture(d):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "地表沉降"
    ws.append(["110二期 X3#井 地表沉降监测"])
    ws.append(["点号", "2024-07-30", "2024-08-30", "2024-09-30"])
    for name, vals in (("DL1", [0, -3.2, -5.5]), ("DL2", [0, -1.1, -2.2]),
                       ("DL10", [0, 0.5, 1.0])):
        ws.append([name] + vals)
    ws2 = wb.create_sheet("测斜TST1")
    ws2.append(["TST1 深层水平位移"])
    ws2.append(["深度(m)", "2024-07-30", "2024-08-30"])
    for dep, a, b in ((0.5, 0, 0.2), (5, 1.2, 3.4), (10, -0.4, 0.1)):
        ws2.append([dep, a, b])
    summary = os.path.join(d, "summary.xlsx")
    wb.save(summary)

    charts = os.path.join(d, "png")
    os.makedirs(os.path.join(charts, "incl"))
    Image.new("RGB", (1200, 360), "white").save(os.path.join(charts, "a.png"))
    Image.new("RGB", (1200, 360), "white").save(os.path.join(charts, "b.png"))
    # 故意不用旧模板名：验证按孔号兜底匹配
    Image.new("RGB", (700, 1600), "white").save(
        os.path.join(charts, "incl", "X3 TST1 位移-深度曲线 全过程.png"))

    groups = [
        {"sheet": "地表沉降", "title": "地表沉降 时程曲线 DL1~DL2", "file": "a.png",
         "points": ["DL1", "DL2"], "dates": ["2024-07-30", "2024-08-30", "2024-09-30"],
         "series": {"DL1": [0, -3.2, -5.5], "DL2": [0, -1.1, -2.2]}},
        {"sheet": "地表沉降", "title": "地表沉降 时程曲线 DL10", "file": "b.png",
         "points": ["DL10"], "dates": ["2024-07-30", "2024-08-30", "2024-09-30"],
         "series": {"DL10": [0, 0.5, 1.0]}},
    ]
    gp = os.path.join(d, "_groups.json")
    json.dump(groups, open(gp, "w", encoding="utf-8"), ensure_ascii=False)
    sel = os.path.join(d, "_sel.json")
    json.dump([{"孔号": "TST1", "选定日期": ["2024-07-30", "2024-08-30"]}],
              open(sel, "w", encoding="utf-8"), ensure_ascii=False)
    return summary, gp, charts, sel


@unittest.skipUnless(HAVE_DEPS, "需要 openpyxl + Pillow")
class TestDeliverySmoke(unittest.TestCase):
    def test_end_to_end(self):
        d = tempfile.mkdtemp()
        summary, gp, charts, sel = build_fixture(d)
        out = os.path.join(d, "out")
        r = subprocess.run([sys.executable, MAKE, "--summary", summary, "--groups", gp,
                            "--charts", charts, "--sel", sel,
                            "--title", "测试项目", "--out", out],
                           capture_output=True, text=True)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.assertIn("识别: 测项 1 个 / 时程图 2 张 / 测斜图 1 张", r.stdout)
        self.assertEqual(len([f for f in os.listdir(out) if not f.startswith(".")]), 5)

        v = subprocess.run([sys.executable, VERIFY, "--dir", out, "--title", "测试项目",
                            "--expect-incl", "1"], capture_output=True, text=True)
        self.assertEqual(v.returncode, 0, v.stdout + v.stderr)
        self.assertIn("PDF 页数", v.stdout)
        self.assertIn("筛选表嵌图", v.stdout)

        # PDF 页数 = 封面 + ceil(2/4) + 1 孔（Pillow 只能写不能读 PDF，直接数对象）
        pdf = os.path.join(out, "测试项目_曲线图.pdf")
        with open(pdf, "rb") as fh:
            data = fh.read()
        self.assertEqual(len(re.findall(rb"/Type\s*/Page[^s]", data)),
                         1 + int(math.ceil(2 / 4.0)) + 1)


if __name__ == "__main__":
    unittest.main()
