#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""零依赖冒烟测试：python3 -m unittest discover -s tests -v"""
import csv
import os
import subprocess
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "scripts"))

from pick_representatives import pick  # noqa: E402
import datetime  # noqa: E402


class TestLongToWide(unittest.TestCase):
    def test_sample(self):
        tmp = tempfile.mkdtemp()
        out = os.path.join(tmp, "wide.csv")
        r = subprocess.run(
            [sys.executable, os.path.join(ROOT, "scripts", "long_to_wide.py"),
             os.path.join(ROOT, "examples", "sample_long.csv"), out],
            capture_output=True, text=True)
        self.assertEqual(r.returncode, 0, r.stderr)

        with open(out, encoding="utf-8-sig") as f:
            rows = list(csv.reader(f))
        header, body = rows[0], {r[0]: r for r in rows[1:]}

        self.assertEqual(header[0], "entity")
        # 空实体被丢弃，非法值被跳过 → 只剩 3 个实体
        self.assertEqual(sorted(body), ["TST1", "TST2", "ZQT1"])
        # 2024/08/07 归一成标准日期列
        self.assertIn("2024-08-07", header)
        self.assertNotIn("2024/08/07", header)
        # 冲突后写覆盖先写：2.60 → 2.61
        self.assertEqual(body["TST2"][header.index("2024-08-05")], "2.610")
        # 空值不产生列（没测 ≠ 0）
        self.assertNotIn("2024-08-10", header)
        # 未测的格子也是空
        self.assertEqual(body["TST1"][header.index("2024-08-20")], "")


class TestPick(unittest.TestCase):
    def _matrix(self):
        base = datetime.date(2024, 7, 30)
        m = {}
        for i in range(120):
            v = 26 * min(i, 20) / 20 + (1 * max(0, i - 20) / 100)
            m[base + datetime.timedelta(days=i)] = [v]
        return m

    def test_metric_spreads_evenly(self):
        dates, amp = pick(self._matrix(), k=6, mode="metric")
        vals = [amp[d] for d in dates]
        self.assertEqual(len(dates), 6)
        self.assertEqual(vals, sorted(vals))
        self.assertLess(vals[0], 1.0)
        self.assertGreater(vals[-1], 26.0)
        # 中间几条不能全挤在峰值附近
        self.assertLess(vals[3], 25.0)

    def test_time_crowds_at_end(self):
        dates, amp = pick(self._matrix(), k=6, mode="time")
        vals = [amp[d] for d in dates]
        self.assertGreater(sum(1 for v in vals[1:] if v > 25.0), 3)


if __name__ == "__main__":
    unittest.main()
