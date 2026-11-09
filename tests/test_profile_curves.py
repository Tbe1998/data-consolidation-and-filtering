#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pick_profile_curves.py 的单测：曲线数量随位移发展复杂度自适应。"""
import csv
import datetime
import os
import subprocess
import sys
import tempfile
import unittest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(ROOT, "scripts", "pick_profile_curves.py")
sys.path.insert(0, os.path.join(ROOT, "scripts"))

from pick_profile_curves import choose, read_matrix  # noqa: E402


def make_csv(path, n=60, bump=False, dup=False):
    base = datetime.date(2024, 7, 30)
    dates = [(base + datetime.timedelta(days=2 * i)).isoformat() for i in range(n)]
    depths = [0.5 * (i + 1) for i in range(20)]
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["深度"] + dates)
        for d in depths:
            row = []
            for i, _ in enumerate(dates):
                # 0~20 期增长，20~40 平台，40 期后回落
                amp = min(i, 20) * 0.8 if i <= 20 else (16.0 if i <= 40 else 16.0 - (i - 40) * 0.2)
                v = amp * (d / 10.0) if d <= 10 else amp * (1 - (d - 10) / 20.0)
                if bump and 28 <= i <= 30:
                    v += 6.0 * (d / 10.0)        # 中段一次明显突变
                if dup and i == 25:
                    v = (min(24, 20) * 0.8) * (d / 10.0 if d <= 10 else 1 - (d - 10) / 20.0)
                row.append(round(v, 3))
            w.writerow([d] + row)
    return dates


class TestProfileCurves(unittest.TestCase):
    def test_endpoints_peak_and_sorted(self):
        p = os.path.join(tempfile.mkdtemp(), "m.csv")
        dates = make_csv(p)
        ds, data = read_matrix(p)
        r = choose(ds, data)
        self.assertEqual(r["selected"][0], dates[0])
        self.assertEqual(r["selected"][-1], dates[-1])
        self.assertIn(r["peak_date"], r["selected"])
        self.assertEqual(r["selected"], sorted(r["selected"]))

    def test_error_under_tolerance_or_capped(self):
        p = os.path.join(tempfile.mkdtemp(), "m.csv")
        make_csv(p)
        ds, data = read_matrix(p)
        r = choose(ds, data, cap=9)
        self.assertTrue(r["max_reconstruction_error_mm"] <= r["tolerance_mm"] + 1e-9
                        or r["n_selected"] == 9, r)

    def test_more_curves_when_mid_bump(self):
        a = os.path.join(tempfile.mkdtemp(), "a.csv")
        b = os.path.join(tempfile.mkdtemp(), "b.csv")
        make_csv(a)
        make_csv(b, bump=True)
        ra = choose(*read_matrix(a))
        rb = choose(*read_matrix(b))
        self.assertGreaterEqual(rb["n_selected"], ra["n_selected"])

    def test_near_duplicate_pruned(self):
        p = os.path.join(tempfile.mkdtemp(), "d.csv")
        make_csv(p, dup=True)
        ds, data = read_matrix(p)
        r = choose(ds, data, min_k=4)
        dates = [d for d in ds if d in data[list(data)[0]]]
        i = dates.index(ds[25]) if ds[25] in dates else None
        # 与 24 期完全相同的 25 期不应被单独选中
        if i:
            self.assertFalse(ds[25] in r["selected"] and ds[25] not in
                             (ds[0], ds[-1], r["peak_date"]))

    def test_cli_text(self):
        p = os.path.join(tempfile.mkdtemp(), "m.csv")
        make_csv(p)
        out = subprocess.run([sys.executable, SCRIPT, p, "--format", "text"],
                             capture_output=True, text=True)
        self.assertEqual(out.returncode, 0, out.stderr)
        self.assertIn("选", out.stdout)


if __name__ == "__main__":
    unittest.main()
