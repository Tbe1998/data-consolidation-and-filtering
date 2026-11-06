#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""演示「按度量分位等距」vs「按时间等距」选代表点的差别。

场景：一条位移曲线前 20 天涨到 26mm，之后 100 天只从 26 涨到 27mm
（典型的「前快后慢、采样前密后疏」）。

- 按时间等距：后 5 条全挤在 26.0~27.0 之间，画出来叠成一团；
- 按度量分位：0 → 5 → 10 → 16 → 22 → 27，在图上均匀铺开。

运行：python3 examples/demo.py
"""
import datetime
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
from pick_representatives import pick  # noqa: E402

BASE = datetime.date(2024, 7, 30)


def make_matrix(n=120, fast_days=20, fast_to=26.0, slow_to=27.0):
    m = {}
    for i in range(n):
        d = BASE + datetime.timedelta(days=i)
        v = fast_to * min(i, fast_days) / fast_days
        if i > fast_days:
            v += (slow_to - fast_to) * (i - fast_days) / (n - 1 - fast_days)
        m[d] = [round(v, 3)]
    return m


def main():
    m = make_matrix()
    rows = []
    for mode, label in (("metric", "按度量分位等距"), ("time", "按时间等距")):
        picked, amp = pick(m, k=6, mode=mode)
        rows.append((label, [(d.strftime("%Y-%m-%d"), round(amp[d], 2)) for d in picked]))

    print("| 选法 | 6 个代表点（日期 → 特征值 mm） |")
    print("|---|---|")
    for label, picked in rows:
        print("| %s | %s |" % (label, " ・ ".join("%s→%.2f" % p for p in picked)))

    print()
    print("结论：metric 第 1 条≈最小值、第 6 条≈最大值，中间均匀；")
    print("      time  后 5 条都挤在 26~27，画图会叠成一团。")


if __name__ == "__main__":
    main()
