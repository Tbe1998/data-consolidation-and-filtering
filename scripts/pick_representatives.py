#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从「实体×时间」矩阵里挑 k 个代表列（日期），供画图/抽样用。

核心区别（用户纠正过两次的点）：
  - 画图展示要「铺得开」→ 按**度量值分位等距**挑（默认 metric 模式）
  - 做趋势判断要「全过程」→ 按**时间轴等距**挑（mode="time"）
按序号等距是最差的选择：采样前密后疏时会挤成一团。

用法：
    from pick_representatives import pick
    dates, amps = pick(matrix, k=6, mode="metric")
    # matrix: {date: [v1, v2, ...]}  或  {date: {depth: v}}

命令行自测：
    python3 pick_representatives.py
"""
import datetime


def _metric(values, how="maxabs"):
    vals = [abs(float(v)) for v in values if v is not None]
    if not vals:
        return 0.0
    if how == "maxabs":
        return max(vals)
    if how == "range":
        return max(vals) - min(vals)
    if how == "mean":
        return sum(vals) / len(vals)
    raise ValueError(how)


def pick(matrix, k=6, mode="metric", how="maxabs", tie="earlier"):
    """matrix: {date: iterable_of_values}

    返回 (picked_dates, amp_dict)。
    mode="metric"：在 [min amp, max amp] 上 k 等分取目标值，各取最接近且未用过的日期，
                   返回按 amp 升序排列（第 1 个最小、第 k 个最大）。
    mode="time"  ：在首末日期之间 k 等分，吸附到最近的实测日期（同距取较早）。
    """
    dates = sorted(matrix)
    if not dates:
        return [], {}
    amp = {d: _metric(matrix[d], how) for d in dates}
    if len(dates) <= k:
        return dates, amp

    if mode == "time":
        lo, hi = dates[0], dates[-1]
        span = (hi - lo).total_seconds()
        used, out = set(), []
        for i in range(k):
            target = lo + datetime.timedelta(seconds=span * i / (k - 1))
            cand = sorted(
                ((abs((d - target).total_seconds()), d) for d in dates if d not in used),
                key=lambda x: (x[0], x[1]),
            )
            used.add(cand[0][1])
            out.append(cand[0][1])
        return sorted(out), amp

    lo_a, hi_a = min(amp.values()), max(amp.values())
    if hi_a - lo_a < 1e-9:          # 度量几乎不变 → 退化为时间等距
        return pick(matrix, k, mode="time", how=how, tie=tie)

    used, out = set(), []
    for i in range(k):
        target = lo_a + (hi_a - lo_a) * i / (k - 1)
        cand = sorted(
            ((abs(amp[d] - target), d) for d in dates if d not in used),
            key=lambda x: (x[0], x[1] if tie == "earlier" else -x[1].toordinal()),
        )
        used.add(cand[0][1])
        out.append(cand[0][1])
    return sorted(out, key=lambda d: amp[d]), amp


def _selftest():
    base = datetime.date(2024, 7, 30)
    # 模拟"前密后疏、位移前快后慢"：前 20 天涨到 26，之后 100 天只涨到 27
    m = {}
    for i in range(120):
        d = base + datetime.timedelta(days=i)
        v = 26 * min(i, 20) / 20 + 1 * max(0, i - 20) / 100
        m[d] = [round(v, 3)]
    for mode in ("metric", "time"):
        picked, amp = pick(m, k=6, mode=mode)
        print(mode, [(d.strftime("%Y-%m-%d"), round(amp[d], 2)) for d in picked])
    # metric 期望：特征值大致均匀分布在 0→27；time 期望：日期均匀、末条才到峰值


if __name__ == "__main__":
    _selftest()
