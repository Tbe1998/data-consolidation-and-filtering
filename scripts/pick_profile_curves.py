#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""深度剖面曲线的「数量自适应」选点（零第三方依赖）。

场景：测斜/沉降剖面有几十上百期，全画会糊成一团，固定取 6 期又可能漏掉关键转折。
做法：不预设张数，用「线性插值重构误差」决定需要几张：

    1) 首测、末测、峰值期必选；
    2) 反复找「与左右两张已选曲线线性插值差得最远」的中间期，
       差值 > tol 就加入；
    3) 相邻两张剖面最大差值 < 1.5 mm 视为视觉重复（保留首/末/峰值）；
    4) 不足 min_k 张时按最大时间间隔补点。

tol 默认 max(2.0 mm, 峰值×10%)，上限 cap 张。位移发展越复杂，自动多给几张。

矩阵格式（CSV，首列深度，其余列表头为日期）：
    深度,2024-07-30,2024-08-01,...
    0.5,0.669,1.649,...

用法：
    python3 scripts/pick_profile_curves.py tst2.csv
    python3 scripts/pick_profile_curves.py tst2.csv --cap 8 --tolerance 3 --format csv
"""
import argparse
import csv
import json
import sys


def read_matrix(path):
    with open(path, encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    head = [h.strip() for h in rows[0]]
    dates = head[1:]
    data = {}
    for r in rows[1:]:
        if not r or not r[0].strip():
            continue
        try:
            depth = float(r[0])
        except ValueError:
            continue
        vals = {}
        for d, v in zip(dates, r[1:]):
            try:
                vals[d] = float(v)
            except (TypeError, ValueError):
                pass
        if vals:
            data[depth] = vals
    return dates, data


def choose(dates, data, cap=9, min_k=5, tol=None, tol_ratio=0.10, tol_min=2.0,
           dup=1.5):
    dep = sorted(data)
    idx = {d: i for i, d in enumerate(dates)}
    P = {d: [data[z].get(d) for z in dep] for d in dates}
    peak_val = {d: max(abs(x) for x in P[d] if x is not None) for d in dates}
    peak = max(dates, key=lambda d: peak_val[d])
    if tol is None:
        tol = max(tol_min, tol_ratio * peak_val[peak])

    def dev(c, a, b):
        w = (idx[c] - idx[a]) / float(idx[b] - idx[a])
        ds = [abs(z - ((1 - w) * x + w * y))
              for x, y, z in zip(P[a], P[b], P[c])
              if x is not None and y is not None and z is not None]
        return max(ds) if ds else 0.0

    def dist(a, b):
        ds = [abs(x - y) for x, y in zip(P[a], P[b])
              if x is not None and y is not None]
        return max(ds) if ds else 0.0

    sel = sorted({dates[0], peak, dates[-1]}, key=idx.get)
    while len(sel) < cap:
        worst = None
        for a, b in zip(sel, sel[1:]):
            for c in dates[idx[a] + 1:idx[b]]:
                dv = dev(c, a, b)
                if worst is None or dv > worst[0]:
                    worst = (dv, c)
        if not worst or worst[0] <= tol:
            break
        sel = sorted(set(sel) | {worst[1]}, key=idx.get)

    changed = True
    while changed and len(sel) > max(3, min_k - 1):
        changed = False
        for a, b in zip(sel, sel[1:]):
            if (dist(a, b) < dup and b not in (peak, dates[-1])
                    and a != dates[0]):
                sel.remove(a)
                changed = True
                break

    while len(sel) < min(min_k, len(dates)):
        gaps = [(idx[b] - idx[a], a, b) for a, b in zip(sel, sel[1:])]
        _, a, b = max(gaps)
        mid = dates[(idx[a] + idx[b]) // 2]
        if mid in sel:
            break
        sel = sorted(set(sel) | {mid}, key=idx.get)

    # 选完后重新计算最大重构误差，作为质量指标
    err = 0.0
    for a, b in zip(sel, sel[1:]):
        for c in dates[idx[a] + 1:idx[b]]:
            err = max(err, dev(c, a, b))
    return {
        "selected": sel,
        "n_selected": len(sel),
        "n_dates": len(dates),
        "peak_date": peak,
        "peak_mm": round(peak_val[peak], 3),
        "tolerance_mm": round(tol, 3),
        "max_reconstruction_error_mm": round(err, 3),
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description="深度剖面曲线数量自适应选点")
    ap.add_argument("matrix", help="CSV：首列深度，其余列表头为日期")
    ap.add_argument("--cap", type=int, default=9)
    ap.add_argument("--min-k", type=int, default=5)
    ap.add_argument("--tolerance", type=float, default=None,
                    help="固定误差阈值 mm；默认 max(2.0, 峰值×10%%)")
    ap.add_argument("--tol-ratio", type=float, default=0.10)
    ap.add_argument("--tol-min", type=float, default=2.0)
    ap.add_argument("--dup", type=float, default=1.5, help="视觉重复阈值 mm")
    ap.add_argument("--format", choices=["json", "csv", "text"], default="json")
    a = ap.parse_args(argv)

    dates, data = read_matrix(a.matrix)
    if not dates or not data:
        print("!! 矩阵为空或格式不对", file=sys.stderr)
        return 2
    r = choose(dates, data, a.cap, a.min_k, a.tolerance, a.tol_ratio,
               a.tol_min, a.dup)
    if a.format == "json":
        print(json.dumps(r, ensure_ascii=False, indent=2))
    elif a.format == "csv":
        print("date,max_mm")
        for d in r["selected"]:
            mx = max(abs(data[z][d]) for z in data if d in data[z])
            print("%s,%.3f" % (d, mx))
    else:
        print("实测 %d 期 → 选 %d 期  峰值 %.1f mm（%s）"
              % (r["n_dates"], r["n_selected"], r["peak_mm"], r["peak_date"]))
        print("  " + " / ".join(r["selected"]))
        print("  阈值 %.2f mm，最大重构误差 %.2f mm"
              % (r["tolerance_mm"], r["max_reconstruction_error_mm"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
