#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""长表 CSV → 宽表 CSV（实体 × 日期），零第三方依赖。

输入：长表 CSV，至少三列（列名可配置）：
    entity,date,value[,source]
    点号,2024-08-01,12.5,日报1.xlsx

输出：
    wide.csv        行=实体，列=日期，值=度量
    conflicts.csv   同一 实体+日期 出现多次时，后写覆盖先写，并在这里留痕

用法：
    python3 long_to_wide.py long.csv wide.csv \
        --entity entity --date date --value value [--source source] [--precision 3]

设计要点（见 SKILL.md）：
  - 冲突不静默：每一次覆盖都写进 conflicts.csv；
  - 日期统一成 YYYY-MM-DD 字符串，避免 Excel serial / 时区串列；
  - 缺失就是空，不补 0；
  - 确定性输出：列按日期排序，行按实体排序。
"""
import argparse
import csv
import datetime
import sys


def norm_date(s):
    s = (s or "").strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d"):
        try:
            return datetime.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    # Excel serial（1900 起算，含闰年 bug）
    try:
        n = float(s)
        if 20000 < n < 80000:
            return (datetime.date(1899, 12, 30) + datetime.timedelta(days=int(n))).strftime("%Y-%m-%d")
    except ValueError:
        pass
    raise ValueError("无法识别的日期: %r" % s)


def norm_key(s):
    # 去首尾空白、换行、全角空格；不改变大小写（ZQT1 与 zqt1 可能是两个实体）
    return (s or "").replace("\u3000", " ").strip().replace("\n", "").replace("\r", "")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--entity", default="entity")
    ap.add_argument("--date", default="date")
    ap.add_argument("--value", default="value")
    ap.add_argument("--source", default="source")
    ap.add_argument("--precision", type=int, default=3)
    args = ap.parse_args()

    wide, conflicts, src_of = {}, [], {}
    with open(args.src, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        need = {args.entity, args.date, args.value}
        if not need.issubset(reader.fieldnames or []):
            sys.exit("缺少列：%s，实际表头：%s" % (sorted(need), reader.fieldnames))
        for row in reader:
            e = norm_key(row[args.entity])
            if not e:
                continue
            d = norm_date(row[args.date])
            raw = (row[args.value] or "").strip()
            if raw == "":
                continue                      # 没测 ≠ 0
            try:
                v = float(raw)
            except ValueError:
                conflicts.append([e, d, row.get(args.value), "", "非数值，已跳过"])
                continue
            cell = wide.setdefault(e, {})
            if d in cell:
                conflicts.append([e, d, cell[d], v, src_of.get((e, d), "")])
            cell[d] = v
            src_of[(e, d)] = row.get(args.source, "")

    dates = sorted({d for cell in wide.values() for d in cell})
    entities = sorted(wide)
    fmt = "%." + str(args.precision) + "f"

    with open(args.out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow([args.entity] + dates)
        for e in entities:
            w.writerow([e] + [fmt % wide[e][d] if d in wide[e] else "" for d in dates])

    cpath = args.out.rsplit(".", 1)[0] + "_conflicts.csv"
    with open(cpath, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["实体", "日期", "原值", "新值", "来源", "备注"])
        w.writerows(conflicts)

    print("实体 %d × 日期 %d，冲突 %d 条 → %s" % (len(entities), len(dates), len(conflicts), args.out))
    if conflicts:
        print("冲突明细见 %s（后写覆盖先写，已留痕）" % cpath)


if __name__ == "__main__":
    main()
