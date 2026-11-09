#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""给工作簿里的「测斜XX」矩阵表批量加联动选点页。

每个孔新增一张「选点XX」：
- H2 起是下拉选点日期（数据验证，来源＝该孔表头日期行）；
- 下方 80 行深度用 INDEX/MATCH 自动取对应日期的剖面（改日期即取新数据）；
- 右侧散点图绑定这些公式单元格，改日期/改数即重绘。

用法：
    python3 xlsx_inclinometer_link.py 曲线图.xlsx 输出.xlsx
    python3 xlsx_inclinometer_link.py 曲线图.xlsx 输出.xlsx --prefix 测斜 --slots 6
"""
import argparse
import os
import re
import sys

import openpyxl
from openpyxl.chart import Reference, ScatterChart, Series
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import SeriesLabel
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HDR = PatternFill('solid', fgColor='DDEBF7')
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
COLORS = ['0072B2', 'E69F00', '009E73', 'CC79A7', '56B4E9', 'D55E4F',
          '7B5EA7', '8C564B', '1F1F1F']
MARKERS = ['circle', 'square', 'triangle', 'diamond', 'x', 'star',
           'dot', 'plus', 'dash']


def sheet_layout(ws):
    """定位矩阵：表头行（含日期）、深度列、首末数据行。"""
    hdr_row = depth_col = None
    for r in range(1, 6):
        vals = [ws.cell(row=r, column=c).value for c in range(1, 4)]
        if any(isinstance(v, str) and re.search('深度', v) for v in vals):
            hdr_row = r
            depth_col = next(c for c, v in enumerate(vals, 1)
                             if isinstance(v, str) and '深度' in v)
            break
    if hdr_row is None:
        return None
    first = hdr_row + 1
    last = first
    while ws.cell(row=last + 1, column=depth_col).value is not None:
        last += 1
    return hdr_row, depth_col, first, last


def add_link(ws_src, wb, slots):
    name = ws_src.title
    tag = re.sub(r'^测斜', '', name) or name
    lay = sheet_layout(ws_src)
    if not lay:
        return None
    hdr_row, dcol, first, last = lay
    # 日期列：表头行里第一个日期到最后
    dcol0 = dcol + 1
    dcol1 = ws_src.max_column
    dL = get_column_letter(dcol)
    d0 = get_column_letter(dcol0)
    d1 = get_column_letter(dcol1)

    ws = wb.create_sheet(('选点' + tag)[:31])
    ws['A1'] = '%s 联动：改 H2 起的日期，数据与曲线自动更新' % tag
    ws['A1'].font = Font(bold=True, size=12)
    ws['A2'] = '选点日期'
    ws['A3'] = '深度'
    for c in ('A2', 'A3'):
        ws[c].font = Font(bold=True, size=9)

    for j in range(slots):
        L = get_column_letter(8 + j)
        c = ws.cell(row=2, column=8 + j)
        c.fill = HDR
        c.alignment = Alignment(horizontal='center')
        h = ws.cell(row=3, column=8 + j, value='=%s2' % L)
        h.font = Font(bold=True, size=9)
        h.alignment = Alignment(horizontal='center')
        h.border = BORDER
        dv = DataValidation(type='list',
                            formula1="'%s'!$%s$%d:$%s$%d" % (name, d0, hdr_row, d1, hdr_row),
                            allow_blank=True)
        dv.errorTitle = '日期不在总表内'
        dv.error = '请从该孔已有的观测日期里选'
        ws.add_data_validation(dv)
        dv.add('%s2' % L)

    q = chr(34)
    for i, row in enumerate(range(4, 4 + (last - first) + 1)):
        dcell = ws.cell(row=row, column=dcol, value=ws_src.cell(row=first + i, column=dcol).value)
        dcell.number_format = '0.0'
        for j in range(slots):
            L = get_column_letter(8 + j)
            tpl = ("=IFERROR(INDEX('{s}'!${d0}${r0}:${d1}${r1},MATCH(${dl}{row},"
                   "'{s}'!${dl}${r0}:${dl}${r1},0),MATCH({L}$3,"
                   "'{s}'!${d0}${h}:${d1}${h},0))," + q + q + ")")
            f = tpl.format(s=name, d0=d0, d1=d1, dl=dL, r0=first, r1=last, h=hdr_row,
                           row=row, L=L)
            cell = ws.cell(row=row, column=8 + j, value=f)
            cell.number_format = '0.000'

    ws.column_dimensions['A'].width = 9
    for j in range(slots):
        ws.column_dimensions[get_column_letter(8 + j)].width = 10
    ws.freeze_panes = 'B4'

    ch = ScatterChart()
    ch.title = '%s 深层水平位移' % tag
    ch.x_axis.title = '位移（mm）'
    ch.y_axis.title = '深度（m）'
    ch.y_axis.scaling.orientation = 'maxMin'
    ch.x_axis.tickLblPos = 'high'
    ch.height, ch.width = 21, 9.2
    ch.legend.position = 'r'
    for j in range(slots):
        L = get_column_letter(8 + j)
        s = Series(Reference(ws, min_col=dcol, min_row=4, max_row=4 + (last - first)),
                   Reference(ws, min_col=8 + j, min_row=4, max_row=4 + (last - first)),
                   title_from_data=False)
        s.tx = SeriesLabel(strRef=StrRef("'%s'!$%s$3" % (ws.title, L)))
        s.marker = Marker(symbol=MARKERS[j % len(MARKERS)], size=4)
        s.graphicalProperties.line.solidFill = COLORS[j % len(COLORS)]
        s.graphicalProperties.line.width = 15000
        ch.series.append(s)
    ws.add_chart(ch, 'O2')
    return ws.title


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('src')
    ap.add_argument('out')
    ap.add_argument('--prefix', default='测斜')
    ap.add_argument('--slots', type=int, default=6, help='选点槽位数（3~9）')
    a = ap.parse_args()
    slots = max(3, min(9, a.slots))

    wb = openpyxl.load_workbook(a.src)
    targets = [ws for ws in wb.worksheets if ws.title.startswith(a.prefix)]
    if not targets:
        sys.exit('没找到以「%s」开头的工作表' % a.prefix)
    made = []
    for ws in targets:
        t = add_link(ws, wb, slots)
        if t:
            made.append(t)
        # 默认日期：首测 / 末测 / 中间各取一个
        lay = sheet_layout(ws)
        if lay:
            h, _, _, _ = lay
            ws2 = wb[t]
            dates = [ws.cell(row=h, column=c).value for c in range(2, ws.max_column + 1)]
            dates = [d for d in dates if d]
            picks = [dates[0], dates[-1]] + ([dates[len(dates) // 2]] if dates else [])
            for j in range(slots):
                ws2.cell(row=2, column=8 + j, value=picks[j % len(picks)])
    wb.calculation.fullCalcOnLoad = True
    wb.save(a.out)
    print('生成 %d 张选点表 → %s' % (len(made), a.out))


if __name__ == '__main__':
    main()
