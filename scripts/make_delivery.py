#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""监测数据「四件套」交付生成器。

输入：
  --summary  汇总工作簿（测项宽表 + 测斜XX 矩阵）
  --groups   xlsx_charts_to_png.py 产出的 _groups.json
  --charts   时程曲线 PNG 目录
  --incl     测斜全过程曲线 PNG 目录
  --sel      选点说明 JSON（[{孔号, 选定日期:[...]}]，可省）
输出（--out 目录）：
  <title>_数据汇总表.xlsx        全部采集数据
  <title>_筛选后数据汇总表.xlsx  按图分块：每块上图 + 绘图数据
  <title>_各点最大累计值汇总表.xlsx
  <title>_曲线图.pdf             A4，时程 4 张/页、测斜 1 张/页，带封面
  <title>_曲线图_PNG.zip         全部 PNG 源文件 + 索引

依赖：openpyxl、Pillow（venv 里装好即可），不依赖 Excel/WPS。
"""
import argparse
import csv
import datetime
import json
import os
import re
import zipfile

import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image, ImageDraw, ImageFont

HDR_FILL = PatternFill('solid', fgColor='DDEBF7')
THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

UNIT_BY_KEY = [('倾斜', '‰'), ('轴力', 'kN'), ('水位', 'mm'), ('沉降', 'mm'),
               ('位移', 'mm'), ('测斜', 'mm')]


def natkey(s):
    return [int(t) if t.isdigit() else str(t).lower()
            for t in re.split(r'(\d+)', str(s))]


def unit_of(name):
    for k, u in UNIT_BY_KEY:
        if k in name:
            return u
    return ''


def style_header(ws, row, ncol):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(bold=True, size=9)
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = BORDER


def autosize(ws, first=13, width=10.5):
    ws.column_dimensions['A'].width = first
    for c in range(2, ws.max_column + 1):
        ws.column_dimensions[get_column_letter(c)].width = width


def load_summary(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


# ---------------------------------------------------------------- 1) 数据汇总表
def build_summary_xlsx(data, out, items):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in data.items():
        ws = wb.create_sheet(name[:31])
        body = rows
        if name in items and len(rows) > 2:
            body = rows[:2] + sorted(rows[2:], key=lambda r: natkey(r[0] if r else ''))
        for r in body:
            ws.append(r)
        if rows:
            ws['A1'].font = Font(bold=True, size=12)
            if len(rows) > 1:
                style_header(ws, 2, max(len(x) for x in rows[:2]))
            ws.freeze_panes = 'B3'
            autosize(ws)
    wb.save(out)


# ---------------------------------------------------------------- 2) 筛选后（按图分块）
def build_plotdata(groups, data, out, charts_dir, incl_dir, sel, title):
    wb = openpyxl.Workbook()
    ws0 = wb.active
    ws0.title = '说明'
    notes = [
        ('%s 绘图数据表（筛选后）' % title, None),
        ('用途：按曲线图分块，每块上方是该图的 PNG 原件，下方是画这张图用的数据。', None),
        ('分块规则：先按测点前缀分类，类内自然排序（1、2、3…10），每 5 点一张图。', None),
        ('日期为行、测点为列；只保留曲线图实际使用的观测日期。', None),
        ('单位：沉降/位移/水位 mm；倾斜 ‰；轴力 kN。', None),
        ('', None),
        ('孔号', '选定日期（用于绘图）'),
    ]
    for r in notes:
        ws0.append(list(r))
    for h, d in (sel or {}).items():
        ws0.append([h, ' / '.join(d)])
    ws0['A1'].font = Font(bold=True, size=13)
    hdr = len(notes)
    ws0.cell(row=hdr, column=1).font = Font(bold=True)
    ws0.cell(row=hdr, column=2).font = Font(bold=True)
    ws0.column_dimensions['A'].width = 12
    ws0.column_dimensions['B'].width = 95

    by_sheet = {}
    for g in groups:
        by_sheet.setdefault(g['sheet'].replace('图1 ', ''), []).append(g)
    for sheet, blocks in by_sheet.items():
        ws = wb.create_sheet(sheet[:31])
        r = 1
        for b in blocks:
            ws.cell(row=r, column=1, value=b['title']).font = Font(bold=True, size=11)
            r += 1
            p = os.path.join(charts_dir, b['file'])
            if os.path.exists(p):
                img = XLImage(p)
                img.width = int(186 / 25.4 * 96)
                img.height = int(56 / 25.4 * 96)
                ws.add_image(img, 'A%d' % r)
            r += 12
            for j, v in enumerate(['观测日期'] + list(b['points']), 1):
                c = ws.cell(row=r, column=j, value=v)
                c.font = Font(bold=True, size=9)
                c.fill = HDR_FILL
                c.alignment = Alignment(horizontal='center')
                c.border = BORDER
            r += 1
            for i, d in enumerate(b['dates']):
                ws.cell(row=r, column=1, value=d).font = Font(size=9)
                for j, name in enumerate(b['points'], 2):
                    vals = b['series'].get(name, [])
                    v = vals[i] if i < len(vals) else None
                    if v in (None, ''):
                        continue
                    try:
                        fv = float(v)
                    except ValueError:
                        continue
                    cell = ws.cell(row=r, column=j, value=round(fv, 3))
                    cell.number_format = '0.000'
                    cell.font = Font(size=9)
                r += 1
            r += 2
        ws.column_dimensions['A'].width = 13
        for c in range(2, 8):
            ws.column_dimensions[get_column_letter(c)].width = 10
        ws.freeze_panes = 'B3'

    # 测斜
    for k, v in data.items():
        if not k.startswith('测斜'):
            continue
        rows = v
        dates = [str(x)[:10] if x else '' for x in rows[1][1:]]
        keep = [i for i, d in enumerate(dates) if d in (sel or {}).get(k[2:], [])]
        if not keep:
            keep = list(range(0, len(dates), max(1, len(dates) // 6)))[:6]
        ws = wb.create_sheet(k[:31])
        ws.cell(row=1, column=1, value='%s %s 深层水平位移' % (title, k[2:])).font = Font(bold=True, size=12)
        png = find_incl_png(incl_dir, k[2:])
        if os.path.exists(png):
            im = XLImage(png)
            im.width = int(90 / 25.4 * 96)
            im.height = int(210 / 25.4 * 96)
            ws.add_image(im, 'A2')
        r = 43
        ws.cell(row=r, column=1, value='深度(m)').font = Font(bold=True, size=9)
        for j, i in enumerate(keep, start=2):
            c = ws.cell(row=r, column=j, value=dates[i])
            c.font = Font(bold=True, size=9)
            c.fill = HDR_FILL
        r += 1
        for row in rows[2:]:
            if row[0] is None:
                continue
            ws.cell(row=r, column=1, value=row[0])
            for j, i in enumerate(keep, start=2):
                val = row[i + 1] if i + 1 < len(row) else None
                if isinstance(val, (int, float)):
                    cell = ws.cell(row=r, column=j, value=round(val, 3))
                    cell.number_format = '0.000'
            r += 1
        ws.column_dimensions['A'].width = 12
        for c in range(2, ws.max_column + 1):
            ws.column_dimensions[get_column_letter(c)].width = 11
    wb.save(out)


# ---------------------------------------------------------------- 3) 各点最大累计值
def build_max(data, out, items):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '各点最大累计值'
    head = ['序号', '测项', '点号', '单位', '最大累计值', '绝对最大值', '出现日期', '数据期数']
    ws.append(head)
    style_header(ws, 1, len(head))
    rows_out = []
    for item, unit in items.items():
        if item not in data:
            continue
        rows = data[item]
        dates = [str(x)[:10] if x else '' for x in rows[1][1:]]
        for r in rows[2:]:
            name = r[0] if r else None
            if not name:
                continue
            best, cnt = None, 0
            for j, v in enumerate(r[1:]):
                if not isinstance(v, (int, float)):
                    continue
                cnt += 1
                if best is None or abs(v) > abs(best[0]):
                    best = (v, dates[j] if j < len(dates) else '')
            if best:
                rows_out.append([item, str(name), unit, round(best[0], 3),
                                 round(abs(best[0]), 3), best[1], cnt])
    order = list(items)
    rows_out.sort(key=lambda r: (order.index(r[0]), natkey(r[1])))
    for i, r in enumerate(rows_out, 1):
        ws.append([i] + r)

    ws2 = wb.create_sheet('测斜孔最大累计值')
    head2 = ['序号', '孔号', '单位', '最大累计位移', '绝对最大值', '出现日期', '出现深度(m)', '数据期数']
    ws2.append(head2)
    style_header(ws2, 1, len(head2))
    n = 0
    for k, rows in data.items():
        if not k.startswith('测斜'):
            continue
        dates = [str(x)[:10] if x else '' for x in rows[1][1:]]
        best = None
        cnt = 0
        for i in range(1, len(rows[1])):
            col_vals = [r[i] for r in rows[2:] if i < len(r) and isinstance(r[i], (int, float))]
            if col_vals:
                cnt += 1
            for r in rows[2:]:
                v = r[i] if i < len(r) else None
                if not isinstance(v, (int, float)):
                    continue
                if best is None or abs(v) > abs(best[0]):
                    best = (v, dates[i - 1] if i - 1 < len(dates) else '', r[0])
        if best:
            n += 1
            ws2.append([n, k[2:], 'mm', round(best[0], 3), round(abs(best[0]), 3),
                        best[1], best[2], cnt])
    for sheet in (ws, ws2):
        autosize(sheet, first=8, width=15)
        sheet.freeze_panes = 'A2'
        for row in sheet.iter_rows(min_row=2):
            for c in row:
                c.border = BORDER
    wb.save(out)


def incl_pngs(incl_dir):
    """测斜 PNG：文件名须含「全过程」。目录不存在时返回空表而不是抛异常。"""
    if not incl_dir or not os.path.isdir(incl_dir):
        return []
    return sorted(n for n in os.listdir(incl_dir)
                  if n.endswith('.png') and '全过程' in n)


def find_incl_png(incl_dir, hole):
    """按孔号找测斜 PNG；不依赖文件名模板，优先「全过程」。找不到返回空串。"""
    names = incl_pngs(incl_dir)
    for n in names:
        if hole in n and '全过程' in n:
            return os.path.join(incl_dir, n)
    for n in names:
        if hole in n:
            return os.path.join(incl_dir, n)
    return ''


# ---------------------------------------------------------------- 4) PDF
def build_pdf(groups, charts_dir, incl_dir, out, title, dpi=220):
    mm = dpi / 25.4
    PW, PH = int(210 * mm), int(297 * mm)
    font_path = '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc'

    def font(size):
        try:
            return ImageFont.truetype(font_path, size)
        except Exception:
            return None

    files = [g['file'] for g in groups]
    pages = []
    cover = Image.new('RGB', (PW, PH), 'white')
    d = ImageDraw.Draw(cover)
    d.text((int(PW * 0.1), int(PH * 0.28)), title, fill='#1F1F1F', font=font(64))
    d.text((int(PW * 0.1), int(PH * 0.28) + 100), '监测曲线图', fill='#1F1F1F', font=font(64))
    n_incl = len(incl_pngs(incl_dir))
    d.text((int(PW * 0.1), int(PH * 0.28) + 230),
           '时程曲线 %d 张 · 测斜全过程曲线 %d 张' % (len(files), n_incl),
           fill='#555555', font=font(34))
    pages.append(cover)

    for i in range(0, len(files), 4):
        page = Image.new('RGB', (PW, PH), 'white')
        y = int(18 * mm)
        for name in files[i:i + 4]:
            p = os.path.join(charts_dir, name)
            if not os.path.exists(p):
                continue
            im = Image.open(p)
            page.paste(im, (int((PW - im.size[0]) / 2), y))
            y += im.size[1] + int(11 * mm)
        pages.append(page)

    for name in incl_pngs(incl_dir):
        im = Image.open(os.path.join(incl_dir, name))
        w, h = int(90 * mm), int(210 * mm)
        page = Image.new('RGB', (PW, PH), 'white')
        page.paste(im.resize((w, h), Image.LANCZOS), (int((PW - w) / 2), int((PH - h) / 2)))
        pages.append(page)

    pages[0].save(out, save_all=True, append_images=pages[1:], resolution=dpi)
    return len(pages)


# ---------------------------------------------------------------- 5) PNG 包
def build_zip(groups, charts_dir, incl_dir, out, title):
    rows = []
    for i, g in enumerate(groups, 1):
        dst = '%02d_%s' % (i, g['file'])
        rows.append([i, '时程曲线', g['title'], dst, '186x56', os.path.join(charts_dir, g['file'])])
    for name in incl_pngs(incl_dir):
        rows.append([len(rows) + 1, '测斜全过程曲线', name[:-4], '%02d_%s' % (len(rows) + 1, name),
                     '90x210', os.path.join(incl_dir, name)])
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for r in rows:
            if os.path.exists(r[5]):
                z.write(r[5], r[3])
        buf = ['序号,类型,图名,文件名,图幅mm'] + [','.join(str(x) for x in r[:5]) for r in rows]
        z.writestr('_索引.csv', ('\ufeff' + '\n'.join(buf)).encode('utf-8'))
    return len(rows)


def main():
    ap = argparse.ArgumentParser(description='监测数据四件套交付')
    ap.add_argument('--summary', required=True)
    ap.add_argument('--groups', required=True, help='_groups.json')
    ap.add_argument('--charts', required=True, help='时程 PNG 目录')
    ap.add_argument('--incl', default=None,
                    help='测斜 PNG 目录（文件名须含“全过程”）；缺省自动找 <charts>/incl')
    ap.add_argument('--sel', default=None, help='选点说明 JSON')
    ap.add_argument('--units', default=None, help='{"工作表名":"单位"} 覆盖自动判断')
    ap.add_argument('--title', default='监测项目')
    ap.add_argument('--out', required=True)
    a = ap.parse_args()

    os.makedirs(a.out, exist_ok=True)
    data = load_summary(a.summary)
    groups = json.load(open(a.groups, encoding='utf-8'))
    if not isinstance(groups, list):
        raise SystemExit('--groups 必须是 _groups.json 的数组；收到 %s' % type(groups).__name__)

    charts_dir = a.charts
    incl_dir = a.incl
    if not incl_dir:
        cand = os.path.join(charts_dir, 'incl')
        incl_dir = cand if os.path.isdir(cand) else charts_dir
    if not os.path.isdir(incl_dir):
        print('WARN 测斜目录不存在：%s（PDF/PNG 包将不含测斜全过程曲线）' % incl_dir)
    if not incl_pngs(incl_dir):
        print('WARN %s 下没有文件名含“全过程”的 PNG，测斜曲线不会进交付件' % incl_dir)
    if not groups:
        print('WARN _groups.json 为空：筛选后汇总表将没有时程图块')
    sel_raw = json.load(open(a.sel, encoding='utf-8')) if a.sel else []
    sel = {r['孔号']: r['选定日期'] for r in sel_raw} if sel_raw else {}

    items = {}
    for k, rows in data.items():
        if k.startswith('测斜') or len(rows) < 3:
            continue
        # 测项宽表：第 2 行是「点号 + 日期」，其余是说明/明细/检查表
        if not (str(rows[1][0]).strip() == '点号'
                and re.match(r'\d{4}-\d{2}-\d{2}', str(rows[1][1]))):
            continue
        items[k] = unit_of(k)
    if a.units:
        items.update(json.load(open(a.units, encoding='utf-8')))
    if not items:
        raise SystemExit('汇总表里没有识别到测项宽表。要求：第 2 行 = “点号” + 首个日期(YYYY-MM-DD)，'
                         '第 3 行起为数据；详见 references/agent-runbook.md')

    n_incl = len(incl_pngs(incl_dir))
    print('识别: 测项 %d 个 / 时程图 %d 张 / 测斜图 %d 张' % (len(items), len(groups), n_incl))

    p_sum = os.path.join(a.out, '%s_数据汇总表.xlsx' % a.title)
    p_plot = os.path.join(a.out, '%s_筛选后数据汇总表.xlsx' % a.title)
    p_max = os.path.join(a.out, '%s_各点最大累计值汇总表.xlsx' % a.title)
    p_pdf = os.path.join(a.out, '%s_曲线图.pdf' % a.title)
    p_zip = os.path.join(a.out, '%s_曲线图_PNG.zip' % a.title)

    build_summary_xlsx(data, p_sum, items)
    build_plotdata(groups, data, p_plot, charts_dir, incl_dir, sel, a.title)
    build_max(data, p_max, items)
    n_pages = build_pdf(groups, charts_dir, incl_dir, p_pdf, a.title)
    n_png = build_zip(groups, charts_dir, incl_dir, p_zip, a.title)
    print('数据汇总表 %s' % p_sum)
    print('筛选后汇总 %s' % p_plot)
    print('最大累计值 %s' % p_max)
    print('曲线图 PDF %s（%d 页）' % (p_pdf, n_pages))
    print('曲线图 PNG %s（%d 张）' % (p_zip, n_png))


if __name__ == '__main__':
    main()
