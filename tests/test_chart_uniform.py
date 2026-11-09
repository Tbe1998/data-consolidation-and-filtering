#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""xlsx_chart_uniform.py 的单测：用最小 xlsx 夹具验证「套版式但不改数据」。

夹具自己拼 zip，不依赖 openpyxl，CI 上零安装可跑。
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SCRIPT = os.path.join(ROOT, "scripts", "xlsx_chart_uniform.py")


def _chart(title, layout, amin, amax, aunit, xmin=0.1, xmax=9.9,
           legend_entry=False, series_name="45503"):
    le = "<c:legendEntry><c:idx val=\"0\"/><c:delete val=\"1\"/></c:legendEntry>" if legend_entry else ""
    return (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:rPr sz="1600" b="1"/>'
        '<a:t>' + title + '</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>'
        '<c:plotArea><c:layout>' + layout + '</c:layout><c:scatterChart>'
        '<c:scatterStyle val="lineMarker"/><c:ser><c:idx val="0"/><c:order val="0"/>'
        '<c:tx><c:strRef><c:f>SHEET!$B$2</c:f><c:strCache><c:ptCount val="1"/>'
        '<c:pt idx="0"><c:v>' + series_name + '</c:v></c:pt></c:strCache></c:strRef></c:tx>'
        '<c:spPr><a:ln><a:solidFill><a:srgbClr val="3366CC"/></a:solidFill></a:ln></c:spPr>'
        '<c:marker><c:symbol val="circle"/></c:marker>'
        '<c:xVal><c:numRef><c:f>SHEET!$B$3:$B$4</c:f><c:numCache><c:ptCount val="2"/>'
        '<c:pt idx="0"><c:v>' + str(xmin) + '</c:v></c:pt>'
        '<c:pt idx="1"><c:v>' + str(xmax) + '</c:v></c:pt></c:numCache></c:numRef></c:xVal>'
        '<c:yVal><c:numRef><c:f>SHEET!$C$3:$C$4</c:f><c:numCache><c:ptCount val="2"/>'
        '<c:pt idx="0"><c:v>0.5</c:v></c:pt><c:pt idx="1"><c:v>40</c:v></c:pt>'
        '</c:numCache></c:numRef></c:yVal></c:ser></c:scatterChart>'
        '<c:valAx><c:axId val="11"/><c:crossAx val="22"/><c:axPos val="t"/>'
        '<c:scaling><c:orientation val="minMax"/><c:min val="' + str(amin) + '"/>'
        '<c:max val="' + str(amax) + '"/><c:majorUnit val="' + str(aunit) + '"/></c:scaling>'
        '<c:title><c:tx><c:rich><a:p><a:r><a:t>位移（mm）</a:t></a:r></a:p></c:rich></c:tx></c:title>'
        '</c:valAx>'
        '<c:valAx><c:axId val="22"/><c:crossAx val="11"/><c:axPos val="l"/>'
        '<c:scaling><c:orientation val="maxMin"/><c:min val="0.5"/><c:max val="40"/>'
        '<c:majorUnit val="2"/></c:scaling>'
        '<c:title><c:tx><c:rich><a:p><a:r><a:t>深度（m）</a:t></a:r></a:p></c:rich></c:tx></c:title>'
        '</c:valAx></c:plotArea>'
        '<c:legend><c:legendPos val="r"/>' + le +
        '<c:layout><c:manualLayout><c:x val="0.7"/><c:y val="0.8"/></c:manualLayout></c:layout>'
        '</c:legend></c:chart></c:chartSpace>')


def build_book(path):
    """两个工作表：图1 各 2 张图（第二张需被统一）；图2 两页各 1 张（测图例日期）。"""
    charts = {
        # 基准：布局 0.1/0.2/0.7/0.6，轴 -10~20 步长 5
        "chart1.xml": _chart("A井 地表沉降 时程曲线", "<c:manualLayout><c:x val=\"0.1\"/>"
                             "<c:y val=\"0.2\"/><c:w val=\"0.7\"/><c:h val=\"0.6\"/>"
                             "</c:manualLayout>", -10, 20, 5),
        # 目标：自有布局、单位、n/m 序号、轴范围差很多（应被基准覆盖）
        "chart2.xml": _chart("A井 地表沉降 时程曲线（mm）2/3",
                             "<c:manualLayout><c:x val=\"0.2\"/><c:y val=\"0.1\"/>"
                             "<c:w val=\"0.5\"/><c:h val=\"0.5\"/></c:manualLayout>",
                             -40, 60, 20, legend_entry=True),
        "chart3.xml": _chart("A井 TST1 位移-深度曲线", "<c:manualLayout><c:x val=\"0.1\"/>"
                             "<c:y val=\"0.2\"/><c:w val=\"0.7\"/><c:h val=\"0.6\"/>"
                             "</c:manualLayout>", -10, 30, 5, series_name="45503"),
        "chart4.xml": _chart("A井 TST2 位移-深度曲线", "<c:manualLayout><c:x val=\"0.3\"/>"
                             "<c:y val=\"0.3\"/><c:w val=\"0.4\"/><c:h val=\"0.4\"/>"
                             "</c:manualLayout>", -40, 60, 20, series_name="45561"),
    }
    sheets = {
        "sheet1.xml": ("图1 地表沉降", ["chart1.xml", "chart2.xml"]),
        "sheet2.xml": ("图2 测斜TST1", ["chart3.xml"]),
        "sheet3.xml": ("图2 测斜TST2", ["chart4.xml"]),
    }
    ct = ('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="xml" ContentType="application/xml"/></Types>')
    wb = ('<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
          ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
          + "".join('<sheet name="%s" sheetId="%d" r:id="rId%d"/>' % (n, i + 1, i + 1)
                    for i, (n, _) in enumerate(sheets.values()))
          + "</sheets></workbook>")
    wbrels = ('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              + "".join('<Relationship Id="rId%d" Type="worksheet" Target="worksheets/sheet%d.xml"/>'
                        % (i + 1, i + 1) for i in range(len(sheets)))
              + "</Relationships>")
    z = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    z.writestr("[Content_Types].xml", ct)
    z.writestr("xl/workbook.xml", wb)
    z.writestr("xl/_rels/workbook.xml.rels", wbrels)
    z.writestr("xl/sharedStrings.xml",
               '<?xml version="1.0"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
               '<si><t>A井 地表沉降 时程曲线（mm） 1/4 [数据页:地表沉降]</t></si></sst>')
    for i, (fname, (name, chs)) in enumerate(sheets.items(), 1):
        rows = ('<row r="2"><c r="B2" s="1"><v>45503</v></c><c r="C2" s="1"><v>45561</v></c></row>')
        z.writestr("xl/worksheets/" + fname,
                   '<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
                   ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                   '<sheetData>' + rows + '</sheetData><drawing r:id="rId1"/></worksheet>')
        z.writestr("xl/worksheets/_rels/sheet%d.xml.rels" % i,
                   '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Type="drawing" Target="../drawings/drawing%d.xml"/>'
                   '</Relationships>' % i)
        z.writestr("xl/drawings/drawing%d.xml" % i,
                   '<?xml version="1.0"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
                   ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
                   ' xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">'
                   + "".join('<xdr:graphicFrame><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                             '<a:graphicData><c:chart r:id="rId%d"/></a:graphicData></a:graphic></xdr:graphicFrame>'
                             % (k + 1) for k in range(len(chs)))
                   + "</xdr:wsDr>")
        z.writestr("xl/drawings/_rels/drawing%d.xml.rels" % i,
                   '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   + "".join('<Relationship Id="rId%d" Type="chart" Target="../charts/%s"/>'
                             % (k + 1, c) for k, c in enumerate(chs))
                   + "</Relationships>")
    for cname, body in charts.items():
        z.writestr("xl/charts/" + cname, body)
    z.close()


class TestChartUniform(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.src = os.path.join(self.tmp, "in.xlsx")
        self.out = os.path.join(self.tmp, "out.xlsx")
        self.cfg = os.path.join(self.tmp, "cfg.json")
        build_book(self.src)
        with open(self.cfg, "w", encoding="utf-8") as f:
            json.dump({
                "groups": [
                    {"name": "地表", "sheets": ["图1 地表沉降"], "ref": 0},
                    {"name": "测斜", "sheets": ["图2 测斜TST1", "图2 测斜TST2"], "ref": 0,
                     "rename": {"位移-深度曲线": "时程曲线图"},
                     "legend_dates": {"row": 2, "from_col": 2, "count": 2}},
                ],
                "clean_strings": {"contains": ["[数据页"], "strip_units": True,
                                   "strip_index": True},
            }, f, ensure_ascii=False)
        r = subprocess.run([sys.executable, SCRIPT, self.src, self.out, "--config", self.cfg],
                           capture_output=True, text=True)
        self.assertEqual(r.returncode, 0, r.stdout + r.stderr)
        self.zip = zipfile.ZipFile(self.out)

    def chart(self, name):
        return self.zip.read("xl/charts/" + name).decode()

    def test_title_strips_unit_and_index(self):
        c2 = self.chart("chart2.xml")
        self.assertEqual(re.findall(r"<a:t>([^<]*)</a:t>", c2)[0], "A井 地表沉降 时程曲线")

    def test_rename_applied(self):
        c4 = self.chart("chart4.xml")
        self.assertEqual(re.findall(r"<a:t>([^<]*)</a:t>", c4)[0], "A井 TST2 时程曲线图")

    def test_layout_copied(self):
        c1, c2 = self.chart("chart1.xml"), self.chart("chart2.xml")

        def lay(s):
            return re.search(r"<c:layout>.*?</c:layout>", s, re.S).group(0)
        self.assertEqual(lay(c2), lay(c1))

    def test_axis_copied_and_expanded(self):
        c2 = self.chart("chart2.xml")
        ax = re.search(r"<c:valAx>.*?</c:valAx>", c2, re.S).group(0)
        self.assertEqual(re.findall(r'<c:(?:min|max|majorUnit) val="([-\d.]+)"', ax),
                         ["-10", "20", "5"])

    def test_legend_entry_dropped(self):
        c2 = self.chart("chart2.xml")
        leg = re.search(r"<c:legend>.*?</c:legend>", c2, re.S).group(0)
        self.assertNotIn("legendEntry", leg)

    def test_data_unchanged(self):
        zin = zipfile.ZipFile(self.src)
        for n in ("chart1.xml", "chart2.xml", "chart3.xml", "chart4.xml"):
            a = re.findall(r"<c:numCache>.*?</c:numCache>",
                           zin.read("xl/charts/" + n).decode(), re.S)
            b = re.findall(r"<c:numCache>.*?</c:numCache>", self.chart(n), re.S)
            self.assertEqual(a, b, n)

    def test_legend_dates_from_header(self):
        c4 = self.chart("chart4.xml")
        v = re.findall(r"<c:tx><c:strRef>.*?<c:v>([^<]*)</c:v>", c4, re.S)
        self.assertEqual(v[0], "2024-07-30")
        self.assertIn("$B$2", c4)

    def test_shared_strings_cleaned(self):
        s = self.zip.read("xl/sharedStrings.xml").decode()
        self.assertIn("A井 地表沉降 时程曲线 [数据页:地表沉降]", s)


if __name__ == "__main__":
    unittest.main()
