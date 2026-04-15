#!/usr/bin/env python3
"""重新生成常州招标 Excel（修复版）"""
import urllib.request, ssl, re, io, zipfile, os
from datetime import datetime

BASE = 'http://ggzy.xzsp.changzhou.gov.cn'
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

req = urllib.request.Request(
    BASE + '/jyzx/001001/tradeInfonew.html?category=001001',
    headers={'User-Agent': 'Mozilla/5.0'}
)
html = urllib.request.urlopen(req, timeout=15, context=ctx).read().decode('utf-8', 'ignore')

REL = ['市政', '道路', '公路', '桥梁', '排水', '照明', '绿化', '交通', '施工', '总承包', '新建', '改建']
sc = lambda t: sum(1 for k in REL if k in t)

items = []
seen = set()
for m in re.finditer(r'<tr[^>]*>([\s\S]*?)</tr>', html):
    row = m.group(1)
    segs = row.split('</td>')
    if len(segs) < 4:
        continue
    om = re.search(r"tzjydetail\('([^']+)',\s*'([^']+)'\s*,", segs[1])
    if not om or om.group(2) in segs[0]:
        continue
    tm = re.search(r'title="([^"]{5,120})"', segs[1])
    if not tm:
        continue
    uuid = om.group(2)
    if uuid in seen:
        continue
    seen.add(uuid)
    area = re.sub('<[^>]+>', '', segs[2]).replace('\xa0', ' ').strip()
    date = re.sub('<[^>]+>', '', segs[3]).replace('\xa0', ' ').strip()
    items.append({
        'cat': om.group(1), 'uuid': uuid,
        'title': tm.group(1)[:120], 'area': area, 'date': date
    })

filt = [x for x in items if '2026-03-01' <= x['date'] <= '2026-04-02']
filt.sort(key=lambda x: (sc(x['title']), x['date']), reverse=True)
top = filt[:3]

print(f'Top {len(top)} items:')
for i, it in enumerate(top):
    print(f'  {i+1}. [{it["date"]}] {it["area"]} | {it["title"][:55]}')

# ─── 生成 Excel ───
headers = [
    '序号', '项目名称', '建设单位', '控制价（万元）',
    '入围方法', '评标办法', '开标日期', '开标时间',
    '付款方式', '资质要求', '企业业绩要求', '项目经理业绩要求',
    '是否编写技术标', '保证金（万元）'
]
col_widths = [6, 52, 28, 14, 20, 24, 13, 10, 42, 42, 42, 42, 16, 16]

def col_ref(c, r):
    return chr(65 + c) + str(r)

def xml_esc(s):
    s = str(s)
    s = s.replace('&', '&amp;')
    s = s.replace('<', '&lt;')
    s = s.replace('>', '&gt;')
    s = s.replace('"', '&quot;')
    s = s.replace("'", '&apos;')
    return s

# Build rows
rows_xml = []
# Header row (style 1 = header bold)
rows_xml.append(
    '<row r="1">' +
    ''.join(
        f'<c r="{col_ref(i, 1)}" s="1" t="inlineStr"><is><t>{xml_esc(h)}</t></is></c>'
        for i, h in enumerate(headers)
    ) +
    '</row>'
)
# Data rows (style 2 = wrap text)
for ri, row in enumerate(top, 2):
    fields = [
        str(ri - 1),
        row['title'],
        '',
        '',
        '',
        '',
        row['date'],
        '',
        '',
        '',
        '',
        '',
        '',
        '',
    ]
    rows_xml.append(
        '<row r="' + str(ri) + '">' +
        ''.join(
            f'<c r="{col_ref(i, ri)}" s="2" t="inlineStr"><is><t>{xml_esc(v)}</t></is></c>'
            for i, v in enumerate(fields)
        ) +
        '</row>'
    )

cols_xml = ''.join(
    f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>'
    for i, w in enumerate(col_widths)
)

sheet_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<sheetViews>'
    '<sheetView workbookViewId="0" showGridLines="1">'
    '<selection activeCell="A1" sqref="A1"/>'
    '</sheetView>'
    '</sheetViews>'
    '<cols>' + cols_xml + '</cols>'
    '<sheetData>' + ''.join(rows_xml) + '</sheetData>'
    '<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75"/>'
    '</worksheet>'
)

wb_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets>'
    '<sheet name="招标项目明细" sheetId="1" r:id="rId1"/>'
    '</sheets>'
    '</workbook>'
)

styles_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts>'
    '<font><sz val="11"/><name val="微软雅黑"/></font>'
    '<font><sz val="11"/><b/><name val="微软雅黑"/></font>'
    '</fonts>'
    '<fills>'
    '<fill><patternFill patternType="none"/></fill>'
    '<fill><patternFill patternType="gray125"/></fill>'
    '<fill><patternFill patternType="solid"><fgColor rgb="FF028090"/></fgColor></patternFill>'
    '</fills>'
    '<borders><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment horizontal="center"/></xf>'
    '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment wrapText="1"/></xf>'
    '</cellXfs>'
    '</styleSheet>'
)

rels_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>'
)

ct_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '</Types>'
)

root_xml = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>'
)

buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct_xml)
    z.writestr('_rels/.rels', root_xml)
    z.writestr('xl/workbook.xml', wb_xml)
    z.writestr('xl/_rels/workbook.xml.rels', rels_xml)
    z.writestr('xl/styles.xml', styles_xml)
    z.writestr('xl/worksheets/sheet1.xml', sheet_xml)

out = '/workspace/skills/changzhou-bid-search/outputs/常州招标_2026-04-01.xlsx'
os.makedirs(os.path.dirname(out), exist_ok=True)
with open(out, 'wb') as f:
    f.write(buf.getvalue())

sz = os.path.getsize(out)
print(f'\nExcel saved: {out}')
print(f'Size: {sz} bytes')
print(f'Rows: {len(top)} data rows + 1 header')
print('\nContents:')
for i, it in enumerate(top):
    print(f'  Row {i+2}: [{it["date"]}] {it["area"]} | {it["title"][:60]}')
print('\nDone.')
