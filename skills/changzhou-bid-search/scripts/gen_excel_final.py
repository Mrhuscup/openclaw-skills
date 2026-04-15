#!/usr/bin/env python3
"""生成常州招标Excel - 整合跨站抓取结果（完整19列）"""
import io, zipfile, os, re

def make_xlsx(data, out_path):
    headers = ['序号', '分类', '项目名称', '地区', '发布日期',
               '建设单位', '控制价(万元)', '评标办法', '入围方法',
               '资格审查', '开标日期', '开标时间', '工期(天)',
               '资质要求', '企业业绩要求', '项目经理业绩要求',
               '技术标', '保证金(万元)', '付款方式', '联系人', '联系电话', '来源']

    col_widths = [5, 22, 45, 8, 12,
                  22, 12, 22, 18,
                  10, 13, 10, 8,
                  38, 42, 42,
                  14, 14, 38, 14, 18, 12]

    def cr(c, r): return chr(65+c)+str(r)
    def esc(s): return str(s).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

    strings, si_map = [], {}
    def si(s):
        s = str(s)
        if s not in si_map:
            si_map[s] = len(strings)
            strings.append(s)
        return si_map[s]

    for h in headers: si(h)

    rows_xml = ['<row r="1">' + ''.join(f'<c r="{cr(i,1)}" s="1" t="s"><v>{si(h)}</v></c>' for i,h in enumerate(headers)) + '</row>']
    for ri, d in enumerate(data, 2):
        fields = [
            str(ri-1),
            d.get('cat_name', ''),
            d.get('title', ''),
            d.get('area', ''),
            d.get('date', ''),
            d.get('client', ''),
            d.get('budget', ''),
            d.get('bid_method', ''),
            d.get('enter_method', ''),
            d.get('qual_check', ''),
            d.get('open_date', ''),
            d.get('open_time', ''),
            d.get('duration', ''),
            d.get('qual_req', ''),
            d.get('corp_perf', ''),
            d.get('pm_perf', ''),
            d.get('tech_bid', ''),
            d.get('deposit', ''),
            d.get('payment', ''),
            d.get('contact', ''),
            d.get('phone', ''),
            d.get('_source', ''),
        ]
        for v in fields: si(v)
        rows_xml.append('<row r="'+str(ri)+'">' + ''.join(f'<c r="{cr(i,ri)}" s="2" t="s"><v>{si(v)}</v></c>' for i,v in enumerate(fields)) + '</row>')

    cols = ''.join(f'<col min="{i+1}" max="{i+1}" width="{w}" customWidth="1"/>' for i,w in enumerate(col_widths))
    sheet = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="1"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><cols>{cols}</cols><sheetData>{"".join(rows_xml)}</sheetData><pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75"/></worksheet>'
    ss_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(strings)}" uniqueCount="{len(strings)}">{"".join(f"<si><t>{esc(s)}</t></si>" for s in strings)}</sst>'
    wb = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="\u62db\u6807\u9879\u76ee\u660e\u7ec6" sheetId="1" r:id="rId1"/></sheets></workbook>'
    styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><sz val="11"/><name val="\u5fae\u8f6f\u96c5\u9ed1"/></font><font><sz val="11"/><b/><name val="\u5fae\u8f6f\u96c5\u9ed1"/></font></fonts><fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF028090"/></fgColor></patternFill></fills><borders><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment wrapText="1"/></xf></cellXfs></styleSheet>'
    wb_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
    root_rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
    ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', ct)
        z.writestr('_rels/.rels', root_rels)
        z.writestr('xl/workbook.xml', wb)
        z.writestr('xl/_rels/workbook.xml.rels', wb_rels)
        z.writestr('xl/styles.xml', styles)
        z.writestr('xl/worksheets/sheet1.xml', sheet)
        z.writestr('xl/sharedStrings.xml', ss_xml)

    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, 'wb') as f:
        f.write(buf.getvalue())
    print(f'  Excel: {out_path}  ({os.path.getsize(out_path)} bytes)')


# ═══════════════════════════════════════════════════════════════════════════
# 2026-04-02 今日抓取数据
# 数据来源: 常州公共资源交易中心(列表) + 千里马/剑鱼标讯(详情)
# ═══════════════════════════════════════════════════════════════════════════
projects = [

    # ── 001001001 招标公告/资审公告 ──────────────────────────────────────
    {
        'cat_name': '招标公告/资审公告',
        'title': '薛冶路（嫩江路-浏阳河路）工程项目市政施工总承包',
        'area': '市辖区',
        'date': '2026-04-01',
        'client': '常州黑牡丹建设投资有限公司',
        'budget': '6800',
        'bid_method': '综合评估法—评定分离（两阶段开标）',
        'enter_method': '评定分离（综合评估法—资格后审）',
        'qual_check': '资格后审',
        'open_date': '2026-04-28',
        'open_time': '09:30',
        'duration': '210',
        'qual_req': '市政公用工程施工总承包一级及以上',
        'corp_perf': '2021-04-01以来，企业或项目经理承担单项合同≥5000万元市政工程竣工业绩（需中标通知书+合同+四方验收材料）',
        'pm_perf': '同企业业绩要求（2021-04-01以来，单项合同≥5000万元市政工程）',
        'tech_bid': '是（须提供施工组织设计，技术复杂工程）',
        'deposit': '50万元（电汇/网银/保函/保单/信用承诺）',
        'payment': '预付款10%+月进度款80%+竣工结算付至97%+质保金3%（贰年保修期满无息返还）',
        'contact': '丁工（招标人） / 李工（招标代理）',
        'phone': '0519-85103252 / 0519-85225031',
        '_source': '常州站ggzy.xzsp.changzhou.gov.cn（详情页）',
    },

    # ── 001001005 中标候选人/评标结果公示 ────────────────────────────────
    {
        'cat_name': '中标候选人/评标结果公示',
        'title': '常州市城市轨道交通5号线工程土建施工09标未入围公示',
        'area': '市辖区',
        'date': '2026-04-01',
        'client': '常州地铁集团有限公司',
        'budget': '16415.88',
        'bid_method': '综合评估法',
        'enter_method': '有限数量制',
        'qual_check': '资格预审',
        'open_date': '待确认',
        'open_time': '待确认',
        'duration': '580',
        'qual_req': '市政公用工程施工总承包壹级及以上（联合体须具备相应资质）',
        'corp_perf': '2020-01-01以来，地铁车站（地下两层及以上）+盾构隧道（≥600m）工程业绩',
        'pm_perf': '壹级注册建造师（铁路/市政）+建安B证+高级工程师+55岁以下',
        'tech_bid': '待确认',
        'deposit': '待确认',
        'payment': '待PDF确认',
        'contact': '待确认（详情登录会员查看）',
        'phone': '待确认',
        '_source': '剑鱼标讯jianyu360.cn（需登录查看完整）',
    },

    # ── 001001008 中标合同 ─────────────────────────────────────────────
    {
        'cat_name': '中标合同',
        'title': '江苏省溧阳中等专业学校改扩建工程项目',
        'area': '溧阳市',
        'date': '2026-04-01',
        'client': '待确认',
        'budget': '待确认',
        'bid_method': '待确认',
        'enter_method': '待确认',
        'qual_check': '待确认',
        'open_date': '待确认',
        'open_time': '待确认',
        'duration': '待确认',
        'qual_req': '待确认',
        'corp_perf': '待确认',
        'pm_perf': '待确认',
        'tech_bid': '待确认',
        'deposit': '待确认',
        'payment': '待确认',
        'contact': '待确认',
        'phone': '待确认',
        '_source': '待抓取（需登录江苏站查看）',
    },

]

out = '/workspace/skills/changzhou-bid-search/outputs/常州招标_2026-04-02_完整版.xlsx'
make_xlsx(projects, out)
print(f'\n完成，共 {len(projects)} 条 | 生成: {out}')
