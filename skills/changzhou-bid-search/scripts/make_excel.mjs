import JSZip from 'jszip';
import fs from 'fs';

const C = {
  primary: '028090',
  secondary: '00A896',
  dark: '1E293B',
  light: 'F0FDFA',
  white: 'FFFFFF',
  gray: '64748B',
  coral: 'EF4444',
  gold: 'F59E0B',
};

// Build shared strings
const sharedStrings = [];
const stringIndex = new Map();

function addStr(s) {
  s = String(s ?? '');
  if (!stringIndex.has(s)) {
    stringIndex.set(s, sharedStrings.length);
    sharedStrings.push(s);
  }
  return stringIndex.get(s);
}

// Fields: 序号, 项目名称, 建设单位, 控制价, 入围方法, 评标办法, 开标日期, 开标时间, 付款方式, 资质要求, 企业业绩要求, 项目经理业绩要求, 是否编写技术标, 保证金
const headers = ['序号','项目名称','建设单位','控制价（万元）','入围方法','评标办法','开标日期','开标时间','付款方式','资质要求','企业业绩要求','项目经理业绩要求','是否编写技术标','保证金（万元）'];

const rows = [
  {
    title: '未来智慧城幼儿园、小学、初级中学及共享健身中心（学校室内体育馆）项目塑胶、硅PU及路灯工程市政施工总承包',
    client: '江苏常州天宁经济开发区管理委员会',
    budget: '420',
    bidMethod: '☑合理低价法',
    enterMethod: '资格后审',
    openDate: '2026-04-14',
    openTime: '09:30',
    payment: '', // 待从PDF提取
    qualReq: '市政公用工程施工总承包三级及以上',
    corpPerf: '', // 3.4.1未勾选，无业绩要求
    pmPerf: '',
    techBid: '', // 页面正文未提及技术标
    deposit: '', // 待从PDF提取（3.4.1）
  },
];

const zip = new JSZip();

zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`);

zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);

zip.file('xl/workbook.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="招标项目明细" sheetId="1" r:id="rId1"/></sheets>
</workbook>`);

zip.file('xl/_rels/workbook.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);

zip.file('xl/styles.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts>
    <font><sz val="11"/><name val="微软雅黑"/></font>
    <font><sz val="11"/><b/><name val="微软雅黑"/></font>
  </fonts>
  <fills>
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF028090"/></fgColor></patternFill></fill>
  </fills>
  <borders><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment horizontal="center"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment wrapText="1"/></xf>
  </cellXfs>
</styleSheet>`);

// Build rows
const colWidths = [6, 48, 28, 12, 18, 22, 13, 10, 40, 42, 42, 42, 16, 16];

function buildCellRef(col, row) {
  return String.fromCharCode(65 + col) + String(row);
}

function escapeXml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

const sheetRows = [];

// Header row
sheetRows.push(
  '<row r="1">' +
  headers.map((h, ci) =>
    `<c r="${buildCellRef(ci, 1)}" s="1"><is><t>${escapeXml(h)}</t></is></c>`
  ).join('') +
  '</row>'
);

// Data rows
for (let ri = 0; ri < rows.length; ri++) {
  const r = rows[ri];
  const rNum = ri + 2;
  const fields = [
    String(ri + 1),   // 序号
    r.title,          // 项目名称
    r.client,         // 建设单位
    r.budget,         // 控制价
    r.enterMethod,    // 入围方法
    r.bidMethod,      // 评标办法
    r.openDate,       // 开标日期
    r.openTime,       // 开标时间
    r.payment,        // 付款方式
    r.qualReq,        // 资质要求
    r.corpPerf,       // 企业业绩要求
    r.pmPerf,         // 项目经理业绩要求
    r.techBid,        // 是否编写技术标
    r.deposit,        // 保证金
  ];
  sheetRows.push(
    '<row r="' + rNum + '">' +
    fields.map((val, ci) =>
      `<c r="${buildCellRef(ci, rNum)}" s="2"><is><t>${escapeXml(String(val))}</t></is></c>`
    ).join('') +
    '</row>'
  );
}

const sheetData = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetViews>
    <sheetView workbookViewId="0" showGridLines="1">
      <selection activeCell="A1" sqref="A1"/>
    </sheetView>
  </sheetViews>
  <cols>${colWidths.map((w, i) => `<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('')}</cols>
  <sheetData>${sheetRows.join('')}</sheetData>
  <pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75"/>
</worksheet>`;

zip.file('xl/worksheets/sheet1.xml', sheetData);

// Shared strings
const ssXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${sharedStrings.length}" uniqueCount="${sharedStrings.length}">
${sharedStrings.map(s => `<si><t>${escapeXml(s)}</t></si>`).join('\n')}
</sst>`;
zip.file('xl/sharedStrings.xml', ssXml);

const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
const outPath = '/workspace/skills/changzhou-bid-search/outputs/常州招标_2026-03-31_示例.xlsx';
fs.mkdirSync('/workspace/skills/changzhou-bid-search/outputs', { recursive: true });
fs.writeFileSync(outPath, buf);
console.log('Excel created:', outPath, '| Size:', buf.length, 'bytes');
console.log('Rows:', rows.length);
