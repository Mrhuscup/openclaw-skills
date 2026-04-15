#!/usr/bin/env node
'use strict';

const https = require('https');
const http = require('http');
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

const BASE_URL = 'ggzy.xzsp.changzhou.gov.cn';
const PROTOCOL = 'http';

const CATEGORIES = {
  '001001': '建设工程',
  '001002': '交通工程',
  '001003': '水利工程',
  '001004': '政府采购',
  '001005': '土地矿产',
  '001006': '国有产权',
  '001009': '其他交易',
};

const RELEVANT_KEYWORDS = [
  '市政', '道路', '公路', '桥梁', '隧道', '排水', '照明',
  '绿化', '交通', '施工', '总承包', '改建', '新建', '提升',
];

function httpGet(url) {
  return new Promise(function(resolve, reject) {
    const lib = url.startsWith('https') ? https : http;
    const req = lib.get(url, { headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36' } }, function(res) {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return httpGet(res.headers.location).then(resolve).catch(reject);
      }
      let data = '';
      res.on('data', function(chunk) { data += chunk; });
      res.on('end', function() { resolve(data); });
    });
    req.on('error', reject);
    req.setTimeout(10000, function() { req.destroy(); reject(new Error('Request timeout')); });
  });
}

function parseDateFromUrl(url) {
  const m = String(url).match(/\/(\d{8})\//);
  if (m) return m[1].slice(0, 4) + '-' + m[1].slice(4, 6) + '-' + m[1].slice(6, 8);
  return '';
}

function scoreRelevance(title) {
  const t = title.toLowerCase();
  return RELEVANT_KEYWORDS.reduce(function(s, k) {
    return s + (t.indexOf(k.toLowerCase()) !== -1 ? 1 : 0);
  }, 0);
}

function delay(ms) {
  return new Promise(function(r) { setTimeout(r, ms); });
}

function extractField(text, keywords) {
  for (let i = 0; i < keywords.length; i++) {
    const kw = keywords[i];
    const idx = text.indexOf(kw);
    if (idx !== -1) {
      const snippet = text.substring(idx, idx + 300);
      const colonIdx = snippet.indexOf('\uff1a');
      const newlineIdx = snippet.indexOf('\n');
      let end = 300;
      if (colonIdx !== -1 && colonIdx < end) end = colonIdx + 1;
      if (newlineIdx !== -1 && newlineIdx < end) end = newlineIdx;
      const val = snippet.substring(colonIdx !== -1 ? colonIdx + 1 : 0, end).trim();
      if (val && val.length > 1) return val.slice(0, 200);
    }
  }
  return '';
}

function extractBudget(text) {
  const m = text.match(/(?:[\d,\.．]+)[万千佰零〇\d]*(?:万元|万)?/);
  return m ? m[0].slice(0, 50) : '';
}

function parseListHtml(html, catCode) {
  const items = [];
  // Extract from URL patterns like href="/jyzx/001001/..."
  const hrefMatches = html.match(/href="(\/jyzx\/[^"]+)"/g) || [];
  const seen = {};
  for (let i = 0; i < hrefMatches.length; i++) {
    const match = hrefMatches[i].match(/href="([^"]+)"/);
    if (!match) continue;
    const href = match[1];
    if (!href || seen[href]) continue;
    seen[href] = true;
    if (href.indexOf('/jyzx/') === -1) continue;
    // extract title from anchor text nearby
    const dateMatch = href.match(/\/(\d{8})\//);
    const fullUrl = href.startsWith('http') ? href : PROTOCOL + '://' + BASE_URL + href;
    // Try to find title near this href in HTML
    const titleRe = new RegExp('>([^<]{5,80})<' + href.replace(/\//g, '\\/').replace(/\./g, '\\.') + '["\']?');
    const titleMatch2 = html.match(titleRe);
    let title = titleMatch2 ? titleMatch2[1].replace(/^\[\d{2}-\d{2}\]/, '').trim() : '';
    items.push({
      url: fullUrl,
      title: title || '（无标题）',
      date: dateMatch ? parseDateFromUrl(dateMatch[1]) : '',
      category: CATEGORIES[catCode] || catCode,
    });
  }
  return items;
}

async function fetchListPage(catCode, keyword, page) {
  const wd = encodeURIComponent(keyword || '\u5e02\u653f');
  const searchUrl = PROTOCOL + '://' + BASE_URL + '/search/fullsearch.html?cnum=' + catCode + '&wd=' + wd + '&page=' + page;
  try {
    const html = await httpGet(searchUrl);
    return parseListHtml(html, catCode);
  } catch (e) {
    return [];
  }
}

async function fetchDetail(url) {
  try {
    const html = await httpGet(url);
    return parseDetailHtml(html, url);
  } catch (e) {
    return { url: url, error: e.message };
  }
}

function parseDetailHtml(html, url) {
  try {
    // Simple text extraction - get all text from body
    const bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    const bodyText = bodyMatch ? bodyMatch[1].replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim() : '';
    var result = {
      url: url,
      title: '',
      client: extractField(bodyText, ['\u5efa\u8bbe\u5355\u4f4d', '\u62db\u6807\u4eba', '\u53d1\u5305\u4eba', '\u9879\u76ee\u5355\u4f4d', '\u59d4\u6258\u5355\u4f4d']),
      contact: extractField(bodyText, ['\u8054\u7cfb\u4eba', '\u8054\u7cfb\u7535\u8bdd', '\u7535\u8bdd']),
      budget: extractBudget(bodyText),
      qualReq: extractField(bodyText, ['\u8d44\u8d28\u8981\u6c42', '\u8d44\u8d28\u6761\u4ef6', '\u6295\u6807\u8d44\u8d28']),
      openDate: extractField(bodyText, ['\u5f00\u6807\u65f6\u95f4', '\u5f00\u6807\u65e5\u671f', '\u6295\u6807\u622a\u6b62\u65f6\u95f4']),
      payment: extractField(bodyText, ['\u4ed8\u6b3e\u65b9\u5f0f', '\u652f\u4ed8\u65b9\u5f0f', '\u4ed8\u6b3e\u529e\u6cd5']),
      bidMethod: extractField(bodyText, ['\u8bc4\u6807\u529e\u6cd5', '\u8bc4\u5ba1\u529e\u6cd5', '\u8bc4\u5ba1\u65b9\u6cd5', '\u5165\u56f4\u65b9\u6cd5']),
      deposit: extractBudget(bodyText),
    };
    return result;
  } catch (e) {
    return { url: url, error: e.message };
  }
}

async function search(keyword, startDate, endDate, category) {
  var cats = category ? [[category, CATEGORIES[category] || category]] : Object.entries(CATEGORIES);
  var results = [];
  console.log('\n\u2705 \u5f00\u59cb\u641c\u7d22: \u5173\u952e\u8bcd="' + keyword + '"  \u5206\u7c7b=' + cats.map(function(c) { return c[1]; }).join(', '));
  console.log('  \u65e5\u671f\u8303\u56f4: ' + (startDate || '\u4e0d\u9650') + ' ~ ' + (endDate || '\u4e0d\u9650') + '\n');

  for (var ci = 0; ci < cats.length; ci++) {
    var catCode = cats[ci][0];
    var catName = cats[ci][1];
    try {
      var items = await fetchListPage(catCode, keyword, 1);
      await delay(400);
      if (items.length === 0) {
        console.log('  ' + catName + ': \u672a\u627e\u5230\u5173\u8054\u5185\u5bb9');
        continue;
      }
      // Filter by date
      var filtered = items.filter(function(it) {
        if (!it.date) return true;
        if (startDate && it.date < startDate) return false;
        if (endDate && it.date > endDate) return false;
        return true;
      });
      results = results.concat(filtered);
      console.log('  ' + catName + ': \u627e\u5230 ' + items.length + ' \u6761\uff0c\u65e5\u671f\u8303\u56f4\u5185 ' + filtered.length + ' \u6761');
    } catch (e) {
      console.log('  ' + catName + ': \u83b7\u53d6\u5931\u8d25 - ' + e.message);
    }
  }

  // Sort by relevance + date
  results.sort(function(a, b) {
    var sa = scoreRelevance(a.title);
    var sb = scoreRelevance(b.title);
    if (sa !== sb) return sb - sa;
    return (b.date || '').localeCompare(a.date || '');
  });

  console.log('\n\u5171\u627e\u5230 ' + results.length + ' \u6761\u516c\u544a\uff08\u5173\u8054\u5185\u5bb9\uff09\n');
  return results;
}

async function fetchDetails(items, maxItems) {
  var details = [];
  var count = Math.min(items.length, maxItems || 20);
  console.log('\u2705 \u5f00\u59cb\u62d3\u53d6\u8be6\u60c5\uff08\u524d ' + count + ' \u6761\uff09...\n');
  for (var i = 0; i < count; i++) {
    var it = items[i];
    process.stdout.write('  [' + (i + 1) + '/' + count + '] ' + (it.title || '').slice(0, 40) + '... ');
    var detail = await fetchDetail(it.url);
    var merged = { title: it.title, date: it.date, category: it.category, url: it.url };
    for (var k in detail) merged[k] = detail[k];
    details.push(merged);
    console.log(detail.error ? '\u2717 ' + detail.error.slice(0, 40) : '\u2713');
    await delay(500);
  }
  return details;
}

function escapeXml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildCellRef(col, row) {
  return String.fromCharCode(65 + col) + row;
}

async function generateExcel(data, outputPath) {
  console.log('\n\u2705 \u6b63\u5728\u751f\u6210 Excel...\n');
  var headers = [
    '\u5e8f\u53f7', '\u9879\u76ee\u540d\u79f0', '\u5efa\u8bbe\u5355\u4f4d', '\u63a7\u5236\u4ef7\uff08\u4e07\u5143\uff09',
    '\u5165\u56f4\u65b9\u6cd5', '\u8bc4\u6807\u529e\u6cd5', '\u5f00\u6807\u65e5\u671f', '\u5f00\u6807\u65f6\u95f4',
    '\u4ed8\u6b3e\u65b9\u5f0f', '\u8d44\u8d28\u8981\u6c42', '\u4f01\u4e1a\u4e1a\u7ee9\u8981\u6c42',
    '\u9879\u76ee\u7ecf\u7406\u4e1a\u7ee9\u8981\u6c42', '\u662f\u5426\u7f16\u5199\u6280\u672f\u6807',
    '\u4fdd\u8bc1\u91d1\uff08\u4e07\u5143\uff09'
  ];
  var colWidths = [6, 45, 25, 15, 22, 35, 13, 10, 40, 40, 40, 40, 15, 15];

  var zip = new JSZip();

  // [Content_Types].xml
  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');

  // _rels/.rels
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');

  // xl/workbook.xml
  zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="\u62db\u6807\u9879\u76ee\u660e\u7ec6" sheetId="1" r:id="rId1"/></sheets></workbook>');

  // xl/_rels/workbook.xml.rels
  zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');

  // xl/styles.xml
  zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><sz val="11"/><name val="\u5fae\u8f6f\u96c5\u9ed1"/></font><font><sz val="11"/><b/><name val="\u5fae\u8f6f\u96c5\u9ed1"/></font></fonts><fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF028090"/></fgColor></patternFill></fills><borders><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment wrapText="1"/></xf></cellXfs></styleSheet>');

  // Build sheet data
  var rows = [];

  // Header row (row 1)
  var headerCells = [];
  for (var hi = 0; hi < headers.length; hi++) {
    headerCells.push('<c r="' + buildCellRef(hi, 1) + '" s="1"><is><t>' + escapeXml(headers[hi]) + '</t></is></c>');
  }
  rows.push('<row r="1">' + headerCells.join('') + '</row>');

  // Data rows
  for (var ri = 0; ri < data.length; ri++) {
    var row = data[ri];
    var rowNum = ri + 2;
    var fields = [
      String(ri + 1),
      row.title || '',
      row.client || '',
      row.budget || '',
      row.bidMethod || '',
      row.bidMethod || '',
      row.date || '',
      row.openDate || '',
      row.payment || '',
      row.qualReq || '',
      '',
      '',
      '',
      row.deposit || '',
    ];
    var cells = [];
    for (var fi = 0; fi < fields.length; fi++) {
      cells.push('<c r="' + buildCellRef(fi, rowNum) + '" s="2"><is><t>' + escapeXml(String(fields[fi])) + '</t></is></c>');
    }
    rows.push('<row r="' + rowNum + '">' + cells.join('') + '</row>');
  }

  var sheetData = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="1"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><cols>' + colWidths.map(function(w, i) { return '<col min="' + (i+1) + '" max="' + (i+1) + '" width="' + w + '" customWidth="1"/>'; }).join('') + '</cols><sheetData>' + rows.join('') + '</sheetData><pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75"/></worksheet>';
  zip.file('xl/worksheets/sheet1.xml', sheetData);

  var buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, buf);
  console.log('\u2705 Excel \u5df2\u751f\u6210: ' + outputPath + '  (' + data.length + ' \u6761\u8bb0\u5f55)\n');
}

async function main() {
  var args = process.argv.slice(2);
  var keyword = args[0] || '';
  var startDate = args[1] || '';
  var endDate = args[2] || '';
  var category = args[3] || '';

  console.log('================================================================');
  console.log('  \u5e38\u5dde\u5e02\u516c\u5171\u8d44\u6e90\u4ea4\u6613\u4ea4\u6613\u4e2d\u5fc3\u62db\u6807\u4fe1\u606f\u62d3\u53d6\u5de5\u5177');
  console.log('  \u9002\u7528\uff1a\u6c5f\u82cf\u516b\u8fbe\u8def\u6865\u6709\u9650\u516c\u53f8');
  console.log('================================================================');
  console.log('  \u5173\u952e\u8bcd\uff1a' + (keyword || '(\u672a\u6307\u5b9a\uff0c\u9ed8\u8ba4\u5e02\u653f)'));
  console.log('  \u8d77\u59cb\uff1a' + (startDate || '\u4e0d\u9650') + '  \u622a\u6b62\uff1a' + (endDate || '\u4e0d\u9650'));
  console.log('  \u5206\u7c7b\uff1a' + (category ? (CATEGORIES[category] || category) : '\u5168\u90e8') + '\n');

  var items = await search(keyword, startDate, endDate, category);

  if (items.length === 0) {
    console.log('\u672a\u627e\u5230\u4efb\u4f55\u516c\u544a\uff0c\u8bf7\u68c0\u67e5\u5173\u952e\u8bcd\u6216\u65e5\u671f\u8303\u56f4\u3002');
    return;
  }

  var validItems = items.filter(function(it) { return it.url && it.url.indexOf('/jyzx/') !== -1; });
  var details = await fetchDetails(validItems, 20);

  var outputPath = path.join(__dirname, '..', 'outputs', '\u5e38\u5dde\u62db\u6807_' + new Date().toISOString().slice(0, 10) + '.xlsx');
  await generateExcel(details, outputPath);
}

main().catch(function(e) { console.error('\u9519\u8bef:', e); process.exit(1); });
