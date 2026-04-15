#!/usr/bin/env node
/**
 * 常州招标抓取脚本 v2（适配 2026-04 网站改版）
 * 默认抓取"建设工程-不限"分类，按日期倒序
 *
 * 使用方法：
 *   node scrape_v2.js [开始日期YYYY-MM-DD] [结束日期YYYY-MM-DD] [最大条数]
 *
 * 示例：
 *   node scrape_v2.js 2026-03-01 2026-04-30 20
 *   node scrape_v2.js 2026-04-01 2026-04-30 50
 *   node scrape_v2.js                               # 全部日期，默认20条
 */

'use strict';

const https = require('https');
const http = require('http');
const fs = require('fs');
const path = require('path');

const BASE_URL = 'ggzy.xzsp.changzhou.gov.cn';
const PROTOCOL = 'http';

function httpGet(url) {
  return new Promise(function(resolve, reject) {
    const lib = url.startsWith('https') ? https : http;
    const req = lib.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'zh-CN,zh;q=0.9',
      }
    }, function(res) {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        const redirectUrl = new URL(res.headers.location, url).toString();
        return httpGet(redirectUrl).then(resolve).catch(reject);
      }
      let data = '';
      res.on('data', function(chunk) { data += chunk; });
      res.on('end', function() { resolve(data); });
    });
    req.on('error', reject);
    req.setTimeout(15000, function() { req.destroy(); reject(new Error('timeout: ' + url)); });
  });
}

function delay(ms) { return new Promise(function(r) { setTimeout(r, ms); }); }

/**
 * 解析列表页 HTML
 * HTML 结构: <tr><td>n</td><td><a onclick="tzjydetail('cat','uuid','','')" title="标题">text</a></td><td>地区</td><td>日期</td></tr>
 * 策略: 逐行扫描 <tr>，uuid 在 segments[1]（第2格）
 */
function parseListHtml(html) {
  const items = [];
  const seen = {};

  const trStartRe = /<tr[^>]*>/gi;
  const trEndRe = /<\/tr>/gi;

  const trStarts = [...html.matchAll(trStartRe)].map(function(m) { return m.index; });
  const trEnds = [...html.matchAll(trEndRe)].map(function(m) { return m.index; });

  for (let i = 0; i < trStarts.length; i++) {
    const start = trStarts[i];
    const endIdx = trEnds.findIndex(function(e) { return e >= start; });
    if (endIdx === -1) continue;
    const rowEnd = trEnds[endIdx] + 5;
    const rowHtml = html.slice(start, rowEnd);

    const segs = rowHtml.split('</td>');
    if (segs.length < 4) continue;

    const onclickInSeg1 = segs[1].match(/tzjydetail\('([^']+)',\s*'([^']+)'\s*,/);
    if (!onclickInSeg1) continue;
    const cat = onclickInSeg1[1];
    const uuid = onclickInSeg1[2];

    // uuid 必须在 segments[1]（第2格），不在 segments[0]（序号格）
    if (segs[0].indexOf(uuid) !== -1) continue;

    const titleMatch = segs[1].match(/title="([^"]{5,120})"/);
    if (!titleMatch) continue;
    const title = titleMatch[1];

    function strip(s) {
      return String(s).replace(/<[^>]*>/g, '').replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').trim();
    }
    const area = strip(segs[2]);
    const date = strip(segs[3]);

    if (seen[uuid]) continue;
    seen[uuid] = true;

    items.push({ category: cat, uuid, title: title.slice(0, 120), area, date });
  }

  return items;
}

async function fetchDetail(category, uuid) {
  const url = PROTOCOL + '://' + BASE_URL + '/jyzx/001001/tradeInfonew.html?category=' + category + '&infoId=' + uuid;
  try {
    const html = await httpGet(url);
    return extractFields(html, uuid);
  } catch (e) {
    return { uuid, error: e.message };
  }
}

function extractFields(html, targetUuid) {
  const text = html
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();

  function find(keywords, len) {
    len = len || 200;
    for (const kw of keywords) {
      const idx = text.indexOf(kw);
      if (idx !== -1) {
        const snippet = text.slice(idx, idx + len);
        const ci = snippet.indexOf('\uff1a');
        const ni = snippet.indexOf('\n');
        let end = len;
        if (ci !== -1) end = Math.min(end, ci + 1);
        if (ni !== -1) end = Math.min(end, ni);
        const val = snippet.slice(ci !== -1 ? ci + 1 : 0, end).trim();
        if (val.length > 1) return val.slice(0, 200);
      }
    }
    return '';
  }

  function findBidMethod() {
    for (const m of ['\u2611合理低价法', '\u2611综合评估法', '\u2611经评审的最低投标价法']) {
      if (text.indexOf(m) !== -1) return m.replace('\u2611', '');
    }
    return '';
  }

  function findAccessMethod() {
    // PDF第三章 2.1.2 评标入围方法：找所有被勾选的项目
    const checked = text.match(/[\u2611\u25a1]\s*[^\n<]{2,30}/g) || [];
    return checked.filter(s => s.startsWith('\u2611')).map(s => s.replace('\u2611', '').trim()).join('；');
  }

  function findBudget() {
    for (const m of ['\u5de5\u7a0b\u5408\u540c\u4f30\u7b97\u4ef7', '\u63a7\u5236\u4ef7', '\u5408\u540c\u4f30\u7b97\u4ef7']) {
      const idx = text.indexOf(m);
      if (idx !== -1) {
        const snippet = text.slice(idx, idx + 100);
        const nums = snippet.match(/[\d,\.．]+/);
        if (nums) return nums[0].replace(/[,\．]/g, '.') + ' \u4e07\u5143';
      }
    }
    return '';
  }

  return {
    uuid: targetUuid,
    title: find(['\u9879\u76ee\u540d\u79f0', '\u5de5\u7a0b\u540d\u79f0']),
    client: find(['\u62db\u6807\u4eba', '\u5efa\u8bbe\u5355\u4f4d', '\u9879\u76ee\u6cd5\u4eba']),
    contact: find(['\u8054\u7cfb\u4eba', '\u8054\u7cfb\u65b9\u5f0f']),
    budget: findBudget(),
    qualReq: find(['\u6295\u6807\u4eba\u8d44\u8d28', '\u8d44\u8d28\u8981\u6c42', '\u8d44\u8d28\u7b49\u7ea7']),
    // 开标日期：PDF第二章 7.1 投标截止时间
    openDate: find(['7.1 \u6295\u6807\u622a\u6b62\u65f6\u95f4', '\u6295\u6807\u622a\u6b62\u65f6\u95f4', '\u622a\u6b62\u65f6\u95f4']),
    bidMethod: findBidMethod(),
    // 入围方法：PDF第三章 2.1.2 评标入围方法（找☑勾选内容）
    accessMethod: findAccessMethod(),
    deposit: find(['\u6295\u6807\u4fdd\u8bc1\u91d1', '\u4fdd\u8bc1\u91d1']),
    payment: find(['\u4ed8\u6b3e\u65b9\u5f0f', '\u4ed8\u6b3e\u5468\u671f']),
  };
}

// ─── 主程序 ───

async function main() {
  const args = process.argv.slice(2);
  const startDate = args[0] || '';
  const endDate = args[1] || '';
  const maxItems = parseInt(args[2] || '20', 10);

  console.log('================================================================');
  console.log('  常州市公共资源交易中心招标信息抓取工具 v2');
  console.log('  分类：建设工程-不限  |  2026-04-01 更新版');
  console.log('  适用：江苏八达路桥有限公司');
  console.log('================================================================');
  console.log('  日期范围：' + (startDate || '不限') + ' ~ ' + (endDate || '不限'));
  console.log('  最大条数：' + maxItems + '\n');

  console.log('[Step 1] 抓取列表页...');
  const listUrl = PROTOCOL + '://' + BASE_URL + '/jyzx/001001/tradeInfonew.html?category=001001';
  let html;
  try {
    html = await httpGet(listUrl);
  } catch (e) {
    console.error('  ❌ 列表页请求失败:', e.message);
    return;
  }

  const allItems = parseListHtml(html);
  console.log('   解析结果：' + allItems.length + ' 条\n');

  if (allItems.length === 0) {
    console.error('  ❌ 列表页解析失败（0条）。请确认网站结构是否有变化。\n');
    return;
  }

  // 打印前几条
  for (let i = 0; i < Math.min(allItems.length, 3); i++) {
    const it = allItems[i];
    console.log('   [' + (i+1) + '] ' + it.date + ' | ' + it.area + ' | ' + it.title.slice(0, 50));
    console.log('        uuid=' + it.uuid.slice(0, 8) + '... cat=' + it.category);
  }
  console.log('');

  // 日期过滤
  let filtered = allItems.filter(function(it) {
    if (!it.date) return true;
    if (startDate && it.date < startDate) return false;
    if (endDate && it.date > endDate) return false;
    return true;
  });

  // 按日期倒序（最新在前）
  filtered.sort(function(a, b) {
    return (b.date || '').localeCompare(a.date || '');
  });

  console.log('\u260f\ufe0f 过滤结果：' + filtered.length + ' 条');
  for (let i = 0; i < Math.min(filtered.length, 10); i++) {
    const it = filtered[i];
    console.log('   [' + (i+1) + '] ' + it.date + ' | ' + it.area + ' | ' + it.title.slice(0, 48));
  }
  console.log('');

  const toProcess = filtered.slice(0, maxItems);

  console.log('\u270b [Step 2] 抓取详情（' + toProcess.length + ' 条）...');
  console.log('  注：详情页为 JS 动态渲染；字段提取自 HTML 正文，完整数据请下载 PDF\n');

  const details = [];
  for (let i = 0; i < toProcess.length; i++) {
    const it = toProcess[i];
    const detail = await fetchDetail(it.category, it.uuid);
    const merged = Object.assign({}, it, detail);
    details.push(merged);
    const ok = !detail.error;
    console.log('  [' + (i+1) + '/' + toProcess.length + '] ' + (ok ? '\u2705' : '\u26a0\ufe0f') +
      ' | ' + it.date + ' | ' + it.title.slice(0, 40));
    if (detail.error) console.log('         \u9519\u8bef: ' + detail.error);
    await delay(500);
  }

  console.log('\n\u270b [Step 3] 生成 Excel...');
  await generateExcel(details);

  console.log('\n\u2701 \u5b8c\u6210\uff01\u5171\u5904\u7406 ' + details.length + ' \u6761\u8bb0\u5f55\uff01\n');
}

// ─── Excel 生成 ───

function escapeXml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function buildCellRef(col, row) { return String.fromCharCode(65 + col) + row; }

async function generateExcel(data) {
  const JSZip = require('jszip');
  const zip = new JSZip();

  const headers = [
    '\u5e8f\u53f7', '\u9879\u76ee\u540d\u79f0', '\u5efa\u8bbe\u5355\u4f4d', '\u63a7\u5236\u4ef7\uff08\u4e07\u5143\uff09',
    '\u5165\u56f4\u65b9\u6cd5', '\u8bc4\u6807\u529e\u6cd5', '\u5f00\u6807\u65e5\u671f', '\u5f00\u6807\u65f6\u95f4',
    '\u4ed8\u6b3e\u65b9\u5f0f', '\u8d44\u8d28\u8981\u6c42', '\u4f01\u4e1a\u4e1a\u7ee9\u8981\u6c42',
    '\u9879\u76ee\u7ecf\u7406\u4e1a\u7ee9\u8981\u6c42', '\u662f\u5426\u7f16\u5199\u6280\u672f\u6807', '\u4fdd\u8bc1\u91d1\uff08\u4e07\u5143\uff09'
  ];
  const colWidths = [6, 52, 28, 14, 20, 24, 13, 10, 42, 42, 42, 42, 16, 16];

  zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
  zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>');
  zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="\u62db\u6807\u9879\u76ee\u660e\u7ec6" sheetId="1" r:id="rId1"/></sheets></workbook>');
  zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>');
  zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts><font><sz val="11"/><name val="\u5fae\u8f6f\u96c5\u9ed1"/></font><font><sz val="11"/><b/><name val="\u5fae\u8f6f\u96c5\u9ed1"/></font></fonts><fills><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF028090"/></fgColor></patternFill></fills><borders><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0"><alignment horizontal="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"><alignment wrapText="1"/></xf></cellXfs></styleSheet>');

  // sharedStrings
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

  const sheetRows = [];
  sheetRows.push('<row r="1">' + headers.map(function(h, ci) {
    return '<c r="' + buildCellRef(ci, 1) + '" s="1" t="s"><v>' + addStr(h) + '</v></c>';
  }).join('') + '</row>');

  for (let ri = 0; ri < data.length; ri++) {
    const row = data[ri];
    const rNum = ri + 2;
    const fields = [
      String(ri + 1),
      row.title || '',
      row.client || '',
      row.budget || '',
      row.accessMethod || '',
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
    sheetRows.push('<row r="' + rNum + '">' + fields.map(function(val, ci) {
      return '<c r="' + buildCellRef(ci, rNum) + '" s="2" t="s"><v>' + addStr(String(val)) + '</v></c>';
    }).join('') + '</row>');
  }

  const colsXml = colWidths.map(function(w, i) {
    return '<col min="' + (i+1) + '" max="' + (i+1) + '" width="' + w + '" customWidth="1"/>';
  }).join('');

  const sheetData = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" showGridLines="1"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews><cols>' + colsXml + '</cols><sheetData>' + sheetRows.join('') + '</sheetData><pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75"/></worksheet>';
  zip.file('xl/worksheets/sheet1.xml', sheetData);

  const ssXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + sharedStrings.length + '" uniqueCount="' + sharedStrings.length + '">\n' + sharedStrings.map(function(s) { return '<si><t>' + escapeXml(s) + '</t></si>'; }).join('\n') + '\n</sst>';
  zip.file('xl/sharedStrings.xml', ssXml);

  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const outPath = path.join(__dirname, '..', 'outputs', '\u5e38\u5dde\u62db\u6807_' + new Date().toISOString().slice(0, 10) + '.xlsx');
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, buf);
  console.log('\n  \u2705 Excel \u5df2\u4fdd\u5b58: ' + outPath + '  (' + buf.length + ' bytes)\n');
}

main().catch(function(e) {
  console.error('\u811a\u672c\u9519\u8bef:', e);
  process.exit(1);
});
