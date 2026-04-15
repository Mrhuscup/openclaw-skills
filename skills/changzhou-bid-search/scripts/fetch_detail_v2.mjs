/**
 * fetch_detail_v2.mjs
 * 抓取常州招标详情页 PDF 附件链接
 * 用法: node fetch_detail_v2.mjs <cat> <uuid>
 */
import https from 'node:https';
import http from 'node:http';
import { URL } from 'node:url';

const BASE = 'http://ggzy.xzsp.changzhou.gov.cn';

function fetch(url) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const mod = parsed.protocol === 'https:' ? https : http;
    const req = mod.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Referer': BASE + '/jyzx/001001/tradeInfonew.html',
      }
    }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetch(res.headers.location).then(resolve).catch(reject);
      }
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve(data));
    });
    req.on('error', reject);
    req.setTimeout(15000, () => { req.destroy(); reject(new Error('timeout')); });
  });
}

async function getRedirectUrl(cat, uuid) {
  // Try the trad_notice.html path first (gonggaomode=1)
  try {
    const html = await fetch(`${BASE}/trad_notice.html?infoid=${uuid}&chanquantype=${cat}`);
    if (html.length > 5000) {
      return { html, url: `${BASE}/trad_notice.html?infoid=${uuid}&chanquantype=${cat}` };
    }
  } catch(e) {}
  
  // Try the tradeInfonew detail path
  try {
    const html = await fetch(`${BASE}/jyzx/001001/tradeInfonew.html?category=${cat}&infoId=${uuid}`);
    return { html, url: `${BASE}/jyzx/001001/tradeInfonew.html?category=${cat}&infoId=${uuid}` };
  } catch(e) {}
  
  return null;
}

async function main() {
  const cat = process.argv[2] || '001001001001';
  const uuid = process.argv[3];
  
  if (!uuid) {
    // Demo: get first item from list page
    console.log('Fetching list page to find items...');
    const listHtml = await fetch(BASE + '/jyzx/001001/tradeInfonew.html?category=001001001');
    
    // Extract items
    const itemRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
    const items = [];
    let m;
    while ((m = itemRe.exec(listHtml)) !== null) {
      const row = m[1];
      const segs = row.split('</td>');
      if (segs.length < 4) continue;
      const onclick = segs[1];
      const om = onclick.match(/tzjydetail\s*\(\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*\)/);
      if (!om) continue;
      const tm = segs[1].match(/title="([^"]{5,120})"/);
      if (!tm) continue;
      const area = segs[2].replace(/<[^>]+>/g, '').replace(/\xa0/g, ' ').trim();
      const date = segs[3].replace(/<[^>]+>/g, '').replace(/\xa0/g, ' ').trim();
      items.push({
        cat: om[1], uuid: om[2], gonggaomode: om[3], cqtype: om[4],
        title: tm[1].slice(0, 60), area, date
      });
    }
    
    console.log(`\nFound ${items.length} items:\n`);
    items.slice(0, 5).forEach((it, i) => {
      console.log(`[${i+1}] ${it.date} | ${it.area} | ${it.title}`);
      console.log(`    cat=${it.cat} uuid=${it.uuid} gonggaomode=${JSON.stringify(it.gonggaomode)}`);
    });
    
    if (items.length > 0) {
      console.log('\n--- Testing first item detail page ---');
      const first = items[0];
      return getRedirectUrl(first.cat, first.uuid);
    }
    return;
  }
  
  return getRedirectUrl(cat, uuid);
}

main().then(result => {
  if (!result) { console.log('No result'); process.exit(0); }
  const { html, url } = result;
  console.log(`\nDetail URL: ${url}`);
  console.log(`HTML length: ${html.length}`);
  
  // Extract PDF links
  const pdfRe = /href=["']([^"']*\.pdf[^"']*)["']/gi;
  const docRe = /href=["']([^"']*(?:docx?|zip|rar|txt|gif|bmp)[^"']*)["']/gi;
  const links = [];
  let m;
  while ((m = pdfRe.exec(html)) !== null) links.push({ type: 'pdf', url: m[1] });
  while ((m = docRe.exec(html)) !== null) links.push({ type: 'doc', url: m[1] });
  
  // Also look for attachment div/section
  const attachSection = html.match(/招标文件附件[\s\S]{0,500}/i);
  
  console.log(`\nLinks found (${links.length}):`);
  links.forEach(l => console.log(`  [${l.type}] ${l.url}`));
  
  if (attachSection) console.log('\nAttachment section:', attachSection[0].slice(0, 300));
  
  // Show a snippet of the HTML around key areas
  const titleArea = html.match(/<title>[^<]+/i);
  if (titleArea) console.log('\nPage title:', titleArea[0]);
  
}).catch(e => { console.error('Error:', e.message); process.exit(1); });
