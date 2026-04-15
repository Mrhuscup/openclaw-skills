import https from 'node:https';
import http from 'node:http';
import { URL } from 'node:url';
import fs from 'fs';

const BASE = 'http://ggzy.xzsp.changzhou.gov.cn';

function fetchPage(url) {
  return new Promise((resolve, reject) => {
    const mod = new URL(url).protocol === 'https:' ? https : http;
    const req = mod.get(url, { headers: { 'User-Agent': 'Mozilla/5.0', 'Referer': BASE } },
      res => {
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          const loc = res.headers.location;
          return fetchPage(loc.startsWith('http') ? loc : BASE + loc)
            .then(resolve).catch(reject);
        }
        let d = '';
        res.on('data', c => d += c);
        res.on('end', () => resolve({ status: res.statusCode, body: d, url }));
      }
    );
    req.on('error', reject);
    req.setTimeout(15000, () => { req.destroy(); reject(new Error('timeout')); });
  });
}

async function main() {
  // Get the list page
  const listPage = await fetchPage(BASE + '/jyzx/001001/tradeInfonew.html?category=001001001');
  console.log('List page status:', listPage.status, 'len:', listPage.body.length);
  
  // Extract items
  const items = [];
  const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let m;
  while ((m = trRe.exec(listPage.body)) !== null) {
    const segs = m[1].split('</td>');
    if (segs.length < 4) continue;
    const om = segs[1].match(/tzjydetail\s*\(\s*'([^']+)'\s*,\s*'([^']+)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*\)/);
    if (!om) continue;
    const tm = segs[1].match(/title="([^"]{5,120})"/);
    if (!tm) continue;
    const area = segs[2].replace(/<[^>]+>/g, '').replace(/\xa0/g, ' ').trim();
    const date = segs[3].replace(/<[^>]+>/g, '').replace(/\xa0/g, ' ').trim();
    items.push({ cat: om[1], uuid: om[2], mode: om[3], cqtype: om[4], title: tm[1], area, date });
  }
  
  fs.writeFileSync('/workspace/skills/changzhou-bid-search/outputs/items_list.json', JSON.stringify(items, null, 2));
  console.log('Items saved:', items.length);
  items.slice(0, 3).forEach((it, i) => console.log(` [${i+1}] ${it.date} | ${it.area} | ${it.title.slice(0,40)} | mode=${it.mode}`));
  
  // Try to find siteGuid in the page
  const body = listPage.body;
  const guidRe = /[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}/gi;
  const guids = [...new Set(body.match(guidRe) || [])];
  console.log('\nGUIDs in page:', guids.slice(0, 3));
  
  // Find what URL patterns are in the JS blocks
  const jsBlocks = body.match(/<script[^>]*>([\s\S]*?)</script>/gi) || [];
  for (const block of jsBlocks.slice(0, 5)) {
    if (block.includes('siteInfo') || block.includes('projectName') || block.includes('frontPage')) {
      console.log('\nRelevant JS block:');
      console.log(block.slice(0, 500));
    }
  }
  
  // Try the redirect API with various paths
  const first = items[0];
  console.log('\n--- Testing first item ---');
  console.log('UUID:', first.uuid, 'CAT:', first.cat);
  
  // Try different API paths
  const paths = [
    '/frontPageRedirctAction.action?cmd=pageRedirect&infoid=' + first.uuid + '&siteGuid=&categorynum=' + first.cat,
    '/jyzx/frontPageRedirctAction.action?cmd=pageRedirect&infoid=' + first.uuid + '&categorynum=' + first.cat,
  ];
  
  for (const p of paths) {
    try {
      const r = await fetchPage(BASE + p);
      console.log(p.slice(0, 60), '->', r.status, r.body.slice(0, 100));
    } catch(e) {
      console.log(p.slice(0, 60), '-> ERROR:', e.message);
    }
  }
}

main().catch(e => { console.error('Fatal:', e.message); process.exit(1); });
