#!/usr/bin/env node
// 调试 parseListHtml 问题
const http = require('http');
const https = require('https');

const BASE_URL = 'ggzy.xzsp.changzhou.gov.cn';

function httpGet(url) {
  return new Promise(function(resolve, reject) {
    const lib = url.startsWith('https') ? https : http;
    lib.get(url, {headers:{'User-Agent':'Mozilla/5.0'}}, function(res) {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve(data));
    }).on('error', reject).setTimeout(10000, function() { this.destroy(); reject(new Error('timeout')); });
  });
}

async function main() {
  const html = await httpGet('http://' + BASE_URL + '/jyzx/001001/tradeInfonew.html?category=001001');
  console.log('HTML length:', html.length);

  // Find all <tr...> positions
  const trStartRe = /<tr[^>]*>/gi;
  const trStarts = [...html.matchAll(trStartRe)].map(m => m.index);
  console.log('tr starts:', trStarts.slice(0, 5), '... count:', trStarts.length);

  const trEndRe = /<\/tr>/gi;
  const trEnds = [...html.matchAll(trEndRe)].map(m => m.index);
  console.log('tr ends:', trEnds.slice(0, 5), '... count:', trEnds.length);

  // Show raw content of first data row (skip thead)
  if (trStarts.length >= 2 && trEnds.length >= 1) {
    // Find first </tr> after first <tr>
    const firstStart = trStarts[0];
    const firstEnd = trEnds.find(e => e >= firstStart);
    console.log('\nFirst <tr>:', html.slice(firstStart, firstStart + 50));
    console.log('First </tr> at:', firstEnd);
    console.log('First row content (50-200):', html.slice(firstStart, firstEnd).slice(50, 200));
    
    // Check: where is 'tzjydetail' in first row?
    const firstRowContent = html.slice(firstStart, firstEnd);
    const firstTdIdx = firstRowContent.indexOf('tzjydetail');
    console.log('\ntzjydetail in first row at offset:', firstTdIdx, 'of first row');
    console.log('First 5 chars around:', firstRowContent.slice(firstTdIdx-20, firstTdIdx+50));
    
    // Now check segs
    const segs = firstRowContent.split('</td>');
    console.log('\nSegs count:', segs.length);
    for (let i = 0; i < Math.min(segs.length, 5); i++) {
      const stripped = segs[i].replace(/<[^>]*>/g, '').trim().slice(0, 40);
      console.log(`  segs[${i}]: "${stripped}"`);
    }
  }
}

main().catch(e => { console.error(e); process.exit(1); });
