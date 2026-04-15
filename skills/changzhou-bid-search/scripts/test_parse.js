#!/usr/bin/env node
// 快速调试：检查列表页 HTML 结构
const http = require('http');
const https = require('https');

const url = 'http://ggzy.xzsp.changzhou.gov.cn/jyzx/001001/tradeInfonew.html?category=001001';

function httpGet(url) {
  return new Promise((resolve, reject) => {
    const lib = url.startsWith('https') ? https : http;
    lib.get(url, {headers:{'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}}, res => {
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve(data));
    }).on('error', reject).setTimeout(10000, function() { this.destroy(); reject(new Error('timeout')); });
  });
}

async function main() {
  const html = await httpGet(url);
  console.log('HTML length:', html.length);
  
  // Try different regex
  const re1 = /tzjydetail\('([^']+)',\s*'([^']+)'/g;
  const re2 = /tzjydetail\("([^"]+)",\s*"([^"]+)"/g;
  const re3 = /tzjydetail\(([^)]+)\)/g;
  
  const m1 = [...html.matchAll(re1)];
  console.log('re1 (single quotes) matches:', m1.length);
  const m2 = [...html.matchAll(re2)];
  console.log('re2 (double quotes) matches:', m2.length);
  const m3 = [...html.matchAll(re3)];
  console.log('re3 (any) matches:', m3.length);
  
  if (m1.length > 0) {
    console.log('First match:', m1[0]);
  }
  
  // Show a portion of raw HTML around first tzjydetail
  const idx = html.indexOf('tzjydetail');
  if (idx !== -1) {
    console.log('\nRaw HTML around first tzjydetail:');
    console.log(html.slice(Math.max(0, idx-100), idx+200));
  } else {
    console.log('No tzjydetail found!');
    // Check if there's any onclick
    const onclickMatches = html.match(/onclick=["'][^"']+["']/g);
    console.log('onclick attrs found:', onclickMatches ? onclickMatches.length : 0);
    if (onclickMatches) onclickMatches.slice(0,3).forEach(m => console.log(' ', m));
  }
}

main().catch(e => { console.error(e); process.exit(1); });
