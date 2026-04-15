import https from 'node:https';
import http from 'node:http';
import fs from 'fs';
import { URL } from 'node:url';

const BASE = 'http://ggzy.xzsp.changzhou.gov.cn';

function fetchHtml(url) {
  return new Promise((resolve, reject) => {
    const parsed = new URL(url);
    const mod = parsed.protocol === 'https:' ? https : http;
    mod.get(url, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'zh-CN,zh;q=0.9',
        'Referer': BASE + '/',
      }
    }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        return fetchHtml(res.headers.location).then(resolve).catch(reject);
      }
      let data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve({ html: data, status: res.statusCode }));
    }).on('error', reject).setTimeout(15000, function() { this.destroy(); reject(new Error('timeout')); });
  });
}

async function main() {
  // First item from list
  const uuid = 'e6a45629-b663-475a-9319-85637b49045c';
  const cat = '001001001001';
  
  const { html, status } = await fetchHtml(`${BASE}/trad_notice.html?infoid=${uuid}&chanquantype=${cat}`);
  
  console.log(`Status: ${status}, Length: ${html.length}`);
  
  // Save raw HTML for inspection
  fs.writeFileSync('/workspace/skills/changzhou-bid-search/outputs/detail_sample.html', html);
  
  // Extract links
  const links = [];
  const re = /href=["']([^"']+)["']/g;
  let m;
  while ((m = re.exec(html)) !== null) {
    if (!m[1].startsWith('javascript') && !m[1].startsWith('#')) {
      links.push(m[1]);
    }
  }
  console.log('\nLinks (' + links.length + '):');
  links.forEach(l => console.log(' ', l));
  
  // Extract title
  const titleM = html.match(/<title>([^<]+)<\/title>/i);
  console.log('\nTitle:', titleM ? titleM[1] : 'none');
  
  // Clean text
  const text = html
    .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ').trim();
  
  fs.writeFileSync('/workspace/skills/changzhou-bid-search/outputs/detail_sample.txt', text);
  console.log('\nText length:', text.length);
  console.log('First 1500 chars:\n', text.slice(0, 1500));
}

main().catch(e => console.error('Error:', e.message));
