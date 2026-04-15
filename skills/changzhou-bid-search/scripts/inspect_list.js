const http = require('http');
const fs = require('fs');

function httpGet(url) {
  return new Promise(function(resolve, reject) {
    http.get(url, {headers: {'User-Agent': 'Mozilla/5.0'}}, function(res) {
      var data = '';
      res.on('data', c => data += c);
      res.on('end', () => resolve(data));
    }).on('error', reject);
  });
}

async function main() {
  // Get tradeInfonew.html
  var html = await httpGet('http://ggzy.xzsp.changzhou.gov.cn/jyzx/001001/tradeInfonew.html');
  
  // Extract all jyzx URLs
  var links = html.match(/href="(\/jyzx\/[^"]+)"/g) || [];
  fs.writeFileSync('/workspace/skills/changzhou-bid-search/list_links.txt', JSON.stringify(links, null, 2));
  console.log('Links found:', links.length);
  
  // Find date patterns in URLs
  var dateLinks = html.match(/href="(\/jyzx\/[^"]*\/\d{8}\/[^"]+)"/g) || [];
  fs.writeFileSync('/workspace/skills/changzhou-bid-search/date_links.txt', JSON.stringify(dateLinks.slice(0, 20), null, 2));
  console.log('Date links found:', dateLinks.length, dateLinks.slice(0, 5));
  
  // Extract text content
  var text = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ');
  fs.writeFileSync('/workspace/skills/changzhou-bid-search/list_text.txt', text.slice(0, 5000));
  console.log('Text sample:', text.slice(0, 2000));
}

main().catch(console.error);
