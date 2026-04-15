import { PDFParse } from 'pdf-parse';
import { createWriteStream, readFileSync } from 'fs';
import https from 'https';
import http from 'http';

function download(url, dest) {
  return new Promise(function(resolve, reject) {
    var lib = url.startsWith('https') ? https : http;
    var file = createWriteStream(dest);
    lib.get(url, {headers: {'User-Agent': 'Mozilla/5.0'}}, function(res) {
      res.pipe(file);
      file.on('finish', resolve);
    }).on('error', reject);
  });
}

async function main() {
  var pdfUrl = 'http://ggzy.xzsp.changzhou.gov.cn/czggzyweb/WebbuilderMIS/attach/downloadZtbAttach.jspx?attachGuid=07bf6cef-7769-46a9-9dfe-2759c94536b4&appUrlFlag=ztb007&siteGuid=7eb5f7f1-9041-43ad-8e13-8fcb82ea831a';
  var dest = '/workspace/招标文件正文_未来智慧城.pdf';
  
  await download(pdfUrl, dest);
  var buf = readFileSync(dest);
  console.log('PDF downloaded, size:', buf.length);
  
  var parser = new PDFParse({data: buf});
  await parser.load();
  var result = await parser.getText();
  var text = result.text;
  console.log('Pages:', result.pages, '| Chars:', text.length);
  
  var keywords = ['3.4.1', '保证金', '投标保证金', '12.4.1', '付款周期', '付款', '2.1.2', '评标入围', '评审'];
  for (var kw of keywords) {
    var idx = text.indexOf(kw);
    if (idx !== -1) {
      console.log('\n=== [' + kw + '] ===');
      console.log(text.slice(Math.max(0, idx - 20), idx + 400));
    }
  }
}

main().catch(function(e) { console.error('ERROR:', e.message, e.stack); process.exit(1); });
