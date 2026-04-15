const { PDFParse } = require('pdf-parse');
const fs = require('fs');
console.log('PDFParse type:', typeof PDFParse);
console.log('PDFParse keys:', Object.keys(PDFParse));
// Try different API
if (typeof PDFParse === 'function') {
  PDFParse(fs.readFileSync('/workspace/招标文件正文_合成生物产业园.pdf')).then(r => {
    console.log('Pages:', r.pages, 'Text length:', r.text.length);
    fs.writeFileSync('/workspace/skills/changzhou-bid-search/parse_result.txt', r.text);
  }).catch(e => console.error('ERROR:', e.message));
} else {
  // It's a class or object
  const parser = new PDFParse();
  console.log('Instance methods:', Object.keys(Object.getPrototypeOf(parser)));
}
