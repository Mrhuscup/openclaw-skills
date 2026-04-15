import { PDFParse } from 'pdf-parse';
import fs from 'fs';

const buf = fs.readFileSync('/workspace/招标文件正文_合成生物产业园.pdf');
const parser = new PDFParse({ data: buf });
await parser.load();
const result = await parser.getText();
const text = result.text;
console.log('Pages:', result.pages, '| Total chars:', result.total);

const keywords = [
  '2.1.2', '12.4.1', '投标人须知前附表', '保证金',
  '付款周期', '3.4资格审查可选条件', '企业业绩', '项目经理业绩',
  '评标办法', '投标保证金', '入围方法'
];

const findings = [];
for (const kw of keywords) {
  const idx = text.indexOf(kw);
  if (idx !== -1) {
    findings.push(`\n=== [${kw}] 位置:${idx} ===\n${text.slice(Math.max(0, idx - 20), idx + 400)}`);
    console.log(`[${kw}] found at ${idx}`);
  }
}

fs.writeFileSync('/workspace/skills/changzhou-bid-search/pdf_result.txt', findings.join('\n'));
console.log('\nDone. Sections:', findings.length);
