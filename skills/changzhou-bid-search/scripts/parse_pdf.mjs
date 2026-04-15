import { PDFParse } from 'pdf-parse';
import fs from 'fs';

const buf = fs.readFileSync('/workspace/招标文件正文_合成生物产业园.pdf');
const parser = new PDFParse({ data: buf });
await parser.load();
const text = await parser.getText();
console.log('getText type:', typeof text, 'isArray:', Array.isArray(text));
if (typeof text === 'string') {
  console.log('Text length:', text.length);
} else if (Array.isArray(text)) {
  console.log('Array length:', text.length, 'First 200 chars:', JSON.stringify(text.slice(0,3)));
} else {
  console.log('Text keys:', Object.keys(text));
}
