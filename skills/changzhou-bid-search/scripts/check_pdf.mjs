import * as pdfModule from 'pdf-parse';
import fs from 'fs';

const cls = pdfModule.PDFParse;
console.log('pdfModule.PDFParse type:', typeof cls);
if (typeof cls === 'function') {
  console.log('PDFParse function name:', cls.name);
  console.log('PDFParse prototype:', Object.getOwnPropertyNames(cls.prototype));
} else {
  console.log('PDFParse is not a function, keys:', Object.keys(cls).slice(0, 10));
}
