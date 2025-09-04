import fs from 'fs';
import JSZip from 'jszip';
import mime from 'mime-types';

export async function detectMime(buffer, filename) {
const byExt = mime.lookup(filename) || '';
// simple fallback by magic number if needed
if (buffer?.length > 4) {
const head = buffer.slice(0, 4).toString('hex');
if (head === '25504446') return 'application/pdf';
if (head === '504b0304') {
// could be docx/pptx
if ((filename||'').endsWith('.docx')) return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
if ((filename||'').endsWith('.pptx')) return 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
}
}
return byExt || 'application/octet-stream';
}


export function inferExt(filename, mimeType) {
const extFromMime = mime.extension(mimeType || '') || '';
const extFromName = (filename || '').split('.').pop();
if (extFromMime) return `.${extFromMime}`;
if (extFromName) return `.${extFromName}`;
return '.bin';
}


export async function zipOutput(files) {
const zip = new JSZip();
for (const f of files) {
zip.file(f.name, f.data);
}
return await zip.generateAsync({ type: 'nodebuffer' });
}


export function readFileBuffer(p) { return fs.promises.readFile(p); }
export function writeFileBuffer(p, b) { return fs.promises.writeFile(p, b); }
