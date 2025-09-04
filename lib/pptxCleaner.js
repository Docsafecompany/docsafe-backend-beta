import JSZip from 'jszip';


export async function cleanPPTX(buffer) {
const zip = await JSZip.loadAsync(buffer);
// Remove core/app properties
const remove = [
'docProps/core.xml',
'docProps/app.xml',
'docProps/custom.xml'
];
remove.forEach(p => { if (zip.file(p)) zip.remove(p); });


// Extract text: concatenate slide text
let text = '';
const slideRegex = /^ppt\/slides\/slide\d+\.xml$/;
zip.forEach((relPath, file) => {
if (slideRegex.test(relPath)) {
// naive text extraction
file.async('string').then(xml => {
const t = xml.replace(/<[^>]+>/g, ' ').replace(/\s{2,}/g, ' ').trim();
text += (t ? (t + '\n') : '');
});
}
});
// ATTENTION: zip.forEach async – on regénère après un court délai pour capture
await new Promise(r => setTimeout(r, 100));


const outBuffer = await zip.generateAsync({ type: 'nodebuffer' });
return { outBuffer, text };
}
