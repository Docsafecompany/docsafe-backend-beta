import JSZip from 'jszip';


export async function cleanDOCX(buffer) {
const zip = await JSZip.loadAsync(buffer);
// Remove core/app properties & comments/custom xml
const remove = [
'docProps/core.xml',
'docProps/app.xml',
'docProps/custom.xml',
'word/comments.xml',
'word/commentsExtended.xml',
'customXml/item1.xml', 'customXml/itemProps1.xml'
];
remove.forEach(p => { if (zip.file(p)) zip.remove(p); });


// Extract text (very basic): read document.xml and strip tags
let text = '';
const docXml = await zip.file('word/document.xml')?.async('string');
if (docXml) {
text = docXml.replace(/<[^>]+>/g, ' ') // strip tags
.replace(/\s{2,}/g, ' ')
.trim();
}


const outBuffer = await zip.generateAsync({ type: 'nodebuffer' });
return { outBuffer, text };
}
