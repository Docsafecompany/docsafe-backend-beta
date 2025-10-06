// lib/officeCorrect.js
import JSZip from 'jszip';

// helpers XML
const decode = s => s
  .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
const encode = s => s
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&apos;');

// Applique correctFn(string)=>Promise<string> à CHAQUE <w:t> sans changer les balises
export async function correctDOCXText(buffer, correctFn) {
  const zip = await JSZip.loadAsync(buffer);
  const targets = ['word/document.xml', ...Object.keys(zip.files).filter(k => /word\/(header|footer)\d+\.xml$/.test(k))];

  for (const p of targets) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async('string');

    // Remplacement async des <w:t>
    const chunks = [];
    let lastIndex = 0;
    const re = /<w:t([^>]*)>([\s\S]*?)<\/w:t>/g;
    let m;
    while ((m = re.exec(xml)) !== null) {
      const [full, attrs, inner] = m;
      const start = m.index;
      chunks.push(xml.slice(lastIndex, start));
      const original = decode(inner);
      const corrected = await correctFn(original);
      chunks.push(`<w:t${attrs}>${encode(corrected)}</w:t>`);
      lastIndex = start + full.length;
    }
    chunks.push(xml.slice(lastIndex));
    xml = chunks.join('');

    zip.file(p, xml);
  }
  return await zip.generateAsync({ type: 'nodebuffer' });
}

// Applique correctFn à CHAQUE <a:t> des slides
export async function correctPPTXText(buffer, correctFn) {
  const zip = await JSZip.loadAsync(buffer);
  const slides = Object.keys(zip.files).filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k));

  for (const sp of slides) {
    let xml = await zip.file(sp).async('string');

    const chunks = [];
    let lastIndex = 0;
    const re = /<a:t>([\s\S]*?)<\/a:t>/g;
    let m;
    while ((m = re.exec(xml)) !== null) {
      const [full, inner] = m;
      const start = m.index;
      chunks.push(xml.slice(lastIndex, start));
      const original = decode(inner);
      const corrected = await correctFn(original);
      chunks.push(`<a:t>${encode(corrected)}</a:t>`);
      lastIndex = start + full.length;
    }
    chunks.push(xml.slice(lastIndex));
    xml = chunks.join('');

    zip.file(sp, xml);
  }
  return await zip.generateAsync({ type: 'nodebuffer' });
}

export default { correctDOCXText, correctPPTXText };
