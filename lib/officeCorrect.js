// lib/officeCorrect.js
import JSZip from 'jszip';

const decode = s => s
  .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
const encode = s => s
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&apos;');

function pushExample(examples, before, after, max=12) {
  if (examples.length >= max) return;
  const b = (before||'').slice(0,140);
  const a = (after ||'').slice(0,140);
  if (b !== a) examples.push({ before: b, after: a });
}

// Applique correctFn(string)=>Promise<string> à CHAQUE <w:t> et renvoie stats
export async function correctDOCXText(buffer, correctFn) {
  const zip = await JSZip.loadAsync(buffer);
  const targets = ['word/document.xml', ...Object.keys(zip.files).filter(k => /word\/(header|footer)\d+\.xml$/.test(k))];
  const stats = { totalTextNodes: 0, changedTextNodes: 0, examples: [] };

  for (const p of targets) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async('string');

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
      stats.totalTextNodes++;
      if (corrected !== original) {
        stats.changedTextNodes++;
        pushExample(stats.examples, original, corrected);
      }
      chunks.push(`<w:t${attrs}>${encode(corrected)}</w:t>`);
      lastIndex = start + full.length;
    }
    chunks.push(xml.slice(lastIndex));
    xml = chunks.join('');

    zip.file(p, xml);
  }
  return { outBuffer: await zip.generateAsync({ type: 'nodebuffer' }), stats };
}

// Applique correctFn à CHAQUE <a:t> des slides, renvoie stats
export async function correctPPTXText(buffer, correctFn) {
  const zip = await JSZip.loadAsync(buffer);
  const slides = Object.keys(zip.files).filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k));
  const stats = { totalTextNodes: 0, changedTextNodes: 0, examples: [] };

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
      stats.totalTextNodes++;
      if (corrected !== original) {
        stats.changedTextNodes++;
        pushExample(stats.examples, original, corrected);
      }
      chunks.push(`<a:t>${encode(corrected)}</a:t>`);
      lastIndex = start + full.length;
    }
    chunks.push(xml.slice(lastIndex));
    xml = chunks.join('');

    zip.file(sp, xml);
  }
  return { outBuffer: await zip.generateAsync({ type: 'nodebuffer' }), stats };
}

export default { correctDOCXText, correctPPTXText };
