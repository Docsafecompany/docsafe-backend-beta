import JSZip from 'jszip';
import { normalizeText } from './textCleaner.js';

export async function cleanPPTX(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  // 1) Retirer métadonnées
  ['docProps/core.xml', 'docProps/app.xml', 'docProps/custom.xml'].forEach(p => {
    if (zip.file(p)) zip.remove(p);
  });

  // 2) Lister slides
  const slideRegex = /^ppt\/slides\/slide\d+\.xml$/;
  const slidePaths = [];
  zip.forEach((relPath) => { if (slideRegex.test(relPath)) slidePaths.push(relPath); });

  // 3) Normaliser texte dans chaque <a:t>
  for (const sp of slidePaths) {
    const xml = await zip.file(sp).async('string');
    const cleanedXml = xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (_m, inner) => {
      const decoded = inner
        .replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
        .replace(/&quot;/g,'"').replace(/&apos;/g,"'");
      const normalized = normalizeText(decoded);
      const enc = normalized
        .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
        .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
      return `<a:t>${enc}</a:t>`;
    });
    zip.file(sp, cleanedXml);
  }

  // 4) Texte pour report
  let text = '';
  for (const sp of slidePaths) {
    const xml = await zip.file(sp).async('string');
    text += xml.replace(/<[^>]+>/g, ' ').replace(/\s{2,}/g, ' ').trim() + '\n';
  }

  // 5) Binaire PPTX
  const outBuffer = await zip.generateAsync({ type: 'nodebuffer' });
  return { outBuffer, text: text.trim() };
}

