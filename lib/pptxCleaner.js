// lib/pptxcleaner.js
import JSZip from 'jszip';
import { normalizeText } from './textCleaner.js';

// Nettoie PPTX : supprime métadonnées + threads de commentaires,
// enlève images/ink, normalise <a:t>, et nettoie ppt/media/*
export async function cleanPPTX(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  // 1) Retirer métadonnées
  ['docProps/core.xml', 'docProps/app.xml', 'docProps/custom.xml'].forEach(p => {
    if (zip.file(p)) zip.remove(p);
  });

  // 2) Supprimer commentaires & auteurs
  for (const k of Object.keys(zip.files)) {
    if (/^ppt\/comments\/.+\.xml$/.test(k)) zip.remove(k);
  }
  if (zip.file('ppt/commentAuthors.xml')) zip.remove('ppt/commentAuthors.xml');

  // 3) Purger relations de commentaires sur chaque slide
  for (const rels of Object.keys(zip.files).filter(k => /ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(k))) {
    let relXml = await zip.file(rels).async('string');
    relXml = relXml.replace(/<Relationship[^>]*Type="[^"]*comments[^"]*"[^>]*\/>/g, '');
    zip.file(rels, relXml);
  }

  // 4) Lister slides
  const slideRegex = /^ppt\/slides\/slide\d+\.xml$/;
  const slidePaths = [];
  zip.forEach((relPath) => { if (slideRegex.test(relPath)) slidePaths.push(relPath); });

  // 5) Nettoyer chaque slide : enlever images/ink + normaliser <a:t>
  for (const sp of slidePaths) {
    let xml = await zip.file(sp).async('string');

    // a) Enlever ink (Office 2010)
    xml = xml.replace(/<a:graphicData[^>]*uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/ink"[\s\S]*?<\/a:graphicData>/g, '');
    xml = xml.replace(/<a14:ink[\s\S]*?<\/a14:ink>/g, '');

    // b) Enlever images (conserve zones de texte et formes sans image)
    xml = xml.replace(/<p:pic[\s\S]*?<\/p:pic>/g, '');

    // c) Normaliser le texte dans <a:t>
    xml = xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (_m, inner) => {
      const decoded = inner
        .replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
        .replace(/&quot;/g,'"').replace(/&apos;/g,"'");
      const normalized = normalizeText(decoded);
      const enc = normalized
        .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
        .replace(/"/g,'&quot;').replace(/'/g,'&apos;');
      return `<a:t>${enc}</a:t>`;
    });

    zip.file(sp, xml);
  }

  // 6) Retirer médias si images supprimées
  for (const k of Object.keys(zip.files)) {
    if (/^ppt\/media\//.test(k)) zip.remove(k);
  }

  // 7) Construire texte pour report
  let text = '';
  for (const sp of slidePaths) {
    const xml = await zip.file(sp).async('string');
    text += xml.replace(/<[^>]+>/g, ' ').replace(/\s{2,}/g, ' ').trim() + '\n';
  }

  // 8) Binaire PPTX
  const outBuffer = await zip.generateAsync({ type: 'nodebuffer' });
  return { outBuffer, text: text.trim() };
}

export default { cleanPPTX };
