import JSZip from 'jszip';
import { normalizeText } from './textCleaner.js';

// Nettoie DOCX en conservant la mise en forme : on traite chaque <w:t>
export async function cleanDOCX(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  // 1) Retirer métadonnées/commentaires
  const remove = [
    'docProps/core.xml',
    'docProps/app.xml',
    'docProps/custom.xml',
    'word/comments.xml',
    'word/commentsExtended.xml',
    'customXml/item1.xml', 'customXml/itemProps1.xml'
  ];
  remove.forEach(p => { if (zip.file(p)) zip.remove(p); });

  // 2) Charger document.xml
  const p = 'word/document.xml';
  const xml = await zip.file(p)?.async('string');
  if (!xml) {
    const outBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    return { outBuffer, text: '' };
  }

  // 3) Normaliser texte dans chaque <w:t>…</w:t>
  const cleanedXml = xml.replace(/<w:t(?:[^>]*)>([\s\S]*?)<\/w:t>/g, (match, inner) => {
    const decoded = inner
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");
    const normalized = normalizeText(decoded);
    const reEncoded = normalized
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
    const openTagEnd = match.indexOf('>') + 1;
    const openTag = match.slice(0, openTagEnd);
    return openTag + reEncoded + '</w:t>';
  });

  // 4) Écrire & extraire texte pour report
  zip.file(p, cleanedXml);
  const text = cleanedXml.replace(/<[^>]+>/g, ' ').replace(/\s{2,}/g, ' ').trim();

  // 5) Binaire DOCX
  const outBuffer = await zip.generateAsync({ type: 'nodebuffer' });
  return { outBuffer, text };
}
