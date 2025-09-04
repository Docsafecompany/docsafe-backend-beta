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

  // 3) Normaliser le texte dans CHAQUE <w:t>…</w:t>
  const cleanedXml = xml.replace(/<w:t(?:[^>]*)>([\s\S]*?)<\/w:t>/g, (match, inner) => {
    // Décoder entités XML de base
    const decoded = inner
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'");

    const normalized = normalizeText(decoded);

    // Ré-encoder
    const reEncoded = normalized
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');

    // Conserver la balise d'ouverture telle quelle (xml:space, etc.)
    const openTagEnd = match.indexOf('>') + 1;
    const openTag = match.slice(0, openTagEnd);
    const closeTag = '</w:t>';
    return openTag + reEncoded + closeTag;
  });

  // 4) Écrire le XML nettoyé
  zip.file(p, cleanedXml);

  // 5) Extraire un texte “plat” pour le report
  const text = cleanedXml.replace(/<[^>]+>/g, ' ').replace(/\s{2,}/g, ' ').trim();

  // 6) Recréer le binaire DOCX
  const outBuffer = await zip.generateAsync({ type: 'nodebuffer' });
  return { outBuffer, text };
}
