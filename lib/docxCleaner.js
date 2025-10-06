// lib/docxcleaner.js
import JSZip from 'jszip';
import { normalizeText } from './textCleaner.js';

// helpers
const removeAll = (xml, re) => xml.replace(re, '');
const unwrapTag = (xml, tagName) => {
  const open = new RegExp(`<${tagName}[^>]*>`, 'g');
  const close = new RegExp(`</${tagName}>`, 'g');
  return xml.replace(open, '').replace(close, '');
};

// Nettoie DOCX : supprime commentaires/métadonnées, accepte révisions,
// retire dessins/images, normalise texte dans <w:t>
export async function cleanDOCX(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  // 1) Retirer métadonnées/commentaires/CustomXML
  const removeParts = [
    'docProps/core.xml',
    'docProps/app.xml',
    'docProps/custom.xml',
    'word/comments.xml',
    'word/commentsExtended.xml',
    'customXml/item1.xml',
    'customXml/itemProps1.xml',
  ];
  removeParts.forEach(p => { if (zip.file(p)) zip.remove(p); });

  // 1b) Nettoyer [Content_Types] des overrides de comments
  if (zip.file('[Content_Types].xml')) {
    let ct = await zip.file('[Content_Types].xml').async('string');
    ct = ct.replace(/<Override[^>]*PartName="\/word\/comments(?:Extended)?\.xml"[^>]*\/>/g, '');
    zip.file('[Content_Types].xml', ct);
  }

  // 1c) Supprimer relations vers comments
  const relsPath = 'word/_rels/document.xml.rels';
  if (zip.file(relsPath)) {
    let rels = await zip.file(relsPath).async('string');
    rels = rels.replace(/<Relationship[^>]*Type="[^"]*comments[^"]*"[^>]*\/>/g, '');
    zip.file(relsPath, rels);
  }

  // 2) Cibles XML à modifier (document + headers/footers)
  const targets = [
    'word/document.xml',
    ...Object.keys(zip.files).filter(k => /word\/(header|footer)\d+\.xml$/.test(k)),
  ];

  let aggregatedText = '';

  for (const p of targets) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async('string');

    // a) Retirer marqueurs de commentaires inline
    xml = removeAll(xml, /<w:commentRangeStart[^>]*\/>/g);
    xml = removeAll(xml, /<w:commentRangeEnd[^>]*\/>/g);
    xml = removeAll(xml, /<w:commentReference[^>]*\/>/g);

    // b) Accepter révisions : supprimer deletions, dérouler insertions
    xml = removeAll(xml, /<w:del\b[\s\S]*?<\/w:del>/g);
    xml = unwrapTag(xml, 'w:ins');

    // c) Retirer dessins/images (DrawingML + VML legacy)
    xml = removeAll(xml, /<w:drawing[\s\S]*?<\/w:drawing>/g); // images/shapes
    xml = removeAll(xml, /<w:pict[\s\S]*?<\/w:pict>/g);       // VML pict
    xml = removeAll(xml, /<v:shape[\s\S]*?<\/v:shape>/g);     // VML shapes

    // d) Normaliser texte dans chaque <w:t>…</w:t> (conserve le style)
    xml = xml.replace(/<w:t(?:[^>]*)>([\s\S]*?)<\/w:t>/g, (match, inner) => {
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

    zip.file(p, xml);

    // e) Texte brut pour le report
    aggregatedText += ' ' + xml.replace(/<[^>]+>/g, ' ');
  }

  // 3) Retirer fichiers médias si on a supprimé dessins/images
  for (const k of Object.keys(zip.files)) {
    if (/^word\/media\//.test(k)) zip.remove(k);
  }

  // 4) Binaire DOCX + texte
  const outBuffer = await zip.generateAsync({ type: 'nodebuffer' });
  const text = aggregatedText.replace(/\s{2,}/g, ' ').trim();
  return { outBuffer, text };
}

export default { cleanDOCX };
