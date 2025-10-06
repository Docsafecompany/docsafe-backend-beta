// lib/docxCleaner.js
import JSZip from 'jszip';

const removeAll = (xml, re) => xml.replace(re, '');
const unwrapTag = (xml, tag) =>
  xml.replace(new RegExp(`<${tag}[^>]*>`, 'g'), '').replace(new RegExp(`</${tag}>`, 'g'), '');

/**
 * drawPolicy:
 *  - "auto" (default): supprime ink/doodles (a14:ink, VML) ; conserve <w:drawing> (images/logos)
 *  - "all": supprime tous dessins/images (w:drawing, pict, v:shape) => texte-only visuel
 *  - "none": ne supprime rien côté dessins/images
 */
export async function cleanDOCX(buffer, { drawPolicy = "auto" } = {}) {
  const zip = await JSZip.loadAsync(buffer);

  // 1) Privacy parts
  [
    'docProps/core.xml','docProps/app.xml','docProps/custom.xml',
    'word/comments.xml','word/commentsExtended.xml',
    'customXml/item1.xml','customXml/itemProps1.xml',
  ].forEach(p => { if (zip.file(p)) zip.remove(p); });

  // Content types (comments)
  if (zip.file('[Content_Types].xml')) {
    let ct = await zip.file('[Content_Types].xml').async('string');
    ct = ct.replace(/<Override[^>]*PartName="\/word\/comments(?:Extended)?\.xml"[^>]*\/>/g, '');
    zip.file('[Content_Types].xml', ct);
  }

  // Rels vers comments
  const rels = 'word/_rels/document.xml.rels';
  if (zip.file(rels)) {
    let relXml = await zip.file(rels).async('string');
    relXml = relXml.replace(/<Relationship[^>]*Type="[^"]*comments[^"]*"[^>]*\/>/g, '');
    zip.file(rels, relXml);
  }

  // 2) Documents cibles
  const targets = ['word/document.xml', ...Object.keys(zip.files).filter(k => /word\/(header|footer)\d+\.xml$/.test(k))];

  for (const p of targets) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async('string');

    // Comments in-text + tracked changes
    xml = removeAll(xml, /<w:commentRangeStart[^>]*\/>/g);
    xml = removeAll(xml, /<w:commentRangeEnd[^>]*\/>/g);
    xml = removeAll(xml, /<w:commentReference[^>]*\/>/g);
    xml = removeAll(xml, /<w:del\b[\s\S]*?<\/w:del>/g); // supprime deletions
    xml = unwrapTag(xml, 'w:ins');                      // garde le texte des insertions

    // Dessins selon policy
    if (drawPolicy === "all") {
      xml = removeAll(xml, /<w:drawing[\s\S]*?<\/w:drawing>/g);
      xml = removeAll(xml, /<w:pict[\s\S]*?<\/w:pict>/g);
      xml = removeAll(xml, /<v:shape[\s\S]*?<\/v:shape>/g);
      xml = removeAll(xml, /<a:graphicData[^>]*uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/ink"[\s\S]*?<\/a:graphicData>/g);
      xml = removeAll(xml, /<a14:ink[\s\S]*?<\/a14:ink>/g);
    } else if (drawPolicy === "auto") {
      // enlève ink/doodles et vieux VML; garde <w:drawing>
      xml = removeAll(xml, /<a:graphicData[^>]*uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/ink"[\s\S]*?<\/a:graphicData>/g);
      xml = removeAll(xml, /<a14:ink[\s\S]*?<\/a14:ink>/g);
      xml = removeAll(xml, /<w:pict[\s\S]*?<\/w:pict>/g);
      xml = removeAll(xml, /<v:shape[\s\S]*?<\/v:shape>/g);
    }

    zip.file(p, xml);
  }

  if (drawPolicy === "all") {
    for (const k of Object.keys(zip.files)) {
      if (/^word\/media\//.test(k)) zip.remove(k);
    }
  }

  return { outBuffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

export default { cleanDOCX };
