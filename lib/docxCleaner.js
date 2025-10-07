// lib/docxCleaner.js
import JSZip from 'jszip';

const removeAllCount = (xml, re) => {
  let count = 0;
  xml = xml.replace(re, () => { count++; return ''; });
  return { xml, count };
};
const unwrapTagCount = (xml, tag) => {
  let count = 0;
  const open = new RegExp(`<${tag}[^>]*>`, 'g');
  const close = new RegExp(`</${tag}>`, 'g');
  xml = xml.replace(open, () => { count++; return ''; });
  xml = xml.replace(close, () => { count++; return ''; });
  return { xml, countOpenClose: count };
};

/**
 * drawPolicy:
 *  - "auto" (default): supprime ink/doodles (a14:ink, VML) ; conserve <w:drawing> (images/logos)
 *  - "all": supprime tous dessins/images (w:drawing, pict, v:shape) => texte-only visuel
 *  - "none": ne supprime rien côté dessins/images
 */
export async function cleanDOCX(buffer, { drawPolicy = "auto" } = {}) {
  const stats = {
    metaRemoved: 0,
    commentsXmlRemoved: 0,
    commentMarkersRemoved: 0,
    revisionsAccepted: { deletionsRemoved: 0, insertionsUnwrapped: 0 },
    drawingsRemoved: 0,
    vmlRemoved: 0,
    inkRemoved: 0,
    mediaDeleted: 0,
  };

  const zip = await JSZip.loadAsync(buffer);

  // 1) Privacy parts
  const parts = [
    'docProps/core.xml','docProps/app.xml','docProps/custom.xml',
    'word/comments.xml','word/commentsExtended.xml',
    'customXml/item1.xml','customXml/itemProps1.xml',
  ];
  parts.forEach(p => {
    if (zip.file(p)) {
      zip.remove(p); stats.metaRemoved++;
      if (p.includes('comments')) stats.commentsXmlRemoved++;
    }
  });

  // 2) [Content_Types] overrides de comments
  if (zip.file('[Content_Types].xml')) {
    let ct = await zip.file('[Content_Types].xml').async('string');
    const before = ct.length;
    ct = ct.replace(/<Override[^>]*PartName="\/word\/comments(?:Extended)?\.xml"[^>]*\/>/g, '');
    if (ct.length !== before) stats.commentsXmlRemoved++;
    zip.file('[Content_Types].xml', ct);
  }

  // 3) Rels vers comments
  const rels = 'word/_rels/document.xml.rels';
  if (zip.file(rels)) {
    let relXml = await zip.file(rels).async('string');
    const n = (relXml.match(/Type="[^"]*comments[^"]*"/g) || []).length;
    if (n) {
      relXml = relXml.replace(/<Relationship[^>]*Type="[^"]*comments[^"]*"[^>]*\/>/g, '');
      stats.commentsXmlRemoved += n;
      zip.file(rels, relXml);
    }
  }

  // 4) XML cibles (document + headers/footers)
  const targets = ['word/document.xml', ...Object.keys(zip.files).filter(k => /word\/(header|footer)\d+\.xml$/.test(k))];

  for (const p of targets) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async('string');

    // a) Marqueurs de commentaires
    let out = removeAllCount(xml, /<w:commentRangeStart[^>]*\/>/g); xml = out.xml; stats.commentMarkersRemoved += out.count;
    out = removeAllCount(xml, /<w:commentRangeEnd[^>]*\/>/g);       xml = out.xml; stats.commentMarkersRemoved += out.count;
    out = removeAllCount(xml, /<w:commentReference[^>]*\/>/g);      xml = out.xml; stats.commentMarkersRemoved += out.count;

    // b) Révisions (accepter)
    xml = xml.replace(/<w:del\b[\s\S]*?<\/w:del>/g, () => { stats.revisionsAccepted.deletionsRemoved++; return ''; });
    const unwrap = unwrapTagCount(xml, 'w:ins'); xml = unwrap.xml;
    stats.revisionsAccepted.insertionsUnwrapped += unwrap.countOpenClose ? Math.ceil(unwrap.countOpenClose/2) : 0;

    // c) Dessins selon policy
    if (drawPolicy === "all") {
      out = removeAllCount(xml, /<w:drawing[\s\S]*?<\/w:drawing>/g); xml = out.xml; stats.drawingsRemoved += out.count;
      out = removeAllCount(xml, /<w:pict[\s\S]*?<\/w:pict>/g);       xml = out.xml; stats.vmlRemoved     += out.count;
      out = removeAllCount(xml, /<v:shape[\s\S]*?<\/v:shape>/g);     xml = out.xml; stats.vmlRemoved     += out.count;
      out = removeAllCount(xml, /<a:graphicData[^>]*uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/ink"[\s\S]*?<\/a:graphicData>/g);
      xml = out.xml; stats.inkRemoved += out.count;
      out = removeAllCount(xml, /<a14:ink[\s\S]*?<\/a14:ink>/g);     xml = out.xml; stats.inkRemoved     += out.count;
    } else if (drawPolicy === "auto") {
      // enlève ink/doodles et vieux VML; garde <w:drawing>
      out = removeAllCount(xml, /<a:graphicData[^>]*uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/ink"[\s\S]*?<\/a:graphicData>/g);
      xml = out.xml; stats.inkRemoved += out.count;
      out = removeAllCount(xml, /<a14:ink[\s\S]*?<\/a14:ink>/g);     xml = out.xml; stats.inkRemoved += out.count;
      out = removeAllCount(xml, /<w:pict[\s\S]*?<\/w:pict>/g);       xml = out.xml; stats.vmlRemoved += out.count;
      out = removeAllCount(xml, /<v:shape[\s\S]*?<\/v:shape>/g);     xml = out.xml; stats.vmlRemoved += out.count;
    } // "none": rien

    zip.file(p, xml);
  }

  // 5) Médias orphelins si drawPolicy=all
  if (drawPolicy === "all") {
    for (const k of Object.keys(zip.files)) {
      if (/^word\/media\//.test(k)) { zip.remove(k); stats.mediaDeleted++; }
    }
  }

  return { outBuffer: await zip.generateAsync({ type: 'nodebuffer' }), stats };
}

export default { cleanDOCX };
