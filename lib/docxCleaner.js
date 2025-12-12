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

// ============================================================
// HIDDEN / WHITE TEXT DETECTION (DOCX)
// ============================================================

function decodeXmlText(s = '') {
  return String(s)
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)));
}

function extractRunText(runXml) {
  const texts = [];
  const re = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(runXml))) {
    texts.push(decodeXmlText(m[1] || ''));
  }
  return texts.join('');
}

function getRunPr(runXml) {
  const m = runXml.match(/<w:rPr[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : '';
}

function isVanished(runPrXml) {
  return /<w:(vanish|specVanish)\b[^\/]*\/?>/i.test(runPrXml);
}

function isWhiteText(runPrXml) {
  // w:color val="FFFFFF" or w:val="FFFFFF"
  const m =
    runPrXml.match(/<w:color\b[^>]*w:val="([^"]+)"/i) ||
    runPrXml.match(/<w:color\b[^>]*val="([^"]+)"/i);

  if (!m) return false;
  const val = String(m[1] || '').replace('#', '').toUpperCase();
  return val === 'FFFFFF';
}

function scanHiddenTextInPart(xml, partName) {
  const detections = [];

  const paraRe = /<w:p\b[\s\S]*?<\/w:p>/g;
  let pIndex = 0;
  let pMatch;

  while ((pMatch = paraRe.exec(xml))) {
    const pXml = pMatch[0];

    const runRe = /<w:r\b[\s\S]*?<\/w:r>/g;
    let rIndex = 0;
    let rMatch;

    while ((rMatch = runRe.exec(pXml))) {
      const rXml = rMatch[0];
      const rPr = getRunPr(rXml);

      const vanished = isVanished(rPr);
      const white = isWhiteText(rPr);

      if (!vanished && !white) { rIndex++; continue; }

      const text = extractRunText(rXml).trim();
      if (!text) { rIndex++; continue; }

      detections.push({
        type: white ? 'white_text' : 'vanished_text',
        reason: white ? 'white_color' : 'vanish_property',
        content: text,
        location: {
          part: partName,      // word/document.xml, word/header1.xml, etc.
          paragraph: pIndex,
          run: rIndex
        }
      });

      rIndex++;
    }

    pIndex++;
  }

  return detections;
}

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

    // NEW
    hiddenTextFound: 0,
  };

  // NEW: detections returned for UI/reporting
  const detections = {
    hiddenContent: [],
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
  const targets = [
    'word/document.xml',
    ...Object.keys(zip.files).filter(k => /word\/(header|footer)\d+\.xml$/.test(k))
  ];

  for (const p of targets) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async('string');

    // NEW: detect hidden/white text BEFORE cleaning
    const hiddenFound = scanHiddenTextInPart(xml, p);
    if (hiddenFound.length) {
      detections.hiddenContent.push(...hiddenFound);
      stats.hiddenTextFound += hiddenFound.length;
    }

    // a) Marqueurs de commentaires
    let out = removeAllCount(xml, /<w:commentRangeStart[^>]*\/>/g); xml = out.xml; stats.commentMarkersRemoved += out.count;
    out = removeAllCount(xml, /<w:commentRangeEnd[^>]*\/>/g);       xml = out.xml; stats.commentMarkersRemoved += out.count;
    out = removeAllCount(xml, /<w:commentReference[^>]*\/>/g);      xml = out.xml; stats.commentMarkersRemoved += out.count;

    // b) Révisions (accepter)
    xml = xml.replace(/<w:del\b[\s\S]*?<\/w:del>/g, () => { stats.revisionsAccepted.deletionsRemoved++; return ''; });
    const unwrap = unwrapTagCount(xml, 'w:ins'); xml = unwrap.xml;
    stats.revisionsAccepted.insertionsUnwrapped += unwrap.countOpenClose ? Math.ceil(unwrap.countOpenClose / 2) : 0;

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

  return {
    outBuffer: await zip.generateAsync({ type: 'nodebuffer' }),
    stats,
    detections, // NEW
  };
}

export default { cleanDOCX };
