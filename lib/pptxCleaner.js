// lib/pptxCleaner.js
import JSZip from 'jszip';

/**
 * drawPolicy:
 *  - "auto" (default): enlève ink (a14:ink) ; conserve p:pic (images/logos)
 *  - "all": supprime p:pic (images) + ink => slides texte-only
 *  - "none": ne touche pas aux dessins/images
 */
export async function cleanPPTX(buffer, { drawPolicy = "auto" } = {}) {
  const stats = {
    metaRemoved: 0,
    commentsXmlRemoved: 0,
    relsCommentsRemoved: 0,
    inkRemoved: 0,
    picturesRemoved: 0,
    mediaDeleted: 0,
  };

  const zip = await JSZip.loadAsync(buffer);

  // 1) Privacy
  ['docProps/core.xml','docProps/app.xml','docProps/custom.xml'].forEach(p => { if (zip.file(p)) { zip.remove(p); stats.metaRemoved++; } });
  for (const k of Object.keys(zip.files)) {
    if (/^ppt\/comments\/.+\.xml$/.test(k)) { zip.remove(k); stats.commentsXmlRemoved++; }
  }
  if (zip.file('ppt/commentAuthors.xml')) { zip.remove('ppt/commentAuthors.xml'); stats.commentsXmlRemoved++; }

  // Rels vers comments
  for (const rels of Object.keys(zip.files).filter(k => /ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(k))) {
    let relXml = await zip.file(rels).async('string');
    const n = (relXml.match(/Type="[^"]*comments[^"]*"/g) || []).length;
    if (n) {
      relXml = relXml.replace(/<Relationship[^>]*Type="[^"]*comments[^"]*"[^>]*\/>/g, '');
      stats.relsCommentsRemoved += n;
      zip.file(rels, relXml);
    }
  }

  // 2) Slides
  for (const sp of Object.keys(zip.files).filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k))) {
    let xml = await zip.file(sp).async('string');

    if (drawPolicy === "all" || drawPolicy === "auto") {
      xml = xml.replace(/<a14:ink[\s\S]*?<\/a14:ink>/g, () => { stats.inkRemoved++; return ''; });
      xml = xml.replace(/<a:graphicData[^>]*uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/ink"[\s\S]*?<\/a:graphicData>/g, () => { stats.inkRemoved++; return ''; });
    }
    if (drawPolicy === "all") {
      xml = xml.replace(/<p:pic[\s\S]*?<\/p:pic>/g, () => { stats.picturesRemoved++; return ''; });
    }

    zip.file(sp, xml);
  }

  if (drawPolicy === "all") {
    for (const k of Object.keys(zip.files)) {
      if (/^ppt\/media\//.test(k)) { zip.remove(k); stats.mediaDeleted++; }
    }
  }

  return { outBuffer: await zip.generateAsync({ type: 'nodebuffer' }), stats };
}

export default { cleanPPTX };
