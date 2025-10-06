// lib/pptxCleaner.js
import JSZip from 'jszip';

/**
 * drawPolicy:
 *  - "auto" (default): enlève ink (a14:ink) ; conserve p:pic (images/logos)
 *  - "all": supprime p:pic (images) + ink => slides texte-only
 *  - "none": ne touche pas aux dessins/images
 */
export async function cleanPPTX(buffer, { drawPolicy = "auto" } = {}) {
  const zip = await JSZip.loadAsync(buffer);

  // 1) Privacy
  ['docProps/core.xml','docProps/app.xml','docProps/custom.xml'].forEach(p => { if (zip.file(p)) zip.remove(p); });
  for (const k of Object.keys(zip.files)) if (/^ppt\/comments\/.+\.xml$/.test(k)) zip.remove(k);
  if (zip.file('ppt/commentAuthors.xml')) zip.remove('ppt/commentAuthors.xml');
  for (const rels of Object.keys(zip.files).filter(k => /ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(k))) {
    let relXml = await zip.file(rels).async('string');
    relXml = relXml.replace(/<Relationship[^>]*Type="[^"]*comments[^"]*"[^>]*\/>/g, '');
    zip.file(rels, relXml);
  }

  // 2) Slides
  for (const sp of Object.keys(zip.files).filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k))) {
    let xml = await zip.file(sp).async('string');

    if (drawPolicy === "all" || drawPolicy === "auto") {
      xml = xml.replace(/<a:graphicData[^>]*uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/ink"[\s\S]*?<\/a:graphicData>/g, '');
      xml = xml.replace(/<a14:ink[\s\S]*?<\/a14:ink>/g, '');
    }
    if (drawPolicy === "all") {
      xml = xml.replace(/<p:pic[\s\S]*?<\/p:pic>/g, '');
    }

    zip.file(sp, xml);
  }

  if (drawPolicy === "all") {
    for (const k of Object.keys(zip.files)) {
      if (/^ppt\/media\//.test(k)) zip.remove(k);
    }
  }

  return { outBuffer: await zip.generateAsync({ type: 'nodebuffer' }) };
}

export default { cleanPPTX };
