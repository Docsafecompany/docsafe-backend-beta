// lib/pdfCleaner.js
import { PDFDocument, PDFName, StandardFonts } from 'pdf-lib';

async function buildTextOnlyPdf(text, title = "DocSafe Text-Only") {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const margin = 48, pageW = 595.28, pageH = 841.89, maxW = pageW - margin*2, lh = 14, size = 11;
  let page = doc.addPage([pageW, pageH]); let y = pageH - margin;
  const words = (text || '').replace(/\s+/g, ' ').trim().split(' ').filter(Boolean);
  const w = s => font.widthOfTextAtSize(s, size);
  let line = '';
  for (const word of words) {
    const test = line ? line + ' ' + word : word;
    if (w(test) <= maxW) { line = test; continue; }
    if (y - lh < margin) { page = doc.addPage([pageW, pageH]); y = pageH - margin; }
    if (line) { page.drawText(line, { x: margin, y: y - lh, size, font }); y -= lh; line = word; }
    else { page.drawText(word, { x: margin, y: y - lh, size, font }); y -= lh; }
  }
  if (line) {
    if (y - lh < margin) { page = doc.addPage([pageW, pageH]); y = pageH - margin; }
    page.drawText(line, { x: margin, y: y - lh, size, font });
  }
  doc.setTitle(title); doc.setCreator("DocSafe");
  return Buffer.from(await doc.save());
}

/**
 * pdfMode:
 *  - "sanitize" (default): wipe metadata, remove annotations & embedded files (visuel intact)
 *  - "text-only": reconstruit un PDF texte-only (supprime tout graphisme)
 * extractTextFn(inputBuffer) optionnel pour text-only
 */
export async function cleanPDF(inputBuffer, { pdfMode = "sanitize", extractTextFn } = {}) {
  if (pdfMode === "text-only") {
    const text = typeof extractTextFn === 'function' ? await extractTextFn(inputBuffer) : '';
    const out = await buildTextOnlyPdf(text);
    return { outBuffer: out, text };
  }

  const pdf = await PDFDocument.load(inputBuffer);
  // metadata
  pdf.setTitle(''); pdf.setAuthor(''); pdf.setCreator(''); pdf.setProducer(''); pdf.setSubject(''); pdf.setKeywords([]);
  // annots
  for (const page of pdf.getPages()) {
    page.node.set(PDFName.of('Annots'), pdf.context.obj([]));
  }
  // embedded files
  const catalog = pdf.catalog;
  const Names = catalog.get(PDFName.of('Names'));
  if (Names) {
    const maybeEF = Names.lookupMaybe(PDFName.of('EmbeddedFiles'));
    if (maybeEF) Names.set(PDFName.of('EmbeddedFiles'), pdf.context.obj({ Names: [] }));
  }

  return { outBuffer: Buffer.from(await pdf.save()), text: '' };
}

export default { cleanPDF };
