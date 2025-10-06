// lib/pdfcleaner.js
import { PDFDocument, PDFName } from 'pdf-lib';

// Nettoie PDF : efface métadonnées, annotations (Annots), et EmbeddedFiles
export async function cleanPDF(inputBuffer, { strict = false } = {}) {
  const pdf = await PDFDocument.load(inputBuffer);

  // 1) Wipe metadata (Info + XMP le cas échéant)
  pdf.setTitle('');
  pdf.setAuthor('');
  pdf.setCreator('');
  pdf.setProducer('');
  pdf.setSubject('');
  pdf.setKeywords([]);

  // 2) Supprimer annotations visibles
  for (const page of pdf.getPages()) {
    page.node.set(PDFName.of('Annots'), pdf.context.obj([]));
  }

  // 3) Supprimer les pièces jointes intégrées (EmbeddedFiles) si présentes
  const catalog = pdf.catalog; // Root
  const Names = catalog.get(PDFName.of('Names'));
  if (Names) {
    const maybeEF = Names.lookupMaybe(PDFName.of('EmbeddedFiles'));
    if (maybeEF) {
      Names.set(PDFName.of('EmbeddedFiles'), pdf.context.obj({ Names: [] }));
    }
  }

  // 4) Strict mode placeholder (ex : flatten forms, etc. si besoin)
  // (non requis pour l’instant)

  const outBuffer = await pdf.save();
  return { outBuffer: Buffer.from(outBuffer), text: '' };
}

export default { cleanPDF };
