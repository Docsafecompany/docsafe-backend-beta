import { PDFDocument } from 'pdf-lib';

export async function cleanPDF(inputBuffer, { strict = false } = {}) {
  const pdf = await PDFDocument.load(inputBuffer);

  // Wipe metadata
  pdf.setTitle('');
  pdf.setAuthor('');
  pdf.setCreator('');
  pdf.setProducer('');
  pdf.setSubject('');
  pdf.setKeywords([]);

  // Limitation: pas d'extraction fiable du texte ici
  let text = '';

  // Strict PDF (placeholder)
  const outBuffer = await pdf.save();
  return { outBuffer, text };
}
