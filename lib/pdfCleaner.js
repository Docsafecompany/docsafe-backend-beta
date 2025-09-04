import { PDFDocument, rgb } from 'pdf-lib';


export async function cleanPDF(inputBuffer, { strict = false } = {}) {
const pdf = await PDFDocument.load(inputBuffer);


// 1) Wipe metadata
pdf.setTitle('');
pdf.setAuthor('');
pdf.setCreator('');
pdf.setProducer('');
pdf.setSubject('');
pdf.setKeywords([]);


// 2) Extract text (coarse). pdf-lib n’extrait pas directement; on garde simple
// -> On ne peut pas extraire facilement le texte avec pdf-lib uniquement.
// Fallback: signaler contenu textuel inconnu, mais on renvoie une chaîne vide.
let text = '';


// 3) Strict: tentative de retirer texte blanc/invisible (heuristique: non-op – limitation)
// NOTE: Sans parser/renderer avancé, on ne peut pas modifier finement le contenu.
// On laisse un placeholder pour une V3 avec pdfjs ou ghostscript côté worker.


const outBuffer = await pdf.save();
return { outBuffer, text };
}
