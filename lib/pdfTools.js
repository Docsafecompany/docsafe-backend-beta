// lib/pdfTools.js
import { PDFDocument } from "pdf-lib";

// --- IMPORTANT ---
// On ne charge PAS pdf-parse au top-level pour éviter l'ENOENT.
// On le charge dynamiquement depuis son module interne "lib/pdf-parse.js".
async function getPdfParse() {
  try {
    // Certaines versions exportent en CJS, d’autres en ESM — on gère les deux.
    const mod = await import("pdf-parse/lib/pdf-parse.js");
    return (mod && (mod.default || mod)) || null;
  } catch (e1) {
    try {
      // Fallback éventuel si l’export interne change
      const mod2 = await import("pdf-parse");
      return (mod2 && (mod2.default || mod2)) || null;
    } catch (e2) {
      console.warn("pdf-parse unavailable, PDF text extraction will be skipped.", e1?.message || e1, e2?.message || e2);
      return null;
    }
  }
}

export async function stripPdfMetadata(buffer) {
  const pdfDoc = await PDFDocument.load(buffer, { updateMetadata: true });
  try { pdfDoc.setTitle(""); } catch {}
  try { pdfDoc.setAuthor(""); } catch {}
  try { pdfDoc.setSubject(""); } catch {}
  try { pdfDoc.setKeywords([]); } catch {}
  try { pdfDoc.setProducer(""); } catch {}
  try { pdfDoc.setCreator(""); } catch {}
  try { pdfDoc.setCreationDate(undefined); } catch {}
  try { pdfDoc.setModificationDate(new Date()); } catch {}
  return await pdfDoc.save();
}

export async function extractPdfText(buffer) {
  const pdfParse = await getPdfParse();
  if (!pdfParse) return "";
  try {
    const data = await pdfParse(buffer);
    return String(data?.text || "");
  } catch (e) {
    console.warn("pdf-parse failed at runtime; returning empty text.", e?.message || e);
    return "";
  }
}

// Heuristique : filtre texte “bruit” (watermarks/headers/lignes courtes)
export function filterExtractedLines(text, { strictPdf = false } = {}) {
  const lines = String(text || "").replace(/\r\n/g, "\n").split("\n");
  return lines
    .filter((ln) => {
      const s = ln.trim();
      if (!s) return false;
      if (s.length <= 1) return false;
      if (/^[A-Z0-9 .\-]{1,10}$/.test(s)) return false;
      if (strictPdf && /^[A-Z0-9 .\-]{1,16}$/.test(s)) return false;
      return true;
    })
    .join("\n")
    .trim();
}

