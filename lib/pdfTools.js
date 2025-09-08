import { PDFDocument } from "pdf-lib";

async function getPdfParse() {
  try {
    const mod = await import("pdf-parse/lib/pdf-parse.js");
    return (mod && (mod.default || mod)) || null;
  } catch (e1) {
    try {
      const mod2 = await import("pdf-parse");
      return (mod2 && (mod2.default || mod2)) || null;
    } catch (e2) {
      console.warn("pdf-parse unavailable; text extraction skipped.", e1?.message || e1, e2?.message || e2);
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
    console.warn("pdf-parse failed; returning empty text.", e?.message || e);
    return "";
  }
}

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
