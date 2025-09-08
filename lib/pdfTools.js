// lib/pdfTools.js
import { PDFDocument } from "pdf-lib";
import pdfParse from "pdf-parse";

export async function stripPdfMetadata(buffer) {
  const pdfDoc = await PDFDocument.load(buffer, { updateMetadata: true });
  // Efface les champs courants
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
  const data = await pdfParse(buffer);
  return String(data.text || "");
}

// Heuristique : élimine filigranes/entêtes courts en MAJ, lignes trop courtes, etc.
export function filterExtractedLines(text, { strictPdf = false } = {}) {
  const lines = String(text || "").replace(/\r\n/g, "\n").split("\n");
  return lines
    .filter((ln) => {
      const s = ln.trim();
      if (!s) return false;
      if (s.length <= 1) return false;
      if (/^[A-Z0-9 .\-]{1,10}$/.test(s)) return false;       // bruit évident
      if (strictPdf && /^[A-Z0-9 .\-]{1,16}$/.test(s)) return false; // plus agressif
      return true;
    })
    .join("\n")
    .trim();
}
