// lib/textExtractors.js
import JSZip from "jszip";

export async function extractTextFromDocx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const files = [
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/header1.xml",
    "word/header2.xml",
    "word/header3.xml"
  ];
  let out = "";
  for (const f of files) {
    const file = zip.file(f);
    if (!file) continue;
    const xml = await file.async("string");
    out += xmlToPlain(xml) + "\n";
  }
  return out.trim();
}

export async function extractTextFromPptx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const slideFiles = Object.keys(zip.files).filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k)).sort();
  let out = "";
  for (const f of slideFiles) {
    const xml = await zip.file(f).async("string");
    out += xmlToPlain(xml) + "\n\n";
  }
  return out.trim();
}

export async function extractTextFromPdf(buffer, { strictPdf = false } = {}) {
  try {
    const pdfParse = (await import("pdf-parse")).default;
    const data = await pdfParse(buffer);
    const lines = String(data.text || "").replace(/\r\n/g, "\n").split("\n");
    let filtered = lines;
    if (strictPdf) {
      filtered = lines.filter(ln => {
        const s = ln.trim();
        if (s.length <= 2) return false;
        if (/^[A-Z0-9 .\-]{1,10}$/.test(s)) return false; // heuristique watermark/entête
        return true;
      });
    }
    return filtered.join("\n").trim();
  } catch {
    return "";
  }
}

function xmlToPlain(xml) {
  let text = String(xml || "")
    .replace(/<w:tab\/>/g, " ")
    .replace(/<w:br\/>/g, "\n")
    .replace(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g, (_, g1) => g1)
    .replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (_, g1) => g1)
    .replace(/<[^>]+>/g, " ");
  return text.replace(/[ \t]{2,}/g, " ").replace(/\n{3,}/g, "\n\n").trim();
}

