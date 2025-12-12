// lib/textExtractors.js
import JSZip from "jszip";

export async function extractTextFromDocx(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  const targets = [
    "word/document.xml",
    // headers & footers (tous)
    ...Object.keys(zip.files).filter((k) => /word\/header\d+\.xml$/i.test(k)),
    ...Object.keys(zip.files).filter((k) => /word\/footer\d+\.xml$/i.test(k)),
    // footnotes / endnotes
    "word/footnotes.xml",
    "word/endnotes.xml",
  ];

  let out = "";
  const seen = new Set();

  for (const p of targets) {
    if (seen.has(p)) continue;
    seen.add(p);

    const file = zip.file(p);
    if (!file) continue;

    const xml = await file.async("string");
    const extracted = xmlToPlain(xml, { mode: "docx" });

    if (extracted) out += (out ? "\n" : "") + extracted;

    // BONUS: textbox content (souvent déjà inclus, mais on renforce)
    const tbxMatches = xml.matchAll(/<w:txbxContent\b[^>]*>[\s\S]*?<\/w:txbxContent>/gi);
    for (const m of tbxMatches) {
      const tbxText = xmlToPlain(m[0], { mode: "docx" });
      if (tbxText) out += "\n" + tbxText;
    }
  }

  return normalizeExtracted(out);
}

export async function extractTextFromPptx(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  // slides + notesSlides
  const files = Object.keys(zip.files)
    .filter(
      (k) =>
        ((k.startsWith("ppt/slides/slide") && k.endsWith(".xml")) ||
          (k.startsWith("ppt/notesSlides/") && k.endsWith(".xml")))
    )
    .sort((a, b) => {
      // tri intelligent: slide1, slide2...
      const na = parseInt((a.match(/(\d+)\.xml$/) || [])[1] || "0", 10);
      const nb = parseInt((b.match(/(\d+)\.xml$/) || [])[1] || "0", 10);
      return na - nb;
    });

  let out = "";
  for (const p of files) {
    const f = zip.file(p);
    if (!f) continue;
    const xml = await f.async("string");
    const extracted = xmlToPlain(xml, { mode: "pptx" });
    if (!extracted) continue;
    out += (out ? "\n\n" : "") + extracted;
  }

  return normalizeExtracted(out);
}

export async function extractTextFromPdf(buffer, { strictPdf = false } = {}) {
  try {
    const pdfParse = (await import("pdf-parse")).default;
    const data = await pdfParse(buffer);

    const lines = String(data.text || "").replace(/\r\n/g, "\n").split("\n");
    let filtered = lines;

    if (strictPdf) {
      filtered = lines.filter((ln) => {
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

// ============================================================
// Core XML -> Plain text (NO fake spaces)
// ============================================================

function xmlToPlain(xml, { mode = "generic" } = {}) {
  let x = String(xml || "");
  if (!x.trim()) return "";

  // --- Preserve structure with real separators (avoid injecting fake spaces) ---
  // DOCX
  if (mode === "docx") {
    x = x
      .replace(/<w:tab\b[^\/>]*\/>/gi, "\t")
      .replace(/<w:br\b[^\/>]*\/>/gi, "\n")
      .replace(/<\/w:p>/gi, "\n")
      .replace(/<\/w:tr>/gi, "\n")
      .replace(/<\/w:tc>/gi, "\t");
  }

  // PPTX
  if (mode === "pptx") {
    x = x
      .replace(/<a:br\b[^\/>]*\/>/gi, "\n")
      .replace(/<\/a:p>/gi, "\n");
  }

  // Extract explicit text nodes first (keeps true spacing better)
  // DOCX: w:t / w:instrText / w:delText
  // PPTX: a:t
  // NOTE: we keep them in place by not collapsing everything into spaces.
  x = x
    .replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/gi, (_, g1) => decodeXmlEntities(g1))
    .replace(/<w:instrText[^>]*>([\s\S]*?)<\/w:instrText>/gi, (_, g1) => decodeXmlEntities(g1))
    .replace(/<w:delText[^>]*>([\s\S]*?)<\/w:delText>/gi, (_, g1) => decodeXmlEntities(g1))
    .replace(/<a:t[^>]*>([\s\S]*?)<\/a:t>/gi, (_, g1) => decodeXmlEntities(g1));

  // Remove remaining tags WITHOUT forcing a space
  // (forcing spaces is exactly what created "The Imp", "act of", etc.)
  x = x.replace(/<[^>]+>/g, "");

  // Decode entities again (in case they exist outside handled nodes)
  x = decodeXmlEntities(x);

  // Normalize whitespace while preserving \n and \t
  x = x
    .replace(/[ \f\r\v]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/\t{2,}/g, "\t")
    .trim();

  return x;
}

function decodeXmlEntities(s) {
  return String(s || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&#34;/g, '"');
}

function normalizeExtracted(s) {
  return String(s || "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}
