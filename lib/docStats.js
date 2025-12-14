// lib/docStats.js (REPLACEMENT - stable output for UI)
// No new modules. Uses JSZip + xml2js + pdf-lib (already in project)

import JSZip from "jszip";
import xml2js from "xml2js";
import { PDFDocument } from "pdf-lib";

const parseStringPromise = xml2js.parseStringPromise;

// ---------- utils ----------
const countOcc = (s, re) => (String(s || "").match(re) || []).length;

async function safeRead(zip, p) {
  try {
    const f = zip?.file(p);
    if (!f) return "";
    // ✅ JSZip expects "text" (not "string") for reliable reading
    return await f.async("text");
  } catch {
    return "";
  }
}

function toIntOrNull(v) {
  const n = parseInt(String(v ?? ""), 10);
  return Number.isFinite(n) ? n : null;
}

// DOCX pages: best effort from docProps/app.xml
async function getDocxPagesFromApp(zip) {
  const xml = await safeRead(zip, "docProps/app.xml");
  if (!xml) return null;

  // 1) robust regex fallback (fast + reliable)
  const m = /<Pages>\s*(\d+)\s*<\/Pages>/i.exec(xml);
  if (m) return toIntOrNull(m[1]);

  // 2) xml2js parsing fallback (in case structure differs)
  try {
    const parsed = await parseStringPromise(xml);
    const pages = parsed?.Properties?.Pages?.[0];
    return toIntOrNull(pages);
  } catch {
    return null;
  }
}

function emptyStats() {
  return {
    pages: null,
    slides: null,
    sheets: null,
    tables: null,
    paragraphs: null,
    approxWords: null,
  };
}

/**
 * Extract structural stats. Always returns stable shape:
 * { pages, slides, sheets, tables, paragraphs, approxWords }
 *
 * @param {Object} params
 * @param {"pdf"|"docx"|"pptx"|"xlsx"} params.ext
 * @param {Buffer} params.buffer
 * @param {JSZip} [params.zip]
 * @param {string} [params.fullText]
 */
export async function extractDocStats({ ext, buffer, zip, fullText }) {
  const out = emptyStats();
  const e = String(ext || "").toLowerCase();

  // ---------------- PDF ----------------
  if (e === "pdf") {
    try {
      const pdfDoc = await PDFDocument.load(buffer);
      out.pages = pdfDoc.getPageCount();
    } catch {
      out.pages = null;
    }

    // optional (if you pass fullText from analysis)
    if (fullText && String(fullText).trim()) {
      out.approxWords = String(fullText).trim().split(/\s+/).length;
    }

    // tables/paragraphs not reliable in PDF without layout analysis
    return out;
  }

  // For Office formats: ensure zip exists
  if (!zip) zip = await JSZip.loadAsync(buffer);

  // ---------------- DOCX ----------------
  if (e === "docx") {
    // pages (best effort, can be null)
    out.pages = await getDocxPagesFromApp(zip);

    // tables/paragraphs from main doc + headers/footers
    let tables = 0;
    let paragraphs = 0;

    const main = await safeRead(zip, "word/document.xml");
    tables += countOcc(main, /<w:tbl\b/g);
    paragraphs += countOcc(main, /<w:p\b/g);

    // include headers/footers (common source of missed tables)
    const hf = Object.keys(zip.files).filter((k) => /^word\/(header|footer)\d+\.xml$/i.test(k));
    for (const p of hf) {
      const xml = await safeRead(zip, p);
      if (!xml) continue;
      tables += countOcc(xml, /<w:tbl\b/g);
      paragraphs += countOcc(xml, /<w:p\b/g);
    }

    out.tables = tables;
    out.paragraphs = paragraphs;

    if (fullText && String(fullText).trim()) {
      out.approxWords = String(fullText).trim().split(/\s+/).length;
    }

    return out;
  }

  // ---------------- PPTX ----------------
  if (e === "pptx") {
    const slideFiles = Object.keys(zip.files)
      .filter((k) => /^ppt\/slides\/slide\d+\.xml$/i.test(k))
      .sort((a, b) => {
        const na = parseInt((a.match(/slide(\d+)\.xml/i) || [])[1] || "0", 10);
        const nb = parseInt((b.match(/slide(\d+)\.xml/i) || [])[1] || "0", 10);
        return na - nb;
      });

    out.slides = slideFiles.length;

    let tables = 0;

    // tables in slides
    for (const p of slideFiles) {
      const xml = await safeRead(zip, p);
      tables += countOcc(xml, /<a:tbl\b/g);
    }

    // tables in speaker notes (optional but useful)
    const notesFiles = Object.keys(zip.files)
      .filter((k) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(k));

    for (const p of notesFiles) {
      const xml = await safeRead(zip, p);
      tables += countOcc(xml, /<a:tbl\b/g);
    }

    out.tables = tables;

    if (fullText && String(fullText).trim()) {
      out.approxWords = String(fullText).trim().split(/\s+/).length;
    }

    return out;
  }

  // ---------------- XLSX ----------------
  if (e === "xlsx") {
    const wb = await safeRead(zip, "xl/workbook.xml");
    // sheet nodes (workbook sheets list)
    const sheetsCount = countOcc(wb, /<sheet\b/gi);
    out.sheets = sheetsCount || 0;

    // structured tables are stored in xl/tables/tableN.xml
    const tableFiles = Object.keys(zip.files).filter((k) =>
      /^xl\/tables\/table\d+\.xml$/i.test(k)
    );
    out.tables = tableFiles.length || 0;

    if (fullText && String(fullText).trim()) {
      out.approxWords = String(fullText).trim().split(/\s+/).length;
    }

    return out;
  }

  return out;
}
