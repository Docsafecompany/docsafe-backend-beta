import JSZip from "jszip";
import xml2js from "xml2js";
import { PDFDocument } from "pdf-lib";

const parseStringPromise = xml2js.parseStringPromise;

const countOcc = (s, re) => (String(s || "").match(re) || []).length;

async function safeRead(zip, path) {
  const f = zip?.file(path);
  if (!f) return "";
  return await f.async("string");
}

async function getDocxPagesFromApp(zip) {
  try {
    const xml = await safeRead(zip, "docProps/app.xml");
    if (!xml) return null;
    const parsed = await parseStringPromise(xml);
    const pages = parsed?.Properties?.Pages?.[0];
    const n = pages != null ? parseInt(pages, 10) : NaN;
    return Number.isFinite(n) ? n : null;
  } catch {
    return null;
  }
}

export async function extractDocStats({ ext, buffer, zip, fullText }) {
  if (ext === "pdf") {
    const pdfDoc = await PDFDocument.load(buffer);
    const pageCount = pdfDoc.getPageCount();
    const approxWords =
      fullText && fullText.trim()
        ? fullText.trim().split(/\s+/).length
        : null;

    return {
      pages: pageCount,
      approxWords,
      tables: null, // not reliable in PDF without layout/OCR
    };
  }

  if (!zip) zip = await JSZip.loadAsync(buffer);

  if (ext === "docx") {
    const docXml = await safeRead(zip, "word/document.xml");

    const pages = await getDocxPagesFromApp(zip); // best-effort
    const tables = countOcc(docXml, /<w:tbl\b/g);
    const paragraphs = countOcc(docXml, /<w:p\b/g);

    return {
      pages,                 // can be null if missing
      tables,
      paragraphs,
    };
  }

  if (ext === "pptx") {
    const slideFiles = Object.keys(zip.files)
      .filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))
      .sort();

    let tables = 0;
    for (const p of slideFiles) {
      const xml = await safeRead(zip, p);
      tables += countOcc(xml, /<a:tbl\b/g);
    }

    return {
      slides: slideFiles.length,
      tables,
    };
  }

  if (ext === "xlsx") {
    const wb = await safeRead(zip, "xl/workbook.xml");
    const sheets = countOcc(wb, /<sheet\b/g);

    const tableFiles = Object.keys(zip.files).filter((k) =>
      /^xl\/tables\/table\d+\.xml$/.test(k)
    );

    return {
      sheets,
      tables: tableFiles.length, // real Excel "structured tables"
    };
  }

  return {};
}
