// lib/pdfCleaner.js
import {
  PDFDocument,
  PDFName,
  PDFDict,
  PDFArray,
  PDFString,
  PDFHexString,
  StandardFonts,
} from "pdf-lib";

/**
 * Build a "text-only" PDF from extracted text
 */
async function buildTextOnlyPdf(text, title = "Qualion Text-Only") {
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);

  const margin = 48;
  const pageW = 595.28;
  const pageH = 841.89;
  const maxW = pageW - margin * 2;
  const lh = 14;
  const size = 11;

  let page = doc.addPage([pageW, pageH]);
  let y = pageH - margin;

  const words = (text || "")
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .filter(Boolean);

  const w = (s) => font.widthOfTextAtSize(s, size);

  let line = "";
  for (const word of words) {
    const test = line ? `${line} ${word}` : word;

    if (w(test) <= maxW) {
      line = test;
      continue;
    }

    if (y - lh < margin) {
      page = doc.addPage([pageW, pageH]);
      y = pageH - margin;
    }

    if (line) {
      page.drawText(line, { x: margin, y: y - lh, size, font });
      y -= lh;
      line = word;
    } else {
      page.drawText(word, { x: margin, y: y - lh, size, font });
      y -= lh;
      line = "";
    }
  }

  if (line) {
    if (y - lh < margin) {
      page = doc.addPage([pageW, pageH]);
      y = pageH - margin;
    }
    page.drawText(line, { x: margin, y: y - lh, size, font });
  }

  doc.setTitle(title);
  doc.setCreator("Qualion");
  doc.setProducer("Qualion");
  return Buffer.from(await doc.save());
}

// ---------- Safe helpers (pdf-lib can vary by version) ----------
const N = (s) => PDFName.of(s);

function asDict(maybe) {
  return maybe instanceof PDFDict ? maybe : null;
}
function asArray(maybe) {
  return maybe instanceof PDFArray ? maybe : null;
}
function safeLookup(dict, keyName) {
  try {
    const d = asDict(dict);
    if (!d) return null;
    return d.lookup(N(keyName));
  } catch {
    return null;
  }
}
function safeGet(dict, keyName) {
  try {
    const d = asDict(dict);
    if (!d) return null;
    return d.get(N(keyName));
  } catch {
    return null;
  }
}
function safeSet(dict, keyName, value) {
  try {
    const d = asDict(dict);
    if (!d) return false;
    d.set(N(keyName), value);
    return true;
  } catch {
    return false;
  }
}
function safeDelete(dict, keyName) {
  try {
    const d = asDict(dict);
    if (!d) return false;
    d.delete(N(keyName));
    return true;
  } catch {
    return false;
  }
}

function countArrayEntries(maybeArr) {
  const arr = asArray(maybeArr);
  if (!arr) return 0;
  try {
    return arr.size();
  } catch {
    // fallback if size() not available (rare)
    try {
      const v = arr.asArray?.();
      return Array.isArray(v) ? v.length : 0;
    } catch {
      return 0;
    }
  }
}

/**
 * pdfMode:
 *  - "sanitize" (default): wipe metadata, remove annotations, embedded files, and neutralize forms/js/actions
 *  - "text-only": rebuild a text-only PDF (removes all graphics)
 * extractTextFn(inputBuffer) optional for text-only
 */
export async function cleanPDF(inputBuffer, { pdfMode = "sanitize", extractTextFn } = {}) {
  const stats = {
    metadataCleared: false,
    annotsRemoved: 0,
    embeddedFilesRemoved: 0,
    acroFormRemoved: false,
    javascriptRemoved: false,
    openActionRemoved: false,
    additionalActionsRemoved: 0,
    textOnly: pdfMode === "text-only",
  };

  // ---------------- text-only ----------------
  if (pdfMode === "text-only") {
    const text = typeof extractTextFn === "function" ? await extractTextFn(inputBuffer) : "";
    const out = await buildTextOnlyPdf(text);
    return { outBuffer: out, stats: { ...stats, metadataCleared: true } };
  }

  // ---------------- sanitize ----------------
  let pdf;
  try {
    pdf = await PDFDocument.load(inputBuffer, { updateMetadata: false });
  } catch (e) {
    // if load fails, fail gracefully with clear error
    throw new Error(`PDF load failed: ${e?.message || e}`);
  }

  // ---- Count annotations before ----
  let annotsBefore = 0;
  try {
    for (const page of pdf.getPages()) {
      const a = page.node.lookup(N("Annots"));
      annotsBefore += countArrayEntries(a);
    }
  } catch (e) {
    console.warn("[PDF] Failed counting annotations:", e?.message || e);
  }

  // ---- Metadata wipe ----
  try {
    pdf.setTitle("");
    pdf.setAuthor("");
    pdf.setCreator("");
    pdf.setProducer("");
    pdf.setSubject("");
    pdf.setKeywords([]);
    stats.metadataCleared = true;
  } catch (e) {
    console.warn("[PDF] Failed clearing metadata:", e?.message || e);
  }

  // ---- Remove annotations ----
  try {
    for (const page of pdf.getPages()) {
      // set Annots to empty array
      page.node.set(N("Annots"), pdf.context.obj([]));
    }
    stats.annotsRemoved = annotsBefore;
  } catch (e) {
    console.warn("[PDF] Failed clearing annotations:", e?.message || e);
  }

  // ---- Neutralize OpenAction + AA (catalog-level) ----
  try {
    const catalog = pdf.catalog;
    const catDict = asDict(catalog.dict);
    if (catDict) {
      if (safeGet(catDict, "OpenAction")) {
        safeDelete(catDict, "OpenAction");
        stats.openActionRemoved = true;
      }
      // Additional Actions at catalog level
      if (safeGet(catDict, "AA")) {
        safeDelete(catDict, "AA");
        stats.additionalActionsRemoved += 1;
      }
    }
  } catch (e) {
    console.warn("[PDF] Failed removing OpenAction/AA:", e?.message || e);
  }

  // ---- Neutralize page-level AA + actions inside annotations already removed ----
  try {
    for (const page of pdf.getPages()) {
      const pDict = asDict(page.node);
      if (!pDict) continue;
      if (safeGet(pDict, "AA")) {
        safeDelete(pDict, "AA");
        stats.additionalActionsRemoved += 1;
      }
    }
  } catch (e) {
    console.warn("[PDF] Failed removing page AA:", e?.message || e);
  }

  // ---- Remove AcroForm (forms) ----
  try {
    const catalog = pdf.catalog;
    const catDict = asDict(catalog.dict);
    if (catDict && safeGet(catDict, "AcroForm")) {
      safeDelete(catDict, "AcroForm");
      stats.acroFormRemoved = true;
    }
  } catch (e) {
    console.warn("[PDF] Failed removing AcroForm:", e?.message || e);
  }

  // ---- Remove JavaScript name tree entries ----
  // JS can be stored under Catalog -> Names -> JavaScript -> Names [...]
  try {
    const catalog = pdf.catalog;
    const catDict = asDict(catalog.dict);
    const namesObj = catDict ? safeLookup(catDict, "Names") : null;
    const namesDict = asDict(namesObj);

    if (namesDict) {
      // EmbeddedFiles
      const efTree = safeLookup(namesDict, "EmbeddedFiles");
      const efDict = asDict(efTree);
      if (efDict) {
        const efNames = safeLookup(efDict, "Names");
        const pairs = countArrayEntries(efNames) / 2; // [name, spec] pairs
        if (pairs > 0) stats.embeddedFilesRemoved = pairs;

        // Reset embedded files name tree (safe minimal)
        safeSet(namesDict, "EmbeddedFiles", pdf.context.obj({ Names: [] }));
      }

      // JavaScript
      const jsTree = safeLookup(namesDict, "JavaScript");
      if (jsTree) {
        // best effort count
        const jsDict = asDict(jsTree);
        const jsNames = jsDict ? safeLookup(jsDict, "Names") : null;
        const jsPairs = countArrayEntries(jsNames) / 2;
        if (jsPairs > 0) stats.javascriptRemoved = true;

        // Remove the JavaScript entry entirely (cleanest)
        safeDelete(namesDict, "JavaScript");
        stats.javascriptRemoved = true;
      }
    }
  } catch (e) {
    console.warn("[PDF] Failed removing embedded files / JS:", e?.message || e);
  }

  // ---- Also remove document-level /OpenAction in trailer root if present (rare) ----
  // pdf-lib exposes catalog already, but some PDFs place actions oddly; this is best-effort.
  try {
    const catalog = pdf.catalog;
    const catDict = asDict(catalog.dict);
    if (catDict && safeGet(catDict, "AA")) {
      safeDelete(catDict, "AA");
      stats.additionalActionsRemoved += 1;
    }
  } catch {}

  // Save
  try {
    const out = await pdf.save({ useObjectStreams: false });
    return { outBuffer: Buffer.from(out), stats };
  } catch (e) {
    throw new Error(`PDF save failed: ${e?.message || e}`);
  }
}

export default { cleanPDF };
