// lib/applySpellingFixes.js
import JSZip from "jszip";
import xml2js from "xml2js";
import { PDFDocument, StandardFonts } from "pdf-lib";
import { extractPdfText } from "./pdfTools.js";

const parseXml = async (xml) => xml2js.parseStringPromise(xml, { explicitArray: false });
const buildXml = (obj) => new xml2js.Builder({ headless: true }).buildObject(obj);

function pickSelected(issues, selectedIdsOrIdx) {
  const set = new Set(selectedIdsOrIdx || []);
  // support ids or indexes
  return (issues || []).filter((it, idx) => set.has(it.id) || set.has(idx));
}

function replaceWithAnchor(original, issue) {
  // safest apply: locate using contextBefore + error + contextAfter
  const before = issue.contextBefore || "";
  const after = issue.contextAfter || "";
  const err = issue.error || "";
  const corr = issue.correction || "";
  if (!err || !corr) return { text: original, applied: false };

  // Try anchored match first
  if (before || after) {
    const anchor = before + err + after;
    const idx = original.indexOf(anchor);
    if (idx !== -1) {
      const start = idx + before.length;
      const end = start + err.length;
      const out = original.slice(0, start) + corr + original.slice(end);
      return { text: out, applied: true };
    }
  }

  // fallback: replace first occurrence of exact error
  const idx2 = original.indexOf(err);
  if (idx2 !== -1) {
    const out = original.slice(0, idx2) + corr + original.slice(idx2 + err.length);
    return { text: out, applied: true };
  }

  return { text: original, applied: false };
}

function deepReplaceRuns(obj, issue) {
  let count = 0;

  const walk = (node) => {
    if (!node) return;
    if (typeof node === "string") return;

    for (const k of Object.keys(node)) {
      const v = node[k];

      // DOCX text
      if (k === "w:t") {
        if (typeof v === "string") {
          const r = replaceWithAnchor(v, issue);
          if (r.applied) { node[k] = r.text; count++; }
        } else if (Array.isArray(v)) {
          v.forEach((vv, i) => {
            const val = typeof vv === "string" ? vv : (vv?._ || "");
            const r = replaceWithAnchor(val, issue);
            if (r.applied) {
              if (typeof vv === "string") v[i] = r.text;
              else vv._ = r.text;
              count++;
            }
          });
        } else if (v?._) {
          const r = replaceWithAnchor(v._, issue);
          if (r.applied) { v._ = r.text; count++; }
        }
      }

      // PPTX text
      if (k === "a:t") {
        if (typeof v === "string") {
          const r = replaceWithAnchor(v, issue);
          if (r.applied) { node[k] = r.text; count++; }
        } else if (Array.isArray(v)) {
          v.forEach((vv, i) => {
            const val = typeof vv === "string" ? vv : (vv?._ || "");
            const r = replaceWithAnchor(val, issue);
            if (r.applied) {
              if (typeof vv === "string") v[i] = r.text;
              else vv._ = r.text;
              count++;
            }
          });
        } else if (v?._) {
          const r = replaceWithAnchor(v._, issue);
          if (r.applied) { v._ = r.text; count++; }
        }
      }

      if (Array.isArray(v)) v.forEach(walk);
      else if (typeof v === "object") walk(v);
    }
  };

  walk(obj);
  return count;
}

export async function applyOfficeFixes(buffer, ext, issues, selectedIdsOrIdx) {
  const selected = pickSelected(issues, selectedIdsOrIdx);

  if (ext === "docx") {
    const zip = await JSZip.loadAsync(buffer);
    const path = "word/document.xml";
    const xml = await zip.file(path)?.async("string");
    if (!xml) return { out: buffer, applied: [], skipped: selected };

    const json = await parseXml(xml);

    const applied = [];
    const skipped = [];

    for (const issue of selected) {
      const n = deepReplaceRuns(json, issue);
      if (n > 0) applied.push({ id: issue.id, error: issue.error, correction: issue.correction, count: n });
      else skipped.push(issue);
    }

    zip.file(path, buildXml(json));
    const out = await zip.generateAsync({ type: "nodebuffer" });
    return { out, applied, skipped };
  }

  if (ext === "pptx") {
    const zip = await JSZip.loadAsync(buffer);
    const slideFiles = Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p));

    const applied = [];
    const skipped = new Set(selected.map(s => s.id));

    for (const sf of slideFiles) {
      const xml = await zip.file(sf)?.async("string");
      if (!xml) continue;
      const json = await parseXml(xml);

      for (const issue of selected) {
        const n = deepReplaceRuns(json, issue);
        if (n > 0) {
          applied.push({ id: issue.id, error: issue.error, correction: issue.correction, count: n, part: sf });
          skipped.delete(issue.id);
        }
      }

      zip.file(sf, buildXml(json));
    }

    const out = await zip.generateAsync({ type: "nodebuffer" });
    return { out, applied, skipped: selected.filter(s => skipped.has(s.id)) };
  }

  if (ext === "xlsx") {
    const zip = await JSZip.loadAsync(buffer);

    // sharedStrings covers most visible strings
    const path = "xl/sharedStrings.xml";
    const xml = await zip.file(path)?.async("string");
    if (!xml) return { out: buffer, applied: [], skipped: selected };

    const json = await parseXml(xml);

    const applied = [];
    const skipped = [];

    for (const issue of selected) {
      const n = deepReplaceRuns(json, issue);
      if (n > 0) applied.push({ id: issue.id, error: issue.error, correction: issue.correction, count: n });
      else skipped.push(issue);
    }

    zip.file(path, buildXml(json));
    const out = await zip.generateAsync({ type: "nodebuffer" });
    return { out, applied, skipped };
  }

  throw new Error(`applyOfficeFixes: unsupported ext ${ext}`);
}

// PDF: rebuild clean (simple layout)
export async function applyPdfFixesRebuild(buffer, issues, selectedIdsOrIdx) {
  const selected = pickSelected(issues, selectedIdsOrIdx);

  let text = "";
  try {
    const raw = await extractPdfText(buffer);
    text = String(raw || "");
  } catch {
    text = "";
  }

  // apply all selected globally (anchored per line isn't possible reliably in PDF)
  let corrected = text;
  for (const issue of selected) {
    if (!issue.error || !issue.correction) continue;
    corrected = corrected.split(issue.error).join(issue.correction);
  }

  const pdfDoc = await PDFDocument.create();
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);
  const page = pdfDoc.addPage([595, 842]); // A4
  const fontSize = 11;
  const margin = 50;
  const maxWidth = 595 - margin * 2;
  const lineHeight = fontSize * 1.4;

  const words = corrected.replace(/\s+/g, " ").trim().split(" ").filter(Boolean);

  let x = margin;
  let y = 842 - margin;
  let line = "";

  const drawLine = (l) => {
    if (!l) return;
    page.drawText(l, { x: margin, y, size: fontSize, font });
    y -= lineHeight;
  };

  for (const w of words) {
    const candidate = line ? `${line} ${w}` : w;
    const width = font.widthOfTextAtSize(candidate, fontSize);
    if (width <= maxWidth) {
      line = candidate;
    } else {
      drawLine(line);
      line = w;
      if (y < margin) break; // (simple, 1 page) – tu peux étendre multi-pages si tu veux
    }
  }
  drawLine(line);

  const out = await pdfDoc.save();
  return { out: Buffer.from(out), applied: selected, skipped: [] };
}
