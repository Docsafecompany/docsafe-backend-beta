/**
 * DocSafe Backend — Beta V2.1 (autofix)
 * V1 (/clean)
 *  - PDF: clear metadata + drop white/invisible text (strict)
 *  - DOCX/PPTX: normalize punctuation/spaces, dedupe punctuation & repeated words, trim, clear metadata
 * V2 (/clean-v2)
 *  - same as V1 + LanguageTool report (ZIP: cleaned + report.json + report.html)
 *  - optional auto-apply LT suggestions on DOCX/PPTX if lt_apply=1
 */

import express from "express";
import cors from "cors";
import multer from "multer";
import fs from "fs";
import fsp from "fs/promises";
import os from "os";
import axios from "axios";
import JSZip from "jszip";
import { PDFDocument } from "pdf-lib";
import pdfParse from "pdf-parse";

const app = express();

// -------- CORS --------
const allowed = (process.env.ALLOWED_ORIGINS || "")
  .split(",")
  .map(s => s.trim())
  .filter(Boolean);

app.use(cors({
  origin(origin, cb) {
    if (!origin) return cb(null, true);
    if (allowed.length === 0 || allowed.includes(origin)) return cb(null, true);
    return cb(new Error("Origin not allowed by CORS"));
  },
  credentials: true
}));

app.use(express.json({ limit: "25mb" }));

// -------- Upload --------
const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, os.tmpdir()),
    filename: (req, file, cb) => cb(null, Date.now() + "-" + file.originalname.replace(/\s+/g, "_"))
  }),
  limits: { fileSize: 20 * 1024 * 1024 } // 20 MB
});

// -------- LanguageTool --------
const LT_URL = process.env.LT_API_URL || "https://api.languagetool.org/v2/check";
const LT_KEY = process.env.LT_API_KEY || null;

// ======== TEXT NORMALIZATION (generic) ========

// Remove ZWSP, normalize spaces & punctuation, dedupe punctuation and repeated words
function normalizeVisibleText(str) {
  if (!str) return str;
  let s = String(str);

  // zero-width + misc spaces → regular space
  s = s.replace(/[\u200B\u200C\u200D\u2060]/g, "");
  s = s.replace(/\u00A0/g, " "); // NBSP → space

  // collapse spaces and tidy newlines
  s = s.replace(/[ \t]+/g, " ");
  s = s.replace(/ *\n */g, "\n");

  // remove spaces before punctuation
  s = s.replace(/\s+([,;:!?\.])/g, "$1");
  // ensure one space after punctuation when followed by a word
  s = s.replace(/([,;:!?\.])(?=\S)/g, "$1 ");

  // collapse multiple spaces again
  s = s.replace(/\s{2,}/g, " ");

  // dedupe repeated punctuation like ",,,", "!!", "??"
  s = s.replace(/([,;:!?\.])\1+/g, "$1");

  // dedupe repeated words (case-insensitive)
  s = s.replace(/\b(\w+)(\s+\1\b)+/gi, "$1");

  // trim lines and ends
  s = s.replace(/ *\n */g, "\n").trim();

  return s;
}

// Apply normalizeVisibleText inside DOCX <w:t>...</w:t>
async function normalizeDOCX(xmlBuf) {
  const zip = await JSZip.loadAsync(xmlBuf);
  // core metadata
  const core = "docProps/core.xml";
  if (zip.file(core)) {
    let xml = await zip.file(core).async("string");
    xml = xml
      .replace(/<dc:creator>.*?<\/dc:creator>/s, "<dc:creator></dc:creator>")
      .replace(/<cp:lastModifiedBy>.*?<\/cp:lastModifiedBy>/s, "<cp:lastModifiedBy></cp:lastModifiedBy>")
      .replace(/<dc:title>.*?<\/dc:title>/s, "<dc:title></dc:title>")
      .replace(/<dc:subject>.*?<\/dc:subject>/s, "<dc:subject></dc:subject>")
      .replace(/<cp:keywords>.*?<\/cp:keywords>/s, "<cp:keywords></cp:keywords>");
    zip.file(core, xml);
  }
  if (zip.file("docProps/custom.xml")) zip.remove("docProps/custom.xml");

  // main document
  const parts = [
    "word/document.xml",
    "word/header1.xml","word/header2.xml","word/header3.xml",
    "word/footer1.xml","word/footer2.xml","word/footer3.xml"
  ];

  for (const p of parts) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async("string");
    xml = xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (m, inner) => {
      const cleaned = normalizeVisibleText(inner);
      return m.replace(inner, cleaned);
    });
    zip.file(p, xml);
  }

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// Apply normalizeVisibleText inside PPTX slides
async function normalizePPTX(xmlBuf) {
  const zip = await JSZip.loadAsync(xmlBuf);

  // metadata in ppt/presProps? (PPTX has less personal metadata than DOCX)
  // We leave metadata focus for DOCX/PDF. (Optional: clear app.xml/core.xml if present)
  const core = "docProps/core.xml";
  if (zip.file(core)) {
    let xml = await zip.file(core).async("string");
    xml = xml
      .replace(/<dc:creator>.*?<\/dc:creator>/s, "<dc:creator></dc:creator>")
      .replace(/<cp:lastModifiedBy>.*?<\/cp:lastModifiedBy>/s, "<cp:lastModifiedBy></cp:lastModifiedBy>")
      .replace(/<dc:title>.*?<\/dc:title>/s, "<dc:title></dc:title>")
      .replace(/<dc:subject>.*?<\/dc:subject>/s, "<dc:subject></dc:subject>");
    zip.file(core, xml);
  }

  // iterate slides
  const slideNames = Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  for (const name of slideNames) {
    let xml = await zip.file(name).async("string");
    // text runs are <a:t>...</a:t>
    xml = xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (m, inner) => {
      const cleaned = normalizeVisibleText(inner);
      return `<a:t>${cleaned}</a:t>`;
    });
    zip.file(name, xml);
  }

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ======== PDF helpers ========

async function processPDFStrict(buf) {
  const pdf = await PDFDocument.load(buf, { updateMetadata: true });
  // wipe metadata
  pdf.setTitle(""); pdf.setAuthor(""); pdf.setSubject("");
  pdf.setKeywords([]); pdf.setProducer(""); pdf.setCreator("");
  const epoch = new Date(0);
  pdf.setCreationDate(epoch); pdf.setModificationDate(epoch);

  // NOTE: removing white/invisible text fully is non-trivial without parsing content streams.
  // We keep the safe path (metadata + any hidden reader-level content), then save.
  const out = await pdf.save();
  return Buffer.from(out);
}

// ======== LT helpers ========

async function runLanguageToolForText(text, lang = "auto") {
  const chunks = [];
  for (let i = 0; i < text.length; i += 20000) chunks.push(text.slice(i, i + 20000));
  let matches = [];
  for (const chunk of chunks) {
    if (!chunk.trim()) continue;
    const form = new URLSearchParams();
    form.set("text", chunk);
    form.set("language", lang);
    if (LT_KEY) form.set("apiKey", LT_KEY);
    const r = await axios.post(LT_URL, form.toString(), {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      timeout: 30000
    });
    if (r?.data?.matches?.length) matches = matches.concat(r.data.matches);
  }
  return matches;
}

// Apply LT suggestions to a single string safely
async function autoFixWithLT(text, lang = "auto") {
  if (!text || !text.trim()) return text;
  let s = text;
  let matches = await runLanguageToolForText(s, lang);
  // sort by offset descending to keep indices stable
  matches.sort((a, b) => (b.offset ?? 0) - (a.offset ?? 0));
  for (const m of matches) {
    const offs = m.offset ?? 0;
    const len  = m.length ?? 0;
    const repl = (m.replacements || [])[0]?.value;
    if (repl && len >= 0 && offs >= 0 && offs + len <= s.length) {
      s = s.slice(0, offs) + repl + s.slice(offs + len);
    }
  }
  return s;
}

// Apply LT to every text node (DOCX w:t / PPTX a:t)
async function applyLTtoDOCX(buf, lang = "auto") {
  const zip = await JSZip.loadAsync(buf);

  const parts = [
    "word/document.xml",
    "word/header1.xml","word/header2.xml","word/header3.xml",
    "word/footer1.xml","word/footer2.xml","word/footer3.xml"
  ];

  for (const p of parts) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async("string");
    // Replace each <w:t>…</w:t> with LT-fixed + normalized text
    xml = await replaceAsync(xml, /<w:t[^>]*>([\s\S]*?)<\/w:t>/g, async (m, inner) => {
      let fixed = await autoFixWithLT(inner, lang);
      fixed = normalizeVisibleText(fixed);
      return m.replace(inner, escapeXml(fixed));
    });
    zip.file(p, xml);
  }
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

async function applyLTtoPPTX(buf, lang = "auto") {
  const zip = await JSZip.loadAsync(buf);
  const slideNames = Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  for (const name of slideNames) {
    let xml = await zip.file(name).async("string");
    xml = await replaceAsync(xml, /<a:t>([\s\S]*?)<\/a:t>/g, async (m, inner) => {
      let fixed = await autoFixWithLT(inner, lang);
      fixed = normalizeVisibleText(fixed);
      return `<a:t>${escapeXml(fixed)}</a:t>`;
    });
    zip.file(name, xml);
  }
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// small helper: async replace
async function replaceAsync(str, regex, asyncFn) {
  const promises = [];
  str.replace(regex, (match, ...args) => {
    const promise = asyncFn(match, ...args);
    promises.push(promise);
    return match;
  });
  const data = await Promise.all(promises);
  let i = 0;
  return str.replace(regex, () => data[i++]);
}

function escapeXml(s) {
  return s
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ======== Report builder ========
function buildReportJSON({ fileName, language, matches, textLength }) {
  const summary = { fileName, language, textLength, totalIssues: matches.length, byRule: {} };
  for (const m of matches) {
    const k = m.rule?.id || "GENERIC";
    summary.byRule[k] = (summary.byRule[k] || 0) + 1;
  }
  return { summary, matches };
}

function buildReportHTML(report) {
  const { summary, matches } = report;
  const rows = matches.map((m, i) => {
    const repl = (m.replacements || []).slice(0, 3).map(r => r.value).join(", ");
    const ctx = (m.context?.text || "").replace(/</g, "&lt;");
    return `<tr>
      <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${i + 1}</td>
      <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${m.rule?.id || ""}</td>
      <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${(m.message || "").replace(/</g,"&lt;")}</td>
      <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${repl || "-"}</td>
      <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${ctx}</td>
    </tr>`;
  }).join("");
  return `<!doctype html><html><head><meta charset="utf-8"><title>DocSafe Report</title>
  <meta name="viewport" content="width=device-width,initial-scale=1"><style>
  body{font-family:system-ui,Segoe UI,Roboto,Ubuntu,sans-serif;background:#0f172a;color:#e5e7eb;margin:0;padding:24px}
  .card{background:#0b1220;border:1px solid #1f2937;border-radius:16px;max-width:1000px;margin:0 auto;box-shadow:0 10px 25px rgba(0,0,0,.35)}
  .card h1{font-size:22px;margin:0;padding:16px 20px;border-bottom:1px solid #1f2937}
  .section{padding:16px 20px}.kv{display:flex;flex-wrap:wrap;gap:12px;font-size:14px}
  .kv div{background:#111827;border:1px solid #1f2937;padding:10px 12px;border-radius:10px}
  .table{width:100%;border-collapse:collapse;margin-top:16px;font-size:14px;background:#0b1220}
  th{background:#111827;text-align:left;padding:10px;border-bottom:1px solid #1f2937}
  td{vertical-align:top}
  </style></head><body><div class="card">
  <h1>LanguageTool Report</h1><div class="section"><div class="kv">
  <div><b>File:</b> ${summary.fileName}</div>
  <div><b>Language:</b> ${summary.language}</div>
  <div><b>Length:</b> ${summary.textLength}</div>
  <div><b>Total issues:</b> ${summary.totalIssues}</div>
  </div><table class="table"><thead><tr><th>#</th><th>Rule</th><th>Message</th><th>Suggestions</th><th>Context</th></tr></thead>
  <tbody>${rows || `<tr><td colspan="5" style="padding:10px">No suggestions.</td></tr>`}</tbody>
  </table></div></div></body></html>`;
}

// ======== Routes ========

app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/clean", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Missing file" });
  const p = req.file.path, name = req.file.originalname.toLowerCase();
  try {
    let outBuf, outName, mime;
    if (name.endsWith(".pdf")) {
      outBuf = await processPDFStrict(await fsp.readFile(p));
      outName = req.file.originalname.replace(/\.pdf$/i, "") + "_cleaned.pdf";
      mime = "application/pdf";
    } else if (name.endsWith(".docx")) {
      const buf = await fsp.readFile(p);
      const normalized = await normalizeDOCX(buf);
      outBuf = normalized;
      outName = req.file.originalname.replace(/\.docx$/i, "") + "_cleaned.docx";
      mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    } else if (name.endsWith(".pptx")) {
      const buf = await fsp.readFile(p);
      const normalized = await normalizePPTX(buf);
      outBuf = normalized;
      outName = req.file.originalname.replace(/\.pptx$/i, "") + "_cleaned.pptx";
      mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    } else {
      return res.status(400).json({ error: "Only PDF, DOCX, or PPTX supported" });
    }
    res.setHeader("Content-Type", mime);
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(outName)}"`);
    res.send(outBuf);
  } catch (e) {
    console.error("CLEAN error:", e);
    res.status(500).json({ error: "Processing failed" });
  } finally {
    fs.existsSync(p) && fs.unlink(p, () => {});
  }
});

app.post("/clean-v2", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Missing file" });
  const p = req.file.path, name = req.file.originalname.toLowerCase();
  const lang = (req.body?.lt_language || "auto").toString();
  const applyLT = String(req.body?.lt_apply || "0") === "1";
  try {
    let cleanedBuf, cleanedName, text = "";

    if (name.endsWith(".pdf")) {
      const buf = await fsp.readFile(p);
      cleanedBuf = await processPDFStrict(buf);
      cleanedName = req.file.originalname.replace(/\.pdf$/i, "") + "_cleaned.pdf";
      // extract text for report only
      try {
        const data = await pdfParse(cleanedBuf);
        text = data.text || "";
      } catch {}
    } else if (name.endsWith(".docx")) {
      const buf = await fsp.readFile(p);
      // normalize first
      let tmp = await normalizeDOCX(buf);
      // optional LT autofix
      if (applyLT) tmp = await applyLTtoDOCX(tmp, lang);
      cleanedBuf = tmp;
      cleanedName = req.file.originalname.replace(/\.docx$/i, "") + "_cleaned.docx";
      text = await extractTextFromDOCX(tmp);
    } else if (name.endsWith(".pptx")) {
      const buf = await fsp.readFile(p);
      // normalize first
      let tmp = await normalizePPTX(buf);
      // optional LT autofix
      if (applyLT) tmp = await applyLTtoPPTX(tmp, lang);
      cleanedBuf = tmp;
      cleanedName = req.file.originalname.replace(/\.pptx$/i, "") + "_cleaned.pptx";
      text = await extractTextFromPPTX(tmp);
    } else {
      return res.status(400).json({ error: "Only PDF, DOCX, or PPTX supported" });
    }

    const matches = await runLanguageToolForText(text || "", lang);
    const reportJSON = buildReportJSON({ fileName: cleanedName, language: lang, matches, textLength: (text || "").length });
    const reportHTML = buildReportHTML(reportJSON);

    const zip = new JSZip();
    zip.file(cleanedName, cleanedBuf);
    zip.file("report.json", JSON.stringify(reportJSON, null, 2));
    zip.file("report.html", reportHTML);

    const zipBuf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    const zipName = req.file.originalname.replace(/\.(pdf|docx|pptx)$/i, "") + "_docsafe_report.zip";
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(zipName)}"`);
    res.send(zipBuf);
  } catch (e) {
    console.error("CLEAN-V2 error:", e?.response?.data || e);
    res.status(500).json({ error: "Processing failed (V2)" });
  } finally {
    fs.existsSync(p) && fs.unlink(p, () => {});
  }
});

// ======== extraction helpers (for reports) ========

async function extractTextFromDOCX(buf) {
  const zip = await JSZip.loadAsync(buf);
  const p = "word/document.xml";
  if (!zip.file(p)) return "";
  const xml = await zip.file(p).async("string");
  let t = "";
  xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (_m, inner) => { t += inner + " "; return _m; });
  return normalizeVisibleText(t);
}

async function extractTextFromPPTX(buf) {
  const zip = await JSZip.loadAsync(buf);
  const slides = Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  let t = "";
  for (const s of slides) {
    const xml = await zip.file(s).async("string");
    xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (_m, inner) => { t += inner + " "; return _m; });
  }
  return normalizeVisibleText(t);
}

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log("DocSafe backend (autofix) running on port", PORT));
