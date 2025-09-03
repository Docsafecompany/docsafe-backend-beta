/**
 * DocSafe Backend — Beta V2 (clean PDF/DOCX/PPTX + LT report)
 */
import express from "express";
import cors from "cors";
import multer from "multer";
import fs from "fs";
import fsp from "fs/promises";
import os from "os";
import axios from "axios";
import JSZip from "jszip";
import { PDFDocument, PDFName } from "pdf-lib";
import zlib from "zlib";

let pdfParse = null; // lazy import

const app = express();

/* ---------------- CORS ---------------- */
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

/* --------------- Upload --------------- */
const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, os.tmpdir()),
    filename: (req, file, cb) => cb(null, Date.now() + "-" + file.originalname.replace(/\s+/g, "_"))
  }),
  limits: { fileSize: 20 * 1024 * 1024 }
});

/* -------------- Helpers --------------- */
const LT_URL = process.env.LT_API_URL || "https://api.languagetool.org/v2/check";
const LT_KEY = process.env.LT_API_KEY || null;

function cleanTextBasic(str) {
  if (!str) return str;
  let s = String(str);

  // Remove zero-width & soft characters
  s = s.replace(/\u200B|\u2060|\uFEFF/g, "");

  // Normalize spaces & line breaks
  s = s.replace(/[ \t]+/g, " ")
       .replace(/ *\n */g, "\n");

  // Normalize punctuation spacing (keep "..." intact)
  s = s.replace(/ ?([,;:!?]) ?/g, "$1 ")
       .replace(/ \./g, ".");

  // Collapse duplicates
  s = s
    .replace(/,{2,}/g, ",")
    .replace(/@{2,}/g, "@")
    .replace(/;{2,}/g, ";")
    .replace(/:{2,}/g, ":")
    .replace(/!{2,}/g, "!")
    .replace(/\?{2,}/g, "?");

  // Final trim
  s = s.replace(/[ ]{2,}/g, " ").trim();
  return s;
}

/** Collapse duplicates that fall ACROSS XML text nodes.
 *  Example: ",</w:t><w:t>,"  => ",</w:t><w:t>"
 */
function collapseAcrossNodes(xml, tag) {
  // duplicate punctuations across nodes: , ; : ! ? @
  const reDupe = new RegExp(`([,;:!?@])<\\/${tag}>\\s*<${tag}[^>]*>\\1`, "g");
  xml = xml.replace(reDupe, (_m, p1) => `${p1}</${tag}><${tag}>`);

  // remove weird spaces around node boundaries (space before/after)
  const reSpace1 = new RegExp(`\\s+<\\/${tag}>\\s*<${tag}[^>]*>`, "g");
  xml = xml.replace(reSpace1, `</${tag}><${tag}>`);
  return xml;
}

/* -------------- DOCX ------------------ */
async function processDOCXBasic(buf) {
  const zip = await JSZip.loadAsync(buf);

  // strip core/custom props
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

  // clean text nodes + cross-node collapse
  const docXml = "word/document.xml";
  if (zip.file(docXml)) {
    let xml = await zip.file(docXml).async("string");
    // Per-node cleaning
    xml = xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (m, inner) => m.replace(inner, cleanTextBasic(inner)));
    // Cross-node punctuation duplicates
    xml = collapseAcrossNodes(xml, "w:t");
    zip.file(docXml, xml);
  }
  return await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

async function extractTextFromDOCX(buf) {
  const zip = await JSZip.loadAsync(buf);
  const p = "word/document.xml";
  if (!zip.file(p)) return "";
  const xml = await zip.file(p).async("string");
  let t = "";
  xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (_m, inner) => { t += inner + " "; return _m; });
  return cleanTextBasic(t);
}

/* -------------- PPTX ------------------ */
async function processPPTXBasic(buf) {
  const zip = await JSZip.loadAsync(buf);

  // strip core/custom props
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

  // Clean slide text nodes
  const slideFiles = Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/i.test(p));
  for (const p of slideFiles) {
    let xml = await zip.file(p).async("string");
    // Per-node cleaning
    xml = xml.replace(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g, (m, inner) => m.replace(inner, cleanTextBasic(inner)));
    // Cross-node collapse
    xml = collapseAcrossNodes(xml, "a:t");
    zip.file(p, xml);
  }
  return await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

async function extractTextFromPPTX(buf) {
  const zip = await JSZip.loadAsync(buf);
  let out = "";
  const slideFiles = Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/i.test(p));
  for (const p of slideFiles) {
    const xml = await zip.file(p).async("string");
    xml.replace(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g, (_m, inner) => { out += inner + " "; return _m; });
  }
  return cleanTextBasic(out);
}

/* --------------- PDF ------------------ */
async function processPDFBasic(buf, { strict = false } = {}) {
  const pdf = await PDFDocument.load(buf, { updateMetadata: true });

  // Clear standard metadata
  pdf.setTitle(""); pdf.setAuthor(""); pdf.setSubject("");
  pdf.setKeywords([]); pdf.setProducer(""); pdf.setCreator("");
  const epoch = new Date(0);
  pdf.setCreationDate(epoch); pdf.setModificationDate(epoch);

  // Remove XMP metadata stream if any
  try { pdf.catalog?.dict?.delete(PDFName.of("Metadata")); } catch {}

  // Remove potentially sensitive catalog entries
  try {
    const dict = pdf.catalog?.dict;
    if (dict) {
      dict.delete(PDFName.of("Names"));
      dict.delete(PDFName.of("OpenAction"));
      dict.delete(PDFName.of("AcroForm"));
      dict.delete(PDFName.of("AA"));
    }
  } catch {}

  let outBuf = Buffer.from(await pdf.save({ useObjectStreams: false }));

  if (strict) outBuf = tryStripHiddenText(outBuf);
  return outBuf;
}

// Heuristic removal of invisible/white text in content streams
function tryStripHiddenText(pdfBuf) {
  const src = pdfBuf.toString("binary");
  const parts = [];
  let idx = 0;

  while (true) {
    const sIdx = src.indexOf("stream", idx);
    if (sIdx === -1) { parts.push([idx, src.length, null]); break; }
    const eLine = src.indexOf("\n", sIdx);
    const eCR = src.indexOf("\r\n", sIdx);
    const eol = (eCR !== -1 && (eLine === -1 || eCR < eLine)) ? eCR + 2 : eLine + 1;
    const endIdx = src.indexOf("endstream", eol);
    if (endIdx === -1) { parts.push([idx, src.length, null]); break; }
    parts.push([idx, sIdx, null]);
    parts.push([eol, endIdx, "stream"]);
    idx = endIdx;
  }

  let rebuilt = "";
  for (const [a, b, kind] of parts) {
    const chunk = src.slice(a, b);
    if (kind !== "stream") { rebuilt += chunk; continue; }

    let data = chunk;
    let inflated = null;
    try { inflated = zlib.inflateSync(Buffer.from(chunk, "binary")); data = inflated.toString("binary"); } catch {}

    // remove text after "3 Tr" (invisible) or with white fill "1 1 1 rg"
    let cleaned = data.replace(/\r\n/g, "\n");
    cleaned = cleaned.replace(/(?:^|\n)([^]*?)\b3\s+Tr[^\n]*\n/g, line =>
      line.replace(/\((?:\\.|[^\\)])*\)\s*T[Jj]/g, "")
    );
    cleaned = cleaned.replace(/(?:^|\n)([^]*?)\b1(?:\.0+)?\s+1(?:\.0+)?\s+1(?:\.0+)?\s+rg[^\n]*\n/g, line =>
      line.replace(/\((?:\\.|[^\\)])*\)\s*T[Jj]/g, "")
    );
    cleaned = cleaned.replace(/[ ]{2,}/g, " ");

    let out;
    try {
      if (inflated && cleaned !== data) {
        const deflated = zlib.deflateSync(Buffer.from(cleaned, "binary"));
        out = deflated.toString("binary");
      } else {
        out = inflated ? data : cleaned;
      }
    } catch { out = cleaned; }

    rebuilt += out;
  }
  return Buffer.from(rebuilt, "binary");
}

async function extractTextFromPDF(buf) {
  try {
    if (!pdfParse) {
      const mod = await import("pdf-parse");
      pdfParse = mod.default || mod;
    }
    const data = await pdfParse(buf);
    return cleanTextBasic(data.text || "");
  } catch { return ""; }
}

/* --------- LanguageTool ---------- */
async function runLanguageTool(fullText, lang = "auto") {
  const chunks = [];
  for (let i = 0; i < fullText.length; i += 20000) chunks.push(fullText.slice(i, i + 20000));
  let matches = [];
  for (const chunk of chunks) {
    if (!chunk.trim()) continue;
    const form = new URLSearchParams();
    form.set("text", chunk);
    form.set("language", lang);
    if (LT_KEY) form.set("apiKey", LT_KEY);
    const r = await axios.post(LT_URL, form.toString(), {
      headers: { "Content-Type": "application/x-www-form-urlencoded" }, timeout: 30000
    });
    if (r?.data?.matches?.length) matches = matches.concat(r.data.matches);
  }
  return matches;
}

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
  </div>${report.infoBanner || ""}
  <table class="table"><thead><tr><th>#</th><th>Rule</th><th>Message</th><th>Suggestions</th><th>Context</th></tr></thead>
  <tbody>${rows || `<tr><td colspan="5" style="padding:10px">No suggestions.</td></tr>`}</tbody>
  </table></div></div></body></html>`;
}

/* --------------- Routes --------------- */
app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/clean", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Missing file" });
  const p = req.file.path, name = req.file.originalname.toLowerCase();
  const strict = String(req.query?.strict || "") === "1";
  try {
    let outBuf, outName, mime;
    if (name.endsWith(".pdf")) {
      outBuf = await processPDFBasic(await fsp.readFile(p), { strict });
      outName = req.file.originalname.replace(/\.pdf$/i, "") + "_cleaned.pdf";
      mime = "application/pdf";
    } else if (name.endsWith(".docx")) {
      outBuf = await processDOCXBasic(await fsp.readFile(p));
      outName = req.file.originalname.replace(/\.docx$/i, "") + "_cleaned.docx";
      mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    } else if (name.endsWith(".pptx")) {
      outBuf = await processPPTXBasic(await fsp.readFile(p));
      outName = req.file.originalname.replace(/\.pptx$/i, "") + "_cleaned.pptx";
      mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    } else {
      return res.status(400).json({ error: "Only PDF, DOCX or PPTX supported" });
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
  const strict = String(req.query?.strict || "") === "1";
  try {
    let cleanedBuf, cleanedName, text = "";
    if (name.endsWith(".pdf")) {
      const buf = await fsp.readFile(p);
      cleanedBuf = await processPDFBasic(buf, { strict });
      cleanedName = req.file.originalname.replace(/\.pdf$/i, "") + "_cleaned.pdf";
      text = await extractTextFromPDF(cleanedBuf);
    } else if (name.endsWith(".docx")) {
      const buf = await fsp.readFile(p);
      cleanedBuf = await processDOCXBasic(buf);
      cleanedName = req.file.originalname.replace(/\.docx$/i, "") + "_cleaned.docx";
      text = await extractTextFromDOCX(cleanedBuf);
    } else if (name.endsWith(".pptx")) {
      const buf = await fsp.readFile(p);
      cleanedBuf = await processPPTXBasic(buf);
      cleanedName = req.file.originalname.replace(/\.pptx$/i, "") + "_cleaned.pptx";
      text = await extractTextFromPPTX(cleanedBuf);
    } else {
      return res.status(400).json({ error: "Only PDF, DOCX or PPTX supported" });
    }

    let matches = [];
    let ltNote = null;
    try {
      matches = await runLanguageTool(text || "", lang);
    } catch (e) {
      ltNote = "LanguageTool unavailable (network/quota). Report generated without suggestions.";
    }

    const reportJSON = buildReportJSON({ fileName: cleanedName, language: lang, matches, textLength: (text || "").length });
    const reportHTML = buildReportHTML({ ...reportJSON, infoBanner: ltNote ? `<div style="margin:12px 0;padding:10px 12px;border-radius:10px;background:#3b82f6;color:#fff">${ltNote}</div>` : "" });

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

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log("DocSafe backend running on port", PORT));
