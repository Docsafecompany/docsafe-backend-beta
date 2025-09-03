/**
 * DocSafe Backend — Beta V2 (strict cleaner)
 * Endpoints:
 *  - POST /clean     -> retourne le fichier nettoyé (PDF/DOCX/PPTX)
 *  - POST /clean-v2  -> nettoyé + rapport LanguageTool en ZIP (cleaned + report.json + report.html)
 *  - GET  /health    -> OK
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
app.use(
  cors({
    origin(origin, cb) {
      if (!origin) return cb(null, true);
      if (allowed.length === 0 || allowed.includes(origin)) return cb(null, true);
      return cb(new Error("Origin not allowed by CORS"));
    },
    credentials: true,
  })
);

app.use(express.json({ limit: "25mb" }));

// -------- Upload --------
const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, os.tmpdir()),
    filename: (req, file, cb) =>
      cb(null, Date.now() + "-" + file.originalname.replace(/\s+/g, "_")),
  }),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20 Mo
});

// -------- Helpers --------
const LT_URL =
  process.env.LT_API_URL || "https://api.languagetool.org/v2/check";
const LT_KEY = process.env.LT_API_KEY || null;

/**
 * Clean text aggressively:
 * - Collapse duplicate punctuation/symbols (", ; : ! ? . @ # % & * € $")
 * - Normalize spaces and spacing around punctuation
 * - Remove repeated consecutive words (case-insensitive)
 * - Trim lines and collapse excessive blank lines
 */
function cleanText(str) {
  if (!str) return str;
  let s = String(str).replace(/\r\n?/g, "\n");

  // (1) Collapse duplicate punctuation & symbols
  s = s.replace(/([,;:!?])(?:\s*\1)+/g, "$1"); // , ; : ! ?
  // Keep "." strict: any run of dots -> single "."
  s = s.replace(/\.{2,}/g, ".");
  s = s.replace(/([@#%&*€$])\1+/g, "$1"); // @ # % & * € $

  // (2) Normalize spaces and spacing around punctuation
  s = s.replace(/[ \t]+/g, " "); // collapse horizontal spaces
  s = s.replace(/\s+([,;:!?\.])/g, "$1"); // no space before punctuation
  s = s.replace(/([,;:!?\.])(?![\s)\]\}»”'])/g, "$1 "); // ensure one space after
  s = s.replace(/[ ]{2,}/g, " "); // cleanup doubles

  // (3) Parentheses / quotes tightening
  s = s.replace(/\(\s+/g, "(").replace(/\s+\)/g, ")");
  s = s.replace(/\s+(['”»])/g, "$1").replace(/([«“'])\s+/g, "$1");

  // (4) Remove consecutive duplicate words (>= 2 chars)
  s = s.replace(/\b([A-Za-zÀ-ÖØ-öø-ÿ]{2,})\b(\s+\1\b)+/gi, "$1");

  // (5) Strip leftovers like " ." , " ;" etc.
  s = s
    .replace(/ \./g, ".")
    .replace(/ ,/g, ",")
    .replace(/ ;/g, ";")
    .replace(/ :/g, ":")
    .replace(/ !/g, "!")
    .replace(/ \?/g, "?");

  // (6) Trim lines & collapse excessive blank lines
  s = s.replace(/^[ \t]+|[ \t]+$/gm, "");
  s = s.replace(/\n{3,}/g, "\n\n");

  return s.trim();
}

/** ---------- DOCX ---------- */
async function processDOCXBasic(buf) {
  const zip = await JSZip.loadAsync(buf);

  // Clean core metadata
  const core = "docProps/core.xml";
  if (zip.file(core)) {
    let xml = await zip.file(core).async("string");
    xml = xml
      .replace(/<dc:creator>.*?<\/dc:creator>/s, "<dc:creator></dc:creator>")
      .replace(
        /<cp:lastModifiedBy>.*?<\/cp:lastModifiedBy>/s,
        "<cp:lastModifiedBy></cp:lastModifiedBy>"
      )
      .replace(/<dc:title>.*?<\/dc:title>/s, "<dc:title></dc:title>")
      .replace(/<dc:subject>.*?<\/dc:subject>/s, "<dc:subject></dc:subject>")
      .replace(/<cp:keywords>.*?<\/cp:keywords>/s, "<cp:keywords></cp:keywords>");
    zip.file(core, xml);
  }
  if (zip.file("docProps/custom.xml")) zip.remove("docProps/custom.xml");

  // Clean document text: replace contents of <w:t>
  const docXml = "word/document.xml";
  if (zip.file(docXml)) {
    let xml = await zip.file(docXml).async("string");
    xml = xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (m, inner) => {
      const cleaned = cleanText(inner);
      return m.replace(inner, cleaned);
    });
    zip.file(docXml, xml);
  }
  return await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
}

async function extractTextFromDOCX(buf) {
  const zip = await JSZip.loadAsync(buf);
  const p = "word/document.xml";
  if (!zip.file(p)) return "";
  const xml = await zip.file(p).async("string");
  let t = "";
  xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (_m, inner) => {
    t += inner + " ";
    return _m;
  });
  return cleanText(t);
}

/** ---------- PPTX ---------- */
async function processPPTXBasic(buf) {
  const zip = await JSZip.loadAsync(buf);

  // Clean core metadata
  const core = "docProps/core.xml";
  if (zip.file(core)) {
    let xml = await zip.file(core).async("string");
    xml = xml
      .replace(/<dc:creator>.*?<\/dc:creator>/s, "<dc:creator></dc:creator>")
      .replace(
        /<cp:lastModifiedBy>.*?<\/cp:lastModifiedBy>/s,
        "<cp:lastModifiedBy></cp:lastModifiedBy>"
      )
      .replace(/<dc:title>.*?<\/dc:title>/s, "<dc:title></dc:title>");
    zip.file(core, xml);
  }
  if (zip.file("docProps/custom.xml")) zip.remove("docProps/custom.xml");

  // Clean slides text: replace contents of <a:t> on all slides
  const folder = zip.folder("ppt/slides");
  if (folder) {
    const files = Object.keys(folder.files).filter((n) => /slide\d+\.xml$/.test(n));
    for (const name of files) {
      let xml = await zip.file(`ppt/slides/${name}`).async("string");
      xml = xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (m, inner) => {
        const cleaned = cleanText(inner);
        return m.replace(inner, cleaned);
      });
      zip.file(`ppt/slides/${name}`, xml);
    }
  }

  return await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
  });
}

async function extractTextFromPPTX(buf) {
  const zip = await JSZip.loadAsync(buf);
  const folder = zip.folder("ppt/slides");
  if (!folder) return "";
  const files = Object.keys(folder.files).filter((n) => /slide\d+\.xml$/.test(n));
  let all = "";
  for (const name of files) {
    const xml = await zip.file(`ppt/slides/${name}`).async("string");
    xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (_m, inner) => {
      all += inner + " ";
      return _m;
    });
  }
  return cleanText(all);
}

/** ---------- PDF ---------- */
async function processPDFBasic(buf) {
  // Clear metadata (content streams are not rewritten)
  const pdf = await PDFDocument.load(buf, { updateMetadata: true });
  pdf.setTitle("");
  pdf.setAuthor("");
  pdf.setSubject("");
  pdf.setKeywords([]);
  pdf.setProducer("");
  pdf.setCreator("");
  const epoch = new Date(0);
  pdf.setCreationDate(epoch);
  pdf.setModificationDate(epoch);
  const out = await pdf.save();
  return Buffer.from(out);
}

async function extractTextFromPDF(buf) {
  try {
    const data = await pdfParse(buf);
    return cleanText(data.text || "");
  } catch {
    return "";
  }
}

/** ---------- LanguageTool & Report ---------- */
async function runLanguageTool(fullText, lang = "auto") {
  const chunks = [];
  for (let i = 0; i < fullText.length; i += 20000)
    chunks.push(fullText.slice(i, i + 20000));
  let matches = [];
  for (const chunk of chunks) {
    if (!chunk.trim()) continue;
    const form = new URLSearchParams();
    form.set("text", chunk);
    form.set("language", lang);
    if (LT_KEY) form.set("apiKey", LT_KEY);
    const r = await axios.post(LT_URL, form.toString(), {
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      timeout: 30000,
    });
    if (r?.data?.matches?.length) matches = matches.concat(r.data.matches);
  }
  return matches;
}

function buildReportJSON({ fileName, language, matches, textLength }) {
  const summary = {
    fileName,
    language,
    textLength,
    totalIssues: matches.length,
    byRule: {},
  };
  for (const m of matches) {
    const k = m.rule?.id || "GENERIC";
    summary.byRule[k] = (summary.byRule[k] || 0) + 1;
  }
  return { summary, matches };
}

function buildReportHTML(report) {
  const { summary, matches } = report;
  const rows = matches
    .map((m, i) => {
      const repl = (m.replacements || [])
        .slice(0, 3)
        .map((r) => r.value)
        .join(", ");
      const ctx = (m.context?.text || "").replace(/</g, "&lt;");
      return `<tr>
      <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${i + 1}</td>
      <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${m.rule?.id || ""}</td>
      <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${(m.message || "").replace(/</g, "&lt;")}</td>
      <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${repl || "-"}</td>
      <td style="padding:8px;border-bottom:1px solid #e5e7eb;">${ctx}</td>
    </tr>`;
    })
    .join("");
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

// -------- Routes --------
app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/clean", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Missing file" });
  const p = req.file.path,
    name = req.file.originalname.toLowerCase();
  try {
    let outBuf,
      outName,
      mime;
    if (name.endsWith(".pdf")) {
      outBuf = await processPDFBasic(await fsp.readFile(p));
      outName = req.file.originalname.replace(/\.pdf$/i, "") + "_cleaned.pdf";
      mime = "application/pdf";
    } else if (name.endsWith(".docx")) {
      outBuf = await processDOCXBasic(await fsp.readFile(p));
      outName =
        req.file.originalname.replace(/\.docx$/i, "") + "_cleaned.docx";
      mime =
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    } else if (name.endsWith(".pptx")) {
      outBuf = await processPPTXBasic(await fsp.readFile(p));
      outName =
        req.file.originalname.replace(/\.pptx$/i, "") + "_cleaned.pptx";
      mime =
        "application/vnd.openxmlformats-officedocument.presentationml.presentation";
    } else {
      return res.status(400).json({ error: "Only PDF, DOCX or PPTX supported" });
    }
    res.setHeader("Content-Type", mime);
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${encodeURIComponent(outName)}"`
    );
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
  const p = req.file.path,
    name = req.file.originalname.toLowerCase();
  const lang = (req.body?.lt_language || "auto").toString();
  try {
    let cleanedBuf,
      cleanedName,
      text = "";
    if (name.endsWith(".pdf")) {
      const buf = await fsp.readFile(p);
      cleanedBuf = await processPDFBasic(buf);
      cleanedName =
        req.file.originalname.replace(/\.pdf$/i, "") + "_cleaned.pdf";
      text = await extractTextFromPDF(cleanedBuf);
    } else if (name.endsWith(".docx")) {
      const buf = await fsp.readFile(p);
      cleanedBuf = await processDOCXBasic(buf);
      cleanedName =
        req.file.originalname.replace(/\.docx$/i, "") + "_cleaned.docx";
      text = await extractTextFromDOCX(cleanedBuf);
    } else if (name.endsWith(".pptx")) {
      const buf = await fsp.readFile(p);
      cleanedBuf = await processPPTXBasic(buf);
      cleanedName =
        req.file.originalname.replace(/\.pptx$/i, "") + "_cleaned.pptx";
      text = await extractTextFromPPTX(cleanedBuf);
    } else {
      return res.status(400).json({ error: "Only PDF, DOCX or PPTX supported" });
    }

    const matches = await runLanguageTool(text || "", lang);
    const reportJSON = buildReportJSON({
      fileName: cleanedName,
      language: lang,
      matches,
      textLength: (text || "").length,
    });
    const reportHTML = buildReportHTML(reportJSON);

    const zip = new JSZip();
    zip.file(cleanedName, cleanedBuf);
    zip.file("report.json", JSON.stringify(reportJSON, null, 2));
    zip.file("report.html", reportHTML);

    const zipBuf = await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
    });
    const zipName = req.file.originalname.replace(/\.(pdf|docx|pptx)$/i, "") + "_docsafe_report.zip";
    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${encodeURIComponent(zipName)}"`
    );
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

