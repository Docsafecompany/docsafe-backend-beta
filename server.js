/**
 * DocSafe Backend — Beta V2 (PDF/DOCX/PPTX + LanguageTool)
 * Endpoints:
 *  - POST /clean       -> retourne le fichier nettoyé (PDF/DOCX/PPTX)
 *  - POST /clean-v2    -> nettoyé + rapport LanguageTool en ZIP (cleaned + report.json + report.html)
 *  - GET  /health      -> OK
 *
 * Notes:
 *  - Option strict (PDF) : ajouter ?strict=1 pour tenter de retirer le texte blanc/invisible.
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

let pdfParse = null; // lazy import pour éviter certains bugs de charge

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
  limits: { fileSize: 20 * 1024 * 1024 } // 20 Mo
});

// -------- Helpers --------
const LT_URL = process.env.LT_API_URL || "https://api.languagetool.org/v2/check";
const LT_KEY = process.env.LT_API_KEY || null;

function cleanTextBasic(str) {
  if (!str) return str;
  return String(str)
    .replace(/\u200B/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/ *\n */g, "\n")
    .replace(/ ?([,;:!?]) ?/g, "$1 ")
    .replace(/ \./g, ".")
    .replace(/[ ]{2,}/g, " ")
    .replace(/,,+/g, ",")
    .replace(/;;+/g, ";")
    .trim();
}

// ----- DOCX -----
async function processDOCXBasic(buf) {
  const zip = await JSZip.loadAsync(buf);
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

  const docXml = "word/document.xml";
  if (zip.file(docXml)) {
    let xml = await zip.file(docXml).async("string");
    xml = xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (m, inner) => m.replace(inner, cleanTextBasic(inner)));
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

// ----- PPTX -----
async function processPPTXBasic(buf) {
  const zip = await JSZip.loadAsync(buf);

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

  // Nettoyage texte dans les slides
  const slideFiles = Object.keys(zip.files).filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p));
  for (const p of slideFiles) {
    let xml = await zip.file(p).async("string");
    xml = xml.replace(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g, (m, inner) => m.replace(inner, cleanTextBasic(inner)));
    zip.file(p, xml);
  }

  return await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

async function extractTextFromPPTX(buf) {
  const zip = await JSZip.loadAsync(buf);
  let out = "";
  const slideFiles = Object.keys(zip.files).filter((p) => /^ppt\/slides\/slide\d+\.xml$/i.test(p));
  for (const p of slideFiles) {
    const xml = await zip.file(p).async("string");
    xml.replace(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g, (_m, inner) => { out += inner + " "; return _m; });
  }
  return cleanTextBasic(out);
}

// ----- PDF -----
async function processPDFBasic(buf, { strict = false } = {}) {
  const pdf = await PDFDocument.load(buf, { updateMetadata: true });
  // Scrub info dictionary
  pdf.setTitle(""); pdf.setAuthor(""); pdf.setSubject("");
  pdf.setKeywords([]); pdf.setProducer(""); pdf.setCreator("");
  const epoch = new Date(0);
  pdf.setCreationDate(epoch); pdf.setModificationDate(epoch);

  // Tentative de suppression de la référence /Metadata dans le Catalog si présente
  try {
    const catalog = pdf.catalog;
    if (catalog?.dict) {
      catalog.dict.delete(PDFName.of("Metadata"));
    }
  } catch {}

  // Tentative de suppression d'AcroForm/JavaScript
  try {
    const dict = pdf.catalog?.dict;
    if (dict) {
      dict.delete(PDFName.of("Names"));
      dict.delete(PDFName.of("OpenAction"));
      dict.delete(PDFName.of("AcroForm"));
      dict.delete(PDFName.of("AA"));
    }
  } catch {}

  let out = await pdf.save({ useObjectStreams: false });
  let outBuf = Buffer.from(out);

  if (strict) {
    // Heuristique: parcourir les objets stream (FlateDecode ou texte brut) et supprimer
    // les instructions Tj/TJ lorsqu'elles sont précédées de "3 Tr" (invisible) ou "1 1 1 rg" (blanc).
    outBuf = tryStripHiddenText(outBuf);
  }
  return outBuf;
}

function tryStripHiddenText(pdfBuf) {
  const src = pdfBuf.toString("binary");

  // Découpe grossière par objets stream
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

    parts.push([idx, sIdx, null]); // header part
    parts.push([eol, endIdx, "stream"]); // actual stream data
    idx = endIdx;
  }

  let rebuilt = "";
  for (let i = 0; i < parts.length; i++) {
    const [a, b, kind] = parts[i];
    const chunk = src.slice(a, b);
    if (kind !== "stream") {
      rebuilt += chunk;
      continue;
    }

    // Tente de décompresser (Flate)
    let data = chunk;
    try {
      const buf = Buffer.from(chunk, "binary");
      const inflated = zlib.inflateSync(buf);
      data = inflated.toString("binary");
    } catch {
      // peut déjà être non compressé, on garde data tel quel
    }

    // Suppression d'opérations texte invisibles/blanches (heuristique best-effort)
    const cleaned = stripOpsInStream(data);

    // Recompresse si on avait réussi à inflater
    let out;
    try {
      if (cleaned !== data) {
        const deflated = zlib.deflateSync(Buffer.from(cleaned, "binary"));
        out = deflated.toString("binary");
      } else {
        out = chunk;
      }
    } catch {
      out = cleaned;
    }
    rebuilt += out;
  }

  return Buffer.from(rebuilt, "binary");
}

// Supprime les Tj/TJ lorsque précédées de "3 Tr" (texte invisible) ou "1 1 1 rg" (blanc)
function stripOpsInStream(data) {
  // Simplifie les EOL pour la détection
  let s = data.replace(/\r\n/g, "\n");

  // 1) Invisible text: "3 Tr" => supprime les Tj/TJ sur la même ligne
  s = s.replace(/(?:^|\n)([^]*?)\b3\s+Tr[^\n]*\n/g, (line) => {
    // retire Tj/TJ dans cette ligne
    return line.replace(/\((?:\\.|[^\\)])*\)\s*T[Jj]/g, "");
  });

  // 2) White text: "1 1 1 rg" => supprime Tj/TJ sur la même ligne
  s = s.replace(/(?:^|\n)([^]*?)\b1(?:\.0+)?\s+1(?:\.0+)?\s+1(?:\.0+)?\s+rg[^\n]*\n/g, (line) => {
    return line.replace(/\((?:\\.|[^\\)])*\)\s*T[Jj]/g, "");
  });

  // Nettoyage: espaces multiples devenus inutiles
  s = s.replace(/[ ]{2,}/g, " ");
  return s;
}

async function extractTextFromPDF(buf) {
  try {
    if (!pdfParse) {
      // lazy import pour éviter bug au lancement dans certains environnements
      const mod = await import("pdf-parse");
      pdfParse = mod.default || mod;
    }
    const data = await pdfParse(buf);
    return cleanTextBasic(data.text || "");
  } catch {
    return "";
  }
}

// ----- LanguageTool -----
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
  <h1>Rapport LanguageTool</h1><div class="section"><div class="kv">
  <div><b>Fichier:</b> ${summary.fileName}</div>
  <div><b>Langue:</b> ${summary.language}</div>
  <div><b>Longueur:</b> ${summary.textLength}</div>
  <div><b>Total issues:</b> ${summary.totalIssues}</div>
  </div><table class="table"><thead><tr><th>#</th><th>Règle</th><th>Message</th><th>Suggestions</th><th>Contexte</th></tr></thead>
  <tbody>${rows || `<tr><td colspan="5" style="padding:10px">Aucune suggestion.</td></tr>`}</tbody>
  </table></div></div></body></html>`;
}

// -------- Routes --------
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

    const matches = await runLanguageTool(text || "", lang);
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

const PORT = process.env.PORT || 8080;
app.listen(PORT, () => console.log("DocSafe backend running on port", PORT));

