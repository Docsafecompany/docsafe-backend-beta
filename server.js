/* server.js (MAJ V2 = rephrased + report uniquement) */
const express = require("express");
const cors = require("cors");
const multer = require("multer");
const fs = require("fs");
const path = require("path");
const os = require("os");
const archiver = require("archiver");
const mammoth = require("mammoth");
const JSZip = require("jszip");
const pdfParse = require("pdf-parse");

const { createDocxFromText } = require("./lib/docxWriter");

const app = express();
app.use(express.json({ limit: "25mb" }));
app.use(express.urlencoded({ extended: true, limit: "25mb" }));

app.use(
  cors({
    origin: "*",
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Accept"],
  })
);

const upload = multer({ storage: multer.memoryStorage() });

function normalizeText(input) {
  if (!input) return "";
  let text = input.replace(/[\u200B-\u200D\uFEFF\u2060\u00AD]/g, "");
  text = text.replace(/\r\n/g, "\n");
  text = text.replace(/[ \t]{2,}/g, " ");
  for (const p of [",", ";", ":", "!", "\\?", "@"]) {
    const re = new RegExp(`(${p})\\s*\\1+`, "g");
    text = text.replace(re, "$1");
  }
  text = text.replace(/\s+([,;:!?])/g, "$1");
  text = text.replace(/([,;:!?])([^\s\d])/g, "$1 $2");
  text = text.replace(/\n{3,}/g, "\n\n");
  return text.trim();
}

function aiSpellGrammarPass(input) {
  if (!input) return "";
  let t = input;
  t = t.replace(/\bi\s+am\b/gi, "I am");
  t = t.replace(/\bim\b/gi, "I'm");
  t = t.replace(/\bteh\b/gi, "the");
  t = t.replace(/\brecieve\b/gi, "receive");
  t = t.replace(/\boccured\b/gi, "occurred");
  t = t.replace(/\bca va\b/gi, "ça va");
  t = t.replace(/(^|\.\s+|\?\s+|!\s+)([a-zà-öø-ÿ])/g, (m, p1, p2) => p1 + p2.toUpperCase());
  return t;
}

function aiRephrase(input) {
  if (!input) return "";
  let t = input;
  t = t.replace(/\btrès\b/gi, "vraiment");
  t = t.replace(/\bimportant\b/gi, "clé");
  t = t.replace(/\bc'est\b/gi, "il est");
  t = t.replace(/\bdonc\b/gi, "ainsi");
  t = t.replace(/\bpar conséquent\b/gi, "en conséquence");
  t = t.replace(/([a-z])\.\s/gi, "$1. ");
  return t;
}

async function extractTextFromFile(fileBuffer, originalName, strictPdf = false) {
  const ext = path.extname(originalName || "").toLowerCase();

  if (ext === ".docx") {
    const { value } = await mammoth.extractRawText({ buffer: fileBuffer });
    return value || "";
  }
  if (ext === ".pdf") {
    const data = await pdfParse(fileBuffer);
    let txt = data.text || "";
    if (strictPdf) {
      txt = txt
        .replace(/[\u200B-\u200D\uFEFF\u2060\u00AD]/g, "")
        .replace(/[ \t]{2,}/g, " ")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
    }
    return txt;
  }
  if (ext === ".pptx") {
    const zip = await JSZip.loadAsync(fileBuffer);
    const slideFiles = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/i.test(f));
    slideFiles.sort((a, b) => parseInt(a.match(/slide(\d+)\.xml/i)[1], 10) - parseInt(b.match(/slide(\d+)\.xml/i)[1], 10));
    let all = [];
    for (const f of slideFiles) {
      const xml = await zip.files[f].async("string");
      const texts = Array.from(xml.matchAll(/<a:t[^>]*>([\s\S]*?)<\/a:t>/gi)).map((m) => m[1]);
      if (texts.length) all.push(texts.join(" "));
    }
    return all.join("\n\n");
  }
  return fileBuffer.toString("utf8");
}

function buildReportHtml({ originalName, lengthBefore, lengthAfter, notes }) {
  return `<!doctype html>
<html lang="fr"><head><meta charset="utf-8"/>
<title>DocSafe — Rapport</title>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<style>body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;padding:24px;line-height:1.6}
h1{font-size:20px;margin:0 0 16px}.kv{margin:8px 0}code{background:#f5f5f5;padding:2px 6px;border-radius:4px}.muted{color:#666}</style>
</head><body>
  <h1>DocSafe — Rapport de nettoyage</h1>
  <div class="kv"><strong>Fichier :</strong> ${originalName || "(sans nom)"}</div>
  <div class="kv"><strong>Caractères (avant) :</strong> ${lengthBefore}</div>
  <div class="kv"><strong>Caractères (après) :</strong> ${lengthAfter}</div>
  <div class="kv"><strong>Gain :</strong> ${Math.max(0, lengthBefore - lengthAfter)}</div>
  <hr/>
  <h2>Détails</h2>
  <ul class="muted">
    <li>Normalisation espaces/ponctuation (doublons supprimés).</li>
    <li>Correction orthographe/grammaire basique (IA locale).</li>
    <li>${notes?.strictPdf ? "Mode strictPdf actif (PDF nettoyé plus agressivement)." : "Mode strictPdf inactif."}</li>
  </ul>
</body></html>`;
}

function tmpPath(name) {
  return path.join(os.tmpdir(), `${Date.now()}_${name}`);
}

async function zipFiles(files) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    const archive = archiver("zip", { zlib: { level: 9 } });
    archive.on("warning", (e) => console.warn("zip warning:", e));
    archive.on("error", (e) => reject(e));
    archive.on("data", (d) => chunks.push(d));
    archive.on("end", () => resolve(Buffer.concat(chunks)));
    archive.append("", { name: ".keep" });
    for (const f of files) archive.file(f.path, { name: f.name });
    archive.finalize();
  });
}

/* ----------- Health/Echo ----------- */
app.get("/health", (_req, res) => res.json({ ok: true, message: "Backend is running ✅" }));
app.get("/_env_ok", (_req, res) => res.json({ ok: true, NODE_ENV: process.env.NODE_ENV || "development" }));
app.get("/_ai_echo", (req, res) => res.json({ ok: true, echo: aiSpellGrammarPass(String(req.query.q || "Hello from AI")) }));
app.get("/_ai_rephrase_echo", (req, res) => res.json({ ok: true, rephrased: aiRephrase(String(req.query.q || "Hello from AI")) }));

/* ----------- V1: cleaned + report ----------- */
app.post("/clean", upload.any(), async (req, res) => {
  try {
    if (!req.files?.length) return res.status(400).json({ ok: false, error: "No file uploaded" });
    const strictPdf = String(req.body.strictPdf || "false") === "true";
    const f = req.files[0];
    const originalName = f.originalname || "input";
    const rawText = await extractTextFromFile(f.buffer, originalName, strictPdf);
    const lengthBefore = rawText.length;

    const corrected = aiSpellGrammarPass(rawText);
    const normalized = normalizeText(corrected);
    const lengthAfter = normalized.length;

    const cleanedPath = tmpPath("cleaned.docx");
    await createDocxFromText(normalized, cleanedPath);

    const reportPath = tmpPath("report.html");
    fs.writeFileSync(
      reportPath,
      buildReportHtml({ originalName, lengthBefore, lengthAfter, notes: { strictPdf } }),
      "utf8"
    );

    const zipBuf = await zipFiles([
      { path: cleanedPath, name: "cleaned.docx" },
      { path: reportPath, name: "report.html" },
    ]);

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", 'attachment; filename="docsafe_v1_result.zip"');
    res.end(zipBuf);
  } catch (error) {
    console.error("CLEAN ERROR", error);
    res.status(500).json({ ok: false, error: "CLEAN_ERROR", message: String(error) });
  }
});

/* ----------- V2: rephrased + report (UNIQUEMENT) ----------- */
app.post("/clean-v2", upload.any(), async (req, res) => {
  try {
    if (!req.files?.length) return res.status(400).json({ ok: false, error: "No file uploaded" });
    const strictPdf = String(req.body.strictPdf || "false") === "true";
    const f = req.files[0];
    const originalName = f.originalname || "input";
    const rawText = await extractTextFromFile(f.buffer, originalName, strictPdf);
    const lengthBefore = rawText.length;

    // pipeline V1 (mais on ne garde PAS le fichier cleaned)
    const corrected = aiSpellGrammarPass(rawText);
    const normalized = normalizeText(corrected);
    const lengthAfter = normalized.length;

    // rephrase puis docx
    const rephrased = normalizeText(aiRephrase(normalized));
    const rephrasedPath = tmpPath("rephrased.docx");
    await createDocxFromText(rephrased, rephrasedPath);

    // report
    const reportPath = tmpPath("report.html");
    fs.writeFileSync(
      reportPath,
      buildReportHtml({ originalName, lengthBefore, lengthAfter, notes: { strictPdf } }),
      "utf8"
    );

    // ZIP: uniquement rephrased + report
    const zipBuf = await zipFiles([
      { path: rephrasedPath, name: "rephrased.docx" },
      { path: reportPath, name: "report.html" },
    ]);

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", 'attachment; filename="docsafe_v2_result.zip"');
    res.end(zipBuf);
  } catch (error) {
    console.error("CLEAN V2 ERROR", error);
    res.status(500).json({ ok: false, error: "CLEAN_V2_ERROR", message: String(error) });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend listening on ${PORT}`));

