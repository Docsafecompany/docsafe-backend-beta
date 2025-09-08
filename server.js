// server.js
import express from "express";
import cors from "cors";
import multer from "multer";
import AdmZip from "adm-zip";
import path from "path";
import { fileURLToPath } from "url";

import { processDocxBuffer, processPptxBuffer } from "./lib/officeXml.js";
import { stripPdfMetadata, extractPdfText, filterExtractedLines } from "./lib/pdfTools.js";
import { buildReportHtml } from "./lib/report.js";
import { createDocxFromText } from "./lib/docxWriter.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors({ origin: "*", credentials: false }));
app.use(express.json({ limit: "32mb" }));
app.use(express.urlencoded({ extended: true, limit: "32mb" }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// =============== IA: correction & reformulation (ULTRA-CONSERVATEUR) ===============

function safeNormalizeSpaces(raw) {
  if (!raw) return "";
  let t = String(raw);
  t = t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  t = t.replace(/[ \t]{2,}/g, " ");                 // compacter espaces
  t = t.replace(/\s+([:;!?)\]\}])/g, "$1");         // pas d’espace avant : ; ! ? ) ] }
  t = t.replace(/([.:;!?])([A-Za-z0-9])/g, "$1 $2");// espace après . : ; ! ? si lettre/num
  t = t.replace(/\n{3,}/g, "\n\n");                  // paragraphes
  return t.trim();
}

function fixObservedArtifacts(text) {
  let t = text;

  // Artefacts intra-mot (observés dans tes fichiers) : soc ial, enablin g, th e, corpo, rations, rig hts…
  t = t.replace(/\bsoc\s+ial\b/gi, "social");
  t = t.replace(/\bcommu\s+nication\b/gi, "communication");
  t = t.replace(/\benablin\s+g\b/gi, "enabling");
  t = t.replace(/\bth\s+e\b/gi, "the");
  t = t.replace(/\bp\s+otential\b/gi, "potential");
  t = t.replace(/\bdis\s+connection\b/gi, "disconnection");
  t = t.replace(/\brig\s+hts\b/gi, "rights");
  t = t.replace(/\bco\s+mmunication\b/gi, "communication");
  t = t.replace(/\b([A-Za-z])\s{2,}([A-Za-z]+)\b/g, (_m, a, rest) => a + rest); // p  otential -> potential
  t = t.replace(/\b([a-z])\1{2,}([a-z]+)/g, (_m, a, rest) => a + rest);         // gggdigital -> digital
  t = t.replace(/\bcorpo,\s*rations\b/gi, "corporations");                       // corpo, rations -> corporations

  // Suffixes scindés par ESPACE (jamais via virgule/point)
  t = t.replace(
    /\b([A-Za-z]{3,})\s+(tion|sion|ment|ness|ship|ance|ence|hood|ward|tial|cial|ing|ed|ly)\b/g,
    (_m, a, suf) => a + suf
  );

  return t;
}

function sanitizeText(raw) {
  let t = String(raw);
  t = fixObservedArtifacts(t);

  t = t
    .replace(/\bTik,?\s*Tok\b/gi, "TikTok")
    .replace(/\bLinked,?\s*In\b/gi, "LinkedIn")
    .replace(/\bface[- ]?to[- ]?face\b/gi, "face-to-face")
    .replace(/\btwenty[- ]?first\b/gi, "twenty-first")
    .replace(/\bdouble[- ]?edged\b/gi, "double-edged")
    .replace(/\btext[- ]?based\b/gi, "text-based");

  t = safeNormalizeSpaces(t);
  return t;
}

// V1 — correction seulement
async function aiCorrectGrammar(text) {
  return sanitizeText(text);
}

// V2 — reformulation légère (pas de synonymes risqués, pas de collage)
async function aiRephrase(text) {
  let t = await aiCorrectGrammar(text);
  t = t.replace(/\bOverall,?\s+(?=[A-Za-z])/g, "");
  t = t.replace(/\bFurthermore,?\s+/g, "Moreover, ");
  t = t.replace(/;(\s+)/g, ". ");
  return sanitizeText(t);
}

// =============== Helpers ===============

function getExt(filename = "") {
  const i = filename.lastIndexOf(".");
  return i >= 0 ? filename.slice(i + 1).toLowerCase() : "";
}

async function transformOfficeBuffer(buffer, ext, transformFn) {
  if (ext === "docx") return await processDocxBuffer(buffer, transformFn);
  if (ext === "pptx") return await processPptxBuffer(buffer, transformFn);
  throw new Error("Unsupported office type: " + ext);
}

// =============== Health / Echo ===============

app.get("/health", (_, res) =>
  res.json({ ok: true, service: "DocSafe Backend", time: new Date().toISOString() })
);
app.get("/_env_ok", (_, res) =>
  res.json({ OPENAI_API_KEY: process.env.OPENAI_API_KEY ? "present" : "absent" })
);
app.get("/_ai_echo", async (req, res) => {
  const sample = req.query.q || "This is   a  test,, please  fix punctuation!! Thanks";
  res.json({ input: sample, corrected: await aiCorrectGrammar(String(sample)) });
});
app.get("/_ai_rephrase_echo", async (req, res) => {
  const sample = req.query.q || "Overall, this is a sentence that could be rephrased furthermore.";
  res.json({ input: sample, rephrased: await aiRephrase(String(sample)) });
});

// =============== V1: /clean ===============

app.post("/clean", upload.any(), async (req, res) => {
  try {
    const strictPdf = String(req.body.strictPdf || "false") === "true";
    if (!req.files?.length) return res.status(400).json({ error: "No files uploaded." });

    const zip = new AdmZip();

    for (const f of req.files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      if (ext === "docx" || ext === "pptx") {
        const cleanedBuf = await transformOfficeBuffer(f.buffer, ext, aiCorrectGrammar);
        zip.addFile(`${base}/cleaned.${ext}`, cleanedBuf);
        zip.addFile(`${base}/report.html`, Buffer.from(
          buildReportHtml({
            original: "[structured file]",
            cleaned: "[structured file corrected]",
            rephrased: null,
            filename: f.originalname,
            mode: "V1",
          }),
          "utf8"
        ));
      } else if (ext === "pdf") {
        const pdfSan = await stripPdfMetadata(f.buffer);
        const rawText = await extractPdfText(pdfSan);
        const filtered = filterExtractedLines(rawText, { strictPdf });
        const cleaned = await aiCorrectGrammar(filtered);
        const cleanedDocx = await createDocxFromText(cleaned, base || "cleaned");

        zip.addFile(`${base}/pdf_sanitized.pdf`, pdfSan);
        zip.addFile(`${base}/cleaned.docx`, cleanedDocx);
        zip.addFile(`${base}/report.html`, Buffer.from(
          buildReportHtml({ original: filtered, cleaned, rephrased: null, filename: f.originalname, mode: "V1" }),
          "utf8"
        ));
      } else {
        const original = f.buffer.toString("utf8");
        const cleaned = await aiCorrectGrammar(original);
        const cleanedDocx = await createDocxFromText(cleaned, base || "cleaned");

        zip.addFile(`${base}/cleaned.docx`, cleanedDocx);
        zip.addFile(`${base}/report.html`, Buffer.from(
          buildReportHtml({ original, cleaned, rephrased: null, filename: f.originalname, mode: "V1" }),
          "utf8"
        ));
      }
    }

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="docsafe_v1.zip"`);
    return res.send(zip.toBuffer());
  } catch (err) {
    console.error("CLEAN ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
});

// =============== V2: /clean-v2 ===============

app.post("/clean-v2", upload.any(), async (req, res) => {
  try {
    const strictPdf = String(req.body.strictPdf || "false") === "true";
    if (!req.files?.length) return res.status(400).json({ error: "No files uploaded." });

    const zip = new AdmZip();

    for (const f of req.files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      if (ext === "docx" || ext === "pptx") {
        const cleanedBuf = await transformOfficeBuffer(f.buffer, ext, aiCorrectGrammar);
        const rephrasedBuf = await transformOfficeBuffer(f.buffer, ext, aiRephrase);

        zip.addFile(`${base}/cleaned.${ext}`, cleanedBuf);
        zip.addFile(`${base}/rephrased.${ext}`, rephrasedBuf);
        zip.addFile(`${base}/report.html`, Buffer.from(
          buildReportHtml({
            original: "[structured file]",
            cleaned: "[structured file corrected]",
            rephrased: "[structured file rephrased]",
            filename: f.originalname,
            mode: "V2",
          }),
          "utf8"
        ));
      } else if (ext === "pdf") {
        const pdfSan = await stripPdfMetadata(f.buffer);
        const rawText = await extractPdfText(pdfSan);
        const filtered = filterExtractedLines(rawText, { strictPdf });

        const cleaned = await aiCorrectGrammar(filtered);
        const rephrased = await aiRephrase(cleaned);

        const cleanedDocx = await createDocxFromText(cleaned, base || "cleaned");
        const rephrasedDocx = await createDocxFromText(rephrased, base || "rephrased");

        zip.addFile(`${base}/pdf_sanitized.pdf`, pdfSan);
        zip.addFile(`${base}/cleaned.docx`, cleanedDocx);
        zip.addFile(`${base}/rephrased.docx`, rephrasedDocx);
        zip.addFile(`${base}/report.html`, Buffer.from(
          buildReportHtml({ original: filtered, cleaned, rephrased, filename: f.originalname, mode: "V2" }),
          "utf8"
        ));
      } else {
        const original = f.buffer.toString("utf8");
        const cleaned = await aiCorrectGrammar(original);
        const rephrased = await aiRephrase(cleaned);
        const cleanedDocx = await createDocxFromText(cleaned, base || "cleaned");
        const rephrasedDocx = await createDocxFromText(rephrased, base || "rephrased");

        zip.addFile(`${base}/cleaned.docx`, cleanedDocx);
        zip.addFile(`${base}/rephrased.docx`, rephrasedDocx);
        zip.addFile(`${base}/report.html`, Buffer.from(
          buildReportHtml({ original, cleaned, rephrased, filename: f.originalname, mode: "V2" }),
          "utf8"
        ));
      }
    }

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="docsafe_v2.zip"`);
    return res.send(zip.toBuffer());
  } catch (err) {
    console.error("CLEAN V2 ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
});

// ================== DEBUG ==================

app.get("/_ping", (req, res) => res.json({ pong: true, time: new Date().toISOString() }));

app.post("/_pdf_test", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No PDF uploaded" });
    const pdfSan = await stripPdfMetadata(req.file.buffer);
    const rawText = await extractPdfText(pdfSan);
    const filtered = filterExtractedLines(rawText, { strictPdf: true });
    res.json({
      filename: req.file.originalname,
      size: req.file.size,
      rawLength: rawText.length,
      filteredLength: filtered.length,
      excerpt: filtered.slice(0, 500)
    });
  } catch (e) {
    console.error("DEBUG _pdf_test error:", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.post("/_office_test", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const ext = req.file.originalname.split(".").pop().toLowerCase();
    let buf;
    if (ext === "docx") {
      buf = await processDocxBuffer(req.file.buffer, (txt) => txt.toUpperCase());
    } else if (ext === "pptx") {
      buf = await processPptxBuffer(req.file.buffer, (txt) => txt.toUpperCase());
    } else {
      return res.status(400).json({ error: "Only DOCX or PPTX allowed" });
    }
    res.setHeader("Content-Type", "application/octet-stream");
    res.setHeader("Content-Disposition", `attachment; filename="debug_out.${ext}"`);
    res.send(buf);
  } catch (e) {
    console.error("DEBUG _office_test error:", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// =============== Listen ===============

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend listening on ${PORT}`));
