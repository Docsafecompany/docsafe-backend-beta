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

// =============== IA: correction & reformulation (sans API, déterministe) ===============

function normalizePunctAndSpaces(raw) {
  if (!raw) return "";
  let t = String(raw);
  t = t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  t = t.split("\n").map(l => l.trim()).join("\n");

  t = t.replace(/([.,;:!?])([^\s\n])/g, "$1 $2"); // espace après ponctuation
  t = t.replace(/\s+([.,;:!?])/g, "$1");          // pas d'espace avant ponctuation
  t = t.replace(/,{2,}/g, ",").replace(/;{2,}/g, ";").replace(/:{2,}/g, ":")
       .replace(/!{2,}/g, "!").replace(/\?{2,}/g, "?").replace(/@{2,}/g, "@");
  t = t.replace(/[ \t]{2,}/g, " ");               // doubles espaces
  t = t.replace(/\n{3,}/g, "\n\n");               // paragraphes
  return t.trim();
}

// Join intraword artifacts: "commu nication"->"communication", "corpo, rations"->"corporations", "th e"->"the"
function fixIntrawordArtifacts(text) {
  let t = text;

  // Ponctuation au milieu ("corpo, rations" -> "corporations")
  t = t.replace(/\b([A-Za-z]{2,})[,\-\/\.]\s*([A-Za-z]{2,})\b/g, (m, a, b) => a + b);

  // Espace intra-mot ("commu nication", "rig hts", "enablin g")
  t = t.replace(/\b([a-z]{2,})\s+([a-z]{1,})\b/g, (m, a, b) => {
    const merged = a + b;
    // heuristique: si les morceaux sont courts ou si merged a l'air plausible
    if (a.length <= 3 || b.length <= 3) return merged;
    if (/(tion|sion|ment|ness|ship|able|ible|ance|ence|hood|ward|tial|cial|ing|ed|ly)$/i.test(merged)) return merged;
    return m;
  });

  // Lettres multiples en tête ("gggdigital" -> "digital")
  t = t.replace(/\b([a-z])\1{2,}([a-z]+)/g, (m, a, rest) => a + rest);

  return t;
}

function sanitizeText(raw) {
  let t = normalizePunctAndSpaces(raw);

  // recollages
  t = fixIntrawordArtifacts(t);

  // Normalisations spécifiques vues couramment
  t = t
    .replace(/\bTik,?\s*Tok\b/gi, "TikTok")
    .replace(/\bLinked,?\s*In\b/gi, "LinkedIn")
    .replace(/\bX\s*\(formerly Twitter\)/gi, "X (formerly Twitter)")
    .replace(/\bface[- ]?to[- ]?face\b/gi, "face-to-face")
    .replace(/\btwenty[- ]?first\b/gi, "twenty-first")
    .replace(/\bdouble[- ]?edged\b/gi, "double-edged")
    .replace(/\btext[- ]?based\b/gi, "text-based");

  return normalizePunctAndSpaces(t);
}

async function aiCorrectGrammar(text) {
  // Ici, pipeline heuristique sans API :
  return sanitizeText(text);
}

async function aiRephrase(text) {
  // Repart de la version corrigée pour éviter de rephraser du bruit
  let t = await aiCorrectGrammar(text);

  // Synonymes simples / restructurations légères
  t = t
    .replace(/\bOne of the most significant\b/gi, "A major")
    .replace(/\bhas redefined\b/gi, "has reshaped")
    .replace(/\bcontributed to\b/gi, "has led to")
    .replace(/\bIn conclusion\b/gi, "To conclude")
    .replace(/;(\s+)/g, ". "); // couper après un ;

  // nettoyage final
  return sanitizeText(t);
}

// =============== Helpers ===============

function getExt(filename = "") {
  const i = filename.lastIndexOf(".");
  return i >= 0 ? filename.slice(i + 1).toLowerCase() : "";
}

// Applique une transformFn sur DOCX/PPTX en conservant structure
async function transformOfficeBuffer(buffer, ext, transformFn) {
  if (ext === "docx") return await processDocxBuffer(buffer, transformFn);
  if (ext === "pptx") return await processPptxBuffer(buffer, transformFn);
  throw new Error("Unsupported office type: " + ext);
}

// =============== Health / Echo ===============

app.get("/health", (_, res) => res.json({ ok: true, service: "DocSafe Backend", time: new Date().toISOString() }));
app.get("/_env_ok", (_, res) => res.json({ OPENAI_API_KEY: process.env.OPENAI_API_KEY ? "present" : "absent" }));
app.get("/_ai_echo", async (req, res) => {
  const sample = req.query.q || "This is   a  test,, please  fix punctuation!!Thanks";
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
        // Correction IA dans la structure
        const cleanedBuf = await transformOfficeBuffer(f.buffer, ext, aiCorrectGrammar);
        // Rapport (extraits : on lit vite-fait le texte brut en interne)
        const originalText = "[structured file]";
        const cleanedText = "[structured file corrected]";
        // Ajout dans un sous-dossier pour éviter d'écraser si multi-fichiers
        zip.addFile(`${base}/cleaned.${ext}`, cleanedBuf);
        zip.addFile(`${base}/report.html`, Buffer.from(
          buildReportHtml({ original: originalText, cleaned: cleanedText, rephrased: null, filename: f.originalname, mode: "V1" }),
          "utf8"
        ));
      } else if (ext === "pdf") {
        // PDF : métadonnées nettoyées + extraction + correction → DOCX
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
        // Fallback
        const cleaned = await aiCorrectGrammar(f.buffer.toString("utf8"));
        const cleanedDocx = await createDocxFromText(cleaned, base || "cleaned");
        zip.addFile(`${base}/cleaned.docx`, cleanedDocx);
        zip.addFile(`${base}/report.html`, Buffer.from(
          buildReportHtml({ original: f.buffer.toString("utf8"), cleaned, rephrased: null, filename: f.originalname, mode: "V1" }),
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
        // V1 (cleaned)
        const cleanedBuf = await transformOfficeBuffer(f.buffer, ext, aiCorrectGrammar);
        // V2 (rephrased)
        const rephrasedBuf = await transformOfficeBuffer(f.buffer, ext, aiRephrase);

        zip.addFile(`${base}/cleaned.${ext}`, cleanedBuf);
        zip.addFile(`${base}/rephrased.${ext}`, rephrasedBuf);
        zip.addFile(`${base}/report.html`, Buffer.from(
          buildReportHtml({
            original: "[structured file]",
            cleaned: "[structured file corrected]",
            rephrased: "[structured file rephrased]",
            filename: f.originalname,
            mode: "V2"
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

// =============== Server listen ===============

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend listening on ${PORT}`));
