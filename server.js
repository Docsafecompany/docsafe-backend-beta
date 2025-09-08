// server.js
// DocSafe Backend (Express + Multer + ZIP + minimal "AI" stubs)
// Endpoints:
//  - POST /clean      -> V1: metadata cleanup + IA orthographe + LT report
//  - POST /clean-v2   -> V2: V1 + reformulation IA
//  - GET  /health, GET /_env_ok, GET /_ai_echo, GET /_ai_rephrase_echo
// Formats: PDF / DOCX / PPTX
// Option strictPdf: best-effort (l'extraction texte ignore déjà la plupart des calques invisibles)

import express from "express";
import cors from "cors";
import multer from "multer";
import AdmZip from "adm-zip";
import path from "path";
import { fileURLToPath } from "url";
import { extractTextFromDocx, extractTextFromPptx, extractTextFromPdf } from "./lib/textExtractors.js";
import { createDocxFromText } from "./lib/docxWriter.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

// --- CORS (corrigé) ---
app.use(cors({ origin: "*", credentials: false }));
app.use(express.json({ limit: "32mb" }));
app.use(express.urlencoded({ extended: true, limit: "32mb" }));

// --- Multer (corrigé: any()) ---
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// ------------------ Utils ------------------

/**
 * Normalisation du texte: espaces/ponctuation/retours à la ligne
 * - supprime doublons de ponctuation (,, ;; :: !! ?? @@)
 * - normalise espaces avant/après la ponctuation
 * - corrige doubles espaces
 * - corrige quelques collages fréquents (face-to-face, twenty-first, etc.)
 * - assure des sauts de ligne propres
 */
function sanitizeText(raw) {
  if (!raw) return "";

  let t = raw;

  // Unifier les fins de ligne
  t = t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

  // Supprimer espaces à début/fin de ligne
  t = t
    .split("\n")
    .map((l) => l.trim())
    .join("\n");

  // Mettre un espace après .,;:!? si une lettre/num suit (sauf fin de ligne)
  t = t.replace(/([.,;:!?])([^\s\n])/g, "$1 $2");

  // Enlever espace avant .,;:!?)
  t = t.replace(/\s+([.,;:!?])/g, "$1");

  // Supprimer doublons de ponctuation listés
  t = t.replace(/,{2,}/g, ",");
  t = t.replace(/;{2,}/g, ";");
  t = t.replace(/:{2,}/g, ":");
  t = t.replace(/!{2,}/g, "!");
  t = t.replace(/\?{2,}/g, "?");
  t = t.replace(/@{2,}/g, "@");

  // Doubles espaces -> simple
  t = t.replace(/[ \t]{2,}/g, " ");

  // Collages courants (sans exagérer pour ne pas casser des mots corrects)
  // "faceto-face" -> "face-to-face"
  t = t.replace(/\bfaceto[- ]face\b/gi, "face-to-face");
  // "twentyfirst" -> "twenty-first"
  t = t.replace(/\btwenty[- ]?first\b/gi, "twenty-first");
  // "doubleedged" -> "double-edged"
  t = t.replace(/\bdouble[- ]?edged\b/gi, "double-edged");
  // "textbased" -> "text-based"
  t = t.replace(/\btext[- ]?based\b/gi, "text-based");
  // "face to face" -> "face-to-face" (si séparé)
  t = t.replace(/\bface to face\b/gi, "face-to-face");
  // "real life" -> "real life" (laisser), mais corriger "reallife"
  t = t.replace(/\breal[- ]?life\b/gi, "real life");

  // Collages typographiques issus d'extractions PDF foireuses: "centurysoc ial" -> "century social"
  // Heuristique: si un mot contient une coupure "so cial", recoller intelligemment
  t = t.replace(/\bsoc ial\b/gi, "social");
  t = t.replace(/\binter personal\b/gi, "interpersonal");
  t = t.replace(/\bnon profit\b/gi, "nonprofit");

  // Ajouter une ligne vide entre paragraphes s'il n'y en a pas
  t = t.replace(/\n{3,}/g, "\n\n");

  // Trim final
  t = t.trim();

  return t;
}

/**
 * "IA" de correction orthographique/grammaire (placeholder).
 * Ici on applique sanitize + quelques micro-heuristiques.
 * Si tu veux brancher OpenAI/LanguageTool, fais-le ici.
 */
async function aiCorrectGrammar(text) {
  // Ici, on applique déjà sanitizeText qui règle la majorité des cas visibles dans tes fichiers.
  // Tu peux ajouter un appel OpenAI si OPENAI_API_KEY est défini.
  const normalized = sanitizeText(text);

  // Démo: corrections simples de tokens fréquents mal extraits
  let t = normalized;
  t = t.replace(/\bTik,?\s*Tok\b/gi, "TikTok");
  t = t.replace(/\bLinked,?\s*In\b/gi, "LinkedIn");
  t = t.replace(/\bX\s*\(formerly Twitter\)/gi, "X (formerly Twitter)");
  t = t.replace(/\bface[- ]?to[- ]?face\b/gi, "face-to-face");

  return t;
}

/**
 * "IA" de reformulation (placeholder).
 * Ici on paraphrase légèrement: on supprime certains "Overall"/"Furthermore" abusifs et on allège.
 */
async function aiRephrase(text) {
  let t = sanitizeText(text);

  // Eviter répétitions stylistiques ("Overall", "Furthermore" en début de phrases)
  t = t.replace(/\bOverall,?\s*/g, "");
  t = t.replace(/\bFurthermore,?\s*/g, "Additionally, ");
  t = t.replace(/\bIn conclusion,?\s*/gi, "To conclude, ");

  // Petits allégements
  t = t.replace(/\bIt is\b/gi, "It's");
  t = t.replace(/\bIt has\b/gi, "It’s");

  // Dernière passe de normalisation
  t = sanitizeText(t);
  return t;
}

// Bricolage simple de "rapport LT" HTML
function buildLtReportHtml({ original, cleaned, rephrased, filename, mode }) {
  const originalLen = original?.length || 0;
  const cleanedLen = cleaned?.length || 0;
  const rephrasedLen = rephrased?.length || 0;

  const ts = new Date().toISOString();
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>DocSafe Report</title>
<style>
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:24px;line-height:1.5}
  h1{font-size:20px;margin:0 0 8px}
  h2{font-size:16px;margin:16px 0 8px}
  code,pre{background:#f6f7f9;padding:8px;border-radius:6px;display:block;white-space:pre-wrap}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
  .meta{color:#555}
</style>
</head>
<body>
  <h1>DocSafe Report</h1>
  <p class="meta">File: <strong>${filename}</strong> • Mode: <strong>${mode}</strong> • ${ts}</p>

  <h2>Summary</h2>
  <ul>
    <li>Original length: ${originalLen} chars</li>
    <li>Cleaned length: ${cleanedLen} chars</li>
    ${mode === "V2" ? `<li>Rephrased length: ${rephrasedLen} chars</li>` : ""}
    <li>Normalizations: spaces/ponctuation de-duplicated, spacing after punctuation, basic hyphenation fixes.</li>
  </ul>

  <div class="grid">
    <div>
      <h2>Original (excerpt)</h2>
      <pre>${escapeHtml(original?.slice(0, 2000) || "")}</pre>
    </div>
    <div>
      <h2>Cleaned (excerpt)</h2>
      <pre>${escapeHtml(cleaned?.slice(0, 2000) || "")}</pre>
    </div>
  </div>

  ${mode === "V2" ? `
  <h2>Rephrased (excerpt)</h2>
  <pre>${escapeHtml(rephrased?.slice(0, 2000) || "")}</pre>` : ""}

</body>
</html>`;
}

function escapeHtml(s) {
  return (s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// Détection simple d'extension
function getExt(filename = "") {
  const ix = filename.lastIndexOf(".");
  return ix >= 0 ? filename.slice(ix + 1).toLowerCase() : "";
}

// Lecture texte selon type
async function readTextByType(buffer, filename, strictPdf) {
  const ext = getExt(filename);
  if (ext === "docx") return await extractTextFromDocx(buffer);
  if (ext === "pptx") return await extractTextFromPptx(buffer);
  if (ext === "pdf") return await extractTextFromPdf(buffer, { strictPdf });
  // fallback brut
  return buffer.toString("utf8");
}

// ------------------ Endpoints ------------------

app.get("/health", (req, res) => {
  res.json({ ok: true, service: "DocSafe Backend", time: new Date().toISOString() });
});

app.get("/_env_ok", (req, res) => {
  const hasKey = !!process.env.OPENAI_API_KEY;
  res.json({ OPENAI_API_KEY: hasKey ? "present" : "absent" });
});

app.get("/_ai_echo", async (req, res) => {
  const sample = req.query.q || "This is   a  test,, please  fix punctuation!!Thanks";
  const corrected = await aiCorrectGrammar(String(sample));
  res.json({ input: sample, corrected });
});

app.get("/_ai_rephrase_echo", async (req, res) => {
  const sample = req.query.q || "Overall, this is a sentence that could be rephrased furthermore.";
  const rephrased = await aiRephrase(String(sample));
  res.json({ input: sample, rephrased });
});

// V1
app.post("/clean", upload.any(), async (req, res) => {
  try {
    const strictPdf = String(req.body.strictPdf || "false") === "true";

    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: "No files uploaded." });
    }

    const zip = new AdmZip();

    for (const f of req.files) {
      const originalText = await readTextByType(f.buffer, f.originalname, strictPdf);
      const corrected = await aiCorrectGrammar(originalText);
      const cleaned = sanitizeText(corrected);

      // DOCX out
      const cleanedDocxBuffer = await createDocxFromText(cleaned, path.parse(f.originalname).name || "cleaned");

      // Report
      const reportHtml = buildLtReportHtml({
        original: originalText,
        cleaned,
        rephrased: null,
        filename: f.originalname,
        mode: "V1",
      });

      // Add to ZIP (names per spec)
      zip.addFile("cleaned.docx", cleanedDocxBuffer);
      zip.addFile("report.html", Buffer.from(reportHtml, "utf8"));
    }

    const out = zip.toBuffer();
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="docsafe_v1.zip"`);
    return res.send(out);
  } catch (err) {
    console.error("CLEAN ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
});

// V2
app.post("/clean-v2", upload.any(), async (req, res) => {
  try {
    const strictPdf = String(req.body.strictPdf || "false") === "true";

    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: "No files uploaded." });
    }

    const zip = new AdmZip();

    for (const f of req.files) {
      const originalText = await readTextByType(f.buffer, f.originalname, strictPdf);
      const corrected = await aiCorrectGrammar(originalText);
      const cleaned = sanitizeText(corrected);
      const rephrased = await aiRephrase(cleaned);

      // DOCX out
      const cleanedDocxBuffer = await createDocxFromText(cleaned, path.parse(f.originalname).name || "cleaned");
      const rephrasedDocxBuffer = await createDocxFromText(rephrased, path.parse(f.originalname).name || "rephrased");

      // Report
      const reportHtml = buildLtReportHtml({
        original: originalText,
        cleaned,
        rephrased,
        filename: f.originalname,
        mode: "V2",
      });

      // Add to ZIP
      zip.addFile("cleaned.docx", cleanedDocxBuffer);
      zip.addFile("rephrased.docx", rephrasedDocxBuffer);
      zip.addFile("report.html", Buffer.from(reportHtml, "utf8"));
    }

    const out = zip.toBuffer();
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="docsafe_v2.zip"`);
    return res.send(out);
  } catch (err) {
    console.error("CLEAN V2 ERROR", err);
    return res.status(500).json({ error: String(err?.message || err) });
  }
});

// Port
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log(`DocSafe backend listening on port ${PORT}`);
});

