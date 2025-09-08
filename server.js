// server.js
import express from "express";
import cors from "cors";
import multer from "multer";
import AdmZip from "adm-zip";
import path from "path";
import { fileURLToPath } from "url";

import { extractTextFromDocx, extractTextFromPptx, extractTextFromPdf } from "./lib/textExtractors.js";
import { createDocxFromText } from "./lib/docxWriter.js";

// --------- nspell (offline spellchecker) ----------
import nspell from "nspell";
import { readFile } from "fs/promises";
import { createRequire } from "module";
const require = createRequire(import.meta.url);
const dictAffPath = require.resolve("dictionary-en-us/index.aff");
const dictDicPath = require.resolve("dictionary-en-us/index.dic");

let SPELL = null;
try {
  const [aff, dic] = await Promise.all([readFile(dictAffPath), readFile(dictDicPath)]);
  SPELL = nspell(aff, dic);
} catch (e) {
  console.warn("Warning: English dictionary failed to load. Heuristics only.", e?.message || e);
}

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors({ origin: "*", credentials: false }));
app.use(express.json({ limit: "32mb" }));
app.use(express.urlencoded({ extended: true, limit: "32mb" }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

// ------------------ Utils ------------------

function normalizePunctAndSpaces(raw) {
  if (!raw) return "";
  let t = String(raw);

  // EOL
  t = t.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

  // trim par ligne
  t = t.split("\n").map(l => l.trim()).join("\n");

  // espaces autour ponctuation
  t = t.replace(/([.,;:!?])([^\s\n])/g, "$1 $2");
  t = t.replace(/\s+([.,;:!?])/g, "$1");

  // doublons ponctuation
  t = t.replace(/,{2,}/g, ",").replace(/;{2,}/g, ";").replace(/:{2,}/g, ":").replace(/!{2,}/g, "!").replace(/\?{2,}/g, "?").replace(/@{2,}/g, "@");

  // espaces multiples
  t = t.replace(/[ \t]{2,}/g, " ");

  // paragraphes
  t = t.replace(/\n{3,}/g, "\n\n");

  return t.trim();
}

// joints inter-mots cassés par ponctuation ou espace (« corpo, rations », « commu nication », « th e »)
function fixIntrawordArtifacts(text) {
  let t = text;

  // 1) Ponctuation au milieu d’un mot: "corpo, rations" -> "corporations"
  t = t.replace(/\b([A-Za-z]{2,})[,\-\/\.]\s*([A-Za-z]{2,})\b/g, (m, a, b) => {
    const merged = a + b;
    return shouldMerge(a, b, merged) ? merged : `${a} ${b}`;
  });

  // 2) Espace intra-mot entre lettres minuscules: "commu nication", "rig hts", "enablin g"
  // On ne merge que si le merge donne un mot plus plausible (spellcheck) ou si les morceaux sont très courts.
  t = t.replace(/\b([a-z]{2,})\s+([a-z]{1,})\b/g, (m, a, b) => {
    const merged = a + b;
    if (shouldMerge(a, b, merged)) return merged;
    return m;
  });

  // 3) Lettres multiples en tête de mot (gggdigital -> digital)
  t = t.replace(/\b([a-z])\1{2,}([a-z]+)/g, (m, a, rest) => a + rest);

  return t;
}

function shouldMerge(a, b, merged) {
  // Règles sûres :
  if (a.length <= 3 || b.length <= 3) return true; // "th e", "rig hts", "enablin g" → ok
  if (merged.length <= 6) return true;           // petits mots
  // Spellchecker si dispo
  if (SPELL) {
    const mergedOk = SPELL.correct(merged);
    const aOk = SPELL.correct(a);
    const bOk = SPELL.correct(b);
    if (mergedOk && !(aOk && bOk)) return true;
  }
  // fallback heuristique : patterns fréquents
  if (/(tion|sion|ment|ness|ship|able|ible|ial|ial|ance|ence|hood|ward|tial|cial|cial)$/i.test(merged)) return true;
  return false;
}

function sanitizeText(raw) {
  let t = normalizePunctAndSpaces(raw);

  // joindre "so cial", "inter personal", etc.
  t = fixIntrawordArtifacts(t);

  // Recollages et normalisations spécifiques vus dans tes fichiers
  t = t
    .replace(/\bTik,?\s*Tok\b/gi, "TikTok")
    .replace(/\bLinked,?\s*In\b/gi, "LinkedIn")
    .replace(/\bX\s*\(formerly Twitter\)/gi, "X (formerly Twitter)")
    .replace(/\bface[- ]?to[- ]?face\b/gi, "face-to-face")
    .replace(/\btwenty[- ]?first\b/gi, "twenty-first")
    .replace(/\bdouble[- ]?edged\b/gi, "double-edged")
    .replace(/\btext[- ]?based\b/gi, "text-based");

  // deuxième passe espaces/ponctuation après merges
  t = normalizePunctAndSpaces(t);
  return t;
}

// IA "correction" (offline) = sanitize + micro ajustements
async function aiCorrectGrammar(text) {
  let t = sanitizeText(text);

  // Si dico dispo, corriger quelques mots isolés évidents (ex: "teh" -> "the")
  if (SPELL) {
    t = t.replace(/\b([A-Za-z]{2,})\b/g, (m, w) => (SPELL.correct(w) ? w : (SPELL.suggest(w)[0] || w)));
  }

  return t;
}

// IA "rephrase" (offline) = rephraser déterministe pour avoir une vraie diff
async function aiRephrase(text) {
  // On part du corrigé pour éviter de rephraser du bruit
  let t = await aiCorrectGrammar(text);

  // 1) Synonymes/variantes
  t = t
    .replace(/\bOne of the most significant\b/gi, "A major")
    .replace(/\bAdditionally\b/gi, "Moreover")
    .replace(/\bAdditionally,\s/gi, "Moreover, ")
    .replace(/\bAdditionally\b/gi, "Furthermore")
    .replace(/\bhas redefined\b/gi, "has reshaped")
    .replace(/\bcontributed to\b/gi, "has led to")
    .replace(/\bIt has\b/g, "It’s")
    .replace(/\bIt is\b/g, "It’s")
    .replace(/\bIn conclusion\b/gi, "To conclude");

  // 2) Restructuration simple de phrases longues (coupe après ; ou : si present)
  t = t.replace(/;(\s+)/g, ". ");

  // 3) Dernière passe de normalisation
  t = sanitizeText(t);
  return t;
}

function escapeHtml(s) {
  return (s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function buildLtReportHtml({ original, cleaned, rephrased, filename, mode }) {
  const ts = new Date().toISOString();
  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8"/>
<title>DocSafe Report</title>
<style>
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:24px;line-height:1.5}
h1{font-size:20px;margin:0 0 8px}h2{font-size:16px;margin:16px 0 8px}
code,pre{background:#f6f7f9;padding:8px;border-radius:6px;display:block;white-space:pre-wrap}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}.meta{color:#555}
</style></head>
<body>
<h1>DocSafe Report</h1>
<p class="meta">File: <strong>${filename}</strong> • Mode: <strong>${mode}</strong> • ${ts}</p>
<h2>Summary</h2>
<ul>
  <li>Original length: ${(original||"").length} chars</li>
  <li>Cleaned length: ${(cleaned||"").length} chars</li>
  ${mode==="V2"?`<li>Rephrased length: ${(rephrased||"").length} chars</li>`:""}
  <li>Operations: intraword merge, de-dup punctuation, spelling (nspell), hyphenation fixes.</li>
</ul>
<div class="grid">
  <div><h2>Original (excerpt)</h2><pre>${escapeHtml((original||"").slice(0,2000))}</pre></div>
  <div><h2>${mode==="V2"?"Rephrased":"Cleaned"} (excerpt)</h2><pre>${escapeHtml(((mode==="V2"?rephrased:cleaned)||"").slice(0,2000))}</pre></div>
</div>
</body></html>`;
}

function getExt(filename = "") {
  const ix = filename.lastIndexOf(".");
  return ix >= 0 ? filename.slice(ix + 1).toLowerCase() : "";
}

async function readTextByType(buffer, filename, strictPdf) {
  const ext = getExt(filename);
  if (ext === "docx") return await extractTextFromDocx(buffer);
  if (ext === "pptx") return await extractTextFromPptx(buffer);
  if (ext === "pdf")  return await extractTextFromPdf(buffer, { strictPdf });
  return buffer.toString("utf8");
}

// ------------------ Endpoints ------------------

app.get("/health", (req, res) => res.json({ ok: true, service: "DocSafe Backend", time: new Date().toISOString() }));
app.get("/_env_ok", (req, res) => res.json({ OPENAI_API_KEY: process.env.OPENAI_API_KEY ? "present" : "absent" }));
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

// V1: cleaned.docx + report.html
app.post("/clean", upload.any(), async (req, res) => {
  try {
    const strictPdf = String(req.body.strictPdf || "false") === "true";
    if (!req.files?.length) return res.status(400).json({ error: "No files uploaded." });

    const zip = new AdmZip();
    for (const f of req.files) {
      const originalText = await readTextByType(f.buffer, f.originalname, strictPdf);
      const cleaned = await aiCorrectGrammar(originalText);
      const cleanedDocxBuffer = await createDocxFromText(cleaned, path.parse(f.originalname).name || "cleaned");
      const reportHtml = buildLtReportHtml({ original: originalText, cleaned, rephrased: null, filename: f.originalname, mode: "V1" });

      zip.addFile("cleaned.docx", cleanedDocxBuffer);
      zip.addFile("report.html", Buffer.from(reportHtml, "utf8"));
    }
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="docsafe_v1.zip"`);
    res.send(zip.toBuffer());
  } catch (err) {
    console.error("CLEAN ERROR", err);
    res.status(500).json({ error: String(err?.message || err) });
  }
});

// V2: cleaned.docx + rephrased.docx + report.html
app.post("/clean-v2", upload.any(), async (req, res) => {
  try {
    const strictPdf = String(req.body.strictPdf || "false") === "true";
    if (!req.files?.length) return res.status(400).json({ error: "No files uploaded." });

    const zip = new AdmZip();
    for (const f of req.files) {
      const originalText = await readTextByType(f.buffer, f.originalname, strictPdf);
      const cleaned = await aiCorrectGrammar(originalText);
      const rephrased = await aiRephrase(cleaned);

      const cleanedDocxBuffer   = await createDocxFromText(cleaned,   path.parse(f.originalname).name || "cleaned");
      const rephrasedDocxBuffer = await createDocxFromText(rephrased, path.parse(f.originalname).name || "rephrased");
      const reportHtml = buildLtReportHtml({ original: originalText, cleaned, rephrased, filename: f.originalname, mode: "V2" });

      zip.addFile("cleaned.docx", cleanedDocxBuffer);
      zip.addFile("rephrased.docx", rephrasedDocxBuffer);
      zip.addFile("report.html", Buffer.from(reportHtml, "utf8"));
    }
    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="docsafe_v2.zip"`);
    res.send(zip.toBuffer());
  } catch (err) {
    console.error("CLEAN V2 ERROR", err);
    res.status(500).json({ error: String(err?.message || err) });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend listening on ${PORT}`));

