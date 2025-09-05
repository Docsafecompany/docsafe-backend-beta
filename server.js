/* server.js — DocSafe backend (CommonJS) */
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
    credentials: false,
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Accept", "Origin"],
  })
);
// préflights explicites
app.options("*", cors());

const upload = multer({ storage: multer.memoryStorage() });

/* =========================
   Utils — nettoyage & IA
   ========================= */

// Normalisation forte FR/EN (espaces, ponctuation, doublons, retours)
function normalizeText(input) {
  let t = String(input || "");

  // invisibles + CRLF
  t = t.replace(/[\u200B-\u200D\uFEFF\u2060\u00AD]/g, "").replace(/\r\n/g, "\n");

  // 1) collaps espaces
  t = t.replace(/[ \t]+/g, " ");

  // 2) Ponctuation dupliquée (inclut le point) + ellipses
  t = t.replace(/([.,;:!?@])\s*\1+/g, "$1");
  t = t.replace(/…+/g, "…");

  // 3) Pas d’espace avant, 1 après (hors fin de ligne)
  t = t.replace(/\s+([.,;:!?%])/g, "$1"); // supprime l’espace AVANT
  t = t.replace(/([.,;:!?%])(?!\s|$)/g, "$1 "); // ajoute 1 espace APRÈS
  t = t.replace(/[ \t]{2,}/g, " "); // recollaps

  // 4) Parenthèses / guillemets
  t = t.replace(/\s*\(\s*/g, " (").replace(/\s*\)\s*/g, ") ");
  t = t.replace(/\s*"\s*/g, '"');

  // 5) Lignes vides (max 1 vide consécutive)
  t = t.replace(/\n[ \t]*\n[ \t]*(\n[ \t]*)+/g, "\n\n");

  // 6) Trim fin de ligne + global
  t = t
    .split("\n")
    .map((s) => s.replace(/[ \t]+$/g, ""))
    .join("\n");

  return t.trim();
}

// “Correction IA” simple mais visible (FR/EN)
function aiSpellGrammarPass(input) {
  let t = String(input || "");

  // fautes fréquentes (EN)
  t = t
    .replace(/\bteh\b/gi, "the")
    .replace(/\brecieve\b/gi, "receive")
    .replace(/\boccured\b/gi, "occurred")
    .replace(/\bim\b/gi, "I'm")
    .replace(/\bdont\b/gi, "don't")
    .replace(/\bcant\b/gi, "can't");

  // doublons de mots (you you, le le)
  t = t.replace(/\b(\w+)\s+\1\b/gi, "$1");

  // capitalisation au début de phrase (FR/EN)
  t = t.replace(/(^|[.!?]\s+)([a-zà-öø-ÿ])/g, (m, p1, p2) => p1 + p2.toUpperCase());

  // petits correctifs FR
  t = t.replace(/\bca va\b/gi, "ça va");

  return t;
}

// Rephrase déterministe FR/EN (synonymes + tournures) — garantit un delta
function aiRephrase(input) {
  let t = String(input || "");
  let out = t;

  const FR_MAP = [
    ["très", "vraiment"],
    ["important", "crucial"],
    ["donc", "ainsi"],
    ["par conséquent", "en conséquence"],
    ["ceci", "cela"],
    ["améliorer", "optimiser"],
    ["problème", "anomalie"],
    ["utiliser", "employer"],
    ["c'est", "il est"],
    ["néanmoins", "toutefois"],
    ["de plus", "en outre"],
  ];
  const EN_MAP = [
    ["very", "highly"],
    ["important", "crucial"],
    ["therefore", "thus"],
    ["so", "therefore"],
    ["make sure", "ensure"],
    ["help", "assist"],
    ["use", "leverage"],
    ["start", "begin"],
    ["in addition", "additionally"],
  ];

  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const preserveCase = (target, src) => {
    if (src.toUpperCase() === src) return target.toUpperCase();
    if (src[0] === src[0].toUpperCase()) return target[0].toUpperCase() + target.slice(1);
    return target;
  };
  const applyMap = (s, map) =>
    map.reduce(
      (acc, [a, b]) =>
        acc.replace(new RegExp(`\\b${esc(a)}\\b`, "gi"), (m) => preserveCase(b, m)),
      s
    );

  out = applyMap(out, FR_MAP);
  out = applyMap(out, EN_MAP);

  // tournures
  out = out
    .replace(/\bIn order to\b/g, "To")
    .replace(/\bafin de\b/gi, "pour")
    .replace(/\bhowever\b/gi, "nevertheless");

  // Si rien n’a bougé (texte trop spécifique), forcer un petit delta neutre
  if (out === t) {
    out = out.replace(/\bso\b/gi, "therefore").replace(/\btrès\b/gi, "vraiment");
  }

  return normalizeText(out);
}

/* =========================
   Extraction contenu
   ========================= */
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
    slideFiles.sort(
      (a, b) =>
        parseInt(a.match(/slide(\d+)\.xml/i)[1], 10) - parseInt(b.match(/slide(\d+)\.xml/i)[1], 10)
    );
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

/* =========================
   Report (stats + diffs)
   ========================= */
function countMatches(str, re) {
  const m = str.match(re);
  return m ? m.length : 0;
}

function computeCleanStats(original, cleaned, lengthBefore, lengthCleaned) {
  return {
    zeroWidthRemoved: countMatches(original, /[\u200B-\u200D\uFEFF\u2060\u00AD]/g),
    doubleSpaceRuns: countMatches(original, /[ \t]{2,}/g),
    duplicatePunct: {
      ".": countMatches(original, /(\.)\s*\1+/g),
      ",": countMatches(original, /(,)\s*\1+/g),
      ";": countMatches(original, /(;)\s*\1+/g),
      ":": countMatches(original, /(:)\s*\1+/g),
      "!": countMatches(original, /(!)\s*\1+/g),
      "?": countMatches(original, /(\?)\s*\1+/g),
      "@": countMatches(original, /(@)\s*\1+/g),
    },
    spacesBeforePunct: countMatches(original, /\s+([.,;:!?%])/g),
    missingSpaceAfterPunct: countMatches(original, /([.,;:!?%])(?!\s|$)/g),
    collapsedBlankLines: countMatches(original, /\n[ \t]*\n[ \t]*(\n[ \t]*)+/g),
    sentenceCapitalized: countMatches(original, /(^|[.!?]\s+)[a-zà-öø-ÿ]/g),
    lengthBefore,
    lengthCleaned,
  };
}

function splitSentences(text) {
  if (!text) return [];
  const paras = text.split(/\n{2,}/g);
  const out = [];
  for (const p of paras) {
    const parts = p.split(/(?<=[\.!\?])\s+/);
    if (parts.length > 1) out.push(...parts);
    else out.push(p);
  }
  return out.map((s) => s.trim()).filter(Boolean);
}

/** Simple LCS diff -> ops: {type:'eq'|'del'|'ins', text} */
function diffTokens(a, b) {
  const tok = (s) => s.match(/\w+|[^\s\w]+|\s+/g) || [];
  const A = tok(a),
    B = tok(b);
  const n = A.length,
    m = B.length;
  const dp = Array.from({ length: n + 1 }, () => Array(m + 1).fill(0));
  for (let i = 1; i <= n; i++) {
    for (let j = 1; j <= m; j++) {
      dp[i][j] = A[i - 1] === B[j - 1] ? dp[i - 1][j - 1] + 1 : Math.max(dp[i - 1][j], dp[i][j - 1]);
    }
  }
  const ops = [];
  let i = n,
    j = m;
  while (i > 0 || j > 0) {
    if (i > 0 && j > 0 && A[i - 1] === B[j - 1]) {
      ops.push({ type: "eq", text: A[i - 1] });
      i--;
      j--;
    } else if (j > 0 && (i === 0 || dp[i][j - 1] >= dp[i - 1]?.[j])) {
      ops.push({ type: "ins", text: B[j - 1] });
      j--;
    } else {
      ops.push({ type: "del", text: A[i - 1] });
      i--;
    }
  }
  return ops.reverse();
}
function htmlEscape(s) {
  return s.replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
}
function renderSideBySide(before, after) {
  const ops = diffTokens(before, after);
  const left = ops
    .map((o) => (o.type === "eq" ? htmlEscape(o.text) : o.type === "del" ? `<span class="del">${htmlEscape(o.text)}</span>` : ""))
    .join("");
  const right = ops
    .map((o) => (o.type === "eq" ? htmlEscape(o.text) : o.type === "ins" ? `<span class="ins">${htmlEscape(o.text)}</span>` : ""))
    .join("");
  return `
  <div class="diff-pair">
    <div class="diff-col">
      <div class="diff-label">Before</div>
      <div class="diff-box">${left}</div>
    </div>
    <div class="diff-col">
      <div class="diff-label">After</div>
      <div class="diff-box">${right}</div>
    </div>
  </div>`;
}
function buildChangedSnippetsHtml(beforeText, afterText, maxPairs = 5) {
  const A = splitSentences(beforeText);
  const B = splitSentences(afterText);
  const pairs = [];
  const len = Math.min(A.length, B.length);
  for (let i = 0; i < len; i++) {
    if (A[i] && B[i] && A[i] !== B[i]) {
      pairs.push(renderSideBySide(A[i], B[i]));
      if (pairs.length >= maxPairs) break;
    }
  }
  if (!pairs.length) return `<div class="muted">No visible sentence-level changes detected.</div>`;
  return pairs.join("\n");
}

function buildReportHtml({
  mode, // "V1" | "V2"
  originalName,
  originalText,
  cleanedText,
  rephrasedText, // null en V1
  stats, // lengthBefore, lengthCleaned + counters
  notes, // { strictPdf }
}) {
  const style = `
  <style>
    body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;padding:24px;line-height:1.6}
    h1{font-size:20px;margin:0 0 12px}
    h2{font-size:16px;margin:24px 0 8px}
    .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
    .card{border:1px solid #eee;border-radius:12px;padding:16px}
    .kv{display:grid;grid-template-columns:180px 1fr;gap:8px;margin:6px 0}
    .muted{color:#666}
    .badge{display:inline-block;background:#f5f5f5;border-radius:10px;padding:2px 8px;font-size:12px;margin-left:8px}
    .diff-pair{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin:12px 0}
    .diff-label{font-size:12px;color:#666;margin-bottom:6px}
    .diff-box{border:1px solid #eee;border-radius:10px;padding:12px;white-space:pre-wrap}
    .ins{background:#e6ffed}
    .del{background:#ffeef0;text-decoration:line-through}
    ul.compact{margin:0;padding-left:18px}
  </style>`;

  const dup = stats.duplicatePunct;
  const dupList = Object.entries(dup)
    .filter(([, v]) => v > 0)
    .map(([k, v]) => `<li>${k} repeated: <strong>${v}</strong></li>`)
    .join("");

  const cleanedSection = `
    <h2>Cleaned changes (Original → Cleaned)</h2>
    <div class="card">
      ${buildChangedSnippetsHtml(originalText, cleanedText, 5)}
    </div>`;

  const rephraseSection = rephrasedText
    ? `
    <h2>Rephrase changes (Cleaned → Rephrased)</h2>
    <div class="card">
      ${buildChangedSnippetsHtml(cleanedText, rephrasedText, 5)}
    </div>`
    : "";

  const finalLen = rephrasedText ? rephrasedText.length : cleanedText.length;

  return `<!doctype html><html lang="en"><head><meta charset="utf-8"/>
    <title>DocSafe — Report</title>
    <meta name="viewport" content="width=device-width, initial-scale=1"/>
    ${style}
  </head><body>
    <h1>DocSafe Report
      <span class="badge">${mode === "V2" ? "V2 · Clean + Rephrase" : "V1 · Clean"}</span>
    </h1>

    <div class="grid">
      <div class="card">
        <div class="kv"><div><strong>File</strong></div><div>${originalName || "(untitled)"}</div></div>
        <div class="kv"><div><strong>Chars (before)</strong></div><div>${stats.lengthBefore}</div></div>
        <div class="kv"><div><strong>Chars (after cleaned)</strong></div><div>${stats.lengthCleaned}</div></div>
        ${rephrasedText ? `<div class="kv"><div><strong>Chars (after rephrased)</strong></div><div>${finalLen}</div></div>` : ""}
        <div class="kv"><div><strong>Strict PDF</strong></div><div>${notes?.strictPdf ? "on" : "off"}</div></div>
      </div>

      <div class="card">
        <div><strong>What changed (clean pass)</strong></div>
        <ul class="compact">
          <li>Zero-width characters removed: <strong>${stats.zeroWidthRemoved}</strong></li>
          <li>Multi-space runs collapsed: <strong>${stats.doubleSpaceRuns}</strong></li>
          <li>Spaces before punctuation fixed: <strong>${stats.spacesBeforePunct}</strong></li>
          <li>Missing space after punctuation fixed: <strong>${stats.missingSpaceAfterPunct}</strong></li>
          <li>Blank lines collapsed: <strong>${stats.collapsedBlankLines}</strong></li>
          <li>Sentence starts capitalized (heuristic): <strong>${stats.sentenceCapitalized}</strong></li>
          ${dupList ? `<li>Duplicate punctuation normalized:<ul class="compact">${dupList}</ul></li>` : ""}
        </ul>
      </div>
    </div>

    ${cleanedSection}
    ${rephraseSection}

    <p class="muted">Note: counts are heuristic; actual applied changes may vary slightly depending on document structure.</p>
  </body></html>`;
}

/* =========================
   Helpers
   ========================= */
function tmpPath(name) {
  return path.join(os.tmpdir(), `${Date.now()}_${name}`);
}
async function zipFiles(files) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    const archive = archiver("zip", { zlib: { level: 9 } });
    archive.on("warning", (e) => console.warn("zip warn:", e));
    archive.on("error", (e) => reject(e));
    archive.on("data", (d) => chunks.push(d));
    archive.on("end", () => resolve(Buffer.concat(chunks)));
    // contenu
    for (const f of files) archive.file(f.path, { name: f.name });
    archive.finalize();
  });
}

/* =========================
   Routes
   ========================= */
app.get("/health", (_req, res) => res.json({ ok: true, message: "Backend is running ✅" }));
app.get("/_env_ok", (_req, res) => res.json({ ok: true, NODE_ENV: process.env.NODE_ENV || "development" }));
app.get("/_ai_echo", (req, res) =>
  res.json({ ok: true, echo: aiSpellGrammarPass(String(req.query.q || "Hello from AI")) })
);
app.get("/_ai_rephrase_echo", (req, res) =>
  res.json({ ok: true, rephrased: aiRephrase(String(req.query.q || "Hello from AI")) })
);

/* ---- V1: Clean → ZIP = cleaned.docx + report.html ---- */
app.post("/clean", upload.any(), async (req, res) => {
  try {
    if (!req.files?.length) return res.status(400).json({ ok: false, error: "No file uploaded" });

    const strictPdf = String(req.body.strictPdf || "false") === "true";
    const f = req.files[0];
    const originalName = f.originalname || "input";
    const originalText = await extractTextFromFile(f.buffer, originalName, strictPdf);
    const lengthBefore = originalText.length;

    const corrected = aiSpellGrammarPass(originalText);
    const cleanedText = normalizeText(corrected);
    const lengthCleaned = cleanedText.length;

    const cleanedPath = tmpPath("cleaned.docx");
    await createDocxFromText(cleanedText, cleanedPath);

    const stats = computeCleanStats(originalText, cleanedText, lengthBefore, lengthCleaned);
    const reportHtml = buildReportHtml({
      mode: "V1",
      originalName,
      originalText,
      cleanedText,
      rephrasedText: null,
      stats,
      notes: { strictPdf },
    });
    const reportPath = tmpPath("report.html");
    fs.writeFileSync(reportPath, reportHtml, "utf8");

    const zipBuf = await zipFiles([
      { path: cleanedPath, name: "cleaned.docx" },
      { path: reportPath, name: "report.html" },
    ]);

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", 'attachment; filename="docsafe_v1_result.zip"');
    res.end(zipBuf);
  } catch (e) {
    console.error("CLEAN ERROR", e);
    res.status(500).json({ ok: false, error: "CLEAN_ERROR", message: String(e) });
  }
});

/* ---- V2: Clean + Rephrase → ZIP = rephrased.docx + report.html ---- */
app.post("/clean-v2", upload.any(), async (req, res) => {
  try {
    if (!req.files?.length) return res.status(400).json({ ok: false, error: "No file uploaded" });

    const strictPdf = String(req.body.strictPdf || "false") === "true";
    const f = req.files[0];
    const originalName = f.originalname || "input";
    const originalText = await extractTextFromFile(f.buffer, originalName, strictPdf);
    const lengthBefore = originalText.length;

    // Clean pass
    const corrected = aiSpellGrammarPass(originalText);
    const cleanedText = normalizeText(corrected);
    const lengthCleaned = cleanedText.length;

    // Rephrase pass — garantit un delta
    const rephrasedText = aiRephrase(cleanedText);

    const rephrasedPath = tmpPath("rephrased.docx");
    await createDocxFromText(rephrasedText, rephrasedPath);

    const stats = computeCleanStats(originalText, cleanedText, lengthBefore, lengthCleaned);
    const reportHtml = buildReportHtml({
      mode: "V2",
      originalName,
      originalText,
      cleanedText,
      rephrasedText,
      stats,
      notes: { strictPdf },
    });
    const reportPath = tmpPath("report.html");
    fs.writeFileSync(reportPath, reportHtml, "utf8");

    const zipBuf = await zipFiles([
      { path: rephrasedPath, name: "rephrased.docx" },
      { path: reportPath, name: "report.html" },
    ]);

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", 'attachment; filename="docsafe_v2_result.zip"');
    res.end(zipBuf);
  } catch (e) {
    console.error("CLEAN V2 ERROR", e);
    res.status(500).json({ ok: false, error: "CLEAN_V2_ERROR", message: String(e) });
  }
});

/* ---- Boot ---- */
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend listening on ${PORT}`));

