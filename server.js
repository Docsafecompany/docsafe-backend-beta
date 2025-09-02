/**
 * DocSafe Backend — Beta V2 (Cleaner amélioré)
 *
 * Nouveautés:
 * - cleanTextBasic: supprime doublons de ponctuation (",,", ";;", "!!", "??", "..."), espaces superflus.
 * - PDF V1: purge métadonnées (Info + dates), enlève JS/Annots/Attachments/AcroForm/OpenAction (durcit la surface d'attaque).
 * - Mode strict PDF (optionnel): tentative de retrait de texte blanc/invisible dans les flux (regex best-effort).
 *   Activez-le avec:  1) env PDF_STRICT=1  OU  2) query ?strict=1 sur /clean (PDF uniquement)
 *
 * Endpoints:
 *  - POST /clean     -> PDF/DOCX nettoyé
 *  - POST /clean-v2  -> nettoyé + rapport LanguageTool (ZIP: cleaned + report.json + report.html)
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
import pako from "pako";

const app = express();

/* ---------------- CORS (ouvert pour valider rapidement) ---------------- */
app.use(cors({ origin: true, credentials: true }));
app.options("*", cors());

app.use(express.json({ limit: "25mb" }));

/* ---------------- Upload ---------------- */
const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, os.tmpdir()),
    filename: (req, file, cb) =>
      cb(null, Date.now() + "-" + file.originalname.replace(/\s+/g, "_")),
  }),
  limits: { fileSize: 20 * 1024 * 1024 }, // 20 Mo
});

/* ---------------- Helpers ---------------- */
const LT_URL = process.env.LT_API_URL || "https://api.languagetool.org/v2/check";
const LT_KEY = process.env.LT_API_KEY || null;

function cleanTextBasic(str) {
  if (!str) return str;
  let s = String(str);

  // Normalisations simples
  s = s.replace(/\u200B/g, "");                   // zero-width space
  s = s.replace(/[ \t]+/g, " ");                  // espaces multiples
  s = s.replace(/ *\n */g, "\n");                 // espaces autour des retours à la ligne

  // Ponctuation: collapse doublons
  s = s.replace(/([,;:!?])\1+/g, "$1");           // ",,", "!!", "??", ";;", "::"
  s = s.replace(/\.{3,}/g, "...");                // "....." -> "..."
  // Espaces autour de la ponctuation
  s = s.replace(/ ?([,;:!?]) ?/g, "$1 ");         // " , " -> ", "
  s = s.replace(/ \./g, ".");                     // " ." -> "."
  s = s.replace(/[ ]{2,}/g, " ");                 // espaces doublons

  // Exemple FR: espace fines avant ;:!? non gérées ici -> on uniformise simple
  return s.trim();
}

/* ---------------- DOCX ---------------- */
async function processDOCXBasic(buf) {
  const zip = await JSZip.loadAsync(buf);
  // Métadonnées
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

  // Nettoyage du texte
  const docXml = "word/document.xml";
  if (zip.file(docXml)) {
    let xml = await zip.file(docXml).async("string");
    xml = xml.replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (m, inner) => {
      const cleaned = cleanTextBasic(inner);
      return m.replace(inner, cleaned);
    });
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

/* ---------------- PDF ---------------- */
function _nukeCommonCatalogKeys(pdf) {
  try {
    const dict = pdf?.catalog?.dict;
    if (!dict) return;

    // Supprimer clés qui exposent des actions/JS/attachements/formulaires
    const kill = [
      "OpenAction", "AA", "AcroForm", "Lang", "Names", "Outlines",
      "PageMode", "ViewerPreferences", "URI", "JavaScript", "JS",
      "Collection", "Metadata" // XMP metadata
    ];
    // suppression souple (sans PDFName)
    for (const k of Array.from(dict.keys?.() || [])) {
      const keyStr = k?.toString?.() || "";
      if (kill.some(tag => keyStr.includes(`/${tag}`))) {
        dict.delete(k);
      }
    }
  } catch {}
}

function _stripAnnotsAndActions(pdf) {
  try {
    for (const page of pdf.getPages()) {
      // Enlever annotations (liens, commentaires, etc.)
      // Accès souple au dictionnaire interne (pas d'API publique stricte)
      const node = page.node;
      const keys = Array.from(node.dict?.keys?.() || []);
      for (const k of keys) {
        const ks = k?.toString?.() || "";
        if (ks.includes("/Annots")) node.dict.delete(k);
        if (ks.includes("/AA")) node.dict.delete(k);          // actions supplémentaires
      }
    }
  } catch {}
}

// Best-effort: supprime texte blanc/invisible de certains flux
function _tryStripWhiteOrInvisible(streamBytes) {
  // Détection/édition via regex sur contenu texte (si FlateDecode -> décodé avant)
  const str = new TextDecoder("latin1").decode(streamBytes);

  // Retire séquences: "1 1 1 rg ... (text) Tj" ou "[...] TJ"
  const rgWhite = /(?:^|\s)(?:1(?:\.0+)?\s+){2}1(?:\.0+)?\s+rg[\s\S]{0,400}?(?:\([^\)]*\)\s*Tj|\[[^\]]*\]\s*TJ)/g;
  // Retire texte en mode rendu invisible "3 Tr ... (text) Tj/TJ"
  const trInvisible = /(?:^|\s)3\s+Tr[\s\S]{0,400}?(?:\([^\)]*\)\s*Tj|\[[^\]]*\]\s*TJ)/g;

  let out = str.replace(rgWhite, "");
  out = out.replace(trInvisible, "");

  return new TextEncoder().encode(out);
}

async function processPDFBasic(buf, { strict = false } = {}) {
  const pdf = await PDFDocument.load(buf, { updateMetadata: true });

  // 1) Métadonnées "Info"
  pdf.setTitle(""); pdf.setAuthor(""); pdf.setSubject("");
  pdf.setKeywords([]); pdf.setProducer(""); pdf.setCreator("");
  const epoch = new Date(0);
  pdf.setCreationDate(epoch); pdf.setModificationDate(epoch);

  // 2) Purge durcie du catalogue (XMP metadata, JS, actions, attachements, formulaires…)
  _nukeCommonCatalogKeys(pdf);

  // 3) Pages: enlever annotations & actions
  _stripAnnotsAndActions(pdf);

  // 4) (Optionnel) STRIP flux "texte blanc / invisible"
  if (strict) {
    try {
      const context = pdf.context;
      for (const page of pdf.getPages()) {
        // Récupère la réf "Contents" (un stream ou un array de streams)
        const node = page.node;
        const contentsRef = node.dict.get?.(Object.fromEntries([])); // placeholder to keep TS calm
        // Accès plus permissif:
        const dict = node.dict;
        const contentsKey = Array.from(dict.keys?.() || []).find(k => (k?.toString?.() || "").includes("/Contents"));
        if (!contentsKey) continue;
        const raw = dict.get(contentsKey);

        const asArray = Array.isArray(raw?.array) ? raw.array : (raw ? [raw] : []);
        const newStreams = [];

        for (const ref of asArray) {
          const stream = context.lookup(ref); // PDFRawStream or similar
          if (!stream) continue;
          let bytes = stream.contents || stream.getContents?.();
          if (!bytes) continue;

          // Décode si FlateDecode
          const dictS = stream.dict;
          const keys = Array.from(dictS.keys?.() || []);
          const hasFlate = keys.some(k => (k?.toString?.() || "").includes("/Filter")) &&
            String(dictS.get?.(keys.find(k => (k?.toString?.() || "").includes("/Filter")))).includes("Flate");

          let decoded = bytes;
          if (hasFlate) {
            try { decoded = pako.inflate(bytes); } catch {}
          }

          // STRIP
          const cleaned = _tryStripWhiteOrInvisible(decoded);

          // Re-encode si nécessaire
          let finalBytes = cleaned;
          if (hasFlate) {
            try { finalBytes = pako.deflate(cleaned); } catch {}
          }

          // Remplace le contenu et met à jour /Length
          stream.contents = finalBytes;
          const lenKey = Array.from(dictS.keys?.() || []).find(k => (k?.toString?.() || "").includes("/Length"));
          if (lenKey) dictS.set?.(lenKey, context.obj(finalBytes.length));
          newStreams.push(stream);
        }
      }
    } catch {
      // en cas d'échec silencieux, on laisse le PDF tel quel (au pire: seules les métadonnées/annots/JS sont purgées)
    }
  }

  const out = await pdf.save({ useObjectStreams: false });
  return Buffer.from(out);
}

/* ---------------- LT ---------------- */
async function extractTextFromPDF(_buf) {
  // Dans cette bêta, on ne lit pas le texte PDF (pour éviter les crashs sur certains fichiers).
  // => Rapport LT vide pour PDF, mais cleaning OK.
  return "";
}

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
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      timeout: 30000,
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

/* ---------------- Routes ---------------- */
app.get("/health", (req, res) => res.json({ ok: true }));

app.post("/clean", upload.single("file"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "Missing file" });
  const p = req.file.path, name = req.file.originalname.toLowerCase();
  const strict = !!(process.env.PDF_STRICT === "1" || req.query.strict === "1");
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
    } else {
      return res.status(400).json({ error: "Only PDF or DOCX supported" });
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
  const strict = !!(process.env.PDF_STRICT === "1" || req.query.strict === "1");
  try {
    let cleanedBuf, cleanedName, text = "";
    if (name.endsWith(".pdf")) {
      const buf = await fsp.readFile(p);
      cleanedBuf = await processPDFBasic(buf, { strict });
      cleanedName = req.file.originalname.replace(/\.pdf$/i, "") + "_cleaned.pdf";
      text = await extractTextFromPDF(cleanedBuf); // vide (choix de bêta)
    } else if (name.endsWith(".docx")) {
      const buf = await fsp.readFile(p);
      cleanedBuf = await processDOCXBasic(buf);
      cleanedName = req.file.originalname.replace(/\.docx$/i, "") + "_cleaned.docx";
      text = await extractTextFromDOCX(cleanedBuf);
    } else {
      return res.status(400).json({ error: "Only PDF or DOCX supported" });
    }

    const matches = await runLanguageTool(text || "", lang);
    const reportJSON = buildReportJSON({ fileName: cleanedName, language: lang, matches, textLength: (text || "").length });
    const reportHTML = buildReportHTML(reportJSON);

    const zip = new JSZip();
    zip.file(cleanedName, cleanedBuf);
    zip.file("report.json", JSON.stringify(reportJSON, null, 2));
    zip.file("report.html", reportHTML);

    const zipBuf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
    const zipName = req.file.originalname.replace(/\.(pdf|docx)$/i, "") + "_docsafe_report.zip";
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

