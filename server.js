// server.js
import express from 'express';
import cors from 'cors';
import compression from 'compression';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';

import { detectMime, zipOutput, inferExt } from './lib/file.js';
import { normalizeText } from './lib/textCleaner.js';
import { generateReportHTML } from './lib/report.js';
import { ltCheck } from './lib/languagetool.js';
import { cleanPDF } from './lib/pdfCleaner.js';
import { cleanDOCX } from './lib/docxCleaner.js';
import { cleanPPTX } from './lib/pptxCleaner.js';
import { aiProofread, aiRephrase } from './lib/ai.js';
import { createDocxFromText } from './lib/docxWriter.js';
import { extractFromDocx, extractFromPdf, extractFromPptx } from './lib/extract.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(compression());
app.use(cors({ origin: process.env.CORS_ORIGIN || '*' }));

// ---- Logs de base
app.use((req, _res, next) => { console.log(`[REQ] ${req.method} ${req.url}`); next(); });
process.on('unhandledRejection', e => console.error('UNHANDLED REJECTION', e));
process.on('uncaughtException', e => console.error('UNCAUGHT EXCEPTION', e));

const upload = multer({ storage: multer.memoryStorage() });

// ---- Health & ENV
app.get('/health', (_req, res) => res.json({ ok: true, message: 'Backend is running ✅' }));
app.get('/_env_ok', (_req, res) => {
  const hasKey = typeof process.env.OPENAI_API_KEY === 'string' && process.env.OPENAI_API_KEY.length > 10;
  res.json({
    ok: true,
    AI_PROVIDER: process.env.AI_PROVIDER || null,
    AI_MODEL: process.env.AI_MODEL || null,
    OPENAI_API_KEY_present: hasKey
  });
});

// ---- DIAG (tolérants)
app.get('/_ai_echo', async (_req, res) => {
  try {
    const sample = 'soc ial enablin g commu nication, dis   connection.';
    const proof = await aiProofread(sample);
    if (!proof) return res.json({ ok: false, where: '_ai_echo', notice: 'AI unavailable (rate-limited or error)', proof: null });
    res.json({ ok: true, in: sample, proof });
  } catch (e) {
    res.json({ ok: false, where: '_ai_echo', error: e?.message || String(e) });
  }
});

app.get('/_ai_rephrase_echo', async (_req, res) => {
  try {
    const sample = 'Notre solution réduit fortement les erreurs et améliore la qualité des documents. Elle s’intègre facilement aux outils existants.';
    const proof = await aiProofread(sample);
    const base = proof || sample;
    const reph = await aiRephrase(base);
    if (!reph) return res.json({ ok: false, where: '_ai_rephrase_echo', notice: 'AI unavailable (rate-limited or error)', proof, rephrase: null });
    res.json({ ok: true, in: sample, proof, rephrase: reph });
  } catch (e) {
    res.json({ ok: false, where: '_ai_rephrase_echo', error: e?.message || String(e) });
  }
});

// ---- Pipeline principal
async function processFile({ buffer, filename, strictPdf = false, mode = 'v1' }) {
  if (!buffer || !filename) throw new Error('No file uploaded');

  const mime = await detectMime(buffer, filename);
  const ext = inferExt(filename, mime);

  // 1) Nettoyage binaire (PDF / DOCX / PPTX)
  let cleanedBuffer = buffer;
  if (ext === '.pdf') {
    const { outBuffer } = await cleanPDF(buffer, { strict: strictPdf });
    cleanedBuffer = outBuffer;
  } else if (ext === '.docx') {
    const { outBuffer } = await cleanDOCX(buffer);
    cleanedBuffer = outBuffer;
  } else if (ext === '.pptx') {
    const { outBuffer } = await cleanPPTX(buffer);
    cleanedBuffer = outBuffer;
  } else {
    throw new Error('Format non supporté. Utilisez PDF/DOCX/PPTX.');
  }

  // 2) Extraction robuste
  let rawText = '';
  if (ext === '.pdf') rawText = await extractFromPdf(cleanedBuffer);
  if (ext === '.docx') rawText = await extractFromDocx(cleanedBuffer);
  if (ext === '.pptx') rawText = await extractFromPptx(cleanedBuffer);

  // 3) Normalisation mécanique (toujours)
  const normalized = normalizeText(rawText || '');

  // 4) LanguageTool (facultatif pour rapport)
  let ltResult = null;
  if (process.env.LT_ENDPOINT && normalized) {
    try { ltResult = await ltCheck(normalized); } catch { ltResult = null; }
  }

  // 5) IA (avec fallbacks)
  let proofText = null;
  let rephraseText = null;

  if (normalized) {
    try { proofText = await aiProofread(normalized); } catch (e) { console.error('aiProofread error', e); }
    if (mode === 'v2') {
      const base = proofText || normalized;
      try { rephraseText = await aiRephrase(base); } catch (e) { console.error('aiRephrase error', e); }
    }
  }

  // 6) Sorties
  const files = [];

  // (a) Binaire nettoyé
  files.push({ name: `cleaned-binary${ext}`, data: cleanedBuffer });

  // (b) V1: cleaned.docx — toujours
  const v1Text = proofText || normalized;
  const v1Docx = await createDocxFromText(v1Text, { title: 'DocSafe Cleaned (V1)' });
  files.push({ name: 'cleaned.docx', data: v1Docx });

  // (c) V2: rephrased.docx — toujours si mode v2
  if (mode === 'v2') {
    const v2Text = rephraseText || proofText || normalized;
    const v2Docx = await createDocxFromText(v2Text, { title: 'DocSafe Rephrased (V2)' });
    files.push({ name: 'rephrased.docx', data: v2Docx });
  }

  // (d) Report
  const reportHtml = generateReportHTML({
    filename,
    mime,
    baseStats: { length: (normalized || '').length },
    lt: ltResult,
    ai: {
      applied: true,
      mode,
      proofed: !!proofText,
      rephrased: !!rephraseText
    }
  });
  files.push({ name: 'report.html', data: Buffer.from(reportHtml, 'utf-8') });

  return await zipOutput(files);
}

// ---- Routes upload (tolérantes): on accepte n'importe quel nom de champ
app.post('/clean', upload.any(), async (req, res) => {
  try {
    const strictPdf = req.body?.strictPdf === 'true';
    const first = (req.files || [])[0];
    if (!first?.buffer) throw new Error('No file uploaded (backend did not receive a file field)');
    const zipBuffer = await processFile({
      buffer: first.buffer,
      filename: first.originalname,
      strictPdf,
      mode: 'v1'
    });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="docsafe-v1.zip"');
    res.send(zipBuffer);
  } catch (e) {
    console.error('CLEAN ERROR', e, 'headers:', req.headers['content-type']);
    res.status(500).json({ ok: false, route: '/clean', error: e?.message || String(e) });
  }
});

app.post('/clean-v2', upload.any(), async (req, res) => {
  try {
    const strictPdf = req.body?.strictPdf === 'true';
    const first = (req.files || [])[0];
    if (!first?.buffer) throw new Error('No file uploaded (backend did not receive a file field)');
    const zipBuffer = await processFile({
      buffer: first.buffer,
      filename: first.originalname,
      strictPdf,
      mode: 'v2'
    });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="docsafe-v2.zip"');
    res.send(zipBuffer);
  } catch (e) {
    console.error('CLEAN-V2 ERROR', e, 'headers:', req.headers['content-type']);
    res.status(500).json({ ok: false, route: '/clean-v2', error: e?.message || String(e) });
  }
});

// ---- Listen
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend running on :${PORT}`));

