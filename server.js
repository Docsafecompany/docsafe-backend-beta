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

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(compression());
app.use(cors({ origin: process.env.CORS_ORIGIN || '*' }));

// ===== LOGGING & DIAG =====
app.use((req, _res, next) => {
  console.log(`[REQ] ${req.method} ${req.url}`);
  next();
});
process.on('unhandledRejection', (e) => {
  console.error('UNHANDLED REJECTION', e);
});
process.on('uncaughtException', (e) => {
  console.error('UNCAUGHT EXCEPTION', e);
});

const upload = multer({ storage: multer.memoryStorage() });

// Health
app.get('/health', (_req, res) => {
  res.json({ ok: true, message: 'Backend is running ✅' });
});

// ENV diag (ne retourne PAS la clé)
app.get('/_env_ok', (_req, res) => {
  const hasKey = typeof process.env.OPENAI_API_KEY === 'string' && process.env.OPENAI_API_KEY.length > 10;
  res.json({
    ok: true,
    AI_PROVIDER: process.env.AI_PROVIDER || null,
    AI_MODEL: process.env.AI_MODEL || null,
    OPENAI_API_KEY_present: hasKey
  });
});

// IA diagnostics
app.get('/_ai_echo', async (_req, res) => {
  try {
    const sample = 'soc ial enablin g commu nication, dis   connection.';
    const proof = await aiProofread(sample);
    res.json({ ok: true, in: sample, proof });
  } catch (e) {
    console.error('AI_ECHO ERROR', e);
    res.status(500).json({ ok: false, where: '_ai_echo', error: e?.message || String(e) });
  }
});

app.get('/_ai_rephrase_echo', async (_req, res) => {
  try {
    const sample = 'Notre solution réduit fortement les erreurs et améliore la qualité des documents. Elle s’intègre facilement aux outils existants.';
    const proof = await aiProofread(sample);
    const reph = await aiRephrase(proof || sample);
    res.json({ ok: true, in: sample, proof, rephrase: reph });
  } catch (e) {
    console.error('AI_REPHRASE_ECHO ERROR', e);
    res.status(500).json({ ok: false, where: '_ai_rephrase_echo', error: e?.message || String(e) });
  }
});

/**
 * mode:
 *  - 'v1' => nettoyage + orthographe IA (sans reformulation)
 *  - 'v2' => v1 + reformulation IA
 */
async function processFile({ buffer, filename, strictPdf = false, mode = 'v1' }) {
  if (!buffer || !filename) throw new Error('No file uploaded');

  const mime = await detectMime(buffer, filename);
  const ext = inferExt(filename, mime);

  let cleanedBuffer = buffer;
  let extractedText = '';
  let ltResult = null;

  if (mime === 'application/pdf' || ext === '.pdf') {
    const { outBuffer, text } = await cleanPDF(buffer, { strict: strictPdf });
    cleanedBuffer = outBuffer; extractedText = text || '';
  } else if (mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || ext === '.docx') {
    const { outBuffer, text } = await cleanDOCX(buffer);
    cleanedBuffer = outBuffer; extractedText = text || '';
  } else if (mime === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' || ext === '.pptx') {
    const { outBuffer, text } = await cleanPPTX(buffer);
    cleanedBuffer = outBuffer; extractedText = text || '';
  } else {
    throw new Error('Format non supporté. Utilisez PDF/DOCX/PPTX.');
  }

  // Normalisation mécanique
  const normalized = normalizeText(extractedText || '');

  // LanguageTool optionnel
  if (process.env.LT_ENDPOINT) {
    try { ltResult = await ltCheck(normalized); } catch (e) { console.warn('LT error', e?.message || e); ltResult = null; }
  }

  // --- IA ---
  let proofText = null;
  let rephraseText = null;

  // V1: correction IA (= orthographe/espaces sans reformulation)
  try {
    proofText = await aiProofread(normalized);
  } catch (e) {
    console.error('aiProofread error', e);
    proofText = null;
  }

  // Si tu veux “fail-fast”, décommente :
  // if (!proofText) throw new Error('AI proofread unavailable (check OPENAI_API_KEY/provider).');

  // V2: reformulation (sur le texte corrigé si dispo)
  if (mode === 'v2') {
    const baseForRephrase = proofText || normalized;
    try { rephraseText = await aiRephrase(baseForRephrase); }
    catch (e) { console.error('aiRephrase error', e); rephraseText = null; }
  }

  // --- Sorties ZIP ---
  const files = [];

  // 1) V1: cleaned.docx depuis le texte corrigé par IA (si dispo), sinon binaire nettoyé
  if (proofText) {
    const docx = await createDocxFromText(proofText, { title: 'DocSafe Cleaned (V1)' });
    files.push({ name: 'cleaned.docx', data: docx });
  } else {
    files.push({ name: `cleaned${ext}`, data: cleanedBuffer });
  }

  // 2) V2: rephrased.docx (toujours généré en mode v2)
  if (mode === 'v2') {
    const base = rephraseText || proofText || normalized || '';
    const reDocx = await createDocxFromText(base, { title: 'DocSafe Rephrased (V2)' });
    files.push({ name: 'rephrased.docx', data: reDocx });
  }

  // 3) Report
  const reportHtml = generateReportHTML({
    filename,
    mime,
    baseStats: { length: (normalized || '').length },
    lt: ltResult,
    ai: { applied: true, mode }
  });
  files.push({ name: 'report.html', data: Buffer.from(reportHtml, 'utf-8') });

  const zipBuffer = await zipOutput(files);
  return zipBuffer;
}

// --- Routes principales ---
app.post('/clean', upload.single('file'), async (req, res) => {
  try {
    const strictPdf = req.body?.strictPdf === 'true';
    const zipBuffer = await processFile({
      buffer: req.file?.buffer,
      filename: req.file?.originalname,
      strictPdf,
      mode: 'v1'
    });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="docsafe-v1.zip"');
    res.send(zipBuffer);
  } catch (e) {
    console.error('CLEAN ERROR', e);
    res.status(500).json({ ok: false, route: '/clean', error: e?.message || String(e) });
  }
});

app.post('/clean-v2', upload.single('file'), async (req, res) => {
  try {
    const strictPdf = req.body?.strictPdf === 'true';
    const zipBuffer = await processFile({
      buffer: req.file?.buffer,
      filename: req.file?.originalname,
      strictPdf,
      mode: 'v2'
    });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="docsafe-v2.zip"');
    res.send(zipBuffer);
  } catch (e) {
    console.error('CLEAN-V2 ERROR', e);
    res.status(500).json({ ok: false, route: '/clean-v2', error: e?.message || String(e) });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend running on :${PORT}`));

