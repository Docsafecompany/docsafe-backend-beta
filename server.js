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

// CORS permissif (front Vercel → backend Render)
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*'); // si tu veux restreindre → mets l’URL Vercel
  res.header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
  next();
});
app.options('*', (req, res) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET,POST,OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
  return res.sendStatus(204);
});

app.use(compression());
app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ extended: true, limit: '20mb' }));

// Logs simples
app.use((req, _res, next) => { console.log(`[REQ] ${req.method} ${req.url}`); next(); });

const upload = multer({ storage: multer.memoryStorage() });

// Health
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

// Diag IA
app.get('/_ai_echo', async (_req, res) => {
  try {
    const sample = 'soc ial enablin g commu nication, dis   connection.';
    const proof = await aiProofread(sample);
    res.json({ ok: true, in: sample, proof });
  } catch (e) {
    res.json({ ok: false, where: '_ai_echo', error: e?.message || String(e) });
  }
});

app.get('/_ai_rephrase_echo', async (_req, res) => {
  try {
    const sample = 'Notre solution réduit fortement les erreurs et améliore la qualité des documents.';
    const proof = await aiProofread(sample);
    const reph = await aiRephrase(proof || sample);
    res.json({ ok: true, in: sample, proof, rephrase: reph });
  } catch (e) {
    res.json({ ok: false, where: '_ai_rephrase_echo', error: e?.message || String(e) });
  }
});

// Fonction principale
async function processFile({ buffer, filename, strictPdf = false, mode = 'v1' }) {
  if (!buffer) throw new Error('No file uploaded');

  const mime = await detectMime(buffer, filename);
  const ext = inferExt(filename, mime);

  let cleanedBuffer = buffer;
  if (ext === '.pdf') {
    const { outBuffer } = await cleanPDF(buffer, { strict: strictPdf, sections: [] });
    cleanedBuffer = outBuffer;
  } else if (ext === '.docx') {
    const { outBuffer } = await cleanDOCX(buffer, { sections: [] });
    cleanedBuffer = outBuffer;
  } else if (ext === '.pptx') {
    const { outBuffer } = await cleanPPTX(buffer, { sections: [] });
    cleanedBuffer = outBuffer;
  }

  let rawText = '';
  if (ext === '.pdf') rawText = await extractFromPdf(cleanedBuffer);
  if (ext === '.docx') rawText = await extractFromDocx(cleanedBuffer);
  if (ext === '.pptx') rawText = await extractFromPptx(cleanedBuffer);

  const normalized = normalizeText(rawText || '');

  let proofText = null;
  let rephraseText = null;

  if (normalized) {
    try { proofText = await aiProofread(normalized); } catch {}
    if (mode === 'v2') {
      try { rephraseText = await aiRephrase(proofText || normalized); } catch {}
    }
  }

  const files = [];
  files.push({ name: `cleaned-binary${ext}`, data: cleanedBuffer });

  const v1Docx = await createDocxFromText(proofText || normalized, { title: 'DocSafe Cleaned (V1)' });
  files.push({ name: 'cleaned.docx', data: v1Docx });

  if (mode === 'v2') {
    const v2Docx = await createDocxFromText(rephraseText || proofText || normalized, { title: 'DocSafe Rephrased (V2)' });
    files.push({ name: 'rephrased.docx', data: v2Docx });
  }

  const reportHtml = generateReportHTML({
    filename,
    mime,
    baseStats: { length: (normalized || '').length },
    ai: { applied: true, mode, proofed: !!proofText, rephrased: !!rephraseText }
  });
  files.push({ name: 'report.html', data: Buffer.from(reportHtml, 'utf-8') });

  return await zipOutput(files);
}

// Routes
app.post('/clean', upload.any(), async (req, res) => {
  try {
    const file = (req.files || [])[0];
    if (!file?.buffer) throw new Error('No file uploaded');
    const zipBuffer = await processFile({ buffer: file.buffer, filename: file.originalname, strictPdf: req.body?.strictPdf === 'true', mode: 'v1' });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="docsafe-v1.zip"');
    res.send(zipBuffer);
  } catch (e) {
    console.error('CLEAN ERROR', e);
    res.status(500).json({ ok: false, route: '/clean', error: e?.message || String(e) });
  }
});

app.post('/clean-v2', upload.any(), async (req, res) => {
  try {
    const file = (req.files || [])[0];
    if (!file?.buffer) throw new Error('No file uploaded');
    const zipBuffer = await processFile({ buffer: file.buffer, filename: file.originalname, strictPdf: req.body?.strictPdf === 'true', mode: 'v2' });
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
