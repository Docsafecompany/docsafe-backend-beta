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

const upload = multer({ storage: multer.memoryStorage() });

app.get('/health', (_req, res) => {
  res.json({ ok: true, message: 'Backend is running ✅' });
});

/**
 * mode:
 *  - 'v1' => nettoyage + orthographe IA (sans reformulation)
 *  - 'v2' => v1 + reformulation IA
 */
async function processFile({ buffer, filename, strictPdf = false, mode = 'v1' }) {
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
  const normalized = normalizeText(extractedText);

  // LanguageTool optionnel
  if (process.env.LT_ENDPOINT) {
    try { ltResult = await ltCheck(normalized); } catch { ltResult = null; }
  }

  // --- IA ---
  let proofText = null;
  let rephraseText = null;

  // V1: correction IA (= orthographe/espaces sans reformulation)
  try { proofText = await aiProofread(normalized); } catch { proofText = null; }

  // V2: reformulation en plus (sur le texte corrigé si dispo)
  if (mode === 'v2') {
    const baseForRephrase = proofText || normalized;
    try { rephraseText = await aiRephrase(baseForRephrase); } catch { rephraseText = null; }
  }

  // --- Sorties ZIP ---
  const files = [];

  // 1) V1: cleaned.docx depuis le texte corrigé par IA (si dispo)
  if (proofText) {
    const docx = await createDocxFromText(proofText, { title: 'DocSafe Cleaned (V1)' });
    files.push({ name: 'cleaned.docx', data: docx });
  } else {
    // fallback: renvoyer au moins le binaire nettoyé
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
      buffer: req.file.buffer,
      filename: req.file.originalname,
      strictPdf,
      mode: 'v1'
    });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="docsafe-v1.zip"');
    res.send(zipBuffer);
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

app.post('/clean-v2', upload.single('file'), async (req, res) => {
  try {
    const strictPdf = req.body?.strictPdf === 'true';
    const zipBuffer = await processFile({
      buffer: req.file.buffer,
      filename: req.file.originalname,
      strictPdf,
      mode: 'v2'
    });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="docsafe-v2.zip"');
    res.send(zipBuffer);
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

// --- Routes DIAG ---
app.get('/_ai_echo', async (_req, res) => {
  try {
    const sample = 'soc ial enablin g commu nication, dis   connection.';
    const proof = await aiProofread(sample);
    res.json({ ok: true, in: sample, proof });
  } catch (e) {
    res.status(500).json({ ok: false, where: '_ai_echo', error: String(e.message || e) });
  }
});

app.get('/_ai_rephrase_echo', async (_req, res) => {
  try {
    const sample = 'Notre solution réduit fortement les erreurs et améliore la qualité des documents. Elle s’intègre facilement aux outils existants.';
    const proof = await aiProofread(sample);
    const reph = await aiRephrase(proof || sample);
    res.json({ ok: true, in: sample, proof, rephrase: reph });
  } catch (e) {
    res.status(500).json({ ok: false, where: '_ai_rephrase_echo', error: String(e.message || e) });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend running on :${PORT}`));
