// server.js
import express from 'express';
import cors from 'cors';
import compression from 'compression';
import multer from 'multer';
import path from 'path';
import { fileURLToPath } from 'url';
import { v4 as uuidv4 } from 'uuid';

import { detectMime, zipOutput, inferExt } from './lib/file.js';
import { normalizeText } from './lib/textCleaner.js';
import { generateReportHTML } from './lib/report.js';
import { ltCheck } from './lib/languagetool.js';
import { cleanPDF } from './lib/pdfCleaner.js';
import { cleanDOCX } from './lib/docxCleaner.js';
import { cleanPPTX } from './lib/pptxCleaner.js';
import { aiRephrase } from './lib/ai.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(compression());
app.use(cors({ origin: process.env.CORS_ORIGIN || '*' }));

const upload = multer({ storage: multer.memoryStorage() });

app.get('/health', (_req, res) => {
  res.json({ ok: true, message: 'Backend is running ✅' });
});

// ----- core pipeline -----
async function processFile({ buffer, filename, strictPdf = false, doAI = false }) {
  const id = uuidv4();
  const mime = await detectMime(buffer, filename);
  const ext = inferExt(filename, mime);

  let cleanedBuffer = buffer;
  let extractedText = '';
  let ltResult = null;

  if (mime === 'application/pdf' || ext === '.pdf') {
    const { outBuffer, text } = await cleanPDF(buffer, { strict: strictPdf });
    cleanedBuffer = outBuffer; extractedText = text || '';
  } else if (
    mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
    ext === '.docx'
  ) {
    const { outBuffer, text } = await cleanDOCX(buffer);
    cleanedBuffer = outBuffer; extractedText = text || '';
  } else if (
    mime === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' ||
    ext === '.pptx'
  ) {
    const { outBuffer, text } = await cleanPPTX(buffer);
    cleanedBuffer = outBuffer; extractedText = text || '';
  } else {
    throw new Error('Format non supporté. Utilisez PDF/DOCX/PPTX.');
  }

  const normalized = normalizeText(extractedText);

  if (process.env.LT_ENDPOINT) {
    try { ltResult = await ltCheck(normalized); } catch { ltResult = null; }
  }

  let aiText = null;
  if (doAI) {
    try { aiText = await aiRephrase(normalized); } catch { aiText = null; }
  }

  const reportHtml = generateReportHTML({
    filename,
    mime,
    baseStats: { length: (normalized || '').length },
    lt: ltResult,
    ai: aiText ? { applied: true } : { applied: false },
  });

  const files = [
    { name: `cleaned${ext}`, data: cleanedBuffer },
    { name: 'report.html', data: Buffer.from(reportHtml, 'utf-8') },
  ];
  const zipBuffer = await zipOutput(files);
  return zipBuffer;
}

// ----- routes -----
app.post('/clean', upload.single('file'), async (req, res) => {
  try {
    const strictPdf = req.body?.strictPdf === 'true';
    const zipBuffer = await processFile({
      buffer: req.file.buffer,
      filename: req.file.originalname,
      strictPdf,
      doAI: false
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
      doAI: true
    });
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', 'attachment; filename="docsafe-v2.zip"');
    res.send(zipBuffer);
  } catch (e) {
    console.error(e);
    res.status(500).json({ ok: false, error: e.message });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend running on :${PORT}`));
