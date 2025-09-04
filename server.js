import express from 'express';
import cors from 'cors';
import compression from 'compression';
import multer from 'multer';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { v4 as uuidv4 } from 'uuid';
import { detectMime, zipOutput, readFileBuffer, writeFileBuffer, inferExt } from './lib/file.js';
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


app.get('/health', (req, res) => {
res.json({ ok: true, message: 'Backend is running ✅' });
});


/**
* Core pipeline (V1 & V2)
* - 1) Détection format
* - 2) Extraction texte + nettoyage base (espaces/ponctuation) + métadonnées
* - 3) LanguageTool (si configuré)
* - 4) (V2) Reformulation IA
* - 5) Génération report.html + ZIP (cleaned.ext + report.html)
*/
async function processFile({ buffer, filename, strictPdf = false, doAI = false }) {
const id = uuidv4();
const mime = await detectMime(buffer, filename);
const ext = inferExt(filename, mime);


let cleanedBuffer = buffer; // par défaut
let extractedText = '';
let ltResult = null;


if (mime === 'application/pdf' || ext === '.pdf') {
const { outBuffer, text } = await cleanPDF(buffer, { strict: strictPdf });
cleanedBuffer = outBuffer; extractedText = text;
} else if (mime === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' || ext === '.docx') {
const { outBuffer, text } = await cleanDOCX(buffer);
cleanedBuffer = outBuffer; extractedText = text;
} else if (mime === 'application/vnd.openxmlformats-officedocument.presentationml.presentation' || ext === '.pptx') {
const { outBuffer, text } = await cleanPPTX(buffer);
app.listen(PORT, () => console.log(`DocSafe backend running on :${PORT}`));
