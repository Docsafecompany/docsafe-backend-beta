// lib/extract.js
import JSZip from 'jszip';
import mammoth from 'mammoth';
import { normalizeText } from './textCleaner.js';

// --- DOCX via mammoth ---
export async function extractFromDocx(buffer) {
  try {
    const { value } = await mammoth.extractRawText({ buffer });
    return normalizeText(value || '');
  } catch {
    return '';
  }
}

// --- PPTX en lisant le XML (a:t) ---
export async function extractFromPptx(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const slideRegex = /^ppt\/slides\/slide\d+\.xml$/;
    const parts = [];
    zip.forEach((p) => { if (slideRegex.test(p)) parts.push(p); });
    let text = '';
    for (const p of parts) {
      const xml = await zip.file(p).async('string');
      const pieces = [];
      xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (_m, inner) => {
        pieces.push(
          inner
            .replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
            .replace(/&quot;/g,'"').replace(/&apos;/g,"'")
        );
        return '';
      });
      text += ' ' + pieces.join(' ');
    }
    return normalizeText(text);
  } catch {
    return '';
  }
}

// --- PDF via pdfjs-dist (sans worker en Node) ---
export async function extractFromPdf(buffer) {
  try {
    const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.mjs');
    // En Node, pas de worker requis
    // pdfjsLib.GlobalWorkerOptions.workerSrc = false;

    const loadingTask = pdfjsLib.getDocument({ data: buffer });
    const pdf = await loadingTask.promise;

    let text = '';
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      const pageText = content.items.map((it) => it.str).join(' ');
      text += ' ' + pageText;
    }
    return normalizeText(text);
  } catch {
    return '';
  }
}

