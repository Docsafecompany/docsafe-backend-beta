// lib/extract.js
import JSZip from 'jszip';
import mammoth from 'mammoth';
import { normalizeText } from './textCleaner.js';

// DOCX → texte brut
export async function extractFromDocx(buffer) {
  try {
    const { value } = await mammoth.extractRawText({ buffer });
    return normalizeText(value || '');
  } catch {
    return '';
  }
}

// PPTX → concat <a:t>
export async function extractFromPptx(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const slideRegex = /^ppt\/slides\/slide\d+\.xml$/;
    const slides = [];
    zip.forEach((p) => { if (slideRegex.test(p)) slides.push(p); });
    let text = '';
    for (const p of slides) {
      const xml = await zip.file(p).async('string');
      const parts = [];
      xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (_m, inner) => {
        parts.push(
          inner.replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
               .replace(/&quot;/g,'"').replace(/&apos;/g,"'")
        );
        return '';
      });
      text += ' ' + parts.join(' ');
    }
    return normalizeText(text);
  } catch {
    return '';
  }
}

// PDF → pdfjs-dist (Node)
export async function extractFromPdf(buffer) {
  try {
    const pdfjsLib = await import('pdfjs-dist/legacy/build/pdf.mjs');
    const loadingTask = pdfjsLib.getDocument({ data: buffer });
    const pdf = await loadingTask.promise;
    let text = '';
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      const content = await page.getTextContent();
      text += ' ' + content.items.map(it => it.str).join(' ');
    }
    return normalizeText(text);
  } catch {
    return '';
  }
}
