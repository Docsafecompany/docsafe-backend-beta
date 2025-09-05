// lib/extract.js
import JSZip from 'jszip';
import mammoth from 'mammoth';
import pdfParse from 'pdf-parse';
import { normalizeText } from './textCleaner.js';

export async function extractFromDocx(buffer) {
  try {
    const { value } = await mammoth.extractRawText({ buffer });
    return normalizeText(value || '');
  } catch {
    return '';
  }
}

export async function extractFromPptx(buffer) {
  try {
    const zip = await JSZip.loadAsync(buffer);
    const slideRegex = /^ppt\/slides\/slide\d+\.xml$/;
    const parts = [];
    zip.forEach((p) => { if (slideRegex.test(p)) parts.push(p); });
    let text = '';
    for (const p of parts) {
      const xml = await zip.file(p).async('string');
      text += ' ' + xml.replace(/<a:t>([\s\S]*?)<\/a:t>/g, (_m, inner) =>
        inner
          .replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>')
          .replace(/&quot;/g,'"').replace(/&apos;/g,"'")
      ).replace(/<[^>]+>/g, ' ');
    }
    return normalizeText(text);
  } catch {
    return '';
  }
}

export async function extractFromPdf(buffer) {
  try {
    const data = await pdfParse(buffer);
    return normalizeText(data.text || '');
  } catch {
    return '';
  }
}
