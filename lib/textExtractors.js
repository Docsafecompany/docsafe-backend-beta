// lib/textExtractors.js
// Extraction de texte basique pour DOCX/PPTX/PDF

import JSZip from "jszip";

// DOCX -> texte via XML
export async function extractTextFromDocx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const files = [
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/header1.xml",
    "word/header2.xml",
    "word/header3.xml",
  ];
  let out = "";
  for (const f of files) {
    const file = zip.file(f);
    if (!file) continue;
    const xml = await file.async("string");
    out += xmlToPlain(xml) + "\n";
  }
  return out.trim();
}

// PPTX -> texte depuis slides/*.xml
export async function extractTextFromPptx(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  const slideFiles = Object.keys(zip.files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k));
  let out = "";
  for (const f of slideFiles.sort()) {
    const xml = await zip.file(f).async("string");
    out += xmlToPlain(xml) + "\n\n";
  }
  return out.trim();
}

// PDF -> best-effort: on s'appuie sur pdf-parse si présent, sinon fallback vide
export async function extractTextFromPdf(buffer, { strictPdf = false } = {}) {
  try {
    // Chargement dynamique pour éviter d'alourdir si pas installé
    const pdfParse = (await import("pdf-parse")).default;
    const data = await pdfParse(buffer, { pagerender: undefined });
    // "strictPdf": l'extraction texte ignore déjà les calques invisibles dans la plupart des cas.
    // On peut faire une passe pour retirer des lignes ultra-courtes type filigranes si besoin:
    const lines = String(data.text || "")
      .replace(/\r\n/g, "\n")
      .split("\n");

    let filtered = lines;

    if (strictPdf) {
      // Heuristique: enlever lignes très courtes en majuscules isolées (suspect watermark/header)
      filtered = lines.filter((ln) => {
        const s = ln.trim();
        if (s.length <= 2) return false;
        if (/^[A-Z0-9 .\-]{1,10}$/.test(s)) return false;
        return true;
      });
    }

    return filtered.join("\n").trim();
  } catch (e) {
    // Fallback
    return "";
  }
}

// --- Helpers ---

function xmlToPlain(xml) {
  // Supprimer les tags XML et transformer <w:t> en texte avec espaces raisonnables
  // Extraction simple: remplace balises par espace, puis normalise.
  let text = String(xml || "")
    // Remplacer balises <w:tab/> par tabulation/espaces
    .replace(/<w:tab\/>/g, " ")
    // Remplacer <w:br/> par saut de ligne
    .replace(/<w:br\/>/g, "\n")
    // Extraire contenu de <a:t> (pptx) et <w:t> (docx)
    .replace(/<a:t[^>]*>([\s\S]*?)<\/a:t>/g, (_, g1) => g1)
    .replace(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g, (_, g1) => g1)
    // Supprimer les balises restantes
    .replace(/<[^>]+>/g, " ");

  // Normaliser espaces
  text = text.replace(/[ \t]{2,}/g, " ").replace(/\n{3,}/g, "\n\n").trim();
  return text;
}
