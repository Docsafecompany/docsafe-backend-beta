// lib/officeXml.js
import JSZip from "jszip";

// Échappe le texte pour le réinjecter dans XML
function xmlEscape(s = "") {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

// Remplace le contenu d'un tag texte en conservant les attributs
function replaceTextNodes(xml, openCloseRegex, transformFn) {
  return xml.replace(openCloseRegex, (full, openTag, inner, closeTag) => {
    const newText = transformFn(inner);
    return `${openTag}${xmlEscape(newText)}${closeTag}`;
  });
}

// DOCX : on modifie word/document.xml, headers/footers, notes, comments
export async function processDocxBuffer(buffer, transformFn) {
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) =>
    /^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i.test(k)
  );

  for (const f of targets) {
    const xml = await zip.file(f).async("string");
    const updated = replaceTextNodes(xml, /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g, transformFn);
    zip.file(f, updated);
  }

  return await zip.generateAsync({ type: "nodebuffer" });
}

// PPTX : on modifie chaque slide + notes
export async function processPptxBuffer(buffer, transformFn) {
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) =>
    /^(ppt\/slides\/slide\d+\.xml|ppt\/notesSlides\/notesSlide\d+\.xml)$/i.test(k)
  );

  for (const f of targets) {
    const xml = await zip.file(f).async("string");
    const updated = replaceTextNodes(xml, /(<a:t[^>]*>)([\s\S]*?)(<\/a:t>)/g, transformFn);
    zip.file(f, updated);
  }

  return await zip.generateAsync({ type: "nodebuffer" });
}
