// lib/officeApplyFixes.js
import JSZip from "jszip";
import { normalizeFixes, applyFixesToText } from "./applySpellingFixes.js";

async function applyToXmlFiles(buffer, xmlPaths, fixes) {
  const zip = await JSZip.loadAsync(buffer);
  const normalized = normalizeFixes(fixes);

  let changedFiles = 0;
  let totalReplacements = 0;

  for (const p of xmlPaths) {
    const file = zip.file(p);
    if (!file) continue;

    const xml = await file.async("text");
    const { text: newXml, stats } = applyFixesToText(xml, normalized, 50);

    if (newXml !== xml) {
      zip.file(p, newXml);
      changedFiles++;
      totalReplacements += stats.replacements;
    }
  }

  const outBuffer = await zip.generateAsync({ type: "nodebuffer" });
  return {
    outBuffer,
    stats: {
      changedFiles,
      totalReplacements
    }
  };
}

export async function applyFixesDOCX(buffer, fixes) {
  // DOCX = texte principal + headers/footers (important)
  const xmlPaths = [
    "word/document.xml",
    ...Array.from({ length: 20 }, (_, i) => `word/header${i + 1}.xml`),
    ...Array.from({ length: 20 }, (_, i) => `word/footer${i + 1}.xml`)
  ];
  return applyToXmlFiles(buffer, xmlPaths, fixes);
}

export async function applyFixesPPTX(buffer, fixes) {
  // PPTX = slides + notes (important)
  // On parcourt dynamiquement
  const zip = await JSZip.loadAsync(buffer);
  const xmlPaths = Object.keys(zip.files).filter(
    p =>
      (p.startsWith("ppt/slides/slide") && p.endsWith(".xml")) ||
      (p.startsWith("ppt/notesSlides/notesSlide") && p.endsWith(".xml"))
  );

  // réutilise applyToXmlFiles mais on a déjà zip chargé → simplest: refaire via buffer
  return applyToXmlFiles(buffer, xmlPaths, fixes);
}

export async function applyFixesXLSX(buffer, fixes) {
  // XLSX: sharedStrings (la majorité des textes)
  const xmlPaths = ["xl/sharedStrings.xml"];
  return applyToXmlFiles(buffer, xmlPaths, fixes);
}
