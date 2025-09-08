// lib/docxWriter.js
// Génère un DOCX propre à partir d'un texte normalisé.
// IMPORTANT: utilise le bon schéma: sections: [ { properties: {}, children } ]

import { Document, Packer, Paragraph, HeadingLevel, TextRun } from "docx";

/**
 * Découpe le texte en paragraphes (séparés par lignes vides ou \n).
 */
function toParagraphs(text) {
  const blocks = String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split(/\n{2,}/g) // paragraphes séparés par lignes vides
    .map((s) => s.trim())
    .filter(Boolean);

  if (blocks.length === 0) return [new Paragraph("")];

  const paras = [];
  for (const block of blocks) {
    // Recouper des lignes trop longues proprement
    const lines = block.split(/\n/);
    for (const ln of lines) {
      paras.push(
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({ text: ln, size: 24 })], // 12pt
        })
      );
    }
  }
  return paras;
}

export async function createDocxFromText(text, title = "Document") {
  const children = [];

  // Titre (facultatif)
  if (title) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 300 },
        children: [new TextRun({ text: title, bold: true })],
      })
    );
  }

  // Corps
  children.push(...toParagraphs(text));

  const doc = new Document({
    sections: [
      {
        properties: {},
        children,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  return buffer;
}


