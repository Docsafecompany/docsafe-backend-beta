// lib/docxWriter.js
import { Document, Packer, Paragraph, HeadingLevel, TextRun } from "docx";

/**
 * Crée un .docx basique à partir d'un texte.
 * - title: utilisé en H1 et dans les métadonnées
 * - Retourne un Buffer prêt à être zippé / renvoyé.
 */
export async function createDocxFromText(text, opts = {}) {
  const safeText = typeof text === "string" ? text : String(text || "");
  const title = (opts.title || "DocSafe").toString();

  // 1) Construire les Paragraphs
  const children = [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text: title, bold: true })],
      spacing: { after: 200 },
    }),
  ];

  // Chaque ligne du texte devient un paragraphe
  for (const line of safeText.split(/\r?\n/)) {
    // Paragraphe vide si la ligne est vide pour conserver les retours
    children.push(
      new Paragraph({
        children: [new TextRun({ text: line })],
      })
    );
  }

  // 2) Document avec sections: []
  const doc = new Document({
    creator: "DocSafe",
    title,
    description: title,
    sections: [
      {
        properties: {},
        children, // <= un tableau de Paragraph
      },
    ],
  });

  // 3) Buffer
  const buf = await Packer.toBuffer(doc);
  return Buffer.from(buf);
}
