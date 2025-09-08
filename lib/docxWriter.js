// lib/docxWriter.js
import { Document, Packer, Paragraph, HeadingLevel, TextRun } from "docx";

function toParagraphs(text) {
  const blocks = String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split(/\n{2,}/g)
    .map(s => s.trim())
    .filter(Boolean);

  if (!blocks.length) return [new Paragraph({ children: [new TextRun("")] })];

  const paras = [];
  for (const block of blocks) {
    const lines = block.split(/\n/);
    for (const ln of lines) {
      paras.push(
        new Paragraph({
          spacing: { after: 200 },
          children: [new TextRun({ text: ln, size: 24 })] // 12pt
        })
      );
    }
  }
  return paras;
}

export async function createDocxFromText(text, title = "Document") {
  const children = [];

  if (title) {
    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { after: 300 },
        children: [new TextRun({ text: title, bold: true })]
      })
    );
  }

  children.push(...toParagraphs(text));

  const doc = new Document({
    sections: [
      {
        properties: {},
        children
      }
    ]
  });

  return await Packer.toBuffer(doc);
}
