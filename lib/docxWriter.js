/* lib/docxWriter.js — CommonJS build */
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun } = require("docx");

/**
 * Build a DOCX from plain text.
 * - Fixes "sections is not iterable" by using: sections: [{ properties:{}, children }]
 * - Paragraph = block split by blank line; single \n kept as soft line breaks.
 */
async function createDocxFromText(text, outPath) {
  const paragraphs = splitToParagraphs(text);

  const children = paragraphs.map((p) => {
    const runs = [];
    const parts = p.split("\n");           // keep soft line breaks inside a paragraph
    parts.forEach((line, i) => {
      runs.push(new TextRun({ text: line }));
      if (i !== parts.length - 1) runs.push(new TextRun({ text: "\n" }));
    });
    return new Paragraph({ children: runs });
  });

  const doc = new Document({
    sections: [
      {
        properties: {},
        children,
      },
    ],
  });

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buffer);
}

function splitToParagraphs(text) {
  if (!text) return [""];
  // Split on double newlines; fallback to whole text if no blocks found
  const blocks = text
    .split(/\n{2,}/g)
    .map((b) => b.replace(/\s+$/g, ""))
    .filter((b) => b.trim().length > 0);
  return blocks.length ? blocks : [text];
}

module.exports = { createDocxFromText };

