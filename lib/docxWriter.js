/* lib/docxWriter.js — CommonJS */
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun } = require("docx");

/**
 * DOCX depuis texte brut — corrige “sections is not iterable”
 * (sections: [{ properties: {}, children }])
 */
async function createDocxFromText(text, outPath) {
  const paragraphs = splitToParagraphs(text);

  const children = paragraphs.map((p) => {
    const runs = [];
    const parts = p.split("\n"); // soft line breaks
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
  const blocks = String(text)
    .split(/\n{2,}/g)
    .map((b) => b.replace(/\s+$/g, ""))
    .filter((b) => b.trim().length > 0);
  return blocks.length ? blocks : [String(text)];
}

module.exports = { createDocxFromText };

