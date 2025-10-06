// server.js
import express from "express";
import cors from "cors";
import multer from "multer";
import AdmZip from "adm-zip";
import path from "path";
import { fileURLToPath } from "url";

import {
  processDocxBuffer,
  processPptxBuffer,
  analyzeDocxBuffer, // pour /_docx_probe2
} from "./lib/officeXml.js";
import {
  stripPdfMetadata,
  extractPdfText,
  filterExtractedLines,
} from "./lib/pdfTools.js";
import { buildReportHtml } from "./lib/report.js";
import { createDocxFromText } from "./lib/docxWriter.js";
import { aiCorrectText, aiRephraseText } from "./lib/ai.js";

// ✅ NEW: cleaners
import { cleanDOCX } from "./lib/docxCleaner.js";
import { cleanPPTX } from "./lib/pptxCleaner.js";
import { cleanPDF  } from "./lib/pdfCleaner.js";

// -----------------------------------------------------------------------------
// ES module helpers
// -----------------------------------------------------------------------------
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// -----------------------------------------------------------------------------
// App & middleware
// -----------------------------------------------------------------------------
const app = express();
app.use(cors({ origin: "*", credentials: false }));
app.use(express.json({ limit: "32mb" }));
app.use(express.urlencoded({ extended: true, limit: "32mb" }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }, // 50 MB
});

// -----------------------------------------------------------------------------
// Helpers
// -----------------------------------------------------------------------------
const getExt = (fn = "") =>
  fn.includes(".") ? fn.split(".").pop().toLowerCase() : "";
const outName = (single, base, name) => (single ? name : `${base}_${name}`);

// ✅ NEW: sanitation commun
const sanitizeBuffer = async (originalName, buf, opts = {}) => {
  const ext = getExt(originalName);
  if (ext === "docx") {
    const { outBuffer } = await cleanDOCX(buf);
    return outBuffer;
  }
  if (ext === "pptx") {
    const { outBuffer } = await cleanPPTX(buf);
    return outBuffer;
  }
  if (ext === "pdf") {
    const { outBuffer } = await cleanPDF(buf, { strict: Boolean(opts.strictPdf) });
    return outBuffer;
  }
  return buf; // autres formats inchangés
};

// -----------------------------------------------------------------------------
// Health / Echo
// -----------------------------------------------------------------------------
app.get("/health", (_, res) =>
  res.json({ ok: true, service: "DocSafe Backend", time: new Date().toISOString() })
);

app.get("/_env_ok", (_, res) =>
  res.json({ OPENAI_API_KEY: process.env.OPENAI_API_KEY ? "present" : "absent" })
);

app.get("/_ai_echo", async (req, res) => {
  const sample =
    req.query.q || "This is   a  test,, please  fix punctuation!! Thanks";
  try {
    const out = (await aiCorrectText(String(sample))) || String(sample);
    res.json({ input: sample, corrected: out });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.get("/_ai_rephrase_echo", async (req, res) => {
  const sample =
    req.query.q || "Overall, this is a sentence that could be rephrased furthermore.";
  try {
    const out = (await aiRephraseText(String(sample))) || String(sample);
    res.json({ input: sample, rephrased: out });
  } catch (e) {
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// -----------------------------------------------------------------------------
// V1: /clean  (correction stricte + report)
// -----------------------------------------------------------------------------
app.post("/clean", upload.any(), async (req, res) => {
  try {
    const strictPdf = String(req.body.strictPdf || "false") === "true";
    const files = req.files || [];
    if (!files.length) return res.status(400).json({ error: "No files uploaded." });

    const single = files.length === 1;
    const zip = new AdmZip();

    for (const f of files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      if (ext === "docx") {
        // ✅ sanitize
        const sanitized = await sanitizeBuffer(f.originalname, f.buffer, { strictPdf });
        const cleanedBuf = await processDocxBuffer(sanitized, "V1");

        zip.addFile(outName(single, base, "cleaned.docx"), cleanedBuf);
        zip.addFile(
          outName(single, base, "report.html"),
          Buffer.from(
            buildReportHtml({
              original: "[structured DOCX]",
              cleaned: "[DOCX corrected by AI]",
              rephrased: null,
              filename: f.originalname,
              mode: "V1",
            }),
            "utf8"
          )
        );
      } else if (ext === "pptx") {
        // ✅ sanitize
        const sanitized = await sanitizeBuffer(f.originalname, f.buffer, { strictPdf });
        const cleanedBuf = await processPptxBuffer(sanitized, "V1");

        zip.addFile(outName(single, base, "cleaned.pptx"), cleanedBuf);
        zip.addFile(
          outName(single, base, "report.html"),
          Buffer.from(
            buildReportHtml({
              original: "[structured PPTX]",
              cleaned: "[PPTX corrected by AI]",
              rephrased: null,
              filename: f.originalname,
              mode: "V1",
            }),
            "utf8"
          )
        );
      } else if (ext === "pdf") {
        // ✅ sanitize PDF (annots + embedded files)
        const sanitizedPdf = await sanitizeBuffer(f.originalname, f.buffer, { strictPdf });
        const rawText = await extractPdfText(sanitizedPdf);
        const filtered = filterExtractedLines(rawText, { strictPdf });

        const cleanedTxt = (await aiCorrectText(filtered)) || filtered;
        const cleanedDocx = await createDocxFromText(cleanedTxt, base || "cleaned");

        zip.addFile(outName(single, base, "pdf_sanitized.pdf"), sanitizedPdf);
        zip.addFile(outName(single, base, "cleaned.docx"), cleanedDocx);
        zip.addFile(
          outName(single, base, "report.html"),
          Buffer.from(
            buildReportHtml({
              original: filtered,
              cleaned: cleanedTxt,
              rephrased: null,
              filename: f.originalname,
              mode: "V1",
            }),
            "utf8"
          )
        );
      } else {
        // Fallback texte brut
        const original = f.buffer.toString("utf8");
        const cleanedTxt = (await aiCorrectText(original)) || original;
        const cleanedDocx = await createDocxFromText(cleanedTxt, base || "cleaned");

        zip.addFile(outName(single, base, "cleaned.docx"), cleanedDocx);
        zip.addFile(
          outName(single, base, "report.html"),
          Buffer.from(
            buildReportHtml({
              original,
              cleaned: cleanedTxt,
              rephrased: null,
              filename: f.originalname,
              mode: "V1",
            }),
            "utf8"
          )
        );
      }
    }

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="docsafe_v1_result.zip"`
    );
    res.send(zip.toBuffer());
  } catch (err) {
    console.error("CLEAN ERROR", err);
    res.status(500).json({ error: String(err?.message || err) });
  }
});

// -----------------------------------------------------------------------------
// V2: /clean-v2  (correction + reformulation + report)
// -----------------------------------------------------------------------------
app.post("/clean-v2", upload.any(), async (req, res) => {
  try {
    const strictPdf = String(req.body.strictPdf || "false") === "true";
    const files = req.files || [];
    if (!files.length) return res.status(400).json({ error: "No files uploaded." });

    const single = files.length === 1;
    const zip = new AdmZip();

    for (const f of files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      if (ext === "docx") {
        // ✅ sanitize
        const sanitized = await sanitizeBuffer(f.originalname, f.buffer, { strictPdf });
        const cleanedBuf   = await processDocxBuffer(sanitized, "V1");
        const rephrasedBuf = await processDocxBuffer(sanitized, "V2");

        zip.addFile(outName(single, base, "cleaned.docx"), cleanedBuf);
        zip.addFile(outName(single, base, "rephrased.docx"), rephrasedBuf);
        zip.addFile(
          outName(single, base, "report.html"),
          Buffer.from(
            buildReportHtml({
              original: "[structured DOCX]",
              cleaned: "[DOCX corrected by AI]",
              rephrased: "[DOCX rephrased by AI]",
              filename: f.originalname,
              mode: "V2",
            }),
            "utf8"
          )
        );
      } else if (ext === "pptx") {
        // ✅ sanitize
        const sanitized = await sanitizeBuffer(f.originalname, f.buffer, { strictPdf });
        const cleanedBuf   = await processPptxBuffer(sanitized, "V1");
        const rephrasedBuf = await processPptxBuffer(sanitized, "V2");

        zip.addFile(outName(single, base, "cleaned.pptx"), cleanedBuf);
        zip.addFile(outName(single, base, "rephrased.pptx"), rephrasedBuf);
        zip.addFile(
          outName(single, base, "report.html"),
          Buffer.from(
            buildReportHtml({
              original: "[structured PPTX]",
              cleaned: "[PPTX corrected by AI]",
              rephrased: "[PPTX rephrased by AI]",
              filename: f.originalname,
              mode: "V2",
            }),
            "utf8"
          )
        );
      } else if (ext === "pdf") {
        // ✅ sanitize PDF (annots + embedded files)
        const sanitizedPdf = await sanitizeBuffer(f.originalname, f.buffer, { strictPdf });
        const rawText = await extractPdfText(sanitizedPdf);
        const filtered = filterExtractedLines(rawText, { strictPdf });

        const cleanedTxt   = (await aiCorrectText(filtered))   || filtered;
        const rephrasedTxt = (await aiRephraseText(cleanedTxt)) || cleanedTxt;

        const cleanedDocx   = await createDocxFromText(cleanedTxt,   base || "cleaned");
        const rephrasedDocx = await createDocxFromText(rephrasedTxt, base || "rephrased");

        zip.addFile(outName(single, base, "pdf_sanitized.pdf"), sanitizedPdf);
        zip.addFile(outName(single, base, "cleaned.docx"),   cleanedDocx);
        zip.addFile(outName(single, base, "rephrased.docx"), rephrasedDocx);
        zip.addFile(
          outName(single, base, "report.html"),
          Buffer.from(
            buildReportHtml({
              original: filtered,
              cleaned: cleanedTxt,
              rephrased: rephrasedTxt,
              filename: f.originalname,
              mode: "V2",
            }),
            "utf8"
          )
        );
      } else {
        // Fallback texte brut
        const original     = f.buffer.toString("utf8");
        const cleanedTxt   = (await aiCorrectText(original))   || original;
        const rephrasedTxt = (await aiRephraseText(cleanedTxt)) || cleanedTxt;

        const cleanedDocx   = await createDocxFromText(cleanedTxt,   base || "cleaned");
        const rephrasedDocx = await createDocxFromText(rephrasedTxt, base || "rephrased");

        zip.addFile(outName(single, base, "cleaned.docx"),   cleanedDocx);
        zip.addFile(outName(single, base, "rephrased.docx"), rephrasedDocx);
        zip.addFile(
          outName(single, base, "report.html"),
          Buffer.from(
            buildReportHtml({
              original,
              cleaned: cleanedTxt,
              rephrased: rephrasedTxt,
              filename: f.originalname,
              mode: "V2",
            }),
            "utf8"
          )
        );
      }
    }

    res.setHeader("Content-Type", "application/zip");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="docsafe_v2_result.zip"`
    );
    res.send(zip.toBuffer());
  } catch (err) {
    console.error("CLEAN V2 ERROR", err);
    res.status(500).json({ error: String(err?.message || err) });
  }
});

// -----------------------------------------------------------------------------
// DEBUG routes (inchangées)
// -----------------------------------------------------------------------------
app.post("/_docx_probe", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No DOCX uploaded" });
    const buf = req.file.buffer;
    const before = buf.length;
    const out = await processDocxBuffer(buf, "V1");
    const after = out.length;
    res.json({
      filename: req.file.originalname,
      size_in: before,
      size_out: after,
      changed: before !== after,
    });
  } catch (e) {
    console.error("DEBUG _docx_probe error:", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.post("/_docx_probe2", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No DOCX uploaded" });
    const mode = (String(req.query.mode || "V1").toUpperCase() === "V2") ? "V2" : "V1";
    const stats = await analyzeDocxBuffer(req.file.buffer, mode);
    res.json({
      mode,
      filename: req.file.originalname,
      paragraphs_total: stats.totalParagraphs,
      paragraphs_changed: stats.changedParagraphs,
      parts: stats.perPart,
    });
  } catch (e) {
    console.error("DEBUG _docx_probe2 error:", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.post("/_pdf_test", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No PDF uploaded" });
    const pdfSan = await stripPdfMetadata(req.file.buffer);
    const rawText = await extractPdfText(pdfSan);
    const filtered = filterExtractedLines(rawText, { strictPdf: true });
    res.json({
      filename: req.file.originalname,
      in_size: req.file.size,
      raw_len: rawText.length,
      filtered_len: filtered.length,
      excerpt: filtered.slice(0, 400),
    });
  } catch (e) {
    console.error("DEBUG _pdf_test error:", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

app.post("/_office_test", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const ext = getExt(req.file.originalname);
    let buf;
    if (ext === "docx") buf = await processDocxBuffer(req.file.buffer, "V1");
    else if (ext === "pptx") buf = await processPptxBuffer(req.file.buffer, "V1");
    else return res.status(400).json({ error: "Only DOCX or PPTX allowed" });
    res.setHeader("Content-Type", "application/octet-stream");
    res.setHeader("Content-Disposition", `attachment; filename="debug_out.${ext}"`);
    res.send(buf);
  } catch (e) {
    console.error("DEBUG _office_test error:", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// -----------------------------------------------------------------------------
// Listen
// -----------------------------------------------------------------------------
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend listening on ${PORT}`));


