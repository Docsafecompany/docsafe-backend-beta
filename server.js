// server.js
import express from "express";
import cors from "cors";
import multer from "multer";
import AdmZip from "adm-zip";
import path from "path";
import { fileURLToPath } from "url";

import { processDocxBuffer, processPptxBuffer } from "./lib/officeXml.js";
import { stripPdfMetadata, extractPdfText, filterExtractedLines } from "./lib/pdfTools.js";
import { buildReportHtml } from "./lib/report.js";
import { createDocxFromText } from "./lib/docxWriter.js";
import { aiCorrectText, aiRephraseText } from "./lib/ai.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors({ origin: "*", credentials: false }));
app.use(express.json({ limit: "32mb" }));
app.use(express.urlencoded({ extended: true, limit: "32mb" }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

/** Helpers */
function getExt(filename = "") {
  const i = filename.lastIndexOf(".");
  return i >= 0 ? filename.slice(i + 1).toLowerCase() : "";
}
function outName(single, base, name) {
  return single ? name : `${base}_${name}`;
}

/** Health / Echo */
app.get("/health", (_, res) => res.json({ ok: true, service: "DocSafe Backend", time: new Date().toISOString() }));
app.get("/_env_ok", (_, res) => res.json({ OPENAI_API_KEY: process.env.OPENAI_API_KEY ? "present" : "absent" }));

app.get("/_ai_echo", async (req, res) => {
  const sample = String(req.query.q || "This is   a  test,, please  fix punctuation!! Thanks");
  const out = (await aiCorrectText(sample)) || sample; // fallback si pas de clé
  res.json({ input: sample, corrected: out });
});
app.get("/_ai_rephrase_echo", async (req, res) => {
  const sample = String(req.query.q || "Overall, this is a sentence that could be rephrased furthermore.");
  const out = (await aiRephraseText(sample)) || sample; // fallback si pas de clé
  res.json({ input: sample, rephrased: out });
});

/** V1: correction */
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
        const cleanedBuf = await processDocxBuffer(f.buffer, aiCorrectText, "V1");
        zip.addFile(outName(single, base, `cleaned.docx`), cleanedBuf);
        zip.addFile(outName(single, base, `report.html`), Buffer.from(buildReportHtml({
          original: "[structured DOCX]",
          cleaned: "[DOCX corrected by AI]",
          rephrased: null,
          filename: f.originalname,
          mode: "V1"
        }), "utf8"));
      } else if (ext === "pptx") {
        const cleanedBuf = await processPptxBuffer(f.buffer, "V1");
        zip.addFile(outName(single, base, `cleaned.pptx`), cleanedBuf);
        zip.addFile(outName(single, base, `report.html`), Buffer.from(buildReportHtml({
          original: "[structured PPTX]",
          cleaned: "[PPTX corrected by AI]",
          rephrased: null,
          filename: f.originalname,
          mode: "V1"
        }), "utf8"));
      } else if (ext === "pdf") {
        const pdfSan = await stripPdfMetadata(f.buffer);
        const rawText = await extractPdfText(pdfSan);
        const filtered = filterExtractedLines(rawText, { strictPdf });
        const cleaned = (await aiCorrectText(filtered)) || filtered;
        const cleanedDocx = await createDocxFromText(cleaned, base || "cleaned");

        zip.addFile(outName(single, base, `pdf_sanitized.pdf`), pdfSan);
        zip.addFile(outName(single, base, `cleaned.docx`), cleanedDocx);
        zip.addFile(outName(single, base, `report.html`), Buffer.from(buildReportHtml({
          original: filtered, cleaned, rephrased: null, filename: f.originalname, mode: "V1"
        }), "utf8"));
      } else {
        const original = f.buffer.toString("utf8");
        const cleaned = (await aiCorrectText(original)) || original;
        const cleanedDocx = await createDocxFromText(cleaned, base || "cleaned");
        zip.addFile(outName(single, base, `cleaned.docx`), cleanedDocx);
        zip.addFile(outName(single, base, `report.html`), Buffer.from(buildReportHtml({
          original, cleaned, rephrased: null, filename: f.originalname, mode: "V1"
        }), "utf8"));
      }
    }

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="docsafe_v1_result.zip"`);
    res.send(zip.toBuffer());
  } catch (err) {
    console.error("CLEAN ERROR", err);
    res.status(500).json({ error: String(err?.message || err) });
  }
});

/** V2: correction + reformulation */
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
        const cleanedBuf = await processDocxBuffer(f.buffer, aiCorrectText, "V1");
        const rephrasedBuf = await processDocxBuffer(f.buffer, aiRephraseText, "V2");

        zip.addFile(outName(single, base, `cleaned.docx`), cleanedBuf);
        zip.addFile(outName(single, base, `rephrased.docx`), rephrasedBuf);
        zip.addFile(outName(single, base, `report.html`), Buffer.from(buildReportHtml({
          original: "[structured DOCX]",
          cleaned: "[DOCX corrected by AI]",
          rephrased: "[DOCX rephrased by AI]",
          filename: f.originalname,
          mode: "V2"
        }), "utf8"));
      } else if (ext === "pptx") {
        const cleanedBuf = await processPptxBuffer(f.buffer, "V1");
        const rephrasedBuf = await processPptxBuffer(f.buffer, "V2");

        zip.addFile(outName(single, base, `cleaned.pptx`), cleanedBuf);
        zip.addFile(outName(single, base, `rephrased.pptx`), rephrasedBuf);
        zip.addFile(outName(single, base, `report.html`), Buffer.from(buildReportHtml({
          original: "[structured PPTX]",
          cleaned: "[PPTX corrected by AI]",
          rephrased: "[PPTX rephrased by AI]",
          filename: f.originalname,
          mode: "V2"
        }), "utf8"));
      } else if (ext === "pdf") {
        const pdfSan = await stripPdfMetadata(f.buffer);
        const rawText = await extractPdfText(pdfSan);
        const filtered = filterExtractedLines(rawText, { strictPdf });

        const cleaned = (await aiCorrectText(filtered)) || filtered;
        const rephrased = (await aiRephraseText(cleaned)) || cleaned;

        const cleanedDocx = await createDocxFromText(cleaned, base || "cleaned");
        const rephrasedDocx = await createDocxFromText(rephrased, base || "rephrased");

        zip.addFile(outName(single, base, `pdf_sanitized.pdf`), pdfSan);
        zip.addFile(outName(single, base, `cleaned.docx`), cleanedDocx);
        zip.addFile(outName(single, base, `rephrased.docx`), rephrasedDocx);
        zip.addFile(outName(single, base, `report.html`), Buffer.from(buildReportHtml({
          original: filtered, cleaned, rephrased, filename: f.originalname, mode: "V2"
        }), "utf8"));
      } else {
        const original = f.buffer.toString("utf8");
        const cleaned = (await aiCorrectText(original)) || original;
        const rephrased = (await aiRephraseText(cleaned)) || cleaned;
        const cleanedDocx = await createDocxFromText(cleaned, base || "cleaned");
        const rephrasedDocx = await createDocxFromText(rephrased, base || "rephrased");

        zip.addFile(outName(single, base, `cleaned.docx`), cleanedDocx);
        zip.addFile(outName(single, base, `rephrased.docx`), rephrasedDocx);
        zip.addFile(outName(single, base, `report.html`), Buffer.from(buildReportHtml({
          original, cleaned, rephrased, filename: f.originalname, mode: "V2"
        }), "utf8"));
      }
    }

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="docsafe_v2_result.zip"`);
    res.send(zip.toBuffer());
  } catch (err) {
    console.error("CLEAN V2 ERROR", err);
    res.status(500).json({ error: String(err?.message || err) });
  }
});

/** DEBUG */
app.get("/_ping", (req, res) => res.json({ pong: true, time: new Date().toISOString() }));
app.post("/_pdf_test", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No PDF uploaded" });
    const pdfSan = await stripPdfMetadata(req.file.buffer);
    const rawText = await extractPdfText(pdfSan);
    const filtered = filterExtractedLines(rawText, { strictPdf: true });
    res.json({ filename: req.file.originalname, size: req.file.size, rawLength: rawText.length, filteredLength: filtered.length, excerpt: filtered.slice(0, 500) });
  } catch (e) {
    console.error("DEBUG _pdf_test error:", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});
app.post("/_office_test", upload.single("file"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded" });
    const ext = req.file.originalname.split(".").pop().toLowerCase();
    let buf;
    if (ext === "docx")      buf = await processDocxBuffer(req.file.buffer, aiCorrectText, "V1");
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

const PORT = process.env.PORT || 10000;
// ===== DEBUG PROBE: pour vérifier qu'un DOCX est bien modifié =====
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
      changed: before !== after
    });
  } catch (e) {
    console.error("DEBUG _docx_probe error:", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});
app.listen(PORT, () => console.log(`DocSafe backend listening on ${PORT}`));


