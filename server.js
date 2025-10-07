// server.js
import express from "express";
import cors from "cors";
import multer from "multer";
import AdmZip from "adm-zip";
import path from "path";
import { fileURLToPath } from "url";

import { cleanDOCX } from "./lib/docxCleaner.js";
import { cleanPPTX } from "./lib/pptxCleaner.js";
import { cleanPDF  } from "./lib/pdfCleaner.js";
import { correctDOCXText, correctPPTXText } from "./lib/officeCorrect.js";
import { buildReportHtmlDetailed } from "./lib/report.js";

// (facultatif, pour PDF -> DOCX corrigé)
import { extractPdfText, filterExtractedLines } from "./lib/pdfTools.js";
import { createDocxFromText } from "./lib/docxWriter.js";
import { aiCorrectText } from "./lib/ai.js"; // doit corriger sans reformuler

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors({ origin: "*", credentials: false }));
app.use(express.json({ limit: "32mb" }));
app.use(express.urlencoded({ extended: true, limit: "32mb" }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

const getExt = (fn = "") => fn.includes(".") ? fn.split(".").pop().toLowerCase() : "";
const outName = (single, base, name) => (single ? name : `${base}_${name}`);

app.get("/health", (_, res) =>
  res.json({ ok: true, service: "DocSafe Backend", time: new Date().toISOString() })
);

// ✅ Route principale — Clean + Correction + Report
app.post("/clean", upload.any(), async (req, res) => {
  try {
    const files = req.files || [];
    if (!files.length) return res.status(400).json({ error: "No files uploaded." });

    // options depuis le frontend
    const drawPolicy = (req.body.drawPolicy || "auto").toLowerCase();    // auto|all|none
    const pdfMode    = (req.body.pdfMode || "sanitize").toLowerCase();   // sanitize|text-only
    const includePdfDocx = String(req.body.pdfDocx || "false") === "true";

    const single = files.length === 1;
    const zip = new AdmZip();

    for (const f of files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      if (ext === "docx") {
        // 1) Clean (stats)
        const cleaned = await cleanDOCX(f.buffer, { drawPolicy });
        // 2) Correction ciblée des nœuds texte <w:t> (stats)
        const corrected = await correctDOCXText(cleaned.outBuffer, aiCorrectText);

        zip.addFile(outName(single, base, "cleaned.docx"), corrected.outBuffer);

        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: { drawPolicy },
          cleaning: cleaned.stats,
          correction: corrected.stats
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));

      } else if (ext === "pptx") {
        const cleaned = await cleanPPTX(f.buffer, { drawPolicy });
        const corrected = await correctPPTXText(cleaned.outBuffer, aiCorrectText);

        zip.addFile(outName(single, base, "cleaned.pptx"), corrected.outBuffer);

        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: { drawPolicy },
          cleaning: cleaned.stats,
          correction: corrected.stats
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));

      } else if (ext === "pdf") {
        const cleaned = await cleanPDF(f.buffer, {
          pdfMode: pdfMode === 'text-only' ? 'text-only' : 'sanitize',
          extractTextFn: async (b) => filterExtractedLines(await extractPdfText(b), { strictPdf: true })
        });

        zip.addFile(outName(single, base, pdfMode === 'text-only' ? "text_only.pdf" : "sanitized.pdf"), cleaned.outBuffer);

        // Option: produire en plus un DOCX corrigé à partir du PDF (pour édition)
        let correctionStats = null;
        if (includePdfDocx) {
          const raw = await extractPdfText(cleaned.outBuffer);
          const filtered = filterExtractedLines(raw, { strictPdf: true });
          const correctedTxt = await aiCorrectText(filtered);
          const docxFromPdf = await createDocxFromText(correctedTxt || filtered, base || "corrected");
          zip.addFile(outName(single, base, "corrected_from_pdf.docx"), docxFromPdf);

          correctionStats = {
            totalTextNodes: 0,
            changedTextNodes: correctedTxt && filtered ? (correctedTxt.trim() === filtered.trim() ? 0 : 1) : 0,
            examples: correctedTxt && filtered && correctedTxt.trim() !== filtered.trim()
              ? [{ before: filtered.slice(0,140), after: correctedTxt.slice(0,140) }]
              : []
          };
        }

        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: { pdfMode },
          cleaning: cleaned.stats,
          correction: correctionStats
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));

      } else {
        // autres formats → on laisse tel quel + report basique
        zip.addFile(outName(single, base, f.originalname), f.buffer);
        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: {},
          cleaning: {},
          correction: null
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));
      }
    }

    res.setHeader("Content-Type", "application/zip");
    res.setHeader("Content-Disposition", `attachment; filename="docsafe_result.zip"`);
    res.send(zip.toBuffer());
  } catch (e) {
    console.error("CLEAN ERROR", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend listening on ${PORT}`));
