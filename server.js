// server.js
import express from "express";
import cors from "cors";
import multer from "multer";
import AdmZip from "adm-zip";
import path from "path";
import { fileURLToPath } from "url";

// === TES FONCTIONS IA EXISTANTES (orthographe/grammaire) ===
import { aiCorrectText } from "./lib/ai.js"; // doit corriger sans reformuler

// === OUTILS PDF EXISTANTS (si tu veux un DOCX à partir du PDF) ===
import { extractPdfText, filterExtractedLines } from "./lib/pdfTools.js";
import { createDocxFromText } from "./lib/docxWriter.js";

// === CLEANERS (privacy + dessins selon policy) ===
import { cleanDOCX } from "./lib/docxCleaner.js";
import { cleanPPTX } from "./lib/pptxCleaner.js";
import { cleanPDF  } from "./lib/pdfCleaner.js";

// === CORRECTEURS NON DESTRUCTIFS DES TEXTES DOCX/PPTX ===
import { correctDOCXText, correctPPTXText } from "./lib/officeCorrect.js";

// ------------------------------------------------------------------
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors({ origin: "*", credentials: false }));
app.use(express.json({ limit: "32mb" }));
app.use(express.urlencoded({ extended: true, limit: "32mb" }));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

const getExt = (fn = "") => fn.includes(".") ? fn.split(".").pop().toLowerCase() : "";
const outName = (single, base, name) => (single ? name : `${base}_${name}`);

// Helpers sanitation
async function sanitizeOffice(originalName, buf, { drawPolicy }) {
  const ext = getExt(originalName);
  if (ext === "docx")  return (await cleanDOCX(buf, { drawPolicy })).outBuffer;
  if (ext === "pptx")  return (await cleanPPTX(buf, { drawPolicy })).outBuffer;
  return buf;
}
async function sanitizePdf(buf, { pdfMode }) {
  if (pdfMode === "text-only") {
    const { outBuffer } = await cleanPDF(buf, {
      pdfMode,
      extractTextFn: async (b) => filterExtractedLines(await extractPdfText(b), { strictPdf: true })
    });
    return outBuffer;
  }
  return (await cleanPDF(buf, { pdfMode })).outBuffer;
}

// ------------------------------------------------------------------
// Health
app.get("/health", (_, res) => res.json({ ok: true, service: "DocSafe Backend", time: new Date().toISOString() }));

// ------------------------------------------------------------------
// CLEAN (privacy + correction textuelle, mise en page intacte)
app.post("/clean", upload.any(), async (req, res) => {
  try {
    const files = req.files || [];
    if (!files.length) return res.status(400).json({ error: "No files uploaded." });

    // Front flags
    const drawPolicy = (req.body.drawPolicy || "auto").toLowerCase();   // auto|all|none
    const pdfMode    = (req.body.pdfMode || "sanitize").toLowerCase();  // sanitize|text-only
    const includePdfDocx = String(req.body.pdfDocx || "false") === "true"; // option: DOCX corrigé depuis PDF

    const single = files.length === 1;
    const zip = new AdmZip();

    for (const f of files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      if (ext === "docx") {
        // 1) Privacy + doodles selon policy (images conservées en auto)
        const sanitized = await sanitizeOffice(f.originalname, f.buffer, { drawPolicy });
        // 2) Correction ciblée des <w:t>
        const corrected = await correctDOCXText(sanitized, aiCorrectText);
        zip.addFile(outName(single, base, "cleaned.docx"), corrected);

      } else if (ext === "pptx") {
        const sanitized = await sanitizeOffice(f.originalname, f.buffer, { drawPolicy });
        const corrected = await correctPPTXText(sanitized, aiCorrectText);
        zip.addFile(outName(single, base, "cleaned.pptx"), corrected);

      } else if (ext === "pdf") {
        // PDF: privacy-clean par défaut
        const sanitizedPdf = await sanitizePdf(f.buffer, { pdfMode: "sanitize" });
        zip.addFile(outName(single, base, "sanitized.pdf"), sanitizedPdf);

        if (pdfMode === "text-only") {
          const textOnly = await sanitizePdf(f.buffer, { pdfMode: "text-only" });
          zip.addFile(outName(single, base, "text_only.pdf"), textOnly);
        }

        // Optionnel: livrer aussi un DOCX corrigé depuis le texte PDF
        if (includePdfDocx) {
          const raw = await extractPdfText(sanitizedPdf);
          const filtered = filterExtractedLines(raw, { strictPdf: true });
          const correctedTxt = await aiCorrectText(filtered);
          const docxFromPdf = await createDocxFromText(correctedTxt || filtered, base || "corrected");
          zip.addFile(outName(single, base, "corrected_from_pdf.docx"), docxFromPdf);
        }

      } else {
        // autres formats (no-op)
        zip.addFile(outName(single, base, f.originalname), f.buffer);
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

// (Optionnel) /clean-v2 si tu veux une variante — sinon garde /clean unique.
// ------------------------------------------------------------------

const PORT = process.env.PORT || 10000;
app.listen(PORT, () => console.log(`DocSafe backend listening on ${PORT}`));
