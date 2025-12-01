// server.js - VERSION FINALE AVEC /analyze ET SPELLING ERRORS DANS REPORTS
import express from "express";
import cors from "cors";
import multer from "multer";
import AdmZip from "adm-zip";
import path from "path";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";

// Tes imports existants
import { cleanDOCX } from "./lib/docxCleaner.js";
import { cleanPPTX } from "./lib/pptxCleaner.js";
import { cleanPDF } from "./lib/pdfCleaner.js";
import { correctDOCXText, correctPPTXText } from "./lib/officeCorrect.js";
import { buildReportHtmlDetailed } from "./lib/report.js";
import { extractPdfText, filterExtractedLines } from "./lib/pdfTools.js";
import { createDocxFromText } from "./lib/docxWriter.js";
import { aiCorrectText } from "./lib/ai.js";

// NOUVEAU: Import de documentAnalyzer
import { analyzeDocument, calculateSummary } from "./lib/documentAnalyzer.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors({
  origin: [
    "https://mindorion.com", 
    "https://www.mindorion.com", 
    "http://localhost:5173",
    "http://localhost:8080",
    /\.lovableproject\.com$/,
    /\.lovable\.app$/
  ],
  methods: ["GET", "POST"],
  allowedHeaders: ["Content-Type", "Authorization"],
}));
app.use(express.json({ limit: "100mb" }));
app.use(express.urlencoded({ extended: true, limit: "100mb" }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 100 * 1024 * 1024 },
});

// ---------- Helpers ----------
const getExt = (fn = "") => (fn.includes(".") ? fn.split(".").pop().toLowerCase() : "");
const getMimeFromExt = (ext) => {
  const map = {
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pdf: 'application/pdf'
  };
  return map[ext] || 'application/octet-stream';
};
const outName = (single, base, name) => (single ? name : `${base}_${name}`);
const baseName = (filename = "document") => filename.replace(/\.[^.]+$/, "");

function sendZip(res, zip, zipName) {
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);
  res.send(zip.toBuffer());
}

// Helper: Calcul du score de risque (0-100, 100 = safe)
function calculateRiskScore(summary) {
  let score = 100;
  score -= summary.critical * 25;
  score -= summary.high * 10;
  score -= summary.medium * 5;
  score -= summary.low * 2;
  return Math.max(0, Math.min(100, score));
}

// Helper: Génération des recommandations
function generateRecommendations(detections, summary) {
  const recommendations = [];
  
  if (detections.metadata?.length > 0) {
    recommendations.push("Remove document metadata to protect author and organization information.");
  }
  if (detections.comments?.length > 0) {
    recommendations.push(`Review and remove ${detections.comments.length} comment(s) before sharing externally.`);
  }
  if (detections.trackChanges?.length > 0) {
    recommendations.push("Accept or reject all tracked changes to finalize the document.");
  }
  if (detections.hiddenContent?.length > 0 || detections.hiddenSheets?.length > 0) {
    recommendations.push("Remove hidden content that could expose confidential information.");
  }
  if (detections.macros?.length > 0) {
    recommendations.push("Remove macros for security - they can contain executable code.");
  }
  if (detections.sensitiveData?.length > 0) {
    const types = [...new Set(detections.sensitiveData.map(d => d.type))];
    recommendations.push(`Review sensitive data detected: ${types.join(', ')}.`);
  }
  if (detections.embeddedObjects?.length > 0) {
    recommendations.push("Remove embedded objects that may contain hidden data.");
  }
  if (detections.spellingErrors?.length > 0) {
    recommendations.push(`${detections.spellingErrors.length} spelling/grammar issue(s) were corrected.`);
  }
  if (recommendations.length === 0) {
    recommendations.push("Document appears clean. Minor review recommended before external sharing.");
  }
  
  return recommendations;
}

// ---------- Health ----------
app.get("/health", (_, res) =>
  res.json({ 
    ok: true, 
    service: "Qualion-Doc Backend", 
    version: "2.1",
    endpoints: ["/analyze", "/clean", "/rephrase"],
    time: new Date().toISOString() 
  })
);

// ===================================================================
// POST /analyze
// ===================================================================
app.post("/analyze", upload.single("file"), async (req, res) => {
  const startTime = Date.now();
  
  try {
    if (!req.file) {
      return res.status(400).json({ error: "No file uploaded." });
    }
    
    const ext = getExt(req.file.originalname);
    const supportedExts = ['docx', 'pptx', 'xlsx', 'pdf'];
    
    if (!supportedExts.includes(ext)) {
      return res.status(400).json({ 
        error: `Unsupported file type: .${ext}. Supported: ${supportedExts.join(', ')}` 
      });
    }
    
    console.log(`[ANALYZE] Processing ${req.file.originalname} (${(req.file.size / 1024).toFixed(1)} KB)`);
    
    const fileType = getMimeFromExt(ext);
    const detections = await analyzeDocument(req.file.buffer, fileType);
    const summary = calculateSummary(detections);
    const riskScore = calculateRiskScore(summary);
    
    const result = {
      documentId: uuidv4(),
      fileName: req.file.originalname,
      fileType: ext,
      fileSize: req.file.size,
      detections: {
        metadata: detections.metadata || [],
        comments: detections.comments || [],
        trackChanges: detections.trackChanges || [],
        hiddenContent: detections.hiddenContent || [],
        hiddenSheets: detections.hiddenSheets || [],
        sensitiveFormulas: detections.sensitiveFormulas || [],
        embeddedObjects: detections.embeddedObjects || [],
        macros: detections.macros || [],
        sensitiveData: detections.sensitiveData || [],
        spellingErrors: detections.spellingErrors || [],
        brokenLinks: detections.brokenLinks || [],
        businessInconsistencies: detections.businessInconsistencies || [],
        complianceRisks: detections.complianceRisks || []
      },
      summary: {
        totalIssues: summary.totalIssues,
        critical: summary.critical,
        high: summary.high,
        medium: summary.medium,
        low: summary.low,
        riskScore: riskScore,
        riskLevel: riskScore >= 90 ? 'safe' : riskScore >= 70 ? 'low' : riskScore >= 50 ? 'medium' : riskScore >= 25 ? 'high' : 'critical',
        recommendations: generateRecommendations(detections, summary)
      },
      processingTime: Date.now() - startTime
    };
    
    console.log(`[ANALYZE] Complete in ${result.processingTime}ms - ${summary.totalIssues} issues found, risk score: ${riskScore}`);
    
    res.json(result);
    
  } catch (e) {
    console.error("[ANALYZE ERROR]", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ===================================================================
// POST /clean  → MISE À JOUR AVEC SPELLING ERRORS DANS LE RAPPORT
// ===================================================================
app.post("/clean", upload.any(), async (req, res) => {
  try {
    const files = req.files || [];
    if (!files.length) return res.status(400).json({ error: "No files uploaded." });

    const drawPolicy = (req.body.drawPolicy || "auto").toLowerCase();
    const pdfMode = (req.body.pdfMode || "sanitize").toLowerCase();
    const includePdfDocx = String(req.body.pdfDocx || "false") === "true";
    
    const cleaningOptions = {
      removeMetadata: req.body.removeMetadata !== "false",
      removeComments: req.body.removeComments !== "false",
      acceptTrackChanges: req.body.acceptTrackChanges !== "false",
      removeHiddenContent: req.body.removeHiddenContent !== "false",
      removeEmbeddedObjects: req.body.removeEmbeddedObjects !== "false",
      removeMacros: req.body.removeMacros !== "false",
      correctSpelling: req.body.correctSpelling !== "false",
    };

    const single = files.length === 1;
    const zip = new AdmZip();

    for (const f of files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      console.log(`[CLEAN] Processing ${f.originalname} with options:`, cleaningOptions);

      // ============================================================
      // NOUVEAU: Analyser le document AVANT le nettoyage pour avoir
      // les spellingErrors et le summary pour le rapport
      // ============================================================
      let analysisResult = null;
      try {
        const fileType = getMimeFromExt(ext);
        const detections = await analyzeDocument(f.buffer, fileType);
        const summary = calculateSummary(detections);
        const riskScore = calculateRiskScore(summary);
        
        analysisResult = {
          detections,
          summary: {
            ...summary,
            riskScore,
            riskLevel: riskScore >= 90 ? 'safe' : riskScore >= 70 ? 'low' : riskScore >= 50 ? 'medium' : riskScore >= 25 ? 'high' : 'critical',
            recommendations: generateRecommendations(detections, summary)
          }
        };
        
        console.log(`[CLEAN] Analysis found ${detections.spellingErrors?.length || 0} spelling errors`);
      } catch (analysisError) {
        console.warn(`[CLEAN] Analysis failed, continuing without:`, analysisError.message);
      }

      if (ext === "docx") {
        const cleaned = await cleanDOCX(f.buffer, { drawPolicy, ...cleaningOptions });
        
        let finalBuffer = cleaned.outBuffer;
        let correctionStats = null;
        
        if (cleaningOptions.correctSpelling) {
          const corrected = await correctDOCXText(cleaned.outBuffer, aiCorrectText);
          finalBuffer = corrected.outBuffer;
          correctionStats = corrected.stats;
        }

        zip.addFile(outName(single, base, "cleaned.docx"), finalBuffer);

        // MISE À JOUR: Passer analysis et spellingErrors au rapport
        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: { drawPolicy, ...cleaningOptions },
          cleaning: cleaned.stats,
          correction: correctionStats,
          analysis: analysisResult,  // ← NOUVEAU
          spellingErrors: analysisResult?.detections?.spellingErrors || []  // ← NOUVEAU
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));
        
      } else if (ext === "pptx") {
        const cleaned = await cleanPPTX(f.buffer, { drawPolicy, ...cleaningOptions });
        
        let finalBuffer = cleaned.outBuffer;
        let correctionStats = null;
        
        if (cleaningOptions.correctSpelling) {
          const corrected = await correctPPTXText(cleaned.outBuffer, aiCorrectText);
          finalBuffer = corrected.outBuffer;
          correctionStats = corrected.stats;
        }

        zip.addFile(outName(single, base, "cleaned.pptx"), finalBuffer);

        // MISE À JOUR: Passer analysis et spellingErrors au rapport
        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: { drawPolicy, ...cleaningOptions },
          cleaning: cleaned.stats,
          correction: correctionStats,
          analysis: analysisResult,  // ← NOUVEAU
          spellingErrors: analysisResult?.detections?.spellingErrors || []  // ← NOUVEAU
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));
        
      } else if (ext === "pdf") {
        const cleaned = await cleanPDF(f.buffer, {
          pdfMode: pdfMode === "text-only" ? "text-only" : "sanitize",
          extractTextFn: async (b) => filterExtractedLines(await extractPdfText(b), { strictPdf: true }),
        });

        zip.addFile(
          outName(single, base, pdfMode === "text-only" ? "text_only.pdf" : "sanitized.pdf"),
          cleaned.outBuffer
        );

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
              ? [{ before: filtered.slice(0, 140), after: correctedTxt.slice(0, 140) }]
              : [],
          };
        }

        // MISE À JOUR: Passer analysis et spellingErrors au rapport
        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: { pdfMode, ...cleaningOptions },
          cleaning: cleaned.stats,
          correction: correctionStats,
          analysis: analysisResult,  // ← NOUVEAU
          spellingErrors: analysisResult?.detections?.spellingErrors || []  // ← NOUVEAU
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));
        
      } else if (ext === "xlsx") {
        // Pour Excel, on ajoute juste le fichier tel quel pour l'instant
        zip.addFile(outName(single, base, f.originalname), f.buffer);
        
        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: cleaningOptions,
          cleaning: {},
          correction: null,
          analysis: analysisResult,  // ← NOUVEAU
          spellingErrors: analysisResult?.detections?.spellingErrors || []  // ← NOUVEAU
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));
        
      } else {
        zip.addFile(outName(single, base, f.originalname), f.buffer);
        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: {},
          cleaning: {},
          correction: null,
          analysis: null,
          spellingErrors: []
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));
      }
    }

    const zipName = files.length === 1
      ? `${baseName(files[0].originalname)} cleaned.zip`
      : "qualion_doc_cleaned.zip";

    sendZip(res, zip, zipName);
  } catch (e) {
    console.error("CLEAN ERROR", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ===================================================================
// POST /rephrase
// ===================================================================
app.post("/rephrase", upload.any(), async (req, res) => {
  try {
    const files = req.files || [];
    if (!files.length) return res.status(400).json({ error: "No files uploaded." });

    const drawPolicy = (req.body.drawPolicy || "auto").toLowerCase();
    const single = files.length === 1;
    const zip = new AdmZip();

    for (const f of files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      // Analyse pour le rapport
      let analysisResult = null;
      try {
        const fileType = getMimeFromExt(ext);
        const detections = await analyzeDocument(f.buffer, fileType);
        const summary = calculateSummary(detections);
        const riskScore = calculateRiskScore(summary);
        
        analysisResult = {
          detections,
          summary: {
            ...summary,
            riskScore,
            riskLevel: riskScore >= 90 ? 'safe' : riskScore >= 70 ? 'low' : riskScore >= 50 ? 'medium' : riskScore >= 25 ? 'high' : 'critical',
            recommendations: generateRecommendations(detections, summary)
          }
        };
      } catch (analysisError) {
        console.warn(`[REPHRASE] Analysis failed:`, analysisError.message);
      }

      if (ext === "docx") {
        const cleaned = await cleanDOCX(f.buffer, { drawPolicy });
        const rephrased = await correctDOCXText(cleaned.outBuffer, aiCorrectText, {
          mode: "rephrase",
        });

        zip.addFile(outName(single, base, "rephrased.docx"), rephrased.outBuffer);

        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: { drawPolicy, mode: "rephrase" },
          cleaning: cleaned.stats,
          correction: rephrased.stats,
          analysis: analysisResult,
          spellingErrors: analysisResult?.detections?.spellingErrors || []
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));
        
      } else if (ext === "pptx") {
        const cleaned = await cleanPPTX(f.buffer, { drawPolicy });
        const rephrased = await correctPPTXText(cleaned.outBuffer, aiCorrectText, {
          mode: "rephrase",
        });

        zip.addFile(outName(single, base, "rephrased.pptx"), rephrased.outBuffer);

        const report = buildReportHtmlDetailed({
          filename: f.originalname,
          ext,
          policy: { drawPolicy, mode: "rephrase" },
          cleaning: cleaned.stats,
          correction: rephrased.stats,
          analysis: analysisResult,
          spellingErrors: analysisResult?.detections?.spellingErrors || []
        });
        zip.addFile(outName(single, base, "report.html"), Buffer.from(report, "utf8"));
        
      } else if (ext === "pdf") {
        return res.status(400).json({ error: "Rephrase for PDF is disabled. Convert to DOCX/PPTX first." });
      } else {
        return res.status(400).json({ error: `Unsupported file for rephrase: .${ext}` });
      }
    }

    const zipName = files.length === 1
      ? `${baseName(files[0].originalname)} rephrased.zip`
      : "qualion_doc_rephrased.zip";

    sendZip(res, zip, zipName);
  } catch (e) {
    console.error("REPHRASE ERROR", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ---------- Boot ----------
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log(`✅ Qualion-Doc Backend v2.1 listening on port ${PORT}`);
  console.log(`   Endpoints: GET /health, POST /analyze, POST /clean, POST /rephrase`);
});
