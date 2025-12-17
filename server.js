// server.js - VERSION 2.9.0
// ✅ Fix Render crash (brace structure)
// ✅ docStats signature fixed (extractDocStats({ ext, buffer }))
// ✅ Select Items to Clean => REAL removal (selected only when provided)
// ✅ Keeps backward compatibility (old + new payload shapes)

import express from "express";
import cors from "cors";
import multer from "multer";
import AdmZip from "adm-zip";
import path from "path";
import { fileURLToPath } from "url";
import { v4 as uuidv4 } from "uuid";

// Imports existants
import { cleanDOCX } from "./lib/docxCleaner.js";
import { cleanPPTX } from "./lib/pptxCleaner.js";
import { cleanPDF } from "./lib/pdfCleaner.js";
import { correctDOCXText, correctPPTXText, correctXLSXText } from "./lib/officeCorrect.js";
import { buildReportHtmlDetailed, buildReportData } from "./lib/report.js";
import { extractPdfText, filterExtractedLines } from "./lib/pdfTools.js";
import { createDocxFromText } from "./lib/docxWriter.js";
import { aiCorrectText } from "./lib/ai.js";
import { cleanXLSX } from "./lib/xlsxCleaner.js";

// Import de documentAnalyzer
import { analyzeDocument } from "./lib/documentAnalyzer.js";

// 🆕 Import docStats (NEW SIGNATURE)
import { extractDocStats } from "./lib/docStats.js";

// 🆕 Import sensitive data cleaner
import {
  removeSensitiveDataFromDOCX,
  removeSensitiveDataFromPPTX,
  removeSensitiveDataFromXLSX,
  removeHiddenContentFromDOCX,
  removeHiddenContentFromPPTX,
  removeVisualObjectsFromPPTX,
} from "./lib/sensitiveDataCleaner.js";

// ✅ NEW: extracted helpers/policies (refactor only)
import { sendZip, outName, baseName } from "./lib/http/zip.js";
import { calculateRiskScore, calculateAfterScore, getRiskLevel } from "./lib/policy/riskScore.js";
import { generateRecommendations } from "./lib/policy/recommendations.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();

app.use(
  cors({
    origin: [
      "https://mindorion.com",
      "https://www.mindorion.com",
      "http://localhost:5173",
      "http://localhost:8080",
      /\.lovableproject\.com$/,
      /\.lovable\.app$/,
    ],
    methods: ["GET", "POST"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

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
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    pdf: "application/pdf",
  };
  return map[ext] || "application/octet-stream";
};

// ✅ Safe wrapper for doc stats (never breaks pipeline)
// IMPORTANT: extractDocStats signature is now extractDocStats({ ext, buffer, ... })
async function safeExtractDocStats(buffer, ext) {
  try {
    const stats = await extractDocStats({ ext, buffer });
    return (
      stats || {
        pages: null,
        slides: null,
        sheets: null,
        tables: null,
        paragraphs: null,
        approxWords: null,
      }
    );
  } catch (e) {
    console.warn("[DOC STATS] Failed:", e?.message || e);
    return {
      pages: null,
      slides: null,
      sheets: null,
      tables: null,
      paragraphs: null,
      approxWords: null,
    };
  }
}

// 🆕 Safe JSON parse helper
function safeJsonParse(str, fallback = []) {
  if (str === undefined || str === null) return fallback;
  if (typeof str !== "string") return fallback;
  if (!str.trim()) return fallback;
  try {
    const parsed = JSON.parse(str);
    return Array.isArray(parsed) ? parsed : fallback;
  } catch (e) {
    console.warn("[JSON PARSE] Failed:", e?.message);
    return fallback;
  }
}

// Helper: detect if a key was provided (even if empty array)
function bodyHasKey(req, key) {
  return Object.prototype.hasOwnProperty.call(req.body || {}, key);
}

function addReportsToZip(zip, single, base, reportParams) {
  const reportHtml = buildReportHtmlDetailed(reportParams);
  zip.addFile(outName(single, base, "report.html"), Buffer.from(reportHtml, "utf8"));

  const reportJson = buildReportData(reportParams);
  zip.addFile(outName(single, base, "report.json"), Buffer.from(JSON.stringify(reportJson, null, 2), "utf8"));
}

// ---------- Health ----------
app.get("/health", (_, res) =>
  res.json({
    ok: true,
    service: "Qualion-Doc Backend",
    version: "2.9.0",
    endpoints: ["/analyze", "/clean", "/rephrase"],
    features: [
      "approvedSpellingErrors",
      "premium-json-report",
      "accurate-risk-score",
      "score-impacts",
      "fixed-detections",
      "documentStats-before-after",
      "sensitive-data-removal",
      "hidden-content-removal",
      "visual-objects-removal",
      "selective-cleaning-by-checkbox",
    ],
    time: new Date().toISOString(),
  })
);

// ===================================================================
// POST /analyze - VERSION 2.9.0
// ===================================================================
app.post("/analyze", upload.single("file"), async (req, res) => {
  const startTime = Date.now();

  try {
    if (!req.file) return res.status(400).json({ error: "No file uploaded." });

    const ext = getExt(req.file.originalname);
    const supportedExts = ["docx", "pptx", "xlsx", "pdf"];
    if (!supportedExts.includes(ext)) {
      return res.status(400).json({
        error: `Unsupported file type: .${ext}. Supported: ${supportedExts.join(", ")}`,
      });
    }

    console.log(
      `[ANALYZE] Processing ${req.file.originalname} (${(req.file.size / 1024).toFixed(1)} KB)`
    );

    const fileType = getMimeFromExt(ext);
    const analysisResult = await analyzeDocument(req.file.buffer, fileType);
    const detections = analysisResult.detections;

    const rawSummary = analysisResult.summary;
    const summary = {
      totalIssues: rawSummary.totalIssues,
      critical: rawSummary.criticalIssues,
      high: rawSummary.highIssues,
      medium: rawSummary.mediumIssues,
      low: rawSummary.lowIssues,
    };

    const { score: riskScore, breakdown } = calculateRiskScore(summary, detections);

    // ✅ documentStats
    const documentStats = await safeExtractDocStats(req.file.buffer, ext);

    res.json({
      documentId: uuidv4(),
      fileName: req.file.originalname,
      fileType: ext,
      fileSize: req.file.size,

      documentStats,

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
        complianceRisks: detections.complianceRisks || [],
        visualObjects: detections.visualObjects || [],
        orphanData: detections.orphanData || [],
        excelHiddenData: detections.excelHiddenData || [],
      },

      summary: {
        ...summary,
        riskScore,
        riskLevel: getRiskLevel(riskScore),
        riskBreakdown: breakdown,
        recommendations: generateRecommendations(detections),
      },

      processingTime: Date.now() - startTime,
    });
  } catch (e) {
    console.error("[ANALYZE ERROR]", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ===================================================================
// POST /clean - VERSION 2.9.0
// ✅ Select Items to Clean => REAL removal
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

    // ✅ Parse approved spelling errors
    const approvedSpellingErrors = safeJsonParse(req.body.approvedSpellingErrors, []);

    // ✅ Detect if front is sending selective lists (even empty)
    const hasSelectiveSensitive =
      bodyHasKey(req, "sensitiveDataToClean") || bodyHasKey(req, "removeSensitiveData");
    const hasSelectiveHidden =
      bodyHasKey
