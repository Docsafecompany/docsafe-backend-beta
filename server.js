// server.js - VERSION 2.9.1
// ✅ Refactor-only: extract zip helpers + risk scoring + recommendations into lib/
// ✅ No behavior change intended (pipeline identical)
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

// ✅ NEW: extracted helpers
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
    version: "2.9.1",
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
      "refactor-riskScore-zip-helpers",
    ],
    time: new Date().toISOString(),
  })
);

// ===================================================================
// POST /analyze - VERSION 2.9.1
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

    console.log(`[ANALYZE] Processing ${req.file.originalname} (${(req.file.size / 1024).toFixed(1)} KB)`);

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
// POST /clean - VERSION 2.9.1
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
    const hasSelectiveSensitive = bodyHasKey(req, "sensitiveDataToClean") || bodyHasKey(req, "removeSensitiveData");
    const hasSelectiveHidden = bodyHasKey(req, "hiddenContentToClean") || bodyHasKey(req, "hiddenContentToCleanRaw"); // (safe)
    const hasSelectiveVisual = bodyHasKey(req, "visualObjectsToClean");

    // ✅ Accept BOTH payload shapes from frontend (old + new)
    const removeSensitiveDataRaw =
      safeJsonParse(req.body.removeSensitiveData, null) ?? safeJsonParse(req.body.sensitiveDataToClean, []);

    const hiddenContentToCleanRaw = safeJsonParse(req.body.hiddenContentToClean, []);
    const visualObjectsToCleanRaw = safeJsonParse(req.body.visualObjectsToClean, []);

    console.log(`[CLEAN] removeSensitiveData raw count: ${removeSensitiveDataRaw?.length || 0}`);
    console.log(`[CLEAN] hiddenContentToClean raw count: ${hiddenContentToCleanRaw?.length || 0}`);
    console.log(`[CLEAN] visualObjectsToClean raw count: ${visualObjectsToCleanRaw?.length || 0}`);
    console.log(`[CLEAN] Selective modes: sensitive=${hasSelectiveSensitive}, hidden=${hasSelectiveHidden}, visual=${hasSelectiveVisual}`);

    const single = files.length === 1;
    const zip = new AdmZip();

    for (const f of files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      console.log(`[CLEAN] Processing ${f.originalname} with options:`, cleaningOptions);

      // BEFORE structural stats
      const documentStatsBefore = await safeExtractDocStats(f.buffer, ext);

      // analysis (optional)
      let analysisResult = null;
      let spellingErrors = [];
      let beforeRiskScore = 100;
      let riskBreakdown = {};
      let detections = null;
      let summary = null;

      try {
        const fileType = getMimeFromExt(ext);
        const fullAnalysis = await analyzeDocument(f.buffer, fileType);
        detections = fullAnalysis.detections;

        const rawSummary = fullAnalysis.summary;
        summary = {
          totalIssues: rawSummary.totalIssues,
          critical: rawSummary.criticalIssues,
          high: rawSummary.highIssues,
          medium: rawSummary.mediumIssues,
          low: rawSummary.lowIssues,
        };

        const riskResult = calculateRiskScore(summary, detections);
        beforeRiskScore = riskResult.score;
        riskBreakdown = riskResult.breakdown;

        spellingErrors = detections.spellingErrors || [];

        analysisResult = {
          detections,
          documentStats: documentStatsBefore,
          summary: {
            ...summary,
            riskScore: beforeRiskScore,
            beforeRiskScore,
            riskBreakdown,
            riskLevel: getRiskLevel(beforeRiskScore),
            recommendations: generateRecommendations(detections),
          },
        };
      } catch (analysisError) {
        console.warn(`[CLEAN] Analysis failed, continuing without:`, analysisError?.message || analysisError);
      }

      // Map selection -> full objects
      let sensitiveDataToRemove = [];
      if ((removeSensitiveDataRaw?.length || 0) > 0 && detections?.sensitiveData) {
        if (typeof removeSensitiveDataRaw[0] === "string") {
          sensitiveDataToRemove = detections.sensitiveData.filter((d) => removeSensitiveDataRaw.includes(d.id));
        } else {
          sensitiveDataToRemove = removeSensitiveDataRaw;
        }
      }

      let hiddenContentToRemove = [];
      if ((hiddenContentToCleanRaw?.length || 0) > 0 && detections?.hiddenContent) {
        if (typeof hiddenContentToCleanRaw[0] === "string") {
          hiddenContentToRemove = detections.hiddenContent.filter((d) => hiddenContentToCleanRaw.includes(d.id));
        } else {
          hiddenContentToRemove = hiddenContentToCleanRaw;
        }
      }

      let visualObjectsToRemove = [];
      if ((visualObjectsToCleanRaw?.length || 0) > 0 && detections?.visualObjects) {
        if (typeof visualObjectsToCleanRaw[0] === "string") {
          visualObjectsToRemove = detections.visualObjects.filter((d) => visualObjectsToCleanRaw.includes(d.id));
        } else {
          visualObjectsToRemove = visualObjectsToCleanRaw;
        }
      }

      console.log(`[CLEAN] Selected: sensitive=${sensitiveDataToRemove.length}, hidden=${hiddenContentToRemove.length}, visual=${visualObjectsToRemove.length}`);

      // If approvedSpellingErrors is provided and not empty, use it
      const spellingFixList =
        Array.isArray(approvedSpellingErrors) && approvedSpellingErrors.length > 0 ? approvedSpellingErrors : spellingErrors;

      // Track extra removals for score calculation
      const extraRemovals = {
        sensitiveDataRemoved: 0,
        hiddenContentRemoved: 0,
      };

      // ---------------- DOCX ----------------
      if (ext === "docx") {
        let currentBuffer = f.buffer;

        // Step 1: Standard cleaning (uses toggles)
        const cleaned = await cleanDOCX(currentBuffer, { drawPolicy, ...cleaningOptions });
        currentBuffer = cleaned.outBuffer;

        // Step 2: Sensitive removal (SELECTED ONLY if list provided)
        if (hasSelectiveSensitive) {
          if (sensitiveDataToRemove.length > 0) {
            const sensitiveResult = await removeSensitiveDataFromDOCX(currentBuffer, sensitiveDataToRemove);
            currentBuffer = sensitiveResult.outBuffer;
            extraRemovals.sensitiveDataRemoved = sensitiveResult.stats.removed;
          }
        }

        // Step 3: Hidden content removal (SELECTED ONLY if list provided)
        if (hasSelectiveHidden) {
          if (hiddenContentToRemove.length > 0) {
            const hiddenResult = await removeHiddenContentFromDOCX(currentBuffer, hiddenContentToRemove);
            currentBuffer = hiddenResult.outBuffer;
            extraRemovals.hiddenContentRemoved = hiddenResult.stats.removed;
          }
        }

        // Step 4: Correct spelling
        let correctionStats = null;
        if (cleaningOptions.correctSpelling) {
          const corrected = await correctDOCXText(currentBuffer, aiCorrectText, {
            spellingErrors: spellingFixList,
          });
          currentBuffer = corrected.outBuffer;
          correctionStats = corrected.stats;
        }

        // AFTER stats
        const documentStatsAfter = await safeExtractDocStats(currentBuffer, ext);

        zip.addFile(outName(single, base, "cleaned.docx"), currentBuffer);

        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats, riskBreakdown, extraRemovals);

        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { drawPolicy, ...cleaningOptions },
          cleaning: {
            ...cleaned.stats,
            sensitiveDataRemoved: extraRemovals.sensitiveDataRemoved,
            hiddenContentRemoved: extraRemovals.hiddenContentRemoved,
          },
          correction: correctionStats,
          analysis: analysisResult,
          spellingErrors,
          approvedSpellingErrors,
          beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts,
          documentStatsBefore,
          documentStatsAfter,
        });

        continue;
      }

      // ---------------- PPTX ----------------
      if (ext === "pptx") {
        let currentBuffer = f.buffer;

        // Step 1: Standard cleaning
        const cleaned = await cleanPPTX(currentBuffer, { drawPolicy, ...cleaningOptions });
        currentBuffer = cleaned.outBuffer;

        // Step 2: Sensitive (selected only)
        if (hasSelectiveSensitive) {
          if (sensitiveDataToRemove.length > 0) {
            const sensitiveResult = await removeSensitiveDataFromPPTX(currentBuffer, sensitiveDataToRemove);
            currentBuffer = sensitiveResult.outBuffer;
            extraRemovals.sensitiveDataRemoved = sensitiveResult.stats.removed;
          }
        }

        // Step 3: Hidden (selected only)
        if (hasSelectiveHidden) {
          if (hiddenContentToRemove.length > 0) {
            const hiddenResult = await removeHiddenContentFromPPTX(currentBuffer, hiddenContentToRemove);
            currentBuffer = hiddenResult.outBuffer;
            extraRemovals.hiddenContentRemoved = hiddenResult.stats.removed;
          }
        }

        // Step 4: Visual objects (selected only)
        if (hasSelectiveVisual) {
          if (visualObjectsToRemove.length > 0) {
            const visualResult = await removeVisualObjectsFromPPTX(currentBuffer, visualObjectsToRemove);
            currentBuffer = visualResult.outBuffer;
          }
        }

        // Step 5: Correct spelling
        let correctionStats = null;
        if (cleaningOptions.correctSpelling) {
          const corrected = await correctPPTXText(currentBuffer, aiCorrectText, {
            spellingErrors: spellingFixList,
          });
          currentBuffer = corrected.outBuffer;
          correctionStats = corrected.stats;
        }

        // AFTER stats
        const documentStatsAfter = await safeExtractDocStats(currentBuffer, ext);

        zip.addFile(outName(single, base, "cleaned.pptx"), currentBuffer);

        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats, riskBreakdown, extraRemovals);

        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { drawPolicy, ...cleaningOptions },
          cleaning: {
            ...cleaned.stats,
            sensitiveDataRemoved: extraRemovals.sensitiveDataRemoved,
            hiddenContentRemoved: extraRemovals.hiddenContentRemoved,
          },
          correction: correctionStats,
          analysis: analysisResult,
          spellingErrors,
          approvedSpellingErrors,
          beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts,
          documentStatsBefore,
          documentStatsAfter,
        });

        continue;
      }

      // ---------------- PDF ----------------
      if (ext === "pdf") {
        const cleaned = await cleanPDF(f.buffer, {
          pdfMode: pdfMode === "text-only" ? "text-only" : "sanitize",
          extractTextFn: async (b) => filterExtractedLines(await extractPdfText(b), { strictPdf: true }),
        });

        zip.addFile(outName(single, base, pdfMode === "text-only" ? "text_only.pdf" : "sanitized.pdf"), cleaned.outBuffer);

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
            examples:
              correctedTxt && filtered && correctedTxt.trim() !== filtered.trim()
                ? [{ before: filtered.slice(0, 140), after: correctedTxt.slice(0, 140) }]
                : [],
          };
        }

        const documentStatsAfter = await safeExtractDocStats(cleaned.outBuffer, ext);

        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats, riskBreakdown, extraRemovals);

        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { pdfMode, ...cleaningOptions },
          cleaning: cleaned.stats,
          correction: correctionStats,
          analysis: analysisResult,
          spellingErrors,
          approvedSpellingErrors,
          beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts,
          documentStatsBefore,
          documentStatsAfter,
        });

        continue;
      }

      // ---------------- XLSX ----------------
      if (ext === "xlsx") {
        let currentBuffer = f.buffer;

        // Step 1: Standard cleaning
        const cleaned = await cleanXLSX(currentBuffer, cleaningOptions);
        currentBuffer = cleaned.outBuffer;

        // Step 2: Sensitive (selected only)
        if (hasSelectiveSensitive) {
          if (sensitiveDataToRemove.length > 0) {
            const sensitiveResult = await removeSensitiveDataFromXLSX(currentBuffer, sensitiveDataToRemove);
            currentBuffer = sensitiveResult.outBuffer;
            extraRemovals.sensitiveDataRemoved = sensitiveResult.stats.removed;
          }
        }

        // Step 3: Correct spelling
        let correctionStats = null;
        if (cleaningOptions.correctSpelling) {
          const corrected = await correctXLSXText(currentBuffer, aiCorrectText, {
            spellingErrors: spellingFixList,
          });
          currentBuffer = corrected.outBuffer;
          correctionStats = corrected.stats;
        }

        const documentStatsAfter = await safeExtractDocStats(currentBuffer, ext);

        zip.addFile(outName(single, base, "cleaned.xlsx"), currentBuffer);

        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats, riskBreakdown, extraRemovals);

        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: cleaningOptions,
          cleaning: { ...cleaned.stats, sensitiveDataRemoved: extraRemovals.sensitiveDataRemoved },
          correction: correctionStats,
          analysis: analysisResult,
          spellingErrors,
          approvedSpellingErrors,
          beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts,
          documentStatsBefore,
          documentStatsAfter,
        });

        continue;
      }

      // ---------------- Other ----------------
      zip.addFile(outName(single, base, f.originalname), f.buffer);

      addReportsToZip(zip, single, base, {
        filename: f.originalname,
        ext,
        policy: {},
        cleaning: {},
        correction: null,
        analysis: null,
        spellingErrors: [],
        approvedSpellingErrors: [],
        beforeRiskScore: 100,
        afterRiskScore: 100,
        scoreImpacts: {},
        documentStatsBefore,
        documentStatsAfter: documentStatsBefore,
      });
    }

    const zipName = files.length === 1 ? `${baseName(files[0].originalname)} cleaned.zip` : "qualion_doc_cleaned.zip";
    sendZip(res, zip, zipName);
  } catch (e) {
    console.error("CLEAN ERROR", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ===================================================================
// POST /rephrase - VERSION 2.9.1
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

      const documentStatsBefore = await safeExtractDocStats(f.buffer, ext);

      let analysisResult = null;
      let spellingErrors = [];
      let beforeRiskScore = 100;
      let riskBreakdown = {};
      let detections = null;
      let summary = null;

      try {
        const fileType = getMimeFromExt(ext);
        const fullAnalysis = await analyzeDocument(f.buffer, fileType);
        detections = fullAnalysis.detections;

        const rawSummary = fullAnalysis.summary;
        summary = {
          totalIssues: rawSummary.totalIssues,
          critical: rawSummary.criticalIssues,
          high: rawSummary.highIssues,
          medium: rawSummary.mediumIssues,
          low: rawSummary.lowIssues,
        };

        const riskResult = calculateRiskScore(summary, detections);
        beforeRiskScore = riskResult.score;
        riskBreakdown = riskResult.breakdown;

        spellingErrors = detections.spellingErrors || [];

        analysisResult = {
          detections,
          documentStats: documentStatsBefore,
          summary: {
            ...summary,
            riskScore: beforeRiskScore,
            beforeRiskScore,
            riskBreakdown,
            riskLevel: getRiskLevel(beforeRiskScore),
            recommendations: generateRecommendations(detections),
          },
        };
      } catch (analysisError) {
        console.warn(`[REPHRASE] Analysis failed:`, analysisError?.message || analysisError);
      }

      if (ext === "docx") {
        const cleaned = await cleanDOCX(f.buffer, { drawPolicy });

        const rephrased = await correctDOCXText(cleaned.outBuffer, aiCorrectText, {
          mode: "rephrase",
          spellingErrors,
        });

        const documentStatsAfter = await safeExtractDocStats(rephrased.outBuffer, ext);

        zip.addFile(outName(single, base, "rephrased.docx"), rephrased.outBuffer);

        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, rephrased.stats, riskBreakdown, {});

        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { drawPolicy, mode: "rephrase" },
          cleaning: cleaned.stats,
          correction: rephrased.stats,
          analysis: analysisResult,
          spellingErrors,
          beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts,
          documentStatsBefore,
          documentStatsAfter,
        });
      } else if (ext === "pptx") {
        const cleaned = await cleanPPTX(f.buffer, { drawPolicy });

        const rephrased = await correctPPTXText(cleaned.outBuffer, aiCorrectText, {
          mode: "rephrase",
          spellingErrors,
        });

        const documentStatsAfter = await safeExtractDocStats(rephrased.outBuffer, ext);

        zip.addFile(outName(single, base, "rephrased.pptx"), rephrased.outBuffer);

        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, rephrased.stats, riskBreakdown, {});

        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { drawPolicy, mode: "rephrase" },
          cleaning: cleaned.stats,
          correction: rephrased.stats,
          analysis: analysisResult,
          spellingErrors,
          beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts,
          documentStatsBefore,
          documentStatsAfter,
        });
      } else if (ext === "pdf") {
        return res.status(400).json({
          error: "Rephrase for PDF is disabled. Convert to DOCX/PPTX first.",
        });
      } else {
        return res.status(400).json({ error: `Unsupported file for rephrase: .${ext}` });
      }
    }

    const zipName = files.length === 1 ? `${baseName(files[0].originalname)} rephrased.zip` : "qualion_doc_rephrased.zip";
    sendZip(res, zip, zipName);
  } catch (e) {
    console.error("REPHRASE ERROR", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ---------- Boot ----------
const PORT = process.env.PORT || 10000;
app.listen(PORT, () => {
  console.log(`✅ Qualion-Doc Backend v2.9.1 listening on port ${PORT}`);
  console.log(`   Endpoints: GET /health, POST /analyze, POST /clean, POST /rephrase`);
  console.log(`   Features: selective-cleaning-by-checkbox + docStats fixed + refactor helpers`);
});
