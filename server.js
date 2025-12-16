// server.js - VERSION 2.8.0 (adds sensitive data & hidden content ACTUAL REMOVAL)
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
import {
  correctDOCXText,
  correctPPTXText,
  correctXLSXText,
} from "./lib/officeCorrect.js";
import { buildReportHtmlDetailed, buildReportData } from "./lib/report.js";
import { extractPdfText, filterExtractedLines } from "./lib/pdfTools.js";
import { createDocxFromText } from "./lib/docxWriter.js";
import { aiCorrectText } from "./lib/ai.js";
import { cleanXLSX } from "./lib/xlsxCleaner.js";

// Import de documentAnalyzer
import { analyzeDocument } from "./lib/documentAnalyzer.js";

// 🆕 Import docStats
import { extractDocStats } from "./lib/docStats.js";

// 🆕 Import sensitive data cleaner (NEW FILE)
import {
  removeSensitiveDataFromDOCX,
  removeSensitiveDataFromPPTX,
  removeSensitiveDataFromXLSX,
  removeHiddenContentFromDOCX,
  removeHiddenContentFromPPTX,
  removeVisualObjectsFromPPTX,
} from "./lib/sensitiveDataCleaner.js";

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
const getExt = (fn = "") =>
  fn.includes(".") ? fn.split(".").pop().toLowerCase() : "";

const getMimeFromExt = (ext) => {
  const map = {
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    pdf: "application/pdf",
  };
  return map[ext] || "application/octet-stream";
};

const outName = (single, base, name) => (single ? name : `${base}_${name}`);
const baseName = (filename = "document") => filename.replace(/\.[^.]+$/, "");

function sendZip(res, zip, zipName) {
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);
  res.send(zip.toBuffer());
}

// 🆕 Safe wrapper for doc stats (never breaks pipeline)
async function safeExtractDocStats(buffer, ext) {
  try {
    const stats = await extractDocStats(buffer, ext);
    return (
      stats || {
        pages: null,
        slides: null,
        sheets: null,
        tables: null,
      }
    );
  } catch (e) {
    console.warn("[DOC STATS] Failed:", e?.message || e);
    return { pages: null, slides: null, sheets: null, tables: null };
  }
}

// 🆕 Safe JSON parse helper
function safeJsonParse(str, fallback = []) {
  if (!str) return fallback;
  try {
    const parsed = JSON.parse(str);
    return Array.isArray(parsed) ? parsed : fallback;
  } catch (e) {
    console.warn("[JSON PARSE] Failed:", e?.message);
    return fallback;
  }
}

// ============================================================
// Calcul du score de risque basé sur les détections
// ============================================================
function calculateRiskScore(summary, detections = null) {
  let score = 100;
  const breakdown = {};

  if (summary.critical > 0) {
    const penalty = summary.critical * 25;
    score -= penalty;
    breakdown.critical = penalty;
  }
  if (summary.high > 0) {
    const penalty = summary.high * 10;
    score -= penalty;
    breakdown.high = penalty;
  }
  if (summary.medium > 0) {
    const penalty = summary.medium * 5;
    score -= penalty;
    breakdown.medium = penalty;
  }
  if (summary.low > 0) {
    const penalty = summary.low * 2;
    score -= penalty;
    breakdown.low = penalty;
  }

  if (detections) {
    const sensitiveCount = detections.sensitiveData?.length || 0;
    if (sensitiveCount > 0) {
      const penalty = Math.min(sensitiveCount * 25, 50);
      score -= penalty;
      breakdown.sensitiveData = penalty;
    }

    const macrosCount = detections.macros?.length || 0;
    if (macrosCount > 0) {
      const penalty = Math.min(macrosCount * 15, 30);
      score -= penalty;
      breakdown.macros = penalty;
    }

    const hiddenCount =
      (detections.hiddenContent?.length || 0) +
      (detections.hiddenSheets?.length || 0);
    if (hiddenCount > 0) {
      const penalty = Math.min(hiddenCount * 8, 24);
      score -= penalty;
      breakdown.hiddenContent = penalty;
    }

    const commentsCount = detections.comments?.length || 0;
    if (commentsCount > 0) {
      const penalty = Math.min(commentsCount * 3, 15);
      score -= penalty;
      breakdown.comments = penalty;
    }

    const trackChangesCount = detections.trackChanges?.length || 0;
    if (trackChangesCount > 0) {
      const penalty = Math.min(trackChangesCount * 3, 15);
      score -= penalty;
      breakdown.trackChanges = penalty;
    }

    const metadataCount = detections.metadata?.length || 0;
    if (metadataCount > 0) {
      const penalty = Math.min(metadataCount * 2, 10);
      score -= penalty;
      breakdown.metadata = penalty;
    }

    const embeddedCount = detections.embeddedObjects?.length || 0;
    if (embeddedCount > 0) {
      const penalty = Math.min(embeddedCount * 5, 15);
      score -= penalty;
      breakdown.embeddedObjects = penalty;
    }

    const spellingCount = detections.spellingErrors?.length || 0;
    if (spellingCount > 0) {
      const penalty = Math.min(spellingCount * 1, 10);
      score -= penalty;
      breakdown.spellingGrammar = penalty;
    }

    const brokenLinksCount = detections.brokenLinks?.length || 0;
    if (brokenLinksCount > 0) {
      const penalty = Math.min(brokenLinksCount * 4, 12);
      score -= penalty;
      breakdown.brokenLinks = penalty;
    }

    const complianceCount = detections.complianceRisks?.length || 0;
    if (complianceCount > 0) {
      const penalty = Math.min(complianceCount * 12, 36);
      score -= penalty;
      breakdown.complianceRisks = penalty;
    }
  } else {
    const classifiedIssues =
      (summary.critical || 0) +
      (summary.high || 0) +
      (summary.medium || 0) +
      (summary.low || 0);
    const unclassifiedIssues = Math.max(
      0,
      (summary.totalIssues || 0) - classifiedIssues
    );

    if (unclassifiedIssues > 0) {
      const penalty = unclassifiedIssues * 5;
      score -= penalty;
      breakdown.unclassified = penalty;
    }
  }

  if (summary.totalIssues > 10) {
    const penalty = (summary.totalIssues - 10) * 2;
    score -= penalty;
    breakdown.volumePenalty = penalty;
  }

  const finalScore = Math.max(0, Math.min(100, score));
  console.log(
    `[RISK SCORE] Calculated: ${finalScore} (total issues: ${summary.totalIssues})`,
    breakdown
  );
  return { score: finalScore, breakdown };
}

// ============================================================
// Calcul du score APRÈS nettoyage
// ============================================================
function calculateAfterScore(beforeScore, cleaningStats, correctionStats, riskBreakdown = {}, extraRemovals = {}) {
  let improvement = 0;
  const scoreImpacts = {};

  if (cleaningStats?.metaRemoved > 0 && riskBreakdown.metadata) {
    const impact = Math.min(cleaningStats.metaRemoved * 2, riskBreakdown.metadata);
    improvement += impact;
    scoreImpacts.metadata = impact;
  } else if (cleaningStats?.metaRemoved > 0) {
    const impact = Math.min(cleaningStats.metaRemoved * 2, 10);
    improvement += impact;
    scoreImpacts.metadata = impact;
  }

  if (cleaningStats?.commentsXmlRemoved > 0 && riskBreakdown.comments) {
    const impact = Math.min(cleaningStats.commentsXmlRemoved * 3, riskBreakdown.comments);
    improvement += impact;
    scoreImpacts.comments = impact;
  } else if (cleaningStats?.commentsXmlRemoved > 0) {
    const impact = Math.min(cleaningStats.commentsXmlRemoved * 3, 15);
    improvement += impact;
    scoreImpacts.comments = impact;
  }

  const trackChangesTotal =
    (cleaningStats?.revisionsAccepted?.deletionsRemoved || 0) +
    (cleaningStats?.revisionsAccepted?.insertionsUnwrapped || 0);

  if (trackChangesTotal > 0 && riskBreakdown.trackChanges) {
    const impact = Math.min(trackChangesTotal * 3, riskBreakdown.trackChanges);
    improvement += impact;
    scoreImpacts.trackChanges = impact;
  } else if (trackChangesTotal > 0) {
    const impact = Math.min(trackChangesTotal * 3, 15);
    improvement += impact;
    scoreImpacts.trackChanges = impact;
  }

  if (cleaningStats?.hiddenRemoved > 0 && riskBreakdown.hiddenContent) {
    const impact = Math.min(cleaningStats.hiddenRemoved * 8, riskBreakdown.hiddenContent);
    improvement += impact;
    scoreImpacts.hiddenContent = impact;
  } else if (cleaningStats?.hiddenRemoved > 0) {
    const impact = Math.min(cleaningStats.hiddenRemoved * 8, 24);
    improvement += impact;
    scoreImpacts.hiddenContent = impact;
  }

  if (cleaningStats?.macrosRemoved > 0 && riskBreakdown.macros) {
    const impact = Math.min(cleaningStats.macrosRemoved * 15, riskBreakdown.macros);
    improvement += impact;
    scoreImpacts.macros = impact;
  } else if (cleaningStats?.macrosRemoved > 0) {
    const impact = Math.min(cleaningStats.macrosRemoved * 15, 30);
    improvement += impact;
    scoreImpacts.macros = impact;
  }

  const embeddedTotal = (cleaningStats?.mediaDeleted || 0) + (cleaningStats?.picturesRemoved || 0);
  if (embeddedTotal > 0 && riskBreakdown.embeddedObjects) {
    const impact = Math.min(embeddedTotal * 5, riskBreakdown.embeddedObjects);
    improvement += impact;
    scoreImpacts.embeddedObjects = impact;
  } else if (embeddedTotal > 0) {
    const impact = Math.min(embeddedTotal * 5, 15);
    improvement += impact;
    scoreImpacts.embeddedObjects = impact;
  }

  if (correctionStats?.changedTextNodes > 0 && riskBreakdown.spellingGrammar) {
    const impact = Math.min(correctionStats.changedTextNodes * 1, riskBreakdown.spellingGrammar);
    improvement += impact;
    scoreImpacts.spellingGrammar = impact;
  } else if (correctionStats?.changedTextNodes > 0) {
    const impact = Math.min(correctionStats.changedTextNodes * 1, 10);
    improvement += impact;
    scoreImpacts.spellingGrammar = impact;
  }

  // 🆕 Account for sensitive data removal
  if (extraRemovals.sensitiveDataRemoved > 0 && riskBreakdown.sensitiveData) {
    const impact = Math.min(extraRemovals.sensitiveDataRemoved * 25, riskBreakdown.sensitiveData);
    improvement += impact;
    scoreImpacts.sensitiveData = impact;
  }

  // 🆕 Account for hidden content removal
  if (extraRemovals.hiddenContentRemoved > 0 && riskBreakdown.hiddenContent) {
    const impact = Math.min(extraRemovals.hiddenContentRemoved * 8, riskBreakdown.hiddenContent);
    improvement += impact;
    scoreImpacts.hiddenContentExtra = impact;
  }

  const afterScore = Math.min(100, beforeScore + improvement);
  return { score: afterScore, scoreImpacts, improvement };
}

function generateRecommendations(detections) {
  const recommendations = [];
  if (detections.metadata?.length > 0)
    recommendations.push(
      "Remove document metadata to protect author and organization information."
    );
  if (detections.comments?.length > 0)
    recommendations.push(
      `Review and remove ${detections.comments.length} comment(s) before sharing externally.`
    );
  if (detections.trackChanges?.length > 0)
    recommendations.push(
      "Accept or reject all tracked changes to finalize the document."
    );
  if (detections.hiddenContent?.length > 0 || detections.hiddenSheets?.length > 0)
    recommendations.push(
      "Remove hidden content that could expose confidential information."
    );
  if (detections.macros?.length > 0)
    recommendations.push(
      "Remove macros for security - they can contain executable code."
    );
  if (detections.sensitiveData?.length > 0) {
    const types = [...new Set(detections.sensitiveData.map((d) => d.type))];
    recommendations.push(`Review sensitive data detected: ${types.join(", ")}.`);
  }
  if (detections.embeddedObjects?.length > 0)
    recommendations.push("Remove embedded objects that may contain hidden data.");
  if (detections.spellingErrors?.length > 0)
    recommendations.push(
      `${detections.spellingErrors.length} spelling/grammar issue(s) were detected.`
    );
  if (recommendations.length === 0)
    recommendations.push(
      "Document appears clean. Minor review recommended before external sharing."
    );
  return recommendations;
}

function getRiskLevel(score) {
  if (score >= 90) return "safe";
  if (score >= 70) return "low";
  if (score >= 50) return "medium";
  if (score >= 25) return "high";
  return "critical";
}

function addReportsToZip(zip, single, base, reportParams) {
  const reportHtml = buildReportHtmlDetailed(reportParams);
  zip.addFile(outName(single, base, "report.html"), Buffer.from(reportHtml, "utf8"));

  const reportJson = buildReportData(reportParams);
  zip.addFile(
    outName(single, base, "report.json"),
    Buffer.from(JSON.stringify(reportJson, null, 2), "utf8")
  );
}

// ---------- Health ----------
app.get("/health", (_, res) =>
  res.json({
    ok: true,
    service: "Qualion-Doc Backend",
    version: "2.8.0",
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
    ],
    time: new Date().toISOString(),
  })
);

// ===================================================================
// POST /analyze - VERSION 2.8.0
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
      `[ANALYZE] Processing ${req.file.originalname} (${(
        req.file.size / 1024
      ).toFixed(1)} KB)`
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

    // documentStats
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
// POST /clean  → VERSION 2.8.0 (ACTUAL REMOVAL of sensitive data & hidden content)
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

// ✅ Accept BOTH payload shapes from frontend (old + new)
// Old: removeSensitiveData / hiddenContentToClean / visualObjectsToClean
// New: sensitiveDataToClean / hiddenContentToClean / visualObjectsToClean
const removeSensitiveDataRaw =
  safeJsonParse(req.body.removeSensitiveData, null) ??
  safeJsonParse(req.body.sensitiveDataToClean, []);

const hiddenContentToCleanRaw =
  safeJsonParse(req.body.hiddenContentToClean, []);

const visualObjectsToCleanRaw =
  safeJsonParse(req.body.visualObjectsToClean, []);

// Optional debug: if full objects, log the values
try {
  if (Array.isArray(removeSensitiveDataRaw) && removeSensitiveDataRaw.length > 0) {
    const first = removeSensitiveDataRaw[0];
    if (typeof first === "object" && first?.value) {
      console.log(
        "[CLEAN] sensitiveData values:",
        removeSensitiveDataRaw.map(s => s?.value).filter(Boolean)
      );
    }
  }
} catch (e) {
  console.warn("[CLEAN] sensitiveData log error:", e?.message || e);
}

console.log(`[CLEAN] removeSensitiveData: ${removeSensitiveDataRaw?.length || 0} items`);
console.log(`[CLEAN] hiddenContentToClean: ${hiddenContentToCleanRaw.length} items`);
console.log(`[CLEAN] visualObjectsToClean: ${visualObjectsToCleanRaw.length} items`);

const single = files.length === 1;
const zip = new AdmZip();

for (const f of files) {
  const ext = getExt(f.originalname);
  const base = path.parse(f.originalname).name;

  console.log(`[CLEAN] Processing ${f.originalname} with options:`, cleaningOptions);
  // ...
}


      // BEFORE structural stats
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
        console.warn(`[CLEAN] Analysis failed, continuing without:`, analysisError.message);
      }

      // 🆕 Map IDs to full detection objects for sensitive data
      let sensitiveDataToRemove = [];
      if (removeSensitiveDataRaw.length > 0 && detections?.sensitiveData) {
        // If raw contains IDs (strings), find matching detections
        if (typeof removeSensitiveDataRaw[0] === "string") {
          sensitiveDataToRemove = detections.sensitiveData.filter((d) =>
            removeSensitiveDataRaw.includes(d.id)
          );
        } else {
          // Already full objects
          sensitiveDataToRemove = removeSensitiveDataRaw;
        }
      }

      // 🆕 Map IDs to full detection objects for hidden content
      let hiddenContentToRemove = [];
      if (hiddenContentToCleanRaw.length > 0 && detections?.hiddenContent) {
        if (typeof hiddenContentToCleanRaw[0] === "string") {
          hiddenContentToRemove = detections.hiddenContent.filter((d) =>
            hiddenContentToCleanRaw.includes(d.id)
          );
        } else {
          hiddenContentToRemove = hiddenContentToCleanRaw;
        }
      }

      // 🆕 Map IDs for visual objects
      let visualObjectsToRemove = [];
      if (visualObjectsToCleanRaw.length > 0 && detections?.visualObjects) {
        if (typeof visualObjectsToCleanRaw[0] === "string") {
          visualObjectsToRemove = detections.visualObjects.filter((d) =>
            visualObjectsToCleanRaw.includes(d.id)
          );
        } else {
          visualObjectsToRemove = visualObjectsToCleanRaw;
        }
      }

      console.log(`[CLEAN] Will remove: ${sensitiveDataToRemove.length} sensitive, ${hiddenContentToRemove.length} hidden, ${visualObjectsToRemove.length} visual`);

      // If approvedSpellingErrors is provided and not empty, use it
      const spellingFixList =
        Array.isArray(approvedSpellingErrors) && approvedSpellingErrors.length > 0
          ? approvedSpellingErrors
          : spellingErrors;

      // Track extra removals for score calculation
      const extraRemovals = {
        sensitiveDataRemoved: 0,
        hiddenContentRemoved: 0,
      };

      if (ext === "docx") {
        let currentBuffer = f.buffer;

        // Step 1: Standard cleaning
        const cleaned = await cleanDOCX(currentBuffer, { drawPolicy, ...cleaningOptions });
        currentBuffer = cleaned.outBuffer;

        // 🆕 Step 2: Remove sensitive data
        if (sensitiveDataToRemove.length > 0) {
          const sensitiveResult = await removeSensitiveDataFromDOCX(currentBuffer, sensitiveDataToRemove);
          currentBuffer = sensitiveResult.outBuffer;
          extraRemovals.sensitiveDataRemoved = sensitiveResult.stats.removed;
        }

        // 🆕 Step 3: Remove hidden content
        if (hiddenContentToRemove.length > 0) {
          const hiddenResult = await removeHiddenContentFromDOCX(currentBuffer, hiddenContentToRemove);
          currentBuffer = hiddenResult.outBuffer;
          extraRemovals.hiddenContentRemoved = hiddenResult.stats.removed;
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

        // AFTER structural stats
        const documentStatsAfter = await safeExtractDocStats(currentBuffer, ext);

        zip.addFile(outName(single, base, "cleaned.docx"), currentBuffer);

        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats, riskBreakdown, extraRemovals);

        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { drawPolicy, ...cleaningOptions },
          cleaning: { ...cleaned.stats, sensitiveDataRemoved: extraRemovals.sensitiveDataRemoved, hiddenContentRemoved: extraRemovals.hiddenContentRemoved },
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

      } else if (ext === "pptx") {
        let currentBuffer = f.buffer;

        // Step 1: Standard cleaning
        const cleaned = await cleanPPTX(currentBuffer, { drawPolicy, ...cleaningOptions });
        currentBuffer = cleaned.outBuffer;

        // 🆕 Step 2: Remove sensitive data
        if (sensitiveDataToRemove.length > 0) {
          const sensitiveResult = await removeSensitiveDataFromPPTX(currentBuffer, sensitiveDataToRemove);
          currentBuffer = sensitiveResult.outBuffer;
          extraRemovals.sensitiveDataRemoved = sensitiveResult.stats.removed;
        }

        // 🆕 Step 3: Remove hidden content
        if (hiddenContentToRemove.length > 0) {
          const hiddenResult = await removeHiddenContentFromPPTX(currentBuffer, hiddenContentToRemove);
          currentBuffer = hiddenResult.outBuffer;
          extraRemovals.hiddenContentRemoved = hiddenResult.stats.removed;
        }

        // 🆕 Step 4: Remove visual objects (covering shapes)
        if (visualObjectsToRemove.length > 0) {
          const visualResult = await removeVisualObjectsFromPPTX(currentBuffer, visualObjectsToRemove);
          currentBuffer = visualResult.outBuffer;
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

        // AFTER structural stats
        const documentStatsAfter = await safeExtractDocStats(currentBuffer, ext);

        zip.addFile(outName(single, base, "cleaned.pptx"), currentBuffer);

        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats, riskBreakdown, extraRemovals);

        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { drawPolicy, ...cleaningOptions },
          cleaning: { ...cleaned.stats, sensitiveDataRemoved: extraRemovals.sensitiveDataRemoved, hiddenContentRemoved: extraRemovals.hiddenContentRemoved },
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
            changedTextNodes:
              correctedTxt && filtered ? (correctedTxt.trim() === filtered.trim() ? 0 : 1) : 0,
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

      } else if (ext === "xlsx") {
        let currentBuffer = f.buffer;

        // Step 1: Standard cleaning
        const cleaned = await cleanXLSX(currentBuffer, cleaningOptions);
        currentBuffer = cleaned.outBuffer;

        // 🆕 Step 2: Remove sensitive data
        if (sensitiveDataToRemove.length > 0) {
          const sensitiveResult = await removeSensitiveDataFromXLSX(currentBuffer, sensitiveDataToRemove);
          currentBuffer = sensitiveResult.outBuffer;
          extraRemovals.sensitiveDataRemoved = sensitiveResult.stats.removed;
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

      } else {
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
    }

    const zipName =
      files.length === 1
        ? `${baseName(files[0].originalname)} cleaned.zip`
        : "qualion_doc_cleaned.zip";

    sendZip(res, zip, zipName);
  } catch (e) {
    console.error("CLEAN ERROR", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ===================================================================
// POST /rephrase - VERSION 2.8.0
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
        console.warn(`[REPHRASE] Analysis failed:`, analysisError.message);
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

    const zipName =
      files.length === 1
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
  console.log(`✅ Qualion-Doc Backend v2.8.0 listening on port ${PORT}`);
  console.log(`   Endpoints: GET /health, POST /analyze, POST /clean, POST /rephrase`);
  console.log(`   Features: sensitive-data-removal, hidden-content-removal`);
});
