// server.js - VERSION 2.5 avec scoreImpacts pour before/after détaillé
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
import { correctDOCXText, correctPPTXText } from "./lib/officeCorrect.js";
import { buildReportHtmlDetailed, buildReportData } from "./lib/report.js";
import { extractPdfText, filterExtractedLines } from "./lib/pdfTools.js";
import { createDocxFromText } from "./lib/docxWriter.js";
import { aiCorrectText } from "./lib/ai.js";

// Import de documentAnalyzer
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

// ============================================================
// Calcul du score de risque basé sur les détections
// Score 0-100 où 100 = document sûr, 0 = risque critique
// Retourne aussi le breakdown des pénalités par catégorie
// ============================================================
function calculateRiskScore(summary, detections = null) {
  let score = 100;
  const breakdown = {};
  
  // 1. Pénalités par sévérité classifiée (si disponible)
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
  
  // 2. Si detections fourni, calculer les pénalités par type
  if (detections) {
    // Sensitive Data = CRITICAL (-25 par item, max -50)
    const sensitiveCount = detections.sensitiveData?.length || 0;
    if (sensitiveCount > 0) {
      const penalty = Math.min(sensitiveCount * 25, 50);
      score -= penalty;
      breakdown.sensitiveData = penalty;
    }
    
    // Macros = HIGH (-15 par item, max -30)
    const macrosCount = detections.macros?.length || 0;
    if (macrosCount > 0) {
      const penalty = Math.min(macrosCount * 15, 30);
      score -= penalty;
      breakdown.macros = penalty;
    }
    
    // Hidden Content = MEDIUM (-8 par item, max -24)
    const hiddenCount = (detections.hiddenContent?.length || 0) + 
                        (detections.hiddenSheets?.length || 0);
    if (hiddenCount > 0) {
      const penalty = Math.min(hiddenCount * 8, 24);
      score -= penalty;
      breakdown.hiddenContent = penalty;
    }
    
    // Comments = LOW (-3 par item, max -15)
    const commentsCount = detections.comments?.length || 0;
    if (commentsCount > 0) {
      const penalty = Math.min(commentsCount * 3, 15);
      score -= penalty;
      breakdown.comments = penalty;
    }
    
    // Track Changes = LOW (-3 par item, max -15)
    const trackChangesCount = detections.trackChanges?.length || 0;
    if (trackChangesCount > 0) {
      const penalty = Math.min(trackChangesCount * 3, 15);
      score -= penalty;
      breakdown.trackChanges = penalty;
    }
    
    // Metadata = LOW (-2 par item, max -10)
    const metadataCount = detections.metadata?.length || 0;
    if (metadataCount > 0) {
      const penalty = Math.min(metadataCount * 2, 10);
      score -= penalty;
      breakdown.metadata = penalty;
    }
    
    // Embedded Objects = MEDIUM (-5 par item, max -15)
    const embeddedCount = detections.embeddedObjects?.length || 0;
    if (embeddedCount > 0) {
      const penalty = Math.min(embeddedCount * 5, 15);
      score -= penalty;
      breakdown.embeddedObjects = penalty;
    }
    
    // Spelling Errors = LOW (-1 par item, max -10)
    const spellingCount = detections.spellingErrors?.length || 0;
    if (spellingCount > 0) {
      const penalty = Math.min(spellingCount * 1, 10);
      score -= penalty;
      breakdown.spellingGrammar = penalty;
    }
    
    // Broken Links = MEDIUM (-4 par item, max -12)
    const brokenLinksCount = detections.brokenLinks?.length || 0;
    if (brokenLinksCount > 0) {
      const penalty = Math.min(brokenLinksCount * 4, 12);
      score -= penalty;
      breakdown.brokenLinks = penalty;
    }
    
    // Compliance Risks = HIGH (-12 par item, max -36)
    const complianceCount = detections.complianceRisks?.length || 0;
    if (complianceCount > 0) {
      const penalty = Math.min(complianceCount * 12, 36);
      score -= penalty;
      breakdown.complianceRisks = penalty;
    }
  } else {
    // 3. Fallback: Si pas de detections, utiliser totalIssues
    const classifiedIssues = (summary.critical || 0) + (summary.high || 0) + 
                            (summary.medium || 0) + (summary.low || 0);
    const unclassifiedIssues = Math.max(0, (summary.totalIssues || 0) - classifiedIssues);
    
    if (unclassifiedIssues > 0) {
      const penalty = unclassifiedIssues * 5;
      score -= penalty;
      breakdown.unclassified = penalty;
    }
  }
  
  // 4. Pénalité supplémentaire si beaucoup d'issues totales
  if (summary.totalIssues > 10) {
    const penalty = (summary.totalIssues - 10) * 2;
    score -= penalty;
    breakdown.volumePenalty = penalty;
  }
  
  const finalScore = Math.max(0, Math.min(100, score));
  
  console.log(`[RISK SCORE] Calculated: ${finalScore} (total issues: ${summary.totalIssues}, breakdown:`, breakdown, ')');
  
  return { score: finalScore, breakdown };
}

// ============================================================
// Calcul du score APRÈS nettoyage avec scoreImpacts détaillés
// ============================================================
function calculateAfterScore(beforeScore, cleaningStats, correctionStats, riskBreakdown = {}) {
  let improvement = 0;
  const scoreImpacts = {};
  
  // Récupérer les points perdus par catégorie du breakdown
  // Et calculer combien on récupère via le nettoyage
  
  // Metadata supprimées → récupérer jusqu'à 100% des points perdus
  if (cleaningStats?.metaRemoved > 0 && riskBreakdown.metadata) {
    const impact = Math.min(cleaningStats.metaRemoved * 2, riskBreakdown.metadata);
    improvement += impact;
    scoreImpacts.metadata = impact;
  } else if (cleaningStats?.metaRemoved > 0) {
    const impact = Math.min(cleaningStats.metaRemoved * 2, 10);
    improvement += impact;
    scoreImpacts.metadata = impact;
  }
  
  // Comments supprimés
  if (cleaningStats?.commentsXmlRemoved > 0 && riskBreakdown.comments) {
    const impact = Math.min(cleaningStats.commentsXmlRemoved * 3, riskBreakdown.comments);
    improvement += impact;
    scoreImpacts.comments = impact;
  } else if (cleaningStats?.commentsXmlRemoved > 0) {
    const impact = Math.min(cleaningStats.commentsXmlRemoved * 3, 15);
    improvement += impact;
    scoreImpacts.comments = impact;
  }
  
  // Track Changes acceptées
  const trackChangesTotal = (cleaningStats?.revisionsAccepted?.deletionsRemoved || 0) + 
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
  
  // Hidden content supprimé
  if (cleaningStats?.hiddenRemoved > 0 && riskBreakdown.hiddenContent) {
    const impact = Math.min(cleaningStats.hiddenRemoved * 8, riskBreakdown.hiddenContent);
    improvement += impact;
    scoreImpacts.hiddenContent = impact;
  } else if (cleaningStats?.hiddenRemoved > 0) {
    const impact = Math.min(cleaningStats.hiddenRemoved * 8, 24);
    improvement += impact;
    scoreImpacts.hiddenContent = impact;
  }
  
  // Macros supprimés
  if (cleaningStats?.macrosRemoved > 0 && riskBreakdown.macros) {
    const impact = Math.min(cleaningStats.macrosRemoved * 15, riskBreakdown.macros);
    improvement += impact;
    scoreImpacts.macros = impact;
  } else if (cleaningStats?.macrosRemoved > 0) {
    const impact = Math.min(cleaningStats.macrosRemoved * 15, 30);
    improvement += impact;
    scoreImpacts.macros = impact;
  }
  
  // Embedded objects / media supprimés
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
  
  // Corrections orthographiques appliquées
  if (correctionStats?.changedTextNodes > 0 && riskBreakdown.spellingGrammar) {
    const impact = Math.min(correctionStats.changedTextNodes * 1, riskBreakdown.spellingGrammar);
    improvement += impact;
    scoreImpacts.spellingGrammar = impact;
  } else if (correctionStats?.changedTextNodes > 0) {
    const impact = Math.min(correctionStats.changedTextNodes * 1, 10);
    improvement += impact;
    scoreImpacts.spellingGrammar = impact;
  }
  
  // Le score après ne peut pas dépasser 100
  const afterScore = Math.min(100, beforeScore + improvement);
  
  console.log(`[AFTER SCORE] Before: ${beforeScore}, After: ${afterScore}, Improvement: +${improvement}, Impacts:`, scoreImpacts);
  
  return { score: afterScore, scoreImpacts, improvement };
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

// Helper: Déterminer le niveau de risque à partir du score
function getRiskLevel(score) {
  if (score >= 90) return 'safe';
  if (score >= 70) return 'low';
  if (score >= 50) return 'medium';
  if (score >= 25) return 'high';
  return 'critical';
}

// Helper: Ajouter les deux rapports (HTML + JSON) au ZIP
function addReportsToZip(zip, single, base, reportParams) {
  // Générer le rapport HTML (legacy)
  const reportHtml = buildReportHtmlDetailed(reportParams);
  zip.addFile(outName(single, base, "report.html"), Buffer.from(reportHtml, "utf8"));
  
  // Générer le rapport JSON (pour PremiumSecurityReport)
  const reportJson = buildReportData(reportParams);
  zip.addFile(outName(single, base, "report.json"), Buffer.from(JSON.stringify(reportJson, null, 2), "utf8"));
  
  console.log(`[REPORT] Generated HTML + JSON reports for ${reportParams.filename} (before: ${reportParams.beforeRiskScore}, after: ${reportParams.afterRiskScore})`);
}

// ---------- Health ----------
app.get("/health", (_, res) =>
  res.json({ 
    ok: true, 
    service: "Qualion-Doc Backend", 
    version: "2.5",
    endpoints: ["/analyze", "/clean", "/rephrase"],
    features: ["spelling-correction", "premium-json-report", "accurate-risk-score", "score-impacts"],
    time: new Date().toISOString() 
  })
);

// ===================================================================
// POST /analyze - VERSION 2.5
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
    
    // Passer les detections pour un score précis avec breakdown
    const { score: riskScore, breakdown } = calculateRiskScore(summary, detections);
    
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
        riskLevel: getRiskLevel(riskScore),
        riskBreakdown: breakdown,
        recommendations: generateRecommendations(detections, summary)
      },
      processingTime: Date.now() - startTime
    };
    
    console.log(`[ANALYZE] Complete in ${result.processingTime}ms - ${summary.totalIssues} issues found, risk score: ${riskScore} (${getRiskLevel(riskScore)})`);
    
    res.json(result);
    
  } catch (e) {
    console.error("[ANALYZE ERROR]", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ===================================================================
// POST /clean  → VERSION 2.5 avec scoreImpacts
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
      // ÉTAPE 1: Analyser le document AVANT le nettoyage
      // ============================================================
      let analysisResult = null;
      let spellingErrors = [];
      let beforeRiskScore = 100;
      let riskBreakdown = {};
      let detections = null;
      
      try {
        const fileType = getMimeFromExt(ext);
        detections = await analyzeDocument(f.buffer, fileType);
        const summary = calculateSummary(detections);
        
        // Passer les detections pour un score précis avec breakdown
        const riskResult = calculateRiskScore(summary, detections);
        beforeRiskScore = riskResult.score;
        riskBreakdown = riskResult.breakdown;
        
        spellingErrors = detections.spellingErrors || [];
        
        analysisResult = {
          detections,
          summary: {
            ...summary,
            riskScore: beforeRiskScore,
            beforeRiskScore: beforeRiskScore,
            riskBreakdown: riskBreakdown,
            riskLevel: getRiskLevel(beforeRiskScore),
            recommendations: generateRecommendations(detections, summary)
          }
        };
        
        console.log(`[CLEAN] Analysis found ${spellingErrors.length} spelling errors, before score: ${beforeRiskScore} (${getRiskLevel(beforeRiskScore)})`);
      } catch (analysisError) {
        console.warn(`[CLEAN] Analysis failed, continuing without:`, analysisError.message);
      }

      // ============================================================
      // TRAITEMENT DOCX
      // ============================================================
      if (ext === "docx") {
        const cleaned = await cleanDOCX(f.buffer, { drawPolicy, ...cleaningOptions });
        
        let finalBuffer = cleaned.outBuffer;
        let correctionStats = null;
        
        if (cleaningOptions.correctSpelling) {
          const corrected = await correctDOCXText(cleaned.outBuffer, aiCorrectText, {
            spellingErrors: spellingErrors
          });
          finalBuffer = corrected.outBuffer;
          correctionStats = corrected.stats;
          
          console.log(`[CLEAN DOCX] Applied ${correctionStats.changedTextNodes} corrections`);
        }

        zip.addFile(outName(single, base, "cleaned.docx"), finalBuffer);

        // Calculer le score après nettoyage avec scoreImpacts
        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats, riskBreakdown);
        
        // Ajouter les deux rapports (HTML + JSON)
        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { drawPolicy, ...cleaningOptions },
          cleaning: cleaned.stats,
          correction: correctionStats,
          analysis: analysisResult,
          spellingErrors: spellingErrors,
          beforeRiskScore: beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts
        });
        
      // ============================================================
      // TRAITEMENT PPTX
      // ============================================================
      } else if (ext === "pptx") {
        const cleaned = await cleanPPTX(f.buffer, { drawPolicy, ...cleaningOptions });
        
        let finalBuffer = cleaned.outBuffer;
        let correctionStats = null;
        
        if (cleaningOptions.correctSpelling) {
          const corrected = await correctPPTXText(cleaned.outBuffer, aiCorrectText, {
            spellingErrors: spellingErrors
          });
          finalBuffer = corrected.outBuffer;
          correctionStats = corrected.stats;
          
          console.log(`[CLEAN PPTX] Applied ${correctionStats.changedTextNodes} corrections`);
        }

        zip.addFile(outName(single, base, "cleaned.pptx"), finalBuffer);

        // Calculer le score après nettoyage avec scoreImpacts
        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats, riskBreakdown);

        // Ajouter les deux rapports (HTML + JSON)
        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { drawPolicy, ...cleaningOptions },
          cleaning: cleaned.stats,
          correction: correctionStats,
          analysis: analysisResult,
          spellingErrors: spellingErrors,
          beforeRiskScore: beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts
        });
        
      // ============================================================
      // TRAITEMENT PDF
      // ============================================================
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

        // Calculer le score après nettoyage avec scoreImpacts
        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats, riskBreakdown);

        // Ajouter les deux rapports (HTML + JSON)
        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { pdfMode, ...cleaningOptions },
          cleaning: cleaned.stats,
          correction: correctionStats,
          analysis: analysisResult,
          spellingErrors: spellingErrors,
          beforeRiskScore: beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts
        });
        
      // ============================================================
      // TRAITEMENT XLSX
      // ============================================================
      } else if (ext === "xlsx") {
        zip.addFile(outName(single, base, f.originalname), f.buffer);
        
        const afterResult = calculateAfterScore(beforeRiskScore, {}, null, riskBreakdown);
        
        // Ajouter les deux rapports (HTML + JSON)
        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: cleaningOptions,
          cleaning: {},
          correction: null,
          analysis: analysisResult,
          spellingErrors: spellingErrors,
          beforeRiskScore: beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts
        });
        
      // ============================================================
      // AUTRES TYPES
      // ============================================================
      } else {
        zip.addFile(outName(single, base, f.originalname), f.buffer);
        
        // Ajouter les deux rapports (HTML + JSON)
        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: {},
          cleaning: {},
          correction: null,
          analysis: null,
          spellingErrors: [],
          beforeRiskScore: 100,
          afterRiskScore: 100,
          scoreImpacts: {}
        });
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
// POST /rephrase - VERSION 2.5 avec scoreImpacts
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
      let spellingErrors = [];
      let beforeRiskScore = 100;
      let riskBreakdown = {};
      let detections = null;
      
      try {
        const fileType = getMimeFromExt(ext);
        detections = await analyzeDocument(f.buffer, fileType);
        const summary = calculateSummary(detections);
        
        // Passer les detections pour un score précis avec breakdown
        const riskResult = calculateRiskScore(summary, detections);
        beforeRiskScore = riskResult.score;
        riskBreakdown = riskResult.breakdown;
        
        spellingErrors = detections.spellingErrors || [];
        
        analysisResult = {
          detections,
          summary: {
            ...summary,
            riskScore: beforeRiskScore,
            beforeRiskScore: beforeRiskScore,
            riskBreakdown: riskBreakdown,
            riskLevel: getRiskLevel(beforeRiskScore),
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
          spellingErrors: spellingErrors
        });

        zip.addFile(outName(single, base, "rephrased.docx"), rephrased.outBuffer);

        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, rephrased.stats, riskBreakdown);

        // Ajouter les deux rapports (HTML + JSON)
        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { drawPolicy, mode: "rephrase" },
          cleaning: cleaned.stats,
          correction: rephrased.stats,
          analysis: analysisResult,
          spellingErrors: spellingErrors,
          beforeRiskScore: beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts
        });
        
      } else if (ext === "pptx") {
        const cleaned = await cleanPPTX(f.buffer, { drawPolicy });
        
        const rephrased = await correctPPTXText(cleaned.outBuffer, aiCorrectText, {
          mode: "rephrase",
          spellingErrors: spellingErrors
        });

        zip.addFile(outName(single, base, "rephrased.pptx"), rephrased.outBuffer);

        const afterResult = calculateAfterScore(beforeRiskScore, cleaned.stats, rephrased.stats, riskBreakdown);

        // Ajouter les deux rapports (HTML + JSON)
        addReportsToZip(zip, single, base, {
          filename: f.originalname,
          ext,
          policy: { drawPolicy, mode: "rephrase" },
          cleaning: cleaned.stats,
          correction: rephrased.stats,
          analysis: analysisResult,
          spellingErrors: spellingErrors,
          beforeRiskScore: beforeRiskScore,
          afterRiskScore: afterResult.score,
          scoreImpacts: afterResult.scoreImpacts
        });
        
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
  console.log(`✅ Qualion-Doc Backend v2.5 listening on port ${PORT}`);
  console.log(`   Endpoints: GET /health, POST /analyze, POST /clean, POST /rephrase`);
  console.log(`   Features: Spelling corrections, Premium JSON reports, Score impacts breakdown`);
});
