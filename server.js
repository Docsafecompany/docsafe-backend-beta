// server.js - VERSION 2.4 avec score de risque corrigé
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
// CORRECTION: Calcul du score de risque basé sur les détections
// Score 0-100 où 100 = document sûr, 0 = risque critique
// ============================================================
function calculateRiskScore(summary, detections = null) {
  let score = 100;
  
  // 1. Pénalités par sévérité classifiée (si disponible)
  score -= (summary.critical || 0) * 25;
  score -= (summary.high || 0) * 10;
  score -= (summary.medium || 0) * 5;
  score -= (summary.low || 0) * 2;
  
  // 2. NOUVEAU: Si detections fourni, calculer les pénalités par type
  if (detections) {
    // Sensitive Data = CRITICAL (-25 par item, max -50)
    const sensitiveCount = detections.sensitiveData?.length || 0;
    if (sensitiveCount > 0) {
      score -= Math.min(sensitiveCount * 25, 50);
    }
    
    // Macros = HIGH (-15 par item, max -30)
    const macrosCount = detections.macros?.length || 0;
    if (macrosCount > 0) {
      score -= Math.min(macrosCount * 15, 30);
    }
    
    // Hidden Content = MEDIUM (-8 par item, max -24)
    const hiddenCount = (detections.hiddenContent?.length || 0) + 
                        (detections.hiddenSheets?.length || 0);
    if (hiddenCount > 0) {
      score -= Math.min(hiddenCount * 8, 24);
    }
    
    // Comments = LOW (-3 par item, max -15)
    const commentsCount = detections.comments?.length || 0;
    if (commentsCount > 0) {
      score -= Math.min(commentsCount * 3, 15);
    }
    
    // Track Changes = LOW (-3 par item, max -15)
    const trackChangesCount = detections.trackChanges?.length || 0;
    if (trackChangesCount > 0) {
      score -= Math.min(trackChangesCount * 3, 15);
    }
    
    // Metadata = LOW (-2 par item, max -10)
    const metadataCount = detections.metadata?.length || 0;
    if (metadataCount > 0) {
      score -= Math.min(metadataCount * 2, 10);
    }
    
    // Embedded Objects = MEDIUM (-5 par item, max -15)
    const embeddedCount = detections.embeddedObjects?.length || 0;
    if (embeddedCount > 0) {
      score -= Math.min(embeddedCount * 5, 15);
    }
    
    // Spelling Errors = LOW (-1 par item, max -10)
    const spellingCount = detections.spellingErrors?.length || 0;
    if (spellingCount > 0) {
      score -= Math.min(spellingCount * 1, 10);
    }
    
    // Broken Links = MEDIUM (-4 par item, max -12)
    const brokenLinksCount = detections.brokenLinks?.length || 0;
    if (brokenLinksCount > 0) {
      score -= Math.min(brokenLinksCount * 4, 12);
    }
    
    // Compliance Risks = HIGH (-12 par item, max -36)
    const complianceCount = detections.complianceRisks?.length || 0;
    if (complianceCount > 0) {
      score -= Math.min(complianceCount * 12, 36);
    }
  } else {
    // 3. Fallback: Si pas de detections, utiliser totalIssues
    const classifiedIssues = (summary.critical || 0) + (summary.high || 0) + 
                            (summary.medium || 0) + (summary.low || 0);
    const unclassifiedIssues = Math.max(0, (summary.totalIssues || 0) - classifiedIssues);
    
    // Chaque issue non classifiée = -5 points (considérée medium par défaut)
    score -= unclassifiedIssues * 5;
  }
  
  // 4. Pénalité supplémentaire si beaucoup d'issues totales
  if (summary.totalIssues > 10) {
    score -= (summary.totalIssues - 10) * 2;
  }
  
  const finalScore = Math.max(0, Math.min(100, score));
  
  console.log(`[RISK SCORE] Calculated: ${finalScore} (total issues: ${summary.totalIssues})`);
  
  return finalScore;
}

// Helper: Calcul du score APRÈS nettoyage (toujours meilleur)
function calculateAfterScore(beforeScore, cleaningStats, correctionStats) {
  let improvement = 0;
  
  // Amélioration basée sur le nettoyage effectué
  if (cleaningStats) {
    if (cleaningStats.metaRemoved > 0) improvement += 5;
    if (cleaningStats.commentsXmlRemoved > 0) improvement += 10;
    if (cleaningStats.revisionsAccepted?.total > 0) improvement += 10;
    if (cleaningStats.hiddenRemoved > 0) improvement += 8;
    if (cleaningStats.macrosRemoved > 0) improvement += 15;
    if (cleaningStats.mediaDeleted > 0) improvement += 5;
  }
  
  // Amélioration basée sur les corrections de texte
  if (correctionStats?.changedTextNodes > 0) {
    improvement += Math.min(correctionStats.changedTextNodes * 2, 15);
  }
  
  // Le score après ne peut pas dépasser 100
  return Math.min(100, beforeScore + improvement);
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
  
  console.log(`[REPORT] Generated HTML + JSON reports for ${reportParams.filename}`);
}

// ---------- Health ----------
app.get("/health", (_, res) =>
  res.json({ 
    ok: true, 
    service: "Qualion-Doc Backend", 
    version: "2.4",
    endpoints: ["/analyze", "/clean", "/rephrase"],
    features: ["spelling-correction", "premium-json-report", "accurate-risk-score"],
    time: new Date().toISOString() 
  })
);

// ===================================================================
// POST /analyze - VERSION 2.4 avec score corrigé
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
    
    // ✅ CORRECTION: Passer les detections pour un score précis
    const riskScore = calculateRiskScore(summary, detections);
    
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
// POST /clean  → VERSION 2.4 avec score corrigé
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
      let detections = null;
      
      try {
        const fileType = getMimeFromExt(ext);
        detections = await analyzeDocument(f.buffer, fileType);
        const summary = calculateSummary(detections);
        
        // ✅ CORRECTION: Passer les detections pour un score précis
        beforeRiskScore = calculateRiskScore(summary, detections);
        
        spellingErrors = detections.spellingErrors || [];
        
        analysisResult = {
          detections,
          summary: {
            ...summary,
            riskScore: beforeRiskScore,
            beforeRiskScore: beforeRiskScore,
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

        // Calculer le score après nettoyage
        const afterRiskScore = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats);
        
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
          afterRiskScore: afterRiskScore
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

        // Calculer le score après nettoyage
        const afterRiskScore = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats);

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
          afterRiskScore: afterRiskScore
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

        // Calculer le score après nettoyage
        const afterRiskScore = calculateAfterScore(beforeRiskScore, cleaned.stats, correctionStats);

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
          afterRiskScore: afterRiskScore
        });
        
      // ============================================================
      // TRAITEMENT XLSX
      // ============================================================
      } else if (ext === "xlsx") {
        zip.addFile(outName(single, base, f.originalname), f.buffer);
        
        const afterRiskScore = calculateAfterScore(beforeRiskScore, {}, null);
        
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
          afterRiskScore: afterRiskScore
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
          afterRiskScore: 100
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
// POST /rephrase - VERSION 2.4 avec score corrigé
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
      let detections = null;
      
      try {
        const fileType = getMimeFromExt(ext);
        detections = await analyzeDocument(f.buffer, fileType);
        const summary = calculateSummary(detections);
        
        // ✅ CORRECTION: Passer les detections pour un score précis
        beforeRiskScore = calculateRiskScore(summary, detections);
        
        spellingErrors = detections.spellingErrors || [];
        
        analysisResult = {
          detections,
          summary: {
            ...summary,
            riskScore: beforeRiskScore,
            beforeRiskScore: beforeRiskScore,
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

        const afterRiskScore = calculateAfterScore(beforeRiskScore, cleaned.stats, rephrased.stats);

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
          afterRiskScore: afterRiskScore
        });
        
      } else if (ext === "pptx") {
        const cleaned = await cleanPPTX(f.buffer, { drawPolicy });
        
        const rephrased = await correctPPTXText(cleaned.outBuffer, aiCorrectText, {
          mode: "rephrase",
          spellingErrors: spellingErrors
        });

        zip.addFile(outName(single, base, "rephrased.pptx"), rephrased.outBuffer);

        const afterRiskScore = calculateAfterScore(beforeRiskScore, cleaned.stats, rephrased.stats);

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
          afterRiskScore: afterRiskScore
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
  console.log(`✅ Qualion-Doc Backend v2.4 listening on port ${PORT}`);
  console.log(`   Endpoints: GET /health, POST /analyze, POST /clean, POST /rephrase`);
  console.log(`   Features: Spelling corrections, Premium JSON reports, Accurate risk scoring`);
});
