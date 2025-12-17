// server.js - VERSION 3.1.0
// ✅ Compatible documentAnalyzer v3.2.2
// ✅ Adds Universal Risk Model (4 surfaces) => riskObjects[] + riskSummary (Client-Ready gate)
// ✅ Does NOT break existing detections/summary payloads
// ✅ Keeps selective cleaning behavior (non-empty lists only)

import express from "express";
import cors from "cors";
import multer from "multer";
import AdmZip from "adm-zip";
import path from "path";
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

// Import docStats
import { extractDocStats } from "./lib/docStats.js";

// Import sensitive data cleaner
import {
  removeSensitiveDataFromDOCX,
  removeSensitiveDataFromPPTX,
  removeSensitiveDataFromXLSX,
  removeHiddenContentFromDOCX,
  removeHiddenContentFromPPTX,
  removeVisualObjectsFromPPTX,
} from "./lib/sensitiveDataCleaner.js";

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

const outName = (single, base, name) => (single ? name : `${base}_${name}`);
const baseName = (filename = "document") => filename.replace(/\.[^.]+$/, "");

function sendZip(res, zip, zipName) {
  res.setHeader("Content-Type", "application/zip");
  res.setHeader("Content-Disposition", `attachment; filename="${zipName}"`);
  res.send(zip.toBuffer());
}

// ✅ Safe wrapper for doc stats (never breaks pipeline)
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

// ✅ Robust JSON parse that accepts stringified JSON array or already-parsed arrays
function safeJsonArray(input, fallback = []) {
  if (input === undefined || input === null) return fallback;
  if (Array.isArray(input)) return input;
  if (typeof input !== "string") return fallback;

  const s = input.trim();
  if (!s) return fallback;

  try {
    const parsed = JSON.parse(s);
    return Array.isArray(parsed) ? parsed : fallback;
  } catch (e) {
    console.warn("[JSON PARSE] Failed:", e?.message);
    return fallback;
  }
}

// ============================================================
// Risk score (legacy score kept)
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

    // ✅ include excelHiddenData in hidden penalty (XLSX)
    const hiddenCount =
      (detections.hiddenContent?.length || 0) +
      (detections.hiddenSheets?.length || 0) +
      (detections.excelHiddenData?.length || 0);

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
      (summary.critical || 0) + (summary.high || 0) + (summary.medium || 0) + (summary.low || 0);
    const unclassifiedIssues = Math.max(0, (summary.totalIssues || 0) - classifiedIssues);

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
  console.log(`[RISK SCORE] Calculated: ${finalScore} (total issues: ${summary.totalIssues})`, breakdown);
  return { score: finalScore, breakdown };
}

// ============================================================
// Universal Risk Model (NEW) - 4 Surfaces + 4 Executive Categories
// ============================================================

// 4 Exposure Surfaces
const SURFACE = {
  VISIBLE: "visible",
  HIDDEN: "hidden",
  LOGIC: "logic",
  METADATA: "metadata",
};

// 4 Executive Risk Categories
const EXEC_CAT = {
  MARGIN: "margin",
  DELIVERY: "delivery",
  NEGOTIATION: "negotiation",
  CREDIBILITY: "credibility",
};

function normalizeSeverity(sev) {
  const s = String(sev || "low").toLowerCase();
  if (["critical", "high", "medium", "low"].includes(s)) return s;
  return "low";
}

function severityRank(sev) {
  const s = normalizeSeverity(sev);
  return s === "critical" ? 4 : s === "high" ? 3 : s === "medium" ? 2 : 1;
}

// Simple deterministic points (per spec)
function pointsForRiskObject(ro) {
  let pts = 0;

  // Hidden but accessible
  if (ro.surface === SURFACE.HIDDEN) pts += 3;

  // Generated/derived logic
  if (ro.surface === SURFACE.LOGIC) pts += 2;

  // Metadata revealing intent/org
  if (ro.surface === SURFACE.METADATA) pts += 2;

  // Cross-object dependency hint (best-effort based on detection types)
  if (ro.ruleId === "XLSX_FORMULA" || ro.ruleId === "XLSX_EXTERNAL_REF" || ro.ruleId === "BROKEN_LINK") pts += 2;

  // Financial proximity
  if (ro.category === EXEC_CAT.MARGIN) pts += 2;

  return pts;
}

// Map detection-type => executive category (deterministic)
function categorizeDetection(detType) {
  const t = String(detType || "").toLowerCase();

  // explicit financial signals
  if (["price", "iban", "credit_card"].includes(t)) return EXEC_CAT.MARGIN;

  // strong negotiation leakage candidates
  if (["project_code", "internal_url", "file_path", "server_path", "ip_address"].includes(t)) return EXEC_CAT.NEGOTIATION;

  // personal data is more "negotiation/credibility" (client trust / compliance)
  if (["email", "phone", "ssn"].includes(t)) return EXEC_CAT.CREDIBILITY;

  return EXEC_CAT.CREDIBILITY;
}

function mapDetectionsToRiskObjects(fileType, detections) {
  const riskObjects = [];
  const ext = String(fileType || "").toLowerCase();

  const pushRO = (ro) => {
    riskObjects.push({
      id: ro.id || `ro_${uuidv4()}`,
      surface: ro.surface, // visible|hidden|logic|metadata
      category: ro.category, // margin|delivery|negotiation|credibility
      severity: normalizeSeverity(ro.severity),
      fileType: ext,
      ruleId: ro.ruleId, // deterministic string
      reason: ro.reason, // short neutral sentence
      location: ro.location || "Document",
      fixability: ro.fixability || "manual", // auto-fix|manual|not-fixable
      source: ro.source || null, // { detectionGroup, detectionId }
      meta: ro.meta || {},
    });
  };

  // ---------------- METADATA => surface: metadata
  (detections.metadata || []).forEach((m) => {
    pushRO({
      id: `ro_meta_${m.id || uuidv4()}`,
      surface: SURFACE.METADATA,
      category: EXEC_CAT.NEGOTIATION,
      severity: m.severity || "medium",
      ruleId: "META_EXPOSURE",
      reason: "Metadata detected that may reveal internal organization or authorship.",
      location: m.location || "Document Properties",
      fixability: "auto-fix",
      source: { detectionGroup: "metadata", detectionId: m.id || null },
      meta: { key: m.key || m.type || null },
    });
  });

  // ---------------- COMMENTS => hidden
  (detections.comments || []).forEach((c) => {
    pushRO({
      id: `ro_comment_${c.id || uuidv4()}`,
      surface: SURFACE.HIDDEN,
      category: EXEC_CAT.CREDIBILITY,
      severity: c.severity || "medium",
      ruleId: c.type === "tracked_change" ? "TRACKED_CHANGE" : "COMMENT_EXPOSURE",
      reason:
        c.type === "tracked_change"
          ? "Tracked changes detected that may expose draft edits or internal review."
          : "Comments detected that may expose internal review discussions.",
      location: c.location || "Document",
      fixability: "auto-fix",
      source: { detectionGroup: "comments", detectionId: c.id || null },
      meta: { author: c.author || null, type: c.type || "comment" },
    });
  });

  // ---------------- TRACK CHANGES (legacy group) => hidden
  (detections.trackChanges || []).forEach((tc) => {
    pushRO({
      id: `ro_tc_${tc.id || uuidv4()}`,
      surface: SURFACE.HIDDEN,
      category: EXEC_CAT.CREDIBILITY,
      severity: tc.severity || "medium",
      ruleId: "TRACKED_CHANGE",
      reason: "Tracked changes detected that may expose draft edits or internal review.",
      location: tc.location || "Document",
      fixability: "auto-fix",
      source: { detectionGroup: "trackChanges", detectionId: tc.id || null },
      meta: { author: tc.author || null, changeType: tc.type || null },
    });
  });

  // ---------------- HIDDEN CONTENT => hidden
  (detections.hiddenContent || []).forEach((h) => {
    pushRO({
      id: `ro_hidden_${h.id || uuidv4()}`,
      surface: SURFACE.HIDDEN,
      category: EXEC_CAT.NEGOTIATION,
      severity: h.severity || "high",
      ruleId: "HIDDEN_CONTENT",
      reason: "Hidden content detected that may expose internal assumptions.",
      location: h.location || "Document",
      fixability: "auto-fix",
      source: { detectionGroup: "hiddenContent", detectionId: h.id || null },
      meta: { type: h.type || null },
    });
  });

  // ---------------- XLSX HIDDEN SHEETS (legacy group) => hidden
  (detections.hiddenSheets || []).forEach((hs) => {
    pushRO({
      id: `ro_hidden_sheet_${hs.id || uuidv4()}`,
      surface: SURFACE.HIDDEN,
      category: EXEC_CAT.MARGIN,
      severity: hs.severity || "high",
      ruleId: "XLSX_HIDDEN_SHEET",
      reason: "Hidden worksheet detected that may contain internal data.",
      location: hs.sheetName ? `Sheet: ${hs.sheetName}` : "Workbook",
      fixability: "manual",
      source: { detectionGroup: "hiddenSheets", detectionId: hs.id || null },
      meta: { sheetName: hs.sheetName || null, state: hs.type || null },
    });
  });

  // ---------------- XLSX FORMULAS (legacy group) => logic
  (detections.sensitiveFormulas || []).forEach((f) => {
    pushRO({
      id: `ro_formula_${f.id || uuidv4()}`,
      surface: SURFACE.LOGIC,
      category: EXEC_CAT.MARGIN,
      severity: f.severity || f.risk || "medium",
      ruleId: f.reason && String(f.reason).toLowerCase().includes("external") ? "XLSX_EXTERNAL_REF" : "XLSX_FORMULA",
      reason: "Spreadsheet formulas detected that may expose internal logic or dependencies.",
      location: f.sheet ? `Sheet: ${f.sheet}` : "Workbook",
      fixability: "manual",
      source: { detectionGroup: "sensitiveFormulas", detectionId: f.id || null },
      meta: { sheet: f.sheet || null },
    });
  });

  // ---------------- XLSX consolidated excelHiddenData => hidden/logic
  (detections.excelHiddenData || []).forEach((x) => {
    const type = String(x.type || "").toLowerCase();
    const isFormula = type.includes("formula");
    pushRO({
      id: `ro_excelhidden_${x.id || uuidv4()}`,
      surface: isFormula ? SURFACE.LOGIC : SURFACE.HIDDEN,
      category: isFormula ? EXEC_CAT.MARGIN : EXEC_CAT.MARGIN,
      severity: x.severity || "high",
      ruleId: isFormula ? "XLSX_FORMULA" : "XLSX_HIDDEN_DATA",
      reason: isFormula
        ? "Spreadsheet logic detected that may expose pricing or internal assumptions."
        : "Hidden spreadsheet data detected that may expose internal information.",
      location: x.location || "Workbook",
      fixability: "manual",
      source: { detectionGroup: "excelHiddenData", detectionId: x.id || null },
      meta: { name: x.name || null },
    });
  });

  // ---------------- EMBEDDED OBJECTS => logic (cross-object)
  (detections.embeddedObjects || []).forEach((e) => {
    pushRO({
      id: `ro_embed_${e.id || uuidv4()}`,
      surface: SURFACE.LOGIC,
      category: EXEC_CAT.NEGOTIATION,
      severity: e.severity || "medium",
      ruleId: "EMBEDDED_OBJECT",
      reason: "Embedded objects detected that may contain hidden data.",
      location: e.path || "Document",
      fixability: "manual",
      source: { detectionGroup: "embeddedObjects", detectionId: e.id || null },
      meta: { filename: e.filename || null },
    });
  });

  // ---------------- MACROS => logic (critical)
  (detections.macros || []).forEach((m) => {
    pushRO({
      id: `ro_macro_${m.id || uuidv4()}`,
      surface: SURFACE.LOGIC,
      category: EXEC_CAT.CREDIBILITY,
      severity: m.severity || "critical",
      ruleId: "MACRO_PRESENT",
      reason: "Executable macros detected that may create security and trust risks.",
      location: m.location || "VBA Project",
      fixability: "manual",
      source: { detectionGroup: "macros", detectionId: m.id || null },
      meta: { name: m.name || null },
    });
  });

  // ---------------- SENSITIVE DATA => visible
  (detections.sensitiveData || []).forEach((s) => {
    const cat = categorizeDetection(s.type);
    pushRO({
      id: `ro_sensitive_${s.id || uuidv4()}`,
      surface: SURFACE.VISIBLE,
      category: cat === EXEC_CAT.CREDIBILITY ? EXEC_CAT.NEGOTIATION : cat, // pragmatic default
      severity: s.severity || "high",
      ruleId: "VISIBLE_SENSITIVE",
      reason: "Visible sensitive information detected that may not be necessary for client understanding.",
      location: s.location || "Document",
      fixability: "manual",
      source: { detectionGroup: "sensitiveData", detectionId: s.id || null },
      meta: { type: s.type || null, category: s.category || null },
    });
  });

  // ---------------- SPELLING / DRAFT signals => credibility
  (detections.spellingErrors || []).forEach((sp) => {
    pushRO({
      id: `ro_spell_${sp.id || uuidv4()}`,
      surface: SURFACE.VISIBLE,
      category: EXEC_CAT.CREDIBILITY,
      severity: sp.severity || "low",
      ruleId: "CREDIBILITY_SPELLING",
      reason: "Spelling or grammar issues detected that may affect professional credibility.",
      location: sp.location || "Document",
      fixability: "auto-fix",
      source: { detectionGroup: "spellingErrors", detectionId: sp.id || null },
      meta: {},
    });
  });

  // ---------------- BROKEN LINKS => logic (dependency risk)
  (detections.brokenLinks || []).forEach((b) => {
    pushRO({
      id: `ro_link_${b.id || uuidv4()}`,
      surface: SURFACE.LOGIC,
      category: EXEC_CAT.CREDIBILITY,
      severity: b.severity || "low",
      ruleId: "BROKEN_LINK",
      reason: "Links detected that may not be accessible to external recipients.",
      location: b.location || "Document",
      fixability: "manual",
      source: { detectionGroup: "brokenLinks", detectionId: b.id || null },
      meta: { type: b.type || null },
    });
  });

  // ---------------- COMPLIANCE RISKS => credibility
  (detections.complianceRisks || []).forEach((c) => {
    pushRO({
      id: `ro_compliance_${c.id || uuidv4()}`,
      surface: SURFACE.VISIBLE,
      category: EXEC_CAT.CREDIBILITY,
      severity: c.severity || "high",
      ruleId: "COMPLIANCE_SIGNAL",
      reason: "Compliance-related signals detected that require review before external sharing.",
      location: c.location || "Document",
      fixability: "manual",
      source: { detectionGroup: "complianceRisks", detectionId: c.id || null },
      meta: { type: c.type || null },
    });
  });

  // ---------------- VISUAL OBJECTS (PPT shapes/textboxes etc.) => hidden/visible mix; we tag hidden if "covering" or missing alt
  (detections.visualObjects || []).forEach((v) => {
    const t = String(v.type || "").toLowerCase();
    const surface = t.includes("cover") ? SURFACE.HIDDEN : SURFACE.VISIBLE;

    pushRO({
      id: `ro_visual_${v.id || uuidv4()}`,
      surface,
      category: EXEC_CAT.CREDIBILITY,
      severity: v.severity || "low",
      ruleId: "VISUAL_OBJECT",
      reason:
        surface === SURFACE.HIDDEN
          ? "Visual elements detected that may hide or obscure content."
          : "Visual elements detected that may require review before external sharing.",
      location: v.location || "Document",
      fixability: "manual",
      source: { detectionGroup: "visualObjects", detectionId: v.id || null },
      meta: { type: v.type || null },
    });
  });

  // ---------------- ORPHAN DATA => credibility (low)
  (detections.orphanData || []).forEach((o) => {
    pushRO({
      id: `ro_orphan_${o.id || uuidv4()}`,
      surface: SURFACE.VISIBLE,
      category: EXEC_CAT.CREDIBILITY,
      severity: o.severity || "low",
      ruleId: "ORPHAN_SIGNAL",
      reason: "Residual formatting or orphan signals detected that may require a quick review.",
      location: o.location || "Document",
      fixability: "manual",
      source: { detectionGroup: "orphanData", detectionId: o.id || null },
      meta: { type: o.type || null },
    });
  });

  // Delivery category (best effort)
  // If you later add explicit detections for delivery/commitment language, map them here.

  return riskObjects;
}

function summarizeRiskObjects(riskObjects) {
  const byCategory = {
    [EXEC_CAT.MARGIN]: { points: 0, count: 0, maxSeverity: "low" },
    [EXEC_CAT.DELIVERY]: { points: 0, count: 0, maxSeverity: "low" },
    [EXEC_CAT.NEGOTIATION]: { points: 0, count: 0, maxSeverity: "low" },
    [EXEC_CAT.CREDIBILITY]: { points: 0, count: 0, maxSeverity: "low" },
  };

  let anyCritical = false;
  let maxOverallRank = 1;

  for (const ro of riskObjects) {
    const cat = ro.category || EXEC_CAT.CREDIBILITY;
    const pts = pointsForRiskObject(ro);
    byCategory[cat].points += pts;
    byCategory[cat].count += 1;

    const sev = normalizeSeverity(ro.severity);
    if (sev === "critical") anyCritical = true;

    const r = severityRank(sev);
    if (r > severityRank(byCategory[cat].maxSeverity)) byCategory[cat].maxSeverity = sev;
    if (r > maxOverallRank) maxOverallRank = r;
  }

  const categorySeverityFromPoints = (pts, maxSev) => {
    // Thresholds per spec (0–3 low, 4–6 medium, 7+ high/critical)
    if (maxSev === "critical") return "critical";
    if (pts >= 7) return "high";
    if (pts >= 4) return "medium";
    return "low";
  };

  const perCat = {};
  let overallSeverity = "low";
  let overallRank = 1;

  for (const cat of Object.keys(byCategory)) {
    const pts = byCategory[cat].points;
    const maxSev = byCategory[cat].maxSeverity;
    const sev = categorySeverityFromPoints(pts, maxSev);

    perCat[cat] = {
      points: pts,
      count: byCategory[cat].count,
      severity: sev,
    };

    const r = severityRank(sev);
    if (r > overallRank) {
      overallRank = r;
      overallSeverity = sev;
    }
  }

  // Executive signals (only the 4 labels, no internals)
  const executiveSignals = [];
  if (perCat[EXEC_CAT.MARGIN].points > 0) executiveSignals.push("Margin Exposure");
  if (perCat[EXEC_CAT.DELIVERY].points > 0) executiveSignals.push("Delivery & Commitment Risk");
  if (perCat[EXEC_CAT.NEGOTIATION].points > 0) executiveSignals.push("Negotiation Leakage");
  if (perCat[EXEC_CAT.CREDIBILITY].points > 0) executiveSignals.push("Professional Credibility Risk");

  // Blocking issues list: high/critical first (short neutral sentence)
  const blocking = riskObjects
    .filter((ro) => ["high", "critical"].includes(normalizeSeverity(ro.severity)))
    .sort((a, b) => severityRank(b.severity) - severityRank(a.severity))
    .slice(0, 12)
    .map((ro) => ({
      id: ro.id,
      severity: normalizeSeverity(ro.severity),
      category: ro.category,
      fixability: ro.fixability,
      reason: ro.reason,
      surface: ro.surface,
      location: ro.location,
      ruleId: ro.ruleId,
    }));

  // Client-Ready gate
  // Spec: any critical => NO. Also if any category points >= 7 => NO (High).
  const anyHighCategory =
    perCat[EXEC_CAT.MARGIN].points >= 7 ||
    perCat[EXEC_CAT.DELIVERY].points >= 7 ||
    perCat[EXEC_CAT.NEGOTIATION].points >= 7 ||
    perCat[EXEC_CAT.CREDIBILITY].points >= 7;

  const clientReady = anyCritical || anyHighCategory ? "NO" : "YES";

  const overallSeverityLabel =
    overallSeverity === "critical"
      ? "Critical"
      : overallSeverity === "high"
      ? "High"
      : overallSeverity === "medium"
      ? "Medium"
      : "Low";

  return {
    clientReady,
    overallSeverity: overallSeverityLabel,
    executiveSignals,
    byCategory: {
      margin: perCat[EXEC_CAT.MARGIN],
      delivery: perCat[EXEC_CAT.DELIVERY],
      negotiation: perCat[EXEC_CAT.NEGOTIATION],
      credibility: perCat[EXEC_CAT.CREDIBILITY],
    },
    blockingIssues: blocking,
    totalRiskObjects: riskObjects.length,
  };
}

// ============================================================
// After score
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

  if (extraRemovals.sensitiveDataRemoved > 0 && riskBreakdown.sensitiveData) {
    const impact = Math.min(extraRemovals.sensitiveDataRemoved * 25, riskBreakdown.sensitiveData);
    improvement += impact;
    scoreImpacts.sensitiveData = impact;
  }

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
    recommendations.push("Remove document metadata to protect author and organization information.");
  if (detections.comments?.length > 0)
    recommendations.push(`Review and remove ${detections.comments.length} comment(s) before sharing externally.`);
  if (detections.trackChanges?.length > 0)
    recommendations.push("Accept or reject all tracked changes to finalize the document.");
  if (
    (detections.hiddenContent?.length || 0) > 0 ||
    (detections.hiddenSheets?.length || 0) > 0 ||
    (detections.excelHiddenData?.length || 0) > 0
  )
    recommendations.push("Remove hidden content that could expose confidential information.");
  if (detections.macros?.length > 0)
    recommendations.push("Remove macros for security - they can contain executable code.");
  if (detections.sensitiveData?.length > 0) {
    const types = [...new Set(detections.sensitiveData.map((d) => d.type))];
    recommendations.push(`Review sensitive data detected: ${types.join(", ")}.`);
  }
  if (detections.embeddedObjects?.length > 0)
    recommendations.push("Remove embedded objects that may contain hidden data.");
  if (detections.spellingErrors?.length > 0)
    recommendations.push(`${detections.spellingErrors.length} spelling/grammar issue(s) were detected.`);
  if (recommendations.length === 0)
    recommendations.push("Document appears clean. Minor review recommended before external sharing.");
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
  zip.addFile(outName(single, base, "report.json"), Buffer.from(JSON.stringify(reportJson, null, 2), "utf8"));
}

// ---------- Health ----------
app.get("/health", (_, res) =>
  res.json({
    ok: true,
    service: "Qualion-Doc Backend",
    version: "3.1.0",
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
      "excelHiddenData-in-risk-score",
      "robust-json-array-parsing",
      "selective-mode-non-empty-only",
      "universal-riskObjects-4-surfaces",
      "client-ready-gate",
    ],
    time: new Date().toISOString(),
  })
);

// ===================================================================
// POST /analyze - VERSION 3.1.0
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

    // ✅ docStats as pipeline reference (analyzer stats only fallback)
    const docStats = await safeExtractDocStats(req.file.buffer, ext);
    const documentStats =
      docStats && (docStats.pages || docStats.slides || docStats.sheets || docStats.tables)
        ? docStats
        : analysisResult.documentStats || docStats;

    // ✅ NEW: universal riskObjects + executive summary
    const riskObjects = mapDetectionsToRiskObjects(ext, detections || {});
    const riskSummary = summarizeRiskObjects(riskObjects);

    res.json({
      documentId: uuidv4(),
      fileName: req.file.originalname,
      fileType: ext,
      fileSize: req.file.size,

      documentStats,

      // ✅ keep current detections structure (front compatibility)
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

      // ✅ keep legacy summary shape
      summary: {
        ...summary,
        riskScore,
        riskLevel: getRiskLevel(riskScore),
        riskBreakdown: breakdown,
        recommendations: generateRecommendations(detections),
      },

      // ✅ NEW OUTPUT (non-breaking additions)
      riskObjects,
      riskSummary,

      processingTime: Date.now() - startTime,
    });
  } catch (e) {
    console.error("[ANALYZE ERROR]", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ===================================================================
// POST /clean - VERSION 3.1.0
// ✅ Selective mode triggers ONLY when list is NON-EMPTY
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

    const approvedSpellingErrors = safeJsonArray(req.body.approvedSpellingErrors, []);

    const removeSensitiveDataRaw =
      safeJsonArray(req.body.removeSensitiveData, null) ?? safeJsonArray(req.body.sensitiveDataToClean, []);

    const hiddenContentToCleanRaw = safeJsonArray(req.body.hiddenContentToClean, []);
    const visualObjectsToCleanRaw = safeJsonArray(req.body.visualObjectsToClean, []);

    const hasSelectiveSensitive = Array.isArray(removeSensitiveDataRaw) && removeSensitiveDataRaw.length > 0;
    const hasSelectiveHidden = Array.isArray(hiddenContentToCleanRaw) && hiddenContentToCleanRaw.length > 0;
    const hasSelectiveVisual = Array.isArray(visualObjectsToCleanRaw) && visualObjectsToCleanRaw.length > 0;

    console.log(`[CLEAN] removeSensitiveData raw count: ${removeSensitiveDataRaw?.length || 0}`);
    console.log(`[CLEAN] hiddenContentToClean raw count: ${hiddenContentToCleanRaw?.length || 0}`);
    console.log(`[CLEAN] visualObjectsToClean raw count: ${visualObjectsToCleanRaw?.length || 0}`);
    console.log(
      `[CLEAN] Selective modes (NON-empty lists): sensitive=${hasSelectiveSensitive}, hidden=${hasSelectiveHidden}, visual=${hasSelectiveVisual}`
    );

    const single = files.length === 1;
    const zip = new AdmZip();

    for (const f of files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      console.log(`[CLEAN] Processing ${f.originalname} with options:`, cleaningOptions);

      // BEFORE stats
      const documentStatsBefore = await safeExtractDocStats(f.buffer, ext);

      // analysis (optional)
      let analysisResult = null;
      let spellingErrors = [];
      let beforeRiskScore = 100;
      let riskBreakdown = {};
      let detections = null;
      let summary = null;

      // ✅ NEW: universal riskObjects/riskSummary kept for report
      let riskObjects = [];
      let riskSummary = null;

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

        // ✅ NEW: universal mapping
        riskObjects = mapDetectionsToRiskObjects(ext, detections || {});
        riskSummary = summarizeRiskObjects(riskObjects);

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

            // ✅ include in report analysis summary
            riskSummary,
          },

          // ✅ include in report analysis payload
          riskObjects,
        };
      } catch (analysisError) {
        console.warn(`[CLEAN] Analysis failed, continuing without:`, analysisError?.message || analysisError);
      }

      // Map selection -> full objects
      let sensitiveDataToRemove = [];
      if (hasSelectiveSensitive && detections?.sensitiveData) {
        if (typeof removeSensitiveDataRaw[0] === "string") {
          sensitiveDataToRemove = detections.sensitiveData.filter((d) => removeSensitiveDataRaw.includes(d.id));
        } else {
          sensitiveDataToRemove = removeSensitiveDataRaw;
        }
      }

      let hiddenContentToRemove = [];
      if (hasSelectiveHidden && detections?.hiddenContent) {
        if (typeof hiddenContentToCleanRaw[0] === "string") {
          hiddenContentToRemove = detections.hiddenContent.filter((d) => hiddenContentToCleanRaw.includes(d.id));
        } else {
          hiddenContentToRemove = hiddenContentToCleanRaw;
        }
      }

      let visualObjectsToRemove = [];
      if (hasSelectiveVisual && detections?.visualObjects) {
        if (typeof visualObjectsToCleanRaw[0] === "string") {
          visualObjectsToRemove = detections.visualObjects.filter((d) => visualObjectsToCleanRaw.includes(d.id));
        } else {
          visualObjectsToRemove = visualObjectsToCleanRaw;
        }
      }

      console.log(
        `[CLEAN] Selected: sensitive=${sensitiveDataToRemove.length}, hidden=${hiddenContentToRemove.length}, visual=${visualObjectsToRemove.length}`
      );

      const spellingFixList =
        Array.isArray(approvedSpellingErrors) && approvedSpellingErrors.length > 0
          ? approvedSpellingErrors
          : spellingErrors;

      const extraRemovals = {
        sensitiveDataRemoved: 0,
        hiddenContentRemoved: 0,
      };

      // ---------------- DOCX ----------------
      if (ext === "docx") {
        let currentBuffer = f.buffer;

        const cleaned = await cleanDOCX(currentBuffer, { drawPolicy, ...cleaningOptions });
        currentBuffer = cleaned.outBuffer;

        if (hasSelectiveSensitive && sensitiveDataToRemove.length > 0) {
          const sensitiveResult = await removeSensitiveDataFromDOCX(currentBuffer, sensitiveDataToRemove);
          currentBuffer = sensitiveResult.outBuffer;
          extraRemovals.sensitiveDataRemoved = sensitiveResult.stats.removed;
        }

        if (hasSelectiveHidden && hiddenContentToRemove.length > 0) {
          const hiddenResult = await removeHiddenContentFromDOCX(currentBuffer, hiddenContentToRemove);
          currentBuffer = hiddenResult.outBuffer;
          extraRemovals.hiddenContentRemoved = hiddenResult.stats.removed;
        }

        let correctionStats = null;
        if (cleaningOptions.correctSpelling) {
          const corrected = await correctDOCXText(currentBuffer, aiCorrectText, {
            spellingErrors: spellingFixList,
          });
          currentBuffer = corrected.outBuffer;
          correctionStats = corrected.stats;
        }

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

        const cleaned = await cleanPPTX(currentBuffer, { drawPolicy, ...cleaningOptions });
        currentBuffer = cleaned.outBuffer;

        if (hasSelectiveSensitive && sensitiveDataToRemove.length > 0) {
          const sensitiveResult = await removeSensitiveDataFromPPTX(currentBuffer, sensitiveDataToRemove);
          currentBuffer = sensitiveResult.outBuffer;
          extraRemovals.sensitiveDataRemoved = sensitiveResult.stats.removed;
        }

        if (hasSelectiveHidden && hiddenContentToRemove.length > 0) {
          const hiddenResult = await removeHiddenContentFromPPTX(currentBuffer, hiddenContentToRemove);
          currentBuffer = hiddenResult.outBuffer;
          extraRemovals.hiddenContentRemoved = hiddenResult.stats.removed;
        }

        if (hasSelectiveVisual && visualObjectsToRemove.length > 0) {
          const visualResult = await removeVisualObjectsFromPPTX(currentBuffer, visualObjectsToRemove);
          currentBuffer = visualResult.outBuffer;
        }

        let correctionStats = null;
        if (cleaningOptions.correctSpelling) {
          const corrected = await correctPPTXText(currentBuffer, aiCorrectText, {
            spellingErrors: spellingFixList,
          });
          currentBuffer = corrected.outBuffer;
          correctionStats = corrected.stats;
        }

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

        const cleaned = await cleanXLSX(currentBuffer, cleaningOptions);
        currentBuffer = cleaned.outBuffer;

        if (hasSelectiveSensitive && sensitiveDataToRemove.length > 0) {
          const sensitiveResult = await removeSensitiveDataFromXLSX(currentBuffer, sensitiveDataToRemove);
          currentBuffer = sensitiveResult.outBuffer;
          extraRemovals.sensitiveDataRemoved = sensitiveResult.stats.removed;
        }

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
// POST /rephrase - VERSION 3.1.0
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

      // ✅ NEW: risk objects for report
      let riskObjects = [];
      let riskSummary = null;

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

        // ✅ NEW
        riskObjects = mapDetectionsToRiskObjects(ext, detections || {});
        riskSummary = summarizeRiskObjects(riskObjects);

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
            riskSummary,
          },
          riskObjects,
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
  console.log(`✅ Qualion-Doc Backend v3.1.0 listening on port ${PORT}`);
  console.log(`   Endpoints: GET /health, POST /analyze, POST /clean, POST /rephrase`);
  console.log(`   Features: universal riskObjects + client-ready gate + selective(non-empty)`);
});
