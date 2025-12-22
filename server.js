// server.js - VERSION 3.2.0
// ✅ Compatible documentAnalyzer (keeps existing detections + summary payloads)
// ✅ Adds Qualion Part 2: Business Risk (5 categories) using deterministic rules
// ✅ Keeps Universal Risk Model (riskObjects[] + riskSummary)
// ✅ Non-breaking: existing keys remain, new key: qualionReport + businessSignals

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
// Universal Risk Model (kept) + Qualion Business Risk V1 (NEW)
// ============================================================

// 4 Exposure Surfaces
const SURFACE = {
  VISIBLE: "visible",
  HIDDEN: "hidden",
  LOGIC: "logic",
  METADATA: "metadata",
};

// 5 Business Risk Categories (Qualion Part 2)
const EXEC_CAT = {
  MARGIN: "margin",
  DELIVERY: "delivery",
  NEGOTIATION: "negotiation",
  COMPLIANCE: "compliance",
  CREDIBILITY: "credibility",
};

const EXEC_CAT_LABEL = {
  margin: "Margin Exposure Risk",
  delivery: "Delivery & Commitment Risk",
  negotiation: "Negotiation Power Leakage",
  compliance: "Compliance & Confidentiality Risk",
  credibility: "Professional Credibility Risk",
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

// Simple deterministic points
function pointsForRiskObject(ro) {
  let pts = 0;

  if (ro.surface === SURFACE.HIDDEN) pts += 3;
  if (ro.surface === SURFACE.LOGIC) pts += 2;
  if (ro.surface === SURFACE.METADATA) pts += 2;

  if (ro.ruleId === "XLSX_FORMULA" || ro.ruleId === "XLSX_EXTERNAL_REF" || ro.ruleId === "BROKEN_LINK") pts += 2;

  // Financial proximity
  if (ro.category === EXEC_CAT.MARGIN) pts += 2;

  return pts;
}

// Deterministic mapping for sensitive data types -> business category
function categorizeSensitiveType(detType) {
  const t = String(detType || "").toLowerCase();

  // Margin exposure identifiers
  if (["price", "rate", "rate_card", "iban", "credit_card"].includes(t)) return EXEC_CAT.MARGIN;

  // Negotiation leakage / internal posture
  if (["project_code", "internal_url", "file_path", "server_path", "ip_address"].includes(t)) return EXEC_CAT.NEGOTIATION;

  // Compliance & confidentiality identifiers
  if (["email", "phone", "ssn", "pii", "personal_data"].includes(t)) return EXEC_CAT.COMPLIANCE;

  return EXEC_CAT.CREDIBILITY;
}

/**
 * NEW: Deterministic business flags feeding Part 2.
 * This is where you extend rules later (especially DELIVERY).
 */
function deriveBusinessSignals(ext, detections = {}) {
  const signals = [];

  const push = (s) =>
    signals.push({
      id: s.id || `bs_${uuidv4()}`,
      category: s.category,          // margin|delivery|negotiation|compliance|credibility
      severity: normalizeSeverity(s.severity || "low"),
      ruleId: s.ruleId,              // stable string
      signalFamily: s.signalFamily,  // stable family key (for analytics)
      reason: s.reason,              // neutral executive framing (no intent)
      location: s.location || "Document",
      source: s.source || null,
    });

  // ---- Margin exposure (XLSX / pricing logic)
  (detections.sensitiveFormulas || []).forEach((f) => {
    push({
      category: EXEC_CAT.MARGIN,
      severity: f.severity || f.risk || "medium",
      ruleId: "MARGIN_FORMULA_OR_LOGIC",
      signalFamily: "pricing_calculation_artifacts",
      reason: "Spreadsheet formulas detected that could reveal pricing logic.",
      location: f.sheet ? `Sheet: ${f.sheet}` : "Workbook",
      source: { group: "sensitiveFormulas", id: f.id || null },
    });
  });

  (detections.hiddenSheets || []).forEach((hs) => {
    push({
      category: EXEC_CAT.MARGIN,
      severity: hs.severity || "high",
      ruleId: "MARGIN_HIDDEN_SHEET",
      signalFamily: "pricing_structure_signals",
      reason: "Hidden worksheets detected that may contain pricing or cost structures.",
      location: hs.sheetName ? `Sheet: ${hs.sheetName}` : "Workbook",
      source: { group: "hiddenSheets", id: hs.id || null },
    });
  });

  (detections.excelHiddenData || []).forEach((x) => {
    const type = String(x.type || "").toLowerCase();
    const isFormula = type.includes("formula");

    push({
      category: EXEC_CAT.MARGIN,
      severity: x.severity || (isFormula ? "high" : "medium"),
      ruleId: isFormula ? "MARGIN_HIDDEN_FORMULA" : "MARGIN_HIDDEN_DATA",
      signalFamily: isFormula ? "pricing_calculation_artifacts" : "pricing_structure_signals",
      reason: isFormula
        ? "Hidden spreadsheet logic detected that could reveal pricing assumptions."
        : "Hidden spreadsheet data detected that may expose internal cost or pricing elements.",
      location: x.location || "Workbook",
      source: { group: "excelHiddenData", id: x.id || null },
    });
  });

  // ---- Negotiation leakage (metadata, hidden content, embedded objects, internal traces)
  (detections.metadata || []).forEach((m) => {
    push({
      category: EXEC_CAT.NEGOTIATION,
      severity: m.severity || "medium",
      ruleId: "NEGOTIATION_METADATA",
      signalFamily: "traceability_artifacts",
      reason: "Metadata detected that may reveal internal authorship or organization context.",
      location: m.location || "Document Properties",
      source: { group: "metadata", id: m.id || null },
    });
  });

  (detections.hiddenContent || []).forEach((h) => {
    push({
      category: EXEC_CAT.NEGOTIATION,
      severity: h.severity || "high",
      ruleId: "NEGOTIATION_HIDDEN_CONTENT",
      signalFamily: "internal_assumption_signals",
      reason: "Hidden content detected that may expose internal assumptions.",
      location: h.location || "Document",
      source: { group: "hiddenContent", id: h.id || null },
    });
  });

  (detections.embeddedObjects || []).forEach((e) => {
    push({
      category: EXEC_CAT.NEGOTIATION,
      severity: e.severity || "medium",
      ruleId: "NEGOTIATION_EMBEDDED_OBJECT",
      signalFamily: "traceability_artifacts",
      reason: "Embedded objects detected that may contain additional internal data.",
      location: e.path || "Document",
      source: { group: "embeddedObjects", id: e.id || null },
    });
  });

  // ---- Compliance & confidentiality (explicit compliance risks + PII-like sensitive data)
  (detections.complianceRisks || []).forEach((c) => {
    push({
      category: EXEC_CAT.COMPLIANCE,
      severity: c.severity || "high",
      ruleId: "COMPLIANCE_SIGNAL",
      signalFamily: "identifier_exposure_signals",
      reason: "Confidential or regulated identifiers detected in a client-facing document.",
      location: c.location || "Document",
      source: { group: "complianceRisks", id: c.id || null },
    });
  });

  (detections.sensitiveData || []).forEach((s) => {
    const cat = categorizeSensitiveType(s.type);

    // Only route to compliance when applicable; otherwise negotiation/margin/credibility
    push({
      category: cat,
      severity: s.severity || "high",
      ruleId: "VISIBLE_SENSITIVE",
      signalFamily: cat === EXEC_CAT.COMPLIANCE ? "identifier_exposure_signals" : "client_facing_exposure",
      reason:
        cat === EXEC_CAT.COMPLIANCE
          ? "Sensitive identifiers detected in a client-facing document."
          : "Sensitive information detected that may not be necessary for client understanding.",
      location: s.location || "Document",
      source: { group: "sensitiveData", id: s.id || null, type: s.type || null },
    });
  });

  // ---- Professional credibility (comments, track changes, spelling, formatting artifacts)
  (detections.comments || []).forEach((c) => {
    push({
      category: EXEC_CAT.CREDIBILITY,
      severity: c.severity || "medium",
      ruleId: c.type === "tracked_change" ? "CRED_TRACKED_CHANGE" : "CRED_COMMENT",
      signalFamily: "draft_artifact_signals",
      reason:
        c.type === "tracked_change"
          ? "Tracked changes detected that may expose draft edits or internal review."
          : "Comments detected that may expose internal review discussions.",
      location: c.location || "Document",
      source: { group: "comments", id: c.id || null },
    });
  });

  (detections.trackChanges || []).forEach((tc) => {
    push({
      category: EXEC_CAT.CREDIBILITY,
      severity: tc.severity || "medium",
      ruleId: "CRED_TRACKED_CHANGE",
      signalFamily: "draft_artifact_signals",
      reason: "Tracked changes detected that may expose draft edits or internal review.",
      location: tc.location || "Document",
      source: { group: "trackChanges", id: tc.id || null },
    });
  });

  (detections.spellingErrors || []).forEach((sp) => {
    push({
      category: EXEC_CAT.CREDIBILITY,
      severity: sp.severity || "low",
      ruleId: "CRED_SPELLING",
      signalFamily: "formatting_consistency_issues",
      reason: "Spelling or grammar issues detected that may affect professional credibility.",
      location: sp.location || "Document",
      source: { group: "spellingErrors", id: sp.id || null },
    });
  });

  (detections.orphanData || []).forEach((o) => {
    push({
      category: EXEC_CAT.CREDIBILITY,
      severity: o.severity || "low",
      ruleId: "CRED_FORMATTING_ARTIFACT",
      signalFamily: "formatting_consistency_issues",
      reason: "Residual formatting artifacts detected that may require a quick review.",
      location: o.location || "Document",
      source: { group: "orphanData", id: o.id || null },
    });
  });

  // ---- Delivery & commitment (PLACEHOLDER deterministic mapping)
  // NOTE: Your analyzer currently doesn't output explicit delivery-language detections.
  // If businessInconsistencies has commitment-related entries, map them here deterministically.
  (detections.businessInconsistencies || []).forEach((b) => {
    const t = String(b.type || b.key || "").toLowerCase();
    const hint = String(b.reason || b.message || "").toLowerCase();

    const looksDelivery =
      t.includes("commit") ||
      t.includes("delivery") ||
      t.includes("deadline") ||
      t.includes("scope") ||
      hint.includes("commit") ||
      hint.includes("deadline") ||
      hint.includes("scope");

    if (!looksDelivery) return;

    push({
      category: EXEC_CAT.DELIVERY,
      severity: b.severity || "medium",
      ruleId: "DELIVERY_COMMITMENT_SIGNAL",
      signalFamily: "engagement_language_signals",
      reason: "Commitment language detected without clear delivery boundaries.",
      location: b.location || "Document",
      source: { group: "businessInconsistencies", id: b.id || null },
    });
  });

  return signals;
}

/**
 * Kept: mapDetectionsToRiskObjects (Universal model) - now updated to 5 cats where appropriate.
 */
function mapDetectionsToRiskObjects(fileType, detections) {
  const riskObjects = [];
  const ext = String(fileType || "").toLowerCase();

  const pushRO = (ro) => {
    riskObjects.push({
      id: ro.id || `ro_${uuidv4()}`,
      surface: ro.surface,
      category: ro.category, // margin|delivery|negotiation|compliance|credibility
      severity: normalizeSeverity(ro.severity),
      fileType: ext,
      ruleId: ro.ruleId,
      reason: ro.reason,
      location: ro.location || "Document",
      fixability: ro.fixability || "manual",
      source: ro.source || null,
      meta: ro.meta || {},
    });
  };

  // METADATA => negotiation surface: metadata
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

  // COMMENTS => hidden => credibility
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

  // TRACK CHANGES => hidden => credibility
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

  // HIDDEN CONTENT => hidden => negotiation
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

  // XLSX HIDDEN SHEETS => hidden => margin
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

  // XLSX FORMULAS => logic => margin
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

  // XLSX excelHiddenData => hidden/logic => margin
  (detections.excelHiddenData || []).forEach((x) => {
    const type = String(x.type || "").toLowerCase();
    const isFormula = type.includes("formula");
    pushRO({
      id: `ro_excelhidden_${x.id || uuidv4()}`,
      surface: isFormula ? SURFACE.LOGIC : SURFACE.HIDDEN,
      category: EXEC_CAT.MARGIN,
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

  // EMBEDDED OBJECTS => logic => negotiation
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

  // MACROS => logic => credibility (trust)
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

  // SENSITIVE DATA => visible => compliance/negotiation/margin/credibility (deterministic)
  (detections.sensitiveData || []).forEach((s) => {
    const cat = categorizeSensitiveType(s.type);
    pushRO({
      id: `ro_sensitive_${s.id || uuidv4()}`,
      surface: SURFACE.VISIBLE,
      category: cat,
      severity: s.severity || "high",
      ruleId: "VISIBLE_SENSITIVE",
      reason:
        cat === EXEC_CAT.COMPLIANCE
          ? "Sensitive identifiers detected in a client-facing document."
          : "Visible sensitive information detected that may not be necessary for client understanding.",
      location: s.location || "Document",
      fixability: "manual",
      source: { detectionGroup: "sensitiveData", detectionId: s.id || null },
      meta: { type: s.type || null, category: s.category || null },
    });
  });

  // SPELLING => credibility
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

  // BROKEN LINKS => credibility
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

  // COMPLIANCE RISKS => compliance
  (detections.complianceRisks || []).forEach((c) => {
    pushRO({
      id: `ro_compliance_${c.id || uuidv4()}`,
      surface: SURFACE.VISIBLE,
      category: EXEC_CAT.COMPLIANCE,
      severity: c.severity || "high",
      ruleId: "COMPLIANCE_SIGNAL",
      reason: "Compliance or confidentiality signals detected that require review before external sharing.",
      location: c.location || "Document",
      fixability: "manual",
      source: { detectionGroup: "complianceRisks", detectionId: c.id || null },
      meta: { type: c.type || null },
    });
  });

  // VISUAL OBJECTS => credibility
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

  // ORPHAN DATA => credibility
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

  // DELIVERY is handled through businessSignals for now (until analyzer emits explicit delivery detections)

  return riskObjects;
}

function summarizeRiskObjects(riskObjects) {
  const byCategory = {
    [EXEC_CAT.MARGIN]: { points: 0, count: 0, maxSeverity: "low" },
    [EXEC_CAT.DELIVERY]: { points: 0, count: 0, maxSeverity: "low" },
    [EXEC_CAT.NEGOTIATION]: { points: 0, count: 0, maxSeverity: "low" },
    [EXEC_CAT.COMPLIANCE]: { points: 0, count: 0, maxSeverity: "low" },
    [EXEC_CAT.CREDIBILITY]: { points: 0, count: 0, maxSeverity: "low" },
  };

  let anyCritical = false;

  for (const ro of riskObjects) {
    const cat = ro.category || EXEC_CAT.CREDIBILITY;
    const pts = pointsForRiskObject(ro);
    byCategory[cat].points += pts;
    byCategory[cat].count += 1;

    const sev = normalizeSeverity(ro.severity);
    if (sev === "critical") anyCritical = true;

    const r = severityRank(sev);
    if (r > severityRank(byCategory[cat].maxSeverity)) byCategory[cat].maxSeverity = sev;
  }

  const categorySeverityFromPoints = (pts, maxSev) => {
    if (maxSev === "critical") return "critical";
    if (pts >= 7) return "high";
    if (pts >= 4) return "medium";
    if (pts > 0) return "low";
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

  const executiveSignals = [];
  if (perCat[EXEC_CAT.MARGIN].points > 0) executiveSignals.push(EXEC_CAT_LABEL.margin);
  if (perCat[EXEC_CAT.DELIVERY].points > 0) executiveSignals.push(EXEC_CAT_LABEL.delivery);
  if (perCat[EXEC_CAT.NEGOTIATION].points > 0) executiveSignals.push(EXEC_CAT_LABEL.negotiation);
  if (perCat[EXEC_CAT.COMPLIANCE].points > 0) executiveSignals.push(EXEC_CAT_LABEL.compliance);
  if (perCat[EXEC_CAT.CREDIBILITY].points > 0) executiveSignals.push(EXEC_CAT_LABEL.credibility);

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

  const anyHighCategory =
    perCat[EXEC_CAT.MARGIN].points >= 7 ||
    perCat[EXEC_CAT.DELIVERY].points >= 7 ||
    perCat[EXEC_CAT.NEGOTIATION].points >= 7 ||
    perCat[EXEC_CAT.COMPLIANCE].points >= 7 ||
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
      compliance: perCat[EXEC_CAT.COMPLIANCE],
      credibility: perCat[EXEC_CAT.CREDIBILITY],
    },
    blockingIssues: blocking,
    totalRiskObjects: riskObjects.length,
  };
}

// ---------------- Qualion Report (Part 1 + Part 2) ----------------

function buildPart1Checklist(detections = {}) {
  const has = (arr) => Array.isArray(arr) && arr.length > 0;

  return [
    {
      key: "hidden_content",
      present: has(detections.hiddenContent) || has(detections.hiddenSheets) || has(detections.excelHiddenData),
      label: "Hidden content detected in client-facing document",
    },
    {
      key: "comments_present",
      present: has(detections.comments),
      label: "Internal comments present",
    },
    {
      key: "track_changes_detected",
      present: has(detections.trackChanges),
      label: "Track changes detected",
    },
    {
      key: "metadata_exposure",
      present: has(detections.metadata),
      label: "Metadata revealing internal context detected",
    },
    {
      key: "embedded_objects",
      present: has(detections.embeddedObjects),
      label: "Embedded objects or attachments detected",
    },
    {
      key: "structural_hidden_elements",
      present: has(detections.hiddenSheets) || has(detections.visualObjects),
      label: "Structural issues detected (hidden slides, sheets, layers)",
    },
    {
      key: "formatting_anomalies",
      present: has(detections.orphanData) || has(detections.visualObjects),
      label: "Formatting anomalies or draft artifacts detected",
    },
  ];
}

function riskLevelFromSeverityLabel(s) {
  const v = String(s || "").toLowerCase();
  if (v === "critical" || v === "high") return "High";
  if (v === "medium") return "Medium";
  if (v === "low") return "Low";
  return "None";
}

function buildBusinessCategoryOutput(catKey, businessSignals, riskSummary) {
  // Risk level comes from riskSummary severity by category
  const sev = riskSummary?.byCategory?.[catKey]?.severity || "low";
  const level = riskLevelFromSeverityLabel(sev);

  // Key drivers from businessSignals (signalFamily)
  const drivers = Array.from(
    new Set(
      (businessSignals || [])
        .filter((s) => s.category === catKey)
        .map((s) => s.signalFamily || s.ruleId)
    )
  ).slice(0, 5);

  const sentences = [];
  if (level === "None") {
    sentences.push("No client-facing business risk signals detected in this category.");
  } else {
    sentences.push(
      level === "High"
        ? "High-risk signals detected that may increase exposure if the document is sent externally."
        : "Signals detected that may increase exposure if the document is sent externally."
    );
    sentences.push("Signals are derived from deterministic hygiene, structure, and identifier checks.");
    if (drivers.length) sentences.push(`Key indicators: ${drivers.slice(0, 3).join("; ")}.`);
  }

  return {
    category: catKey,
    title: EXEC_CAT_LABEL[catKey],
    riskLevel: level, // None / Low / Medium / High
    summary: sentences.slice(0, 3),
  };
}

function computeBusinessRecommendation(clientReady, overallSeverityLabel) {
  const sev = String(overallSeverityLabel || "").toLowerCase();

  if (clientReady === "NO") {
    return "Review flagged business risk items before external sharing, then re-run Qualion Clean.";
  }
  if (sev === "medium") {
    return "Consider reviewing highlighted items to reduce exposure before external sharing.";
  }
  return "No blocking business risks detected. Document is suitable for client-facing export.";
}

function buildQualionReport({ documentId, fileName, ext, detections, riskObjects, riskSummary, businessSignals }) {
  const part1 = {
    title: "Technical & Content Hygiene Report",
    checklist: buildPart1Checklist(detections || {}),
  };

  const part2 = {
    title: "Business Risk Report",
    collapsedByDefault: true,
    clientReady: riskSummary?.clientReady || "YES",
    overallSeverity: riskSummary?.overallSeverity || "Low",
    categories: [
      buildBusinessCategoryOutput("margin", businessSignals, riskSummary),
      buildBusinessCategoryOutput("delivery", businessSignals, riskSummary),
      buildBusinessCategoryOutput("negotiation", businessSignals, riskSummary),
      buildBusinessCategoryOutput("compliance", businessSignals, riskSummary),
      buildBusinessCategoryOutput("credibility", businessSignals, riskSummary),
    ],
    recommendation: computeBusinessRecommendation(riskSummary?.clientReady || "YES", riskSummary?.overallSeverity || "Low"),
  };

  return {
    meta: {
      documentId,
      fileName,
      fileType: ext,
      analyzedAt: new Date().toISOString(),
      version: "v1",
    },
    part1,
    part2,
    // keep internal helpers if you want to debug; front can ignore
    internal: {
      businessSignalsCount: businessSignals?.length || 0,
      riskObjectsCount: riskObjects?.length || 0,
    },
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
    version: "3.2.0",
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
      "qualion-report-part1-part2",
      "business-risk-5-categories",
      "deterministic-business-signals",
    ],
    time: new Date().toISOString(),
  })
);

// ===================================================================
// POST /analyze - VERSION 3.2.0
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

    const documentId = uuidv4();

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

    // docStats
    const docStats = await safeExtractDocStats(req.file.buffer, ext);
    const documentStats =
      docStats && (docStats.pages || docStats.slides || docStats.sheets || docStats.tables)
        ? docStats
        : analysisResult.documentStats || docStats;

    // Universal risk model (kept)
    const riskObjects = mapDetectionsToRiskObjects(ext, detections || {});
    const riskSummary = summarizeRiskObjects(riskObjects);

    // NEW: Business signals feeding Qualion Part 2
    const businessSignals = deriveBusinessSignals(ext, detections || {});

    // NEW: Qualion Part 1 + Part 2 report
    const qualionReport = buildQualionReport({
      documentId,
      fileName: req.file.originalname,
      ext,
      detections,
      riskObjects,
      riskSummary,
      businessSignals,
    });

    res.json({
      documentId,
      fileName: req.file.originalname,
      fileType: ext,
      fileSize: req.file.size,

      documentStats,

      // keep current detections structure
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

      // keep legacy summary shape
      summary: {
        ...summary,
        riskScore,
        riskLevel: getRiskLevel(riskScore),
        riskBreakdown: breakdown,
        recommendations: generateRecommendations(detections),
      },

      // kept outputs
      riskObjects,
      riskSummary,

      // NEW outputs for Qualion V1
      businessSignals,
      qualionReport,

      processingTime: Date.now() - startTime,
    });
  } catch (e) {
    console.error("[ANALYZE ERROR]", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ===================================================================
// POST /clean - VERSION 3.2.0 (same behavior, now also includes qualionReport in report payload)
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

      const documentStatsBefore = await safeExtractDocStats(f.buffer, ext);

      let analysisResult = null;
      let spellingErrors = [];
      let beforeRiskScore = 100;
      let riskBreakdown = {};
      let detections = null;
      let summary = null;

      // universal + business additions for report
      let riskObjects = [];
      let riskSummary = null;
      let businessSignals = [];
      let qualionReport = null;

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

        riskObjects = mapDetectionsToRiskObjects(ext, detections || {});
        riskSummary = summarizeRiskObjects(riskObjects);

        businessSignals = deriveBusinessSignals(ext, detections || {});
        qualionReport = buildQualionReport({
          documentId: uuidv4(),
          fileName: f.originalname,
          ext,
          detections,
          riskObjects,
          riskSummary,
          businessSignals,
        });

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
          businessSignals,
          qualionReport,
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

      // Other
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
// POST /rephrase - VERSION 3.2.0 (unchanged behavior, enriched report payload)
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

      let riskObjects = [];
      let riskSummary = null;
      let businessSignals = [];
      let qualionReport = null;

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

        riskObjects = mapDetectionsToRiskObjects(ext, detections || {});
        riskSummary = summarizeRiskObjects(riskObjects);

        businessSignals = deriveBusinessSignals(ext, detections || {});
        qualionReport = buildQualionReport({
          documentId: uuidv4(),
          fileName: f.originalname,
          ext,
          detections,
          riskObjects,
          riskSummary,
          businessSignals,
        });

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
          businessSignals,
          qualionReport,
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
  console.log(`✅ Qualion-Doc Backend v3.2.0 listening on port ${PORT}`);
  console.log(`   Endpoints: GET /health, POST /analyze, POST /clean, POST /rephrase`);
  console.log(`   Features: qualionReport(part1+part2) + 5 business risk categories`);
});
