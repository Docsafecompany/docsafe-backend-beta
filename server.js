// server.js - VERSION 3.3.0
// ✅ Qualion Clean V1 aligned (SINGLE SOURCE OF TRUTH)
// ✅ Adds Part 2: Business Risk (5 categories) with deterministic flags
// ✅ Keeps existing detections/summary/riskObjects/riskSummary (non-breaking)
// ✅ Adds executive scoring (25/25/25/25) + critical gate + override-ready payload
// ✅ No AI guessing. No semantic interpretation. No legal advice.

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
// 0) QUALION CLEAN V1 — SOURCE OF TRUTH CONSTANTS
// ============================================================

const QUALION_V1 = {
  supportedFileTypes: ["pdf", "docx", "xlsx", "pptx"],
  fileTypeFeatures: {
    pdf: {
      primaryRisk: "Irreversible external exposure.",
      features: [
        "Metadata detection & removal",
        "Visible & hidden annotations",
        "Layered content detection",
        "Embedded objects & attachments",
        "Hyperlinks detection",
        "Form fields detection",
        "Unsafe redaction detection",
      ],
      businessValue: [
        "Prevent irreversible IP leakage",
        "Protect executive credibility",
        "Ensure final documents are truly client-ready",
      ],
    },
    docx: {
      primaryRisk: "Unintended commitments and internal exposure.",
      features: [
        "Track changes detection",
        "Internal comments detection",
        "Hidden sections & non-printable content",
        "Template & style origin detection",
        "Automatic fields detection",
        "Version residue detection",
      ],
      businessValue: [
        "Avoid exposing internal debates or uncertainty",
        "Prevent accidental delivery commitments",
        "Strengthen negotiation posture",
      ],
    },
    xlsx: {
      primaryRisk: "Margin destruction and negotiation leverage loss.",
      features: [
        "Hidden & very hidden sheets detection",
        "Formula exposure detection",
        "External links detection",
        "Pricing scenarios detection",
        "Cell comments detection",
        "Macro presence detection (no execution)",
      ],
      businessValue: [
        "Protect pricing logic and margins",
        "Prevent procurement reverse-engineering",
        "Maintain commercial leverage",
      ],
    },
    pptx: {
      primaryRisk: "Strategic narrative exposure and credibility loss.",
      features: [
        "Speaker notes detection",
        "Hidden slides detection",
        "Off-slide content detection",
        "Embedded media & objects",
        "Reviewer comments",
        "Metadata & template origin",
        "Hidden text boxes",
      ],
      businessValue: [
        "Prevent leakage of internal strategy or talking points",
        "Protect executive storytelling",
        "Avoid client confusion or credibility loss",
      ],
    },
  },
  businessRiskCategories: {
    margin: {
      title: "Margin Exposure Risk",
      definition: "Any signal allowing a client to infer pricing logic or margins.",
      signals: ["Pricing logic references", "Cost or margin assumptions", "Excel formulas & scenarios", "Rate references"],
      businessValue: ["Preserve negotiation power", "Avoid margin erosion before negotiation even starts"],
    },
    delivery: {
      title: "Delivery & Commitment Risk",
      definition: "Signals indicating unintended or unbounded delivery obligations.",
      signals: [
        "Strong engagement language",
        "Open-ended deliverables",
        "Fixed price without boundaries",
        "Deadlines without dependencies",
      ],
      businessValue: [
        "Prevent scope creep",
        "Protect delivery margins",
        "Align sales promises with execution reality",
      ],
    },
    negotiation: {
      title: "Negotiation Power Leakage",
      definition: "Information weakening commercial leverage.",
      signals: ["Internal assumptions", "Pricing options A/B/C", "Client dependency mentions", "Internal benchmarks"],
      businessValue: ["Maintain deal control", "Reduce procurement leverage"],
    },
    compliance: {
      title: "Compliance & Confidentiality Risk",
      definition: "Exposure of confidential or regulated information.",
      signals: ["Emails and identifiers", "Confidential markers", "Project codes", "Personal data indicators"],
      businessValue: ["Reduce audit and legal exposure", "Protect corporate reputation"],
    },
    credibility: {
      title: "Professional Credibility Risk",
      definition: "Signals degrading perceived professionalism.",
      signals: ["Comments & track changes", "Structural inconsistencies", "Hidden content", "Formatting anomalies"],
      businessValue: ["Strengthen trust", "Reinforce premium brand positioning"],
    },
  },

  // Executive weighting model (per your Section 4)
  // NOTE: Compliance is a CRITICAL gate category (does not need to be part of 25/25/25/25)
  scoringWeights: {
    margin: 0.25,
    delivery: 0.25,
    negotiation: 0.25,
    credibility: 0.25,
  },
};

// ============================================================
// 1) LEGACY SCORE (kept)
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

function getRiskLevel(score) {
  if (score >= 90) return "safe";
  if (score >= 70) return "low";
  if (score >= 50) return "medium";
  if (score >= 25) return "high";
  return "critical";
}

// ============================================================
// 2) QUALION CLEAN V1 — PART 2 BUSINESS RISK (DETERMINISTIC)
// ============================================================

const BIZ_SEVERITY = {
  NONE: "None",
  LOW: "Low",
  MEDIUM: "Medium",
  HIGH: "High",
  CRITICAL: "Critical",
};

// deterministic ranking
function bizRank(level) {
  const s = String(level || "").toLowerCase();
  if (s === "critical") return 5;
  if (s === "high") return 4;
  if (s === "medium") return 3;
  if (s === "low") return 2;
  return 1;
}

// Business “flag” object (Part 2)
function makeBizFlag({
  category,
  level,
  ruleId,
  reason,
  location = "Document",
  evidence = null,
  source = null,
}) {
  return {
    id: `bf_${uuidv4()}`,
    category, // margin|delivery|negotiation|compliance|credibility
    level, // None/Low/Medium/High/Critical
    ruleId, // stable string
    reason, // neutral exec wording
    location,
    evidence, // optional: small excerpt or signal key (NOT interpretation)
    source, // optional reference
  };
}

// Extract text deterministically for language-based rules
async function getDeterministicText(ext, buffer, analysisResult) {
  // Prefer analyzer-provided extracted text if available
  const candidate =
    analysisResult?.extractedText ||
    analysisResult?.text ||
    analysisResult?.documentText ||
    null;

  if (typeof candidate === "string" && candidate.trim()) return candidate;

  // PDF fallback: extract text using existing tools (deterministic)
  if (ext === "pdf") {
    try {
      const raw = await extractPdfText(buffer);
      const filtered = filterExtractedLines(raw, { strictPdf: true });
      return String(filtered || "");
    } catch (e) {
      console.warn("[BIZ TEXT] PDF extraction failed:", e?.message || e);
      return "";
    }
  }

  // DOCX/PPTX/XLSX: rely on analyzer output (or implement dedicated extractors later)
  return "";
}

// deterministic keyword/regex library (STRICT, no AI)
const RULES = {
  delivery: {
    strongEngagement: [
      /\bwe\s+will\b/gi,
      /\bwe\s+commit\b/gi,
      /\bwe\s+guarantee\b/gi,
      /\bwe\s+ensure\b/gi,
      /\bdeliver\s+by\b/gi,
      /\bcommitment\b/gi,
    ],
    openEndedDeliverables: [
      /\bas\s+needed\b/gi,
      /\bunlimited\b/gi,
      /\bongoing\b/gi,
      /\bcontinuous\b/gi,
      /\bsupport\s+until\b/gi,
      /\bfull\s+ownership\b/gi,
      /\bend-to-end\b/gi,
    ],
    fixedPriceNoBoundaries: [
      /\bfixed\s+price\b/gi,
      /\bflat\s+fee\b/gi,
      /\ball-inclusive\b/gi,
      /\bturnkey\b/gi,
    ],
    deadlinesNoDependencies: [
      /\bby\s+(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/gi,
      /\bby\s+\d{1,2}\/\d{1,2}(\/\d{2,4})?\b/gi,
      /\bno\s+later\s+than\b/gi,
      /\bdeadline\b/gi,
    ],
    dependencyMarkers: [
      /\bsubject\s+to\b/gi,
      /\bassuming\b/gi,
      /\bdependent\s+on\b/gi,
      /\bclient\s+to\s+provide\b/gi,
      /\bprerequisite\b/gi,
    ],
  },
  negotiation: {
    internalAssumptions: [
      /\binternal\s+assumption\b/gi,
      /\bassumption\b/gi,
      /\bworking\s+hypothesis\b/gi,
      /\bnot\s+for\s+client\b/gi,
      /\bdo\s+not\s+share\b/gi,
    ],
    optionsABC: [
      /\boption\s+a\b/gi,
      /\boption\s+b\b/gi,
      /\boption\s+c\b/gi,
      /\balternative\s+a\b/gi,
      /\balternative\s+b\b/gi,
    ],
    clientDependency: [
      /\bwe\s+need\s+this\s+deal\b/gi,
      /\bstrategic\s+client\b/gi,
      /\bmust-win\b/gi,
      /\bpriority\s+account\b/gi,
    ],
    internalBenchmarks: [
      /\bbenchmark\b/gi,
      /\btarget\s+rate\b/gi,
      /\bwalk-away\b/gi,
      /\breservation\s+price\b/gi,
      /\bmargin\s+target\b/gi,
    ],
  },
  margin: {
    pricingKeywords: [
      /\brate\b/gi,
      /\brate\s+card\b/gi,
      /\bunit\s+cost\b/gi,
      /\bcost\b/gi,
      /\bmargin\b/gi,
      /\bmarkup\b/gi,
      /\bdiscount\b/gi,
      /\bpricing\b/gi,
    ],
  },
  compliance: {
    confidentialMarkers: [
      /\bconfidential\b/gi,
      /\bproprietary\b/gi,
      /\binternal\s+use\s+only\b/gi,
      /\bnda\b/gi,
      /\bexport\s+control\b/gi,
      /\bitar\b/gi,
      /\bear\b/gi,
    ],
    projectCodeLike: [
      /\bproj[-_\s]?\d{3,}\b/gi,
      /\bprj[-_\s]?\d{3,}\b/gi,
      /\bpo[-_\s]?\d{4,}\b/gi,
      /\bwo[-_\s]?\d{4,}\b/gi,
      /\bso[-_\s]?\d{4,}\b/gi,
    ],
    emailLike: [
      /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi,
    ],
  },
};

// helper: count regex hits (bounded)
function countHits(text, patterns, max = 50) {
  if (!text || !patterns?.length) return 0;
  let hits = 0;
  for (const re of patterns) {
    const m = text.match(re);
    if (m?.length) hits += m.length;
    if (hits >= max) return max;
  }
  return hits;
}

// Determine Part 2 business flags using:
// - structural detections (existing Part 1 signals reused)
// - deterministic text rules (no AI)
async function buildBusinessRiskFlags({ ext, buffer, analysisResult, detections }) {
  const flags = [];
  const text = await getDeterministicText(ext, buffer, analysisResult);

  // -------------------------
  // Margin Exposure Risk
  // -------------------------
  const hasExcelLogic =
    (detections?.sensitiveFormulas?.length || 0) > 0 ||
    (detections?.excelHiddenData?.length || 0) > 0 ||
    (detections?.hiddenSheets?.length || 0) > 0;

  if (ext === "xlsx" && hasExcelLogic) {
    const level = (detections?.hiddenSheets?.length || 0) > 0 ? BIZ_SEVERITY.HIGH : BIZ_SEVERITY.MEDIUM;
    flags.push(
      makeBizFlag({
        category: "margin",
        level,
        ruleId: "MARGIN_EXCEL_PRICING_STRUCTURES",
        reason: "Pricing-related spreadsheet structures detected that could allow a customer to infer margin logic.",
        location: "Workbook",
        evidence: {
          hiddenSheets: detections?.hiddenSheets?.length || 0,
          formulas: detections?.sensitiveFormulas?.length || 0,
          excelHiddenData: detections?.excelHiddenData?.length || 0,
        },
        source: { group: "xlsx", signals: ["hiddenSheets", "sensitiveFormulas", "excelHiddenData"] },
      })
    );
  }

  // keyword-only support (doc/ppt/pdf) if text is available
  const marginHits = countHits(text, RULES.margin.pricingKeywords, 30);
  if (marginHits >= 6) {
    flags.push(
      makeBizFlag({
        category: "margin",
        level: BIZ_SEVERITY.MEDIUM,
        ruleId: "MARGIN_PRICING_LANGUAGE",
        reason: "Margin-related language detected that may require review before external sharing.",
        location: "Document Text",
        evidence: { hits: marginHits },
        source: { group: "textRules", ruleSet: "margin.pricingKeywords" },
      })
    );
  }

  // -------------------------
  // Delivery & Commitment Risk
  // -------------------------
  // deterministic: strong engagement language + lack of dependency markers => higher risk
  const strongHits = countHits(text, RULES.delivery.strongEngagement, 20);
  const openHits = countHits(text, RULES.delivery.openEndedDeliverables, 20);
  const fixedHits = countHits(text, RULES.delivery.fixedPriceNoBoundaries, 10);
  const deadlineHits = countHits(text, RULES.delivery.deadlinesNoDependencies, 20);
  const depHits = countHits(text, RULES.delivery.dependencyMarkers, 20);

  if (strongHits + openHits + fixedHits + deadlineHits > 0) {
    // escalate if there is commitment language AND no dependency markers
    const escalator = depHits === 0 && (strongHits + fixedHits + deadlineHits) > 0;
    const level =
      escalator || fixedHits > 0
        ? BIZ_SEVERITY.HIGH
        : strongHits + openHits + deadlineHits >= 4
        ? BIZ_SEVERITY.MEDIUM
        : BIZ_SEVERITY.LOW;

    flags.push(
      makeBizFlag({
        category: "delivery",
        level,
        ruleId: "DELIVERY_COMMITMENT_LANGUAGE",
        reason: "Commitment language detected without clear delivery boundaries.",
        location: "Document Text",
        evidence: { strongHits, openHits, fixedHits, deadlineHits, dependencyMarkers: depHits },
        source: { group: "textRules", ruleSet: "delivery.*" },
      })
    );
  }

  // -------------------------
  // Negotiation Power Leakage
  // -------------------------
  const assumHits = countHits(text, RULES.negotiation.internalAssumptions, 15);
  const optHits = countHits(text, RULES.negotiation.optionsABC, 15);
  const depClientHits = countHits(text, RULES.negotiation.clientDependency, 10);
  const benchHits = countHits(text, RULES.negotiation.internalBenchmarks, 15);

  // also reuse hiddenContent/metadata as negotiation leakage proxies
  const metaCount = detections?.metadata?.length || 0;
  const hiddenCount = detections?.hiddenContent?.length || 0;

  if (assumHits + optHits + depClientHits + benchHits > 0 || metaCount > 0 || hiddenCount > 0) {
    const level =
      optHits > 0 || benchHits > 0
        ? BIZ_SEVERITY.HIGH
        : (assumHits + depClientHits >= 3) || hiddenCount > 0
        ? BIZ_SEVERITY.MEDIUM
        : BIZ_SEVERITY.LOW;

    flags.push(
      makeBizFlag({
        category: "negotiation",
        level,
        ruleId: "NEGOTIATION_LEVERAGE_SIGNALS",
        reason: "Internal assumptions or negotiation structures detected that could reduce negotiation leverage.",
        location: "Document",
        evidence: { assumHits, optHits, depClientHits, benchHits, metadata: metaCount, hiddenContent: hiddenCount },
        source: { group: "textRules+struct", ruleSet: "negotiation.* + metadata/hiddenContent" },
      })
    );
  }

  // -------------------------
  // Compliance & Confidentiality Risk
  // -------------------------
  const complianceDetected = (detections?.complianceRisks?.length || 0) > 0;
  const piiDetected = (detections?.sensitiveData?.length || 0) > 0;

  const confHits = countHits(text, RULES.compliance.confidentialMarkers, 15);
  const codeHits = countHits(text, RULES.compliance.projectCodeLike, 20);
  const emailHits = countHits(text, RULES.compliance.emailLike, 30);

  if (complianceDetected || piiDetected || confHits > 0 || codeHits > 0 || emailHits > 0) {
    const level = complianceDetected || piiDetected ? BIZ_SEVERITY.CRITICAL : (confHits + codeHits + emailHits >= 5 ? BIZ_SEVERITY.HIGH : BIZ_SEVERITY.MEDIUM);

    flags.push(
      makeBizFlag({
        category: "compliance",
        level,
        ruleId: "COMPLIANCE_IDENTIFIER_EXPOSURE",
        reason: "Confidential identifiers detected in a client-facing document.",
        location: "Document",
        evidence: {
          complianceRisks: detections?.complianceRisks?.length || 0,
          sensitiveData: detections?.sensitiveData?.length || 0,
          confHits,
          projectCodeHits: codeHits,
          emailHits,
        },
        source: { group: "detections+textRules", ruleSet: "compliance.*" },
      })
    );
  }

  // -------------------------
  // Professional Credibility Risk
  // -------------------------
  const comments = detections?.comments?.length || 0;
  const track = detections?.trackChanges?.length || 0;
  const spell = detections?.spellingErrors?.length || 0;
  const orphan = detections?.orphanData?.length || 0;
  const hiddenStruct = (detections?.hiddenContent?.length || 0) + (detections?.hiddenSheets?.length || 0);

  const credTotal = comments + track + spell + orphan + hiddenStruct;
  if (credTotal > 0) {
    const level =
      (comments + track) > 0
        ? BIZ_SEVERITY.HIGH
        : credTotal >= 8
        ? BIZ_SEVERITY.MEDIUM
        : BIZ_SEVERITY.LOW;

    flags.push(
      makeBizFlag({
        category: "credibility",
        level,
        ruleId: "CREDIBILITY_DRAFT_ARTIFACTS",
        reason: "Internal draft artifacts detected that may impact professional credibility.",
        location: "Document",
        evidence: { comments, trackChanges: track, spelling: spell, formattingArtifacts: orphan, hiddenStructural: hiddenStruct },
        source: { group: "detections", signals: ["comments", "trackChanges", "spellingErrors", "orphanData", "hiddenContent/hiddenSheets"] },
      })
    );
  }

  return flags;
}

// Executive scoring model (25/25/25/25) + compliance critical gate + client-ready decision
function summarizeBusinessRisk(flags) {
  const byCat = {
    margin: { level: BIZ_SEVERITY.NONE, flags: [] },
    delivery: { level: BIZ_SEVERITY.NONE, flags: [] },
    negotiation: { level: BIZ_SEVERITY.NONE, flags: [] },
    compliance: { level: BIZ_SEVERITY.NONE, flags: [] },
    credibility: { level: BIZ_SEVERITY.NONE, flags: [] },
  };

  for (const f of flags || []) {
    if (!byCat[f.category]) continue;
    byCat[f.category].flags.push(f);
    if (bizRank(f.level) > bizRank(byCat[f.category].level)) byCat[f.category].level = f.level;
  }

  // Weighted score (0-100) using 4 categories as per spec section 4
  // Map levels to points deterministically
  const levelToScore = (lvl) => {
    const s = String(lvl || "").toLowerCase();
    if (s === "critical") return 0;
    if (s === "high") return 25;
    if (s === "medium") return 60;
    if (s === "low") return 85;
    return 100; // none
  };

  const weighted =
    levelToScore(byCat.margin.level) * QUALION_V1.scoringWeights.margin +
    levelToScore(byCat.delivery.level) * QUALION_V1.scoringWeights.delivery +
    levelToScore(byCat.negotiation.level) * QUALION_V1.scoringWeights.negotiation +
    levelToScore(byCat.credibility.level) * QUALION_V1.scoringWeights.credibility;

  const businessRiskScore = Math.round(weighted);

  // Client-ready logic (per spec):
  // - Any CRITICAL signal => NO
  // - Otherwise if any category is HIGH => NO
  const anyCritical = flags.some((f) => String(f.level).toLowerCase() === "critical");
  const anyHigh =
    ["margin", "delivery", "negotiation", "compliance", "credibility"].some(
      (k) => String(byCat[k].level || "").toLowerCase() === "high"
    );

  const clientReady = anyCritical || anyHigh ? "NO" : "YES";

  // Executive neutral recommendation
  const recommendation =
    clientReady === "NO"
      ? "Fix flagged items and re-run Qualion Clean, or acknowledge and override with audit logging."
      : "No blocking business risks detected. Document is suitable for client-facing export.";

  // Blocking issues (top 10 high/critical)
  const blocking = (flags || [])
    .filter((f) => ["high", "critical"].includes(String(f.level || "").toLowerCase()))
    .sort((a, b) => bizRank(b.level) - bizRank(a.level))
    .slice(0, 10);

  return {
    clientReady,
    businessRiskScore,
    byCategory: {
      margin: { level: byCat.margin.level, count: byCat.margin.flags.length },
      delivery: { level: byCat.delivery.level, count: byCat.delivery.flags.length },
      negotiation: { level: byCat.negotiation.level, count: byCat.negotiation.flags.length },
      compliance: { level: byCat.compliance.level, count: byCat.compliance.flags.length },
      credibility: { level: byCat.credibility.level, count: byCat.credibility.flags.length },
    },
    blockingFlags: blocking,
    recommendation,
    override: {
      allowed: true,
      requiredForExportWhenClientReadyNo: true,
      // backend does not apply override by itself; front should submit and store audit event in DB
      auditFields: ["userId", "documentId", "timestamp", "reason", "acknowledgedFlags"],
    },
  };
}

// Build Qualion Clean V1 report payload (Part1 + Part2)
function buildQualionCleanV1Report({ documentId, fileName, ext, detections, businessFlags, businessSummary }) {
  // Part 1 already exists in your UI; we still provide a clean structure to support new site update
  const part1 = {
    title: "Technical & Content Hygiene Report",
    defaultView: true,
    style: "checklist_neutral",
    signals: {
      hiddenContentDetected:
        (detections?.hiddenContent?.length || 0) > 0 ||
        (detections?.hiddenSheets?.length || 0) > 0 ||
        (detections?.excelHiddenData?.length || 0) > 0,
      commentsDetected: (detections?.comments?.length || 0) > 0,
      trackChangesDetected: (detections?.trackChanges?.length || 0) > 0,
      metadataDetected: (detections?.metadata?.length || 0) > 0,
      embeddedObjectsDetected: (detections?.embeddedObjects?.length || 0) > 0,
      structuralIssuesDetected: (detections?.visualObjects?.length || 0) > 0,
      formattingAnomaliesDetected:
        (detections?.orphanData?.length || 0) > 0 || (detections?.spellingErrors?.length || 0) > 0,
    },
  };

  const part2 = {
    title: "Business Risk Report",
    collapsedByDefault: true,
    clientReady: businessSummary.clientReady,
    businessRiskScore: businessSummary.businessRiskScore,
    categories: [
      {
        key: "margin",
        title: QUALION_V1.businessRiskCategories.margin.title,
        riskLevel: businessSummary.byCategory.margin.level,
        definition: QUALION_V1.businessRiskCategories.margin.definition,
        executiveSummary: [
          businessSummary.byCategory.margin.level === BIZ_SEVERITY.NONE
            ? "No margin exposure signals detected."
            : "Margin-related signals detected that could allow a customer to infer pricing or margin logic.",
        ].slice(0, 3),
      },
      {
        key: "delivery",
        title: QUALION_V1.businessRiskCategories.delivery.title,
        riskLevel: businessSummary.byCategory.delivery.level,
        definition: QUALION_V1.businessRiskCategories.delivery.definition,
        executiveSummary: [
          businessSummary.byCategory.delivery.level === BIZ_SEVERITY.NONE
            ? "No delivery commitment signals detected."
            : "Strong engagement language without boundaries detected.",
        ].slice(0, 3),
      },
      {
        key: "negotiation",
        title: QUALION_V1.businessRiskCategories.negotiation.title,
        riskLevel: businessSummary.byCategory.negotiation.level,
        definition: QUALION_V1.businessRiskCategories.negotiation.definition,
        executiveSummary: [
          businessSummary.byCategory.negotiation.level === BIZ_SEVERITY.NONE
            ? "No negotiation leverage leakage signals detected."
            : "Internal assumptions or negotiation structures detected that could reduce leverage.",
        ].slice(0, 3),
      },
      {
        key: "compliance",
        title: QUALION_V1.businessRiskCategories.compliance.title,
        riskLevel: businessSummary.byCategory.compliance.level,
        definition: QUALION_V1.businessRiskCategories.compliance.definition,
        executiveSummary: [
          businessSummary.byCategory.compliance.level === BIZ_SEVERITY.NONE
            ? "No compliance or confidentiality signals detected."
            : "Confidential identifiers detected in a client-facing document.",
        ].slice(0, 3),
      },
      {
        key: "credibility",
        title: QUALION_V1.businessRiskCategories.credibility.title,
        riskLevel: businessSummary.byCategory.credibility.level,
        definition: QUALION_V1.businessRiskCategories.credibility.definition,
        executiveSummary: [
          businessSummary.byCategory.credibility.level === BIZ_SEVERITY.NONE
            ? "No professionalism or credibility signals detected."
            : "Draft artifacts detected that may impact professional credibility.",
        ].slice(0, 3),
      },
    ],
    blockingFlags: businessSummary.blockingFlags,
    recommendation: businessSummary.recommendation,
    governance: {
      noAdvice: true,
      deterministicOnly: true,
      decisionPaths: ["Fix and re-run", "Acknowledge risk and override", "Export client-ready version"],
      override: businessSummary.override,
    },
  };

  return {
    meta: {
      product: "Qualion Clean",
      version: "V1",
      documentId,
      fileName,
      fileType: ext,
      analyzedAt: new Date().toISOString(),
      rule: "Any document shared externally must pass Qualion Clean.",
      supportedFileTypes: QUALION_V1.supportedFileTypes,
    },
    fileTypeContext: QUALION_V1.fileTypeFeatures[ext] || null,
    part1,
    part2,
    // For debug or analytics (front can ignore)
    internal: {
      businessFlagsCount: businessFlags.length,
    },
  };
}

// ============================================================
// 3) REPORT ZIP HELPERS
// ============================================================

function addReportsToZip(zip, single, base, reportParams) {
  const reportHtml = buildReportHtmlDetailed(reportParams);
  zip.addFile(outName(single, base, "report.html"), Buffer.from(reportHtml, "utf8"));

  const reportJson = buildReportData(reportParams);
  zip.addFile(outName(single, base, "report.json"), Buffer.from(JSON.stringify(reportJson, null, 2), "utf8"));
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

// ============================================================
// HEALTH
// ============================================================

app.get("/health", (_, res) =>
  res.json({
    ok: true,
    service: "Qualion-Doc Backend",
    version: "3.3.0",
    endpoints: ["/analyze", "/clean", "/rephrase"],
    features: [
      "qualion-clean-v1-part2-business-risk",
      "5-business-risk-categories",
      "deterministic-language-rules-pdf",
      "override-ready-payload",
      "no-ai-guessing",
      "non-breaking-output",
    ],
    time: new Date().toISOString(),
  })
);

// ===================================================================
// POST /analyze - VERSION 3.3.0
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
    const detections = analysisResult.detections || {};

    const rawSummary = analysisResult.summary;
    const summary = {
      totalIssues: rawSummary.totalIssues,
      critical: rawSummary.criticalIssues,
      high: rawSummary.highIssues,
      medium: rawSummary.mediumIssues,
      low: rawSummary.lowIssues,
    };

    const { score: riskScore, breakdown } = calculateRiskScore(summary, detections);

    const docStats = await safeExtractDocStats(req.file.buffer, ext);
    const documentStats =
      docStats && (docStats.pages || docStats.slides || docStats.sheets || docStats.tables)
        ? docStats
        : analysisResult.documentStats || docStats;

    // ✅ NEW: Qualion Clean V1 Part 2 flags + summary
    const businessFlags = await buildBusinessRiskFlags({
      ext,
      buffer: req.file.buffer,
      analysisResult,
      detections,
    });
    const businessSummary = summarizeBusinessRisk(businessFlags);

    // ✅ NEW: Combined Qualion Clean V1 report (Part1 + Part2)
    const qualionCleanV1 = buildQualionCleanV1Report({
      documentId,
      fileName: req.file.originalname,
      ext,
      detections,
      businessFlags,
      businessSummary,
    });

    res.json({
      documentId,
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

      // ✅ NEW: Part 2 business risk outputs (clean + deterministic)
      businessRisk: {
        flags: businessFlags,
        summary: businessSummary,
      },

      // ✅ NEW: full Qualion Clean V1 report object (for your site / UI update)
      qualionCleanV1,

      processingTime: Date.now() - startTime,
    });
  } catch (e) {
    console.error("[ANALYZE ERROR]", e);
    res.status(500).json({ error: String(e?.message || e) });
  }
});

// ===================================================================
// POST /clean - VERSION 3.3.0
// (unchanged behavior, but still returns report.zip with analysis payload inside report.json)
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

    const single = files.length === 1;
    const zip = new AdmZip();

    for (const f of files) {
      const ext = getExt(f.originalname);
      const base = path.parse(f.originalname).name;

      const documentStatsBefore = await safeExtractDocStats(f.buffer, ext);

      // analysis (optional)
      let analysisResult = null;
      let spellingErrors = [];
      let beforeRiskScore = 100;
      let riskBreakdown = {};
      let detections = null;
      let summary = null;

      // NEW: business risk payload for report
      let businessFlags = [];
      let businessSummary = null;
      let qualionCleanV1 = null;

      try {
        const fileType = getMimeFromExt(ext);
        const fullAnalysis = await analyzeDocument(f.buffer, fileType);
        detections = fullAnalysis.detections || {};

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

        businessFlags = await buildBusinessRiskFlags({
          ext,
          buffer: f.buffer,
          analysisResult: fullAnalysis,
          detections,
        });
        businessSummary = summarizeBusinessRisk(businessFlags);

        qualionCleanV1 = buildQualionCleanV1Report({
          documentId: uuidv4(),
          fileName: f.originalname,
          ext,
          detections,
          businessFlags,
          businessSummary,
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
          },
          businessRisk: { flags: businessFlags, summary: businessSummary },
          qualionCleanV1,
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
// POST /rephrase - VERSION 3.3.0
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

      // NEW: business risk payload for report
      let businessFlags = [];
      let businessSummary = null;
      let qualionCleanV1 = null;

      try {
        const fileType = getMimeFromExt(ext);
        const fullAnalysis = await analyzeDocument(f.buffer, fileType);
        detections = fullAnalysis.detections || {};

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

        businessFlags = await buildBusinessRiskFlags({
          ext,
          buffer: f.buffer,
          analysisResult: fullAnalysis,
          detections,
        });
        businessSummary = summarizeBusinessRisk(businessFlags);

        qualionCleanV1 = buildQualionCleanV1Report({
          documentId: uuidv4(),
          fileName: f.originalname,
          ext,
          detections,
          businessFlags,
          businessSummary,
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
          },
          businessRisk: { flags: businessFlags, summary: businessSummary },
          qualionCleanV1,
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
  console.log(`✅ Qualion-Doc Backend v3.3.0 listening on port ${PORT}`);
  console.log(`   Endpoints: GET /health, POST /analyze, POST /clean, POST /rephrase`);
  console.log(`   Features: Qualion Clean V1 Part2 (Business Risk 5 cats + executive scoring + override payload)`);
});
