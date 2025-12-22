// lib/report.js
// Qualion-Doc - v3.4 Enterprise-grade (compatible analyzer v3.2 + v3.4 Business Risks)
// Fixes:
// - Compatible avec analysis.summary.riskScore / criticalIssues / categoryCounts / recommendations(message)
// - Recommendations: supporte {priority,message,category}
// - Hidden Content: supporte hidden_text_details.items (texte + location + reason)
// - Hidden Content: counts cohérents (Executive Overview + Cleaning Summary)
// - UUID fallback si crypto.randomUUID indisponible
// - ✅ NEW: Business Risks Report (riskObjects + riskSummary) with 5 exec categories + Client-Ready gate
// - ✅ NEW (Phase 3 spec): include analysis.businessRiskAssessment in returned report JSON (businessRiskAssessment)

import crypto from "crypto";

// ============================================================
// HELPERS
// ============================================================

const fmt = (n) => (typeof n === "number" && !Number.isNaN(n) ? n : 0);

const uuid = () => {
  if (crypto.randomUUID) return crypto.randomUUID();
  return crypto.randomBytes(16).toString("hex");
};

function escapeHtml(s = "") {
  if (s === null || s === undefined) return "";
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getScoreStyle(score) {
  if (score >= 90) return { color: "#22c55e", label: "Safe", bg: "#dcfce7" };
  if (score >= 70) return { color: "#84cc16", label: "Low Risk", bg: "#ecfccb" };
  if (score >= 50) return { color: "#eab308", label: "Medium Risk", bg: "#fef9c3" };
  if (score >= 25) return { color: "#f97316", label: "High Risk", bg: "#ffedd5" };
  return { color: "#ef4444", label: "Critical Risk", bg: "#fee2e2" };
}

function getComplianceStatus(afterScore, criticalIssues = 0) {
  if (afterScore >= 90 && criticalIssues === 0) return "safe";
  if (afterScore >= 50) return "attention";
  return "not-ready";
}

function getRiskLevelEmoji(level) {
  switch (String(level || "").toLowerCase()) {
    case "critical":
      return "🔥";
    case "high":
      return "🛑";
    case "medium":
      return "🔶";
    case "low":
      return "🟢";
    default:
      return "🟡";
  }
}

function getRiskLevelColor(level) {
  switch (String(level || "").toLowerCase()) {
    case "critical":
      return "#dc2626";
    case "high":
      return "#ef4444";
    case "medium":
      return "#f59e0b";
    case "low":
      return "#10b981";
    default:
      return "#6b7280";
  }
}

function getTypeBadge(type) {
  const badges = {
    fragmented_word: { icon: "🔗", label: "Fragment", class: "type-fragment" },
    fragment: { icon: "🔗", label: "Fragment", class: "type-fragment" },
    spelling: { icon: "✏️", label: "Spelling", class: "type-spelling" },
    grammar: { icon: "📝", label: "Grammar", class: "type-grammar" },
    punctuation: { icon: "⚫", label: "Punct.", class: "type-punctuation" },
    ai_correction: { icon: "🤖", label: "AI Fix", class: "type-ai" },
    ai: { icon: "🤖", label: "AI Fix", class: "type-ai" },
  };
  return badges[type] || badges.spelling;
}

/**
 * Extrait le texte propre d'un commentaire (gère objets JSON & strings)
 */
function extractCommentText(comment) {
  if (!comment) return "";

  if (typeof comment === "string") {
    if (comment.startsWith("{") || comment.startsWith("[")) {
      try {
        const parsed = JSON.parse(comment);
        return extractCommentText(parsed);
      } catch {
        return comment;
      }
    }
    return comment;
  }

  if (typeof comment === "object") {
    if (comment.text) return String(comment.text);
    if (comment.content) return String(comment.content);
    if (comment.comment) return String(comment.comment);
    if (comment["w:t"]) return String(comment["w:t"]);

    if (comment["w:p"] || comment.p) {
      const para = comment["w:p"] || comment.p;
      return extractTextFromWordParagraph(para);
    }

    if (Array.isArray(comment)) {
      return comment.map(extractCommentText).filter(Boolean).join(" ");
    }

    for (const key of Object.keys(comment)) {
      if (key === "t" || key === "w:t" || key === "text" || key === "_") {
        const val = comment[key];
        if (typeof val === "string") return val;
        if (Array.isArray(val))
          return val
            .map((v) => (typeof v === "string" ? v : v._ || v.text || ""))
            .join("");
      }
    }
  }

  return "";
}

/**
 * Extrait le texte d'un paragraphe Word XML
 */
function extractTextFromWordParagraph(para) {
  if (!para) return "";

  const texts = [];
  const processNode = (node) => {
    if (typeof node === "string") {
      texts.push(node);
      return;
    }
    if (Array.isArray(node)) {
      node.forEach(processNode);
      return;
    }
    if (typeof node === "object") {
      if (node["w:t"] || node.t || node._) {
        const t = node["w:t"] || node.t || node._;
        texts.push(typeof t === "string" ? t : "");
      }
      if (node["w:r"]) processNode(node["w:r"]);
      if (node.r) processNode(node.r);
    }
  };

  processNode(para);
  return texts.join("").trim();
}

function formatDate(dateStr) {
  if (!dateStr) return "";
  try {
    const date = new Date(dateStr);
    if (isNaN(date.getTime())) return dateStr;
    return date.toLocaleDateString("en-US", { day: "numeric", month: "short", year: "numeric" });
  } catch {
    return dateStr;
  }
}

// ============================================================
// BUSINESS RISKS HELPERS (v3.4)
// ============================================================

const severityRank = (s) => {
  const v = String(s || "low").toLowerCase();
  if (v === "critical") return 4;
  if (v === "high") return 3;
  if (v === "medium") return 2;
  return 1;
};

const severityLabel = (s) => {
  const v = String(s || "low").toLowerCase();
  return v.charAt(0).toUpperCase() + v.slice(1);
};

const severityMax = (a, b) => (severityRank(b) > severityRank(a) ? b : a);

const execBusinessCategories = [
  { key: "margin", label: "Margin Exposure Risk" },
  { key: "delivery", label: "Delivery & Commitment Risk" },
  { key: "negotiation", label: "Negotiation Power Leakage" },
  { key: "compliance", label: "Compliance & Confidentiality Risk" },
  { key: "credibility", label: "Professional Credibility Risk" },
];

// ============================================================
// BUILD REPORT DATA (JSON) - Enterprise-grade
// ============================================================

export function buildReportData({
  filename,
  ext,
  policy,
  cleaning,
  correction,
  analysis,
  spellingErrors = [],
  beforeRiskScore,
  afterRiskScore,
  scoreImpacts = {},
  user = null,
  organization = null,
}) {
  // ============================================================
  // SCORES (compat analyzer v3.2)
  // analyzer: analysis.summary.riskScore (0..100)
  // ============================================================
  const analyzerScore = analysis?.summary?.riskScore; // "safe score" (100 = safe)
  const beforeScore =
    beforeRiskScore ??
    analysis?.summary?.beforeRiskScore ??
    (typeof analyzerScore === "number" ? analyzerScore : 50);

  const afterScore =
    afterRiskScore ??
    analysis?.summary?.afterRiskScore ??
    Math.min(100, Math.max(0, beforeScore + 25));

  const riskReduction =
    beforeScore < 100 ? Math.round(((afterScore - beforeScore) / (100 - beforeScore)) * 100) : 0;

  const totalIssues = fmt(analysis?.summary?.totalIssues);
  const criticalIssues = fmt(analysis?.summary?.criticalIssues);
  const recommendationsRaw = analysis?.summary?.recommendations || [];

  const complianceStatus = getComplianceStatus(afterScore, criticalIssues);

  // ============================================================
  // COLLECTER corrections uniques
  // ============================================================
  const allCorrections = [];
  const seenCorrections = new Set();

  const addCorrection = (before, after, type, context, location) => {
    const beforeStr = String(before || "").trim();
    const afterStr = String(after || "").trim();
    if (!beforeStr || !afterStr || beforeStr === afterStr) return;

    const key = `${beforeStr}|${afterStr}`;
    if (seenCorrections.has(key)) return;
    seenCorrections.add(key);

    allCorrections.push({
      id: uuid(),
      original: beforeStr,
      corrected: afterStr,
      type: type || "spelling",
      context: context || "",
      location: location || "Document",
    });
  };

  if (Array.isArray(spellingErrors) && spellingErrors.length > 0) {
    spellingErrors.forEach((err) => {
      const before = err.error || err.word || err.original || err.before || "";
      const after = err.correction || err.suggestion || err.corrected || err.after || "";
      addCorrection(before, after, err.type || "spelling", err.context || "", err.location || "");
    });
  }

  if (correction?.examples && Array.isArray(correction.examples)) {
    correction.examples.forEach((ex) => {
      const before = ex.before || ex.error || ex.original || "";
      const after = ex.after || ex.correction || ex.corrected || "";
      addCorrection(before, after, ex.type || "ai_correction", "", "");
    });
  }

  // ============================================================
  // ITEMS DÉTAILLÉS PAR CATÉGORIE (compat analyzer v3.2)
  // ============================================================

  // 1) Metadata
  const metadataItems = [];
  if (Array.isArray(analysis?.detections?.metadata)) {
    analysis.detections.metadata.forEach((m) => {
      const label = m.key || m.type || m.name || "Property";
      const value = m.value ? String(m.value).substring(0, 160) : undefined;
      metadataItems.push({ label, value });
    });
  }

  // 2) Comments (texte complet)
  const commentItems = [];
  if (Array.isArray(analysis?.detections?.comments)) {
    analysis.detections.comments.forEach((c) => {
      const author = c.author || c.authorName || "Unknown author";
      const date = formatDate(c.date);

      let commentText = "";
      if (c.text) commentText = extractCommentText(c.text);
      else if (c.content) commentText = extractCommentText(c.content);
      else if (c.comment) commentText = extractCommentText(c.comment);

      commentText = String(commentText || "").trim();
      if (!commentText || commentText === "[object Object]") commentText = "Comment removed";

      const label = date ? `${author} (${date})` : author;
      commentItems.push({ label, value: commentText });
    });
  }

  // 3) Track changes
  const trackChangesItems = [];
  if (Array.isArray(analysis?.detections?.trackChanges)) {
    analysis.detections.trackChanges.forEach((tc) => {
      const type = tc.type || tc.changeType || "change";
      const typeEmoji =
        type === "deletion" ? "🔴 Deleted:" : type === "insertion" ? "🟢 Added:" : "🔄 Changed:";
      const original = tc.originalText || "";
      const newText = tc.newText || "";

      let displayText = "";
      if (original && newText) displayText = `"${original}" → "${newText}"`;
      else if (tc.text) displayText = tc.text;
      else if (original) displayText = `"${original}"`;
      else if (newText) displayText = `"${newText}"`;

      const author = tc.author || "Unknown";
      trackChangesItems.push({
        label: `${typeEmoji} by ${author}`,
        value: displayText.length > 220 ? displayText.substring(0, 220) + "..." : displayText,
      });
    });
  }

  // 4) Hidden content (support hidden_text_details.items)
  const hiddenContentItems = [];
  let hiddenContentDetailedCount = 0;

  if (Array.isArray(analysis?.detections?.hiddenContent)) {
    analysis.detections.hiddenContent.forEach((hc) => {
      const type = hc.type || hc.elementType || "hidden_element";

      const typeLabels = {
        vanished_text: "Hidden text (vanish)",
        white_text: "White/invisible text",
        invisible_text: "Invisible text",
        hidden_slide: "Hidden slide",
        off_slide_content: "Off-slide content",
        embedded_file: "Embedded file",
        hidden_row: "Hidden row",
        hidden_column: "Hidden column",
        hidden_sheet: "Hidden sheet",
        hidden_layer: "Hidden layer (PDF)",
        pdf_javascript: "PDF JavaScript",
      };

      // ✅ Détails
      if (type === "hidden_text_details" && Array.isArray(hc.items) && hc.items.length > 0) {
        hc.items.forEach((it) => {
          hiddenContentDetailedCount++;

          const reasonLabel =
            {
              vanish: "vanish",
              white_color: "white",
              tiny_font: "tiny",
              off_slide: "off-slide",
              hidden_style: "hidden-style",
            }[it.reason] || (it.reason || "hidden");

          const loc = it.location || hc.location || "Document";
          const textShown = String(it.text || it.preview || "").trim();
          const value = textShown.length > 320 ? textShown.slice(0, 317) + "..." : textShown;

          hiddenContentItems.push({
            label: `Hidden text (${reasonLabel}) — ${loc}`,
            value,
            reason: it.reason || null,
            location: loc,
            color: it.color || null,
            fontSizePt: it.fontSizePt || null,
            fontSizeHalfPoints: it.fontSizeHalfPoints || null,
          });
        });

        return; // pas de résumé additionnel
      }

      // ✅ Résumés
      const label = typeLabels[type] || "Hidden element";
      let value = "";

      if (hc.content && typeof hc.content === "string") {
        value = hc.content.length > 320 ? hc.content.substring(0, 320) + "..." : hc.content;
      } else if (hc.description) {
        value = hc.description;
      } else if (hc.count) {
        value = `${hc.count} element(s) found`;
      } else if (hc.location) {
        value = hc.location;
      }

      hiddenContentItems.push({ label, value });
    });
  }

  // 5) Macros
  const macroItems = [];
  if (Array.isArray(analysis?.detections?.macros)) {
    analysis.detections.macros.forEach((m) => {
      const type = m.type || "VBA Macro";
      macroItems.push({
        label: `⚠️ ${type}`,
        value: m.description || m.name || "Executable code detected",
      });
    });
  }

  // 6) Sensitive Data (try to show full value when available)
  const sensitiveDataItems = [];
  const sensitiveDataTypes = new Set();
  if (Array.isArray(analysis?.detections?.sensitiveData)) {
    analysis.detections.sensitiveData.forEach((sd) => {
      const type = sd.type || sd.dataType || "Sensitive";
      sensitiveDataTypes.add(type);

      // IMPORTANT:
      // analyzer v3.2 renvoie souvent sd.value = masked || value
      // si tu ajoutes plus tard sd.rawValue / originalValue, on l'affiche automatiquement.
      const fullValue = sd.rawValue || sd.originalValue || sd.value || sd.match || sd.text || "";

      sensitiveDataItems.push({
        label: type.replace(/_/g, " ").toUpperCase(),
        value: String(fullValue),
      });
    });
  }

  // 7) Visual objects
  const visualObjectsItems = [];
  if (Array.isArray(analysis?.detections?.visualObjects)) {
    analysis.detections.visualObjects.forEach((vo) => {
      const type = vo.type || "visual_object";
      const typeLabel =
        {
          shape_covering_text: "Shape covering text",
          missing_alt_text: "Missing alt text",
        }[type] || "Visual object";

      visualObjectsItems.push({
        label: typeLabel,
        value: vo.description || "",
      });
    });
  }

  // 8) Orphan data
  const orphanDataItems = [];
  if (Array.isArray(analysis?.detections?.orphanData)) {
    analysis.detections.orphanData.forEach((od) => {
      const type = od.type || "orphan_data";
      const typeLabel =
        {
          broken_link: "Broken link",
          empty_page: "Empty page",
          trailing_whitespace: "Trailing whitespace",
        }[type] || "Orphan data";

      orphanDataItems.push({
        label: typeLabel,
        value: od.value || od.description || "",
      });
    });
  }

  // 9) Excel hidden data
  const excelHiddenDataItems = [];
  if (Array.isArray(analysis?.detections?.excelHiddenData)) {
    analysis.detections.excelHiddenData.forEach((ed) => {
      const type = ed.type || "excel_hidden";
      const typeLabel =
        {
          hidden_sheet: "Hidden sheet",
          very_hidden_sheet: "Very hidden sheet",
          hidden_column: "Hidden column",
          hidden_row: "Hidden row",
          hidden_formula: "Hidden formula",
        }[type] || "Excel hidden data";

      excelHiddenDataItems.push({
        label: `${typeLabel}: ${ed.name || ""}`.trim(),
        value: ed.description || "",
      });
    });
  }

  if (Array.isArray(analysis?.detections?.hiddenSheets)) {
    analysis.detections.hiddenSheets.forEach((hs) => {
      excelHiddenDataItems.push({
        label: `Hidden Sheet: ${hs.sheetName || "Unknown"}`,
        value: hs.type === "very_hidden" ? "Very Hidden" : "Hidden",
      });
    });
  }

  // 10) Spelling/Grammar items
  const spellingItems = allCorrections.map((c) => ({
    label: c.original,
    value: `→ ${c.corrected}`,
  }));

  // ============================================================
  // EXECUTIVE OVERVIEW (9 catégories)
  // ============================================================

  const calculateCategoryStats = (items, cleaned = null, riskLevel = "medium") => {
    const found = items.length;
    const cleanedCount = cleaned !== null ? cleaned : found;
    return {
      found,
      cleaned: cleanedCount,
      remaining: Math.max(0, found - cleanedCount),
      riskLevel,
      items,
    };
  };

  const sensitiveDataRiskLevel = sensitiveDataItems.some(
    (i) => i.label.includes("IBAN") || i.label.includes("SSN") || i.label.includes("CREDIT")
  )
    ? "critical"
    : sensitiveDataItems.length > 0
    ? "high"
    : "low";

  const executiveOverview = {
    confidentialInfo: calculateCategoryStats(
      sensitiveDataItems,
      fmt(cleaning?.sensitiveDataMasked) || sensitiveDataItems.length,
      sensitiveDataRiskLevel
    ),
    metadataExposure: calculateCategoryStats(
      metadataItems,
      fmt(cleaning?.metaRemoved) || metadataItems.length,
      metadataItems.length > 3 ? "medium" : "low"
    ),
    commentsReview: calculateCategoryStats(
      [...commentItems, ...trackChangesItems],
      (fmt(cleaning?.commentsXmlRemoved) + fmt(cleaning?.commentMarkersRemoved)) ||
        (commentItems.length + trackChangesItems.length),
      commentItems.length > 0 || trackChangesItems.length > 0 ? "medium" : "low"
    ),
    hiddenContent: calculateCategoryStats(
      hiddenContentItems,
      fmt(cleaning?.hiddenRemoved) || hiddenContentDetailedCount || hiddenContentItems.length,
      hiddenContentItems.some((i) => /white|vanish/i.test(i.label)) ? "high" : "low"
    ),
    grammarTone: calculateCategoryStats(
      spellingItems.slice(0, 20),
      allCorrections.length,
      allCorrections.length > 10 ? "medium" : "low"
    ),
    visualObjects: calculateCategoryStats(
      visualObjectsItems,
      visualObjectsItems.length,
      visualObjectsItems.some((i) => i.label.includes("covering")) ? "medium" : "low"
    ),
    orphanData: calculateCategoryStats(orphanDataItems, orphanDataItems.length, "low"),
    macroThreats: calculateCategoryStats(
      macroItems,
      fmt(cleaning?.macrosRemoved) || macroItems.length,
      macroItems.length > 0 ? "critical" : "low"
    ),
    excelHiddenData: calculateCategoryStats(
      excelHiddenDataItems,
      excelHiddenDataItems.length,
      excelHiddenDataItems.length > 0 ? "medium" : "low"
    ),
  };

  // ============================================================
  // CLEANING SUMMARY
  // ============================================================

  const cleaningSummary = {
    metadata: {
      count: metadataItems.length || fmt(cleaning?.metaRemoved) || 0,
      items: metadataItems,
      scoreImpact: scoreImpacts.metadata || Math.min(10, metadataItems.length * 2),
    },
    comments: {
      count:
        commentItems.length ||
        (fmt(cleaning?.commentsXmlRemoved) + fmt(cleaning?.commentMarkersRemoved)) ||
        0,
      items: commentItems,
      scoreImpact: scoreImpacts.comments || Math.min(15, commentItems.length * 3),
    },
    trackChanges: {
      count:
        trackChangesItems.length ||
        (fmt(cleaning?.revisionsAccepted?.deletionsRemoved) +
          fmt(cleaning?.revisionsAccepted?.insertionsUnwrapped)) ||
        0,
      items: trackChangesItems,
      deletions: fmt(cleaning?.revisionsAccepted?.deletionsRemoved) || 0,
      insertions: fmt(cleaning?.revisionsAccepted?.insertionsUnwrapped) || 0,
      scoreImpact: scoreImpacts.trackChanges || Math.min(10, trackChangesItems.length * 2),
    },
    hiddenContent: {
      count:
        hiddenContentDetailedCount ||
        hiddenContentItems.length ||
        fmt(cleaning?.hiddenRemoved) ||
        0,
      items: hiddenContentItems,
      scoreImpact:
        scoreImpacts.hiddenContent ||
        Math.min(20, (hiddenContentDetailedCount || hiddenContentItems.length) * 5),
    },
    macros: {
      count: macroItems.length || fmt(cleaning?.macrosRemoved) || 0,
      items: macroItems,
      scoreImpact: scoreImpacts.macros || (macroItems.length > 0 ? 30 : 0),
    },
    sensitiveData: {
      count: sensitiveDataItems.length,
      items: sensitiveDataItems,
      types: Array.from(sensitiveDataTypes),
      scoreImpact: scoreImpacts.sensitiveData || Math.min(25, sensitiveDataItems.length * 5),
    },
    spellingGrammar: {
      count: allCorrections.length,
      items: spellingItems.slice(0, 10),
      scoreImpact: scoreImpacts.spellingGrammar || Math.min(5, allCorrections.length),
    },
    visualObjects: {
      count: visualObjectsItems.length,
      items: visualObjectsItems,
      scoreImpact: scoreImpacts.visualObjects || Math.min(5, visualObjectsItems.length * 2),
    },
    orphanData: {
      count: orphanDataItems.length,
      items: orphanDataItems,
      scoreImpact: scoreImpacts.orphanData || Math.min(3, orphanDataItems.length),
    },
    excelHiddenData: {
      count: excelHiddenDataItems.length,
      items: excelHiddenDataItems,
      scoreImpact: scoreImpacts.excelHiddenData || Math.min(15, excelHiddenDataItems.length * 4),
    },
  };

  const totalRemoved = Object.values(cleaningSummary).reduce((sum, cat) => sum + (cat.count || 0), 0);

  // ============================================================
  // RISKS DETECTED (dashboard)
  // ============================================================
  const risksDetected = [];

  if (Array.isArray(analysis?.detections?.sensitiveData)) {
    analysis.detections.sensitiveData.forEach((sd) => {
      risksDetected.push({
        id: sd.id || uuid(),
        severity: sd.severity || "high",
        type: sd.type || "sensitive_data",
        description: sd.description || `Detected ${sd.type || "sensitive data"}`,
        context: sd.context || sd.value || "",
        action: "flagged",
      });
    });
  }

  if (Array.isArray(analysis?.detections?.complianceRisks)) {
    analysis.detections.complianceRisks.forEach((cr) => {
      risksDetected.push({
        id: cr.id || uuid(),
        severity: cr.severity || "high",
        type: cr.type || "compliance",
        description: cr.description || cr.message || "Compliance risk detected",
        context: cr.context || "",
        action: "flagged",
      });
    });
  }

  if (Array.isArray(analysis?.detections?.macros) && analysis.detections.macros.length > 0) {
    risksDetected.push({
      id: uuid(),
      severity: "critical",
      type: "macros",
      description: `${analysis.detections.macros.length} macro(s) detected - potential security risk`,
      context: "Macros can contain executable code",
      action: "removed",
    });
  }

  if (Array.isArray(analysis?.detections?.hiddenContent)) {
    const criticalHidden = analysis.detections.hiddenContent.filter(
      (h) => h.type === "white_text" || h.type === "vanished_text"
    );
    criticalHidden.forEach((ch) => {
      risksDetected.push({
        id: ch.id || uuid(),
        severity: "high",
        type: "hidden_content",
        description: ch.description || "Hidden content detected",
        context: ch.content || ch.location || "",
        action: "removed",
      });
    });

    const detailed = analysis.detections.hiddenContent.find((h) => h.type === "hidden_text_details");
    if (detailed?.items?.length) {
      detailed.items.slice(0, 10).forEach((it) => {
        risksDetected.push({
          id: uuid(),
          severity: it.reason === "white_color" || it.reason === "vanish" ? "high" : "medium",
          type: "hidden_text",
          description: `Hidden text detected (${it.reason || "hidden"})`,
          context: `${it.location || "Document"} — ${(it.text || it.preview || "").toString().slice(0, 120)}`,
          action: "removed",
        });
      });
    }
  }

  // ============================================================
  // RECOMMENDATIONS (compat analyzer v3.2: {priority,message,category})
  // ============================================================
  const formattedRecommendations = [];

  if (Array.isArray(recommendationsRaw)) {
    recommendationsRaw.forEach((rec) => {
      if (typeof rec === "string") formattedRecommendations.push(rec);
      else if (rec?.message) {
        const p = rec.priority ? rec.priority.toUpperCase() : "";
        formattedRecommendations.push(`${p ? `[${p}] ` : ""}${rec.message}`);
      } else if (rec?.text) {
        formattedRecommendations.push(`${rec.icon || ""} ${rec.text}`.trim());
      }
    });
  }

  if (formattedRecommendations.length === 0) {
    if (sensitiveDataItems.length > 0)
      formattedRecommendations.push("Review all flagged sensitive data before sharing externally");
    if (macroItems.length > 0)
      formattedRecommendations.push("Remove or disable macros for external document sharing");
    if (hiddenContentItems.length > 0)
      formattedRecommendations.push("Hidden content detected — ensure it was removed before sharing");
    formattedRecommendations.push("Document analysis completed");
  }

  // ============================================================
  // ✅ BUSINESS RISKS (v3.4)
  // ============================================================
  const riskSummary = analysis?.riskSummary || null;
  const riskObjects = Array.isArray(analysis?.riskObjects) ? analysis.riskObjects : [];

  // Build category buckets
  const byCategory = {};
  execBusinessCategories.forEach((c) => {
    byCategory[c.key] = {
      key: c.key,
      label: c.label,
      count: 0,
      maxSeverity: "low",
      items: [],
      bySurface: { Visible: 0, Hidden: 0, Structural: 0, Residual: 0 },
    };
  });

  // Collect
  riskObjects.forEach((r) => {
    const k = r?.riskCategory;
    if (!k || !byCategory[k]) return;

    byCategory[k].count += 1;
    byCategory[k].maxSeverity = severityMax(byCategory[k].maxSeverity, r.severity || "low");

    const surface = r.surface || "Visible";
    if (byCategory[k].bySurface[surface] !== undefined) byCategory[k].bySurface[surface] += 1;

    byCategory[k].items.push({
      id: r.id || uuid(),
      severity: String(r.severity || "low").toLowerCase(),
      surface: surface,
      reason: r.reason || "",
      location: r.location || "",
      fixability: r.fixability || "manual",
      ruleId: r.ruleId || "",
      points: typeof r.points === "number" ? r.points : null,
    });
  });

  // Sort within each category
  Object.values(byCategory).forEach((cat) => {
    cat.items.sort((a, b) => {
      const d = severityRank(b.severity) - severityRank(a.severity);
      if (d !== 0) return d;
      const pb = typeof b.points === "number" ? b.points : 0;
      const pa = typeof a.points === "number" ? a.points : 0;
      return pb - pa;
    });
  });

  // Derive overall if missing
  const derivedOverallSeverity =
    riskSummary?.overallSeverity ||
    Object.values(byCategory).reduce((mx, c) => severityMax(mx, c.maxSeverity), "low");

  const derivedClientReady =
    riskSummary?.clientReady || (severityRank(derivedOverallSeverity) >= 3 ? "NO" : "YES");

  const businessRisks = {
    clientReady: derivedClientReady,
    overallSeverity: String(derivedOverallSeverity || "low").toLowerCase(),
    executiveSignals: Array.isArray(riskSummary?.executiveSignals) ? riskSummary.executiveSignals : [],
    totalsByCategory: riskSummary?.byCategory || null,
    blockingIssues: Array.isArray(riskSummary?.blockingIssues) ? riskSummary.blockingIssues : [],
    categories: execBusinessCategories.map((c) => byCategory[c.key]),
    totalRiskObjects:
      typeof riskSummary?.totalRiskObjects === "number" ? riskSummary.totalRiskObjects : riskObjects.length,
  };

  // ============================================================
  // ✅ NEW (Phase 3 spec): Business Risk Assessment passthrough
  // ============================================================
  const businessRiskAssessment = analysis?.businessRiskAssessment ?? null;

  // ============================================================
  // EXTRA METRICS
  // ============================================================
  const clarityScoreImprovement =
    allCorrections.length > 0 ? Math.min(25, Math.round(allCorrections.length * 2.5)) : 0;

  const documentSizeReduction =
    (fmt(cleaning?.mediaDeleted) + fmt(cleaning?.embeddedFilesRemoved)) > 0
      ? Math.min(30, (fmt(cleaning?.mediaDeleted) + fmt(cleaning?.embeddedFilesRemoved)) * 5)
      : 0;

  console.log(
    `[REPORT JSON v3.4] beforeScore=${beforeScore}, afterScore=${afterScore}, riskReduction=${riskReduction}%, complianceStatus=${complianceStatus}, corrections=${allCorrections.length}, comments=${commentItems.length}, sensitiveData=${sensitiveDataItems.length}, hiddenDetails=${hiddenContentDetailedCount}, businessRiskObjects=${businessRisks.totalRiskObjects}, businessOverall=${businessRisks.overallSeverity}, clientReady=${businessRisks.clientReady}, hasBusinessRiskAssessment=${!!businessRiskAssessment}`
  );

  // ============================================================
  // FINAL JSON (PremiumReportData)
  // ============================================================
  return {
    document: {
      name: filename,
      type: ext?.toUpperCase() || "UNKNOWN",
      processedAt: new Date().toISOString(),
      user: user || null,
      organization: organization || null,
    },
    beforeScore: Math.round(beforeScore),
    afterScore: Math.round(afterScore),
    riskReduction: Math.max(0, riskReduction),
    complianceStatus,
    summary: {
      totalIssuesFound: totalIssues,
      elementsRemoved: totalRemoved,
      correctionsApplied: allCorrections.length,
      criticalRisksResolved: criticalIssues,
    },
    executiveOverview,
    clarityScoreImprovement,
    documentSizeReduction,
    cleaningSummary,
    textCorrections: allCorrections,
    risksDetected,
    recommendations: formattedRecommendations,

    // ✅ NEW (v3.4)
    businessRisks,

    // ✅ NEW (Phase 3 spec)
    businessRiskAssessment,
  };
}

// ============================================================
// BUILD REPORT HTML - Enterprise-grade Template
// ============================================================

export function buildReportHtmlDetailed({
  filename,
  ext,
  policy,
  cleaning,
  correction,
  analysis,
  spellingErrors = [],
  beforeRiskScore,
  afterRiskScore,
  user = null,
  organization = null,
}) {
  const reportData = buildReportData({
    filename,
    ext,
    policy,
    cleaning,
    correction,
    analysis,
    spellingErrors,
    beforeRiskScore,
    afterRiskScore,
    user,
    organization,
  });

  const beforeStyle = getScoreStyle(reportData.beforeScore);
  const afterStyle = getScoreStyle(reportData.afterScore);

  // ============================================================
  // MODIF 1 — Compliance labels + subtitles
  // ============================================================
  const complianceColors = {
    safe: {
      emoji: "✅",
      label: "Ready to Send",
      subtitle:
        "This document has been cleaned and is now safe to share with clients or external partners.",
      color: "#10b981",
      bg: "#dcfce7",
    },
    attention: {
      emoji: "🟡",
      label: "Attention Required",
      subtitle: "This document requires attention before sharing externally.",
      color: "#f59e0b",
      bg: "#fef3c7",
    },
    "not-ready": {
      emoji: "🔴",
      label: "Not Client-Ready",
      subtitle: "This document contains elements that should not be shared externally.",
      color: "#ef4444",
      bg: "#fee2e2",
    },
  };

  const compliance = complianceColors[reportData.complianceStatus] || complianceColors["attention"];

  const getCategoryRow = (name, stats) => {
    if (!stats || (stats.found === 0 && stats.cleaned === 0)) return "";
    return `
      <tr>
        <td>${name}</td>
        <td class="center">${stats.found}</td>
        <td class="center">${stats.cleaned}</td>
        <td class="center">${stats.remaining}</td>
        <td class="center">
          <span class="risk-level" style="background: ${getRiskLevelColor(stats.riskLevel)}20; color: ${getRiskLevelColor(
      stats.riskLevel
    )}">
            ${getRiskLevelEmoji(stats.riskLevel)} ${
      stats.riskLevel.charAt(0).toUpperCase() + stats.riskLevel.slice(1)
    }
          </span>
        </td>
      </tr>
    `;
  };

  const executiveOverviewRows = `
    ${getCategoryRow("Confidential & Sensitive Information", reportData.executiveOverview.confidentialInfo)}
    ${getCategoryRow("Metadata Exposure", reportData.executiveOverview.metadataExposure)}
    ${getCategoryRow("Comments & Review Traces", reportData.executiveOverview.commentsReview)}
    ${getCategoryRow("Hidden & Embedded Content", reportData.executiveOverview.hiddenContent)}
    ${getCategoryRow("Grammar, Tone & Clarity", reportData.executiveOverview.grammarTone)}
    ${getCategoryRow("Visual Objects & Diagrams", reportData.executiveOverview.visualObjects)}
    ${getCategoryRow("Senseless / Orphan Data", reportData.executiveOverview.orphanData)}
    ${getCategoryRow("Macro / Script Threats", reportData.executiveOverview.macroThreats)}
    ${getCategoryRow("Excel Hidden Data", reportData.executiveOverview.excelHiddenData)}
  `;

  const getSectionItemsHtml = (items) => {
    if (!items || items.length === 0) return "";
    return `
      <ul class="items-list">
        ${items
          .slice(0, 15)
          .map(
            (item) => `
          <li>
            <div class="item-content">
              <span class="item-label">${escapeHtml(item.label)}</span>
              ${item.value ? `<span class="item-value">${escapeHtml(item.value)}</span>` : ""}
            </div>
          </li>
        `
          )
          .join("")}
        ${items.length > 15 ? `<li class="more-items">... and ${items.length - 15} more items</li>` : ""}
      </ul>
    `;
  };

  // ============================================================
  // ✅ Business Risks HTML (Part 2)
  // ============================================================
  const br = reportData.businessRisks || null;

  const businessRiskBadge = (sev) => {
    const s = String(sev || "low").toLowerCase();
    const color = getRiskLevelColor(s);
    return `<span class="risk-level" style="background:${color}20;color:${color}">${getRiskLevelEmoji(
      s
    )} ${escapeHtml(severityLabel(s))}</span>`;
  };

  const businessRisksSection = br
    ? `
    <div class="overview-card">
      <h3>🏛️ Part 2 — Business Risks (Executive)</h3>

      <div style="padding: 0 24px 18px; color: var(--gray-700); font-size: 14px;">
        <div style="display:flex; flex-wrap:wrap; gap:10px; align-items:center;">
          <div><strong>Client-Ready Gate:</strong> ${escapeHtml(br.clientReady)}</div>
          <div><strong>Overall Severity:</strong> ${businessRiskBadge(br.overallSeverity)}</div>
          <div><strong>Total Signals:</strong> ${fmt(br.totalRiskObjects)}</div>
        </div>

        ${
          Array.isArray(br.executiveSignals) && br.executiveSignals.length
            ? `<div style="margin-top:10px;"><strong>Executive Signals:</strong> ${escapeHtml(
                br.executiveSignals.join(" • ")
              )}</div>`
            : ""
        }
      </div>

      <table class="overview-table">
        <thead>
          <tr>
            <th>Category</th>
            <th class="center">Count</th>
            <th class="center">Max Severity</th>
            <th class="center">Surfaces (V/H/S/R)</th>
          </tr>
        </thead>
        <tbody>
          ${br.categories
            .map((c) => {
              const s = c.bySurface || {};
              const surf = `${fmt(s.Visible)}/${fmt(s.Hidden)}/${fmt(s.Structural)}/${fmt(
                s.Residual
              )}`;
              return `
                <tr>
                  <td>${escapeHtml(c.label)}</td>
                  <td class="center">${fmt(c.count)}</td>
                  <td class="center">${businessRiskBadge(c.maxSeverity)}</td>
                  <td class="center"><span style="font-family: 'SF Mono', Monaco, monospace;">${escapeHtml(
                    surf
                  )}</span></td>
                </tr>
              `;
            })
            .join("")}
        </tbody>
      </table>

      ${
        Array.isArray(br.blockingIssues) && br.blockingIssues.length
          ? `
        <div style="padding: 16px 24px 24px;">
          <h4 style="margin: 6px 0 10px; font-size: 14px;">🛑 Blocking Issues</h4>
          <ul class="items-list" style="margin:0;">
            ${br.blockingIssues
              .slice(0, 10)
              .map(
                (it) => `
              <li>
                <div class="item-content">
                  <span class="item-label">${escapeHtml(
                    it.label || it.riskCategory || "Issue"
                  )} — ${escapeHtml(it.severity || "")}</span>
                  <span class="item-value">${escapeHtml(it.reason || "")}</span>
                  ${
                    it.fixability
                      ? `<div style="margin-top:6px; font-size:12px; color: var(--gray-500);">Fixability: <strong>${escapeHtml(
                          it.fixability
                        )}</strong></div>`
                      : ""
                  }
                </div>
              </li>
            `
              )
              .join("")}
            ${
              br.blockingIssues.length > 10
                ? `<li class="more-items">... and ${br.blockingIssues.length - 10} more blocking issues</li>`
                : ""
            }
          </ul>
        </div>
        `
          : ""
      }

      ${br.categories
        .filter((c) => Array.isArray(c.items) && c.items.length)
        .map((c) => {
          const top = c.items.slice(0, 6);
          return `
            <section class="detail-section" style="margin: 0 24px 24px;">
              <h2>📌 ${escapeHtml(c.label)} — Top Signals</h2>
              <ul class="items-list">
                ${top
                  .map(
                    (it) => `
                  <li>
                    <div class="item-content">
                      <span class="item-label">${businessRiskBadge(it.severity)} &nbsp; ${escapeHtml(
                      it.surface || ""
                    )} ${it.location ? `— ${escapeHtml(it.location)}` : ""}</span>
                      <span class="item-value">${escapeHtml(it.reason || "")}</span>
                      <div style="margin-top:6px; font-size:12px; color: var(--gray-500);">
                        Fixability: <strong>${escapeHtml(it.fixability || "manual")}</strong>
                        ${it.ruleId ? ` • Rule: <strong>${escapeHtml(it.ruleId)}</strong>` : ""}
                        ${typeof it.points === "number" ? ` • Points: <strong>${it.points}</strong>` : ""}
                      </div>
                    </div>
                  </li>
                `
                  )
                  .join("")}
                ${
                  c.items.length > 6
                    ? `<li class="more-items">... and ${c.items.length - 6} more signals</li>`
                    : ""
                }
              </ul>
            </section>
          `;
        })
        .join("")}
    </div>
  `
    : "";

  const section1 =
    reportData.executiveOverview.confidentialInfo.found > 0
      ? `
    <section class="detail-section">
      <h2>🔒 Sensitive Data Flagged</h2>
      <p class="section-intro">The following sensitive information was detected in your document:</p>
      ${getSectionItemsHtml(reportData.cleaningSummary.sensitiveData.items)}
      <div class="action-taken">
        📌 ${reportData.cleaningSummary.sensitiveData.count} sensitive data element(s) flagged for review
      </div>
    </section>
  `
      : "";

  const section2 =
    reportData.executiveOverview.metadataExposure.found > 0
      ? `
    <section class="detail-section">
      <h2>📄 Metadata Removed</h2>
      <p class="section-intro">The following metadata was removed from your document:</p>
      ${getSectionItemsHtml(reportData.cleaningSummary.metadata.items)}
      <div class="action-taken">
        📌 ${reportData.cleaningSummary.metadata.count} metadata element(s) removed
      </div>
    </section>
  `
      : "";

  const section3 =
    reportData.cleaningSummary.comments.count > 0
      ? `
    <section class="detail-section">
      <h2>💬 Comments Cleaned</h2>
      <p class="section-intro">The following comments were removed from your document:</p>
      ${getSectionItemsHtml(reportData.cleaningSummary.comments.items)}
      <div class="action-taken">
        📌 ${reportData.cleaningSummary.comments.count} comment(s) removed
      </div>
    </section>
  `
      : "";

  const section4 =
    reportData.cleaningSummary.trackChanges.count > 0
      ? `
    <section class="detail-section">
      <h2>📝 Track Changes Removed</h2>
      <p class="section-intro">The following revision marks were accepted/removed:</p>
      ${getSectionItemsHtml(reportData.cleaningSummary.trackChanges.items)}
      <div class="action-taken">
        📌 ${reportData.cleaningSummary.trackChanges.count} track change(s) processed
      </div>
    </section>
  `
      : "";

  const section5 =
    reportData.executiveOverview.hiddenContent.found > 0
      ? `
    <section class="detail-section">
      <h2>👁️ Hidden Content Removed</h2>
      <p class="section-intro">The following hidden elements were detected and removed:</p>
      ${getSectionItemsHtml(reportData.cleaningSummary.hiddenContent.items)}
      <div class="action-taken">
        📌 ${reportData.cleaningSummary.hiddenContent.count} hidden element(s) removed
      </div>
    </section>
  `
      : "";

  const section6 =
    reportData.executiveOverview.macroThreats.found > 0
      ? `
    <section class="detail-section danger-section">
      <h2>⚠️ Macros Disabled</h2>
      <p class="section-intro">The following potentially dangerous elements were disabled:</p>
      ${getSectionItemsHtml(reportData.cleaningSummary.macros.items)}
      <div class="action-taken danger">
        ⚠️ ${reportData.cleaningSummary.macros.count} macro(s) removed for security
      </div>
    </section>
  `
      : "";

  const section7 =
    reportData.textCorrections.length > 0
      ? `
    <section class="detail-section">
      <h2>✏️ Text Corrections Applied</h2>
      <p class="section-intro">${reportData.textCorrections.length} spelling/grammar correction(s) were applied:</p>

      <table class="corrections-table">
        <thead>
          <tr>
            <th>#</th>
            <th>Type</th>
            <th>Original</th>
            <th>Corrected</th>
          </tr>
        </thead>
        <tbody>
          ${reportData.textCorrections
            .slice(0, 25)
            .map((c, i) => {
              const badge = getTypeBadge(c.type);
              return `
              <tr>
                <td class="center">${i + 1}</td>
                <td><span class="type-badge ${badge.class}">${badge.icon} ${badge.label}</span></td>
                <td class="original">${escapeHtml(c.original)}</td>
                <td class="corrected">${escapeHtml(c.corrected)}</td>
              </tr>
            `;
            })
            .join("")}
          ${
            reportData.textCorrections.length > 25
              ? `
            <tr>
              <td colspan="4" class="center more-items">... and ${
                reportData.textCorrections.length - 25
              } more corrections</td>
            </tr>
          `
              : ""
          }
        </tbody>
      </table>
    </section>
  `
      : "";

  const section8 =
    (ext === "xlsx" || ext === "xls") && reportData.executiveOverview.excelHiddenData.found > 0
      ? `
    <section class="detail-section">
      <h2>📊 Excel Hidden Data Removed</h2>
      <p class="section-intro">The following Excel-specific hidden elements were cleaned:</p>
      ${getSectionItemsHtml(reportData.cleaningSummary.excelHiddenData.items)}
      <div class="action-taken">
        📌 ${reportData.cleaningSummary.excelHiddenData.count} hidden Excel element(s) removed
      </div>
    </section>
  `
      : "";

  const recommendationsSection =
    reportData.recommendations.length > 0
      ? `
    <section class="detail-section recommendations-section">
      <h2>💡 Recommendations</h2>
      <ul class="recommendations-list">
        ${reportData.recommendations.map((rec) => `<li>${escapeHtml(rec)}</li>`).join("")}
      </ul>
    </section>
  `
      : "";

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Qualion-Doc Report — ${escapeHtml(filename)}</title>
  <style>
    :root {
      --primary: #3b82f6;
      --primary-dark: #2563eb;
      --success: #22c55e;
      --warning: #f59e0b;
      --danger: #ef4444;
      --gray-50: #f9fafb;
      --gray-100: #f3f4f6;
      --gray-200: #e5e7eb;
      --gray-300: #d1d5db;
      --gray-500: #6b7280;
      --gray-700: #374151;
      --gray-900: #111827;
    }

    * { box-sizing: border-box; margin: 0; padding: 0; }

    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
      background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 50%, #f0fdf4 100%);
      min-height: 100vh;
      padding: 32px 16px;
      color: var(--gray-900);
      line-height: 1.6;
    }

    .container { max-width: 950px; margin: 0 auto; }

    .header-card {
      background: linear-gradient(135deg, var(--primary) 0%, #8b5cf6 100%);
      border-radius: 20px;
      padding: 32px;
      color: white;
      text-align: center;
      margin-bottom: 16px;
      box-shadow: 0 10px 40px rgba(59, 130, 246, 0.3);
    }

    .header-card .logo { font-size: 36px; margin-bottom: 8px; }
    .header-card h1 { font-size: 28px; font-weight: 700; margin-bottom: 4px; }
    .header-card .subtitle { opacity: 0.9; font-size: 14px; margin-bottom: 20px; }

    .doc-info {
      display: flex;
      flex-wrap: wrap;
      justify-content: center;
      gap: 16px;
      margin-top: 16px;
    }

    .doc-info-item {
      background: rgba(255,255,255,0.15);
      padding: 8px 16px;
      border-radius: 8px;
      font-size: 13px;
    }

    .doc-info-item strong { margin-right: 4px; }

    /* ============================================================
       MODIF 2 — New header compliance status block
       ============================================================ */
    .compliance-status-block {
      margin: 18px auto 0;
      max-width: 680px;
      text-align: left;
      background: ${compliance.bg};
      color: ${compliance.color};
      border: 2px solid ${compliance.color}40;
      border-radius: 16px;
      padding: 14px 16px;
      box-shadow: 0 8px 22px rgba(0,0,0,0.12);
    }

    .compliance-status-top {
      display: flex;
      align-items: center;
      gap: 10px;
      font-weight: 800;
      font-size: 16px;
      line-height: 1.2;
    }

    .compliance-status-top .emoji { font-size: 18px; }
    .compliance-status-top .label { letter-spacing: 0.2px; }

    .compliance-status-subtitle {
      margin-top: 6px;
      font-size: 13px;
      line-height: 1.4;
      color: ${compliance.color};
      opacity: 0.95;
    }

    /* ============================================================
       MODIF 3 — Transformation Summary (inserted after header)
       ============================================================ */
    .transformation-summary {
      background: white;
      border-radius: 18px;
      padding: 22px 24px;
      margin: 0 0 24px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.08);
      overflow: hidden;
    }

    .transformation-summary h2 {
      font-size: 16px;
      margin-bottom: 14px;
      display: flex;
      align-items: center;
      gap: 10px;
      color: var(--gray-900);
    }

    .transformation-grid {
      display: grid;
      grid-template-columns: 1fr auto 1fr 1fr;
      gap: 14px;
      align-items: center;
    }

    .transformation-summary .score-box {
      border-radius: 14px;
      padding: 16px;
      text-align: center;
      border: 1px solid var(--gray-200);
      background: var(--gray-50);
    }

    .transformation-summary .score-box.initial {
      background: ${beforeStyle.bg};
      border: 2px solid ${beforeStyle.color}30;
    }

    .transformation-summary .score-box.final {
      background: ${afterStyle.bg};
      border: 2px solid ${afterStyle.color}30;
    }

    .transformation-summary .score-box .score {
      display: block;
      font-size: 22px;
      font-weight: 900;
      line-height: 1.1;
    }

    .transformation-summary .score-box.initial .score { color: ${beforeStyle.color}; }
    .transformation-summary .score-box.final .score { color: ${afterStyle.color}; }

    .transformation-summary .score-box .label {
      display: block;
      margin-top: 6px;
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.8px;
      color: var(--gray-500);
      font-weight: 700;
    }

    .transformation-summary .arrow {
      font-size: 22px;
      font-weight: 900;
      color: var(--gray-700);
      opacity: 0.75;
      text-align: center;
      padding: 0 6px;
    }

    .transformation-summary .improvement {
      border-radius: 14px;
      padding: 16px;
      text-align: center;
      border: 1px solid #bbf7d0;
      background: linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%);
    }

    .transformation-summary .improvement .score {
      display: block;
      font-size: 22px;
      font-weight: 900;
      color: #166534;
      line-height: 1.1;
    }

    .transformation-summary .improvement .label {
      display: block;
      margin-top: 6px;
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.8px;
      color: #166534;
      opacity: 0.9;
      font-weight: 800;
    }

    .score-card {
      background: white;
      border-radius: 20px;
      padding: 32px;
      margin-bottom: 24px;
      box-shadow: 0 4px 20px rgba(0,0,0,0.08);
    }

    .score-card h3 {
      text-align: center;
      font-size: 18px;
      color: var(--gray-700);
      margin-bottom: 24px;
    }

    .score-comparison {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 24px;
      flex-wrap: wrap;
    }

    .score-box {
      text-align: center;
      padding: 20px;
      border-radius: 16px;
      min-width: 160px;
    }

    .score-box.before { background: ${beforeStyle.bg}; border: 2px solid ${beforeStyle.color}40; }
    .score-box.after { background: ${afterStyle.bg}; border: 2px solid ${afterStyle.color}40; }

    .score-box-label {
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 1px;
      color: var(--gray-500);
      margin-bottom: 8px;
    }

    .score-circle {
      width: 100px;
      height: 100px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0 auto 12px;
      position: relative;
    }

    .score-box.before .score-circle {
      background: conic-gradient(${beforeStyle.color} ${reportData.beforeScore}%, var(--gray-200) ${reportData.beforeScore}%);
    }

    .score-box.after .score-circle {
      background: conic-gradient(${afterStyle.color} ${reportData.afterScore}%, var(--gray-200) ${reportData.afterScore}%);
    }

    .score-circle::before {
      content: '';
      width: 76px;
      height: 76px;
      background: white;
      border-radius: 50%;
      position: absolute;
    }

    .score-number {
      position: relative;
      z-index: 1;
      font-size: 28px;
      font-weight: 800;
    }

    .score-box.before .score-number { color: ${beforeStyle.color}; }
    .score-box.after .score-number { color: ${afterStyle.color}; }

    .score-status {
      font-size: 12px;
      font-weight: 600;
      padding: 4px 12px;
      border-radius: 12px;
    }

    .score-box.before .score-status { background: ${beforeStyle.color}20; color: ${beforeStyle.color}; }
    .score-box.after .score-status { background: ${afterStyle.color}20; color: ${afterStyle.color}; }

    .score-improvement {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 4px;
    }

    .score-improvement-arrow { font-size: 24px; color: var(--success); }

    .score-improvement-value {
      background: linear-gradient(135deg, var(--success) 0%, #16a34a 100%);
      color: white;
      padding: 8px 16px;
      border-radius: 20px;
      font-weight: 700;
      font-size: 14px;
    }

    .score-stats {
      display: flex;
      justify-content: center;
      gap: 32px;
      margin-top: 24px;
      padding-top: 24px;
      border-top: 1px solid var(--gray-200);
      flex-wrap: wrap;
    }

    .score-stat { text-align: center; min-width: 80px; }
    .score-stat-number { font-size: 28px; font-weight: 700; color: var(--gray-900); }
    .score-stat-label { font-size: 11px; color: var(--gray-500); text-transform: uppercase; }

    .overview-card {
      background: white;
      border-radius: 16px;
      margin-bottom: 24px;
      box-shadow: 0 2px 12px rgba(0,0,0,0.06);
      overflow: hidden;
    }

    .overview-card h3 {
      padding: 20px 24px;
      border-bottom: 1px solid var(--gray-100);
      font-size: 16px;
      display: flex;
      align-items: center;
      gap: 10px;
    }

    .overview-table {
      width: 100%;
      border-collapse: collapse;
    }

    .overview-table th {
      background: var(--gray-50);
      padding: 12px 16px;
      text-align: left;
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: var(--gray-500);
      border-bottom: 2px solid var(--gray-200);
    }

    .overview-table th.center,
    .overview-table td.center { text-align: center; }

    .overview-table td {
      padding: 14px 16px;
      border-bottom: 1px solid var(--gray-100);
      font-size: 14px;
    }

    .overview-table tr:hover { background: var(--gray-50); }

    .risk-level {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 600;
    }

    .detail-section {
      background: white;
      border-radius: 16px;
      padding: 24px;
      margin-bottom: 20px;
      box-shadow: 0 2px 12px rgba(0,0,0,0.06);
    }

    .detail-section h2 {
      font-size: 16px;
      color: var(--gray-900);
      margin-bottom: 16px;
      padding-bottom: 12px;
      border-bottom: 1px solid var(--gray-100);
    }

    .detail-section.danger-section { border-left: 4px solid var(--danger); }

    .section-intro { color: var(--gray-600); margin-bottom: 12px; font-size: 14px; }

    .items-list {
      list-style: none;
      background: var(--gray-50);
      border-radius: 8px;
      padding: 12px 16px;
      margin: 16px 0;
    }

    .items-list li {
      padding: 12px 0;
      border-bottom: 1px solid var(--gray-200);
      font-size: 14px;
    }

    .items-list li:last-child { border-bottom: none; }

    .item-content { display: flex; flex-direction: column; gap: 4px; }

    .item-label {
      font-weight: 600;
      color: var(--gray-800);
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .item-value {
      color: var(--gray-700);
      font-family: 'SF Mono', Monaco, 'Cascadia Code', monospace;
      font-size: 14px;
      background: white;
      padding: 8px 12px;
      border-radius: 6px;
      border: 1px solid var(--gray-200);
      word-break: break-word;
    }

    .more-items { color: var(--gray-500); font-style: italic; text-align: center; padding: 8px; }

    .action-taken {
      background: linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%);
      border: 1px solid #86efac;
      padding: 12px 16px;
      border-radius: 8px;
      font-size: 13px;
      color: #166534;
      font-weight: 500;
      margin-top: 16px;
    }

    .action-taken.danger {
      background: linear-gradient(135deg, #fee2e2 0%, #fecaca 100%);
      border-color: #fca5a5;
      color: #991b1b;
    }

    .corrections-table { width: 100%; border-collapse: collapse; margin: 16px 0; }

    .corrections-table th {
      background: var(--gray-50);
      padding: 12px;
      text-align: left;
      font-size: 11px;
      text-transform: uppercase;
      color: var(--gray-500);
      border-bottom: 2px solid var(--gray-200);
    }

    .corrections-table td {
      padding: 12px;
      border-bottom: 1px solid var(--gray-100);
      font-size: 13px;
      vertical-align: middle;
    }

    .corrections-table .original {
      color: #991b1b;
      text-decoration: line-through;
      background: #fef2f2;
      padding: 4px 8px;
      border-radius: 4px;
    }

    .corrections-table .corrected {
      color: #166534;
      font-weight: 600;
      background: #f0fdf4;
      padding: 4px 8px;
      border-radius: 4px;
    }

    .type-badge {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 4px 8px;
      border-radius: 10px;
      font-size: 10px;
      font-weight: 600;
    }

    .type-fragment { background: #fef3c7; color: #92400e; }
    .type-spelling { background: #dbeafe; color: #1e40af; }
    .type-grammar { background: #f3e8ff; color: #7c3aed; }
    .type-punctuation { background: #f3f4f6; color: #374151; }
    .type-ai { background: #dcfce7; color: #166534; }

    .recommendations-list { list-style: none; }

    .recommendations-list li {
      padding: 12px 16px;
      background: var(--gray-50);
      border-radius: 8px;
      margin-bottom: 8px;
      font-size: 14px;
      color: var(--gray-700);
      border-left: 3px solid var(--primary);
    }

    .footer {
      text-align: center;
      padding: 32px 16px;
      color: var(--gray-500);
      font-size: 13px;
    }

    .footer a { color: var(--primary); text-decoration: none; font-weight: 500; }

    @media (max-width: 720px) {
      .transformation-grid {
        grid-template-columns: 1fr;
        gap: 10px;
      }
      .transformation-summary .arrow { display: none; }
      .compliance-status-block { text-align: center; }
      .compliance-status-top { justify-content: center; }
      .compliance-status-subtitle { text-align: center; }
    }

    @media (max-width: 640px) {
      body { padding: 16px 12px; }
      .score-comparison { flex-direction: column; }
      .score-improvement { flex-direction: row; }
      .score-stats { flex-direction: column; gap: 16px; }
      .overview-table { font-size: 12px; }
      .overview-table th, .overview-table td { padding: 10px 8px; }
    }

    @media print {
      body { background: white; padding: 0; }
      .detail-section, .score-card, .overview-card, .transformation-summary { box-shadow: none; border: 1px solid var(--gray-200); break-inside: avoid; }
      .header-card { box-shadow: none; }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header-card">
      <div class="logo">🔐</div>
      <h1>Document Clean & Secure Report</h1>
      <p class="subtitle">Generated by Qualion-Doc — AI-powered Compliance Engine</p>

      <div class="doc-info">
        <div class="doc-info-item">📁 <strong>Document:</strong> ${escapeHtml(filename)}</div>
        <div class="doc-info-item">📅 <strong>Processed:</strong> ${new Date().toLocaleDateString("en-US", {
          year: "numeric",
          month: "short",
          day: "numeric",
        })}</div>
        ${user ? `<div class="doc-info-item">👤 <strong>User:</strong> ${escapeHtml(user)}</div>` : ""}
        ${organization ? `<div class="doc-info-item">🏢 <strong>Organization:</strong> ${escapeHtml(organization)}</div>` : ""}
      </div>

      <!-- MODIF 2: Replace old badge with a more visible status block -->
      <div class="compliance-status-block">
        <div class="compliance-status-top">
          <span class="emoji">${compliance.emoji}</span>
          <span class="label">${escapeHtml(compliance.label)}</span>
        </div>
        <div class="compliance-status-subtitle">${escapeHtml(compliance.subtitle || "")}</div>
      </div>
    </div>

    <!-- MODIF 3: Transformation Summary -->
    <div class="transformation-summary">
      <h2>📈 Transformation Summary</h2>
      <div class="transformation-grid">
        <div class="score-box initial">
          <span class="score">${reportData.beforeScore}/10</span>
          <span class="label">Score Initial</span>
        </div>
        <div class="arrow">→</div>
        <div class="score-box final">
          <span class="score">${reportData.afterScore}/10</span>
          <span class="label">Score Final</span>
        </div>
        <div class="improvement">
          <span class="score">+${reportData.riskReduction}%</span>
          <span class="label">Amélioration</span>
        </div>
      </div>
    </div>

    <div class="score-card">
      <h3>🛡️ Security Score Comparison</h3>
      <div class="score-comparison">
        <div class="score-box before">
          <div class="score-box-label">Security Level Before</div>
          <div class="score-circle">
            <span class="score-number">${reportData.beforeScore}</span>
          </div>
          <span class="score-status">${beforeStyle.label}</span>
        </div>

        ${
          reportData.afterScore > reportData.beforeScore
            ? `
        <div class="score-improvement">
          <div class="score-improvement-arrow">→</div>
          <div class="score-improvement-value">+${reportData.afterScore - reportData.beforeScore} pts</div>
          <div style="font-size: 11px; color: var(--gray-500);">🔥 ${reportData.riskReduction}% Risk Reduction</div>
        </div>
        `
            : ""
        }

        <div class="score-box after">
          <div class="score-box-label">After Cleaning</div>
          <div class="score-circle">
            <span class="score-number">${reportData.afterScore}</span>
          </div>
          <span class="score-status">${afterStyle.label}</span>
        </div>
      </div>

      <div class="score-stats">
        <div class="score-stat">
          <div class="score-stat-number">${reportData.summary.totalIssuesFound}</div>
          <div class="score-stat-label">Issues Found</div>
        </div>
        <div class="score-stat">
          <div class="score-stat-number">${reportData.summary.elementsRemoved}</div>
          <div class="score-stat-label">Elements Removed</div>
        </div>
        <div class="score-stat">
          <div class="score-stat-number">${reportData.summary.correctionsApplied}</div>
          <div class="score-stat-label">Corrections</div>
        </div>
        <div class="score-stat">
          <div class="score-stat-number">${reportData.summary.criticalRisksResolved}</div>
          <div class="score-stat-label">Critical Resolved</div>
        </div>
      </div>
    </div>

    <div class="overview-card">
      <h3>📊 Executive Overview</h3>
      <table class="overview-table">
        <thead>
          <tr>
            <th>Category</th>
            <th class="center">Issues Found</th>
            <th class="center">Cleaned</th>
            <th class="center">Remaining</th>
            <th class="center">Risk Level</th>
          </tr>
        </thead>
        <tbody>
          ${executiveOverviewRows}
        </tbody>
      </table>
    </div>

    ${businessRisksSection}

    ${section1}
    ${section2}
    ${section3}
    ${section4}
    ${section5}
    ${section6}
    ${section7}
    ${section8}
    ${recommendationsSection}

    <div class="footer">
      <p>Generated by <a href="https://mindorion.com" target="_blank">Qualion-Doc</a> by Mindorion</p>
      <p style="margin-top: 4px;">${new Date().toLocaleDateString("en-US", {
        weekday: "long",
        year: "numeric",
        month: "long",
        day: "numeric",
        hour: "2-digit",
        minute: "2-digit",
      })}</p>
    </div>
  </div>
</body>
</html>`;
}
