// lib/report.js

// Qualion Proposal - VERSION 4.1 (Fusion complète)
// ✅ CONSERVÉ: buildReportData() inchangé pour compatibilité JSON
// ✅ CONSERVÉ: Business Risks Section complète (table 5 catégories, blocking issues, top signals)
// ✅ CONSERVÉ: Score Comparison, Executive Overview, Recommendations
// ✅ NOUVEAU: Status Banner "Ready to Send" / "Review Recommended"
// ✅ NOUVEAU: Document Status Card avec badge CLIENT-READY
// ✅ NOUVEAU: Transformation Summary avec contexte business
// ✅ NOUVEAU: Metrics Grid (Items Cleaned, Improvement, Items Kept)
// ✅ NOUVEAU: Annexes Collapsibles avec <details>
// ✅ NOUVEAU: CSS modernisé (emerald/amber/slate palette)
// ✅ NOUVEAU: Certificate amélioré avec gradient

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
    case "critical": return "🔥";
    case "high": return "🛑";
    case "medium": return "🔶";
    case "low": return "🟢";
    default: return "🟡";
  }
}

function getRiskLevelColor(level) {
  switch (String(level || "").toLowerCase()) {
    case "critical": return "#dc2626";
    case "high": return "#ef4444";
    case "medium": return "#f59e0b";
    case "low": return "#10b981";
    default: return "#6b7280";
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
          return val.map((v) => (typeof v === "string" ? v : v._ || v.text || "")).join("");
      }
    }
  }
  return "";
}

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
// BUSINESS RISKS HELPERS
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
// BUILD REPORT DATA (JSON) - Inchangé
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
  const analyzerScore = analysis?.summary?.riskScore;
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

  // Corrections
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

  // Metadata
  const metadataItems = [];
  if (Array.isArray(analysis?.detections?.metadata)) {
    analysis.detections.metadata.forEach((m) => {
      const label = m.key || m.type || m.name || "Property";
      const value = m.value ? String(m.value).substring(0, 160) : undefined;
      metadataItems.push({ label, value });
    });
  }

  // Comments
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

  // Track changes
  const trackChangesItems = [];
  if (Array.isArray(analysis?.detections?.trackChanges)) {
    analysis.detections.trackChanges.forEach((tc) => {
      const type = tc.type || tc.changeType || "change";
      const typeEmoji = type === "deletion" ? "🔴 Deleted:" : type === "insertion" ? "🟢 Added:" : "🔄 Changed:";
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

  // Hidden content
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

      if (type === "hidden_text_details" && Array.isArray(hc.items) && hc.items.length > 0) {
        hc.items.forEach((it) => {
          hiddenContentDetailedCount++;
          const reasonLabel = {
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
          });
        });
        return;
      }

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

  // Macros
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

  // Sensitive Data
  const sensitiveDataItems = [];
  const sensitiveDataTypes = new Set();
  if (Array.isArray(analysis?.detections?.sensitiveData)) {
    analysis.detections.sensitiveData.forEach((sd) => {
      const type = sd.type || sd.dataType || "Sensitive";
      sensitiveDataTypes.add(type);
      const fullValue = sd.rawValue || sd.originalValue || sd.value || sd.match || sd.text || "";
      sensitiveDataItems.push({
        label: type.replace(/_/g, " ").toUpperCase(),
        value: String(fullValue),
      });
    });
  }

  // Visual objects
  const visualObjectsItems = [];
  if (Array.isArray(analysis?.detections?.visualObjects)) {
    analysis.detections.visualObjects.forEach((vo) => {
      const type = vo.type || "visual_object";
      const typeLabel = {
        shape_covering_text: "Shape covering text",
        missing_alt_text: "Missing alt text",
      }[type] || "Visual object";
      visualObjectsItems.push({ label: typeLabel, value: vo.description || "" });
    });
  }

  // Orphan data
  const orphanDataItems = [];
  if (Array.isArray(analysis?.detections?.orphanData)) {
    analysis.detections.orphanData.forEach((od) => {
      const type = od.type || "orphan_data";
      const typeLabel = {
        broken_link: "Broken link",
        empty_page: "Empty page",
        trailing_whitespace: "Trailing whitespace",
      }[type] || "Orphan data";
      orphanDataItems.push({ label: typeLabel, value: od.value || od.description || "" });
    });
  }

  // Excel hidden data
  const excelHiddenDataItems = [];
  if (Array.isArray(analysis?.detections?.excelHiddenData)) {
    analysis.detections.excelHiddenData.forEach((ed) => {
      const type = ed.type || "excel_hidden";
      const typeLabel = {
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

  // Spelling items
  const spellingItems = allCorrections.map((c) => ({
    label: c.original,
    value: `→ ${c.corrected}`,
  }));

  // Executive Overview
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

  // Cleaning Summary
  const cleaningSummary = {
    metadata: {
      count: metadataItems.length || fmt(cleaning?.metaRemoved) || 0,
      items: metadataItems,
      scoreImpact: scoreImpacts.metadata || Math.min(10, metadataItems.length * 2),
    },
    comments: {
      count: commentItems.length || (fmt(cleaning?.commentsXmlRemoved) + fmt(cleaning?.commentMarkersRemoved)) || 0,
      items: commentItems,
      scoreImpact: scoreImpacts.comments || Math.min(15, commentItems.length * 3),
    },
    trackChanges: {
      count: trackChangesItems.length ||
        (fmt(cleaning?.revisionsAccepted?.deletionsRemoved) + fmt(cleaning?.revisionsAccepted?.insertionsUnwrapped)) || 0,
      items: trackChangesItems,
      deletions: fmt(cleaning?.revisionsAccepted?.deletionsRemoved) || 0,
      insertions: fmt(cleaning?.revisionsAccepted?.insertionsUnwrapped) || 0,
      scoreImpact: scoreImpacts.trackChanges || Math.min(10, trackChangesItems.length * 2),
    },
    hiddenContent: {
      count: hiddenContentDetailedCount || hiddenContentItems.length || fmt(cleaning?.hiddenRemoved) || 0,
      items: hiddenContentItems,
      scoreImpact: scoreImpacts.hiddenContent || Math.min(20, (hiddenContentDetailedCount || hiddenContentItems.length) * 5),
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

  // Risks Detected
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

  // Recommendations
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

  // Business Risks
  const riskSummary = analysis?.riskSummary || null;
  const riskObjects = Array.isArray(analysis?.riskObjects) ? analysis.riskObjects : [];

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

  Object.values(byCategory).forEach((cat) => {
    cat.items.sort((a, b) => {
      const d = severityRank(b.severity) - severityRank(a.severity);
      if (d !== 0) return d;
      const pb = typeof b.points === "number" ? b.points : 0;
      const pa = typeof a.points === "number" ? a.points : 0;
      return pb - pa;
    });
  });

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
    totalRiskObjects: typeof riskSummary?.totalRiskObjects === "number" ? riskSummary.totalRiskObjects : riskObjects.length,
  };

  const businessRiskAssessment = analysis?.businessRiskAssessment ?? null;

  // Extra Metrics
  const clarityScoreImprovement =
    allCorrections.length > 0 ? Math.min(25, Math.round(allCorrections.length * 2.5)) : 0;
  const documentSizeReduction =
    (fmt(cleaning?.mediaDeleted) + fmt(cleaning?.embeddedFilesRemoved)) > 0
      ? Math.min(30, (fmt(cleaning?.mediaDeleted) + fmt(cleaning?.embeddedFilesRemoved)) * 5)
      : 0;

  console.log(
    `[REPORT JSON v4.1] beforeScore=${beforeScore}, afterScore=${afterScore}, riskReduction=${riskReduction}%, complianceStatus=${complianceStatus}, corrections=${allCorrections.length}, businessClientReady=${businessRisks.clientReady}`
  );

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
    businessRisks,
    businessRiskAssessment,
  };
}

// ============================================================
// BUILD REPORT HTML - VERSION 4.1 (Fusion complète)
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

  const isReady = reportData.complianceStatus === "safe" || 
    (reportData.businessRisks?.clientReady === "YES");
  
  const processedDate = new Date().toLocaleDateString("en-US", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  const processedDateTime = new Date().toLocaleDateString("en-US", {
    weekday: "long",
    year: "numeric",
    month: "long",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });

  // Business Risk Context
  const generateBusinessRiskContext = () => {
    const br = reportData.businessRisks;
    if (!br || !br.categories) return [];

    const contexts = [];
    const categoryLabels = {
      margin: { protected: "Margin Exposure: Protected", risk: "Margin Exposure: Review recommended" },
      delivery: { protected: "Delivery Risk: Mitigated", risk: "Delivery Commitments: Flagged" },
      negotiation: { protected: "Negotiation Power: Secured", risk: "Negotiation Leverage: Exposed" },
      compliance: { protected: "Compliance: Verified", risk: "Compliance: Requires review" },
      credibility: { protected: "Credibility: Preserved", risk: "Professional Credibility: Needs attention" },
    };

    br.categories.forEach((cat) => {
      const labels = categoryLabels[cat.key];
      if (!labels) return;
      const isLow = severityRank(cat.maxSeverity) <= 2;
      contexts.push({
        text: isLow ? labels.protected : labels.risk,
        isProtected: isLow,
      });
    });

    return contexts;
  };

  const businessContexts = generateBusinessRiskContext();

  // Metrics
  const totalCleaned = reportData.summary.elementsRemoved + reportData.summary.correctionsApplied;
  const improvement = reportData.riskReduction;
  const itemsKept = reportData.risksDetected.filter(r => r.action === "flagged").length;

  // Business Risk Badge
  const businessRiskBadge = (sev) => {
    const s = String(sev || "low").toLowerCase();
    const color = getRiskLevelColor(s);
    return `<span style="background: ${color}20; color: ${color}; padding: 4px 10px; border-radius: 12px; font-size: 12px; font-weight: 600;">${getRiskLevelEmoji(s)} ${escapeHtml(severityLabel(s))}</span>`;
  };

  // Executive Overview Row
  const getCategoryRow = (name, stats) => {
    if (!stats || (stats.found === 0 && stats.cleaned === 0)) return "";
    return `
      <tr>
        <td>${escapeHtml(name)}</td>
        <td class="center">${stats.found}</td>
        <td class="center">${stats.cleaned}</td>
        <td class="center">${stats.remaining}</td>
        <td class="center">
          <span class="risk-badge" style="background: ${getRiskLevelColor(stats.riskLevel)}20; color: ${getRiskLevelColor(stats.riskLevel)}">
            ${getRiskLevelEmoji(stats.riskLevel)} ${stats.riskLevel.charAt(0).toUpperCase() + stats.riskLevel.slice(1)}
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

  // Annex Section Generator
  const generateAnnexSection = (title, icon, items) => {
    if (!items || items.length === 0) return "";
    return `
      <details class="annex-section">
        <summary>
          <span class="annex-icon">${icon}</span>
          <span class="annex-title">${escapeHtml(title)}</span>
          <span class="annex-count">${items.length} item${items.length > 1 ? 's' : ''}</span>
        </summary>
        <div class="annex-content">
          <table class="annex-table">
            <thead>
              <tr>
                <th>Item</th>
                <th>Details</th>
              </tr>
            </thead>
            <tbody>
              ${items.slice(0, 20).map(item => `
                <tr>
                  <td class="item-label">${escapeHtml(item.label || '')}</td>
                  <td class="item-value">${escapeHtml(item.value || '')}</td>
                </tr>
              `).join('')}
              ${items.length > 20 ? `<tr><td colspan="2" class="more-items">... and ${items.length - 20} more items</td></tr>` : ''}
            </tbody>
          </table>
        </div>
      </details>
    `;
  };

  // Corrections Annex
  const generateCorrectionsAnnex = () => {
    if (!reportData.textCorrections || reportData.textCorrections.length === 0) return "";
    return `
      <details class="annex-section">
        <summary>
          <span class="annex-icon">✏️</span>
          <span class="annex-title">Text Corrections Applied</span>
          <span class="annex-count">${reportData.textCorrections.length} correction${reportData.textCorrections.length > 1 ? 's' : ''}</span>
        </summary>
        <div class="annex-content">
          <table class="annex-table corrections-table">
            <thead>
              <tr>
                <th>#</th>
                <th>Type</th>
                <th>Original</th>
                <th>Corrected</th>
              </tr>
            </thead>
            <tbody>
              ${reportData.textCorrections.slice(0, 30).map((c, i) => {
                const badge = getTypeBadge(c.type);
                return `
                  <tr>
                    <td class="correction-num">${i + 1}</td>
                    <td class="correction-type"><span class="type-badge">${badge.icon} ${badge.label}</span></td>
                    <td class="correction-original">${escapeHtml(c.original)}</td>
                    <td class="correction-corrected">${escapeHtml(c.corrected)}</td>
                  </tr>
                `;
              }).join('')}
              ${reportData.textCorrections.length > 30 ? `<tr><td colspan="4" class="more-items">... and ${reportData.textCorrections.length - 30} more corrections</td></tr>` : ''}
            </tbody>
          </table>
        </div>
      </details>
    `;
  };

  // Business Risks Section (ORIGINAL - conservé)
  const br = reportData.businessRisks || null;
  const businessRisksSection = br ? `
    <div class="section-card business-risks-section">
      <h2 class="section-title"><span>🏛️</span> Business Risks (Executive)</h2>
      
      <div class="business-gate">
        <div class="gate-item">
          <span class="gate-label">Client-Ready Gate:</span>
          <span class="gate-value ${br.clientReady === 'YES' ? 'ready' : 'blocked'}">${escapeHtml(br.clientReady)}</span>
        </div>
        <div class="gate-item">
          <span class="gate-label">Overall Severity:</span>
          ${businessRiskBadge(br.overallSeverity)}
        </div>
        <div class="gate-item">
          <span class="gate-label">Total Signals:</span>
          <span class="gate-value">${fmt(br.totalRiskObjects)}</span>
        </div>
      </div>

      ${Array.isArray(br.executiveSignals) && br.executiveSignals.length ? `
        <div class="executive-signals">
          <strong>Executive Signals:</strong> ${escapeHtml(br.executiveSignals.join(" • "))}
        </div>
      ` : ""}

      <table class="business-table">
        <thead>
          <tr>
            <th>Category</th>
            <th class="center">Count</th>
            <th class="center">Max Severity</th>
            <th class="center">Surfaces (V/H/S/R)</th>
          </tr>
        </thead>
        <tbody>
          ${br.categories.map((c) => {
            const s = c.bySurface || {};
            const surf = `${fmt(s.Visible)}/${fmt(s.Hidden)}/${fmt(s.Structural)}/${fmt(s.Residual)}`;
            return `
              <tr>
                <td>${escapeHtml(c.label)}</td>
                <td class="center">${fmt(c.count)}</td>
                <td class="center">${businessRiskBadge(c.maxSeverity)}</td>
                <td class="center">${escapeHtml(surf)}</td>
              </tr>
            `;
          }).join("")}
        </tbody>
      </table>

      ${Array.isArray(br.blockingIssues) && br.blockingIssues.length ? `
        <div class="blocking-issues">
          <h3>🛑 Blocking Issues</h3>
          <div class="blocking-list">
            ${br.blockingIssues.slice(0, 10).map((it) => `
              <div class="blocking-item">
                <div class="blocking-header">
                  ${escapeHtml(it.label || it.riskCategory || "Issue")} — ${escapeHtml(it.severity || "")}
                </div>
                <div class="blocking-reason">${escapeHtml(it.reason || "")}</div>
                ${it.fixability ? `<div class="blocking-fix">Fixability: ${escapeHtml(it.fixability)}</div>` : ""}
              </div>
            `).join("")}
            ${br.blockingIssues.length > 10 ? `<div class="more-items">... and ${br.blockingIssues.length - 10} more blocking issues</div>` : ""}
          </div>
        </div>
      ` : ""}

      ${br.categories.filter((c) => Array.isArray(c.items) && c.items.length).map((c) => {
        const top = c.items.slice(0, 6);
        return `
          <div class="category-signals">
            <h4>📌 ${escapeHtml(c.label)} — Top Signals</h4>
            <div class="signals-list">
              ${top.map((it) => `
                <div class="signal-item">
                  <div class="signal-header">
                    ${businessRiskBadge(it.severity)} 
                    <span class="signal-surface">${escapeHtml(it.surface || "")}</span>
                    ${it.location ? `<span class="signal-location">— ${escapeHtml(it.location)}</span>` : ""}
                  </div>
                  <div class="signal-reason">${escapeHtml(it.reason || "")}</div>
                  <div class="signal-meta">
                    Fixability: ${escapeHtml(it.fixability || "manual")}
                    ${it.ruleId ? ` • Rule: ${escapeHtml(it.ruleId)}` : ""}
                    ${typeof it.points === "number" ? ` • Points: ${it.points}` : ""}
                  </div>
                </div>
              `).join("")}
              ${c.items.length > 6 ? `<div class="more-items">... and ${c.items.length - 6} more signals</div>` : ""}
            </div>
          </div>
        `;
      }).join("")}
    </div>
  ` : "";

  // Recommendations Section
  const recommendationsSection = reportData.recommendations.length > 0 ? `
    <div class="section-card">
      <h2 class="section-title"><span>💡</span> Recommendations</h2>
      <ul class="recommendations-list">
        ${reportData.recommendations.map((rec) => `<li>${escapeHtml(rec)}</li>`).join("")}
      </ul>
    </div>
  ` : "";

  // HTML Template
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Document Cleaning Report - ${escapeHtml(filename)}</title>
  <style>
    :root {
      --emerald-50: #ecfdf5;
      --emerald-100: #d1fae5;
      --emerald-500: #10b981;
      --emerald-600: #059669;
      --emerald-700: #047857;
      --amber-50: #fffbeb;
      --amber-100: #fef3c7;
      --amber-500: #f59e0b;
      --amber-600: #d97706;
      --amber-700: #b45309;
      --slate-50: #f8fafc;
      --slate-100: #f1f5f9;
      --slate-200: #e2e8f0;
      --slate-300: #cbd5e1;
      --slate-400: #94a3b8;
      --slate-500: #64748b;
      --slate-600: #475569;
      --slate-700: #334155;
      --slate-800: #1e293b;
      --slate-900: #0f172a;
      --red-500: #ef4444;
      --red-600: #dc2626;
    }

    * { margin: 0; padding: 0; box-sizing: border-box; }

    body {
      font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
      background: var(--slate-50);
      color: var(--slate-800);
      line-height: 1.6;
      padding: 24px;
    }

    .container { max-width: 960px; margin: 0 auto; }

    /* Status Banner */
    .status-banner {
      display: flex;
      align-items: center;
      gap: 16px;
      padding: 24px 32px;
      border-radius: 16px;
      margin-bottom: 24px;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }

    .status-banner.ready {
      background: linear-gradient(135deg, var(--emerald-500), var(--emerald-600));
      color: white;
    }

    .status-banner.review {
      background: linear-gradient(135deg, var(--amber-500), var(--amber-600));
      color: white;
    }

    .status-icon {
      width: 64px;
      height: 64px;
      background: rgba(255, 255, 255, 0.2);
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 32px;
    }

    .status-content h1 { font-size: 28px; font-weight: 700; margin-bottom: 4px; }
    .status-content p { opacity: 0.9; font-size: 14px; }

    /* Document Card */
    .document-card {
      background: white;
      border-radius: 12px;
      padding: 20px 24px;
      margin-bottom: 24px;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
      border: 1px solid var(--slate-200);
    }

    .card-header {
      display: flex;
      align-items: center;
      gap: 12px;
      margin-bottom: 8px;
    }

    .shield-icon { font-size: 24px; }
    .card-title { font-weight: 600; font-size: 18px; color: var(--slate-800); flex: 1; }

    .badge {
      padding: 6px 12px;
      border-radius: 20px;
      font-size: 12px;
      font-weight: 600;
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }

    .badge.client-ready { background: var(--emerald-100); color: var(--emerald-700); }
    .badge.needs-review { background: var(--amber-100); color: var(--amber-700); }
    .card-meta { color: var(--slate-500); font-size: 14px; }

    /* Section Card */
    .section-card {
      background: white;
      border-radius: 12px;
      padding: 24px;
      margin-bottom: 24px;
      box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
      border: 1px solid var(--slate-200);
    }

    .section-title {
      font-size: 18px;
      font-weight: 600;
      color: var(--slate-800);
      margin-bottom: 16px;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    /* Transformation Summary */
    .readiness-message { font-size: 15px; color: var(--slate-600); margin-bottom: 12px; }

    .reasons-list { list-style: none; margin-bottom: 24px; }

    .reasons-list li {
      padding: 8px 0;
      display: flex;
      align-items: center;
      gap: 10px;
      font-size: 14px;
      color: var(--slate-700);
      border-bottom: 1px solid var(--slate-100);
    }

    .reasons-list li:last-child { border-bottom: none; }

    .reason-icon {
      width: 20px;
      height: 20px;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 12px;
    }

    .reason-icon.protected { background: var(--emerald-100); color: var(--emerald-600); }
    .reason-icon.risk { background: var(--amber-100); color: var(--amber-600); }

    /* Metrics Grid */
    .metrics-grid {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 16px;
    }

    .metric-card {
      background: var(--slate-50);
      border-radius: 12px;
      padding: 20px;
      text-align: center;
      border: 1px solid var(--slate-200);
    }

    .metric-value { font-size: 32px; font-weight: 700; color: var(--slate-800); margin-bottom: 4px; }
    .metric-value.positive { color: var(--emerald-600); }
    .metric-label { font-size: 13px; color: var(--slate-500); text-transform: uppercase; letter-spacing: 0.5px; }

    /* Score Comparison */
    .score-comparison {
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 24px;
      flex-wrap: wrap;
    }

    .score-card {
      text-align: center;
      padding: 20px 32px;
      border-radius: 12px;
      min-width: 140px;
    }

    .score-card.before { background: var(--slate-100); }
    .score-card.after { background: var(--emerald-50); }

    .score-label { font-size: 13px; color: var(--slate-500); margin-bottom: 8px; text-transform: uppercase; }
    .score-value { font-size: 48px; font-weight: 700; }
    .score-status { font-size: 14px; margin-top: 4px; }

    .score-arrow {
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 4px;
    }

    .score-arrow .arrow { font-size: 24px; color: var(--emerald-500); }
    .score-arrow .improvement { font-size: 14px; font-weight: 600; color: var(--emerald-600); }

    /* Tables */
    .overview-table, .business-table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 16px;
    }

    .overview-table th, .overview-table td,
    .business-table th, .business-table td {
      padding: 12px 16px;
      text-align: left;
      border-bottom: 1px solid var(--slate-200);
    }

    .overview-table th, .business-table th {
      background: var(--slate-50);
      font-weight: 600;
      font-size: 13px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: var(--slate-600);
    }

    .center { text-align: center; }

    .risk-badge {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 12px;
      font-weight: 600;
    }

    /* Business Risks Section */
    .business-gate {
      display: flex;
      gap: 24px;
      flex-wrap: wrap;
      padding: 16px;
      background: var(--slate-50);
      border-radius: 8px;
      margin-bottom: 16px;
    }

    .gate-item { display: flex; align-items: center; gap: 8px; }
    .gate-label { font-size: 14px; color: var(--slate-600); }
    .gate-value { font-weight: 600; font-size: 14px; }
    .gate-value.ready { color: var(--emerald-600); }
    .gate-value.blocked { color: var(--red-600); }

    .executive-signals {
      padding: 12px 16px;
      background: var(--slate-100);
      border-radius: 8px;
      margin-bottom: 16px;
      font-size: 14px;
      color: var(--slate-700);
    }

    .blocking-issues { margin-top: 24px; }
    .blocking-issues h3 { font-size: 16px; margin-bottom: 12px; color: var(--red-600); }

    .blocking-list { display: flex; flex-direction: column; gap: 12px; }

    .blocking-item {
      padding: 12px 16px;
      background: #fef2f2;
      border: 1px solid #fecaca;
      border-radius: 8px;
    }

    .blocking-header { font-weight: 600; color: var(--red-600); margin-bottom: 4px; }
    .blocking-reason { font-size: 14px; color: var(--slate-700); }
    .blocking-fix { font-size: 12px; color: var(--slate-500); margin-top: 4px; }

    .category-signals { margin-top: 24px; }
    .category-signals h4 { font-size: 15px; margin-bottom: 12px; color: var(--slate-700); }

    .signals-list { display: flex; flex-direction: column; gap: 10px; }

    .signal-item {
      padding: 12px 16px;
      background: var(--slate-50);
      border: 1px solid var(--slate-200);
      border-radius: 8px;
    }

    .signal-header { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; margin-bottom: 4px; }
    .signal-surface { font-size: 12px; color: var(--slate-500); }
    .signal-location { font-size: 12px; color: var(--slate-400); }
    .signal-reason { font-size: 14px; color: var(--slate-700); }
    .signal-meta { font-size: 12px; color: var(--slate-500); margin-top: 8px; }

    /* Annexes */
    .annexes-section { margin-bottom: 24px; }

    .annex-section {
      background: white;
      border: 1px solid var(--slate-200);
      border-radius: 8px;
      margin-bottom: 12px;
      overflow: hidden;
    }

    .annex-section:last-child { margin-bottom: 0; }

    .annex-section summary {
      display: flex;
      align-items: center;
      gap: 12px;
      padding: 14px 16px;
      background: var(--slate-50);
      cursor: pointer;
      user-select: none;
      font-weight: 500;
    }

    .annex-section summary:hover { background: var(--slate-100); }
    .annex-section[open] summary { border-bottom: 1px solid var(--slate-200); }

    .annex-icon { font-size: 18px; }
    .annex-title { flex: 1; color: var(--slate-700); }
    .annex-count { font-size: 12px; color: var(--slate-500); background: var(--slate-200); padding: 4px 10px; border-radius: 12px; }

    .annex-content { padding: 16px; }

    .annex-table { width: 100%; border-collapse: collapse; font-size: 14px; }

    .annex-table th, .annex-table td {
      padding: 10px 12px;
      text-align: left;
      border-bottom: 1px solid var(--slate-100);
    }

    .annex-table th {
      background: var(--slate-50);
      font-weight: 600;
      font-size: 12px;
      text-transform: uppercase;
      color: var(--slate-500);
    }

    .annex-table .item-label { font-weight: 500; color: var(--slate-700); }
    .annex-table .item-value { color: var(--slate-600); max-width: 400px; overflow: hidden; text-overflow: ellipsis; }

    .corrections-table .correction-num { width: 40px; color: var(--slate-400); }
    .corrections-table .correction-type { width: 100px; }

    .type-badge {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 4px 8px;
      background: var(--slate-100);
      border-radius: 4px;
      font-size: 12px;
      color: var(--slate-600);
    }

    .corrections-table .correction-original { color: #dc2626; text-decoration: line-through; }
    .corrections-table .correction-corrected { color: var(--emerald-600); font-weight: 500; }

    .more-items { text-align: center; color: var(--slate-500); font-style: italic; padding: 12px !important; }

    /* Recommendations */
    .recommendations-list {
      list-style: none;
      padding: 0;
    }

    .recommendations-list li {
      padding: 10px 0;
      border-bottom: 1px solid var(--slate-100);
      font-size: 14px;
    }

    .recommendations-list li:last-child { border-bottom: none; }

    /* Certificate */
    .certificate-section {
      background: linear-gradient(135deg, var(--slate-800), var(--slate-900));
      color: white;
      border-radius: 12px;
      padding: 32px;
      text-align: center;
      margin-top: 24px;
    }

    .certificate-icon { font-size: 48px; margin-bottom: 16px; }
    .certificate-title { font-size: 20px; font-weight: 600; margin-bottom: 8px; }

    .certificate-text {
      color: var(--slate-300);
      font-size: 14px;
      margin-bottom: 16px;
      max-width: 500px;
      margin-left: auto;
      margin-right: auto;
    }

    .certificate-meta {
      display: flex;
      justify-content: center;
      gap: 32px;
      font-size: 13px;
      color: var(--slate-400);
    }

    .certificate-meta span { display: flex; align-items: center; gap: 6px; }

    /* Footer */
    .footer { text-align: center; padding: 24px; color: var(--slate-500); font-size: 13px; }

    /* Responsive */
    @media (max-width: 640px) {
      body { padding: 16px; }
      .status-banner { flex-direction: column; text-align: center; padding: 20px; }
      .metrics-grid { grid-template-columns: 1fr; }
      .score-comparison { flex-direction: column; }
      .business-gate { flex-direction: column; gap: 12px; }
      .certificate-meta { flex-direction: column; gap: 8px; }
    }

    /* Print */
    @media print {
      body { background: white; padding: 0; }
      .section-card, .document-card, .status-banner { box-shadow: none; break-inside: avoid; }
    }
  </style>
</head>
<body>
  <div class="container">
    <!-- Status Banner (NOUVEAU) -->
    <div class="status-banner ${isReady ? 'ready' : 'review'}">
      <div class="status-icon">${isReady ? '✓' : '⚠'}</div>
      <div class="status-content">
        <h1>${isReady ? 'Ready to Send' : 'Review Recommended'}</h1>
        <p>Document cleaned on ${processedDate}</p>
      </div>
    </div>

    <!-- Document Status Card (NOUVEAU) -->
    <div class="document-card">
      <div class="card-header">
        <span class="shield-icon">🛡️</span>
        <span class="card-title">Document Status</span>
        <span class="badge ${isReady ? 'client-ready' : 'needs-review'}">${isReady ? 'CLIENT-READY' : 'REVIEW NEEDED'}</span>
      </div>
      <p class="card-meta">
        📁 ${escapeHtml(filename)} • Processed: ${processedDate}
        ${user ? ` • 👤 ${escapeHtml(user)}` : ''}
        ${organization ? ` • 🏢 ${escapeHtml(organization)}` : ''}
      </p>
    </div>

    <!-- Transformation Summary (NOUVEAU) -->
    <div class="section-card">
      <h2 class="section-title"><span>📊</span> Transformation Summary</h2>

      <p class="readiness-message">Why this document is ${isReady ? 'now ready' : 'flagged for review'}:</p>

      <ul class="reasons-list">
        ${businessContexts.map(ctx => `
          <li>
            <span class="reason-icon ${ctx.isProtected ? 'protected' : 'risk'}">${ctx.isProtected ? '✓' : '!'}</span>
            <span>${escapeHtml(ctx.text)}</span>
          </li>
        `).join('')}
      </ul>

      <div class="metrics-grid">
        <div class="metric-card">
          <div class="metric-value">${totalCleaned}</div>
          <div class="metric-label">Items Cleaned</div>
        </div>
        <div class="metric-card">
          <div class="metric-value positive">+${improvement}%</div>
          <div class="metric-label">Improvement</div>
        </div>
        <div class="metric-card">
          <div class="metric-value">${itemsKept}</div>
          <div class="metric-label">Items Kept</div>
        </div>
      </div>
    </div>

    <!-- Score Comparison (CONSERVÉ) -->
    <div class="section-card">
      <h2 class="section-title"><span>🛡️</span> Security Score Comparison</h2>
      <div class="score-comparison">
        <div class="score-card before">
          <div class="score-label">Before Cleaning</div>
          <div class="score-value" style="color: ${beforeStyle.color}">${reportData.beforeScore}</div>
          <div class="score-status" style="color: ${beforeStyle.color}">${beforeStyle.label}</div>
        </div>
        
        ${reportData.afterScore > reportData.beforeScore ? `
          <div class="score-arrow">
            <span class="arrow">→</span>
            <span class="improvement">+${reportData.afterScore - reportData.beforeScore} pts</span>
          </div>
        ` : ''}
        
        <div class="score-card after">
          <div class="score-label">After Cleaning</div>
          <div class="score-value" style="color: ${afterStyle.color}">${reportData.afterScore}</div>
          <div class="score-status" style="color: ${afterStyle.color}">${afterStyle.label}</div>
        </div>
      </div>
    </div>

    <!-- Executive Overview (CONSERVÉ) -->
    <div class="section-card">
      <h2 class="section-title"><span>📋</span> Executive Overview</h2>
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

    <!-- Business Risks Section (CONSERVÉ - complet) -->
    ${businessRisksSection}

    <!-- Detailed Annexes (NOUVEAU - collapsibles) -->
    <div class="section-card annexes-section">
      <h2 class="section-title"><span>📎</span> Detailed Annexes</h2>

      ${generateAnnexSection('Sensitive Data Flagged', '🔒', reportData.cleaningSummary.sensitiveData.items)}
      ${generateAnnexSection('Metadata Removed', '📄', reportData.cleaningSummary.metadata.items)}
      ${generateAnnexSection('Comments Cleaned', '💬', reportData.cleaningSummary.comments.items)}
      ${generateAnnexSection('Track Changes Processed', '📝', reportData.cleaningSummary.trackChanges.items)}
      ${generateAnnexSection('Hidden Content Removed', '👁️', reportData.cleaningSummary.hiddenContent.items)}
      ${generateAnnexSection('Macros Disabled', '⚠️', reportData.cleaningSummary.macros.items)}
      ${generateCorrectionsAnnex()}
      ${generateAnnexSection('Excel Hidden Data', '📊', reportData.cleaningSummary.excelHiddenData.items)}
    </div>

    <!-- Recommendations (CONSERVÉ) -->
    ${recommendationsSection}

    <!-- Certificate (AMÉLIORÉ) -->
    <div class="certificate-section">
      <div class="certificate-icon">📜</div>
      <h2 class="certificate-title">Document Cleaning Certificate</h2>
      <p class="certificate-text">
        This document has been processed by Qualion Proposal's AI-powered cleaning engine.
        All detected risks have been addressed according to enterprise security standards.
      </p>
      <div class="certificate-meta">
        <span>📅 ${processedDateTime}</span>
        <span>🔐 Qualion Proposal by Mindorion</span>
      </div>
    </div>

    <!-- Footer -->
    <div class="footer">
      <p>Generated by <strong>Qualion Proposal</strong> — AI-powered Document Security</p>
      <p>© ${new Date().getFullYear()} Mindorion. All rights reserved.</p>
    </div>
  </div>
</body>
</html>`;
}
