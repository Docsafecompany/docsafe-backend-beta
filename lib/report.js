// lib/report.js
// Génère un rapport HTML et JSON structuré pour Qualion-Doc - v3.0 Enterprise-grade
// Avec Executive Overview (9 catégories), Compliance Status, et sections détaillées
import crypto from 'crypto';

// ============================================================
// HELPER FUNCTIONS
// ============================================================

const fmt = (n) => typeof n === 'number' ? n : 0;

function escapeHtml(s = '') {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function getScoreStyle(score) {
  if (score >= 90) return { color: '#22c55e', label: 'Safe', bg: '#dcfce7' };
  if (score >= 70) return { color: '#84cc16', label: 'Low Risk', bg: '#ecfccb' };
  if (score >= 50) return { color: '#eab308', label: 'Medium Risk', bg: '#fef9c3' };
  if (score >= 25) return { color: '#f97316', label: 'High Risk', bg: '#ffedd5' };
  return { color: '#ef4444', label: 'Critical Risk', bg: '#fee2e2' };
}

function getRiskLevelFromScore(score) {
  if (score >= 90) return 'low';
  if (score >= 70) return 'medium';
  if (score >= 50) return 'high';
  return 'critical';
}

function getComplianceStatus(afterScore, criticalIssues = 0) {
  if (afterScore >= 90 && criticalIssues === 0) return 'safe';
  if (afterScore >= 50) return 'attention';
  return 'not-ready';
}

function getRiskLevelEmoji(level) {
  switch (level) {
    case 'critical': return '🔥';
    case 'high': return '🛑';
    case 'medium': return '🔶';
    case 'low': return '🟢';
    default: return '🟡';
  }
}

function getRiskLevelColor(level) {
  switch (level) {
    case 'critical': return '#dc2626';
    case 'high': return '#ef4444';
    case 'medium': return '#f59e0b';
    case 'low': return '#10b981';
    default: return '#6b7280';
  }
}

function getTypeBadge(type) {
  const badges = {
    'fragmented_word': { icon: '🔗', label: 'Fragment', class: 'type-fragment' },
    'fragment': { icon: '🔗', label: 'Fragment', class: 'type-fragment' },
    'spelling': { icon: '✏️', label: 'Spelling', class: 'type-spelling' },
    'grammar': { icon: '📝', label: 'Grammar', class: 'type-grammar' },
    'punctuation': { icon: '⚫', label: 'Punct.', class: 'type-punctuation' },
    'ai_correction': { icon: '🤖', label: 'AI Fix', class: 'type-ai' },
    'ai': { icon: '🤖', label: 'AI Fix', class: 'type-ai' }
  };
  return badges[type] || badges['spelling'];
}

// ============================================================
// BUILD REPORT DATA (JSON) - Enterprise-grade
// ============================================================

/**
 * Génère les données structurées du rapport (pour JSON)
 * Compatible avec PremiumReportData côté frontend
 * @param {Object} params - Paramètres du rapport
 * @returns {Object} - Données structurées du rapport
 */
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
  organization = null
}) {
  // ============================================================
  // SCORES avec fallbacks
  // ============================================================
  const beforeScore = beforeRiskScore ?? analysis?.summary?.beforeRiskScore ?? 50;
  const afterScore = afterRiskScore ?? Math.min(100, beforeScore + 25);
  const riskReduction = beforeScore > 0 ? Math.round(((afterScore - beforeScore) / (100 - beforeScore)) * 100) : 0;
  
  const totalIssues = fmt(analysis?.summary?.totalIssues);
  const criticalIssues = fmt(analysis?.summary?.critical) || fmt(analysis?.summary?.criticalIssues);
  const recommendations = analysis?.summary?.recommendations || [];
  
  // Compliance status
  const complianceStatus = getComplianceStatus(afterScore, criticalIssues);
  
  // ============================================================
  // COLLECTER toutes les corrections uniques
  // ============================================================
  const allCorrections = [];
  const seenCorrections = new Set();
  
  const addCorrection = (before, after, type, context, location) => {
    if (!before || !after) return;
    const beforeStr = String(before).trim();
    const afterStr = String(after).trim();
    if (!beforeStr || !afterStr || beforeStr === afterStr) return;
    
    const key = `${beforeStr}|${afterStr}`;
    if (seenCorrections.has(key)) return;
    seenCorrections.add(key);
    
    allCorrections.push({
      id: crypto.randomUUID(),
      original: beforeStr,
      corrected: afterStr,
      type: type || 'spelling',
      context: context || '',
      location: location || 'Document'
    });
  };
  
  // Ajouter spellingErrors
  if (Array.isArray(spellingErrors) && spellingErrors.length > 0) {
    spellingErrors.forEach(err => {
      const before = err.error || err.word || err.original || err.before || '';
      const after = err.correction || err.suggestion || err.corrected || err.after || '';
      addCorrection(before, after, err.type || 'spelling', err.context || '', err.location || '');
    });
  }
  
  // Ajouter correction.examples
  if (correction?.examples && Array.isArray(correction.examples)) {
    correction.examples.forEach(ex => {
      const before = ex.before || ex.error || ex.original || '';
      const after = ex.after || ex.correction || ex.corrected || '';
      addCorrection(before, after, ex.type || 'ai_correction', '', '');
    });
  }
  
  // ============================================================
  // CONSTRUIRE les items détaillés pour chaque catégorie
  // ============================================================
  
  // 1. Sensitive Data items (Confidential & Sensitive Information)
  const sensitiveDataItems = [];
  const sensitiveDataTypes = new Set();
  if (analysis?.detections?.sensitiveData && Array.isArray(analysis.detections.sensitiveData)) {
    analysis.detections.sensitiveData.forEach(sd => {
      const type = sd.type || sd.dataType || 'Sensitive';
      sensitiveDataTypes.add(type);
      const preview = sd.value || sd.match || sd.text || '';
      const maskedValue = preview.length > 3 
        ? preview.substring(0, 3) + '***' + preview.substring(preview.length - 2)
        : '***';
      sensitiveDataItems.push({ 
        label: type.replace(/_/g, ' ').toUpperCase(), 
        value: maskedValue,
        location: sd.location || 'Document body'
      });
    });
  }
  
  // 2. Metadata items
  const metadataItems = [];
  if (analysis?.detections?.metadata && Array.isArray(analysis.detections.metadata)) {
    analysis.detections.metadata.forEach(m => {
      const label = m.key || m.type || m.name || 'Property';
      const value = m.value ? String(m.value).substring(0, 50) : undefined;
      metadataItems.push({ label, value, location: m.location || 'Document Properties' });
    });
  }
  
  // 3. Comments items (includes track changes and speaker notes)
  const commentItems = [];
  if (analysis?.detections?.comments && Array.isArray(analysis.detections.comments)) {
    analysis.detections.comments.forEach(c => {
      const author = c.author || c.authorName || 'Unknown';
      const content = c.text || c.content || c.comment || '';
      const type = c.type || 'comment';
      const typeLabel = type === 'speaker_note' ? '🎤 Speaker Note' 
                      : type === 'tracked_change' ? '📝 Track Change' 
                      : '💬 Comment';
      commentItems.push({ 
        label: `${typeLabel} by ${author}`, 
        value: content.length > 60 ? content.substring(0, 60) + '...' : content,
        location: c.location || 'Document'
      });
    });
  }
  
  // Legacy: Track changes items (si séparé)
  const trackChangesItems = [];
  if (analysis?.detections?.trackChanges && Array.isArray(analysis.detections.trackChanges)) {
    analysis.detections.trackChanges.forEach(tc => {
      const type = tc.type || tc.changeType || 'change';
      const typeEmoji = type === 'deletion' ? '🔴' : type === 'insertion' ? '🟢' : '🔄';
      const original = tc.originalText || '';
      const newText = tc.newText || '';
      const text = tc.text || (original && newText ? `"${original}" → "${newText}"` : '');
      trackChangesItems.push({ 
        label: `${typeEmoji} ${type.charAt(0).toUpperCase() + type.slice(1)} by ${tc.author || 'Unknown'}`, 
        value: text.length > 60 ? text.substring(0, 60) + '...' : text,
        location: tc.location || 'Document body'
      });
    });
  }
  
  // 4. Hidden content items
  const hiddenContentItems = [];
  if (analysis?.detections?.hiddenContent && Array.isArray(analysis.detections.hiddenContent)) {
    analysis.detections.hiddenContent.forEach(hc => {
      const type = hc.type || hc.elementType || 'hidden_element';
      const typeLabel = {
        'vanished_text': '👻 Vanished Text',
        'white_text': '⬜ White Text',
        'invisible_text': '🔍 Invisible Text',
        'hidden_slide': '📊 Hidden Slide',
        'off_slide_content': '📐 Off-slide Content',
        'embedded_file': '📎 Embedded File'
      }[type] || '❓ Hidden Element';
      hiddenContentItems.push({ 
        label: typeLabel, 
        value: hc.description || '',
        location: hc.location || 'Document'
      });
    });
  }
  
  // 5. Visual Objects items
  const visualObjectsItems = [];
  if (analysis?.detections?.visualObjects && Array.isArray(analysis.detections.visualObjects)) {
    analysis.detections.visualObjects.forEach(vo => {
      const type = vo.type || 'visual_object';
      const typeLabel = {
        'shape_covering_text': '🔲 Shape Covering Text',
        'missing_alt_text': '♿ Missing Alt Text'
      }[type] || '🎨 Visual Object';
      visualObjectsItems.push({ 
        label: typeLabel, 
        value: vo.description || '',
        location: vo.location || 'Document'
      });
    });
  }
  
  // 6. Orphan Data items
  const orphanDataItems = [];
  if (analysis?.detections?.orphanData && Array.isArray(analysis.detections.orphanData)) {
    analysis.detections.orphanData.forEach(od => {
      const type = od.type || 'orphan_data';
      const typeLabel = {
        'broken_link': '🔗 Broken Link',
        'empty_page': '📄 Empty Page',
        'trailing_whitespace': '⬜ Trailing Whitespace'
      }[type] || '🗑️ Orphan Data';
      orphanDataItems.push({ 
        label: typeLabel, 
        value: od.value || od.description || '',
        location: od.location || 'Document'
      });
    });
  }
  
  // 7. Macros items
  const macroItems = [];
  if (analysis?.detections?.macros && Array.isArray(analysis.detections.macros)) {
    analysis.detections.macros.forEach(m => {
      const type = m.type || 'VBA Macro';
      macroItems.push({ 
        label: `⚠️ ${type}`, 
        value: m.description || 'Executable code detected',
        location: m.location || 'VBA Project'
      });
    });
  }
  
  // 8. Excel Hidden Data items
  const excelHiddenDataItems = [];
  if (analysis?.detections?.excelHiddenData && Array.isArray(analysis.detections.excelHiddenData)) {
    analysis.detections.excelHiddenData.forEach(ed => {
      const type = ed.type || 'excel_hidden';
      const typeLabel = {
        'hidden_sheet': '📊 Hidden Sheet',
        'very_hidden_sheet': '🔒 Very Hidden Sheet',
        'hidden_column': '📏 Hidden Column',
        'hidden_row': '📏 Hidden Row',
        'hidden_formula': '🔢 Hidden Formula'
      }[type] || '📊 Excel Hidden Data';
      excelHiddenDataItems.push({ 
        label: typeLabel, 
        value: ed.name || ed.description || '',
        location: ed.location || 'Spreadsheet'
      });
    });
  }
  
  // Legacy: Hidden sheets & columns
  if (analysis?.detections?.hiddenSheets && Array.isArray(analysis.detections.hiddenSheets)) {
    analysis.detections.hiddenSheets.forEach(hs => {
      excelHiddenDataItems.push({ 
        label: `📊 Hidden Sheet: ${hs.sheetName || 'Unknown'}`, 
        value: hs.type === 'very_hidden' ? 'Very Hidden' : 'Hidden',
        location: 'Workbook'
      });
    });
  }
  
  // 9. Spelling/Grammar items
  const spellingItems = allCorrections.map(c => ({
    label: c.original,
    value: `→ ${c.corrected}`,
    location: c.location
  }));
  
  // ============================================================
  // EXECUTIVE OVERVIEW - 9 catégories
  // ============================================================
  
  const calculateCategoryStats = (items, cleaned = null, riskLevel = 'medium') => {
    const found = items.length;
    const cleanedCount = cleaned !== null ? cleaned : found;
    return {
      found,
      cleaned: cleanedCount,
      remaining: Math.max(0, found - cleanedCount),
      riskLevel,
      items
    };
  };
  
  // Déterminer le niveau de risque pour chaque catégorie
  const sensitiveDataRiskLevel = sensitiveDataItems.some(i => 
    i.label.includes('IBAN') || i.label.includes('SSN') || i.label.includes('PRICING')
  ) ? 'critical' : sensitiveDataItems.length > 0 ? 'high' : 'low';
  
  const executiveOverview = {
    // 1. Confidential & Sensitive Information
    confidentialInfo: calculateCategoryStats(
      sensitiveDataItems,
      fmt(cleaning?.sensitiveDataMasked) || sensitiveDataItems.length,
      sensitiveDataRiskLevel
    ),
    
    // 2. Metadata Exposure
    metadataExposure: calculateCategoryStats(
      metadataItems,
      fmt(cleaning?.metaRemoved) || metadataItems.length,
      metadataItems.length > 3 ? 'medium' : 'low'
    ),
    
    // 3. Comments & Review Traces
    commentsReview: calculateCategoryStats(
      [...commentItems, ...trackChangesItems],
      fmt(cleaning?.commentsXmlRemoved) + fmt(cleaning?.commentMarkersRemoved) || commentItems.length + trackChangesItems.length,
      commentItems.length > 0 || trackChangesItems.length > 0 ? 'medium' : 'low'
    ),
    
    // 4. Hidden & Embedded Content
    hiddenContent: calculateCategoryStats(
      hiddenContentItems,
      fmt(cleaning?.hiddenRemoved) || hiddenContentItems.length,
      hiddenContentItems.some(i => i.label.includes('White Text') || i.label.includes('Vanished')) ? 'high' : 'low'
    ),
    
    // 5. Grammar, Tone & Clarity
    grammarTone: calculateCategoryStats(
      spellingItems.slice(0, 20),
      allCorrections.length,
      allCorrections.length > 10 ? 'medium' : 'low'
    ),
    
    // 6. Visual Objects & Diagrams
    visualObjects: calculateCategoryStats(
      visualObjectsItems,
      visualObjectsItems.length,
      visualObjectsItems.some(i => i.label.includes('Covering')) ? 'medium' : 'low'
    ),
    
    // 7. Senseless / Orphan Data
    orphanData: calculateCategoryStats(
      orphanDataItems,
      orphanDataItems.length,
      'low'
    ),
    
    // 8. Macro / Script Threats
    macroThreats: calculateCategoryStats(
      macroItems,
      fmt(cleaning?.macrosRemoved) || macroItems.length,
      macroItems.length > 0 ? 'critical' : 'low'
    ),
    
    // 9. Excel Hidden Data
    excelHiddenData: calculateCategoryStats(
      excelHiddenDataItems,
      excelHiddenDataItems.length,
      excelHiddenDataItems.length > 0 ? 'medium' : 'low'
    )
  };
  
  // ============================================================
  // CLEANING SUMMARY avec items et scoreImpact
  // ============================================================
  const cleaningSummary = {
    metadata: {
      count: metadataItems.length || fmt(cleaning?.metaRemoved) || 0,
      items: metadataItems,
      scoreImpact: scoreImpacts.metadata || Math.min(10, metadataItems.length * 2)
    },
    comments: {
      count: commentItems.length || fmt(cleaning?.commentsXmlRemoved) + fmt(cleaning?.commentMarkersRemoved) || 0,
      items: commentItems,
      scoreImpact: scoreImpacts.comments || Math.min(15, commentItems.length * 3)
    },
    trackChanges: {
      count: trackChangesItems.length || fmt(cleaning?.revisionsAccepted?.deletionsRemoved) + fmt(cleaning?.revisionsAccepted?.insertionsUnwrapped) || 0,
      items: trackChangesItems,
      deletions: fmt(cleaning?.revisionsAccepted?.deletionsRemoved) || 0,
      insertions: fmt(cleaning?.revisionsAccepted?.insertionsUnwrapped) || 0,
      scoreImpact: scoreImpacts.trackChanges || Math.min(10, trackChangesItems.length * 2)
    },
    hiddenContent: {
      count: hiddenContentItems.length || fmt(cleaning?.hiddenRemoved) || 0,
      items: hiddenContentItems,
      scoreImpact: scoreImpacts.hiddenContent || Math.min(20, hiddenContentItems.length * 5)
    },
    macros: {
      count: macroItems.length || fmt(cleaning?.macrosRemoved) || 0,
      items: macroItems,
      scoreImpact: scoreImpacts.macros || (macroItems.length > 0 ? 30 : 0)
    },
    sensitiveData: {
      count: sensitiveDataItems.length,
      items: sensitiveDataItems,
      types: Array.from(sensitiveDataTypes),
      scoreImpact: scoreImpacts.sensitiveData || Math.min(25, sensitiveDataItems.length * 5)
    },
    // Nouvelles catégories
    spellingGrammar: {
      count: allCorrections.length,
      items: spellingItems.slice(0, 10),
      scoreImpact: scoreImpacts.spellingGrammar || Math.min(5, allCorrections.length)
    },
    visualObjects: {
      count: visualObjectsItems.length,
      items: visualObjectsItems,
      scoreImpact: scoreImpacts.visualObjects || Math.min(5, visualObjectsItems.length * 2)
    },
    orphanData: {
      count: orphanDataItems.length,
      items: orphanDataItems,
      scoreImpact: scoreImpacts.orphanData || Math.min(3, orphanDataItems.length)
    },
    excelHiddenData: {
      count: excelHiddenDataItems.length,
      items: excelHiddenDataItems,
      scoreImpact: scoreImpacts.excelHiddenData || Math.min(15, excelHiddenDataItems.length * 4)
    }
  };
  
  // Calculer le total des éléments supprimés
  const totalRemoved = Object.values(cleaningSummary).reduce((sum, cat) => sum + (cat.count || 0), 0);
  
  // ============================================================
  // RISQUES DÉTECTÉS - groupés par sévérité
  // ============================================================
  const risksDetected = [];
  
  // Ajouter les risques depuis sensitiveData
  if (analysis?.detections?.sensitiveData && Array.isArray(analysis.detections.sensitiveData)) {
    analysis.detections.sensitiveData.forEach(sd => {
      risksDetected.push({
        id: sd.id || crypto.randomUUID(),
        severity: sd.severity || 'high',
        type: sd.type || 'sensitive_data',
        description: sd.description || `Detected ${sd.type || 'sensitive data'}`,
        context: sd.context || sd.value || '',
        action: 'flagged'
      });
    });
  }
  
  // Ajouter les risques de compliance
  if (analysis?.detections?.complianceRisks && Array.isArray(analysis.detections.complianceRisks)) {
    analysis.detections.complianceRisks.forEach(cr => {
      risksDetected.push({
        id: cr.id || crypto.randomUUID(),
        severity: cr.severity || 'high',
        type: cr.type || 'compliance',
        description: cr.description || cr.message || 'Compliance risk detected',
        context: cr.context || '',
        action: 'flagged'
      });
    });
  }
  
  // Ajouter les macros comme risques critiques
  if (analysis?.detections?.macros && analysis.detections.macros.length > 0) {
    risksDetected.push({
      id: crypto.randomUUID(),
      severity: 'critical',
      type: 'macros',
      description: `${analysis.detections.macros.length} macro(s) detected - potential security risk`,
      context: 'Macros can contain executable code',
      action: 'removed'
    });
  }
  
  // Ajouter hidden content critique
  if (analysis?.detections?.hiddenContent) {
    const criticalHidden = analysis.detections.hiddenContent.filter(h => 
      h.type === 'white_text' || h.type === 'vanished_text'
    );
    criticalHidden.forEach(ch => {
      risksDetected.push({
        id: ch.id || crypto.randomUUID(),
        severity: 'high',
        type: 'hidden_content',
        description: ch.description || 'Hidden content detected',
        context: ch.location || '',
        action: 'removed'
      });
    });
  }
  
  // ============================================================
  // RECOMMENDATIONS formatées
  // ============================================================
  const formattedRecommendations = [];
  if (Array.isArray(recommendations)) {
    recommendations.forEach(rec => {
      if (typeof rec === 'string') {
        formattedRecommendations.push(rec);
      } else if (rec.text) {
        formattedRecommendations.push(`${rec.icon || ''} ${rec.text}`);
      }
    });
  }
  
  // ============================================================
  // MÉTRIQUES ADDITIONNELLES
  // ============================================================
  const clarityScoreImprovement = allCorrections.length > 0 
    ? Math.min(25, Math.round(allCorrections.length * 2.5))
    : 0;
  
  const documentSizeReduction = fmt(cleaning?.mediaDeleted) + fmt(cleaning?.embeddedFilesRemoved) > 0
    ? Math.min(30, (fmt(cleaning?.mediaDeleted) + fmt(cleaning?.embeddedFilesRemoved)) * 5)
    : 0;
  
  console.log(`[REPORT JSON v3.0] Generated: beforeScore=${beforeScore}, afterScore=${afterScore}, riskReduction=${riskReduction}%, complianceStatus=${complianceStatus}, corrections=${allCorrections.length}`);
  
  // ============================================================
  // RETOUR FINAL - Compatible avec PremiumReportData
  // ============================================================
  return {
    document: {
      name: filename,
      type: ext?.toUpperCase() || 'UNKNOWN',
      processedAt: new Date().toISOString(),
      user: user || null,
      organization: organization || null
    },
    beforeScore: Math.round(beforeScore),
    afterScore: Math.round(afterScore),
    riskReduction: Math.max(0, riskReduction),
    complianceStatus,
    summary: {
      totalIssuesFound: totalIssues,
      elementsRemoved: totalRemoved,
      correctionsApplied: allCorrections.length,
      criticalRisksResolved: criticalIssues
    },
    executiveOverview,
    clarityScoreImprovement,
    documentSizeReduction,
    cleaningSummary,
    textCorrections: allCorrections,
    risksDetected,
    recommendations: formattedRecommendations
  };
}

// ============================================================
// BUILD REPORT HTML - Enterprise-grade Template
// ============================================================

/**
 * Génère un rapport HTML Enterprise-grade complet
 * @param {Object} params - Paramètres du rapport
 * @returns {string} - HTML complet du rapport
 */
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
  organization = null
}) {
  // Générer les données structurées d'abord
  const reportData = buildReportData({
    filename, ext, policy, cleaning, correction, analysis,
    spellingErrors, beforeRiskScore, afterRiskScore,
    user, organization
  });
  
  const beforeStyle = getScoreStyle(reportData.beforeScore);
  const afterStyle = getScoreStyle(reportData.afterScore);
  
  // Compliance Status Badge
  const complianceColors = {
    'safe': { emoji: '🟢', label: 'Safe', color: '#10b981', bg: '#dcfce7' },
    'attention': { emoji: '🟡', label: 'Attention Required', color: '#f59e0b', bg: '#fef3c7' },
    'not-ready': { emoji: '🔴', label: 'Not Client-Ready', color: '#ef4444', bg: '#fee2e2' }
  };
  const compliance = complianceColors[reportData.complianceStatus];
  
  // ============================================================
  // EXECUTIVE OVERVIEW TABLE ROWS
  // ============================================================
  const getCategoryRow = (name, stats) => {
    if (!stats || (stats.found === 0 && stats.cleaned === 0)) return '';
    return `
      <tr>
        <td>${name}</td>
        <td class="center">${stats.found}</td>
        <td class="center">${stats.cleaned}</td>
        <td class="center">${stats.remaining}</td>
        <td class="center">
          <span class="risk-level" style="background: ${getRiskLevelColor(stats.riskLevel)}20; color: ${getRiskLevelColor(stats.riskLevel)}">
            ${getRiskLevelEmoji(stats.riskLevel)} ${stats.riskLevel.charAt(0).toUpperCase() + stats.riskLevel.slice(1)}
          </span>
        </td>
      </tr>
    `;
  };
  
  const executiveOverviewRows = `
    ${getCategoryRow('Confidential & Sensitive Information', reportData.executiveOverview.confidentialInfo)}
    ${getCategoryRow('Metadata Exposure', reportData.executiveOverview.metadataExposure)}
    ${getCategoryRow('Comments & Review Traces', reportData.executiveOverview.commentsReview)}
    ${getCategoryRow('Hidden & Embedded Content', reportData.executiveOverview.hiddenContent)}
    ${getCategoryRow('Grammar, Tone & Clarity', reportData.executiveOverview.grammarTone)}
    ${getCategoryRow('Visual Objects & Diagrams', reportData.executiveOverview.visualObjects)}
    ${getCategoryRow('Senseless / Orphan Data', reportData.executiveOverview.orphanData)}
    ${getCategoryRow('Macro / Script Threats', reportData.executiveOverview.macroThreats)}
    ${getCategoryRow('Excel Hidden Data', reportData.executiveOverview.excelHiddenData)}
  `;
  
  // ============================================================
  // SECTION ITEMS HTML HELPER
  // ============================================================
  const getSectionItemsHtml = (items) => {
    if (!items || items.length === 0) return '';
    return `
      <ul class="items-list">
        ${items.slice(0, 10).map(item => `
          <li>
            <span class="item-label">${escapeHtml(item.label)}</span>
            ${item.value ? `<span class="item-value">${escapeHtml(item.value)}</span>` : ''}
            ${item.location ? `<span class="item-location">📍 ${escapeHtml(item.location)}</span>` : ''}
          </li>
        `).join('')}
        ${items.length > 10 ? `<li class="more-items">... and ${items.length - 10} more items</li>` : ''}
      </ul>
    `;
  };
  
  // ============================================================
  // DETAILED SECTIONS
  // ============================================================
  
  // Section 1: Confidential & Sensitive Information
  const section1 = reportData.executiveOverview.confidentialInfo.found > 0 ? `
    <section class="detail-section">
      <h2>1️⃣ Confidential & Sensitive Information Exposure (Critical Risk)</h2>
      <p class="section-intro">Includes:</p>
      <ul class="check-list">
        <li>Employee names, internal emails</li>
        <li>Client names mixing (cross-contamination)</li>
        <li>Project codes, equipment IDs</li>
        <li>Price / margin / financial exposure</li>
        <li>Personal data (GDPR)</li>
        <li>IP disclosures</li>
      </ul>
      ${getSectionItemsHtml(reportData.cleaningSummary.sensitiveData.items)}
      <div class="action-taken">
        📌 Action Taken → Removed / masked all detected confidential elements
      </div>
    </section>
  ` : '';
  
  // Section 2: Metadata Leakage
  const section2 = reportData.executiveOverview.metadataExposure.found > 0 ? `
    <section class="detail-section">
      <h2>2️⃣ Metadata Leakage (Compliance Risk)</h2>
      <p class="section-intro">Removed metadata:</p>
      <ul class="check-list">
        <li>Author, company, device info</li>
        <li>Internal file paths + SharePoint/OneDrive links</li>
        <li>Time stamps, revision IDs</li>
        <li>Software versions, printers, localization</li>
      </ul>
      ${getSectionItemsHtml(reportData.cleaningSummary.metadata.items)}
      <div class="action-taken">
        📌 100% metadata sanitized
        <br>📝 Detail list provided in Annex A
      </div>
    </section>
  ` : '';
  
  // Section 3: Comments & Revision Traces
  const section3 = reportData.executiveOverview.commentsReview.found > 0 ? `
    <section class="detail-section">
      <h2>3️⃣ Comments & Revision Traces (Reputation Risk)</h2>
      <p class="section-intro">Cleaned:</p>
      <ul class="check-list">
        <li>Word comments</li>
        <li>PPT speaker notes</li>
        <li>"Resolve Later" flags</li>
        <li>Suggestion marks</li>
        <li>Tracked changes not accepted</li>
      </ul>
      ${getSectionItemsHtml(reportData.cleaningSummary.comments.items)}
      ${getSectionItemsHtml(reportData.cleaningSummary.trackChanges.items)}
      <div class="action-taken">
        📌 All comments removed automatically
      </div>
    </section>
  ` : '';
  
  // Section 4: Hidden & Embedded Content
  const section4 = reportData.executiveOverview.hiddenContent.found > 0 ? `
    <section class="detail-section">
      <h2>4️⃣ Hidden & Embedded Content</h2>
      <p class="section-intro">Detected and cleaned:</p>
      <ul class="check-list">
        <li>Hidden text / hidden elements</li>
        <li>Off-slide graphics in PPT</li>
        <li>Hidden sheets / rows / columns in Excel</li>
        <li>White-on-white text</li>
        <li>Embedded attachments</li>
      </ul>
      ${getSectionItemsHtml(reportData.cleaningSummary.hiddenContent.items)}
      <div class="action-taken">
        📌 Removed ${reportData.executiveOverview.hiddenContent.cleaned} invisible object(s)
      </div>
    </section>
  ` : '';
  
  // Section 5: Grammar & Professional Tone
  const section5 = reportData.textCorrections.length > 0 ? `
    <section class="detail-section">
      <h2>5️⃣ Grammar & Professional Tone (AI-Powered)</h2>
      <p class="section-intro">Corrected:</p>
      <ul class="check-list">
        <li>Grammar & spelling</li>
        <li>Tone standardization (Executive / Pharma / Engineering)</li>
        <li>Inconsistent terminology</li>
        <li>Numeric format / scientific units</li>
        <li>Punctuation + spacing</li>
      </ul>
      
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
          ${reportData.textCorrections.slice(0, 20).map((c, i) => {
            const badge = getTypeBadge(c.type);
            return `
              <tr>
                <td class="center">${i + 1}</td>
                <td><span class="type-badge ${badge.class}">${badge.icon} ${badge.label}</span></td>
                <td class="original">${escapeHtml(c.original)}</td>
                <td class="corrected">${escapeHtml(c.corrected)}</td>
              </tr>
            `;
          }).join('')}
          ${reportData.textCorrections.length > 20 ? `
            <tr>
              <td colspan="4" class="center more-items">... and ${reportData.textCorrections.length - 20} more corrections</td>
            </tr>
          ` : ''}
        </tbody>
      </table>
      
      <div class="action-taken">
        📌 Clarity Score Improvement: +${reportData.clarityScoreImprovement}%
        <br>📌 Optional policy tone (enterprise mode)
      </div>
    </section>
  ` : '';
  
  // Section 6: Visual Objects & Diagrams
  const section6 = reportData.executiveOverview.visualObjects.found > 0 ? `
    <section class="detail-section">
      <h2>6️⃣ Visual Objects & Diagrams</h2>
      <p class="section-intro">Detected and addressed:</p>
      <ul class="check-list">
        <li>Removed text hidden under shapes</li>
        <li>Updated visual alt-text (accessibility)</li>
        <li>Sanitized broken callouts & outdated references</li>
      </ul>
      ${getSectionItemsHtml(reportData.cleaningSummary.visualObjects.items)}
      <div class="action-taken">
        📌 ${reportData.executiveOverview.visualObjects.cleaned} corrections
      </div>
    </section>
  ` : '';
  
  // Section 7: Senseless / Orphan Data
  const section7 = reportData.executiveOverview.orphanData.found > 0 ? `
    <section class="detail-section">
      <h2>7️⃣ Senseless / Orphan Data (Optimization)</h2>
      <p class="section-intro">Cleaned:</p>
      <ul class="check-list">
        <li>Empty pages</li>
        <li>Deprecated tables / leftover numbers</li>
        <li>Broken hyperlinks</li>
        <li>Corrupted content</li>
        <li>Trailing whitespace</li>
      </ul>
      ${getSectionItemsHtml(reportData.cleaningSummary.orphanData.items)}
      <div class="action-taken">
        📌 Document size reduced: ${reportData.documentSizeReduction}%
        <br>⚙️ Impact: performance + clarity + compliance
      </div>
    </section>
  ` : '';
  
  // Section 8: Macro / Script Threats
  const section8 = reportData.executiveOverview.macroThreats.found > 0 ? `
    <section class="detail-section danger-section">
      <h2>8️⃣ Script, Macro & Embedded Threat Scan</h2>
      <p class="section-intro">Security analysis:</p>
      <ul class="check-list">
        <li>Malicious scripts analysis</li>
        <li>Auto-macro disabling</li>
        <li>OLE detachment verification</li>
      </ul>
      ${getSectionItemsHtml(reportData.cleaningSummary.macros.items)}
      <div class="action-taken danger">
        ⚠️ Potential threats removed
      </div>
    </section>
  ` : '';
  
  // Section 9: Excel Context
  const section9 = (ext === 'xlsx' || ext === 'xls') && reportData.executiveOverview.excelHiddenData.found > 0 ? `
    <section class="detail-section">
      <h2>9️⃣ Excel Context</h2>
      <p class="section-intro">Excel-specific cleaning:</p>
      <ul class="check-list">
        <li>Hidden formulas exposing sensitive data</li>
        <li>Sheets hidden / very hidden</li>
        <li>Pivot caches / track changes removed</li>
      </ul>
      ${getSectionItemsHtml(reportData.cleaningSummary.excelHiddenData.items)}
      <div class="action-taken">
        📌 Full cleanup executed
      </div>
    </section>
  ` : '';
  
  // Recommendations section
  const recommendationsSection = reportData.recommendations.length > 0 ? `
    <section class="detail-section recommendations-section">
      <h2>💡 Recommendations</h2>
      <ul class="recommendations-list">
        ${reportData.recommendations.map(rec => `<li>${escapeHtml(rec)}</li>`).join('')}
      </ul>
    </section>
  ` : '';

  // ============================================================
  // COMPLETE HTML TEMPLATE
  // ============================================================
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
    
    /* HEADER */
    .header-card {
      background: linear-gradient(135deg, var(--primary) 0%, #8b5cf6 100%);
      border-radius: 20px;
      padding: 32px;
      color: white;
      text-align: center;
      margin-bottom: 24px;
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
    
    /* COMPLIANCE STATUS */
    .compliance-badge {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 12px 24px;
      border-radius: 30px;
      font-size: 16px;
      font-weight: 700;
      margin-top: 16px;
      background: ${compliance.bg};
      color: ${compliance.color};
      border: 2px solid ${compliance.color}40;
    }
    
    /* SCORE COMPARISON */
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
    
    /* EXECUTIVE OVERVIEW TABLE */
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
    
    /* DETAIL SECTIONS */
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
    
    .detail-section.danger-section {
      border-left: 4px solid var(--danger);
    }
    
    .section-intro {
      color: var(--gray-600);
      margin-bottom: 12px;
      font-size: 14px;
    }
    
    .check-list {
      list-style: none;
      margin-bottom: 16px;
    }
    
    .check-list li {
      padding: 6px 0 6px 24px;
      position: relative;
      font-size: 13px;
      color: var(--gray-700);
    }
    
    .check-list li::before {
      content: '✓';
      position: absolute;
      left: 0;
      color: var(--success);
      font-weight: bold;
    }
    
    .items-list {
      list-style: none;
      background: var(--gray-50);
      border-radius: 8px;
      padding: 12px 16px;
      margin: 16px 0;
    }
    
    .items-list li {
      padding: 8px 0;
      border-bottom: 1px solid var(--gray-200);
      font-size: 13px;
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      align-items: center;
    }
    
    .items-list li:last-child { border-bottom: none; }
    
    .item-label { font-weight: 600; color: var(--gray-800); }
    .item-value { color: var(--gray-600); font-family: monospace; background: white; padding: 2px 6px; border-radius: 4px; }
    .item-location { color: var(--gray-400); font-size: 11px; }
    .more-items { color: var(--gray-500); font-style: italic; }
    
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
    
    /* CORRECTIONS TABLE */
    .corrections-table {
      width: 100%;
      border-collapse: collapse;
      margin: 16px 0;
    }
    
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
    
    /* RECOMMENDATIONS */
    .recommendations-list {
      list-style: none;
    }
    
    .recommendations-list li {
      padding: 12px 16px;
      background: var(--gray-50);
      border-radius: 8px;
      margin-bottom: 8px;
      font-size: 14px;
      color: var(--gray-700);
      border-left: 3px solid var(--primary);
    }
    
    /* FOOTER */
    .footer {
      text-align: center;
      padding: 32px 16px;
      color: var(--gray-500);
      font-size: 13px;
    }
    
    .footer a { color: var(--primary); text-decoration: none; font-weight: 500; }
    
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
      .detail-section, .score-card, .overview-card { box-shadow: none; border: 1px solid var(--gray-200); break-inside: avoid; }
    }
  </style>
</head>
<body>
  <div class="container">
    <!-- HEADER -->
    <div class="header-card">
      <div class="logo">🔐</div>
      <h1>Document Clean & Secure Report</h1>
      <p class="subtitle">Generated by Qualion-Doc — AI-powered Compliance Engine</p>
      
      <div class="doc-info">
        <div class="doc-info-item">📁 <strong>Document:</strong> ${escapeHtml(filename)}</div>
        <div class="doc-info-item">📅 <strong>Processed:</strong> ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' })}</div>
        ${user ? `<div class="doc-info-item">👤 <strong>User:</strong> ${escapeHtml(user)}</div>` : ''}
        ${organization ? `<div class="doc-info-item">🏢 <strong>Organization:</strong> ${escapeHtml(organization)}</div>` : ''}
      </div>
      
      <div class="compliance-badge">
        ${compliance.emoji} Final Compliance Status: ${compliance.label}
      </div>
    </div>
    
    <!-- SCORE COMPARISON -->
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
        
        ${reportData.afterScore > reportData.beforeScore ? `
        <div class="score-improvement">
          <div class="score-improvement-arrow">→</div>
          <div class="score-improvement-value">+${reportData.afterScore - reportData.beforeScore} pts</div>
          <div style="font-size: 11px; color: var(--gray-500);">🔥 ${reportData.riskReduction}% Risk Reduction</div>
        </div>
        ` : ''}
        
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
    
    <!-- EXECUTIVE OVERVIEW -->
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
    
    <!-- DETAILED SECTIONS -->
    ${section1}
    ${section2}
    ${section3}
    ${section4}
    ${section5}
    ${section6}
    ${section7}
    ${section8}
    ${section9}
    ${recommendationsSection}
    
    <!-- FOOTER -->
    <div class="footer">
      <p>Generated by <a href="https://mindorion.com" target="_blank">Qualion-Doc</a> by Mindorion</p>
      <p style="margin-top: 4px;">${new Date().toLocaleDateString('en-US', { 
        weekday: 'long', 
        year: 'numeric', 
        month: 'long', 
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit'
      })}</p>
    </div>
  </div>
</body>
</html>`;
}

export default { buildReportHtmlDetailed, buildReportData };

