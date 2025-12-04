// lib/report.js
// Génère un rapport HTML et JSON structuré pour Qualion-Doc - v2.4 avec JSON
import crypto from 'crypto';

/**
 * Génère les données structurées du rapport (pour JSON)
 * @param {Object} params - Paramètres du rapport
 * @returns {Object} - Données structurées du rapport
 */
export function buildReportData({ filename, ext, policy, cleaning, correction, analysis, spellingErrors = [] }) {
  const fmt = (n) => typeof n === 'number' ? n : 0;
  
  // Score avec fallbacks
  let afterScore = 100;
  if (analysis?.summary?.riskScore !== undefined && analysis?.summary?.riskScore !== null) {
    afterScore = Number(analysis.summary.riskScore);
    if (isNaN(afterScore)) afterScore = 100;
  }
  
  // Score avant nettoyage (si disponible)
  let beforeScore = analysis?.summary?.beforeRiskScore || Math.max(0, afterScore - 30);
  
  const totalIssues = fmt(analysis?.summary?.totalIssues);
  const criticalIssues = fmt(analysis?.summary?.critical) || fmt(analysis?.summary?.criticalIssues);
  const recommendations = analysis?.summary?.recommendations || [];
  
  // Collecter toutes les corrections uniques
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
  
  // Calculer les statistiques de nettoyage par catégorie
  const cleaningSummary = {
    metadata: {
      count: fmt(cleaning?.metaRemoved) || 0,
      items: cleaning?.metadataItems || []
    },
    comments: {
      count: fmt(cleaning?.commentsXmlRemoved) + fmt(cleaning?.commentMarkersRemoved) || 0,
      items: []
    },
    trackChanges: {
      count: fmt(cleaning?.revisionsAccepted?.deletionsRemoved) + fmt(cleaning?.revisionsAccepted?.insertionsUnwrapped) || 0,
      deletions: fmt(cleaning?.revisionsAccepted?.deletionsRemoved) || 0,
      insertions: fmt(cleaning?.revisionsAccepted?.insertionsUnwrapped) || 0
    },
    hiddenContent: {
      count: fmt(cleaning?.hiddenRemoved) || 0,
      items: []
    },
    embeddedObjects: {
      count: fmt(cleaning?.mediaDeleted) + fmt(cleaning?.picturesRemoved) || 0,
      media: fmt(cleaning?.mediaDeleted) || 0,
      pictures: fmt(cleaning?.picturesRemoved) || 0
    },
    macros: {
      count: fmt(cleaning?.macrosRemoved) || 0,
      items: []
    },
    sensitiveData: {
      count: 0,
      types: []
    },
    spellingGrammar: {
      count: allCorrections.length,
      items: []
    }
  };
  
  // Calculer le total des éléments supprimés
  const totalRemoved = Object.values(cleaningSummary).reduce((sum, cat) => sum + (cat.count || 0), 0);
  
  // Grouper les risques par sévérité
  const risks = analysis?.risks || [];
  const risksDetected = risks.map(risk => ({
    id: risk.id || crypto.randomUUID(),
    severity: risk.severity || 'medium',
    type: risk.type || 'unknown',
    description: risk.description || '',
    context: risk.context || '',
    action: risk.resolved ? 'removed' : 'flagged'
  }));
  
  return {
    document: {
      name: filename,
      type: ext?.toUpperCase() || 'UNKNOWN',
      processedAt: new Date().toISOString()
    },
    beforeScore: Math.round(beforeScore),
    afterScore: Math.round(afterScore),
    summary: {
      totalIssuesFound: totalIssues,
      elementsRemoved: totalRemoved,
      correctionsApplied: allCorrections.length,
      criticalRisksResolved: criticalIssues
    },
    cleaningSummary,
    textCorrections: allCorrections,
    risksDetected,
    recommendations
  };
}

/**
 * Génère un rapport HTML détaillé et stylé
 * @param {Object} params - Paramètres du rapport
 * @returns {string} - HTML complet du rapport
 */
export function buildReportHtmlDetailed({ filename, ext, policy, cleaning, correction, analysis, spellingErrors = [] }) {
  const fmt = (n) => typeof n === 'number' ? n : 0;
  
  // ============================================================
  // SCORE - Récupération robuste avec fallbacks
  // ============================================================
  let score = 100;
  if (analysis?.summary?.riskScore !== undefined && analysis?.summary?.riskScore !== null) {
    score = Number(analysis.summary.riskScore);
    if (isNaN(score)) score = 100;
  }
  
  const totalIssues = fmt(analysis?.summary?.totalIssues);
  const criticalIssues = fmt(analysis?.summary?.critical) || fmt(analysis?.summary?.criticalIssues);
  const recommendations = analysis?.summary?.recommendations || [];
  
  // Couleurs selon le score (100 = vert/safe, 0 = rouge/critical)
  const getScoreStyle = (s) => {
    if (s >= 90) return { color: '#22c55e', label: 'Safe', bg: '#dcfce7' };
    if (s >= 70) return { color: '#84cc16', label: 'Low Risk', bg: '#ecfccb' };
    if (s >= 50) return { color: '#eab308', label: 'Medium Risk', bg: '#fef9c3' };
    if (s >= 25) return { color: '#f97316', label: 'High Risk', bg: '#ffedd5' };
    return { color: '#ef4444', label: 'Critical Risk', bg: '#fee2e2' };
  };
  
  const scoreStyle = getScoreStyle(score);
  
  // ============================================================
  // COMBINER toutes les corrections de manière robuste
  // ============================================================
  const allCorrections = [];
  const seenCorrections = new Set();
  
  const addCorrection = (before, after, type, context) => {
    if (!before || !after) return;
    const beforeStr = String(before).trim();
    const afterStr = String(after).trim();
    if (!beforeStr || !afterStr || beforeStr === afterStr) return;
    
    const key = `${beforeStr}|${afterStr}`;
    if (seenCorrections.has(key)) return;
    seenCorrections.add(key);
    
    allCorrections.push({
      before: beforeStr,
      after: afterStr,
      type: type || 'spelling',
      context: context || ''
    });
  };
  
  // 1. Ajouter les spellingErrors
  if (Array.isArray(spellingErrors) && spellingErrors.length > 0) {
    console.log(`[REPORT] Processing ${spellingErrors.length} spellingErrors`);
    spellingErrors.forEach((err, i) => {
      const before = err.error || err.word || err.original || err.before || '';
      const after = err.correction || err.suggestion || err.corrected || err.after || '';
      const type = err.type || 'spelling';
      const context = err.context || '';
      
      console.log(`[REPORT] SpellingError ${i}: "${before}" → "${after}" (${type})`);
      addCorrection(before, after, type, context);
    });
  }
  
  // 2. Ajouter les correction.examples
  if (correction?.examples && Array.isArray(correction.examples) && correction.examples.length > 0) {
    console.log(`[REPORT] Processing ${correction.examples.length} correction.examples`);
    correction.examples.forEach((ex, i) => {
      const before = ex.before || ex.error || ex.original || '';
      const after = ex.after || ex.correction || ex.corrected || '';
      const type = ex.type || 'ai_correction';
      
      console.log(`[REPORT] Example ${i}: "${before}" → "${after}" (${type})`);
      addCorrection(before, after, type, '');
    });
  }
  
  console.log(`[REPORT] Total unique corrections: ${allCorrections.length}`);
  
  // ============================================================
  // GÉNÉRER les lignes de corrections
  // ============================================================
  const getTypeBadge = (type) => {
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
  };
  
  const correctionRows = allCorrections.map((corr, index) => {
    const badge = getTypeBadge(corr.type);
    
    return `<tr>
      <td class="cell-before">
        <div class="correction-number">#${index + 1}</div>
        <div class="code-block error-block">
          <span class="error-text">${escapeHtml(corr.before)}</span>
        </div>
        ${corr.context ? `<div class="context-text">"...${escapeHtml(corr.context.slice(0, 60))}..."</div>` : ''}
      </td>
      <td class="cell-after">
        <div class="code-block success-block">
          <span class="corrected-text">${escapeHtml(corr.after)}</span>
        </div>
        <span class="type-badge ${badge.class}">${badge.icon} ${badge.label}</span>
      </td>
    </tr>`;
  }).join('');
  
  const totalCorrections = allCorrections.length || fmt(correction?.changedTextNodes) || 0;
  
  const recommendationsList = recommendations.map(rec => 
    `<li>${escapeHtml(rec)}</li>`
  ).join('');
  
  // ============================================================
  // CLEANING SUMMARY selon le type de fichier
  // ============================================================
  const getCleaningList = () => {
    if (ext === 'docx') {
      return `
        <div class="stat-item">
          <span class="stat-label">Metadata removed</span>
          <span class="stat-value">${fmt(cleaning?.metaRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Comments removed</span>
          <span class="stat-value">${fmt(cleaning?.commentsXmlRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Comment markers cleaned</span>
          <span class="stat-value">${fmt(cleaning?.commentMarkersRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Deletions removed</span>
          <span class="stat-value">${fmt(cleaning?.revisionsAccepted?.deletionsRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Insertions unwrapped</span>
          <span class="stat-value">${fmt(cleaning?.revisionsAccepted?.insertionsUnwrapped)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Ink annotations removed</span>
          <span class="stat-value">${fmt(cleaning?.inkRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">VML shapes removed</span>
          <span class="stat-value">${fmt(cleaning?.vmlRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Media files deleted</span>
          <span class="stat-value">${fmt(cleaning?.mediaDeleted)}</span>
        </div>
      `;
    }
    if (ext === 'pptx') {
      return `
        <div class="stat-item">
          <span class="stat-label">Metadata removed</span>
          <span class="stat-value">${fmt(cleaning?.metaRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Comments removed</span>
          <span class="stat-value">${fmt(cleaning?.commentsXmlRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Comment relations cleaned</span>
          <span class="stat-value">${fmt(cleaning?.relsCommentsRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Ink removed</span>
          <span class="stat-value">${fmt(cleaning?.inkRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Pictures removed</span>
          <span class="stat-value">${fmt(cleaning?.picturesRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Media deleted</span>
          <span class="stat-value">${fmt(cleaning?.mediaDeleted)}</span>
        </div>
      `;
    }
    if (ext === 'pdf') {
      return `
        <div class="stat-item">
          <span class="stat-label">Metadata cleared</span>
          <span class="stat-value">${cleaning?.metadataCleared ? '✓' : '—'}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Annotations removed</span>
          <span class="stat-value">${fmt(cleaning?.annotsRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Embedded files removed</span>
          <span class="stat-value">${fmt(cleaning?.embeddedFilesRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Processing mode</span>
          <span class="stat-value">${policy?.pdfMode || 'standard'}</span>
        </div>
      `;
    }
    if (ext === 'xlsx') {
      return `
        <div class="stat-item">
          <span class="stat-label">Metadata removed</span>
          <span class="stat-value">${fmt(cleaning?.metaRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Comments removed</span>
          <span class="stat-value">${fmt(cleaning?.commentsRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Hidden sheets handled</span>
          <span class="stat-value">${fmt(cleaning?.hiddenSheetsRemoved)}</span>
        </div>
      `;
    }
    return '<div class="stat-item"><span class="stat-label">Standard cleaning applied</span><span class="stat-value">✓</span></div>';
  };

  // ============================================================
  // HTML TEMPLATE COMPLET
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
    
    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
    }
    
    body {
      font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
      background: linear-gradient(135deg, #f0f9ff 0%, #e0f2fe 50%, #f0fdf4 100%);
      min-height: 100vh;
      padding: 32px 16px;
      color: var(--gray-900);
      line-height: 1.6;
    }
    
    .container {
      max-width: 800px;
      margin: 0 auto;
    }
    
    .header-card {
      background: linear-gradient(135deg, var(--primary) 0%, #8b5cf6 100%);
      border-radius: 20px;
      padding: 32px;
      color: white;
      text-align: center;
      margin-bottom: 24px;
      box-shadow: 0 10px 40px rgba(59, 130, 246, 0.3);
    }
    
    .logo { font-size: 32px; margin-bottom: 8px; }
    .header-card h1 { font-size: 24px; font-weight: 700; margin-bottom: 4px; }
    .header-card .subtitle { opacity: 0.9; font-size: 14px; margin-bottom: 20px; }
    
    .file-badge {
      display: inline-flex;
      align-items: center;
      gap: 16px;
      background: rgba(255, 255, 255, 0.15);
      backdrop-filter: blur(10px);
      padding: 12px 24px;
      border-radius: 12px;
      font-size: 14px;
      flex-wrap: wrap;
      justify-content: center;
    }
    
    .file-badge code {
      background: rgba(255, 255, 255, 0.2);
      padding: 4px 10px;
      border-radius: 6px;
      font-family: 'Monaco', 'Consolas', monospace;
      font-size: 13px;
    }
    
    .score-card {
      background: white;
      border-radius: 20px;
      padding: 40px;
      text-align: center;
      margin-bottom: 24px;
      box-shadow: 0 4px 20px rgba(0, 0, 0, 0.08);
    }
    
    .score-circle {
      width: 140px;
      height: 140px;
      border-radius: 50%;
      background: conic-gradient(${scoreStyle.color} ${score}%, var(--gray-200) ${score}%);
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0 auto 20px;
      position: relative;
      box-shadow: 0 8px 30px ${scoreStyle.color}40;
    }
    
    .score-circle::before {
      content: '';
      width: 110px;
      height: 110px;
      background: white;
      border-radius: 50%;
      position: absolute;
    }
    
    .score-number {
      position: relative;
      z-index: 1;
      font-size: 42px;
      font-weight: 800;
      color: ${scoreStyle.color};
    }
    
    .score-label {
      display: inline-block;
      padding: 8px 20px;
      background: ${scoreStyle.bg};
      color: ${scoreStyle.color};
      border-radius: 24px;
      font-weight: 600;
      font-size: 14px;
      margin-bottom: 12px;
    }
    
    .score-description { color: var(--gray-500); font-size: 14px; }
    
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
    .score-stat-label { font-size: 12px; color: var(--gray-500); text-transform: uppercase; letter-spacing: 0.5px; }
    
    .card {
      background: white;
      border-radius: 16px;
      margin-bottom: 20px;
      box-shadow: 0 2px 12px rgba(0, 0, 0, 0.06);
      overflow: hidden;
    }
    
    .card-header {
      padding: 20px 24px;
      border-bottom: 1px solid var(--gray-100);
      display: flex;
      align-items: center;
      gap: 12px;
    }
    
    .card-header h2 { font-size: 16px; font-weight: 600; color: var(--gray-900); }
    
    .card-icon {
      width: 36px;
      height: 36px;
      border-radius: 10px;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 18px;
    }
    
    .card-icon.blue { background: #dbeafe; }
    .card-icon.green { background: #dcfce7; }
    .card-icon.orange { background: #ffedd5; }
    .card-icon.purple { background: #f3e8ff; }
    
    .card-body { padding: 20px 24px; }
    
    .stats-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; }
    
    .stat-item {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 12px 16px;
      background: var(--gray-50);
      border-radius: 10px;
    }
    
    .stat-label { color: var(--gray-700); font-size: 13px; }
    
    .stat-value {
      font-weight: 600;
      color: var(--gray-900);
      background: white;
      padding: 4px 12px;
      border-radius: 6px;
      font-size: 13px;
    }
    
    .recommendations-list { list-style: none; }
    
    .recommendations-list li {
      padding: 12px 16px;
      background: var(--gray-50);
      border-radius: 10px;
      margin-bottom: 8px;
      font-size: 14px;
      color: var(--gray-700);
      border-left: 3px solid var(--primary);
    }
    
    .recommendations-list li:last-child { margin-bottom: 0; }
    
    .corrections-table { width: 100%; border-collapse: collapse; }
    
    .corrections-table th {
      background: var(--gray-50);
      padding: 14px 16px;
      text-align: left;
      font-weight: 600;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: var(--gray-500);
      border-bottom: 2px solid var(--gray-200);
    }
    
    .corrections-table td {
      padding: 16px;
      vertical-align: top;
      border-bottom: 1px solid var(--gray-100);
    }
    
    .corrections-table tr:hover { background: var(--gray-50); }
    
    .correction-number { font-size: 11px; color: var(--gray-400); margin-bottom: 6px; font-weight: 500; }
    
    .code-block {
      padding: 12px 14px;
      border-radius: 8px;
      font-family: 'Monaco', 'Consolas', 'Courier New', monospace;
      font-size: 13px;
      white-space: pre-wrap;
      word-break: break-word;
      line-height: 1.5;
    }
    
    .error-block { background: linear-gradient(135deg, #fef2f2 0%, #fee2e2 100%); border: 1px solid #fecaca; }
    .error-text { color: #991b1b; text-decoration: line-through; text-decoration-color: #ef4444; text-decoration-thickness: 2px; }
    .success-block { background: linear-gradient(135deg, #f0fdf4 0%, #dcfce7 100%); border: 1px solid #bbf7d0; margin-bottom: 10px; }
    .corrected-text { color: #166534; font-weight: 600; }
    
    .context-text {
      font-size: 11px;
      color: var(--gray-500);
      margin-top: 8px;
      font-style: italic;
      padding: 6px 10px;
      background: var(--gray-100);
      border-radius: 4px;
    }
    
    .type-badge {
      display: inline-flex;
      align-items: center;
      gap: 4px;
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 600;
    }
    
    .type-fragment { background: linear-gradient(135deg, #fef3c7 0%, #fde68a 100%); color: #92400e; border: 1px solid #fcd34d; }
    .type-spelling { background: linear-gradient(135deg, #dbeafe 0%, #bfdbfe 100%); color: #1e40af; border: 1px solid #93c5fd; }
    .type-grammar { background: linear-gradient(135deg, #f3e8ff 0%, #e9d5ff 100%); color: #7c3aed; border: 1px solid #c4b5fd; }
    .type-punctuation { background: linear-gradient(135deg, #f3f4f6 0%, #e5e7eb 100%); color: #374151; border: 1px solid #d1d5db; }
    .type-ai { background: linear-gradient(135deg, #dcfce7 0%, #bbf7d0 100%); color: #166534; border: 1px solid #86efac; }
    
    .footer { text-align: center; padding: 32px 16px; color: var(--gray-500); font-size: 13px; }
    .footer a { color: var(--primary); text-decoration: none; font-weight: 500; }
    .footer a:hover { text-decoration: underline; }
    
    .empty-state { text-align: center; padding: 40px 20px; color: var(--gray-500); }
    .empty-state-icon { font-size: 48px; margin-bottom: 12px; }
    .empty-state p { font-size: 14px; }
    
    .corrections-summary { display: flex; align-items: center; gap: 8px; margin-left: auto; font-size: 13px; color: var(--gray-600); }
    .corrections-count { background: var(--success); color: white; padding: 2px 10px; border-radius: 12px; font-weight: 600; font-size: 12px; }
    
    @media (max-width: 640px) {
      body { padding: 16px 12px; }
      .stats-grid { grid-template-columns: 1fr; }
      .score-stats { flex-direction: column; gap: 16px; }
      .file-badge { flex-direction: column; gap: 8px; }
      .header-card { padding: 24px 16px; }
      .score-card { padding: 24px 16px; }
      .corrections-table th, .corrections-table td { padding: 12px 10px; }
      .code-block { font-size: 12px; padding: 10px; }
    }
    
    @media print {
      body { background: white; padding: 0; }
      .container { max-width: 100%; }
      .card { box-shadow: none; border: 1px solid var(--gray-200); break-inside: avoid; }
      .header-card { box-shadow: none; }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header-card">
      <div class="logo">🛡️</div>
      <h1>Qualion-Doc Security Report</h1>
      <p class="subtitle">Professional Document Sanitization</p>
      <div class="file-badge">
        <span><strong>File:</strong> <code>${escapeHtml(filename)}</code></span>
        <span><strong>Type:</strong> <code>.${(ext || 'unknown').toUpperCase()}</code></span>
      </div>
    </div>
    
    <div class="score-card">
      <div class="score-circle">
        <span class="score-number">${Math.round(score)}</span>
      </div>
      <div class="score-label">${scoreStyle.label}</div>
      <p class="score-description">Security Score after cleaning process</p>
      
      <div class="score-stats">
        <div class="score-stat">
          <div class="score-stat-number">${totalIssues}</div>
          <div class="score-stat-label">Issues Found</div>
        </div>
        <div class="score-stat">
          <div class="score-stat-number">${criticalIssues}</div>
          <div class="score-stat-label">Critical</div>
        </div>
        <div class="score-stat">
          <div class="score-stat-number">${totalCorrections}</div>
          <div class="score-stat-label">Corrections</div>
        </div>
      </div>
    </div>
    
    <div class="card">
      <div class="card-header">
        <div class="card-icon blue">🧹</div>
        <h2>Cleaning Summary</h2>
      </div>
      <div class="card-body">
        <div class="stats-grid">
          ${getCleaningList()}
        </div>
      </div>
    </div>
    
    ${recommendations.length > 0 ? `
    <div class="card">
      <div class="card-header">
        <div class="card-icon orange">💡</div>
        <h2>Recommendations</h2>
      </div>
      <div class="card-body">
        <ul class="recommendations-list">
          ${recommendationsList}
        </ul>
      </div>
    </div>
    ` : ''}
    
    ${allCorrections.length > 0 ? `
    <div class="card">
      <div class="card-header">
        <div class="card-icon green">✏️</div>
        <h2>Text Corrections Applied</h2>
        <div class="corrections-summary">
          <span class="corrections-count">${allCorrections.length}</span>
          <span>corrections</span>
        </div>
      </div>
      <div class="card-body" style="padding: 0;">
        <table class="corrections-table">
          <thead>
            <tr>
              <th style="width: 50%;">Original (Error)</th>
              <th style="width: 50%;">Corrected</th>
            </tr>
          </thead>
          <tbody>
            ${correctionRows}
          </tbody>
        </table>
      </div>
    </div>
    ` : `
    <div class="card">
      <div class="card-header">
        <div class="card-icon green">✏️</div>
        <h2>Text Corrections</h2>
      </div>
      <div class="card-body">
        <div class="empty-state">
          <div class="empty-state-icon">✅</div>
          <p>No text corrections were needed.</p>
          <p style="margin-top: 8px; font-size: 12px;">The document text appears to be error-free.</p>
        </div>
      </div>
    </div>
    `}
    
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

function escapeHtml(s = '') {
  if (s === null || s === undefined) return '';
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export default { buildReportHtmlDetailed, buildReportData };
