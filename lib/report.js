// lib/report.js
// Génère un rapport HTML professionnel pour Qualion-Doc

/**
 * Génère un rapport HTML détaillé et stylé
 * @param {Object} params - Paramètres du rapport
 * @returns {string} - HTML complet du rapport
 */
export function buildReportHtmlDetailed({ filename, ext, policy, cleaning, correction, analysis, spellingErrors = [] }) {
  const fmt = (n) => typeof n === 'number' ? n : 0;
  
  // Récupérer le score depuis analysis - CORRIGÉ pour éviter NaN
  const score = typeof analysis?.summary?.riskScore === 'number' ? analysis.summary.riskScore : 100;
  const totalIssues = analysis?.summary?.totalIssues ?? 0;
  const criticalIssues = analysis?.summary?.criticalIssues ?? analysis?.summary?.critical ?? 0;
  const recommendations = analysis?.summary?.recommendations || [];
  
  // Couleurs selon le score (100 = vert/safe, 0 = rouge/critical)
  const getScoreColor = (s) => {
    if (s >= 90) return '#22c55e'; // Green
    if (s >= 70) return '#84cc16'; // Lime
    if (s >= 50) return '#eab308'; // Yellow
    if (s >= 25) return '#f97316'; // Orange
    return '#ef4444'; // Red
  };
  
  const getScoreLabel = (s) => {
    if (s >= 90) return 'Safe';
    if (s >= 70) return 'Low Risk';
    if (s >= 50) return 'Medium Risk';
    if (s >= 25) return 'High Risk';
    return 'Critical Risk';
  };
  
  const scoreColor = getScoreColor(score);
  const scoreLabel = getScoreLabel(score);
  
  // ============================================================
  // COMBINER toutes les corrections (spellingErrors + correction.examples)
  // ============================================================
  const allCorrections = [];
  
  // 1. Ajouter les spellingErrors (format: {error, correction, type, context})
  if (spellingErrors && spellingErrors.length > 0) {
    spellingErrors.forEach(err => {
      const before = err.error || err.word || err.original || '';
      const after = err.correction || err.suggestion || err.corrected || '';
      if (before && after && before !== after) {
        allCorrections.push({
          before,
          after,
          type: err.type || 'spelling',
          context: err.context || ''
        });
      }
    });
  }
  
  // 2. Ajouter les correction.examples (format: {before, after})
  if (correction?.examples && correction.examples.length > 0) {
    correction.examples.forEach(ex => {
      const before = ex.before || '';
      const after = ex.after || '';
      if (before && after && before !== after) {
        // Éviter les doublons
        const exists = allCorrections.some(c => c.before === before && c.after === after);
        if (!exists) {
          allCorrections.push({
            before,
            after,
            type: ex.type || 'ai_correction',
            context: ''
          });
        }
      }
    });
  }
  
  // Générer les lignes de corrections avec design amélioré
  const correctionRows = allCorrections.map(corr => {
    const typeLabel = {
      'fragmented_word': '🔗 Fragment',
      'spelling': '✏️ Spelling',
      'grammar': '📝 Grammar',
      'punctuation': '⚫ Punctuation',
      'ai_correction': '🤖 AI Fix'
    }[corr.type] || '✏️ Fix';
    
    const typeClass = {
      'fragmented_word': 'type-fragment',
      'spelling': 'type-spelling',
      'grammar': 'type-grammar',
      'punctuation': 'type-punctuation',
      'ai_correction': 'type-ai'
    }[corr.type] || 'type-spelling';
    
    return `<tr>
      <td class="cell-before">
        <div class="code-block">
          <span class="error-text">${escapeHtml(corr.before)}</span>
        </div>
        ${corr.context ? `<div class="context-text">"...${escapeHtml(corr.context)}..."</div>` : ''}
      </td>
      <td class="cell-after">
        <div class="code-block">
          <span class="corrected-text">${escapeHtml(corr.after)}</span>
        </div>
        <span class="type-badge ${typeClass}">${typeLabel}</span>
      </td>
    </tr>`;
  }).join('');
  
  const totalCorrections = allCorrections.length || fmt(correction?.changedTextNodes) || 0;
  
  // Générer les recommandations
  const recommendationsList = recommendations.map(rec => 
    `<li>${escapeHtml(rec)}</li>`
  ).join('');
  
  // Construire la liste de nettoyage selon le type de fichier
  const getCleaningList = () => {
    if (ext === 'docx') {
      return `
        <div class="stat-item">
          <span class="stat-label">Metadata removed</span>
          <span class="stat-value">${fmt(cleaning.metaRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Comments removed</span>
          <span class="stat-value">${fmt(cleaning.commentsXmlRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Comment markers cleaned</span>
          <span class="stat-value">${fmt(cleaning.commentMarkersRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Deletions removed</span>
          <span class="stat-value">${fmt(cleaning.revisionsAccepted?.deletionsRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Insertions unwrapped</span>
          <span class="stat-value">${fmt(cleaning.revisionsAccepted?.insertionsUnwrapped)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Ink annotations removed</span>
          <span class="stat-value">${fmt(cleaning.inkRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">VML shapes removed</span>
          <span class="stat-value">${fmt(cleaning.vmlRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Media files deleted</span>
          <span class="stat-value">${fmt(cleaning.mediaDeleted)}</span>
        </div>
      `;
    }
    if (ext === 'pptx') {
      return `
        <div class="stat-item">
          <span class="stat-label">Metadata removed</span>
          <span class="stat-value">${fmt(cleaning.metaRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Comments removed</span>
          <span class="stat-value">${fmt(cleaning.commentsXmlRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Comment relations cleaned</span>
          <span class="stat-value">${fmt(cleaning.relsCommentsRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Ink removed</span>
          <span class="stat-value">${fmt(cleaning.inkRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Pictures removed</span>
          <span class="stat-value">${fmt(cleaning.picturesRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Media deleted</span>
          <span class="stat-value">${fmt(cleaning.mediaDeleted)}</span>
        </div>
      `;
    }
    if (ext === 'pdf') {
      return `
        <div class="stat-item">
          <span class="stat-label">Metadata cleared</span>
          <span class="stat-value">${cleaning.metadataCleared ? '✓' : '—'}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Annotations removed</span>
          <span class="stat-value">${fmt(cleaning.annotsRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Embedded files removed</span>
          <span class="stat-value">${fmt(cleaning.embeddedFilesRemoved)}</span>
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
          <span class="stat-value">${fmt(cleaning.metaRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Comments removed</span>
          <span class="stat-value">${fmt(cleaning.commentsRemoved)}</span>
        </div>
        <div class="stat-item">
          <span class="stat-label">Hidden sheets handled</span>
          <span class="stat-value">${fmt(cleaning.hiddenSheetsRemoved)}</span>
        </div>
      `;
    }
    return '<div class="stat-item"><span class="stat-label">No specific cleaning for this format</span></div>';
  };

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
    
    /* Header Card */
    .header-card {
      background: linear-gradient(135deg, var(--primary) 0%, #8b5cf6 100%);
      border-radius: 20px;
      padding: 32px;
      color: white;
      text-align: center;
      margin-bottom: 24px;
      box-shadow: 0 10px 40px rgba(59, 130, 246, 0.3);
    }
    
    .logo {
      font-size: 32px;
      margin-bottom: 8px;
    }
    
    .header-card h1 {
      font-size: 24px;
      font-weight: 700;
      margin-bottom: 4px;
    }
    
    .header-card .subtitle {
      opacity: 0.9;
      font-size: 14px;
      margin-bottom: 20px;
    }
    
    .file-badge {
      display: inline-flex;
      align-items: center;
      gap: 12px;
      background: rgba(255, 255, 255, 0.15);
      backdrop-filter: blur(10px);
      padding: 12px 24px;
      border-radius: 12px;
      font-size: 14px;
    }
    
    .file-badge code {
      background: rgba(255, 255, 255, 0.2);
      padding: 4px 10px;
      border-radius: 6px;
      font-family: 'Monaco', 'Consolas', monospace;
      font-size: 13px;
    }
    
    /* Score Card */
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
      background: conic-gradient(${scoreColor} ${score}%, var(--gray-200) ${score}%);
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0 auto 20px;
      position: relative;
      box-shadow: 0 8px 30px ${scoreColor}40;
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
      color: ${scoreColor};
    }
    
    .score-label {
      display: inline-block;
      padding: 8px 20px;
      background: ${scoreColor}15;
      color: ${scoreColor};
      border-radius: 24px;
      font-weight: 600;
      font-size: 14px;
      margin-bottom: 12px;
    }
    
    .score-description {
      color: var(--gray-500);
      font-size: 14px;
    }
    
    .score-stats {
      display: flex;
      justify-content: center;
      gap: 32px;
      margin-top: 24px;
      padding-top: 24px;
      border-top: 1px solid var(--gray-200);
    }
    
    .score-stat {
      text-align: center;
    }
    
    .score-stat-number {
      font-size: 28px;
      font-weight: 700;
      color: var(--gray-900);
    }
    
    .score-stat-label {
      font-size: 12px;
      color: var(--gray-500);
      text-transform: uppercase;
      letter-spacing: 0.5px;
    }
    
    /* Content Cards */
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
    
    .card-header h2 {
      font-size: 16px;
      font-weight: 600;
      color: var(--gray-900);
    }
    
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
    
    .card-body {
      padding: 20px 24px;
    }
    
    /* Stats Grid */
    .stats-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 12px;
    }
    
    .stat-item {
      display: flex;
      justify-content: space-between;
      align-items: center;
      padding: 12px 16px;
      background: var(--gray-50);
      border-radius: 10px;
    }
    
    .stat-label {
      color: var(--gray-700);
      font-size: 13px;
    }
    
    .stat-value {
      font-weight: 600;
      color: var(--gray-900);
      background: white;
      padding: 4px 12px;
      border-radius: 6px;
      font-size: 13px;
    }
    
    /* Recommendations */
    .recommendations-list {
      list-style: none;
    }
    
    .recommendations-list li {
      padding: 12px 16px;
      background: var(--gray-50);
      border-radius: 10px;
      margin-bottom: 8px;
      font-size: 14px;
      color: var(--gray-700);
      border-left: 3px solid var(--primary);
    }
    
    .recommendations-list li:last-child {
      margin-bottom: 0;
    }
    
    /* Corrections Table - AMÉLIORÉ */
    .corrections-table {
      width: 100%;
      border-collapse: collapse;
    }
    
    .corrections-table th {
      background: var(--gray-50);
      padding: 14px 16px;
      text-align: left;
      font-weight: 600;
      font-size: 12px;
      text-transform: uppercase;
      letter-spacing: 0.5px;
      color: var(--gray-500);
    }
    
    .corrections-table td {
      padding: 16px;
      vertical-align: top;
      border-bottom: 1px solid var(--gray-100);
    }
    
    .cell-before .code-block {
      background: #fef2f2;
      color: #991b1b;
      padding: 12px;
      border-radius: 8px;
      font-family: 'Monaco', 'Consolas', monospace;
      font-size: 13px;
      white-space: pre-wrap;
      word-break: break-word;
    }
    
    .cell-before .error-text {
      text-decoration: line-through;
      text-decoration-color: #ef4444;
    }
    
    .cell-after .code-block {
      background: #f0fdf4;
      color: #166534;
      padding: 12px;
      border-radius: 8px;
      font-family: 'Monaco', 'Consolas', monospace;
      font-size: 13px;
      white-space: pre-wrap;
      word-break: break-word;
      margin-bottom: 8px;
    }
    
    .cell-after .corrected-text {
      font-weight: 600;
    }
    
    .context-text {
      font-size: 11px;
      color: var(--gray-500);
      margin-top: 6px;
      font-style: italic;
    }
    
    .type-badge {
      display: inline-block;
      padding: 4px 10px;
      border-radius: 12px;
      font-size: 11px;
      font-weight: 500;
    }
    
    .type-fragment { background: #fef3c7; color: #92400e; }
    .type-spelling { background: #dbeafe; color: #1e40af; }
    .type-grammar { background: #f3e8ff; color: #7c3aed; }
    .type-punctuation { background: #e5e7eb; color: #374151; }
    .type-ai { background: #dcfce7; color: #166534; }
    
    /* Footer */
    .footer {
      text-align: center;
      padding: 32px 16px;
      color: var(--gray-500);
      font-size: 13px;
    }
    
    .footer a {
      color: var(--primary);
      text-decoration: none;
      font-weight: 500;
    }
    
    .footer a:hover {
      text-decoration: underline;
    }
    
    /* Empty State */
    .empty-state {
      text-align: center;
      padding: 32px;
      color: var(--gray-500);
    }
    
    .empty-state-icon {
      font-size: 48px;
      margin-bottom: 12px;
    }
    
    /* Responsive */
    @media (max-width: 640px) {
      .stats-grid {
        grid-template-columns: 1fr;
      }
      
      .score-stats {
        flex-direction: column;
        gap: 16px;
      }
      
      .file-badge {
        flex-direction: column;
        gap: 8px;
      }
    }
  </style>
</head>
<body>
  <div class="container">
    <!-- Header -->
    <div class="header-card">
      <div class="logo">🛡️</div>
      <h1>Qualion-Doc Security Report</h1>
      <p class="subtitle">Professional Document Sanitization</p>
      <div class="file-badge">
        <span><strong>File:</strong> <code>${escapeHtml(filename)}</code></span>
        <span><strong>Type:</strong> <code>.${ext.toUpperCase()}</code></span>
      </div>
    </div>
    
    <!-- Score Card -->
    <div class="score-card">
      <div class="score-circle">
        <span class="score-number">${score}</span>
      </div>
      <div class="score-label">${scoreLabel}</div>
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
    
    <!-- Cleaning Summary -->
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
    
    <!-- Recommendations -->
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
    
    <!-- Text Corrections -->
    ${allCorrections.length > 0 ? `
    <div class="card">
      <div class="card-header">
        <div class="card-icon green">✏️</div>
        <h2>Text Corrections Applied (${allCorrections.length})</h2>
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
          <p>No text corrections were needed or applied.</p>
        </div>
      </div>
    </div>
    `}
    
    <!-- Footer -->
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

/**
 * Échapper les caractères HTML
 */
function escapeHtml(s = '') {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export default { buildReportHtmlDetailed };
