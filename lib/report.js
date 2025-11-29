// lib/report.js
// Génère le rapport HTML détaillé

export function buildReportHtmlDetailed({ filename, ext, policy, cleaning, correction, detections, summary }) {
  const fmt = (n) => typeof n === 'number' ? n : 0;
  
  // Generate text corrections table
  const correctionRows = (correction?.examples || []).map(({ before, after, error, correction: corr }) =>
    `<tr>
      <td><code class="error">${escapeHtml(error || '')}</code></td>
      <td><code class="correction">${escapeHtml(corr || '')}</code></td>
      <td><pre class="context">${escapeHtml(before || '')}</pre></td>
      <td><pre class="context">${escapeHtml(after || '')}</pre></td>
    </tr>`
  ).join('');

  // Generate spelling errors section (from analysis)
  const spellingSection = detections?.spellingErrors?.length > 0 ? `
    <h2>📝 Spelling & Grammar Errors Detected</h2>
    <table>
      <thead>
        <tr>
          <th>Error</th>
          <th>Suggested Correction</th>
          <th>Context</th>
          <th>Severity</th>
        </tr>
      </thead>
      <tbody>
        ${detections.spellingErrors.map(err => `
          <tr class="severity-${err.severity}">
            <td><code class="error">${escapeHtml(err.error)}</code></td>
            <td><code class="correction">${escapeHtml(err.correction)}</code></td>
            <td><span class="context">${escapeHtml(err.context)}</span></td>
            <td><badge class="badge-${err.severity}">${err.severity}</badge></td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  ` : '';

  // Generate sensitive data section
  const sensitiveSection = detections?.sensitiveData?.length > 0 ? `
    <h2>⚠️ Sensitive Data Detected</h2>
    <table>
      <thead>
        <tr>
          <th>Type</th>
          <th>Value</th>
          <th>Category</th>
          <th>Severity</th>
        </tr>
      </thead>
      <tbody>
        ${detections.sensitiveData.map(data => `
          <tr class="severity-${data.severity}">
            <td>${escapeHtml(data.type)}</td>
            <td><code>${escapeHtml(data.value)}</code></td>
            <td>${escapeHtml(data.category || 'N/A')}</td>
            <td><badge class="badge-${data.severity}">${data.severity}</badge></td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  ` : '';

  // Generate risk score section
  const riskColor = summary?.riskLevel === 'critical' ? '#dc2626' :
                    summary?.riskLevel === 'high' ? '#ea580c' :
                    summary?.riskLevel === 'medium' ? '#ca8a04' : '#16a34a';

  const riskSection = summary ? `
    <div class="risk-summary">
      <div class="risk-score" style="border-color: ${riskColor}; color: ${riskColor}">
        <span class="score">${summary.riskScore}</span>
        <span class="label">Risk Score</span>
      </div>
      <div class="risk-details">
        <p><strong>Risk Level:</strong> <badge class="badge-${summary.riskLevel}">${summary.riskLevel.toUpperCase()}</badge></p>
        <p><strong>Total Issues:</strong> ${summary.totalIssues}</p>
        <p><strong>Critical Issues:</strong> ${summary.criticalIssues}</p>
      </div>
    </div>
    ${summary.recommendations?.length > 0 ? `
      <h3>📋 Recommendations</h3>
      <ul class="recommendations">
        ${summary.recommendations.map(rec => `<li>${escapeHtml(rec)}</li>`).join('')}
      </ul>
    ` : ''}
  ` : '';

  // Cleaning summary based on file type
  const cleaningList = (() => {
    if (ext === 'docx') {
      return `
        <li><b>Metadata parts removed:</b> ${fmt(cleaning.metaRemoved)}</li>
        <li><b>Comments parts removed:</b> ${fmt(cleaning.commentsXmlRemoved)}</li>
        <li><b>In-text comment markers removed:</b> ${fmt(cleaning.commentMarkersRemoved)}</li>
        <li><b>Tracked changes:</b> deletions removed ${fmt(cleaning.revisionsAccepted?.deletionsRemoved)}, insertions unwrapped ${fmt(cleaning.revisionsAccepted?.insertionsUnwrapped)}</li>
        <li><b>Ink removed:</b> ${fmt(cleaning.inkRemoved)}</li>
        <li><b>VML shapes removed:</b> ${fmt(cleaning.vmlRemoved)}</li>
        <li><b>Drawings removed:</b> ${fmt(cleaning.drawingsRemoved)} (policy: ${policy?.drawPolicy || 'default'})</li>
        <li><b>Media files deleted:</b> ${fmt(cleaning.mediaDeleted)}</li>
      `;
    }
    if (ext === 'pptx') {
      return `
        <li><b>Metadata parts removed:</b> ${fmt(cleaning.metaRemoved)}</li>
        <li><b>Comment threads/authors removed:</b> ${fmt(cleaning.commentsXmlRemoved)}</li>
        <li><b>Slide rels to comments removed:</b> ${fmt(cleaning.relsCommentsRemoved)}</li>
        <li><b>Ink removed:</b> ${fmt(cleaning.inkRemoved)}</li>
        <li><b>Pictures removed:</b> ${fmt(cleaning.picturesRemoved)} (policy: ${policy?.drawPolicy || 'default'})</li>
        <li><b>Media files deleted:</b> ${fmt(cleaning.mediaDeleted)}</li>
      `;
    }
    if (ext === 'pdf') {
      return `
        <li><b>Metadata cleared:</b> ${cleaning.metadataCleared ? 'yes' : 'no'}</li>
        <li><b>Annotations removed:</b> ${fmt(cleaning.annotsRemoved)}</li>
        <li><b>Embedded files removed:</b> ${fmt(cleaning.embeddedFilesRemoved)}</li>
        <li><b>Mode:</b> ${policy?.pdfMode || 'standard'}${cleaning.textOnly ? ' (text-only PDF built)' : ''}</li>
      `;
    }
    if (ext === 'xlsx') {
      return `
        <li><b>Metadata parts removed:</b> ${fmt(cleaning.metaRemoved)}</li>
        <li><b>Comments removed:</b> ${fmt(cleaning.commentsRemoved)}</li>
        <li><b>Hidden sheets removed:</b> ${fmt(cleaning.hiddenSheetsRemoved)}</li>
        <li><b>VBA macros removed:</b> ${cleaning.macrosRemoved ? 'yes' : 'no'}</li>
      `;
    }
    return `<li>No cleaning performed for .${ext}</li>`;
  })();

  return `<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>Qualion-Doc Report — ${escapeHtml(filename)}</title>
  <style>
    * { box-sizing: border-box; }
    body {
      font-family: system-ui, -apple-system, 'Segoe UI', Roboto, Arial, sans-serif;
      margin: 0;
      padding: 24px;
      line-height: 1.5;
      color: #1f2937;
      background: #f9fafb;
    }
    .container {
      max-width: 900px;
      margin: 0 auto;
      background: white;
      padding: 32px;
      border-radius: 12px;
      box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    h1 {
      font-size: 24px;
      margin: 0 0 8px;
      color: #111827;
      display: flex;
      align-items: center;
      gap: 12px;
    }
    h1 img {
      height: 32px;
    }
    h2 {
      font-size: 18px;
      margin: 28px 0 12px;
      color: #374151;
      border-bottom: 2px solid #e5e7eb;
      padding-bottom: 8px;
    }
    h3 {
      font-size: 16px;
      margin: 20px 0 10px;
      color: #4b5563;
    }
    
    .file-info {
      background: #f3f4f6;
      padding: 12px 16px;
      border-radius: 8px;
      margin-bottom: 24px;
    }
    .file-info code {
      background: #e5e7eb;
      padding: 2px 8px;
      border-radius: 4px;
      font-size: 14px;
    }
    
    .risk-summary {
      display: flex;
      gap: 24px;
      align-items: center;
      padding: 20px;
      background: #fefefe;
      border: 2px solid #e5e7eb;
      border-radius: 12px;
      margin: 20px 0;
    }
    .risk-score {
      width: 100px;
      height: 100px;
      border: 4px solid;
      border-radius: 50%;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      flex-shrink: 0;
    }
    .risk-score .score {
      font-size: 32px;
      font-weight: bold;
    }
    .risk-score .label {
      font-size: 11px;
      text-transform: uppercase;
    }
    .risk-details p {
      margin: 4px 0;
    }
    
    .recommendations {
      background: #eff6ff;
      padding: 16px 16px 16px 32px;
      border-radius: 8px;
      border-left: 4px solid #3b82f6;
    }
    .recommendations li {
      margin: 6px 0;
    }
    
    table {
      width: 100%;
      border-collapse: collapse;
      margin: 12px 0;
      font-size: 14px;
    }
    th, td {
      border: 1px solid #e5e7eb;
      padding: 10px 12px;
      text-align: left;
      vertical-align: top;
    }
    th {
      background: #f9fafb;
      font-weight: 600;
      color: #374151;
    }
    tr:hover {
      background: #f9fafb;
    }
    
    .severity-high, .severity-critical {
      background: #fef2f2 !important;
    }
    .severity-medium {
      background: #fffbeb !important;
    }
    
    code {
      background: #f3f4f6;
      padding: 2px 6px;
      border-radius: 4px;
      font-family: 'SF Mono', Monaco, Consolas, monospace;
      font-size: 13px;
    }
    code.error {
      background: #fee2e2;
      color: #dc2626;
      text-decoration: line-through;
    }
    code.correction {
      background: #dcfce7;
      color: #16a34a;
    }
    
    .context {
      font-size: 12px;
      color: #6b7280;
      white-space: pre-wrap;
      margin: 0;
      max-width: 300px;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    
    badge {
      display: inline-block;
      padding: 2px 8px;
      border-radius: 9999px;
      font-size: 11px;
      font-weight: 600;
      text-transform: uppercase;
    }
    .badge-critical {
      background: #dc2626;
      color: white;
    }
    .badge-high {
      background: #ea580c;
      color: white;
    }
    .badge-medium {
      background: #ca8a04;
      color: white;
    }
    .badge-low {
      background: #16a34a;
      color: white;
    }
    
    ul {
      margin: 8px 0 0 18px;
      padding: 0;
    }
    li {
      margin: 4px 0;
    }
    
    .footer {
      margin-top: 32px;
      padding-top: 16px;
      border-top: 1px solid #e5e7eb;
      color: #9ca3af;
      font-size: 12px;
      text-align: center;
    }
  </style>
</head>
<body>
  <div class="container">
    <h1>
      <span>🛡️</span>
      Qualion-Doc — Security Report
    </h1>
    
    <div class="file-info">
      <strong>File:</strong> <code>${escapeHtml(filename)}</code> &nbsp;|&nbsp;
      <strong>Type:</strong> <code>.${ext}</code> &nbsp;|&nbsp;
      <strong>Generated:</strong> ${new Date().toLocaleString()}
    </div>

    ${riskSection}

    ${sensitiveSection}
    
    ${spellingSection}

    <h2>🧹 Cleaning Actions Performed</h2>
    <ul>${cleaningList}</ul>

    ${correction?.examples?.length > 0 ? `
      <h2>✏️ Text Corrections Applied</h2>
      <p>Text nodes changed: <strong>${fmt(correction.changedTextNodes)}</strong> / ${fmt(correction.totalTextNodes)}</p>
      <table>
        <thead>
          <tr>
            <th>Original</th>
            <th>Corrected</th>
            <th>Before Context</th>
            <th>After Context</th>
          </tr>
        </thead>
        <tbody>${correctionRows}</tbody>
      </table>
    ` : ''}

    <div class="footer">
      <p>Generated by Qualion-Doc • Mindorion</p>
      <p>This report documents the security analysis and cleaning operations performed on your document.</p>
    </div>
  </div>
</body>
</html>`;
}

function escapeHtml(str = '') {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export default { buildReportHtmlDetailed };
