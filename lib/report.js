// lib/report.js
export function buildReportHtmlDetailed({ filename, ext, policy, cleaning, correction }) {
  const fmt = (n) => typeof n === 'number' ? n : 0;
  const exRows = (correction?.examples || []).map(({before, after}) =>
    `<tr><td><pre>${escapeHtml(before)}</pre></td><td><pre>${escapeHtml(after)}</pre></td></tr>`
  ).join('');

  const cleaningList = (() => {
    if (ext === 'docx') {
      return `
        <li><b>Metadata parts removed:</b> ${fmt(cleaning.metaRemoved)}</li>
        <li><b>Comments parts removed:</b> ${fmt(cleaning.commentsXmlRemoved)}</li>
        <li><b>In-text comment markers removed:</b> ${fmt(cleaning.commentMarkersRemoved)}</li>
        <li><b>Tracked changes:</b> deletions removed ${fmt(cleaning.revisionsAccepted?.deletionsRemoved)}, insertions unwrapped ${fmt(cleaning.revisionsAccepted?.insertionsUnwrapped)}</li>
        <li><b>Ink removed:</b> ${fmt(cleaning.inkRemoved)}</li>
        <li><b>VML shapes removed:</b> ${fmt(cleaning.vmlRemoved)}</li>
        <li><b>Drawings removed:</b> ${fmt(cleaning.drawingsRemoved)} (policy: ${policy.drawPolicy})</li>
        <li><b>Media files deleted:</b> ${fmt(cleaning.mediaDeleted)}</li>
      `;
    }
    if (ext === 'pptx') {
      return `
        <li><b>Metadata parts removed:</b> ${fmt(cleaning.metaRemoved)}</li>
        <li><b>Comment threads/authors removed:</b> ${fmt(cleaning.commentsXmlRemoved)}</li>
        <li><b>Slide rels to comments removed:</b> ${fmt(cleaning.relsCommentsRemoved)}</li>
        <li><b>Ink removed:</b> ${fmt(cleaning.inkRemoved)}</li>
        <li><b>Pictures removed:</b> ${fmt(cleaning.picturesRemoved)} (policy: ${policy.drawPolicy})</li>
        <li><b>Media files deleted:</b> ${fmt(cleaning.mediaDeleted)}</li>
      `;
    }
    if (ext === 'pdf') {
      return `
        <li><b>Metadata cleared:</b> ${cleaning.metadataCleared ? 'yes' : 'no'}</li>
        <li><b>Annotations removed:</b> ${fmt(cleaning.annotsRemoved)}</li>
        <li><b>Embedded files removed:</b> ${fmt(cleaning.embeddedFilesRemoved)}</li>
        <li><b>Mode:</b> ${policy.pdfMode}${cleaning.textOnly ? ' (text-only PDF built)' : ''}</li>
      `;
    }
    return `<li>No cleaning performed for .${ext}</li>`;
  })();

  return `<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<title>DocSafe Report — ${escapeHtml(filename)}</title>
<style>
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial;margin:24px;line-height:1.45;color:#111}
h1{font-size:18px;margin:0 0 8px}
h2{font-size:15px;margin:20px 0 8px}
code,badge{background:#f3f4f6;padding:2px 6px;border-radius:6px}
ul{margin:8px 0 0 18px}
table{width:100%;border-collapse:collapse;margin-top:8px}
td,th{border:1px solid #e5e7eb;padding:8px;vertical-align:top}
pre{white-space:pre-wrap;margin:0}
.small{color:#6b7280;font-size:12px;margin-top:12px}
</style></head><body>
<h1>DocSafe — Detailed Report</h1>
<p><b>File:</b> <code>${escapeHtml(filename)}</code> &nbsp; <b>Type:</b> <code>.${ext}</code></p>

<h2>Cleaning</h2>
<ul>${cleaningList}</ul>

${correction ? `
<h2>Text correction</h2>
<ul>
  <li><b>Text nodes changed:</b> ${fmt(correction.changedTextNodes)} / ${fmt(correction.totalTextNodes)}</li>
</ul>
${correction.examples?.length ? `
<table>
  <thead><tr><th>Before</th><th>After</th></tr></thead>
  <tbody>${exRows}</tbody>
</table>` : `<p class="small">No examples captured (few or no changes).</p>`}
` : ''}

<p class="small">DocSafe removed non-client artifacts and corrected typography/grammar while preserving layout, tables, images and styles.</p>
</body></html>`;
}

function escapeHtml(s='') {
  return s.replace(/&/g,'&amp;').replace(/</g,'&lt;/g'.replace('/g','')) // quick guard
          .replace(/</g,'&lt;').replace(/>/g,'&gt;')
          .replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
