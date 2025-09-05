import sanitizeHtml from 'sanitize-html';

export function generateReportHTML({ filename, mime, baseStats, lt, ai }) {
  const ltIssues = lt?.matches?.length || 0;
  const badges = [
    ai?.mode === 'v2' ? 'V2: proofread + rephrase' : 'V1: proofread',
    ai?.proofed ? 'AI proofread ✅' : 'AI proofread ⚠️ (fallback)',
    ai?.mode === 'v2' ? (ai?.rephrased ? 'AI rephrase ✅' : 'AI rephrase ⚠️ (fallback)') : null
  ].filter(Boolean);

  return `<!doctype html>
<html lang="en"><head>
<meta charset="utf-8"/><meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>DocSafe Report</title>
<style>
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Ubuntu,Arial,sans-serif;padding:24px;}
h1{font-size:22px;margin:0 0 16px}
.card{border:1px solid #e5e7eb;border-radius:12px;padding:16px;margin:0 0 12px}
.tag{display:inline-block;background:#f3f4f6;border:1px solid #e5e7eb;border-radius:999px;padding:4px 10px;font-size:12px;margin-right:8px}
.mono{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;color:#111827}
</style>
</head><body>
<h1>DocSafe Report</h1>
<div class="card">
  <div>File: <span class="mono">${sanitizeHtml(filename || '')}</span></div>
  <div>Type: <span class="mono">${sanitizeHtml(mime || '')}</span></div>
  <div>Chars after normalize: <span class="mono">${baseStats?.length ?? 0}</span></div>
  <div>LanguageTool issues: <span class="mono">${ltIssues}</span></div>
  <div style="margin-top:8px;">${badges.map(b=>`<span class="tag">${b}</span>`).join('')}</div>
</div>
<div class="card">
  <div>Notes:</div>
  <ul>
    <li>Spaces & punctuation normalized (no ",," ";;" "::" "!!" "??" "@@" or double spaces).</li>
    <li>Metadata scrubbed; Strict PDF removes invisible text when possible.</li>
    <li>Outputs:
      <ul>
        <li><b>cleaned-binary</b> = original format cleaned</li>
        <li><b>cleaned.docx</b> = AI-corrected (or normalized fallback)</li>
        ${ai?.mode === 'v2' ? '<li><b>rephrased.docx</b> = AI-rephrased (or corrected fallback)</li>' : ''}
      </ul>
    </li>
  </ul>
</div>
</body></html>`;
}
