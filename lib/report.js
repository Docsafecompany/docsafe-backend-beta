import sanitizeHtml from 'sanitize-html';

export function generateReportHTML({ filename, mime, baseStats, lt, ai }) {
  const ltIssues = lt?.matches?.length || 0;
  const aiMode = ai?.mode === 'v2' ? 'V2 (proofread + rephrase)' : 'V1 (proofread only)';

  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" />
<title>DocSafe Report</title>
<style>
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Ubuntu,Arial,sans-serif;padding:24px;}
  h1{font-size:22px;margin:0 0 16px}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
  .card{border:1px solid #e5e7eb;border-radius:12px;padding:16px}
  .muted{color:#6b7280}
  table{width:100%;border-collapse:collapse}
  td,th{padding:8px;border-bottom:1px solid #eee;text-align:left}
  .tag{display:inline-block;background:#f3f4f6;border:1px solid #e5e7eb;border-radius:999px;padding:4px 10px;font-size:12px;margin-right:8px}
</style>
</head>
<body>
  <h1>DocSafe Report</h1>
  <p class="muted">File: <strong>${sanitizeHtml(filename || '')}</strong> — <span>${sanitizeHtml(mime || '')}</span></p>
  <div class="grid">
    <div class="card">
      <h2>Summary</h2>
      <table>
        <tr><th>Characters</th><td>${baseStats?.length ?? 0}</td></tr>
        <tr><th>LanguageTool Issues</th><td>${ltIssues}</td></tr>
        <tr><th>AI Mode</th><td>${aiMode}</td></tr>
      </table>
    </div>
    <div class="card">
      <h2>Notes</h2>
      <p class="muted">• Spaces & punctuation normalized (no ",," ";;" "::" "!!" "??" "@@" double spaces).<br/>
      • Metadata scrubbed.<br/>
      • Strict PDF removes invisible/white text when possible (heuristic).</p>
    </div>
  </div>
</body>
</html>`;
}
