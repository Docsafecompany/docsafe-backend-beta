// lib/report.js
export function buildReportHtml({ original, cleaned, rephrased, filename, mode, notes = [] }) {
  const ts = new Date().toISOString();
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>DocSafe Report</title>
<style>
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;margin:24px;line-height:1.55}
  h1{font-size:20px;margin:0 0 8px}
  h2{font-size:16px;margin:16px 0 8px}
  code,pre{background:#f6f7f9;padding:8px;border-radius:6px;display:block;white-space:pre-wrap}
  .grid{display:grid;grid-template-columns:1fr 1fr;gap:16px}
  .meta{color:#555}
  ul{margin-left:18px}
</style>
</head>
<body>
  <h1>DocSafe Report</h1>
  <p class="meta">File: <strong>${escapeHtml(filename)}</strong> • Mode: <strong>${mode}</strong> • ${ts}</p>

  ${notes?.length ? `<h2>Notes</h2><ul>${notes.map(n=>`<li>${escapeHtml(n)}</li>`).join("")}</ul>` : ""}

  <h2>Summary</h2>
  <ul>
    <li>Original length: ${(original||"").length} chars</li>
    <li>${mode==="V2" ? "Rephrased" : "Cleaned"} length: ${(mode==="V2"?(rephrased||""):(cleaned||"")).length} chars</li>
    ${mode==="V2" ? `<li>Cleaned length: ${(cleaned||"").length} chars</li>` : ""}
    <li>Ops: artifact fixes, punctuation spacing, dedup, safe suffix merges, light rephrase (V2).</li>
  </ul>

  <div class="grid">
    <div>
      <h2>Original (excerpt)</h2>
      <pre>${escapeHtml((original||"").slice(0, 2000))}</pre>
    </div>
    <div>
      <h2>${mode==="V2"?"Rephrased":"Cleaned"} (excerpt)</h2>
      <pre>${escapeHtml(((mode==="V2"?rephrased:cleaned)||"").slice(0, 2000))}</pre>
    </div>
  </div>
</body>
</html>`;
}

function escapeHtml(s=""){return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")}
