// lib/applySpellingFixes.js
// Applique une liste de corrections { error, correction } dans des strings
// Safe: remplace seulement les occurrences exactes, en limitant le nombre de remplacements par item.

export function normalizeFixes(fixes) {
  if (!Array.isArray(fixes)) return [];
  return fixes
    .map(f => ({
      error: String(f?.error || "").trim(),
      correction: String(f?.correction || "").trim(),
      type: f?.type || "spelling",
      severity: f?.severity || "medium",
      message: f?.message || ""
    }))
    .filter(f => f.error && f.correction && f.error.toLowerCase() !== f.correction.toLowerCase())
    // éviter les spans énormes (sécurité)
    .filter(f => f.error.length <= 200 && f.correction.length <= 200);
}

// Remplace les occurrences exactes, au plus `maxPerFix` fois par correction
export function applyFixesToText(text, fixes, maxPerFix = 50) {
  if (!text || !fixes?.length) return { text: text || "", stats: { replacements: 0, perFix: [] } };

  let out = text;
  let total = 0;
  const perFix = [];

  for (const f of fixes) {
    let count = 0;

    // remplacement exact (case-sensitive)
    while (count < maxPerFix) {
      const idx = out.indexOf(f.error);
      if (idx === -1) break;

      out = out.slice(0, idx) + f.correction + out.slice(idx + f.error.length);
      count++;
      total++;
    }

    if (count > 0) perFix.push({ error: f.error, correction: f.correction, count });
  }

  return { text: out, stats: { replacements: total, perFix } };
}
