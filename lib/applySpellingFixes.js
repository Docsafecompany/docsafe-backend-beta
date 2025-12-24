// lib/applySpellingFixes.js
// Applique une liste de corrections { error, correction } dans des strings
// ✅ Fix majeur: NE JAMAIS trim() error/correction (sinon collage de mots)
// ✅ Matching robuste: supporte espaces multiples / retours ligne entre tokens ("comm z unication")
// ✅ Safe boundaries: évite de remplacer au milieu d’un mot
// ✅ Limite de remplacements par fix (maxPerFix)

function escapeRegExp(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Transforme "comm z unication" en regex "comm\s+z\s+unication"
 * et capture le caractère avant (boundary) pour le préserver.
 */
function buildFlexibleWhitespaceRegex(error) {
  const raw = String(error ?? "");
  // Split sur whitespace (inclut \n, \t, espaces multiples)
  const parts = raw.split(/\s+/).filter(Boolean).map(escapeRegExp);
  if (!parts.length) return null;

  const core = parts.join("\\s+");

  // Boundaries: on évite de matcher au milieu d'un mot.
  // - Capture "left boundary" (début de string ou non-\w)
  // - Lookahead pour right boundary (non-\w ou fin)
  return new RegExp(`(^|[^\\w])(${core})(?=[^\\w]|$)`, "g");
}

/**
 * Normalise les fixes sans détruire les espaces.
 * On utilise une version "comparaison" trim/lowercase uniquement pour filtrer,
 * mais on conserve les valeurs RAW pour l’application.
 */
export function normalizeFixes(fixes) {
  if (!Array.isArray(fixes)) return [];

  return fixes
    .map((f) => {
      const rawError = String(f?.error ?? "");
      const rawCorrection = String(f?.correction ?? "");

      const cmpError = rawError.trim().toLowerCase();
      const cmpCorrection = rawCorrection.trim().toLowerCase();

      return {
        error: rawError,                 // ✅ NO TRIM
        correction: rawCorrection,       // ✅ NO TRIM
        _cmpError: cmpError,
        _cmpCorrection: cmpCorrection,
        type: f?.type || "spelling",
        severity: f?.severity || "medium",
        message: f?.message || ""
      };
    })
    .filter((f) => f._cmpError && f._cmpCorrection && f._cmpError !== f._cmpCorrection)
    // éviter spans énormes (sécurité)
    .filter((f) => f.error.length <= 200 && f.correction.length <= 200)
    // cleanup champs internes
    .map(({ _cmpError, _cmpCorrection, ...rest }) => rest);
}

/**
 * Applique les fixes de manière robuste :
 * - match flexible whitespace (ex: "th e", "comm z unication")
 * - boundaries: évite collisions au milieu d’un mot
 * - limite à maxPerFix remplacements par correction
 *
 * NOTE: on préserve le caractère de gauche (boundary) pour éviter de "coller" le texte.
 */
export function applyFixesToText(text, fixes, maxPerFix = 50) {
  if (!text || !fixes?.length) {
    return { text: text || "", stats: { replacements: 0, perFix: [] } };
  }

  let out = String(text);
  let total = 0;
  const perFix = [];

  for (const f of fixes) {
    const error = f?.error ?? "";
    const correction = f?.correction ?? "";

    if (!error || !correction) continue;

    const re = buildFlexibleWhitespaceRegex(error);
    if (!re) continue;

    let count = 0;

    // On remplace via regex et on stoppe au-delà de maxPerFix
    out = out.replace(re, (match, leftBoundary /*, matchedCore */) => {
      if (count >= maxPerFix) return match;
      count++;
      total++;
      return `${leftBoundary}${correction}`;
    });

    if (count > 0) {
      perFix.push({ error, correction, count });
    }
  }

  return { text: out, stats: { replacements: total, perFix } };
}
