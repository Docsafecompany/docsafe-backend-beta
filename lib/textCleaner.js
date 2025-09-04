// Normalisation basique: espaces, ponctuation doublée, guillemets, etc.
const MULTI_PUNCT = /([,;:!?@.])\1+/g; // ",", ";;", "::", "!!", "??", "@@", ".."
const MULTI_SPACE = /\u00A0|\s{2,}/g;   // doubles espaces + nbsp simplifiés
const SPACE_BEFORE_PUNCT = /\s+([,;:!?@.])/g; // pas d’espace avant ces ponctuations
const SPACE_AFTER_COMMA = /,([^\s])/g;       // espace après virgule

export function normalizeText(str) {
  if (!str) return str;
  let out = String(str);
  out = out.replace(MULTI_PUNCT, '$1');
  out = out.replace(MULTI_SPACE, ' ');
  out = out.replace(SPACE_BEFORE_PUNCT, '$1');
  out = out.replace(SPACE_AFTER_COMMA, ', $1');
  // supprime espaces en fin de lignes
  out = out.split('\n').map(l => l.trimEnd()).join('\n');
  return out.trim();
}
