// lib/aiSpellCheck.js
// VERSION 6.0 - AI-first proofreading (context-aware) + regex candidates
// Objectif: détecter orthographe + grammaire + typos + corruption (espaces/ponctuation) avec contexte.
// ✅ Regex = CANDIDATS uniquement (aucune correction automatique)
// ✅ IA tranche et propose corrections (l'utilisateur valide ensuite)
// ✅ Chunking pour récupérer BEAUCOUP plus d'erreurs que 10–20
// ✅ Zéro faux positifs type "a message" -> "amessage"

const STOPWORDS = new Set([
  'a','an','the','to','of','in','on','at','by','for','with','from','into','onto','upon','within','without',
  'and','or','but','so','yet','nor',
  'is','are','was','were','be','been','being',
  'can','could','would','should','will','may','might','must',
  'has','have','had','do','does','did',
  'not','how','one','due','as'
]);

function looksLikeProperName(p1, p2) {
  // "Honesty Walters" => pas un mot cassé, c'est un nom propre
  if (!p1 || !p2) return false;
  return /^[A-Z][a-z]+$/.test(p1) && /^[A-Z][a-z]+$/.test(p2);
}

function isPlausibleWord(w) {
  if (!w) return false;
  if (!/^[A-Za-z]+$/.test(w)) return false;
  if (w.length < 2 || w.length > 40) return false;
  if (/^(.)\1{4,}$/.test(w.toLowerCase())) return false;
  return true;
}

function applyCasingLike(original, corrected) {
  if (!corrected) return corrected;
  const o = (original || '').trim();
  if (!o) return corrected;

  // si l’original commence par Maj -> Capitalize, sinon lower.
  if (/^[A-Z]/.test(o)) return corrected.charAt(0).toUpperCase() + corrected.slice(1);
  return corrected.toLowerCase();
}

/**
 * Découpe le texte en chunks (pour éviter que l'IA sorte 15 erreurs max).
 * 15000 chars -> on split en blocs ~2500 chars
 */
function chunkText(text, chunkSize = 2500) {
  const chunks = [];
  if (!text) return chunks;

  // split par paragraphes pour garder du contexte
  const parts = text.split(/\n{2,}/g);
  let buf = '';

  for (const p of parts) {
    const next = buf ? (buf + '\n\n' + p) : p;
    if (next.length > chunkSize) {
      if (buf) chunks.push(buf);
      if (p.length > chunkSize) {
        // si un paragraphe est énorme, on le coupe brut
        for (let i = 0; i < p.length; i += chunkSize) chunks.push(p.slice(i, i + chunkSize));
        buf = '';
      } else {
        buf = p;
      }
    } else {
      buf = next;
    }
  }
  if (buf) chunks.push(buf);
  return chunks;
}

/**
 * Pré-détecte des "candidats" de corruption à donner à l'IA.
 * IMPORTANT: ne propose PAS de correction ici, juste des spans suspects.
 */
function buildCandidates(text) {
  const candidates = [];
  if (!text) return candidates;

  // 1) multi-spaces dans un mot: "p  otential"
  const multiSpaceWord = /\b([A-Za-z]{1,20})\s{2,}([A-Za-z]{1,20})\b/g;
  let m;
  while ((m = multiSpaceWord.exec(text)) !== null) {
    candidates.push({
      error: m[0],
      hint: 'multiple_spaces_inside_word'
    });
  }

  // 2) espace dans un mot: "soc ial" / "gen uine" / "dar ker" / "commu nication"
  // On évite les stopwords et les noms propres "Honesty Walters"
  const splitWord = /\b([A-Za-z]{2,20})\s+([A-Za-z]{1,20})\b/g;
  while ((m = splitWord.exec(text)) !== null) {
    const p1 = m[1];
    const p2 = m[2];

    if (STOPWORDS.has(p1.toLowerCase()) || STOPWORDS.has(p2.toLowerCase())) continue;
    if (looksLikeProperName(p1, p2)) continue;

    const combined = (p1 + p2).trim();
    if (!isPlausibleWord(combined)) continue;

    // signal: un des fragments court (<=3) OU mots “cassés” typiques
    const strong = (p1.length <= 3) || (p2.length <= 3);
    if (!strong) continue;

    candidates.push({
      error: m[0],
      hint: 'likely_fragmented_word'
    });
  }

  // 3) ponctuation au milieu d’un mot: "corpo,rations"
  const punctInsideWord = /\b([A-Za-z]{2,})[,:;.'-]+([A-Za-z]{2,})\b/g;
  while ((m = punctInsideWord.exec(text)) !== null) {
    candidates.push({
      error: m[0],
      hint: 'punctuation_inside_word'
    });
  }

  // 4) répétitions bizarres: "gggdigital"
  const repeatedPrefix = /\b([a-zA-Z])\1{2,}([A-Za-z]{3,})\b/g;
  while ((m = repeatedPrefix.exec(text)) !== null) {
    candidates.push({
      error: m[0],
      hint: 'repeated_prefix_noise'
    });
  }

  return candidates;
}

// ================= AI PROMPT =================

const PROOFREADER_PROMPT = `
You are an expert English proofreader.

You MUST detect and propose corrections for:
- spelling mistakes
- grammar/conjugation mistakes
- obvious typos
- corrupted words (spaces inside word, multiple spaces, punctuation inside word, repeated junk letters)

Rules:
1) Do NOT rewrite sentences. Only minimal fixes.
2) Use full context to infer intended words (like a human).
3) Return corrections as exact substrings from the input (smallest safe span).
4) If something is ambiguous, still propose the most likely correction, but keep span minimal.

OUTPUT: return ONLY a JSON array:
[
  {
    "error": "exact substring from input",
    "correction": "corrected substring",
    "type": "fragmented_word|multiple_spaces|spelling|grammar|capitalization|punctuation",
    "severity": "high|medium|low",
    "message": "short reason"
  }
]
Return ONLY the JSON array. No extra text.
`;

/**
 * Appel IA sur un chunk avec une liste de candidats suspects.
 * L’IA peut aussi trouver d’autres erreurs de spelling/grammar dans le chunk.
 */
async function aiAnalyzeChunk(chunk, candidates) {
  const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
  if (!OPENAI_API_KEY) return [];

  const payload = {
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: PROOFREADER_PROMPT },
      {
        role: 'user',
        content:
          `TEXT CHUNK:\n<<<\n${chunk}\n>>>\n\n` +
          `SUSPECT CANDIDATES (may be real errors or false alarms):\n` +
          `${JSON.stringify(candidates.slice(0, 80), null, 2)}\n\n` +
          `Task:\n` +
          `1) Validate candidates: keep only real errors and give corrections.\n` +
          `2) Also find other spelling/grammar/typo errors in the chunk.\n` +
          `Return JSON only.`
      }
    ],
    temperature: 0.1,
    max_tokens: 2500
  };

  const response = await fetch('https://api.openai.com/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${OPENAI_API_KEY}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    console.error(`[SPELLCHECK] OpenAI API error: ${response.status}`);
    return [];
  }

  const data = await response.json();
  const content = data.choices?.[0]?.message?.content || '[]';

  try {
    const jsonMatch = content.match(/\[[\s\S]*\]/);
    if (!jsonMatch) return [];
    const arr = JSON.parse(jsonMatch[0]);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    console.error('[SPELLCHECK] Failed to parse AI JSON:', e.message);
    return [];
  }
}

/**
 * Filtre minimal : on enlève uniquement les corrections absurdes.
 * IMPORTANT: on ne “bloque” plus soc ial / gen uine etc.
 */
function sanitizeCorrections(list) {
  const out = [];

  for (const c of (list || [])) {
    const error = (c.error || '').trim();
    const correction = (c.correction || '').trim();
    if (!error || !correction) continue;

    // pas de "no-op"
    if (error.toLowerCase() === correction.toLowerCase()) continue;

    // doit être court-ish (sécurité)
    if (error.length > 120) continue;
    if (correction.length > 120) continue;

    // correction pas vide
    if (correction.length < 1) continue;

    // normalise casing uniquement pour les types de mot cassé
    const type = (c.type || '').trim();
    const fixed = {
      error,
      correction: (type === 'fragmented_word' || type === 'multiple_spaces')
        ? applyCasingLike(error, correction)
        : correction,
      type: type || 'spelling',
      severity: c.severity || 'medium',
      message: c.message || 'AI-detected'
    };

    out.push(fixed);
  }

  // dédup par "error" (case-insensitive)
  const seen = new Set();
  const dedup = [];
  for (const x of out) {
    const k = x.error.toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    dedup.push(x);
  }
  return dedup;
}

// ================= MAIN =================

export async function checkSpellingWithAI(text) {
  console.log('📝 AI spell-check VERSION 6.0 starting...');

  if (!text || text.length < 10) return [];

  const maxChars = 15000;
  const truncatedText = text.length > maxChars ? text.substring(0, maxChars) : text;
  console.log(`[SPELLCHECK] Text length: ${truncatedText.length} characters`);

  // 1) chunks
  const chunks = chunkText(truncatedText, 2500);

  // 2) IA sur chaque chunk
  const all = [];

  for (let i = 0; i < chunks.length; i++) {
    const chunk = chunks[i];

    // candidates local sur le chunk
    const candidates = buildCandidates(chunk);
    console.log(`[SPELLCHECK] Chunk ${i + 1}/${chunks.length} candidates: ${candidates.length}`);

    // IA
    const aiRaw = await aiAnalyzeChunk(chunk, candidates);
    console.log(`[AI RAW] Chunk ${i + 1}: ${aiRaw.length} items`);

    const cleaned = sanitizeCorrections(aiRaw);
    all.push(...cleaned);
  }

  // 3) dédup global
  const seen = new Set();
  const final = [];
  for (const x of all) {
    const k = x.error.toLowerCase();
    if (seen.has(k)) continue;
    seen.add(k);
    final.push(x);
  }

  console.log(`✅ AI spell-check complete: ${final.length} valid issues found`);
  final.slice(0, 25).forEach((issue, i) => {
    console.log(`  ${i + 1}. [${issue.type}] "${issue.error}" → "${issue.correction}" (${issue.message})`);
  });

  return final;
}

export default { checkSpellingWithAI };
