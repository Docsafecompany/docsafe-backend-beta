// lib/aiSpellCheck.js
// VERSION 7.2 - High-recall AI proofreading with overlap + offsets (start/end)
// Goal: detect spelling + grammar + typos + corrupted words with minimal fixes
// ✅ Regex = candidates only
// ✅ AI validates candidates + finds additional issues
// ✅ Overlap chunking to avoid missing boundary errors
// ✅ AI returns offsets (start/end) inside chunk to apply safely downstream
// ✅ PATCH 7.2:
//   - filterInvalidCorrections(): reject merges of valid words + non-real-word outputs
//   - Stronger AI instructions: NEVER merge valid words, only fix truly split words
//   - preDetectFragmentedWords(): stricter merge guardrails (skip if both parts are valid words)

const STOPWORDS = new Set([
  "a","an","the","to","of","in","on","at","by","for","with","from","into","onto","upon","within","without",
  "and","or","but","so","yet","nor",
  "is","are","was","were","be","been","being",
  "can","could","would","should","will","may","might","must",
  "has","have","had","do","does","did",
  "not","how","one","due","as"
]);

/**
 * Known frequent real words (used as an allow-list when we *do* merge).
 * This is NOT meant to be exhaustive; it’s a safety valve for true fragmented words.
 */
const COMMON_WORDS = new Set([
  // core business / compliance / doc words
  "confidential","confidentiality","compliance","documentation","document",
  "enabling","enabled","enablement","alignment","assessment","requirements",
  "deliverable","deliverables","deliver","delivery","commitment","commitments",
  "stakeholder","stakeholders","validation","validated","regulatory",
  "configuration","integration","commissioning","qualification",
  // frequent words
  "therefore","however","because","between","within","without","available",
  "important","information","performance","potential","successful","analysis"
]);

function looksLikeProperName(p1, p2) {
  if (!p1 || !p2) return false;
  return /^[A-Z][a-z]+$/.test(p1) && /^[A-Z][a-z]+$/.test(p2);
}

function isPlausibleWord(w) {
  if (!w) return false;
  if (!/^[A-Za-zÀ-ÿ]+$/.test(w)) return false;
  if (w.length < 2 || w.length > 40) return false;
  if (/^(.)\1{4,}$/.test(w.toLowerCase())) return false; // aaaaaa
  return true;
}

/**
 * Heuristic "real word" gate.
 * If you already have a stronger dictionary-based implementation in your project,
 * you can safely swap this function with yours.
 */
function isRealWord(w) {
  if (!w) return false;
  const s = String(w).trim();
  if (!s) return false;
  // reject multi-word strings here (this helper is for single words)
  if (/\s/.test(s)) return false;
  if (!isPlausibleWord(s)) return false;

  const lower = s.toLowerCase();
  if (STOPWORDS.has(lower)) return true;
  if (COMMON_WORDS.has(lower)) return true;

  // basic vowel check to reject nonsense like "xqtr"
  if (!/[aeiouyàâäéèêëîïôöùûüÿ]/i.test(lower)) return false;

  // allow common suffixes/prefixes patterns
  if (/(ing|ed|ion|ions|ment|ments|able|ible|ity|ies|ers|er|or|ors|ive|ives|ize|ises|ised|ized)$/i.test(lower)) return true;

  // conservative default: accept plausible-looking alphabetic word length >= 3
  return lower.length >= 3;
}

/**
 * Returns true if the input is a valid phrase we must NOT merge.
 * Conservative: if it's 2+ words and each token looks like a real word, treat as valid phrase.
 */
function isValidPhrase(s) {
  if (!s) return false;
  const txt = String(s).trim();
  const parts = txt.split(/\s+/).filter(Boolean);
  if (parts.length < 2) return false;

  // If both parts are real words, it's a phrase (do not merge)
  const allReal = parts.every(p => isRealWord(p));
  if (allReal) return true;

  // Also treat stopword-led phrases as valid even if second is unknown (e.g., "in X")
  if (STOPWORDS.has(parts[0].toLowerCase())) return true;

  return false;
}

function applyCasingLike(original, corrected) {
  if (!corrected) return corrected;
  const o = (original || "").trim();
  if (!o) return corrected;
  if (/^[A-Z]/.test(o)) return corrected.charAt(0).toUpperCase() + corrected.slice(1);
  return corrected.toLowerCase();
}

function chunkTextOverlap(text, chunkSize = 2600, overlap = 250) {
  if (!text) return [];
  const chunks = [];
  let i = 0;
  while (i < text.length) {
    const end = Math.min(text.length, i + chunkSize);
    chunks.push({ start: i, end, text: text.slice(i, end) });
    if (end === text.length) break;
    i = Math.max(0, end - overlap);
  }
  return chunks;
}

/**
 * ✅ NEW: reject unsafe / invalid AI corrections.
 * - Block merges of multiple valid words ("content to" => "contentto")
 * - Block "no change" corrections
 * - Block corrections that don't produce a real word (single-word corrections only)
 */
function filterInvalidCorrections(errors) {
  return (errors || []).filter(err => {
    const error = String(err?.error || "");
    const correction = String(err?.correction || "");

    // allow multiple_spaces fixes (they are expected to keep spaces)
    if (err?.type === "multiple_spaces") return true;

    // 1. Rejeter si la correction ne produit pas un mot réel (only if correction is a single word)
    if (!/\s/.test(correction) && !isRealWord(correction)) {
      console.log(`❌ REJECTED (not a real word): "${error}" → "${correction}"`);
      return false;
    }

    // 2. Rejeter si c'est une phrase valide qu'on essaie de fusionner
    if (isValidPhrase(error)) {
      console.log(`❌ REJECTED (valid phrase): "${error}"`);
      return false;
    }

    // 3. Rejeter les corrections identiques
    if (error.trim() === correction.trim()) {
      console.log(`❌ REJECTED (no change): "${error}"`);
      return false;
    }

    // 4. NOUVEAU : Rejeter TOUTE fusion de 2+ mots valides séparés
    const words = error.toLowerCase().split(/\s+/).filter(Boolean);
    if (words.length >= 2) {
      const allWordsValid = words.every(w => isRealWord(w));
      if (allWordsValid) {
        // Vérifier si ce n'est PAS un vrai mot fragmenté connu
        const correctionLower = correction.toLowerCase().trim();
        if (!COMMON_WORDS.has(correctionLower)) {
          console.log(`❌ REJECTED (all ${words.length} parts are valid words): "${error}" → "${correction}"`);
          return false;
        }
      }
    }

    // 5. NOUVEAU : Rejeter si la correction contient moins d'espaces que l'original
    //    et que les mots originaux sont tous valides
    const originalSpaces = (error.match(/\s/g) || []).length;
    const correctionSpaces = (correction.match(/\s/g) || []).length;
    if (correctionSpaces < originalSpaces && originalSpaces > 0) {
      const originalWords = error.split(/\s+/).filter(Boolean);
      if (originalWords.every(w => isRealWord(w) && w.length > 1)) {
        console.log(`❌ REJECTED (would merge valid words): "${error}" → "${correction}"`);
        return false;
      }
    }

    return true;
  });
}

/**
 * ✅ NEW: pre-detect "fragmented words" (2+ letters + space + 1+ letters),
 * with strict guardrails to avoid merging valid phrases.
 *
 * Mirrors your requested logic:
 * - Only accept if combined is a real word AND the phrase isn't a valid phrase
 * - SKIP when part2 is a valid word of 2+ letters (both parts valid)
 * - Only push if part2 isn't a real word OR combined is in COMMON_WORDS
 */
function preDetectFragmentedWords(text) {
  const issues = [];
  if (!text || text.length < 3) return issues;

  // Pattern: "Confiden tial" / "enablin g"
  // (we keep it conservative; AI will catch more)
  const splitWord = /\b([A-Za-zÀ-ÿ]{2,25})\s+([A-Za-zÀ-ÿ]{1,25})\b/g;
  let m;

  while ((m = splitWord.exec(text)) !== null) {
    const fullMatch = m[0];
    const part1 = m[1];
    const part2 = m[2];

    // quick skip for stopword phrases & proper names
    if (STOPWORDS.has(part1.toLowerCase()) || STOPWORDS.has(part2.toLowerCase())) continue;
    if (looksLikeProperName(part1, part2)) continue;

    const combined = (part1 + part2).trim();
    if (!isRealWord(combined) && !COMMON_WORDS.has(combined.toLowerCase())) continue;

    // requested stricter condition
    if (isValidPhrase(fullMatch.toLowerCase())) continue;

    // NOUVEAU : Ne pas fusionner si part2 est un mot valide de 2+ lettres
    if (part2.length >= 2 && isRealWord(part2)) {
      console.log(`⏭️ SKIP (both parts valid): "${fullMatch}"`);
      continue;
    }

    // part2 seul n'est généralement pas un mot valide isolé
    if (!isRealWord(part2) || COMMON_WORDS.has(combined.toLowerCase())) {
      const start = m.index;
      const end = m.index + fullMatch.length;

      issues.push({
        id: `pred_frag_${issues.length}`,
        error: fullMatch,
        correction: part1 + part2,
        type: "fragmented_word",
        severity: "high",
        message: "Fragmented word detected (pre-detect)",
        start,
        end,
        globalStart: start,
        globalEnd: end,
        context: text.substring(Math.max(0, start - 20), Math.min(text.length, end + 20)),
        location: `Position ${start}`
      });
    }
  }

  return issues;
}

/**
 * ✅ pre-detect spacing issues (no AI)
 * - Doubles espaces (>=2) entre 2 non-espaces
 * - Mots fragmentés "a bc" (lettre + espace + mot) => "abc"
 * - + stronger fragmented word detection (preDetectFragmentedWords)
 * Returns issues with local offsets (start/end) + global offsets.
 */
function preDetectSpacingIssues(text) {
  const errors = [];
  if (!text || text.length < 3) return errors;

  // 1) Multiple spaces between tokens
  const doubleSpaceRegex = /(\S+)\s{2,}(\S+)/g;
  let match;
  while ((match = doubleSpaceRegex.exec(text)) !== null) {
    const start = match.index;
    const end = match.index + match[0].length;

    errors.push({
      id: `spacing_${errors.length}`,
      error: match[0],
      correction: match[1] + " " + match[2],
      type: "multiple_spaces",
      severity: "medium",
      message: "Multiple spaces detected",
      start,
      end,
      globalStart: start,
      globalEnd: end,
      context: text.substring(Math.max(0, start - 20), Math.min(text.length, end + 20)),
      location: `Position ${start}`
    });
  }

  // 2) Fragmented word (single letter + space + letters)
  const fragmentedRegex = /\b([A-Za-zÀ-ÿ])\s+([A-Za-zÀ-ÿ]{2,})\b/g;
  while ((match = fragmentedRegex.exec(text)) !== null) {
    const start = match.index;
    const end = match.index + match[0].length;

    // guardrail: do not merge if second part is a real word (e.g., "a message")
    const part1 = match[1];
    const part2 = match[2];
    if (part2.length >= 2 && isRealWord(part2)) {
      console.log(`⏭️ SKIP (letter + valid word): "${match[0]}"`);
      continue;
    }

    errors.push({
      id: `fragment_${errors.length}`,
      error: match[0],
      correction: part1 + part2,
      type: "fragmented_word",
      severity: "high",
      message: "Fragmented word detected",
      start,
      end,
      globalStart: start,
      globalEnd: end,
      context: text.substring(Math.max(0, start - 20), Math.min(text.length, end + 20)),
      location: `Position ${start}`
    });
  }

  // 3) Stronger fragmented words detection (2+ letters + space + 1+ letters)
  const stronger = preDetectFragmentedWords(text);
  errors.push(...stronger);

  return errors;
}

/**
 * Build "candidates" spans (suspects) for the AI.
 * Candidates are NOT auto-fixes.
 */
function buildCandidates(text) {
  const candidates = [];
  if (!text) return candidates;

  // 1) multi spaces inside word: "p  otential"
  const multiSpaceWord = /\b([A-Za-zÀ-ÿ]{1,25})\s{2,}([A-Za-zÀ-ÿ]{1,25})\b/g;
  let m;
  while ((m = multiSpaceWord.exec(text)) !== null) {
    candidates.push({ error: m[0], hint: "multiple_spaces_inside_word" });
  }

  // 2) fragmented word with single space: "soc ial", "gen uine", "dar ker"
  // improved recall: allow fragments up to 6 each, but still avoid stopwords & proper names
  // NOTE: We keep candidates broad; filterInvalidCorrections + prompt rules will prevent bad merges.
  const splitWord = /\b([A-Za-zÀ-ÿ]{2,25})\s+([A-Za-zÀ-ÿ]{1,25})\b/g;
  while ((m = splitWord.exec(text)) !== null) {
    const p1 = m[1];
    const p2 = m[2];

    if (STOPWORDS.has(p1.toLowerCase()) || STOPWORDS.has(p2.toLowerCase())) continue;
    if (looksLikeProperName(p1, p2)) continue;

    // if both parts are real words -> do not even send as candidate (avoid AI temptation)
    if (p2.length >= 2 && isRealWord(p1) && isRealWord(p2)) continue;
    if (isValidPhrase(m[0].toLowerCase())) continue;

    const combined = (p1 + p2).trim();
    if (!isPlausibleWord(combined)) continue;

    const strong = (p1.length <= 6) || (p2.length <= 6);
    if (!strong) continue;

    candidates.push({ error: m[0], hint: "likely_fragmented_word" });
  }

  // 3) punctuation inside word: "corpo,rations"
  const punctInsideWord = /\b([A-Za-zÀ-ÿ]{2,})[,:;.'-]+([A-Za-zÀ-ÿ]{2,})\b/g;
  while ((m = punctInsideWord.exec(text)) !== null) {
    candidates.push({ error: m[0], hint: "punctuation_inside_word" });
  }

  // 4) repeated prefix noise: "gggdigital"
  const repeatedPrefix = /\b([a-zA-ZÀ-ÿ])\1{2,}([A-Za-zÀ-ÿ]{3,})\b/g;
  while ((m = repeatedPrefix.exec(text)) !== null) {
    candidates.push({ error: m[0], hint: "repeated_prefix_noise" });
  }

  return candidates;
}

// ================= AI PROMPT =================

// Keep the system prompt format strict (JSON-only), but reinforce merge rules in USER content (next).
const PROOFREADER_PROMPT = `
You are an expert English proofreader.

You MUST detect and propose corrections for:
- spelling mistakes
- grammar/conjugation mistakes
- obvious typos
- corrupted words (spaces inside word, multiple spaces, punctuation inside word, repeated junk letters)

Rules:
1) Do NOT rewrite sentences. Only minimal fixes.
2) Use context like a human.
3) Return corrections as smallest safe span.
4) Output MUST include start/end offsets inside the chunk.
5) If ambiguous, propose the most likely fix but keep the span minimal.

OUTPUT: return ONLY a JSON array:
[
  {
    "error": "exact substring from input",
    "correction": "corrected substring",
    "type": "fragmented_word|multiple_spaces|spelling|grammar|capitalization|punctuation",
    "severity": "high|medium|low",
    "message": "short reason",
    "start": 123,
    "end": 140
  }
]
Return ONLY the JSON array. No extra text.
`;

async function aiAnalyzeChunk(chunkText, candidates) {
  const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
  if (!OPENAI_API_KEY) return [];

  // =========================
  // MODIF 2 — stronger explicit instructions to AI (user content)
  // =========================
  const strictInstructions = `Analyze this text and find ALL spelling, grammar, conjugation, syntax, and punctuation errors.

CRITICAL RULES:
- NEVER merge words that are correctly separated by spaces
- "forces shaping" = TWO WORDS (correct) - do NOT suggest "forcesshaping"
- "content to" = TWO WORDS (correct) - do NOT suggest "contentto"
- "It was the" = THREE WORDS (correct) - do NOT suggest "Itwasthe"
- Only suggest merging when a word is INCORRECTLY SPLIT like "Confiden tial" → "Confidential"
- A true fragmented word = one word split by accidental space, not two separate valid words

REMEMBER:
- Only suggest corrections that result in REAL words
- Never merge separate valid words like "In the" or "a message"
- Floating letters usually belong to the LEFT word (e.g., "enablin g" → "enabling")

TEXT TO ANALYZE:
${chunkText.slice(0, 15000)}`;

  const payload = {
    model: "gpt-4o-mini",
    messages: [
      { role: "system", content: PROOFREADER_PROMPT },
      {
        role: "user",
        content:
          `${strictInstructions}\n\n` +
          `SUSPECT CANDIDATES (may be real errors or false alarms):\n` +
          `${JSON.stringify(candidates.slice(0, 120), null, 2)}\n\n` +
          `Task:\n` +
          `1) Validate candidates: keep only real errors and give corrections.\n` +
          `2) Also find other spelling/grammar/typo errors in the chunk.\n` +
          `3) Provide accurate start/end offsets for each "error" inside the chunk.\n` +
          `Return JSON only.`
      }
    ],
    temperature: 0.1,
    max_tokens: 3000
  };

  const response = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    console.error(`[SPELLCHECK] OpenAI API error: ${response.status}`);
    return [];
  }

  const data = await response.json();
  const content = data.choices?.[0]?.message?.content || "[]";

  try {
    const jsonMatch = content.match(/\[[\s\S]*\]/);
    if (!jsonMatch) return [];
    const arr = JSON.parse(jsonMatch[0]);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    console.error("[SPELLCHECK] Failed to parse AI JSON:", e.message);
    return [];
  }
}

function sanitizeCorrections(list, chunkText) {
  const out = [];

  for (const c of (list || [])) {
    const error = (c.error || "");
    const correction = (c.correction || "");
    const type = String(c.type || "").trim() || "spelling";

    const start = Number.isFinite(c.start) ? c.start : null;
    const end = Number.isFinite(c.end) ? c.end : null;

    if (!error || !correction) continue;
    if (error.trim().length === 0 || correction.trim().length === 0) continue;
    if (error.length > 140 || correction.length > 140) continue;

    // must have offsets
    if (start === null || end === null) continue;
    if (start < 0 || end <= start || end > chunkText.length) continue;

    // verify substring matches (super important)
    const slice = chunkText.slice(start, end);
    if (slice !== error) {
      // allow small mismatch (AI sometimes trims) -> try fallback by searching error once
      const idx = chunkText.indexOf(error);
      if (idx === -1) continue;
      // if multiple occurrences, skip to avoid wrong apply
      if (chunkText.indexOf(error, idx + 1) !== -1) continue;
      c.start = idx;
      c.end = idx + error.length;
    }

    if (error.toLowerCase() === correction.toLowerCase()) continue;

    out.push({
      error,
      correction:
        (type === "fragmented_word" || type === "multiple_spaces")
          ? applyCasingLike(error, correction)
          : correction,
      type,
      severity: c.severity || "medium",
      message: c.message || "AI-detected",
      start: c.start,
      end: c.end
    });
  }

  // dedup by span + correction (chunk-local)
  const seen = new Set();
  const dedup = [];
  for (const x of out) {
    const key = `${x.start}:${x.end}:${x.correction}`.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    dedup.push(x);
  }
  return dedup;
}

// ================= MAIN =================

export async function checkSpellingWithAI(text) {
  console.log("📝 AI spell-check VERSION 7.2 starting...");
  if (!text || text.length < 10) return [];

  // ✅ PRE-DETECT (no AI)
  const spacingErrors = preDetectSpacingIssues(text);

  // IMPORTANT: stop truncating at 15k -> raise a lot
  const maxChars = 120000; // ✅ increase recall massively
  const input = text.length > maxChars ? text.slice(0, maxChars) : text;

  console.log(`[SPELLCHECK] Text length used: ${input.length} characters`);
  console.log(`[SPELLCHECK] Pre-detected spacing issues: ${spacingErrors.length}`);

  const chunks = chunkTextOverlap(input, 2600, 250);

  const aiErrors = [];

  for (let i = 0; i < chunks.length; i++) {
    const { start: chunkStart, text: chunkText } = chunks[i];

    const candidates = buildCandidates(chunkText);
    console.log(`[SPELLCHECK] Chunk ${i + 1}/${chunks.length} candidates: ${candidates.length}`);

    const aiRaw = await aiAnalyzeChunk(chunkText, candidates);
    console.log(`[AI RAW] Chunk ${i + 1}: ${aiRaw.length} items`);

    const cleaned = sanitizeCorrections(aiRaw, chunkText);

    // convert offsets to global offsets
    for (const item of cleaned) {
      aiErrors.push({
        ...item,
        globalStart: chunkStart + item.start,
        globalEnd: chunkStart + item.end
      });
    }
  }

  // ✅ merge pre-detect + AI
  const allErrors = [...spacingErrors, ...aiErrors];

  // global dedup by global span + correction
  const seen = new Set();
  const deduped = [];
  for (const x of allErrors) {
    const gs = Number.isFinite(x.globalStart) ? x.globalStart : x.start;
    const ge = Number.isFinite(x.globalEnd) ? x.globalEnd : x.end;
    const corr = String(x.correction || "");
    const key = `${gs}:${ge}:${corr}`.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    deduped.push({
      ...x,
      globalStart: gs,
      globalEnd: ge
    });
  }

  // ✅ MODIF 1 — apply invalid-correction filter (especially blocks word merges)
  const final = filterInvalidCorrections(deduped);

  console.log(`✅ AI spell-check complete: ${final.length} issues found`);
  final.slice(0, 25).forEach((issue, i) => {
    console.log(
      `  ${i + 1}. [${issue.type}] "${issue.error}" → "${issue.correction}" @${issue.globalStart}-${issue.globalEnd}`
    );
  });

  // IMPORTANT: keep shape compatible with old code too
  // (front can ignore offsets if it wants)
  return final;
}

export default { checkSpellingWithAI };
