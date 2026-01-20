// lib/aiSpellCheck.js
// VERSION 7.3 - Parallel chunking (5000 chars) + guardrails from 7.2
// ✅ Parallel chunk processing (MAX_PARALLEL_CHUNKS)
// ✅ Keeps 7.2 safety: pre-detect spacing/fragmentation + filterInvalidCorrections()
// ✅ Accepts (text, openaiApiKey) as requested (backward compatible if key omitted)
// ✅ Returns offsets in globalStart/globalEnd AND offset (compat) for downstream

// -------------------- NEW CONSTANTS --------------------
const CHUNK_SIZE = 5000; // Characters per chunk
const MAX_PARALLEL_CHUNKS = 3; // Process 3 chunks at a time

// -------------------- 7.2 BASE (GUARDRAILS) --------------------
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
  if (/\s/.test(s)) return false; // single-word only
  if (!isPlausibleWord(s)) return false;

  const lower = s.toLowerCase();
  if (STOPWORDS.has(lower)) return true;
  if (COMMON_WORDS.has(lower)) return true;

  // basic vowel check to reject nonsense like "xqtr"
  if (!/[aeiouyàâäéèêëîïôöùûüÿ]/i.test(lower)) return false;

  // allow common suffixes/prefixes patterns
  if (/(ing|ed|ion|ions|ment|ments|able|ible|ity|ies|ers|er|or|ors|ive|ives|ize|ises|ised|ized)$/i.test(lower)) return true;

  // conservative default
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

  const allReal = parts.every(p => isRealWord(p));
  if (allReal) return true;

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

/**
 * ✅ reject unsafe / invalid AI corrections.
 */
function filterInvalidCorrections(errors) {
  return (errors || []).filter(err => {
    const error = String(err?.error || "");
    const correction = String(err?.correction || "");

    // allow multiple_spaces fixes (they are expected to keep spaces)
    if (err?.type === "multiple_spaces") return true;

    // 1) reject if correction is a single word and not a "real" word
    if (!/\s/.test(correction) && !isRealWord(correction)) {
      console.log(`❌ REJECTED (not a real word): "${error}" → "${correction}"`);
      return false;
    }

    // 2) reject if it's a valid phrase that someone tries to merge
    if (isValidPhrase(error)) {
      console.log(`❌ REJECTED (valid phrase): "${error}"`);
      return false;
    }

    // 3) reject no-op
    if (error.trim() === correction.trim()) {
      console.log(`❌ REJECTED (no change): "${error}"`);
      return false;
    }

    // 4) reject merging multiple valid words (unless allow-listed as true fragmented word)
    const words = error.toLowerCase().split(/\s+/).filter(Boolean);
    if (words.length >= 2) {
      const allWordsValid = words.every(w => isRealWord(w));
      if (allWordsValid) {
        const correctionLower = correction.toLowerCase().trim();
        if (!COMMON_WORDS.has(correctionLower)) {
          console.log(`❌ REJECTED (all ${words.length} parts are valid words): "${error}" → "${correction}"`);
          return false;
        }
      }
    }

    // 5) reject if correction reduces spaces but original tokens are valid
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
 * ✅ stricter pre-detect of fragmented words (2+ letters + space + 1+ letters)
 */
function preDetectFragmentedWords(text) {
  const issues = [];
  if (!text || text.length < 3) return issues;

  const splitWord = /\b([A-Za-zÀ-ÿ]{2,25})\s+([A-Za-zÀ-ÿ]{1,25})\b/g;
  let m;

  while ((m = splitWord.exec(text)) !== null) {
    const fullMatch = m[0];
    const part1 = m[1];
    const part2 = m[2];

    if (STOPWORDS.has(part1.toLowerCase()) || STOPWORDS.has(part2.toLowerCase())) continue;
    if (looksLikeProperName(part1, part2)) continue;

    const combined = (part1 + part2).trim();
    if (!isRealWord(combined) && !COMMON_WORDS.has(combined.toLowerCase())) continue;

    if (isValidPhrase(fullMatch.toLowerCase())) continue;

    // do not merge if part2 is itself a valid 2+ letters word
    if (part2.length >= 2 && isRealWord(part2)) {
      continue;
    }

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
        offset: start, // compat
        context: text.substring(Math.max(0, start - 20), Math.min(text.length, end + 20)),
        location: `Position ${start}`
      });
    }
  }

  return issues;
}

/**
 * ✅ pre-detect spacing issues (no AI)
 */
function preDetectSpacingIssues(text) {
  const errors = [];
  if (!text || text.length < 3) return errors;

  // 1) multiple spaces between tokens
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
      offset: start, // compat
      context: text.substring(Math.max(0, start - 20), Math.min(text.length, end + 20)),
      location: `Position ${start}`
    });
  }

  // 2) fragmented word (single letter + space + letters)
  const fragmentedRegex = /\b([A-Za-zÀ-ÿ])\s+([A-Za-zÀ-ÿ]{2,})\b/g;
  while ((match = fragmentedRegex.exec(text)) !== null) {
    const start = match.index;
    const end = match.index + match[0].length;

    const part2 = match[2];
    if (part2.length >= 2 && isRealWord(part2)) continue;

    errors.push({
      id: `fragment_${errors.length}`,
      error: match[0],
      correction: match[1] + match[2],
      type: "fragmented_word",
      severity: "high",
      message: "Fragmented word detected",
      start,
      end,
      globalStart: start,
      globalEnd: end,
      offset: start, // compat
      context: text.substring(Math.max(0, start - 20), Math.min(text.length, end + 20)),
      location: `Position ${start}`
    });
  }

  // 3) stronger fragmented words
  errors.push(...preDetectFragmentedWords(text));
  return errors;
}

/**
 * Build "candidates" spans (suspects) for the AI.
 * Candidates are NOT auto-fixes.
 */
function buildCandidates(text) {
  const candidates = [];
  if (!text) return candidates;

  // multi spaces inside word: "p  otential"
  const multiSpaceWord = /\b([A-Za-zÀ-ÿ]{1,25})\s{2,}([A-Za-zÀ-ÿ]{1,25})\b/g;
  let m;
  while ((m = multiSpaceWord.exec(text)) !== null) {
    candidates.push({ error: m[0], hint: "multiple_spaces_inside_word" });
  }

  // fragmented word with single space: "soc ial"
  const splitWord = /\b([A-Za-zÀ-ÿ]{2,25})\s+([A-Za-zÀ-ÿ]{1,25})\b/g;
  while ((m = splitWord.exec(text)) !== null) {
    const p1 = m[1];
    const p2 = m[2];

    if (STOPWORDS.has(p1.toLowerCase()) || STOPWORDS.has(p2.toLowerCase())) continue;
    if (looksLikeProperName(p1, p2)) continue;

    // if both parts are real words -> do not send as candidate
    if (p2.length >= 2 && isRealWord(p1) && isRealWord(p2)) continue;
    if (isValidPhrase(m[0].toLowerCase())) continue;

    const combined = (p1 + p2).trim();
    if (!isPlausibleWord(combined)) continue;

    const strong = (p1.length <= 6) || (p2.length <= 6);
    if (!strong) continue;

    candidates.push({ error: m[0], hint: "likely_fragmented_word" });
  }

  // punctuation inside word: "corpo,rations"
  const punctInsideWord = /\b([A-Za-zÀ-ÿ]{2,})[,:;.'-]+([A-Za-zÀ-ÿ]{2,})\b/g;
  while ((m = punctInsideWord.exec(text)) !== null) {
    candidates.push({ error: m[0], hint: "punctuation_inside_word" });
  }

  // repeated prefix noise: "gggdigital"
  const repeatedPrefix = /\b([a-zA-ZÀ-ÿ])\1{2,}([A-Za-zÀ-ÿ]{3,})\b/g;
  while ((m = repeatedPrefix.exec(text)) !== null) {
    candidates.push({ error: m[0], hint: "repeated_prefix_noise" });
  }

  return candidates;
}

// -------------------- AI PROMPT (kept, strict JSON array) --------------------
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

// -------------------- AI CALL (now takes apiKey) --------------------
async function aiAnalyzeChunk(chunkText, candidates, openaiApiKey) {
  if (!openaiApiKey) return [];

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
      Authorization: `Bearer ${openaiApiKey}`,
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

    if (start === null || end === null) continue;
    if (start < 0 || end <= start || end > chunkText.length) continue;

    // verify substring matches
    const slice = chunkText.slice(start, end);
    if (slice !== error) {
      const idx = chunkText.indexOf(error);
      if (idx === -1) continue;
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

// -------------------- NEW: single-chunk checker (used by parallel batches) --------------------
async function checkChunkSpelling(chunk, openaiApiKey) {
  try {
    const candidates = buildCandidates(chunk.text);
    const aiRaw = await aiAnalyzeChunk(chunk.text, candidates, openaiApiKey);
    const cleaned = sanitizeCorrections(aiRaw, chunk.text);

    // convert to global offsets
    return cleaned.map(item => {
      const globalStart = (chunk.startOffset || 0) + item.start;
      const globalEnd = (chunk.startOffset || 0) + item.end;
      return {
        ...item,
        globalStart,
        globalEnd,
        offset: globalStart // compat with older consumers
      };
    });
  } catch (error) {
    console.error(`[SPELLCHECK] Chunk error:`, error?.message || error);
    return [];
  }
}

/**
 * Remove duplicate errors (same correction on same global span)
 */
function deduplicateErrors(errors) {
  const seen = new Set();
  return (errors || []).filter(err => {
    const gs = Number.isFinite(err.globalStart) ? err.globalStart : (err.offset || 0);
    const ge = Number.isFinite(err.globalEnd) ? err.globalEnd : (gs + (String(err.error || "").length));
    const key = `${String(err.error || "")}:${gs}:${ge}:${String(err.correction || "")}`.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

/**
 * Check spelling with AI using parallel chunking
 * Replaces the main export function (as requested)
 */
export async function checkSpellingWithAI(text, openaiApiKey) {
  // backward compat: allow calling checkSpellingWithAI(text) if env var exists
  const apiKey = openaiApiKey || process.env.OPENAI_API_KEY;

  if (!apiKey || !text || text.length < 50) {
    console.log("[SPELLCHECK] Skipping - no API key or text too short");
    return [];
  }

  const startTime = Date.now();

  // ✅ PRE-DETECT (no AI) on full text (but keep it bounded)
  const maxChars = 120000; // keep your recall boost
  const input = text.length > maxChars ? text.slice(0, maxChars) : text;
  const spacingErrors = preDetectSpacingIssues(input);

  // Split text into chunks (NO overlap, as per requested spec)
  const chunks = [];
  for (let i = 0; i < input.length; i += CHUNK_SIZE) {
    chunks.push({
      text: input.slice(i, i + CHUNK_SIZE),
      startOffset: i
    });
  }

  console.log(`[SPELLCHECK] Processing ${chunks.length} chunks (${input.length} chars total)`);
  console.log(`[SPELLCHECK] Pre-detected spacing issues: ${spacingErrors.length}`);

  // If only 1 chunk, process directly
  if (chunks.length === 1) {
    const aiErrors = await checkChunkSpelling(chunks[0], apiKey);
    const merged = deduplicateErrors([...spacingErrors, ...aiErrors]);
    const final = filterInvalidCorrections(merged);

    console.log(`[SPELLCHECK] Single chunk completed in ${Date.now() - startTime}ms, found ${final.length} issues`);
    return final;
  }

  // Process chunks in parallel batches
  const allAiErrors = [];

  for (let i = 0; i < chunks.length; i += MAX_PARALLEL_CHUNKS) {
    const batch = chunks.slice(i, i + MAX_PARALLEL_CHUNKS);
    const batchStart = Date.now();
    const batchNum = Math.floor(i / MAX_PARALLEL_CHUNKS) + 1;
    const totalBatches = Math.ceil(chunks.length / MAX_PARALLEL_CHUNKS);

    console.log(`[SPELLCHECK] Processing batch ${batchNum}/${totalBatches}...`);

    const results = await Promise.all(
      batch.map(c => checkChunkSpelling(c, apiKey))
    );

    console.log(`[SPELLCHECK] Batch ${batchNum}/${totalBatches} completed in ${Date.now() - batchStart}ms`);

    results.forEach(chunkErrors => {
      allAiErrors.push(...chunkErrors);
    });
  }

  // Merge + deduplicate + filter
  const merged = deduplicateErrors([...spacingErrors, ...allAiErrors]);
  const final = filterInvalidCorrections(merged);

  console.log(`[SPELLCHECK] Total: ${final.length} issues found in ${Date.now() - startTime}ms`);
  return final;
}

export default { checkSpellingWithAI };
