// lib/aiSpellCheck.js
// VERSION 7.0 - High-recall AI proofreading with overlap + offsets (start/end)
// Goal: detect spelling + grammar + typos + corrupted words with minimal fixes
// ✅ Regex = candidates only
// ✅ AI validates candidates + finds additional issues
// ✅ Overlap chunking to avoid missing boundary errors
// ✅ AI returns offsets (start/end) inside chunk to apply safely downstream

const STOPWORDS = new Set([
  'a','an','the','to','of','in','on','at','by','for','with','from','into','onto','upon','within','without',
  'and','or','but','so','yet','nor',
  'is','are','was','were','be','been','being',
  'can','could','would','should','will','may','might','must',
  'has','have','had','do','does','did',
  'not','how','one','due','as'
]);

function looksLikeProperName(p1, p2) {
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
 * Build "candidates" spans (suspects) for the AI.
 * Candidates are NOT auto-fixes.
 */
function buildCandidates(text) {
  const candidates = [];
  if (!text) return candidates;

  // 1) multi spaces inside word: "p  otential"
  const multiSpaceWord = /\b([A-Za-z]{1,25})\s{2,}([A-Za-z]{1,25})\b/g;
  let m;
  while ((m = multiSpaceWord.exec(text)) !== null) {
    candidates.push({ error: m[0], hint: 'multiple_spaces_inside_word' });
  }

  // 2) fragmented word with single space: "soc ial", "gen uine", "dar ker"
  // improved recall: allow fragments up to 6 each, but still avoid stopwords & proper names
  const splitWord = /\b([A-Za-z]{2,25})\s+([A-Za-z]{2,25})\b/g;
  while ((m = splitWord.exec(text)) !== null) {
    const p1 = m[1];
    const p2 = m[2];

    if (STOPWORDS.has(p1.toLowerCase()) || STOPWORDS.has(p2.toLowerCase())) continue;
    if (looksLikeProperName(p1, p2)) continue;

    const combined = (p1 + p2).trim();
    if (!isPlausibleWord(combined)) continue;

    // stronger recall than v6: accept if either side <= 6 (not only <=3)
    const strong = (p1.length <= 6) || (p2.length <= 6);
    if (!strong) continue;

    candidates.push({ error: m[0], hint: 'likely_fragmented_word' });
  }

  // 3) punctuation inside word: "corpo,rations"
  const punctInsideWord = /\b([A-Za-z]{2,})[,:;.'-]+([A-Za-z]{2,})\b/g;
  while ((m = punctInsideWord.exec(text)) !== null) {
    candidates.push({ error: m[0], hint: 'punctuation_inside_word' });
  }

  // 4) repeated prefix noise: "gggdigital"
  const repeatedPrefix = /\b([a-zA-Z])\1{2,}([A-Za-z]{3,})\b/g;
  while ((m = repeatedPrefix.exec(text)) !== null) {
    candidates.push({ error: m[0], hint: 'repeated_prefix_noise' });
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

  const payload = {
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: PROOFREADER_PROMPT },
      {
        role: 'user',
        content:
          `TEXT CHUNK:\n<<<\n${chunkText}\n>>>\n\n` +
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

function sanitizeCorrections(list, chunkText) {
  const out = [];

  for (const c of (list || [])) {
    const error = (c.error || '');
    const correction = (c.correction || '');
    const type = String(c.type || '').trim() || 'spelling';

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
      correction: (type === 'fragmented_word' || type === 'multiple_spaces')
        ? applyCasingLike(error, correction)
        : correction,
      type,
      severity: c.severity || 'medium',
      message: c.message || 'AI-detected',
      start: c.start,
      end: c.end
    });
  }

  // dedup by span + correction
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
  console.log('📝 AI spell-check VERSION 7.0 starting...');
  if (!text || text.length < 10) return [];

  // IMPORTANT: stop truncating at 15k -> raise a lot
  const maxChars = 120000; // ✅ increase recall massively
  const input = text.length > maxChars ? text.slice(0, maxChars) : text;

  console.log(`[SPELLCHECK] Text length used: ${input.length} characters`);

  const chunks = chunkTextOverlap(input, 2600, 250);

  const all = [];

  for (let i = 0; i < chunks.length; i++) {
    const { start: chunkStart, text: chunkText } = chunks[i];

    const candidates = buildCandidates(chunkText);
    console.log(`[SPELLCHECK] Chunk ${i + 1}/${chunks.length} candidates: ${candidates.length}`);

    const aiRaw = await aiAnalyzeChunk(chunkText, candidates);
    console.log(`[AI RAW] Chunk ${i + 1}: ${aiRaw.length} items`);

    const cleaned = sanitizeCorrections(aiRaw, chunkText);

    // convert offsets to global offsets
    for (const item of cleaned) {
      all.push({
        ...item,
        globalStart: chunkStart + item.start,
        globalEnd: chunkStart + item.end
      });
    }
  }

  // global dedup by global span + correction
  const seen = new Set();
  const final = [];
  for (const x of all) {
    const key = `${x.globalStart}:${x.globalEnd}:${x.correction}`.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    final.push(x);
  }

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
