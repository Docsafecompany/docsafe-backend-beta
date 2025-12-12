// lib/aiProofreadAnchored.js
// VERSION 5.1 — 100% AI + micro-deterministic for short fragmented words
// ✅ AI detects real spelling/grammar errors like a human
// ✅ Micro-deterministic catches "o f" → "of", "i n" → "in" (AI blind spot)
// ✅ Conservative - no false positives
// ✅ Context-scored mapping back to RAW text

const PROMPT = `
You are an expert human proofreader reviewing a professional document.

Your task: find REAL spelling, grammar, and formatting errors that a professional editor would catch.

DETECT ONLY:
- Genuine spelling mistakes (typos, misspellings)
- Grammar errors (subject-verb agreement, wrong tense, conjugation)
- Punctuation errors (missing periods, wrong commas, extra punctuation)
- Double spaces (two or more consecutive spaces)
- Short common words broken by space:
  ✅ "o f" → "of"
  ✅ "i n" → "in"  
  ✅ "t o" → "to"
  ✅ "i s" → "is"
  ✅ "a s" → "as"
- Longer words broken by space ONLY if merging creates a REAL word:
  ✅ "soc ial" → "social" (real word)
  ✅ "confiden tial" → "confidential" (real word)
  ❌ "sty Wal" → DO NOT merge (not a word, likely two names)
  ❌ "com Pro" → DO NOT merge (separate words/names)
- Words stuck together ONLY if splitting creates two REAL words:
  ✅ "suchas" → "such as"
  ✅ "Impactof" → "Impact of"
  ❌ "therefore" → DO NOT split (valid word)

STRICT RULES:
1. Do NOT merge words if the result is NOT a valid English/French word
2. Do NOT change proper nouns (names of people, places, companies)
3. Do NOT flag technical terms, acronyms, product names, or codes
4. Do NOT rewrite sentences - only fix minimal character spans
5. Do NOT change meaning or style
6. When in doubt, DO NOT flag it - be conservative
7. Maximum 30 issues per text chunk (focus on most important errors)

Return JSON array only, no markdown, no explanation:
[
  {
    "error": "exact text to replace (shortest possible span)",
    "correction": "replacement text",
    "type": "spelling|grammar|punctuation|spacing",
    "severity": "high|medium|low",
    "message": "brief explanation (5-10 words max)",
    "contextBefore": "up to 40 chars before the error",
    "contextAfter": "up to 40 chars after the error"
  }
]

If no errors found, return: []
`;

// ============================================================
// Utility functions
// ============================================================

function safeParseJSONArray(raw) {
  if (!raw) return [];
  const m = raw.match(/\[[\s\S]*\]/);
  if (!m) return [];
  try { return JSON.parse(m[0]); } catch { return []; }
}

function normalizeSpaces(s) {
  return (s || "").replace(/\s+/g, " ").trim();
}

function normalizeContext(s) {
  return (s || "").replace(/\s+/g, " ").trim().slice(0, 80);
}

function mkId() {
  return `spell_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function addAnchors(list) {
  return (list || []).map((it) => ({
    id: mkId(),
    error: it.error,
    correction: it.correction,
    type: it.type || "spelling",
    severity: it.severity || "medium",
    message: it.message || "Detected issue",
    contextBefore: it.contextBefore || "",
    contextAfter: it.contextAfter || "",
    startIndex: typeof it.startIndex === "number" ? it.startIndex : null,
    endIndex: typeof it.endIndex === "number" ? it.endIndex : null,
  }));
}

function chunkText(text, chunkSize = 8000, overlap = 450) {
  const t = text || "";
  const chunks = [];
  let i = 0;
  while (i < t.length) {
    const start = i;
    const end = Math.min(t.length, i + chunkSize);
    chunks.push({ start, end, text: t.slice(start, end) });
    if (end === t.length) break;
    i = end - overlap;
    if (i < 0) i = 0;
  }
  return chunks;
}

function isProbablyGarbageSpan(s) {
  const x = String(s || "");
  if (!x.trim()) return true;
  if (x.length > 120) return true;
  if (/[<>]/.test(x) && /(class=|data-|style=|<\w+|<\/\w+)/i.test(x)) return true;
  if (/[\\\/\{\}\[\]@#\$%\^&\*]/.test(x)) return true;
  return false;
}

function computeContextFromFullText(fullText, globalIndex, errorLen) {
  const before = fullText.slice(Math.max(0, globalIndex - 40), globalIndex);
  const after = fullText.slice(
    globalIndex + errorLen,
    Math.min(fullText.length, globalIndex + errorLen + 40)
  );
  return { contextBefore: before, contextAfter: after };
}

// ============================================================
// Noise masking for AI (keeps LENGTH) so anchors/indexing remain valid
// ============================================================

function maskNoiseKeepLength(raw) {
  const s = String(raw || "");

  // mask long numeric runs: 123456789 -> 000000000 (same length)
  let out = s.replace(/\b\d{6,}\b/g, (m) => "0".repeat(m.length));

  // mask short "codes" like A123, BC2345 -> A000, BC0000 (same length)
  out = out.replace(/\b[A-Z]{1,3}\d{2,6}\b/g, (m) => {
    const letters = m.match(/^[A-Z]+/)?.[0] || "";
    const digits = m.slice(letters.length);
    return letters + "0".repeat(digits.length);
  });

  // mask email addresses (keep length)
  out = out.replace(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, (m) => 
    "x".repeat(m.length)
  );

  // mask URLs (keep length)
  out = out.replace(/https?:\/\/[^\s]+/g, (m) => "x".repeat(m.length));

  return out;
}

// ============================================================
// MICRO-DETERMINISTIC: Short fragmented common words only
// This catches patterns the AI struggles with (1-2 char fragments)
// ============================================================

const SHORT_COMMON_WORDS = new Map([
  // 2-letter words
  ["o f", "of"],
  ["i n", "in"],
  ["o n", "on"],
  ["a t", "at"],
  ["t o", "to"],
  ["i s", "is"],
  ["i t", "it"],
  ["a s", "as"],
  ["o r", "or"],
  ["a n", "an"],
  ["b y", "by"],
  ["u p", "up"],
  ["w e", "we"],
  ["h e", "he"],
  ["m y", "my"],
  ["n o", "no"],
  ["s o", "so"],
  ["d o", "do"],
  ["i f", "if"],
  ["m e", "me"],
  ["u s", "us"],
  // 3-letter words (1+2 or 2+1 splits)
  ["t he", "the"],
  ["th e", "the"],
  ["a nd", "and"],
  ["an d", "and"],
  ["f or", "for"],
  ["fo r", "for"],
  ["a re", "are"],
  ["ar e", "are"],
  ["b ut", "but"],
  ["bu t", "but"],
  ["n ot", "not"],
  ["no t", "not"],
  ["y ou", "you"],
  ["yo u", "you"],
  ["a ll", "all"],
  ["al l", "all"],
  ["c an", "can"],
  ["ca n", "can"],
  ["h ad", "had"],
  ["ha d", "had"],
  ["h er", "her"],
  ["he r", "her"],
  ["w as", "was"],
  ["wa s", "was"],
  ["o ne", "one"],
  ["on e", "one"],
  ["o ur", "our"],
  ["ou r", "our"],
  ["o ut", "out"],
  ["ou t", "out"],
  ["d ay", "day"],
  ["da y", "day"],
  ["g et", "get"],
  ["ge t", "get"],
  ["h as", "has"],
  ["ha s", "has"],
  ["h im", "him"],
  ["hi m", "him"],
  ["h is", "his"],
  ["hi s", "his"],
  ["h ow", "how"],
  ["ho w", "how"],
  ["n ew", "new"],
  ["ne w", "new"],
  ["n ow", "now"],
  ["no w", "now"],
  ["o ld", "old"],
  ["ol d", "old"],
  ["s ee", "see"],
  ["se e", "see"],
  ["t wo", "two"],
  ["tw o", "two"],
  ["w ay", "way"],
  ["wa y", "way"],
  ["w ho", "who"],
  ["wh o", "who"],
  ["o il", "oil"],
  ["oi l", "oil"],
  ["s ay", "say"],
  ["sa y", "say"],
  ["s he", "she"],
  ["sh e", "she"],
  ["a ny", "any"],
  ["an y", "any"],
  // 4-letter common words (2+2 splits)
  ["wi th", "with"],
  ["th at", "that"],
  ["ha ve", "have"],
  ["th is", "this"],
  ["wi ll", "will"],
  ["yo ur", "your"],
  ["fr om", "from"],
  ["th ey", "they"],
  ["be en", "been"],
  ["ca ll", "call"],
  ["wh at", "what"],
  ["wh en", "when"],
  ["ma ke", "make"],
  ["ca me", "came"],
  ["co me", "come"],
  ["co ul", "could"],  // partial
  ["ma ny", "many"],
  ["so me", "some"],
  ["ti me", "time"],
  ["ve ry", "very"],
  ["wh en", "when"],
  ["wo rd", "word"],
  ["wo rk", "work"],
  ["ye ar", "year"],
  ["al so", "also"],
  ["ba ck", "back"],
  ["be en", "been"],
  ["bo th", "both"],
  ["ea ch", "each"],
  ["fi nd", "find"],
  ["fi rs", "first"], // partial
  ["gi ve", "give"],
  ["go od", "good"],
  ["ju st", "just"],
  ["kn ow", "know"],
  ["la st", "last"],
  ["li ke", "like"],
  ["li ne", "line"],
  ["lo ng", "long"],
  ["lo ok", "look"],
  ["mo re", "more"],
  ["mu ch", "much"],
  ["mu st", "must"],
  ["na me", "name"],
  ["ne ed", "need"],
  ["ne xt", "next"],
  ["on ly", "only"],
  ["ov er", "over"],
  ["pa rt", "part"],
  ["pe op", "people"], // partial
  ["sa id", "said"],
  ["sa me", "same"],
  ["su ch", "such"],
  ["ta ke", "take"],
  ["th em", "them"],
  ["th en", "then"],
  ["th er", "there"], // partial
  ["th es", "these"], // partial
  ["th in", "think"], // partial
  ["us ed", "used"],
  ["wa nt", "want"],
  ["we ll", "well"],
  ["we re", "were"],
  ["wh er", "where"], // partial
  ["wh ic", "which"], // partial
  ["wo ul", "would"], // partial
  // French common words
  ["l e", "le"],
  ["l a", "la"],
  ["u n", "un"],
  ["d e", "de"],
  ["e t", "et"],
  ["e n", "en"],
  ["à", "à"],  // keep as is
  ["d u", "du"],
  ["a u", "au"],
  ["o u", "ou"],
  ["q ue", "que"],
  ["qu e", "que"],
  ["p ar", "par"],
  ["pa r", "par"],
  ["s ur", "sur"],
  ["su r", "sur"],
  ["p our", "pour"],
  ["po ur", "pour"],
  ["pou r", "pour"],
  ["av ec", "avec"],
  ["ave c", "avec"],
  ["da ns", "dans"],
  ["dan s", "dans"],
  ["so nt", "sont"],
  ["son t", "sont"],
  ["ce tte", "cette"],
  ["cet te", "cette"],
  ["cett e", "cette"],
  ["pl us", "plus"],
  ["plu s", "plus"],
  ["au tre", "autre"],
  ["aut re", "autre"],
  ["autr e", "autre"],
  ["ma is", "mais"],
  ["mai s", "mais"],
  ["to ut", "tout"],
  ["tou t", "tout"],
  ["to us", "tous"],
  ["tou s", "tous"],
  ["bi en", "bien"],
  ["bie n", "bien"],
  ["sa ns", "sans"],
  ["san s", "sans"],
  ["mê me", "même"],
  ["mêm e", "même"],
  ["en tre", "entre"],
  ["ent re", "entre"],
  ["entr e", "entre"],
]);

function detectShortFragmentedWords(text) {
  const out = [];
  const t = String(text || "");
  const tLower = t.toLowerCase();

  for (const [fragment, correction] of SHORT_COMMON_WORDS) {
    const fragLower = fragment.toLowerCase();
    let searchStart = 0;

    while (true) {
      const idx = tLower.indexOf(fragLower, searchStart);
      if (idx === -1) break;

      // Get actual text from original (preserve case for context)
      const actualFragment = t.slice(idx, idx + fragment.length);
      
      // Verify boundaries: should be at word boundaries or surrounded by spaces/punctuation
      const prevChar = idx > 0 ? t[idx - 1] : " ";
      const nextChar = idx + fragment.length < t.length ? t[idx + fragment.length] : " ";
      
      const prevIsWord = /[A-Za-zÀ-ÖØ-öø-ÿ]/.test(prevChar);
      const nextIsWord = /[A-Za-zÀ-ÖØ-öø-ÿ]/.test(nextChar);
      
      // Only flag if it's truly isolated (not part of a larger word)
      if (!prevIsWord && !nextIsWord) {
        const ctx = computeContextFromFullText(t, idx, fragment.length);
        
        out.push({
          error: actualFragment,
          correction: correction,
          type: "spacing",
          severity: "high",
          message: `Fragmented word: "${fragment}" → "${correction}"`,
          contextBefore: ctx.contextBefore,
          contextAfter: ctx.contextAfter,
          startIndex: idx,
          endIndex: idx + fragment.length,
        });
      }

      searchStart = idx + 1;
      if (out.length >= 100) break;
    }

    if (out.length >= 100) break;
  }

  return out;
}

// ============================================================
// Context scoring for accurate position mapping
// ============================================================

function scoreContext(allText, idx, errObj) {
  const e = String(errObj?.error || "");
  if (!e) return -999;

  const before = normalizeSpaces(errObj?.contextBefore || "");
  const after = normalizeSpaces(errObj?.contextAfter || "");

  const winBefore = normalizeSpaces(allText.slice(Math.max(0, idx - 60), idx));
  const winAfter = normalizeSpaces(allText.slice(idx + e.length, idx + e.length + 60));

  let score = 0;

  if (before && winBefore.endsWith(before)) score += 6;
  if (after && winAfter.startsWith(after)) score += 6;

  if (allText.slice(idx, idx + e.length) === e) score += 3;

  const prev = idx > 0 ? allText[idx - 1] : " ";
  const next = idx + e.length < allText.length ? allText[idx + e.length] : " ";
  if (!/[A-Za-zÀ-ÖØ-öø-ÿ]/.test(prev) || !/[A-Za-zÀ-ÖØ-öø-ÿ]/.test(next)) score += 1;

  return score;
}

function findBestOccurrence(allText, errObj) {
  const e = String(errObj?.error || "");
  if (!e) return -1;

  const lcText = allText.toLowerCase();
  const lcErr = e.toLowerCase();

  const positions = [];
  let i = 0;
  while (true) {
    const idx = lcText.indexOf(lcErr, i);
    if (idx === -1) break;
    positions.push(idx);
    i = idx + Math.max(1, lcErr.length);
  }
  if (!positions.length) return -1;

  let bestIdx = positions[0];
  let bestScore = -999;

  for (const p of positions) {
    const sc = scoreContext(allText, p, errObj);
    if (sc > bestScore) {
      bestScore = sc;
      bestIdx = p;
    }
  }

  const hasContext =
    (errObj?.contextBefore && String(errObj.contextBefore).trim()) ||
    (errObj?.contextAfter && String(errObj.contextAfter).trim());

  if (hasContext && bestScore <= 0) return -1;

  return bestIdx;
}

// ============================================================
// OpenAI API call
// ============================================================

async function callAIChunk(apiKey, chunkText) {
  const resp = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      temperature: 0.0,
      max_tokens: 2500,
      messages: [
        { role: "system", content: PROMPT },
        { role: "user", content: `Proofread this text and return JSON array of errors found:\n\n${chunkText}` },
      ],
    }),
  });

  if (!resp.ok) {
    console.error("[AI] API Error:", resp.status);
    return [];
  }

  const data = await resp.json();
  const content = data?.choices?.[0]?.message?.content || "[]";
  return safeParseJSONArray(content);
}

// ============================================================
// Post-processing validation
// ============================================================

function isValidCorrection(error, correction) {
  const e = String(error || "").trim();
  const c = String(correction || "").trim();
  
  if (!e || !c) return false;
  if (e.toLowerCase() === c.toLowerCase()) return false;
  
  // Reject if correction is way longer (likely rewrite)
  if (c.length > e.length * 2 + 10) return false;
  
  // Reject if error is multiple capitalized words (likely names)
  const words = e.split(/\s+/);
  const allCapitalized = words.every(w => /^[A-Z][a-z]*$/.test(w));
  if (allCapitalized && words.length >= 2) return false;
  
  return true;
}

// ============================================================
// Main export
// ============================================================

export async function checkSpellingWithAI(text) {
  console.log("🧠 AI proofread starting (VERSION 5.1 - AI + micro-deterministic)...");

  const fullRaw = String(text || "");
  if (fullRaw.trim().length < 10) {
    console.log("[AI] Text too short, skipping.");
    return [];
  }

  // 1) MICRO-DETERMINISTIC: catch short fragmented words (o f, i n, t o, etc.)
  const microDet = detectShortFragmentedWords(fullRaw);
  console.log(`[MICRO] Found ${microDet.length} short fragmented word(s)`);

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.log("[AI] No OPENAI_API_KEY configured, returning micro-deterministic only.");
    return addAnchors(microDet);
  }

  // 2) AI DETECTION: mask noise and process chunks
  const aiView = maskNoiseKeepLength(fullRaw);
  const chunks = chunkText(aiView, 8000, 450);
  console.log(`[AI] Chunking: ${chunks.length} chunk(s), total chars=${aiView.length}`);

  const collected = [];

  for (let ci = 0; ci < chunks.length; ci++) {
    const ch = chunks[ci];
    console.log(`[AI] Chunk ${ci + 1}/${chunks.length} (${ch.start}-${ch.end})`);

    const raw = await callAIChunk(apiKey, ch.text);
    console.log(`[AI] Chunk ${ci + 1} returned ${raw.length} issue(s)`);

    for (const x of raw) {
      const e = normalizeSpaces(x?.error);
      const c = normalizeSpaces(x?.correction);

      if (!isValidCorrection(e, c)) continue;
      if (e.length > 120) continue;
      if (isProbablyGarbageSpan(e) || isProbablyGarbageSpan(c)) continue;

      const candidate = {
        error: e,
        correction: c,
        type: x?.type || "spelling",
        severity: x?.severity || "medium",
        message: x?.message || "AI-detected issue",
        contextBefore: x?.contextBefore || "",
        contextAfter: x?.contextAfter || "",
      };

      const bestIdx = findBestOccurrence(fullRaw, candidate);

      let contextBefore = candidate.contextBefore;
      let contextAfter = candidate.contextAfter;
      let startIndex = null;
      let endIndex = null;

      if (bestIdx >= 0) {
        const ctx = computeContextFromFullText(fullRaw, bestIdx, e.length);
        contextBefore = ctx.contextBefore;
        contextAfter = ctx.contextAfter;
        startIndex = bestIdx;
        endIndex = bestIdx + e.length;
      }

      collected.push({
        error: e,
        correction: c,
        type: candidate.type,
        severity: candidate.severity,
        message: candidate.message,
        contextBefore,
        contextAfter,
        startIndex,
        endIndex,
      });
    }
  }

  // 3) MERGE: micro-deterministic + AI results
  const merged = [...microDet, ...collected];

  // 4) DEDUPLICATE by error+correction+context
  const seen = new Set();
  const deduped = [];
  for (const it of merged) {
    const key = [
      String(it.error || "").toLowerCase(),
      String(it.correction || "").toLowerCase(),
      normalizeContext(it.contextBefore).toLowerCase().slice(0, 30),
      normalizeContext(it.contextAfter).toLowerCase().slice(0, 30),
    ].join("||");
    if (seen.has(key)) continue;
    seen.add(key);
    deduped.push(it);
  }

  const out = addAnchors(deduped);

  console.log(`✅ AI proofread complete: ${out.length} issues (${microDet.length} micro + ${collected.length} AI)`);
  if (out.length > 0) {
    console.log("First 10 issues:");
    out.slice(0, 10).forEach((x, i) => {
      console.log(`  ${i + 1}. "${x.error}" → "${x.correction}" (${x.type})`);
    });
  }

  return out;
}

export default { checkSpellingWithAI };
