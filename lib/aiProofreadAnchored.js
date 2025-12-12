// lib/aiProofreadAnchored.js
// VERSION 5.0 — 100% AI detection (human-like proofreading)
// ✅ No regex/deterministic false positives
// ✅ Conservative AI prompt
// ✅ Context-scored mapping back to RAW text
// ✅ Noise masking keeps string length (safe indexing)

const PROMPT = `
You are an expert human proofreader reviewing a professional document.

Your task: find REAL spelling, grammar, and formatting errors that a professional editor would catch.

DETECT ONLY:
- Genuine spelling mistakes (typos, misspellings)
- Grammar errors (subject-verb agreement, wrong tense, conjugation)
- Punctuation errors (missing periods, wrong commas, extra punctuation)
- Double spaces (two or more consecutive spaces)
- Words clearly broken by a space ONLY if merging creates a REAL word:
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
  // Reject if looks like code/path
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

  // exact match boost
  if (allText.slice(idx, idx + e.length) === e) score += 3;

  // mild boost if boundaries look word-ish (reduce false positives)
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

  // if context provided but we couldn't score it at all, reject
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
  
  // Reject if correction is longer than error by too much (likely rewrite)
  if (c.length > e.length * 2 + 10) return false;
  
  // Reject if error contains only valid capitalized words (likely names)
  const words = e.split(/\s+/);
  const allCapitalized = words.every(w => /^[A-Z][a-z]*$/.test(w));
  if (allCapitalized && words.length >= 2) return false;
  
  return true;
}

// ============================================================
// Main export
// ============================================================

export async function checkSpellingWithAI(text) {
  console.log("🧠 AI proofread starting (100% AI, no regex)...");

  const fullRaw = String(text || "");
  if (fullRaw.trim().length < 10) {
    console.log("[AI] Text too short, skipping.");
    return [];
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.log("[AI] No OPENAI_API_KEY configured, cannot proofread.");
    return [];
  }

  // Mask noise (numbers, codes, emails, URLs) but keep string length
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

      // Validation filters
      if (!isValidCorrection(e, c)) continue;
      if (e.length > 120) continue;
      if (isProbablyGarbageSpan(e) || isProbablyGarbageSpan(c)) continue;

      // Map back to RAW text via context-scored search
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

  // Deduplicate by error+correction+context
  const seen = new Set();
  const deduped = [];
  for (const it of collected) {
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

  console.log(`✅ AI proofread complete: ${out.length} issues`);
  if (out.length > 0) {
    console.log("First 10 issues:");
    out.slice(0, 10).forEach((x, i) => {
      console.log(`  ${i + 1}. "${x.error}" → "${x.correction}" (${x.type})`);
    });
  }

  return out;
}

export default { checkSpellingWithAI };
