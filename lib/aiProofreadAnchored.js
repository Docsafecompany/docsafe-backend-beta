// lib/aiProofreadAnchored.js
// VERSION 3.3 — AI + deterministic sweep (anchored, chunked, dedup, context-safe)

const PROMPT = `
You are an expert proofreader.

You must detect REAL issues only:
- spelling
- grammar (conjugation, agreement)
- punctuation
- spacing issues (double spaces, space inside a word)
- corrupted words (e.g. "correctionnns" -> "corrections")

STRICT:
1) Do NOT rewrite sentences. Only fix minimal spans.
2) Do NOT merge valid separate words (e.g., "a message" must remain two words).
3) If a word is clearly corrupted by an extra space INSIDE the word ("soc ial"), fix it.
4) Output must be JSON only.

Return JSON array:
[
  {
    "error": "exact substring to replace (shortest possible)",
    "correction": "replacement",
    "type": "spelling|grammar|punctuation|spacing|fragmented_word|typo|capitalization",
    "severity": "high|medium|low",
    "message": "short reason",
    "contextBefore": "up to 40 chars before the error",
    "contextAfter": "up to 40 chars after the error"
  }
]
`;

function safeParseJSONArray(raw) {
  if (!raw) return [];
  const m = raw.match(/\[[\s\S]*\]/);
  if (!m) return [];
  try { return JSON.parse(m[0]); } catch { return []; }
}

function normalize(s) {
  return (s || "").replace(/\s+/g, " ").trim();
}

function mkId() {
  return `spell_${Date.now()}_${Math.random().toString(16).slice(2)}`;
}

function normalizeContext(s) {
  return (s || "").replace(/\s+/g, " ").trim().slice(0, 80);
}

function addAnchors(issues) {
  return (issues || []).map((it) => ({
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
  if (/[<>]/.test(x) && /(class=|data-|style=|<\w+|<\/\w+)/i.test(x)) return true;
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

// -------------------------
// Deterministic spacing sweep
// catches: "so cial" / "searc r h" / "enablin g" etc.
// without merging valid words.
// -------------------------

function isLetter(ch) {
  return /[A-Za-zÀ-ÖØ-öø-ÿ]/.test(ch || "");
}

function makeSpacingFixes(fullText) {
  const text = String(fullText || "");
  const out = [];

  // Patterns like: letter(s) + space(s) + letter(s) INSIDE a word-like sequence
  // We keep it conservative:
  // - only when BOTH sides are short fragments (1-3 chars) OR one side is 1 char,
  // - and surrounding chars are letters (to avoid "a message")
  const re = /([A-Za-zÀ-ÖØ-öø-ÿ]{1,3})\s{1,3}([A-Za-zÀ-ÖØ-öø-ÿ]{1,3})/g;

  let m;
  while ((m = re.exec(text)) !== null) {
    const error = m[0];
    const left = m[1];
    const right = m[2];

    const start = m.index;
    const end = start + error.length;

    const prev = start > 0 ? text[start - 1] : "";
    const next = end < text.length ? text[end] : "";

    // Must be inside a "word run": either previous or next char is a letter
    const looksWordish = isLetter(prev) || isLetter(next);

    // Conservative: avoid merging valid words like "to do" (2+2) or "a message"
    const riskyTwoWords = left.length >= 2 && right.length >= 2 && !looksWordish;
    if (riskyTwoWords) continue;

    // Only accept if it looks like fragmentation:
    // - at least one side is 1 char, OR both sides <= 2 chars, OR looksWordish strongly
    const likelyFragment =
      left.length === 1 ||
      right.length === 1 ||
      (left.length <= 2 && right.length <= 2) ||
      looksWordish;

    if (!likelyFragment) continue;

    const correction = (left + right);
    if (correction.toLowerCase() === error.toLowerCase()) continue;

    const ctx = computeContextFromFullText(text, start, error.length);

    out.push({
      error,
      correction,
      type: "spacing",
      severity: "high",
      message: "Fragmented word (space inside word)",
      contextBefore: ctx.contextBefore,
      contextAfter: ctx.contextAfter,
      startIndex: start,
      endIndex: start + error.length,
    });

    // safety limit
    if (out.length >= 200) break;
  }

  return out;
}

// -------------------------
// OpenAI call
// -------------------------

async function callAIChunk(apiKey, chunkText) {
  const resp = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      temperature: 0.1,
      max_tokens: 2000,
      messages: [
        { role: "system", content: PROMPT },
        { role: "user", content: `Text:\n${chunkText}\n\nReturn JSON only.` }
      ]
    })
  });

  if (!resp.ok) {
    console.error("[AI] Error:", resp.status);
    return [];
  }

  const data = await resp.json();
  const content = data?.choices?.[0]?.message?.content || "[]";
  return safeParseJSONArray(content);
}

export async function checkSpellingWithAI(text) {
  console.log("🧠 AI proofread (anchored+chunked) starting...");

  const full = String(text || "");
  if (full.trim().length < 10) return [];

  // 1) deterministic spacing sweep first (cheap + catches most misses)
  const spacingFixes = makeSpacingFixes(full);

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.log("[AI] No API key, returning deterministic fixes only.");
    const out = addAnchors(spacingFixes);
    console.log(`✅ Proofread complete: ${out.length} issue(s) (deterministic only)`);
    return out;
  }

  // 2) AI pass (chunked)
  const chunks = chunkText(full, 8000, 450);
  console.log(`[AI] Chunking: ${chunks.length} chunk(s), total chars=${full.length}`);

  const collected = [];

  for (let ci = 0; ci < chunks.length; ci++) {
    const ch = chunks[ci];
    console.log(`[AI] Chunk ${ci + 1}/${chunks.length} (${ch.start}-${ch.end})`);

    const raw = await callAIChunk(apiKey, ch.text);

    for (const x of raw) {
      const e = normalize(x?.error);
      const c = normalize(x?.correction);

      if (!e || !c) continue;
      if (e.toLowerCase() === c.toLowerCase()) continue;
      if (e.length > 120) continue;
      if (isProbablyGarbageSpan(e) || isProbablyGarbageSpan(c)) continue;

      const localIdx = ch.text.toLowerCase().indexOf(e.toLowerCase());
      const globalIdx = localIdx >= 0 ? (ch.start + localIdx) : null;

      let contextBefore = x?.contextBefore || "";
      let contextAfter = x?.contextAfter || "";
      if (typeof globalIdx === "number") {
        const ctx = computeContextFromFullText(full, globalIdx, e.length);
        contextBefore = ctx.contextBefore;
        contextAfter = ctx.contextAfter;
      }

      collected.push({
        error: e,
        correction: c,
        type: x?.type || "spelling",
        severity: x?.severity || "medium",
        message: x?.message || "AI-detected issue",
        contextBefore,
        contextAfter,
        startIndex: typeof globalIdx === "number" ? globalIdx : null,
        endIndex: typeof globalIdx === "number" ? (globalIdx + e.length) : null,
      });
    }
  }

  // 3) merge deterministic + AI
  const merged = [...spacingFixes, ...collected];

  // 4) Dedup: same error/correction + similar context
  const seen = new Set();
  const deduped = [];
  for (const it of merged) {
    const key = [
      String(it.error || "").toLowerCase(),
      String(it.correction || "").toLowerCase(),
      normalizeContext(it.contextBefore).toLowerCase(),
      normalizeContext(it.contextAfter).toLowerCase(),
    ].join("||");

    if (seen.has(key)) continue;
    seen.add(key);
    deduped.push(it);
  }

  const out = addAnchors(deduped);

  console.log(`✅ AI proofread complete: ${out.length} issues (deduped)`);
  out.slice(0, 10).forEach((x, i) => {
    console.log(`  ${i + 1}. "${x.error}" -> "${x.correction}" (${x.type})`);
  });

  return out;
}

export default { checkSpellingWithAI };
