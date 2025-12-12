// lib/aiProofreadAnchored.js
// VERSION 4.0 — AI + deterministic sweep (anchors robust, detects missing-space, inside-word spacing, punctuation corruption)

const PROMPT = `
You are an expert proofreader.

Detect REAL issues only:
- spelling
- grammar (agreement, conjugation)
- punctuation
- spacing issues:
  - double spaces
  - space INSIDE a word ("soc ial" -> "social")
  - missing space between words when clearly stuck together ("suchas" -> "such as", "andX" -> "and X", "Impactof" -> "Impact of")
- corrupted words:
  - punctuation inside word ("corpo,rations" -> "corporations")
  - repeated letters that clearly corrupt a word ("gggdigital" -> "digital" or "gdigital" -> "digital")

STRICT:
1) Do NOT rewrite sentences. Only fix minimal spans.
2) Do NOT change meaning.
3) Do NOT split or merge valid words unless the span is clearly corrupted.
4) Output must be JSON only. No markdown.

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
  if (x.length > 140) return true;
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
// Pre-clean to reduce noise (numbers/codes) so AI focuses on words
// -------------------------
function reduceNoiseForProofread(fullText) {
  const t = String(fullText || "");
  // keep punctuation/spaces, but drop long numeric blocks & excel-like cell refs noise
  return t
    .replace(/\b\d{6,}\b/g, " ")                 // long numbers
    .replace(/\b[A-Z]{1,3}\d{2,5}\b/g, " ")      // like A123 / BC2345
    .replace(/\s+/g, " ")
    .trim();
}

// -------------------------
// Deterministic pass #1: space INSIDE a word
// catches: "soc ial" / "searc r h" / "enablin g" etc.
// -------------------------
const LETTER = /[A-Za-zÀ-ÖØ-öø-ÿ]/;
function isLetter(ch) { return LETTER.test(ch || ""); }

function deterministicInsideWordSpaces(text) {
  const out = [];
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
    const looksWordish = isLetter(prev) || isLetter(next);

    // conservative
    const likelyFragment =
      left.length === 1 ||
      right.length === 1 ||
      (left.length <= 2 && right.length <= 2) ||
      looksWordish;

    if (!likelyFragment) continue;

    const correction = left + right;
    const ctx = computeContextFromFullText(text, start, error.length);

    out.push({
      error,
      correction,
      type: "spacing",
      severity: "high",
      message: "Space inside a word (fragmented word)",
      contextBefore: ctx.contextBefore,
      contextAfter: ctx.contextAfter,
      startIndex: start,
      endIndex: end,
    });

    if (out.length >= 250) break;
  }

  return out;
}

// -------------------------
// Deterministic pass #2: punctuation inside word
// catches: "corpo,rations" / "comm,ents" etc.
// -------------------------
function deterministicPunctInsideWord(text) {
  const out = [];
  const re = /([A-Za-zÀ-ÖØ-öø-ÿ]{2,})[,\.;:']([A-Za-zÀ-ÖØ-öø-ÿ]{2,})/g;

  let m;
  while ((m = re.exec(text)) !== null) {
    const error = m[0];
    const correction = m[1] + m[2];
    const start = m.index;
    const ctx = computeContextFromFullText(text, start, error.length);

    out.push({
      error,
      correction,
      type: "typo",
      severity: "high",
      message: "Punctuation inside a word (corrupted token)",
      contextBefore: ctx.contextBefore,
      contextAfter: ctx.contextAfter,
      startIndex: start,
      endIndex: start + error.length,
    });

    if (out.length >= 250) break;
  }

  return out;
}

// -------------------------
// Deterministic pass #3: missing spaces for very common stuck connectors
// catches: suchas, andX, Impactof, Mediaon, individualsto, etc.
// -------------------------
function deterministicMissingSpaces(text) {
  const out = [];

  // Common connectors you WANT to split when stuck (case-insensitive)
  const connectors = [
    "as", "of", "to", "in", "on", "at", "by", "for", "from", "with",
    "and", "or", "the", "a", "an", "is", "are", "was", "were"
  ];

  // Strategy:
  // - find long alphabetic token
  // - if token contains a connector sandwiched between letters: e.g. "suchas" => "such as"
  // - avoid very short tokens
  const tokenRe = /\b[A-Za-zÀ-ÖØ-öø-ÿ]{5,30}\b/g;

  let m;
  while ((m = tokenRe.exec(text)) !== null) {
    const token = m[0];
    const start = m.index;

    const low = token.toLowerCase();

    for (const conn of connectors) {
      // connector must not be at start; must not be at end; and must be surrounded by letters
      const idx = low.indexOf(conn);
      if (idx <= 1) continue;
      if (idx + conn.length >= low.length - 1) continue;

      const beforeCh = token[idx - 1];
      const afterCh = token[idx + conn.length];

      if (!isLetter(beforeCh) || !isLetter(afterCh)) continue;

      // Guard: avoid splitting if both sides are 1-2 chars (too risky)
      const left = token.slice(0, idx);
      const right = token.slice(idx + conn.length);
      if (left.length <= 2 && right.length <= 2) continue;

      const correction = token.slice(0, idx) + " " + token.slice(idx);

      const ctx = computeContextFromFullText(text, start, token.length);

      out.push({
        error: token,
        correction,
        type: "spacing",
        severity: "high",
        message: "Missing space between words (stuck connector)",
        contextBefore: ctx.contextBefore,
        contextAfter: ctx.contextAfter,
        startIndex: start,
        endIndex: start + token.length,
      });

      break; // only one split per token (conservative)
    }

    if (out.length >= 250) break;
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
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: "gpt-4o-mini",
      temperature: 0.0,
      max_tokens: 2000,
      messages: [
        { role: "system", content: PROMPT },
        { role: "user", content: `Text:\n${chunkText}\n\nReturn JSON only.` },
      ],
    }),
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

  const fullRaw = String(text || "");
  const full = reduceNoiseForProofread(fullRaw);
  if (full.trim().length < 10) return [];

  // 1) deterministic passes first
  const det1 = deterministicInsideWordSpaces(fullRaw);
  const det2 = deterministicPunctInsideWord(fullRaw);
  const det3 = deterministicMissingSpaces(fullRaw);
  const deterministic = [...det1, ...det2, ...det3];

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.log("[AI] No API key, returning deterministic fixes only.");
    const out = addAnchors(deterministic);
    console.log(`✅ Proofread complete: ${out.length} issue(s) (deterministic only)`);
    return out;
  }

  // 2) AI pass (on reduced-noise text)
  const chunks = chunkText(full, 8000, 450);
  console.log(`[AI] Chunking: ${chunks.length} chunk(s), total chars=${full.length}`);

  const collected = [];
  for (let ci = 0; ci < chunks.length; ci++) {
    const ch = chunks[ci];
    console.log(`[AI] Chunk ${ci + 1}/${chunks.length} (${ch.start}-${ch.end})`);

    const raw = await callAIChunk(apiKey, ch.text);

    for (const x of raw) {
      const e = normalizeSpaces(x?.error);
      const c = normalizeSpaces(x?.correction);

      if (!e || !c) continue;
      if (e.toLowerCase() === c.toLowerCase()) continue;
      if (e.length > 140) continue;
      if (isProbablyGarbageSpan(e) || isProbablyGarbageSpan(c)) continue;

      // Map index: search in FULL RAW (not reduced) to apply later safely
      const localIdx = fullRaw.toLowerCase().indexOf(e.toLowerCase());
      const globalIdx = localIdx >= 0 ? localIdx : null;

      let contextBefore = x?.contextBefore || "";
      let contextAfter = x?.contextAfter || "";
      if (typeof globalIdx === "number") {
        const ctx = computeContextFromFullText(fullRaw, globalIdx, e.length);
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
        endIndex: typeof globalIdx === "number" ? globalIdx + e.length : null,
      });
    }
  }

  // 3) merge deterministic + AI
  const merged = [...deterministic, ...collected];

  // 4) dedup
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
