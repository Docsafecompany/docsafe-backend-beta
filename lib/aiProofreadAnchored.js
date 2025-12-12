// lib/aiProofreadAnchored.js
// VERSION 4.1 — Robust detection + robust anchoring
// ✅ AI + deterministic sweep
// ✅ Context-scored mapping back to RAW text
// ✅ Missing-space conservative (camelCase + stuck connectors + stoplist)
// ✅ Noise masking keeps string length (safe-ish indexing)

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
  - repeated letters that clearly corrupt a word ("gggdigital" -> "digital")

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
  if (x.length > 160) return true;
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

// ============================================================
// Noise masking for AI (keeps LENGTH) so anchors/indexing remain plausible
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

  return out;
}

// ============================================================
// Context scoring (same spirit as officeCorrect)
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
// Deterministic passes
// ============================================================

const LETTER = /[A-Za-zÀ-ÖØ-öø-ÿ]/;
function isLetter(ch) { return LETTER.test(ch || ""); }

// 1) space INSIDE a word: "soc ial" / "enablin g"
function deterministicInsideWordSpaces(text) {
  const out = [];
  const t = String(text || "");
  const re = /([A-Za-zÀ-ÖØ-öø-ÿ]{1,3})\s{1,3}([A-Za-zÀ-ÖØ-öø-ÿ]{1,3})/g;

  let m;
  while ((m = re.exec(t)) !== null) {
    const error = m[0];
    const left = m[1];
    const right = m[2];

    const start = m.index;
    const end = start + error.length;

    const prev = start > 0 ? t[start - 1] : "";
    const next = end < t.length ? t[end] : "";
    const looksWordish = isLetter(prev) || isLetter(next);

    // conservative fragmentation signal
    const likelyFragment =
      left.length === 1 ||
      right.length === 1 ||
      (left.length <= 2 && right.length <= 2) ||
      looksWordish;

    if (!likelyFragment) continue;

    const correction = left + right;
    const ctx = computeContextFromFullText(t, start, error.length);

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

// 2) punctuation inside word: "corpo,rations"
function deterministicPunctInsideWord(text) {
  const out = [];
  const t = String(text || "");
  const re = /([A-Za-zÀ-ÖØ-öø-ÿ]{2,})[,\.;:']([A-Za-zÀ-ÖØ-öø-ÿ]{2,})/g;

  let m;
  while ((m = re.exec(t)) !== null) {
    const error = m[0];
    const correction = m[1] + m[2];
    const start = m.index;

    const ctx = computeContextFromFullText(t, start, error.length);
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

// 3) missing spaces (CONSERVATIVE):
// - camelCase boundary: "andX" => "and X"
// - stuck connector: "suchas" => "such as", "Impactof" => "Impact of"
// + stoplist to avoid therefore/before/after/etc.
function deterministicMissingSpaces(text) {
  const out = [];
  const t = String(text || "");

  const STOP = new Set([
    "therefore", "before", "after", "whereas", "moreover", "however", "whatever",
    "without", "within", "another", "together", "everywhere"
  ]);

  // A) camelCase boundary
  const camelRe = /\b([A-Za-zÀ-ÖØ-öø-ÿ]{2,})([A-Z][a-z])\b/g;
  let m;
  while ((m = camelRe.exec(t)) !== null) {
    const token = m[0];
    const start = m.index;

    if (STOP.has(token.toLowerCase())) continue;

    const correction = token.replace(/([a-z])([A-Z])/g, "$1 $2");
    if (correction === token) continue;

    const ctx = computeContextFromFullText(t, start, token.length);
    out.push({
      error: token,
      correction,
      type: "spacing",
      severity: "high",
      message: "Missing space (camelCase stuck words)",
      contextBefore: ctx.contextBefore,
      contextAfter: ctx.contextAfter,
      startIndex: start,
      endIndex: start + token.length,
    });

    if (out.length >= 200) break;
  }

  if (out.length >= 200) return out;

  // B) stuck connectors inside long tokens (one split max)
  const connectors = ["as", "of", "to", "in", "on", "and", "the"];
  const tokenRe = /\b[A-Za-zÀ-ÖØ-öø-ÿ]{5,30}\b/g;

  while ((m = tokenRe.exec(t)) !== null) {
    const token = m[0];
    const start = m.index;

    const low = token.toLowerCase();
    if (STOP.has(low)) continue;

    // skip tokens that are ALL caps (often acronyms)
    if (/^[A-Z]{5,}$/.test(token)) continue;

    for (const conn of connectors) {
      const idx = low.indexOf(conn);
      if (idx <= 1) continue;
      if (idx + conn.length >= low.length - 1) continue;

      const left = token.slice(0, idx);
      const right = token.slice(idx);

      // conservative guardrails
      if (left.length < 3) continue;
      if (right.length < 3 && !/^[a-z]*[A-Z]$/.test(right)) continue; // allow "andX"
      if (!isLetter(token[idx - 1]) || !isLetter(token[idx + conn.length])) continue;

      const correction = left + " " + right;

      const ctx = computeContextFromFullText(t, start, token.length);
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

      break;
    }

    if (out.length >= 250) break;
  }

  return out;
}

// ============================================================
// OpenAI call
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

// ============================================================
// Main
// ============================================================

export async function checkSpellingWithAI(text) {
  console.log("🧠 AI proofread (anchored+chunked) starting...");

  const fullRaw = String(text || "");
  if (fullRaw.trim().length < 10) return [];

  // 1) deterministic on RAW (important!)
  const det = [
    ...deterministicInsideWordSpaces(fullRaw),
    ...deterministicPunctInsideWord(fullRaw),
    ...deterministicMissingSpaces(fullRaw),
  ];

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.log("[AI] No API key, returning deterministic fixes only.");
    const out = addAnchors(det);
    console.log(`✅ Proofread complete: ${out.length} issue(s) (deterministic only)`);
    return out;
  }

  // 2) AI on masked-noise text (same length as raw)
  const aiView = maskNoiseKeepLength(fullRaw);

  const chunks = chunkText(aiView, 8000, 450);
  console.log(`[AI] Chunking: ${chunks.length} chunk(s), total chars=${aiView.length}`);

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
      if (e.length > 160) continue;
      if (isProbablyGarbageSpan(e) || isProbablyGarbageSpan(c)) continue;

      // IMPORTANT: map back to RAW via context-scored search (not indexOf first match)
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

  // 3) merge deterministic + AI
  const merged = [...det, ...collected];

  // 4) dedup (error/correction + similar context)
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
