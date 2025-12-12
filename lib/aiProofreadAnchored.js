// lib/aiProofreadAnchored.js
// AI-first detection with anchors to apply safely later (chunked + dedup + context-safe)

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
    message: it.message || "AI-detected issue",
    contextBefore: it.contextBefore || "",
    contextAfter: it.contextAfter || "",
    // optional fields (safe to ignore by frontend)
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
  // block obvious HTML/marker junk
  if (/[<>]/.test(x) && /(class=|data-|style=|<\w+|<\/\w+)/i.test(x)) return true;
  return false;
}

function computeContextFromFullText(fullText, globalIndex, errorLen) {
  const before = fullText.slice(Math.max(0, globalIndex - 40), globalIndex);
  const after = fullText.slice(globalIndex + errorLen, Math.min(fullText.length, globalIndex + errorLen + 40));
  return { contextBefore: before, contextAfter: after };
}

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

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.log("[AI] No API key, returning empty.");
    return [];
  }

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

      // Try to map index inside chunk -> global index
      const localIdx = ch.text.toLowerCase().indexOf(e.toLowerCase());
      let globalIdx = localIdx >= 0 ? (ch.start + localIdx) : null;

      // If not found, try a normalized search (helps with double spaces)
      if (globalIdx === null) {
        const nChunk = normalize(ch.text).toLowerCase();
        const nErr = normalize(e).toLowerCase();
        const nIdx = nChunk.indexOf(nErr);
        if (nIdx >= 0) {
          // no perfect mapping back; keep null but keep contexts
          globalIdx = null;
        }
      }

      // Force anchors to be based on REAL fullText contexts when possible
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
        contextBefore: contextBefore,
        contextAfter: contextAfter,
        startIndex: typeof globalIdx === "number" ? globalIdx : null,
        endIndex: typeof globalIdx === "number" ? (globalIdx + e.length) : null,
      });
    }
  }

  // Dedup: same error/correction + similar context
  const seen = new Set();
  const deduped = [];
  for (const it of collected) {
    const key = [
      it.error.toLowerCase(),
      it.correction.toLowerCase(),
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
