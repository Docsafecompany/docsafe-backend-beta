// lib/aiProofreadAnchored.js
// AI-first detection with anchors to apply safely later

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

function addAnchors(issues) {
  return (issues || []).map((it, i) => ({
    id: `spell_${i}_${Date.now()}`,
    error: it.error,
    correction: it.correction,
    type: it.type || "spelling",
    severity: it.severity || "medium",
    message: it.message || "AI-detected issue",
    contextBefore: it.contextBefore || "",
    contextAfter: it.contextAfter || ""
  }));
}

// IMPORTANT: no regex-merge filtering here. Let AI decide. We apply safely with anchors later.
export async function checkSpellingWithAI(text) {
  console.log("🧠 AI proofread (anchored) starting...");

  const t = (text || "").slice(0, 20000);
  if (!t || t.length < 10) return [];

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.log("[AI] No API key, returning empty.");
    return [];
  }

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
        { role: "user", content:
`Text:
${t}

Return JSON only. Include contextBefore/contextAfter for each fix.`
        }
      ]
    })
  });

  if (!resp.ok) {
    console.error("[AI] Error:", resp.status);
    return [];
  }

  const data = await resp.json();
  const content = data?.choices?.[0]?.message?.content || "[]";
  const raw = safeParseJSONArray(content);

  // basic sanity: must change, must be short
  const filtered = raw.filter(x => {
    const e = normalize(x?.error);
    const c = normalize(x?.correction);
    if (!e || !c) return false;
    if (e.toLowerCase() === c.toLowerCase()) return false;
    if (e.length > 120) return false;
    return true;
  });

  const out = addAnchors(filtered);

  console.log(`✅ AI proofread complete: ${out.length} issues`);
  out.slice(0, 10).forEach((x, i) => {
    console.log(`  ${i + 1}. "${x.error}" -> "${x.correction}" (${x.type})`);
  });

  return out;
}

export default { checkSpellingWithAI };
