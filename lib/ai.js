// lib/ai.js
// Appels OpenAI via fetch (Node 18+). Deux fonctions: aiCorrectText (V1) et aiRephraseText (V2).
// Prompts renforcés pour joindre les mots scindés par un espace à l’intérieur d’un mot.

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";

async function callOpenAI({ system, user, model = "gpt-4o-mini", temperature = 0 }) {
  if (!OPENAI_API_KEY) return null; // fallback côté serveur si pas de clé
  const resp = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${OPENAI_API_KEY}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model,
      temperature,
      messages: [
        { role: "system", content: system },
        { role: "user", content: user }
      ]
    })
  });
  if (!resp.ok) {
    const text = await resp.text().catch(() => "");
    throw new Error(`OpenAI HTTP ${resp.status}: ${text || resp.statusText}`);
  }
  const data = await resp.json();
  return String(data?.choices?.[0]?.message?.content ?? "");
}

// ---------- V1: correction ----------
export async function aiCorrectText(text) {
  const system = [
    "You are a precise multilingual text corrector.",
    "Auto-detect the input language and write the output in the same language.",
    "Correct spelling, diacritics, grammar, spacing, and punctuation.",
    "VERY IMPORTANT: If a single word is split by an internal space between letters (e.g., 'soc ial', 'enablin g', 'th e', 'corpo, rations' -> 'corporations'), MERGE it into the correct single word.",
    "Do not merge two DISTINCT words (e.g., 'social media' must stay two words).",
    "Preserve all line breaks and list/heading structure.",
    "Do NOT paraphrase: only corrections.",
    "Return ONLY the corrected text."
  ].join(" ");
  return await callOpenAI({ system, user: String(text), temperature: 0 });
}

// ---------- V2: reformulation ----------
export async function aiRephraseText(text) {
  const system = [
    "You are a careful multilingual text rewriter.",
    "Auto-detect the input language and return in the same language.",
    "Rewrite to be clearer and more fluent while preserving meaning and tone.",
    "ALSO fix spelling, punctuation, spacing and MERGE any internally split words ('soc ial' -> 'social', 'enablin g' -> 'enabling', 'th e' -> 'the').",
    "Do not add new facts or remove information.",
    "Keep headings and bullet/numbered lists; preserve line breaks.",
    "Return ONLY the rewritten text."
  ].join(" ");
  return await callOpenAI({ system, user: String(text), temperature: 0.3 });
}

