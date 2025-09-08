// lib/ai.js
// Appel direct à l'API OpenAI (chat.completions) via fetch (Node 18+).
// Deux modes : correction (V1) et reformulation (V2).
// Auto-détection de la langue; conservation des retours à la ligne; aucune fantaisie.

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";

async function callOpenAI({ system, user, model = "gpt-4o-mini", temperature = 0 }) {
  if (!OPENAI_API_KEY) {
    // Pas de clé => on renvoie null pour activer le fallback dans le serveur.
    return null;
  }
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
  const out = data?.choices?.[0]?.message?.content ?? "";
  return String(out);
}

// ------- Public API -------

// Correction (V1) — orthographe/ponctuation/espaces, sans changer le fond.
// Conserve retours à la ligne; ne colle pas les mots; ne reformule pas.
export async function aiCorrectText(text) {
  const system = [
    "You are a precise text corrector.",
    "Auto-detect the language (French/English/etc.).",
    "Fix typos, spelling, spacing, punctuation, and diacritics.",
    "Never concatenate separate words or split valid words.",
    "Preserve the original line breaks and list structure.",
    "Do NOT paraphrase: keep the same wording except for corrections.",
    "Return ONLY the corrected text, no explanations."
  ].join(" ");

  return await callOpenAI({
    system,
    user: text,
    temperature: 0
  });
}

// Reformulation (V2) — plus claire et fluide, sens conservé.
// Conserve les retours à la ligne et la structure (titres, listes).
export async function aiRephraseText(text) {
  const system = [
    "You are a careful text rewriter.",
    "Auto-detect the language (French/English/etc.).",
    "Rewrite to be clearer and more fluent while preserving meaning.",
    "Keep headings, bullet/numbered lists, and line breaks.",
    "Do not add new facts; do not remove information.",
    "Return ONLY the rewritten text, no explanations."
  ].join(" ");

  return await callOpenAI({
    system,
    user: text,
    temperature: 0.3
  });
}
