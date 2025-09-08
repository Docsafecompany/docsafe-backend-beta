// lib/ai.js
// IA = seule source de vérité pour la correction et la reformulation.
// Correction V1 : orthographe/grammaire/ponctuation/espaces + fusion des mots scindés + correction des non-mots.
// Reformulation V2 : idem V1 + réécriture plus claire, sens conservé.
// -> Modèle guidé avec consignes fortes + exemples pour corriger des tokens arbitraires (p. ex. "ggggggdigital" -> "digital").
// Aucune dépendance externe : usage direct de fetch vers OpenAI Chat Completions.

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";

// ------------------------ Core call ------------------------
async function callOpenAI({ system, user, model = "gpt-4o-mini", temperature = 0 }) {
  if (!OPENAI_API_KEY) return null; // pas de clé -> le serveur appliquera un fallback
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

// ------------------------ Prompts partagés ------------------------
const EXAMPLES_CORRECT = [
  // Français / Anglais — cas arbitraires, scissions internes, lettres répétées
  "soc ial  → social",
  "commu nication  → communication",
  "enablin g  → enabling",
  "th e  → the",
  "p otential  → potential",
  "rig hts  → rights",
  "gdigital  → digital",
  "gggdigital  → digital",
  "ggggggdigital  → digital",
  "corpo, rations  → corporations",
  "dis connection  → disconnection",
  "o f  → of",
  "c an  → can"
].join("\n");

function buildSystemCorrectionPrompt() {
  return [
    "You are a STRICT multilingual copy editor.",
    "Task: Return the corrected text ONLY, in the SAME language as input.",
    "Perform ALL of the following:",
    "1) Spelling/diacritics and grammar fixes.",
    "2) Punctuation and spacing normalization (no duplicate punctuation, proper spaces).",
    "3) Merge internally split words (letters separated by spaces or linebreaks) into ONE correct word.",
    "4) Replace NON-WORD tokens with the correct real word (e.g., 'gdigital', 'ggggggdigital' -> 'digital').",
    "5) NEVER paraphrase or change meaning; do not add/remove information.",
    "6) Preserve ALL line breaks, headings, lists, and structural markers.",
    "7) Keep separate words separate (e.g., 'social media' must remain two words).",
    "",
    "Examples of required behavior:\n" + EXAMPLES_CORRECT,
    "",
    "Output policy: return ONLY the corrected text, with the same line breaks."
  ].join("\n");
}

function buildSystemRephrasePrompt() {
  return [
    "You are a multilingual rewriter.",
    "Task: Return the rewritten text ONLY, in the SAME language as input.",
    "Perform ALL of the following:",
    "1) Make sentences clearer and more fluent while preserving meaning and tone.",
    "2) ALSO perform strict spelling/grammar/diacritics/punctuation/spacing corrections.",
    "3) Merge internally split words into ONE correct word.",
    "4) Replace NON-WORD tokens with the correct real word (e.g., 'gdigital', 'ggggggdigital' -> 'digital').",
    "5) Do NOT add or remove information.",
    "6) Preserve ALL line breaks, headings, lists, and structural markers.",
    "",
    "Examples of required behavior (before → after):\n" + EXAMPLES_CORRECT,
    "",
    "Output policy: return ONLY the rewritten text, with the same line breaks."
  ].join("\n");
}

// ------------------------ Public API ------------------------
export async function aiCorrectText(text) {
  return await callOpenAI({
    system: buildSystemCorrectionPrompt(),
    user: String(text),
    model: "gpt-4o-mini",
    temperature: 0
  });
}

export async function aiRephraseText(text) {
  return await callOpenAI({
    system: buildSystemRephrasePrompt(),
    user: String(text),
    model: "gpt-4o-mini",
    temperature: 0.2
  });
}
