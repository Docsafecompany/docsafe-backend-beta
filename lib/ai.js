// lib/ai.js
// V1 : correction stricte (même texte, fautes fixées).
// V2 : reformulation CLAIREMENT différente + corrections, structure conservée.

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";

// Appel générique
async function callOpenAI({
  system,
  user,
  model = "gpt-4o-mini",
  temperature = 0,
  frequency_penalty = 0.6,   // décourage répétition des mêmes tournures
  presence_penalty = 0.1,
  top_p = 1
}) {
  if (!OPENAI_API_KEY) return null;
  const resp = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Authorization": `Bearer ${OPENAI_API_KEY}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      model,
      temperature,
      frequency_penalty,
      presence_penalty,
      top_p,
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

// --------- V1: Correction stricte (aucune paraphrase) ---------
function buildSystemCorrectionPrompt() {
  return [
    "You are a STRICT multilingual copy editor.",
    "Return ONLY the corrected text, in the SAME language as input.",
    "Do all of the following:",
    "• Fix spelling/diacritics/grammar.",
    "• Normalize punctuation and spacing; no duplicate punctuation.",
    "• Merge internally split words (e.g., 'soc ial'→'social', 'enablin g'→'enabling').",
    "• Replace non-words with the correct word (e.g., 'gdigital'/'ggggggdigital'→'digital').",
    "• Keep separate words separate (e.g., 'social media' stays two words).",
    "• Preserve ALL line breaks, headings and list markers.",
    "• Do NOT paraphrase; do NOT change meaning.",
    "Output: ONLY the corrected text."
  ].join("\n");
}

export async function aiCorrectText(text) {
  return await callOpenAI({
    system: buildSystemCorrectionPrompt(),
    user: String(text),
    model: "gpt-4o-mini",
    temperature: 0
  });
}

// --------- V2: Reformulation marquée (phrase par phrase) ---------
function buildSystemRephrasePrompt() {
  return [
    "You are a multilingual REWRITER.",
    "Return ONLY the rewritten text, in the SAME language as input.",
    "Objectives (apply ALL):",
    "1) Reformulate EACH sentence so wording is CLEARLY DIFFERENT while preserving meaning and tone.",
    "   • Prefer active voice; improve flow and clarity.",
    "   • You may split long sentences or merge very short ones *without* dropping information.",
    "   • Avoid copying original phrases; avoid n-grams ≥ 4 words from the input.",
    "2) ALSO correct spelling/diacritics/grammar, normalize punctuation/spacing.",
    "   • Merge internally split words (e.g., 'soc ial'→'social').",
    "   • Replace non-words (e.g., 'ggggggdigital'→'digital').",
    "3) STRICTLY preserve the document structure:",
    "   • Keep the SAME line breaks (paragraph boundaries) and list markers in the SAME places.",
    "   • Do NOT add headings or remove existing ones.",
    "4) Do NOT add facts; do NOT omit information.",
    "Output: ONLY the rewritten text, with the same number of lines as input."
  ].join("\n");
}

export async function aiRephraseText(text) {
  return await callOpenAI({
    system: buildSystemRephrasePrompt(),
    user: String(text),
    model: "gpt-4o-mini",
    temperature: 0.7,       // plus de liberté pour reformuler
    frequency_penalty: 0.8, // force des tournures différentes
    presence_penalty: 0.2
  });
}


