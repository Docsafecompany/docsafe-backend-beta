// lib/ai.js
// IA = seule source de vérité.
// V1 : correction stricte (orthographe/ponctuation/espaces/scissions).
// V2 : REFORMULATION CLAIRE ET DISTINCTE + corrections, même langue, même structure.

const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";

// ---- Core call --------------------------------------------------------------
async function callOpenAI({
  system,
  user,
  model = "gpt-4o-mini",
  temperature = 0,
  frequency_penalty = 0.3,   // favorise des tournures différentes
  presence_penalty = 0.0,
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

// ---- Prompts ---------------------------------------------------------------

const EXAMPLES_CORRECT = [
  "soc ial → social",
  "enablin g → enabling",
  "th e → the",
  "gdigital / ggggggdigital → digital",
  "corpo, rations → corporations",
  "dis connection → disconnection"
].join("\n");

function buildSystemCorrectionPrompt() {
  return [
    "You are a STRICT multilingual copy editor.",
    "Return ONLY the corrected text, in the SAME language as the input.",
    "Do all of the following:",
    "• Fix spelling/diacritics/grammar.",
    "• Normalize punctuation and spacing.",
    "• Merge internally split words into one correct word.",
    "• Replace non-words with the correct word (e.g., 'gdigital'→'digital').",
    "• Preserve ALL line breaks, headings and list structure.",
    "• Do NOT paraphrase or change meaning.",
    "",
    "Examples:\n" + EXAMPLES_CORRECT
  ].join("\n");
}

function buildSystemRephrasePrompt() {
  return [
    "You are a multilingual rewriter.",
    "Return ONLY the rewritten text, in the SAME language as the input.",
    "Goals (apply ALL):",
    "1) Rewrite EACH sentence so the wording is CLEARER and NOTICEABLY DIFFERENT while preserving meaning and tone.",
    "2) Prefer active voice, tighten verbosity, improve flow; you may split long sentences or merge very short ones, but keep the same information.",
    "3) ALSO correct spelling/diacritics/grammar, normalize punctuation/spacing, and merge internally split words; fix non-words (e.g., 'gdigital'→'digital').",
    "4) Keep the document structure: keep ALL line breaks, headings and list markers in the SAME places.",
    "5) Do NOT add new facts and do NOT omit information.",
    "",
    "Rewrite policy:",
    "• Avoid copying original phrases verbatim; aim for clearly new phrasing per sentence.",
    "• Keep paragraph boundaries and list bullets exactly where they are.",
    "",
    "Examples of BEFORE → AFTER (style only):",
    "• “Overall, this is a sentence that could be rephrased.” → “Overall, this sentence can be expressed more directly.”",
    "• “Social media is very important today.” → “Today, social media plays a central role.”"
  ].join("\n");
}

// ---- Public API ------------------------------------------------------------

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
    temperature: 0.55,   // un peu plus de liberté pour reformuler
    frequency_penalty: 0.5
  });
}

