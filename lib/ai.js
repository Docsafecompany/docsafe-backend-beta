import axios from 'axios';

const PROVIDER = process.env.AI_PROVIDER || 'openai';

function ensureOpenAI() {
  if (PROVIDER !== 'openai') throw new Error(`AI_PROVIDER non supporté: ${PROVIDER}`);
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error('OPENAI_API_KEY manquant');
  const model = process.env.AI_MODEL || 'gpt-4o-mini';
  return { apiKey, model };
}

/** Correction IA sans reformulation (typos, espaces, grammaire légère) */
export async function aiProofread(text) {
  if (!text) return null;
  const { apiKey, model } = ensureOpenAI();

  const { data } = await axios.post(
    'https://api.openai.com/v1/chat/completions',
    {
      model,
      temperature: 0.2,
      messages: [
        {
          role: 'system',
          content:
            'You are a precise proofreader. Fix typos, split/merge words correctly, fix duplicated letters/spaces/punctuation, and light grammar. DO NOT paraphrase or change style.'
        },
        {
          role: 'user',
          content:
`Correct the following text. Keep wording identical unless fixing mistakes (split/merge words, fix duplicated letters/spaces/punctuation, light grammar). Return only the corrected text.

TEXT:
${text}`
        }
      ]
    },
    { headers: { Authorization: `Bearer ${apiKey}` } }
  );

  return data?.choices?.[0]?.message?.content?.trim() || null;
}

/** Reformulation IA (plus marquée, même langue, même sens) */
export async function aiRephrase(text) {
  if (!text) return null;
  const { apiKey, model } = ensureOpenAI();

  const { data } = await axios.post(
    'https://api.openai.com/v1/chat/completions',
    {
      model,
      temperature: 0.5,
      messages: [
        {
          role: 'system',
          content: [
            'You are a professional editor.',
            'Rewrite the user text to improve clarity and flow, CHANGE THE WORDING noticeably while preserving the original meaning.',
            'Keep the SAME LANGUAGE as the input (French stays French, English stays English).',
            'Keep roughly similar length; do not add new facts.',
            'Preserve lists and headings if present; simplify punctuation/spaces.',
            'Return ONLY the rewritten text, no explanations or markdown.'
          ].join(' ')
        },
        {
          role: 'user',
          content: [
            'Rewrite the following text.',
            'Aim for at least a 30–50% lexical change while keeping the exact meaning.',
            'No quotes, no commentary—just the rewritten text:\n\n',
            text
          ].join(' ')
        }
      ]
    },
    { headers: { Authorization: `Bearer ${apiKey}` } }
  );

  return data?.choices?.[0]?.message?.content?.trim() || null;
}
