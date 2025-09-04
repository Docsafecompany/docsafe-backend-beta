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

  const { data } = await axios.post('https://api.openai.com/v1/chat/completions', {
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
  }, { headers: { Authorization: `Bearer ${apiKey}` } });

  return data?.choices?.[0]?.message?.content?.trim() || null;
}

/** Reformulation IA (clarifie la rédaction en conservant le sens) */
export async function aiRephrase(text) {
  if (!text) return null;
  const { apiKey, model } = ensureOpenAI();

  const { data } = await axios.post('https://api.openai.com/v1/chat/completions', {
    model,
    temperature: 0.3,
    messages: [
      {
        role: 'system',
        content:
          'You are a rewriting assistant. Improve clarity and flow while preserving meaning. Keep formatting simple; do not invent facts.'
      },
      {
        role: 'user',
        content:
`Rewrite the following text in a clear, professional tone, preserving structure and terminology.

TEXT:
${text}`
      }
    ]
  }, { headers: { Authorization: `Bearer ${apiKey}` } });

  return data?.choices?.[0]?.message?.content?.trim() || null;
}
