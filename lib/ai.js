// lib/ai.js
import axios from 'axios';
const PROVIDER = process.env.AI_PROVIDER || 'openai';

export async function aiProofread(text) {
  if (!text) return null;
  if (PROVIDER !== 'openai') throw new Error(`AI_PROVIDER non supporté: ${PROVIDER}`);
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error('OPENAI_API_KEY manquant');

  const model = process.env.AI_MODEL || 'gpt-4o-mini';
  const temperature = Number(process.env.AI_TEMPERATURE || 0.2);

  const { data } = await axios.post('https://api.openai.com/v1/chat/completions', {
    model,
    temperature,
    messages: [
      {
        role: 'system',
        content:
          'You are a precise proofreader. Fix typos, split/merge words correctly, fix spacing and light grammar, but DO NOT paraphrase or change style. Preserve original sentences as much as possible.'
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

export async function aiRephrase(text) {
  if (!text) return null;
  if (PROVIDER !== 'openai') throw new Error(`AI_PROVIDER non supporté: ${PROVIDER}`);
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error('OPENAI_API_KEY manquant');

  const model = process.env.AI_MODEL || 'gpt-4o-mini';
  const temperature = Number(process.env.AI_TEMPERATURE || 0.3);

  const { data } = await axios.post('https://api.openai.com/v1/chat/completions', {
    model,
    temperature,
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
