// lib/ai.js
import axios from 'axios';

const PROVIDER = process.env.AI_PROVIDER || 'openai';
const DEFAULT_MODEL = process.env.AI_MODEL || 'gpt-4o-mini';
const MAX_RETRIES = Number(process.env.AI_MAX_RETRIES || 4);

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function ensureOpenAI() {
  if (PROVIDER !== 'openai') throw new Error(`AI_PROVIDER non supporté: ${PROVIDER}`);
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error('OPENAI_API_KEY manquant');
  return { apiKey, model: DEFAULT_MODEL };
}

async function openAIChatWithRetry(messages, { temperature = 0.2 } = {}) {
  const { apiKey, model } = ensureOpenAI();
  let lastError = null;

  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    try {
      const { data } = await axios.post(
        'https://api.openai.com/v1/chat/completions',
        { model, temperature, messages },
        { headers: { Authorization: `Bearer ${apiKey}` }, timeout: 60000 }
      );
      return String(data?.choices?.[0]?.message?.content ?? '').trim();
    } catch (e) {
      lastError = e;
      const status = e?.response?.status;
      const retriable = status === 429 || (status >= 500 && status <= 599);
      if (retriable && attempt < MAX_RETRIES - 1) {
        const base = 1000 * Math.pow(2, attempt); // 1s,2s,4s...
        const jitter = Math.floor(Math.random() * 400);
        await sleep(base + jitter);
        continue;
      }
      break;
    }
  }
  return null; // fallback handled by server
}

export async function aiProofread(text) {
  if (!text) return null;
  const system = 'You are a precise proofreader. Fix typos, split/merge words, remove duplicated letters/spaces/punctuation, light grammar only. Do NOT paraphrase.';
  const user = ['Correct the following text. Return only the corrected text.', '', 'TEXT:', text].join('\n');
  return await openAIChatWithRetry(
    [{ role: 'system', content: system }, { role: 'user', content: user }],
    { temperature: 0.2 }
  );
}

export async function aiRephrase(text) {
  if (!text) return null;
  const system = 'You are a professional editor. Rewrite to improve clarity and flow with a noticeable wording change while preserving meaning and language. No explanations.';
  const user = ['Rewrite the text below. Target 30–50% lexical change, same meaning, same language.', '', text].join('\n');
  return await openAIChatWithRetry(
    [{ role: 'system', content: system }, { role: 'user', content: user }],
    { temperature: 0.5 }
  );
}


