// lib/ai.js
import axios from 'axios';

const PROVIDER = process.env.AI_PROVIDER || 'openai';
const DEFAULT_MODEL = process.env.AI_MODEL || 'gpt-4o-mini';
const MAX_RETRIES = Number(process.env.AI_MAX_RETRIES || 4);

/** Utilitaire sleep (ms) */
function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function ensureOpenAI() {
  if (PROVIDER !== 'openai') {
    throw new Error(`AI_PROVIDER non supporté: ${PROVIDER}`);
  }
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) throw new Error('OPENAI_API_KEY manquant');
  return { apiKey, model: DEFAULT_MODEL };
}

/** Appel OpenAI avec retry exponentiel sur 429/5xx. Retourne string | null. */
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

      const content = data?.choices?.[0]?.message?.content ?? '';
      return String(content).trim();
    } catch (e) {
      lastError = e;
      const status = e?.response?.status;
      const retriable = status === 429 || (status >= 500 && status <= 599);

      // backoff exponentiel + jitter
      if (retriable && attempt < MAX_RETRIES - 1) {
        const base = 1000 * Math.pow(2, attempt); // 1s, 2s, 4s, 8s...
        const jitter = Math.floor(Math.random() * 400); // + [0..400]ms
        await sleep(base + jitter);
        continue;
      }
      break;
    }
  }
  // En cas d’échec total, on renvoie null (le serveur fera un fallback sans 500)
  return null;
}

/** Correction IA sans reformulation (typos/espaces/grammaire légère) */
export async function aiProofread(text) {
  if (!text) return null;

  const system = 'You are a precise proofreader. Fix typos, split/merge words correctly, remove duplicated letters/spaces/punctuation, and do light grammar fixes. Do NOT paraphrase or change style.';
  const user = [
    'Correct the following text. Keep wording identical unless fixing mistakes (split/merge words, fix duplicated letters/spaces/punctuation, light grammar).',
    'Return only the corrected text.',
    '',
    'TEXT:',
    text
  ].join('\n');

  const out = await openAIChatWithRetry(
    [
      { role: 'system', content: system },
      { role: 'user', content: user }
    ],
    { temperature: 0.2 }
  );

  return out || null;
}

/** Reformulation IA (plus marquée, même langue, même sens) */
export async function aiRephrase(text) {
  if (!text) return null;

  const system = [
    'You are a professional editor.',
    'Rewrite the user text to improve clarity and flow, CHANGE THE WORDING noticeably while preserving the original meaning.',
    'Keep the SAME LANGUAGE as the input (French stays French, English stays English).',
    'Keep roughly similar length; do not add new facts.',
    'Preserve lists and headings if present; simplify punctuation/spaces.',
    'Return ONLY the rewritten text, no explanations or markdown.'
  ].join(' ');

  const user = [
    'Rewrite the following text.',
    'Aim for at least a 30–50% lexical change while keeping the exact meaning.',
    'No quotes, no commentary—just the rewritten text:',
    '',
    text
  ].join('\n');

  const out = await openAIChatWithRetry(
    [
      { role: 'system', content: system },
      { role: 'user', content: user }
    ],
    { temperature: 0.5 }
  );

  return out || null;
}

