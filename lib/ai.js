import axios from 'axios';


const PROVIDER = process.env.AI_PROVIDER || 'openai';


export async function aiRephrase(text) {
if (!text) return null;


if (PROVIDER === 'openai') {
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
content: 'You are a precise rewriting assistant. Improve clarity, grammar, and style while preserving meaning. Keep formatting simple, avoid hallucinations, and never invent facts.'
},
{
role: 'user',
content: `Rewrite the following text in a clear, professional tone, preserving structure and terminology.\n\nTEXT:\n${text}`
}
]
}, {
headers: { Authorization: `Bearer ${apiKey}` }
});


const out = data?.choices?.[0]?.message?.content?.trim();
return out || null;
}


// Hook générique pour d’autres providers (Anthropic, Azure, Mistral, etc.)
throw new Error(`AI_PROVIDER non supporté: ${PROVIDER}`);
}
