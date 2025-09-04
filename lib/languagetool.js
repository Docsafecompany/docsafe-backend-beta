import axios from 'axios';


export async function ltCheck(text) {
if (!process.env.LT_ENDPOINT) return null;
const endpoint = process.env.LT_ENDPOINT;
const language = process.env.LT_LANGUAGE || 'en-US';
const { data } = await axios.post(endpoint, new URLSearchParams({
text,
language
}));
return data; // structure LT standard { matches: [...] }
}
