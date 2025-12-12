// lib/aiSpellCheck.js
// VERSION 5.3 - Zero false-positive merges
// ✅ Regex merge ONLY if combined is in COMMON_WORDS
// ✅ Safe-merge ONLY for multi-space corruption (very strong signal)
// ✅ Single-letter and trailing-letter merges ONLY if combined is in COMMON_WORDS
// ✅ Better casing preservation

const COMMON_WORDS = new Set([
  // ... (TON DICTIONNAIRE INCHANGÉ)
]);

const VALID_PHRASES = new Set([
  // ... (TON SET INCHANGÉ)
]);

// ================= HELPERS =================

function isRealWord(word) {
  if (!word || word.length < 2) return false;
  const normalized = word.toLowerCase().trim();
  return COMMON_WORDS.has(normalized);
}

function isPlausibleWord(word) {
  if (!word) return false;
  const w = word.trim();
  if (!/^[A-Za-z]+$/.test(w)) return false;
  if (w.length < 3 || w.length > 30) return false;
  if (/^(.)\1{4,}$/.test(w.toLowerCase())) return false;
  return true;
}

function isValidPhrase(phrase) {
  const normalized = phrase.toLowerCase().trim();
  if (VALID_PHRASES.has(normalized)) return true;

  const parts = normalized.split(/\s+/);
  if (parts.length === 2) {
    const [first, second] = parts;

    const shortWords = [
      'a', 'an', 'the', 'to', 'of', 'in', 'on', 'at', 'by', 'for', 'with',
      'is', 'are', 'was', 'were', 'be', 'can', 'could', 'would', 'should', 'will',
      'may', 'might', 'must', 'have', 'has', 'had', 'do', 'does', 'did',
      'not', 'how', 'one', 'due'
    ];

    // Si "word + connector" => phrase normale à ne pas merger
    if (shortWords.includes(second) && isRealWord(first)) return true;

    // Si les 2 mots sont valides et que la fusion ne donne pas un vrai mot => phrase normale
    if (isRealWord(first) && isRealWord(second)) {
      const combined = first + second;
      if (!isRealWord(combined)) return true;
    }
  }

  return false;
}

/**
 * Applique une casse "safe":
 * - si l'erreur commence par une majuscule -> Capitalize
 * - sinon -> lower-case
 */
function applyCasingLike(original, combined) {
  if (!combined) return combined;
  const o = (original || '').trim();
  if (!o) return combined;

  const startsUpper = /^[A-Z]/.test(o);
  if (startsUpper) {
    return combined.charAt(0).toUpperCase() + combined.slice(1).toLowerCase();
  }
  return combined.toLowerCase();
}

/**
 * SAFE-MERGE UNIQUEMENT pour multi-spaces
 * Ex: "p  otential" -> "potential", "dis   connection" -> "disconnection"
 * ⚠️ Interdit pour fragments courts (sinon: "a message" -> "amessage")
 */
function canSafeMerge(error, correction) {
  const err = (error || '').trim();
  const corr = (correction || '').trim();

  const errorHasSpace = /\s/.test(err);
  const correctionHasNoSpace = !/\s/.test(corr);
  if (!errorHasSpace || !correctionHasNoSpace) return false;

  if (isValidPhrase(err.toLowerCase())) return false;

  // ✅ seul signal autorisé = multi-spaces
  if (!/\s{2,}/.test(err)) return false;

  return isPlausibleWord(corr);
}

// ================= REGEX PRE-DETECTION =================

function preDetectFragmentedWords(text) {
  const issues = [];

  // Pattern 1: "word word" -> merge ONLY if combined is a REAL WORD in dict.
  const fragmentPattern = /\b([a-zA-Z]{2,})[ ]+([a-zA-Z]{2,12})\b/g;
  let match;

  while ((match = fragmentPattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const part1 = match[1];
    const part2 = match[2];
    const combined = part1 + part2;

    // ✅ no safe merge here
    if (isRealWord(combined) && !isValidPhrase(fullMatch.toLowerCase())) {
      issues.push({
        error: fullMatch,
        correction: applyCasingLike(fullMatch, combined),
        type: 'fragmented_word',
        severity: 'high',
        message: `Word incorrectly split by space: "${fullMatch}" → "${combined}"`
      });
    }
  }

  // Pattern 2: single-letter + word (ex: "c an", "th e")
  // ✅ merge ONLY if combined is in dict (NO safe-merge)
  const singleLetterPattern = /\b([a-zA-Z])[ ]+([a-zA-Z]{1,15})\b/g;

  while ((match = singleLetterPattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const letter = match[1];
    const rest = match[2];
    const combined = letter + rest;

    if (isRealWord(combined) && !isValidPhrase(fullMatch.toLowerCase())) {
      issues.push({
        error: fullMatch,
        correction: applyCasingLike(fullMatch, combined),
        type: 'fragmented_word',
        severity: 'high',
        message: `Single letter fragment: "${fullMatch}" → "${combined}"`
      });
    }
  }

  // Pattern 3: word + trailing single-letter (ex: "essenc e")
  // ✅ merge ONLY if combined is in dict (NO safe-merge)
  const trailingLetterPattern = /\b([a-zA-Z]{2,})[ ]+([a-zA-Z])\b/g;

  while ((match = trailingLetterPattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const word = match[1];
    const letter = match[2];
    const combined = word + letter;

    if (isRealWord(combined) && !isValidPhrase(fullMatch.toLowerCase())) {
      issues.push({
        error: fullMatch,
        correction: applyCasingLike(fullMatch, combined),
        type: 'fragmented_word',
        severity: 'high',
        message: `Trailing letter fragment: "${fullMatch}" → "${combined}"`
      });
    }
  }

  // Pattern 4: multi-spaces inside a word (strong signal)
  const multiSpacePattern = /\b([a-zA-Z]+)[ ]{2,}([a-zA-Z]+)\b/g;

  while ((match = multiSpacePattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const part1 = match[1];
    const part2 = match[2];
    const combined = part1 + part2;

    const okByDict = isRealWord(combined);
    const okBySafe = canSafeMerge(fullMatch, combined);

    if ((okByDict || okBySafe) && !isValidPhrase(fullMatch.toLowerCase())) {
      issues.push({
        error: fullMatch,
        correction: applyCasingLike(fullMatch, combined),
        type: 'multiple_spaces',
        severity: 'high',
        message: `Multiple spaces in word: "${fullMatch}" → "${combined}"`
      });
    }
  }

  return issues;
}

// ================= AI FILTERING =================

function filterInvalidCorrections(corrections) {
  return (corrections || []).filter(c => {
    const error = (c.error || '').trim();
    const correction = (c.correction || '').trim();
    const type = (c.type || '').trim();

    if (!error || !correction) return false;

    if (error.toLowerCase() === correction.toLowerCase()) {
      console.log(`❌ REJECTED (no change): "${error}" → "${correction}"`);
      return false;
    }

    // Never merge valid phrases
    if (isValidPhrase(error.toLowerCase())) {
      console.log(`❌ REJECTED (valid phrase): "${error}"`);
      return false;
    }

    const errorHasSpace = /\s/.test(error);
    const correctionHasNoSpace = !/\s/.test(correction);

    // Merge attempts
    if (errorHasSpace && correctionHasNoSpace) {
      // ✅ accept only if dict OR multi-space safe merge
      if (isRealWord(correction)) {
        console.log(`✅ ACCEPTED (dict merge): "${error}" → "${correction}"`);
        return true;
      }

      if (type === 'multiple_spaces' && canSafeMerge(error, correction)) {
        console.log(`✅ ACCEPTED (safe multi-space merge): "${error}" → "${correction}"`);
        return true;
      }

      console.log(`❌ REJECTED (merge blocked): "${error}" → "${correction}"`);
      return false;
    }

    // Other types: accept with basic sanity checks
    if (type === 'capitalization') {
      if (!/[A-Za-z]/.test(correction)) return false;
      console.log(`✅ ACCEPTED: "${error}" → "${correction}"`);
      return true;
    }

    if (type === 'punctuation' || type === 'spelling' || type === 'grammar') {
      if (correction.length < 1) return false;
      console.log(`✅ ACCEPTED: "${error}" → "${correction}"`);
      return true;
    }

    console.log(`✅ ACCEPTED (fallback): "${error}" → "${correction}"`);
    return true;
  });
}

// ================= AI PROMPT =================

const PROFESSIONAL_PROOFREADER_PROMPT = `
You are an expert English academic proofreader.

Your job is to detect and propose corrections for REAL errors only, using full sentence context.
This is NOT a simple spellchecker: you must reconstruct intended words when the text is clearly corrupted.

CRITICAL PRIORITIES (must catch these):
A) Fragmented words (spaces inside a single intended word)
- Examples: "soc ial"→"social", "enablin g"→"enabling", "gen uine"→"genuine", "dar ker"→"darker"
- Also multiple spaces inside a word: "p  otential"→"potential", "dis   connection"→"disconnection"
- Single-letter fragments: "th e"→"the", "o f"→"of"

B) Corrupted typos where the intended word is obvious in context
- Extra letters: "correctionnns"→"corrections" or "correction"
- Keyboard/near typos: "searcrh"→"search"
- Repeated garbage prefix/suffix: "gggdigital"→"digital"
- Punctuation inserted inside words: "corpo,rations"→"corporations"

C) Capitalization that is clearly incorrect
- Sentence start must be capitalized.
- Proper nouns must be capitalized (names, brands, countries).
- Titles should be in Title Case (e.g., "The Impact of Social Media on Modern Communication").

STRICT RULES:
1) Do NOT invent content, do NOT paraphrase, do NOT rewrite sentences.
   Only fix spelling/spacing/capitalization/punctuation errors.

2) Do NOT merge valid separate words.
   If it is clearly two intended words, keep them separate.
   Example: "become one", "message can", "strike a", "largely due" must remain two words.

3) Return corrections as small exact spans (shortest possible "error" string) so replacement is safe.
   The "error" must match the exact substring in the input.

OUTPUT FORMAT:
Return ONLY a valid JSON array of objects with exactly this schema:
[
  {
    "error": "exact substring from input",
    "correction": "corrected substring",
    "type": "fragmented_word|multiple_spaces|spelling|capitalization|punctuation|grammar",
    "severity": "high|medium|low",
    "message": "short reason"
  }
]

Return ONLY the JSON array, no other text.
`;

// ================= MAIN =================

export async function checkSpellingWithAI(text) {
  console.log('🧪 aiSpellCheck loaded VERSION 5.3');
  console.log('📝 AI spell-check VERSION 5.3 starting...');

  if (!text || text.length < 10) return [];

  const maxChars = 15000;
  const truncatedText = text.length > maxChars ? text.substring(0, maxChars) : text;
  console.log(`[SPELLCHECK] Text length: ${truncatedText.length} characters`);

  const regexIssues = preDetectFragmentedWords(truncatedText);
  console.log(`[REGEX] Found ${regexIssues.length} fragmented word patterns`);

  let aiIssues = [];

  try {
    const OPENAI_API_KEY = process.env.OPENAI_API_KEY;

    if (!OPENAI_API_KEY) {
      console.log('[SPELLCHECK] No OpenAI API key, using regex-only detection');
      return regexIssues;
    }

    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [
          { role: 'system', content: PROFESSIONAL_PROOFREADER_PROMPT },
          { role: 'user', content: `Analyze this text for errors:\n\n${truncatedText}` }
        ],
        temperature: 0.1,
        max_tokens: 2000
      })
    });

    if (!response.ok) {
      console.error(`[SPELLCHECK] OpenAI API error: ${response.status}`);
      return regexIssues;
    }

    const data = await response.json();
    const content = data.choices?.[0]?.message?.content || '[]';

    try {
      const jsonMatch = content.match(/\[[\s\S]*\]/);
      if (jsonMatch) {
        const rawIssues = JSON.parse(jsonMatch[0]);
        console.log(`[AI RAW] Detected ${rawIssues.length} errors`);
        aiIssues = filterInvalidCorrections(rawIssues);
      }
    } catch (parseError) {
      console.error('[SPELLCHECK] Failed to parse AI response:', parseError.message);
    }
  } catch (error) {
    console.error('[SPELLCHECK] AI call failed:', error.message);
  }

  const allIssues = [...regexIssues];
  const seenErrors = new Set(regexIssues.map(i => (i.error || '').toLowerCase()));

  for (const issue of aiIssues) {
    const errorKey = (issue.error || '').toLowerCase();
    if (!seenErrors.has(errorKey)) {
      allIssues.push({
        error: issue.error,
        correction: issue.correction,
        type: issue.type || 'spelling',
        severity: issue.severity || 'medium',
        message: issue.message || 'AI-detected error'
      });
      seenErrors.add(errorKey);
    }
  }

  console.log(`✅ AI spell-check complete: ${allIssues.length} valid errors found`);
  allIssues.slice(0, 25).forEach((issue, i) => {
    console.log(`  ${i + 1}. [${issue.type}] "${issue.error}" → "${issue.correction}" (${issue.message})`);
  });

  return allIssues;
}

export default { checkSpellingWithAI };
