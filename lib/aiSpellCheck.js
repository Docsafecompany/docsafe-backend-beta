// lib/aiSpellCheck.js
// VERSION 3.0 - Correction intelligente avec validation contextuelle

/**
 * Dictionnaire de mots communs pour validation
 */
const COMMON_WORDS = new Set([
  // English common words
  'the', 'and', 'for', 'are', 'but', 'not', 'you', 'all', 'can', 'had', 'her', 'was', 'one', 'our', 'out',
  'has', 'his', 'how', 'its', 'may', 'new', 'now', 'old', 'see', 'way', 'who', 'boy', 'did', 'get', 'let',
  'put', 'say', 'she', 'too', 'use', 'dad', 'mom', 'day', 'man', 'men', 'run', 'set', 'top', 'try', 'ask',
  'big', 'end', 'eye', 'far', 'few', 'got', 'him', 'job', 'lot', 'own', 'pay', 'per', 'sit', 'yes', 'yet',
  'able', 'also', 'back', 'been', 'both', 'call', 'come', 'could', 'each', 'even', 'find', 'from', 'give',
  'good', 'have', 'here', 'high', 'just', 'know', 'last', 'left', 'life', 'like', 'line', 'long', 'look',
  'made', 'make', 'many', 'more', 'most', 'much', 'must', 'name', 'need', 'next', 'only', 'over', 'part',
  'people', 'place', 'point', 'right', 'same', 'show', 'side', 'some', 'such', 'take', 'tell', 'than',
  'that', 'them', 'then', 'there', 'these', 'they', 'this', 'time', 'turn', 'upon', 'used', 'very', 'want',
  'well', 'went', 'what', 'when', 'where', 'which', 'while', 'with', 'word', 'work', 'world', 'would', 'year',
  'about', 'after', 'again', 'being', 'between', 'came', 'change', 'could', 'different', 'does', 'down',
  'during', 'every', 'first', 'follow', 'found', 'great', 'group', 'hand', 'having', 'head', 'help', 'home',
  'house', 'important', 'into', 'keep', 'kind', 'large', 'learn', 'little', 'live', 'mean', 'might', 'move',
  'never', 'number', 'often', 'order', 'other', 'page', 'play', 'possible', 'present', 'problem', 'public',
  'question', 'read', 'really', 'report', 'said', 'school', 'second', 'should', 'since', 'small', 'something',
  'sound', 'spell', 'start', 'state', 'still', 'story', 'study', 'system', 'their', 'think', 'thought',
  'three', 'through', 'together', 'under', 'understand', 'until', 'water', 'without', 'write', 'young',
  // Document/business words
  'enabling', 'individuals', 'landscape', 'communication', 'platform', 'potential', 'connection', 'digital',
  'social', 'media', 'balance', 'message', 'matter', 'sense', 'tool', 'barrier', 'double', 'realm', 'sword',
  'contribution', 'ability', 'significant', 'distances', 'platforms', 'presence', 'letters', 'abroad',
  'travel', 'seconds', 'making', 'global', 'community', 'provides', 'expression', 'contributing',
  'individuality', 'essential', 'strike', 'embracing', 'benefits', 'remains', 'meaningful', 'rather',
  'conclusion', 'stands', 'edged', 'embodies', 'unity', 'creativity', 'regulate', 'guide', 'essence',
  'increasingly', 'century', 'become', 'influential', 'forces', 'shaping', 'connect', 'share', 'information',
  'facebook', 'instagram', 'tiktok', 'twitter', 'dominate', 'daily', 'interactions', 'exchange', 'ideas',
  'across', 'geographical', 'boundaries', 'instantly', 'traditional', 'methods', 'phone', 'calls', 'relied',
  'heavily', 'physical', 'lengthy', 'processes', 'mailing', 'today', 'staying', 'touch', 'loved', 'ones',
  'fostered', 'breaking', 'barriers', 'self', 'cultures', 'identity', 'decline', 'face', 'personal', 'quality',
  'depth', 'nuance', 'struggle', 'replicate', 'screen', 'based', 'often', 'feels', 'impersonal', 'leading',
  'superficial', 'relationships', 'loneliness', 'users', 'surrounded', 'virtual', 'friends', 'genuine',
  'human', 'spread', 'misinformation', 'fake', 'news', 'viral', 'speed', 'corrections', 'issued', 'undermining',
  'trust', 'institutions', 'creating', 'divisions', 'within', 'communities', 'political', 'polarization',
  'recent', 'years', 'example', 'exacerbated', 'circulation', 'unverified', 'claims', 'additionally', 'raised',
  'concerns', 'privacy', 'commodification', 'data', 'every', 'interaction', 'likes', 'comments', 'search',
  'history', 'creates', 'footprints', 'corporations', 'targeted', 'advertising', 'personalization', 'enhance',
  'experience', 'raises', 'ethical', 'questions', 'consent', 'surveillance', 'fully', 'realize', 'extent',
  'habits', 'monitored', 'analyzed', 'growing', 'debates', 'protection', 'rights', 'ultimately', 'impact',
  'modern', 'profound', 'multifaceted', 'revolutionized', 'people', 'breaking', 'providing', 'activism',
  'contributed', 'erosion', 'disconnection', 'falsehoods', 'exploitation', 'responsibility', 'lies',
  'governments', 'critically', 'reflect', 'preserve', 'confidential', 'slide', 'professional', 'content',
  // French common words
  'le', 'la', 'les', 'de', 'du', 'des', 'un', 'une', 'et', 'est', 'en', 'que', 'qui', 'dans', 'ce', 'il',
  'pas', 'plus', 'pour', 'sur', 'par', 'au', 'aux', 'avec', 'son', 'sont', 'nous', 'vous', 'leur', 'cette',
  'bien', 'aussi', 'comme', 'tout', 'elle', 'entre', 'faire', 'fait', 'peut', 'donc', 'sans', 'mais', 'ou',
  'avant', 'après', 'avoir', 'être', 'très', 'encore', 'autre', 'notre', 'votre', 'même', 'quand', 'quel',
  'confidentiel', 'diapositive', 'contenu', 'professionnel', 'numérique', 'document', 'fichier', 'rapport'
]);

/**
 * Phrases valides à NE JAMAIS toucher (ce sont des séparations correctes)
 */
const VALID_PHRASES = new Set([
  // Article + noun/adj patterns
  'a message', 'a matter', 'a sense', 'a tool', 'a balance', 'a platform', 'a double', 'a barrier',
  'a new', 'a big', 'a few', 'a lot', 'a bit', 'a man', 'a day', 'a way', 'a set', 'a key',
  // Preposition patterns  
  'in the', 'of the', 'to the', 'by the', 'on the', 'at the', 'for the', 'from the', 'with the',
  'in a', 'of a', 'to a', 'by a', 'on a', 'at a', 'for a', 'from a', 'with a',
  'in an', 'of an', 'to an', 'by an', 'on an', 'at an', 'for an', 'from an',
  // Verb patterns
  'is the', 'is a', 'is an', 'is its', 'is it', 'is not', 'is also', 'is both',
  'it is', 'it has', 'it was', 'it can', 'it may', 'it will', 'it would', 'it could',
  'to it', 'to be', 'to do', 'to go', 'to see', 'to get', 'to make', 'to take',
  'has been', 'has the', 'has a', 'has its', 'has also', 'has both',
  'on how', 'of dis', // special cases
  // Pronoun patterns
  'i am', 'i was', 'i will', 'i have', 'i can', 'i could', 'i would', 'i should',
  // French patterns
  'de la', 'de le', 'à la', 'à le', 'en un', 'en une', 'il y', 'ce qui', 'ce que',
  'c\'est', 'qu\'il', 'qu\'elle', 'n\'est', 'd\'un', 'd\'une'
]);

/**
 * Vérifie si un mot existe dans le dictionnaire
 */
function isRealWord(word) {
  if (!word || word.length < 2) return false;
  const normalized = word.toLowerCase().trim();
  
  // Mots de 1-3 lettres sont généralement valides s'ils sont dans le dictionnaire
  if (normalized.length <= 3) {
    return COMMON_WORDS.has(normalized);
  }
  
  // Mots plus longs : vérifier le dictionnaire ou accepter si structure valide
  if (COMMON_WORDS.has(normalized)) return true;
  
  // Vérifier si le mot a une structure valide (pas de patterns bizarres)
  // Rejeter les mots avec des combinaisons impossibles
  const invalidPatterns = [
    /^[bcdfghjklmnpqrstvwxz]{4,}/i,  // 4+ consonnes au début
    /[bcdfghjklmnpqrstvwxz]{5,}/i,   // 5+ consonnes consécutives
    /(.)\1{3,}/i,                     // 4+ caractères identiques
    /^[aeiou]{4,}/i                   // 4+ voyelles au début
  ];
  
  for (const pattern of invalidPatterns) {
    if (pattern.test(normalized)) return false;
  }
  
  return true; // Accepter les mots avec structure valide non dans le dictionnaire
}

/**
 * Vérifie si une phrase est valide (ne doit pas être "corrigée")
 */
function isValidPhrase(phrase) {
  if (!phrase) return false;
  const normalized = phrase.toLowerCase().trim();
  return VALID_PHRASES.has(normalized);
}

/**
 * Obtient le contexte autour d'une position
 */
function getContext(text, index, length) {
  const start = Math.max(0, index - 25);
  const end = Math.min(text.length, index + length + 25);
  return text.slice(start, end).replace(/\s+/g, ' ').trim();
}

/**
 * Pré-détection ciblée des vrais mots fragmentés
 */
function preDetectFragmentedWords(text) {
  const issues = [];
  
  // Pattern: Mot + espace + 1-2 lettres qui complètent le mot
  // Ex: "enablin g" -> "enabling", "soc ial" -> "social"
  const fragmentPattern = /\b([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]{3,})\s+([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]{1,3})\b/g;
  
  let match;
  while ((match = fragmentPattern.exec(text)) !== null) {
    const [fullMatch, part1, part2] = match;
    const combined = part1 + part2;
    
    // Vérifier si la combinaison forme un vrai mot ET que les parties séparées ne forment PAS une phrase valide
    if (isRealWord(combined) && !isValidPhrase(fullMatch.toLowerCase())) {
      // Vérification supplémentaire : part2 seul n'est généralement pas un mot valide isolé
      // ou la combinaison est clairement un mot connu
      if (!isRealWord(part2) || COMMON_WORDS.has(combined.toLowerCase())) {
        issues.push({
          error: fullMatch,
          correction: combined,
          context: getContext(text, match.index, fullMatch.length),
          type: 'fragmented_word',
          severity: 'high',
          source: 'regex'
        });
      }
    }
  }
  
  // Pattern: 1-2 lettres + espace + mot (ex: "th e" -> "the")
  // MAIS exclure les patterns valides comme "a new", "I am"
  const prefixPattern = /\b([a-zA-Z]{1,2})\s+([a-zA-Z]{2,})\b/g;
  
  while ((match = prefixPattern.exec(text)) !== null) {
    const [fullMatch, prefix, word] = match;
    const combined = prefix + word;
    
    // Exclure les phrases valides
    if (isValidPhrase(fullMatch.toLowerCase())) continue;
    
    // Exclure "I am", "A new", etc.
    if (prefix.toLowerCase() === 'i' || prefix.toLowerCase() === 'a') {
      if (isRealWord(word)) continue;
    }
    
    // Vérifier si la combinaison forme un vrai mot
    if (isRealWord(combined) && !isRealWord(prefix)) {
      issues.push({
        error: fullMatch,
        correction: combined,
        context: getContext(text, match.index, fullMatch.length),
        type: 'fragmented_word', 
        severity: 'high',
        source: 'regex'
      });
    }
  }
  
  // Pattern: Espaces multiples dans un mot
  const multiSpacePattern = /\b([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]+)\s{2,}([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]+)\b/g;
  
  while ((match = multiSpacePattern.exec(text)) !== null) {
    const combined = match[1] + match[2];
    if (isRealWord(combined)) {
      issues.push({
        error: match[0],
        correction: combined,
        context: getContext(text, match.index, match[0].length),
        type: 'multiple_spaces',
        severity: 'high',
        source: 'regex'
      });
    }
  }
  
  console.log(`[REGEX] Found ${issues.length} fragmented word patterns`);
  return issues;
}

/**
 * Filtre les corrections invalides
 */
function filterInvalidCorrections(errors) {
  return errors.filter(err => {
    // 1. Rejeter si la correction ne produit pas un mot réel
    if (!isRealWord(err.correction)) {
      console.log(`❌ REJECTED (not a real word): "${err.error}" → "${err.correction}"`);
      return false;
    }
    
    // 2. Rejeter si c'est une phrase valide qu'on essaie de fusionner
    if (isValidPhrase(err.error)) {
      console.log(`❌ REJECTED (valid phrase): "${err.error}"`);
      return false;
    }
    
    // 3. Rejeter si erreur = correction (pas de changement)
    if (err.error.replace(/\s+/g, '').toLowerCase() === err.correction.replace(/\s+/g, '').toLowerCase()) {
      if (err.error === err.correction) {
        console.log(`❌ REJECTED (no change): "${err.error}"`);
        return false;
      }
    }
    
    return true;
  });
}

/**
 * Vérifie l'orthographe avec l'IA
 */
export async function checkSpellingWithAI(text) {
  if (!text || text.trim().length < 10) return [];
  
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.warn('⚠️ OPENAI_API_KEY not configured - spelling check disabled');
    return [];
  }

  try {
    console.log('📝 AI spell-check VERSION 3.0 starting...');
    console.log(`[SPELLCHECK] Text length: ${text.length} characters`);
    
    // ÉTAPE 1: Pré-détection regex ciblée
    const regexErrors = preDetectFragmentedWords(text);
    
    // ÉTAPE 2: Appel IA avec prompt intelligent
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [
          {
            role: 'system',
            content: `You are an expert multilingual proofreader. Your job is to detect REAL spelling and grammar errors.

CRITICAL RULES - YOU MUST FOLLOW THESE:

1. **VERIFY EVERY CORRECTION IS A REAL WORD**
   Before suggesting ANY correction, ask yourself: "Does this corrected word exist in the language?"
   - "gindividuals" does NOT exist → DO NOT suggest it
   - "amessage" does NOT exist → DO NOT suggest it  
   - "enabling" EXISTS → You can suggest it
   - "potential" EXISTS → You can suggest it

2. **NEVER MERGE LEGITIMATE SEPARATE WORDS**
   These are CORRECT as separate words - DO NOT flag them as errors:
   - "In the", "of the", "to the", "by the", "on the", "at the"
   - "a message", "a matter", "a sense", "a tool", "a balance", "a platform"
   - "is its", "is the", "It has", "it has", "to it", "in an"
   - "I am", "I was", "A new", "A big"

3. **UNDERSTAND FLOATING LETTERS BELONG TO ADJACENT WORDS**
   When you see text like: "enablin g individuals"
   - The floating "g" belongs to "enablin" → the correction is "enabling"
   - "individuals" is ALREADY a correct, complete word
   - DO NOT create fake words like "gindividuals"
   
   When you see: "soc ial media"
   - "ial" belongs to "soc" → correction is "social"
   - "media" is already correct

4. **ANALYZE CONTEXT BEFORE CORRECTING**
   Read the full sentence to understand:
   - Which word is incomplete/fragmented
   - Which direction the floating letters belong (usually to the LEFT word)
   - Whether the correction makes grammatical sense

5. **DETECT THESE SPECIFIC ERROR TYPES**
   - Fragmented words: "p otential" → "potential", "Addit onally" → "Additionally"
   - Multiple spaces: "dis  connection" → "disconnection"
   - Merged words: "thankyou" → "thank you"
   - Typos with punctuation: "corpo,rations" → "corporations"
   - Repeated characters: "gggdigital" → "digital"
   - Regular spelling errors

6. **SKIP IF NOT 100% CERTAIN**
   If you're not absolutely sure the correction is valid and the word exists, DO NOT include it.
   It's better to miss an error than to suggest a wrong correction.

OUTPUT: Return ONLY errors where you are 100% confident the correction is a real word.`
          },
          {
            role: 'user',
            content: `Analyze this text and find spelling errors. Remember: ONLY suggest corrections that result in REAL words that exist. Never merge separate words incorrectly.\n\n${text.slice(0, 12000)}`
          }
        ],
        tools: [
          {
            type: 'function',
            function: {
              name: 'report_spelling_errors',
              description: 'Report detected spelling errors with validated corrections',
              parameters: {
                type: 'object',
                properties: {
                  errors: {
                    type: 'array',
                    items: {
                      type: 'object',
                      properties: {
                        error: { type: 'string', description: 'The exact incorrect text as it appears' },
                        correction: { type: 'string', description: 'The corrected text (MUST be a real word)' },
                        context: { type: 'string', description: 'Surrounding text for context' },
                        severity: { type: 'string', enum: ['low', 'medium', 'high'] },
                        type: { type: 'string', enum: ['fragmented_word', 'multiple_spaces', 'merged_word', 'spelling', 'grammar', 'punctuation'] }
                      },
                      required: ['error', 'correction', 'severity', 'type']
                    }
                  }
                },
                required: ['errors']
              }
            }
          }
        ],
        tool_choice: { type: 'function', function: { name: 'report_spelling_errors' } },
        temperature: 0.05,
        max_tokens: 3000
      })
    });

    if (!response.ok) {
      console.error('❌ OpenAI API error:', response.status);
      return formatErrors(filterInvalidCorrections(regexErrors));
    }

    const data = await response.json();
    const toolCall = data.choices?.[0]?.message?.tool_calls?.[0];
    
    let aiErrors = [];
    if (toolCall?.function?.arguments) {
      try {
        const parsed = JSON.parse(toolCall.function.arguments);
        aiErrors = parsed.errors || [];
        console.log(`[AI RAW] Detected ${aiErrors.length} errors`);
      } catch (e) {
        console.error('❌ Failed to parse AI response:', e.message);
      }
    }
    
    // ÉTAPE 3: Fusionner et dédupliquer
    const allErrors = mergeErrors(regexErrors, aiErrors);
    
    // ÉTAPE 4: FILTRER les corrections invalides
    const validErrors = filterInvalidCorrections(allErrors);
    
    // ÉTAPE 5: Formater
    const spellingErrors = formatErrors(validErrors);

    console.log(`✅ AI spell-check complete: ${spellingErrors.length} valid errors found`);
    spellingErrors.forEach((err, i) => {
      console.log(`  ${i + 1}. [${err.type}] "${err.error}" → "${err.correction}"`);
    });
    
    return spellingErrors;

  } catch (error) {
    console.error('❌ AI spell-check error:', error.message);
    return [];
  }
}

/**
 * Fusionne les erreurs en évitant les doublons
 */
function mergeErrors(regexErrors, aiErrors) {
  const merged = [...regexErrors];
  const seen = new Set(regexErrors.map(e => e.error.toLowerCase().replace(/\s+/g, '')));
  
  for (const aiErr of aiErrors) {
    const normalized = aiErr.error.toLowerCase().replace(/\s+/g, '');
    if (!seen.has(normalized)) {
      merged.push(aiErr);
      seen.add(normalized);
    }
  }
  
  return merged;
}

/**
 * Formate les erreurs pour le frontend
 */
function formatErrors(errors) {
  return errors.map((err, index) => ({
    id: `spell_${index}_${Date.now()}`,
    error: err.error,
    correction: err.correction,
    context: err.context || '',
    location: `Detection ${index + 1}`,
    severity: err.severity || 'medium',
    type: err.type || 'spelling',
    rule: err.type?.toUpperCase() || 'AI_CORRECTION',
    message: getErrorMessage(err.type),
    category: getCategory(err.type),
    source: err.source || 'ai'
  }));
}

function getErrorMessage(type) {
  const messages = {
    'fragmented_word': 'Word incorrectly split by space',
    'multiple_spaces': 'Multiple spaces inside word',
    'merged_word': 'Words incorrectly merged together',
    'spelling': 'Spelling error',
    'grammar': 'Grammar error',
    'punctuation': 'Punctuation error'
  };
  return messages[type] || 'Spelling/grammar correction';
}

function getCategory(type) {
  return ['fragmented_word', 'multiple_spaces', 'merged_word'].includes(type) 
    ? 'Word Fragments' 
    : type === 'grammar' ? 'Grammar' : 'Spelling';
}

/**
 * Applique les corrections au texte
 */
export function applyCorrections(text, corrections) {
  if (!corrections?.length) return { correctedText: text, examples: [], changedCount: 0 };

  let correctedText = text;
  const examples = [];
  let changedCount = 0;

  // Trier par longueur décroissante
  const sorted = [...corrections].sort((a, b) => (b.error?.length || 0) - (a.error?.length || 0));

  for (const c of sorted) {
    if (!c.error || !c.correction || c.error === c.correction) continue;
    
    const regex = new RegExp(escapeRegExp(c.error), 'gi');
    const matches = correctedText.match(regex);
    
    if (matches?.length) {
      correctedText = correctedText.replace(regex, c.correction);
      changedCount += matches.length;
      if (examples.length < 12) {
        examples.push({ before: c.error, after: c.correction, type: c.type });
      }
    }
  }

  console.log(`[APPLY] ${changedCount} corrections applied`);
  return { correctedText, examples, changedCount };
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export default { checkSpellingWithAI, applyCorrections };

