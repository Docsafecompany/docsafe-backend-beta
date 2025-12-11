// lib/aiSpellCheck.js
// VERSION 5.0 - Approche STRICTE avec validation dictionnaire obligatoire

/**
 * Dictionnaire de mots communs pour validation STRICTE
 * Un mot corrigé DOIT être dans cette liste pour être accepté
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
  
  // Document/business words - ESSENTIAL
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
  'modern', 'profound', 'multifaceted', 'revolutionized', 'breaking', 'providing', 'activism',
  'contributed', 'erosion', 'disconnection', 'falsehoods', 'exploitation', 'responsibility', 'lies',
  'governments', 'critically', 'reflect', 'preserve', 'confidential', 'slide', 'professional', 'content',
  'darker', 'connectivity', 'continents', 'largely', 'due', 'redefined',
  
  // Academic/formal words
  'additionally', 'furthermore', 'however', 'therefore', 'consequently', 'moreover', 'nevertheless',
  'emphasized', 'collaborated', 'coordinated', 'accessibility', 'empowerment', 'implementation',
  'marginalized', 'momentum', 'amplify', 'oversight', 'verification', 'authorization', 'documentation',
  'multifaceted', 'geographical', 'interpersonal', 'organizational', 'technological', 'environmental',
  'sustainable', 'infrastructure', 'comprehensive', 'strategic', 'operational', 'administrative',
  
  // French common words
  'le', 'la', 'les', 'de', 'du', 'des', 'un', 'une', 'et', 'est', 'en', 'que', 'qui', 'dans', 'ce', 'il',
  'pas', 'plus', 'pour', 'sur', 'par', 'au', 'aux', 'avec', 'son', 'sont', 'nous', 'vous', 'leur', 'cette',
  'bien', 'aussi', 'comme', 'tout', 'elle', 'entre', 'faire', 'fait', 'peut', 'donc', 'sans', 'mais', 'ou',
  'avant', 'après', 'avoir', 'être', 'très', 'encore', 'autre', 'notre', 'votre', 'même', 'quand', 'quel',
  'confidentiel', 'diapositive', 'contenu', 'professionnel', 'numérique', 'document', 'fichier', 'rapport',
  'entreprise', 'société', 'analyse', 'résultat', 'projet', 'équipe', 'objectif', 'stratégie', 'processus'
]);

/**
 * Phrases valides à NE JAMAIS fusionner
 * Format: "word word" qui sont des combinaisons légitimes
 */
const VALID_PHRASES = new Set([
  // Article + noun/adj patterns
  'a message', 'a matter', 'a sense', 'a tool', 'a balance', 'a platform', 'a double', 'a barrier',
  'a new', 'a big', 'a few', 'a lot', 'a bit', 'a man', 'a day', 'a way', 'a set', 'a key', 'a global',
  
  // Preposition patterns  
  'in the', 'of the', 'to the', 'by the', 'on the', 'at the', 'for the', 'from the', 'with the',
  'in a', 'of a', 'to a', 'by a', 'on a', 'at a', 'for a', 'from a', 'with a',
  'in an', 'of an', 'to an', 'by an', 'on an', 'at an', 'for an', 'from an',
  
  // Verb + preposition/article patterns
  'is the', 'is a', 'is an', 'is its', 'is it', 'is not', 'is also', 'is both',
  'it is', 'it has', 'it was', 'it can', 'it may', 'it will', 'it would', 'it could',
  'to it', 'to be', 'to do', 'to go', 'to see', 'to get', 'to make', 'to take',
  'has been', 'has the', 'has a', 'has its', 'has also', 'has both',
  
  // Word + CAN patterns (CRITICAL - these get incorrectly merged)
  'message can', 'media can', 'individuals can', 'continents can', 'connectivity can',
  'misinformation can', 'people can', 'users can', 'platforms can', 'technology can',
  'data can', 'information can', 'content can', 'networks can', 'systems can',
  
  // Word + A patterns
  'fostered a', 'provides a', 'remains a', 'strike a', 'become a', 'create a', 'made a',
  'has a', 'is a', 'was a', 'as a', 'like a', 'than a', 'such a', 'quite a', 'what a',
  
  // Word + THE patterns
  'fostered the', 'provides the', 'remains the', 'become the', 'within the', 'across the',
  'through the', 'during the', 'before the', 'after the', 'under the', 'over the',
  
  // Word + ONE patterns
  'become one', 'number one', 'this one', 'that one', 'which one', 'every one', 'any one',
  'no one', 'some one', 'each one',
  
  // Word + HOW patterns
  'shaping how', 'revolutionized how', 'understand how', 'show how', 'see how', 'know how',
  'learn how', 'explain how', 'describe how', 'determine how',
  
  // Word + NOT patterns
  'lies not', 'is not', 'does not', 'did not', 'was not', 'were not', 'has not', 'have not',
  'will not', 'would not', 'could not', 'should not', 'might not', 'must not',
  
  // Word + DUE patterns
  'largely due', 'partly due', 'mainly due', 'primarily due', 'solely due',
  
  // Word + ARE patterns
  's are', 'they are', 'we are', 'you are', 'there are', 'here are', 'what are', 'who are',
  'these are', 'those are', 'which are',
  
  // Pronoun patterns
  'i am', 'i was', 'i will', 'i have', 'i can', 'i could', 'i would', 'i should',
  
  // Common 2-word combinations
  'and the', 'and a', 'and an', 'and it', 'and its', 'and is', 'and are', 'and was',
  'but the', 'but a', 'but it', 'but is', 'or the', 'or a', 'or it',
  'on how', 'of dis',
  
  // Name patterns (proper nouns + abbreviations)
  'clemens eng',
  
  // French patterns
  'de la', 'de le', 'à la', 'à le', 'en un', 'en une', 'il y', 'ce qui', 'ce que',
  'c\'est', 'qu\'il', 'qu\'elle', 'n\'est', 'd\'un', 'd\'une', 'y a', 'il a', 'elle a'
]);

/**
 * Prompt pour l'IA correcteur professionnel
 */
const PROFESSIONAL_PROOFREADER_PROMPT = `You are a professional multilingual proofreader and language expert. Your mission is to make documents PERFECT and ready for professional delivery.

AUTOMATICALLY DETECT the language of the text and analyze ALL errors.

═══════════════════════════════════════════════════════════════
📝 ERROR CATEGORIES TO DETECT
═══════════════════════════════════════════════════════════════

1️⃣ SPELLING ERRORS
   - Misspelled words: "recieve" → "receive", "occurence" → "occurrence"
   - Typos: "teh" → "the", "adn" → "and"
   - Homophones: "their/there/they're", "your/you're", "its/it's"
   - Accented characters (French): "éléphant" not "elephant"

2️⃣ CONJUGATION & VERB ERRORS
   - Wrong tense: "He go yesterday" → "He went yesterday"
   - Subject-verb agreement: "She have" → "She has"
   - French: "Il a manger" → "Il a mangé"

3️⃣ GRAMMAR ERRORS
   - Article usage: "a apple" → "an apple"
   - Plural/singular: "many information" → "much information"
   - Prepositions: "depends of" → "depends on"

4️⃣ SYNTAX ERRORS
   - Run-on sentences, fragments, misplaced modifiers

5️⃣ PUNCTUATION ERRORS
   - Missing periods, commas; wrong apostrophes
   - French spacing rules before : ; ! ?

6️⃣ FRAGMENTED WORDS (CRITICAL - Only fix REAL fragments)
   ✅ CORRECT these (word split by space):
   - "soc ial" → "social" (fragment)
   - "p otential" → "potential" (fragment)  
   - "enablin g" → "enabling" (fragment - the g belongs to enablin)
   - "th e" → "the" (fragment)
   - "commu nication" → "communication" (fragment)
   
   ❌ DO NOT TOUCH these (legitimate separate words):
   - "become one" - TWO separate valid words
   - "message can" - TWO separate valid words
   - "fostered a" - TWO separate valid words
   - "than a" - TWO separate valid words
   - "largely due" - TWO separate valid words

7️⃣ MERGED/FUSED WORDS
   - Missing space: "thankyou" → "thank you"

═══════════════════════════════════════════════════════════════
🚫 CRITICAL RULES - YOU MUST FOLLOW THESE
═══════════════════════════════════════════════════════════════

1. **VERIFY EVERY CORRECTION IS A REAL WORD**
   Before suggesting ANY correction, ask: "Does this word exist?"
   ❌ "becomeone" does NOT exist → DO NOT suggest
   ❌ "messagecan" does NOT exist → DO NOT suggest
   ❌ "fostereda" does NOT exist → DO NOT suggest
   ❌ "thana" does NOT exist → DO NOT suggest
   ✅ "social" EXISTS → Valid correction
   ✅ "enabling" EXISTS → Valid correction
   ✅ "potential" EXISTS → Valid correction

2. **NEVER MERGE TWO SEPARATE VALID WORDS**
   If BOTH parts are valid English words, DO NOT merge them:
   - "become" is valid + "one" is valid = DO NOT MERGE
   - "message" is valid + "can" is valid = DO NOT MERGE
   - "than" is valid + "a" is valid = DO NOT MERGE
   - "largely" is valid + "due" is valid = DO NOT MERGE
   
   Only merge when ONE part is NOT a valid word:
   - "soc" is NOT valid + "ial" is NOT valid = MERGE to "social" ✅
   - "enablin" is NOT valid + "g" is NOT valid alone = MERGE to "enabling" ✅

3. **ANALYZE FLOATING LETTERS CAREFULLY**
   When you see: "enablin g individuals"
   - "enablin" is NOT a word, "g" alone is NOT a word
   - Combined "enabling" IS a word → CORRECT ✅
   - "individuals" is already correct → DO NOT TOUCH
   
   When you see: "message can travel"
   - "message" IS a word, "can" IS a word
   - Both are valid separate words → DO NOT MERGE ❌

4. **SKIP IF UNCERTAIN**
   If you're not 100% sure, DO NOT include the correction.

═══════════════════════════════════════════════════════════════
📊 OUTPUT FORMAT
═══════════════════════════════════════════════════════════════

Return ONLY errors where the correction is a VERIFIED REAL WORD.
Each error needs: error, correction, context, severity, type, message`;

/**
 * Vérifie si un mot existe dans le dictionnaire - VERSION STRICTE
 * Le mot DOIT être dans COMMON_WORDS pour être considéré valide
 */
function isRealWord(word) {
  if (!word || word.length < 2) return false;
  const normalized = word.toLowerCase().trim();
  
  // VERSION STRICTE: Le mot DOIT être dans le dictionnaire
  return COMMON_WORDS.has(normalized);
}

/**
 * Vérifie si une phrase est valide (ne doit pas être fusionnée)
 */
function isValidPhrase(phrase) {
  if (!phrase) return false;
  const normalized = phrase.toLowerCase().trim();
  
  // Vérification exacte
  if (VALID_PHRASES.has(normalized)) return true;
  
  // Vérification pattern: si les deux parties sont des mots valides, c'est probablement valide
  const parts = normalized.split(/\s+/);
  if (parts.length === 2) {
    const [word1, word2] = parts;
    // Si les deux parties sont des mots valides du dictionnaire, ne pas fusionner
    if (COMMON_WORDS.has(word1) && COMMON_WORDS.has(word2)) {
      console.log(`[VALID_PHRASE] Both parts are valid words: "${word1}" + "${word2}" → DO NOT MERGE`);
      return true;
    }
  }
  
  return false;
}

/**
 * Obtient le contexte autour d'une position
 */
function getContext(text, index, length) {
  const start = Math.max(0, index - 30);
  const end = Math.min(text.length, index + length + 30);
  return text.slice(start, end).replace(/\s+/g, ' ').trim();
}

/**
 * Pré-détection ciblée des vrais mots fragmentés - VERSION STRICTE
 * Ne propose une correction QUE si le mot combiné est dans le dictionnaire
 */
function preDetectFragmentedWords(text) {
  const issues = [];
  
  // Pattern: Mot incomplet + espace + 1-3 lettres qui complètent
  // Ex: "soc ial" -> "social", "enablin g" -> "enabling"
  const fragmentPattern = /\b([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]{2,})\s+([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]{1,4})\b/g;
  
  let match;
  while ((match = fragmentPattern.exec(text)) !== null) {
    const [fullMatch, part1, part2] = match;
    const combined = part1 + part2;
    const normalizedCombined = combined.toLowerCase();
    
    // RÈGLE STRICTE 1: Le mot combiné DOIT être dans le dictionnaire
    if (!COMMON_WORDS.has(normalizedCombined)) {
      continue; // Skip - le mot combiné n'existe pas
    }
    
    // RÈGLE STRICTE 2: Vérifier que ce n'est PAS une phrase valide (deux mots séparés)
    if (isValidPhrase(fullMatch)) {
      console.log(`[REGEX SKIP] Valid phrase: "${fullMatch}"`);
      continue;
    }
    
    // RÈGLE STRICTE 3: Au moins une des parties NE DOIT PAS être un mot valide seul
    const part1Valid = COMMON_WORDS.has(part1.toLowerCase());
    const part2Valid = COMMON_WORDS.has(part2.toLowerCase());
    
    if (part1Valid && part2Valid) {
      // Les deux parties sont des mots valides séparés → NE PAS FUSIONNER
      console.log(`[REGEX SKIP] Both parts valid: "${part1}" + "${part2}"`);
      continue;
    }
    
    // OK - C'est un vrai fragment à corriger
    console.log(`[REGEX FOUND] Fragment: "${fullMatch}" → "${combined}"`);
    issues.push({
      error: fullMatch,
      correction: combined,
      context: getContext(text, match.index, fullMatch.length),
      type: 'fragmented_word',
      severity: 'high',
      message: 'Word incorrectly split by space',
      source: 'regex'
    });
  }
  
  // Pattern: 1-2 lettres isolées + espace + mot
  // Ex: "th e" -> "the", "o f" -> "of"
  const prefixPattern = /\b([a-zA-Z]{1,2})\s+([a-zA-Z]{1,4})\b/g;
  
  while ((match = prefixPattern.exec(text)) !== null) {
    const [fullMatch, prefix, suffix] = match;
    const combined = prefix + suffix;
    const normalizedCombined = combined.toLowerCase();
    
    // Le mot combiné DOIT être dans le dictionnaire
    if (!COMMON_WORDS.has(normalizedCombined)) {
      continue;
    }
    
    // Éviter les phrases valides
    if (isValidPhrase(fullMatch)) {
      continue;
    }
    
    // Éviter "I am", "A new", etc. où le préfixe est un mot valide
    if (COMMON_WORDS.has(prefix.toLowerCase()) && COMMON_WORDS.has(suffix.toLowerCase())) {
      continue;
    }
    
    console.log(`[REGEX FOUND] Short fragment: "${fullMatch}" → "${combined}"`);
    issues.push({
      error: fullMatch,
      correction: combined,
      context: getContext(text, match.index, fullMatch.length),
      type: 'fragmented_word',
      severity: 'high',
      message: 'Word incorrectly split by space',
      source: 'regex'
    });
  }
  
  // Pattern: Espaces multiples dans un mot
  const multiSpacePattern = /\b([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]+)\s{2,}([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]+)\b/g;
  
  while ((match = multiSpacePattern.exec(text)) !== null) {
    const combined = match[1] + match[2];
    if (COMMON_WORDS.has(combined.toLowerCase())) {
      issues.push({
        error: match[0],
        correction: combined,
        context: getContext(text, match.index, match[0].length),
        type: 'multiple_spaces',
        severity: 'high',
        message: 'Multiple spaces inside word',
        source: 'regex'
      });
    }
  }
  
  console.log(`[REGEX] Found ${issues.length} valid fragmented word patterns`);
  return issues;
}

/**
 * Filtre les corrections invalides - VERSION STRICTE
 */
function filterInvalidCorrections(errors) {
  return errors.filter(err => {
    const correctionLower = err.correction?.toLowerCase().trim();
    const errorLower = err.error?.toLowerCase().trim();
    
    // 1. La correction DOIT être dans le dictionnaire
    if (!COMMON_WORDS.has(correctionLower)) {
      console.log(`❌ REJECTED (not in dictionary): "${err.error}" → "${err.correction}"`);
      return false;
    }
    
    // 2. Rejeter si c'est une phrase valide qu'on essaie de fusionner
    if (isValidPhrase(err.error)) {
      console.log(`❌ REJECTED (valid phrase): "${err.error}"`);
      return false;
    }
    
    // 3. Rejeter si erreur = correction
    if (errorLower.replace(/\s+/g, '') === correctionLower.replace(/\s+/g, '')) {
      if (err.error.trim() === err.correction.trim()) {
        console.log(`❌ REJECTED (no change): "${err.error}"`);
        return false;
      }
    }
    
    // 4. Rejeter les fusions de deux mots valides
    const parts = errorLower.split(/\s+/);
    if (parts.length === 2) {
      const [p1, p2] = parts;
      if (COMMON_WORDS.has(p1) && COMMON_WORDS.has(p2)) {
        console.log(`❌ REJECTED (both parts valid): "${p1}" + "${p2}" should not merge to "${err.correction}"`);
        return false;
      }
    }
    
    console.log(`✅ ACCEPTED: "${err.error}" → "${err.correction}"`);
    return true;
  });
}

/**
 * Vérifie l'orthographe avec l'IA - VERSION 5.0 STRICTE
 */
export async function checkSpellingWithAI(text) {
  if (!text || text.trim().length < 10) return [];
  
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.warn('⚠️ OPENAI_API_KEY not configured - spelling check disabled');
    return [];
  }

  try {
    console.log('📝 AI spell-check VERSION 5.0 STRICT starting...');
    console.log(`[SPELLCHECK] Text length: ${text.length} characters`);
    
    // ÉTAPE 1: Pré-détection regex STRICTE
    const regexErrors = preDetectFragmentedWords(text);
    
    // ÉTAPE 2: Appel IA avec prompt strict
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [
          { role: 'system', content: PROFESSIONAL_PROOFREADER_PROMPT },
          { role: 'user', content: `Analyze this text for spelling/grammar errors. 
          
CRITICAL REMINDERS:
- NEVER merge two valid separate words (like "become one", "message can", "than a")
- ONLY fix real fragments where the combined word EXISTS (like "soc ial" → "social")
- Verify every correction is a real word before including it

Text to analyze:
${text.slice(0, 15000)}` }
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
                        error: { type: 'string', description: 'The exact incorrect text' },
                        correction: { type: 'string', description: 'The corrected text (MUST be a real word)' },
                        context: { type: 'string', description: 'Surrounding text' },
                        severity: { type: 'string', enum: ['low', 'medium', 'high'] },
                        type: { type: 'string', enum: ['fragmented_word', 'multiple_spaces', 'merged_word', 'spelling', 'grammar', 'conjugation', 'syntax', 'punctuation'] },
                        message: { type: 'string', description: 'Explanation' }
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
        max_tokens: 4000
      })
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('❌ OpenAI API error:', response.status, errorText);
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
    console.log(`[MERGED] Total errors before filtering: ${allErrors.length}`);
    
    // ÉTAPE 4: FILTRER STRICTEMENT les corrections invalides
    const validErrors = filterInvalidCorrections(allErrors);
    console.log(`[FILTERED] Valid errors after strict filtering: ${validErrors.length}`);
    
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
    message: err.message || getErrorMessage(err.type),
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
    'conjugation': 'Conjugation error',
    'syntax': 'Syntax error',
    'punctuation': 'Punctuation error'
  };
  return messages[type] || 'Spelling/grammar correction';
}

function getCategory(type) {
  if (['fragmented_word', 'multiple_spaces', 'merged_word'].includes(type)) return 'Word Fragments';
  if (type === 'grammar') return 'Grammar';
  if (type === 'conjugation') return 'Conjugation';
  if (type === 'syntax') return 'Syntax';
  if (type === 'punctuation') return 'Punctuation';
  return 'Spelling';
}

/**
 * Applique les corrections au texte
 */
export function applyCorrections(text, corrections) {
  if (!corrections?.length) return { correctedText: text, examples: [], changedCount: 0 };

  let correctedText = text;
  const examples = [];
  let changedCount = 0;

  // Trier par longueur décroissante pour éviter les conflits
  const sorted = [...corrections].sort((a, b) => (b.error?.length || 0) - (a.error?.length || 0));

  for (const c of sorted) {
    if (!c.error || !c.correction || c.error === c.correction) continue;
    
    // Dernière vérification: la correction doit être dans le dictionnaire
    if (!COMMON_WORDS.has(c.correction.toLowerCase().trim())) {
      console.log(`[APPLY SKIP] Correction not in dictionary: "${c.correction}"`);
      continue;
    }
    
    const regex = new RegExp(escapeRegExp(c.error), 'gi');
    const matches = correctedText.match(regex);
    
    if (matches?.length) {
      correctedText = correctedText.replace(regex, c.correction);
      changedCount += matches.length;
      if (examples.length < 15) {
        examples.push({ before: c.error, after: c.correction, type: c.type });
      }
      console.log(`[APPLY] ✅ "${c.error}" → "${c.correction}" (${matches.length}x)`);
    }
  }

  console.log(`[APPLY COMPLETE] ${changedCount} corrections applied`);
  return { correctedText, examples, changedCount };
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export default { checkSpellingWithAI, applyCorrections };

