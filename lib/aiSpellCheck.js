// lib/aiSpellCheck.js
// VERSION 2.0 - Détection améliorée des mots fragmentés, espaces multiples et mots fusionnés

/**
 * Pré-détection regex des erreurs courantes avant l'appel IA
 * Garantit la détection des patterns évidents
 */
function preDetectPatterns(text) {
  const issues = [];
  
  // Pattern 1: Lettre isolée + espace + mot (ex: "c an", "o f", "th e")
  const singleLetterPattern = /\b([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ])\s+([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]{2,})\b/g;
  let match;
  while ((match = singleLetterPattern.exec(text)) !== null) {
    const combined = match[1] + match[2];
    // Éviter les faux positifs comme "I am", "A new"
    if (!isValidWord(match[0]) && isLikelyWord(combined)) {
      issues.push({
        error: match[0],
        correction: combined,
        context: getContext(text, match.index, match[0].length),
        type: 'fragmented_word',
        severity: 'high',
        source: 'regex'
      });
    }
  }
  
  // Pattern 2: Espaces multiples dans un mot (ex: "p  otential", "dis   connection")
  const multiSpacePattern = /\b([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]+)\s{2,}([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]+)\b/g;
  while ((match = multiSpacePattern.exec(text)) !== null) {
    issues.push({
      error: match[0],
      correction: match[1] + match[2],
      context: getContext(text, match.index, match[0].length),
      type: 'multiple_spaces',
      severity: 'high',
      source: 'regex'
    });
  }
  
  // Pattern 3: Deux lettres + espace + reste du mot (ex: "th e", "an d", "fo r")
  const twoLetterPattern = /\b([a-zA-Z]{2})\s+([a-zA-Z]{1,3})\b/g;
  while ((match = twoLetterPattern.exec(text)) !== null) {
    const combined = match[1] + match[2];
    // Vérifier si c'est probablement un mot fragmenté
    if (isLikelyWord(combined) && !isValidTwoWordPhrase(match[1], match[2])) {
      issues.push({
        error: match[0],
        correction: combined,
        context: getContext(text, match.index, match[0].length),
        type: 'fragmented_word',
        severity: 'high',
        source: 'regex'
      });
    }
  }
  
  console.log(`[REGEX PRE-DETECT] Found ${issues.length} pattern-based errors`);
  return issues;
}

/**
 * Obtient le contexte autour d'une erreur
 */
function getContext(text, index, length) {
  const start = Math.max(0, index - 20);
  const end = Math.min(text.length, index + length + 20);
  return text.slice(start, end).replace(/\s+/g, ' ').trim();
}

/**
 * Vérifie si un mot combiné est probablement valide
 */
function isLikelyWord(word) {
  const commonWords = [
    'the', 'and', 'for', 'are', 'but', 'not', 'you', 'all', 'can', 'had',
    'her', 'was', 'one', 'our', 'out', 'has', 'his', 'how', 'its', 'may',
    'potential', 'confidential', 'slide', 'content', 'professional', 'digital',
    'social', 'media', 'connection', 'disconnection', 'additionally', 'correction',
    'le', 'la', 'les', 'des', 'une', 'est', 'pour', 'dans', 'avec', 'sur',
    'confidentiel', 'diapositive', 'contenu', 'professionnel', 'numérique'
  ];
  return commonWords.includes(word.toLowerCase()) || word.length >= 4;
}

/**
 * Vérifie si c'est une phrase valide de deux mots
 */
function isValidTwoWordPhrase(word1, word2) {
  const validPhrases = [
    ['in', 'a'], ['to', 'be'], ['to', 'do'], ['at', 'a'], ['on', 'a'],
    ['is', 'a'], ['as', 'a'], ['by', 'a'], ['of', 'a'], ['an', 'a'],
    ['de', 'la'], ['à', 'la'], ['en', 'un'], ['il', 'y'], ['ce', 'qui']
  ];
  return validPhrases.some(([a, b]) => 
    a.toLowerCase() === word1.toLowerCase() && b.toLowerCase() === word2.toLowerCase()
  );
}

/**
 * Vérifie si une séquence est un mot valide (évite faux positifs)
 */
function isValidWord(phrase) {
  const validPatterns = [
    /^[AI]\s+(am|was|will|have|can|could|would|should|must|may|might)$/i,
    /^[A]\s+(new|big|few|lot|bit|man|day|way|set|key)$/i
  ];
  return validPatterns.some(p => p.test(phrase));
}

/**
 * Fusionne les erreurs regex et IA en évitant les doublons
 */
function mergeErrors(regexErrors, aiErrors) {
  const merged = [...regexErrors];
  
  for (const aiErr of aiErrors) {
    // Vérifier si cette erreur existe déjà (regex l'a trouvée)
    const isDuplicate = regexErrors.some(regexErr => {
      const regexNorm = regexErr.error.toLowerCase().replace(/\s+/g, '');
      const aiNorm = aiErr.error.toLowerCase().replace(/\s+/g, '');
      return regexNorm === aiNorm || 
             regexErr.error.toLowerCase().includes(aiErr.error.toLowerCase()) ||
             aiErr.error.toLowerCase().includes(regexErr.error.toLowerCase());
    });
    
    if (!isDuplicate) {
      merged.push(aiErr);
    }
  }
  
  console.log(`[MERGE] Regex: ${regexErrors.length}, AI: ${aiErrors.length}, Total unique: ${merged.length}`);
  return merged;
}

/**
 * Vérifie l'orthographe avec l'IA (détecte les mots fragmentés)
 * @param {string} text - Le texte à analyser
 * @returns {Promise<Array>} - Liste des erreurs groupées
 */
export async function checkSpellingWithAI(text) {
  if (!text || text.trim().length < 10) return [];
  
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.warn('⚠️ OPENAI_API_KEY not configured - spelling check disabled');
    return [];
  }

  try {
    console.log('📝 AI spell-check starting...');
    console.log(`[SPELLCHECK] Text length: ${text.length} characters`);
    
    // ÉTAPE 1: Pré-détection regex
    const regexErrors = preDetectPatterns(text);
    
    // ÉTAPE 2: Appel IA avec prompt amélioré
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
            content: `You are a professional multilingual spelling and grammar checker.
AUTOMATICALLY DETECT the language of the text and analyze in that language.

PRIORITY DETECTION - You MUST find these SPECIFIC error types:

1. ISOLATED LETTERS WITH SPACES:
   - Pattern: single letter + space(s) + rest of word
   - Examples: "c an"→"can", "o f"→"of", "th e"→"the", "t he"→"the", "a nd"→"and"

2. MULTIPLE SPACES INSIDE WORDS:
   - Pattern: word fragments separated by 2+ spaces
   - Examples: "p  otential"→"potential", "dis   connection"→"disconnection", "o  f"→"of"

3. FRAGMENTED WORDS (single space in middle):
   - Pattern: word incorrectly split by ONE space
   - Examples: "Conf dential"→"Confidential", "sl ide"→"slide", "Addit onally"→"Additionally"
   - Examples: "pro fessional"→"professional", "dig ital"→"digital", "soc ial"→"social"

4. MERGED/FUSED WORDS:
   - Pattern: two words incorrectly joined without space
   - Examples: "correctionnews"→"corrections", "darkside"→"dark side", "thankyou"→"thank you"

5. Regular spelling and grammar errors

IMPORTANT RULES:
- Scan the ENTIRE text carefully, character by character for space anomalies
- Return EACH error as a SINGLE correction
- If consecutive fragments form ONE word, return as ONE correction with the full segment
- Include 20-30 characters of surrounding context
- Set severity="high" for all fragmented/merged/multiple-space errors
- DO NOT miss any space anomalies - they are the PRIORITY`
          },
          {
            role: 'user',
            content: `Analyze this text and find ALL spelling errors, especially fragmented words with incorrect spaces:\n\n${text.slice(0, 15000)}`
          }
        ],
        tools: [
          {
            type: 'function',
            function: {
              name: 'report_spelling_errors',
              description: 'Report all detected spelling and grammar errors',
              parameters: {
                type: 'object',
                properties: {
                  errors: {
                    type: 'array',
                    items: {
                      type: 'object',
                      properties: {
                        error: { 
                          type: 'string', 
                          description: 'The exact incorrect text as it appears (include spaces)' 
                        },
                        correction: { 
                          type: 'string', 
                          description: 'The corrected text' 
                        },
                        context: { 
                          type: 'string', 
                          description: 'Surrounding text (20-30 chars before/after)' 
                        },
                        severity: { 
                          type: 'string', 
                          enum: ['low', 'medium', 'high'],
                          description: 'high=fragmented/merged words, medium=spelling, low=grammar/typo'
                        },
                        type: {
                          type: 'string',
                          enum: ['fragmented_word', 'multiple_spaces', 'merged_word', 'spelling', 'grammar', 'punctuation'],
                          description: 'Type of error detected'
                        }
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
        temperature: 0.1,
        max_tokens: 3000
      })
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('❌ OpenAI API error:', response.status, errorText);
      // Retourner au moins les erreurs regex si l'IA échoue
      return formatErrors(regexErrors);
    }

    const data = await response.json();
    const toolCall = data.choices?.[0]?.message?.tool_calls?.[0];
    
    let aiErrors = [];
    if (toolCall?.function?.arguments) {
      try {
        const parsed = JSON.parse(toolCall.function.arguments);
        aiErrors = parsed.errors || [];
        console.log(`[AI] Detected ${aiErrors.length} errors`);
      } catch (e) {
        console.error('❌ Failed to parse AI response:', e.message);
      }
    }
    
    // ÉTAPE 3: Fusionner regex + IA
    const mergedErrors = mergeErrors(regexErrors, aiErrors);
    
    // ÉTAPE 4: Formater pour le frontend/backend
    const spellingErrors = formatErrors(mergedErrors);

    console.log(`✅ AI spell-check complete: ${spellingErrors.length} total errors found`);
    
    // Log détaillé des erreurs trouvées
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
 * Formate les erreurs pour le frontend/backend
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
    rule: getRule(err.type),
    message: getMessage(err.type),
    category: getCategory(err.type),
    source: err.source || 'ai'
  }));
}

function getRule(type) {
  const rules = {
    'fragmented_word': 'FRAGMENTED_WORD',
    'multiple_spaces': 'MULTIPLE_SPACES',
    'merged_word': 'MERGED_WORD',
    'spelling': 'SPELLING',
    'grammar': 'GRAMMAR',
    'punctuation': 'PUNCTUATION'
  };
  return rules[type] || 'AI_CORRECTION';
}

function getMessage(type) {
  const messages = {
    'fragmented_word': 'Fragmented word detected - spaces incorrectly splitting a word',
    'multiple_spaces': 'Multiple spaces detected inside word',
    'merged_word': 'Merged words detected - missing space between words',
    'spelling': 'Spelling error detected',
    'grammar': 'Grammar error detected',
    'punctuation': 'Punctuation error detected'
  };
  return messages[type] || 'Spelling/grammar correction';
}

function getCategory(type) {
  const categories = {
    'fragmented_word': 'Word Fragments',
    'multiple_spaces': 'Word Fragments',
    'merged_word': 'Word Fragments',
    'spelling': 'Spelling',
    'grammar': 'Grammar',
    'punctuation': 'Punctuation'
  };
  return categories[type] || 'General';
}

/**
 * Applique les corrections au texte
 */
export function applyCorrections(text, corrections) {
  if (!corrections || corrections.length === 0) {
    return { correctedText: text, examples: [], changedCount: 0 };
  }

  let correctedText = text;
  const examples = [];
  let changedCount = 0;

  // Trier par longueur décroissante pour appliquer les plus longs d'abord
  const sortedCorrections = [...corrections].sort(
    (a, b) => (b.error?.length || 0) - (a.error?.length || 0)
  );

  for (const correction of sortedCorrections) {
    if (!correction.error || !correction.correction) continue;
    if (correction.error === correction.correction) continue;

    const regex = new RegExp(escapeRegExp(correction.error), 'gi');
    const matches = correctedText.match(regex);
    
    if (matches && matches.length > 0) {
      correctedText = correctedText.replace(regex, correction.correction);
      changedCount += matches.length;

      if (examples.length < 12) {
        examples.push({
          before: correction.error,
          after: correction.correction,
          type: correction.type || 'spelling'
        });
      }
      
      console.log(`[APPLY] "${correction.error}" → "${correction.correction}" (${matches.length}x)`);
    }
  }

  console.log(`[APPLY] Total corrections applied: ${changedCount}`);
  return { correctedText, examples, changedCount };
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Vérifie si les corrections ont été appliquées
 */
export function verifyCorrections(originalText, correctedText, spellingErrors) {
  const remaining = [];
  
  for (const err of spellingErrors) {
    if (correctedText.includes(err.error)) {
      remaining.push(err);
      console.warn(`⚠️ ERROR NOT FIXED: "${err.error}" still present`);
    }
  }
  
  const fixed = spellingErrors.length - remaining.length;
  console.log(`[VERIFY] Fixed: ${fixed}/${spellingErrors.length}, Remaining: ${remaining.length}`);
  
  return {
    success: remaining.length === 0,
    remainingErrors: remaining,
    totalFixed: fixed,
    totalErrors: spellingErrors.length
  };
}

export default { checkSpellingWithAI, applyCorrections, verifyCorrections };

