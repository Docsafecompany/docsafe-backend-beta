// lib/aiSpellCheck.js
// Détection intelligente des fautes d'orthographe avec regroupement des mots fragmentés

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
            content: `Tu es un correcteur orthographique expert. Analyse le texte et détecte :
1. Les MOTS FRAGMENTÉS ou COUPÉS (ex: "Conf dential" → "Confidential", "sl ide" → "slide")
2. Les fautes d'orthographe classiques
3. Les erreurs de grammaire

RÈGLES IMPORTANTES :
- Si plusieurs mots consécutifs forment UN SEUL mot fragmenté, retourne UNE SEULE correction avec le segment complet
- Exemple : "Conf dential sl ide content" → correction unique : "Confidential slide content"
- Ne fais PAS de corrections mot par mot pour les fragments
- Retourne la phrase/segment COMPLET avant et après correction
- Langue : auto-détection (français, anglais, allemand, etc.)`
          },
          {
            role: 'user',
            content: `Analyse ce texte et retourne les corrections nécessaires :\n\n${text.slice(0, 8000)}`
          }
        ],
        tools: [
          {
            type: 'function',
            function: {
              name: 'report_spelling_errors',
              description: 'Report detected spelling and grammar errors with grouped corrections',
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
                          description: 'The incorrect text segment (can be multiple words if fragmented)' 
                        },
                        correction: { 
                          type: 'string', 
                          description: 'The correct text (joined/fixed)' 
                        },
                        context: { 
                          type: 'string', 
                          description: 'Surrounding text for context (20-30 chars before/after)' 
                        },
                        severity: { 
                          type: 'string', 
                          enum: ['low', 'medium', 'high'],
                          description: 'low=typo, medium=grammar, high=fragmented word or major error'
                        },
                        type: {
                          type: 'string',
                          enum: ['fragmented_word', 'spelling', 'grammar', 'punctuation'],
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
        temperature: 0.1, // Très déterministe
        max_tokens: 2000
      })
    });

    if (!response.ok) {
      const errorText = await response.text();
      console.error('❌ OpenAI API error:', response.status, errorText);
      return [];
    }

    const data = await response.json();
    const toolCall = data.choices?.[0]?.message?.tool_calls?.[0];
    
    if (!toolCall?.function?.arguments) {
      console.log('⚠️ No spelling errors detected by AI');
      return [];
    }

    const parsed = JSON.parse(toolCall.function.arguments);
    const errors = parsed.errors || [];

    // Formater pour le frontend
    const spellingErrors = errors.map((err, index) => ({
      id: `spell_${index}_${Date.now()}`,
      error: err.error,
      correction: err.correction,
      context: err.context || '',
      location: `AI Detection ${index + 1}`,
      severity: err.severity || 'medium',
      type: err.type || 'spelling',
      rule: err.type === 'fragmented_word' ? 'FRAGMENTED_WORD' : 'AI_CORRECTION',
      message: err.type === 'fragmented_word' 
        ? 'Fragmented words detected and merged' 
        : 'Spelling/grammar correction',
      category: err.type === 'fragmented_word' ? 'Word Fragments' : 'General'
    }));

    console.log(`✅ AI spell-check: ${spellingErrors.length} errors found`);
    return spellingErrors;

  } catch (error) {
    console.error('❌ AI spell-check error:', error.message);
    return [];
  }
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
  // (évite les conflits quand un segment en contient un autre)
  const sortedCorrections = [...corrections].sort(
    (a, b) => (b.error?.length || 0) - (a.error?.length || 0)
  );

  for (const correction of sortedCorrections) {
    if (!correction.error || !correction.correction) continue;
    if (correction.error === correction.correction) continue;

    // Recherche case-insensitive mais préserve la casse originale
    const regex = new RegExp(escapeRegExp(correction.error), 'gi');
    const matches = correctedText.match(regex);
    
    if (matches && matches.length > 0) {
      correctedText = correctedText.replace(regex, correction.correction);
      changedCount += matches.length;

      if (examples.length < 10) {
        examples.push({
          before: correction.error,
          after: correction.correction,
          type: correction.type || 'spelling'
        });
      }
    }
  }

  return { correctedText, examples, changedCount };
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export default { checkSpellingWithAI, applyCorrections };
