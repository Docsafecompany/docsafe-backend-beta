// lib/languagetool.js
// Détection des fautes d'orthographe et grammaire via LanguageTool API

import axios from 'axios';

/**
 * Vérifie l'orthographe et la grammaire d'un texte
 * @param {string} text - Le texte à analyser
 * @returns {Promise<Array>} - Liste des erreurs formatées pour le frontend
 */
export async function checkSpelling(text) {
  // Si pas de texte ou pas d'endpoint configuré, retourner vide
  if (!text || text.trim().length < 10) return [];
  if (!process.env.LT_ENDPOINT) {
    console.warn('⚠️ LT_ENDPOINT not configured - spelling check disabled');
    return [];
  }

  try {
    const endpoint = process.env.LT_ENDPOINT;
    const language = process.env.LT_LANGUAGE || 'fr'; // Français par défaut

    const { data } = await axios.post(
      endpoint,
      new URLSearchParams({ 
        text, 
        language,
        enabledOnly: 'false'
      }),
      {
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        timeout: 30000 // 30 secondes timeout
      }
    );

    if (!data.matches || !Array.isArray(data.matches)) {
      return [];
    }

    // Formater les résultats pour le frontend Lovable
    const spellingErrors = data.matches
      .map((match, index) => {
        const errorText = text.substring(match.offset, match.offset + match.length);
        const correction = match.replacements?.[0]?.value || '';
        
        // Ne garder que les erreurs avec une correction proposée
        if (!correction || correction === errorText) return null;

        // Extraire le contexte (30 caractères avant/après)
        const contextStart = Math.max(0, match.offset - 30);
        const contextEnd = Math.min(text.length, match.offset + match.length + 30);
        const contextText = text.substring(contextStart, contextEnd);

        // Déterminer la sévérité
        let severity = 'medium';
        if (match.rule?.issueType === 'misspelling') severity = 'high';
        if (match.rule?.issueType === 'typographical') severity = 'low';
        if (match.rule?.category?.id === 'TYPOS') severity = 'high';

        return {
          id: `spell_${index}_${Date.now()}`,
          error: errorText,
          correction: correction,
          context: contextText,
          location: `Position ${match.offset}`,
          severity: severity,
          rule: match.rule?.id || 'UNKNOWN',
          message: match.message || 'Spelling/grammar error detected',
          category: match.rule?.category?.name || 'General'
        };
      })
      .filter(Boolean); // Enlever les null

    console.log(`✅ LanguageTool: ${spellingErrors.length} spelling errors found`);
    return spellingErrors;

  } catch (error) {
    console.error('❌ LanguageTool API error:', error.message);
    return [];
  }
}

/**
 * Applique les corrections au texte
 * @param {string} text - Texte original
 * @param {Array} corrections - Liste des corrections à appliquer
 * @returns {Object} - { correctedText, examples, changedCount }
 */
export function applyCorrections(text, corrections) {
  if (!corrections || corrections.length === 0) {
    return { correctedText: text, examples: [], changedCount: 0 };
  }

  let correctedText = text;
  const examples = [];
  let changedCount = 0;
  let offset = 0;

  // Trier par position (du début à la fin)
  const sortedCorrections = [...corrections].sort((a, b) => {
    const posA = parseInt(a.location?.replace('Position ', '') || '0');
    const posB = parseInt(b.location?.replace('Position ', '') || '0');
    return posA - posB;
  });

  for (const correction of sortedCorrections) {
    if (!correction.error || !correction.correction) continue;

    const errorPos = correctedText.indexOf(correction.error);
    if (errorPos !== -1) {
      // Capturer le contexte avant/après
      const beforeContext = correctedText.substring(Math.max(0, errorPos - 20), errorPos);
      const afterContext = correctedText.substring(
        errorPos + correction.error.length,
        Math.min(correctedText.length, errorPos + correction.error.length + 20)
      );

      // Appliquer la correction
      correctedText = 
        correctedText.substring(0, errorPos) + 
        correction.correction + 
        correctedText.substring(errorPos + correction.error.length);

      // Ajouter un exemple (max 10)
      if (examples.length < 10) {
        examples.push({
          before: `...${beforeContext}[${correction.error}]${afterContext}...`,
          after: `...${beforeContext}[${correction.correction}]${afterContext}...`
        });
      }

      changedCount++;
    }
  }

  return { correctedText, examples, changedCount };
}

/**
 * Version simplifiée pour l'ancien format
 */
export async function ltCheck(text) {
  if (!process.env.LT_ENDPOINT) return null;
  const endpoint = process.env.LT_ENDPOINT;
  const language = process.env.LT_LANGUAGE || 'fr';
  const { data } = await axios.post(endpoint, new URLSearchParams({ text, language }));
  return data;
}

export default { checkSpelling, applyCorrections, ltCheck };
