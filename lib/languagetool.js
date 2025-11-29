// lib/languagetool.js
import axios from 'axios';

/**
 * Check text for spelling/grammar errors using LanguageTool
 * @param {string} text - Text to check
 * @param {string} language - Language code (default: 'en-US')
 * @returns {Promise<Array>} - Array of spelling errors formatted for frontend
 */
export async function checkSpelling(text, language = 'en-US') {
  if (!process.env.LT_ENDPOINT || !text || text.trim().length < 10) {
    return [];
  }

  try {
    const endpoint = process.env.LT_ENDPOINT;
    const lang = process.env.LT_LANGUAGE || language;

    const { data } = await axios.post(
      endpoint,
      new URLSearchParams({ text, language: lang }),
      { timeout: 30000 }
    );

    if (!data?.matches || !Array.isArray(data.matches)) {
      return [];
    }

    // Transform LanguageTool matches to frontend format
    return data.matches.map((match, index) => {
      const error = text.substring(match.offset, match.offset + match.length);
      const correction = match.replacements?.[0]?.value || '';
      
      // Extract context (30 chars before and after)
      const contextStart = Math.max(0, match.offset - 30);
      const contextEnd = Math.min(text.length, match.offset + match.length + 30);
      const context = text.substring(contextStart, contextEnd);

      // Determine severity based on rule category
      let severity = 'low';
      const category = match.rule?.category?.id || '';
      if (category === 'TYPOS' || category === 'MISSPELLING') {
        severity = 'medium';
      } else if (category === 'GRAMMAR' || category === 'PUNCTUATION') {
        severity = 'low';
      } else if (category === 'STYLE' || category === 'REDUNDANCY') {
        severity = 'low';
      }

      return {
        id: `spelling_${index}`,
        error: error,
        correction: correction,
        context: context,
        location: `Character ${match.offset}`,
        severity: severity,
        // Additional info for report
        message: match.message || '',
        rule: match.rule?.id || '',
        category: category
      };
    }).filter(item => item.error && item.correction && item.error !== item.correction);

  } catch (error) {
    console.error('LanguageTool check failed:', error.message);
    return [];
  }
}

/**
 * Apply spelling corrections to text
 * @param {string} text - Original text
 * @param {Array} corrections - Array of corrections with offset info
 * @returns {Object} - { correctedText, examples }
 */
export function applyCorrections(text, corrections) {
  if (!corrections || corrections.length === 0) {
    return { correctedText: text, examples: [] };
  }

  // Sort by offset descending to apply from end to start (avoid offset shift)
  const sorted = [...corrections].sort((a, b) => {
    const offsetA = parseInt(a.location?.replace('Character ', '') || '0');
    const offsetB = parseInt(b.location?.replace('Character ', '') || '0');
    return offsetB - offsetA;
  });

  let correctedText = text;
  const examples = [];

  for (const corr of sorted) {
    const offset = parseInt(corr.location?.replace('Character ', '') || '-1');
    if (offset >= 0 && corr.error && corr.correction) {
      const before = correctedText.substring(
        Math.max(0, offset - 20),
        offset + corr.error.length + 20
      );
      
      correctedText = 
        correctedText.substring(0, offset) + 
        corr.correction + 
        correctedText.substring(offset + corr.error.length);

      const after = correctedText.substring(
        Math.max(0, offset - 20),
        offset + corr.correction.length + 20
      );

      examples.push({
        before: `...${before}...`,
        after: `...${after}...`,
        error: corr.error,
        correction: corr.correction
      });
    }
  }

  return { 
    correctedText, 
    examples: examples.slice(0, 10) // Limit to 10 examples for report
  };
}

export default { checkSpelling, applyCorrections };
