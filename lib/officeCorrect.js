// lib/officeCorrect.js - VERSION CORRIGÉE
import JSZip from 'jszip';

const decode = s => s
  .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
  .replace(/&quot;/g, '"').replace(/&apos;/g, "'");
const encode = s => s
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&apos;');

function pushExample(examples, before, after, max = 12) {
  if (examples.length >= max) return;
  const b = (before || '').slice(0, 140);
  const a = (after || '').slice(0, 140);
  if (b !== a) examples.push({ before: b, after: a });
}

/**
 * Échappe les caractères spéciaux regex
 */
function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Crée un pattern flexible pour les mots fragmentés
 * "Confi dential" → "Confi[\\s\\S]*?dential" pour matcher même avec des tags entre
 */
function createFlexiblePattern(errorText) {
  // Échapper les caractères spéciaux
  const escaped = escapeRegExp(errorText);
  
  // Remplacer les espaces par un pattern qui accepte :
  // - espaces normaux
  // - tags XML entre les caractères (pour les fragments)
  // - caractères unicode spéciaux
  const flexible = escaped
    .split(/\s+/)
    .map(word => {
      // Pour chaque mot, permettre des tags XML entre les caractères
      return word.split('').join('(?:</a:t>\\s*<a:t[^>]*>)?');
    })
    .join('(?:</a:t>\\s*<a:t[^>]*>)?\\s*(?:</a:t>\\s*<a:t[^>]*>)?');
  
  return flexible;
}

/**
 * Applique les corrections de spellingErrors directement au XML
 */
function applySpellingCorrectionsToXML(xml, spellingErrors, stats) {
  if (!spellingErrors || spellingErrors.length === 0) return xml;
  
  let modifiedXml = xml;
  
  // Trier par longueur décroissante (appliquer les plus longs d'abord)
  const sortedErrors = [...spellingErrors].sort(
    (a, b) => (b.error?.length || 0) - (a.error?.length || 0)
  );
  
  for (const err of sortedErrors) {
    const errorText = err.error;
    const correctionText = err.correction;
    
    if (!errorText || !correctionText || errorText === correctionText) continue;
    
    // Méthode 1 : Recherche directe dans le texte (pour les erreurs simples)
    const simplePattern = escapeRegExp(errorText);
    const simpleRegex = new RegExp(simplePattern, 'gi');
    
    if (simpleRegex.test(modifiedXml)) {
      const beforeCount = (modifiedXml.match(simpleRegex) || []).length;
      modifiedXml = modifiedXml.replace(simpleRegex, correctionText);
      if (beforeCount > 0) {
        stats.changedTextNodes += beforeCount;
        pushExample(stats.examples, errorText, correctionText);
        console.log(`[CORRECT] Applied simple: "${errorText}" → "${correctionText}" (${beforeCount}x)`);
      }
      continue;
    }
    
    // Méthode 2 : Recherche dans le texte extrait seulement (entre <a:t> tags)
    // Extraire tout le texte, faire le remplacement, puis reconstruire
    const textContentRegex = /<a:t>([^<]*)<\/a:t>/g;
    let allText = '';
    let match;
    const textPositions = [];
    
    while ((match = textContentRegex.exec(modifiedXml)) !== null) {
      textPositions.push({
        start: match.index,
        end: match.index + match[0].length,
        text: decode(match[1]),
        fullMatch: match[0]
      });
      allText += decode(match[1]);
    }
    
    // Chercher l'erreur dans le texte concaténé
    const errorInText = allText.toLowerCase().includes(errorText.toLowerCase());
    
    if (errorInText) {
      // L'erreur existe dans le texte combiné - on doit trouver et corriger
      // Stratégie : remplacer dans chaque segment de texte qui contient une partie
      
      const errorLower = errorText.toLowerCase();
      const allTextLower = allText.toLowerCase();
      const errorIndex = allTextLower.indexOf(errorLower);
      
      if (errorIndex !== -1) {
        // Trouver quels segments de texte sont concernés
        let currentPos = 0;
        let foundStart = false;
        let correctionApplied = false;
        
        for (let i = 0; i < textPositions.length; i++) {
          const pos = textPositions[i];
          const segmentStart = currentPos;
          const segmentEnd = currentPos + pos.text.length;
          
          // Ce segment contient-il une partie de l'erreur ?
          if (segmentEnd > errorIndex && segmentStart < errorIndex + errorText.length) {
            if (!foundStart) {
              // Premier segment contenant l'erreur - appliquer la correction ici
              foundStart = true;
              
              // Calculer quelle partie de l'erreur est dans ce segment
              const errorStartInSegment = Math.max(0, errorIndex - segmentStart);
              const errorEndInSegment = Math.min(pos.text.length, errorIndex + errorText.length - segmentStart);
              
              // Remplacer dans ce segment
              const newText = pos.text.substring(0, errorStartInSegment) + 
                             correctionText + 
                             pos.text.substring(errorEndInSegment);
              
              // Mettre à jour le XML
              const newTag = `<a:t>${encode(newText)}</a:t>`;
              modifiedXml = modifiedXml.substring(0, pos.start) + newTag + modifiedXml.substring(pos.end);
              
              // Ajuster les positions suivantes
              const lengthDiff = newTag.length - pos.fullMatch.length;
              for (let j = i + 1; j < textPositions.length; j++) {
                textPositions[j].start += lengthDiff;
                textPositions[j].end += lengthDiff;
              }
              
              stats.changedTextNodes++;
              pushExample(stats.examples, errorText, correctionText);
              correctionApplied = true;
              console.log(`[CORRECT] Applied fragmented: "${errorText}" → "${correctionText}"`);
              
            } else if (!correctionApplied) {
              // Segments suivants qui contenaient des parties de l'erreur - les vider
              const newTag = `<a:t></a:t>`;
              modifiedXml = modifiedXml.substring(0, pos.start) + newTag + modifiedXml.substring(pos.end);
              
              const lengthDiff = newTag.length - pos.fullMatch.length;
              for (let j = i + 1; j < textPositions.length; j++) {
                textPositions[j].start += lengthDiff;
                textPositions[j].end += lengthDiff;
              }
            }
          }
          
          currentPos = segmentEnd;
        }
      }
    }
  }
  
  return modifiedXml;
}

/**
 * Applique les corrections de spellingErrors au XML DOCX
 */
function applySpellingCorrectionsToDocxXML(xml, spellingErrors, stats) {
  if (!spellingErrors || spellingErrors.length === 0) return xml;
  
  let modifiedXml = xml;
  
  const sortedErrors = [...spellingErrors].sort(
    (a, b) => (b.error?.length || 0) - (a.error?.length || 0)
  );
  
  for (const err of sortedErrors) {
    const errorText = err.error;
    const correctionText = err.correction;
    
    if (!errorText || !correctionText || errorText === correctionText) continue;
    
    // Recherche directe
    const simplePattern = escapeRegExp(errorText);
    const simpleRegex = new RegExp(simplePattern, 'gi');
    
    if (simpleRegex.test(modifiedXml)) {
      const beforeCount = (modifiedXml.match(simpleRegex) || []).length;
      modifiedXml = modifiedXml.replace(simpleRegex, correctionText);
      if (beforeCount > 0) {
        stats.changedTextNodes += beforeCount;
        pushExample(stats.examples, errorText, correctionText);
        console.log(`[CORRECT DOCX] Applied: "${errorText}" → "${correctionText}" (${beforeCount}x)`);
      }
    }
  }
  
  return modifiedXml;
}

// ===================================================================
// DOCX TEXT CORRECTION - MISE À JOUR
// ===================================================================
export async function correctDOCXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);
  const targets = ['word/document.xml', ...Object.keys(zip.files).filter(k => /word\/(header|footer)\d+\.xml$/.test(k))];
  const stats = { totalTextNodes: 0, changedTextNodes: 0, examples: [] };
  
  const spellingErrors = options.spellingErrors || [];

  for (const p of targets) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async('string');

    // ÉTAPE 1 : Appliquer les spellingErrors détectés (prioritaire)
    if (spellingErrors.length > 0) {
      console.log(`[DOCX] Applying ${spellingErrors.length} spelling corrections to ${p}`);
      xml = applySpellingCorrectionsToDocxXML(xml, spellingErrors, stats);
    }

    // ÉTAPE 2 : Correction IA node par node (si pas de spellingErrors ou en complément)
    if (!spellingErrors.length || options.alsoRunAI) {
      const chunks = [];
      let lastIndex = 0;
      const re = /<w:t([^>]*)>([\s\S]*?)<\/w:t>/g;
      let m;
      while ((m = re.exec(xml)) !== null) {
        const [full, attrs, inner] = m;
        const start = m.index;
        chunks.push(xml.slice(lastIndex, start));
        const original = decode(inner);
        const corrected = await correctFn(original);
        stats.totalTextNodes++;
        if (corrected !== original) {
          stats.changedTextNodes++;
          pushExample(stats.examples, original, corrected);
        }
        chunks.push(`<w:t${attrs}>${encode(corrected)}</w:t>`);
        lastIndex = start + full.length;
      }
      chunks.push(xml.slice(lastIndex));
      xml = chunks.join('');
    }

    zip.file(p, xml);
  }
  
  return { outBuffer: await zip.generateAsync({ type: 'nodebuffer' }), stats };
}

// ===================================================================
// PPTX TEXT CORRECTION - MISE À JOUR
// ===================================================================
export async function correctPPTXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);
  const slides = Object.keys(zip.files).filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k));
  const stats = { totalTextNodes: 0, changedTextNodes: 0, examples: [] };
  
  const spellingErrors = options.spellingErrors || [];

  for (const sp of slides) {
    let xml = await zip.file(sp).async('string');

    // ÉTAPE 1 : Appliquer les spellingErrors détectés (prioritaire)
    if (spellingErrors.length > 0) {
      console.log(`[PPTX] Applying ${spellingErrors.length} spelling corrections to ${sp}`);
      xml = applySpellingCorrectionsToXML(xml, spellingErrors, stats);
    }

    // ÉTAPE 2 : Correction IA node par node (si pas de spellingErrors ou en complément)
    if (!spellingErrors.length || options.alsoRunAI) {
      const chunks = [];
      let lastIndex = 0;
      const re = /<a:t>([\s\S]*?)<\/a:t>/g;
      let m;
      while ((m = re.exec(xml)) !== null) {
        const [full, inner] = m;
        const start = m.index;
        chunks.push(xml.slice(lastIndex, start));
        const original = decode(inner);
        const corrected = await correctFn(original);
        stats.totalTextNodes++;
        if (corrected !== original) {
          stats.changedTextNodes++;
          pushExample(stats.examples, original, corrected);
        }
        chunks.push(`<a:t>${encode(corrected)}</a:t>`);
        lastIndex = start + full.length;
      }
      chunks.push(xml.slice(lastIndex));
      xml = chunks.join('');
    }

    zip.file(sp, xml);
  }
  
  return { outBuffer: await zip.generateAsync({ type: 'nodebuffer' }), stats };
}

export default { correctDOCXText, correctPPTXText };
