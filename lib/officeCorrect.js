// lib/officeCorrect.js - VERSION 3.0 CORRIGÉE
// Correction des mots fragmentés dans PPTX/DOCX
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
 * Normalise le texte pour la comparaison (espaces multiples → un seul)
 */
function normalizeText(text) {
  return (text || '').replace(/\s+/g, ' ').trim();
}

/**
 * Supprime tous les espaces pour comparaison stricte
 */
function removeAllSpaces(text) {
  return (text || '').replace(/\s+/g, '');
}

/**
 * Applique les corrections de spellingErrors directement au XML PPTX
 * VERSION 3.0 - Corrige le problème de désynchronisation des espaces
 */
function applySpellingCorrectionsToXML(xml, spellingErrors, stats) {
  if (!spellingErrors || spellingErrors.length === 0) {
    console.log('[CORRECT PPTX] No spelling errors to apply');
    return xml;
  }
  
  console.log(`[CORRECT PPTX] Starting correction with ${spellingErrors.length} errors`);
  
  let modifiedXml = xml;
  
  // Trier par longueur décroissante (appliquer les plus longs d'abord)
  const sortedErrors = [...spellingErrors].sort(
    (a, b) => (b.error?.length || 0) - (a.error?.length || 0)
  );
  
  for (const err of sortedErrors) {
    const errorText = err.error;
    const correctionText = err.correction;
    
    if (!errorText || !correctionText || errorText === correctionText) {
      console.log(`[CORRECT PPTX] Skipping invalid error: "${errorText}" → "${correctionText}"`);
      continue;
    }
    
    console.log(`[CORRECT PPTX] Processing: "${errorText}" → "${correctionText}" (type: ${err.type})`);
    
    // ========================================
    // MÉTHODE 1 : Recherche directe dans le XML brut
    // ========================================
    const simplePattern = escapeRegExp(errorText);
    const simpleRegex = new RegExp(simplePattern, 'gi');
    
    if (simpleRegex.test(modifiedXml)) {
      const beforeCount = (modifiedXml.match(simpleRegex) || []).length;
      modifiedXml = modifiedXml.replace(simpleRegex, correctionText);
      if (beforeCount > 0) {
        stats.changedTextNodes += beforeCount;
        pushExample(stats.examples, errorText, correctionText);
        console.log(`[CORRECT PPTX] ✅ Method 1 - Applied simple: "${errorText}" → "${correctionText}" (${beforeCount}x)`);
      }
      continue;
    }
    
    // ========================================
    // MÉTHODE 2 : Recherche dans le texte extrait avec espaces
    // FIX: Ajouter des espaces entre les segments pour matcher l'extraction
    // ========================================
    const textContentRegex = /<a:t>([^<]*)<\/a:t>/g;
    let allText = '';
    let match;
    const textPositions = [];
    
    // Réinitialiser le regex
    textContentRegex.lastIndex = 0;
    
    while ((match = textContentRegex.exec(modifiedXml)) !== null) {
      const decodedText = decode(match[1]);
      textPositions.push({
        start: match.index,
        end: match.index + match[0].length,
        text: decodedText,
        fullMatch: match[0],
        textStart: allText.length, // Position dans allText
        textEnd: allText.length + decodedText.length
      });
      // FIX: Ajouter un espace entre les segments (comme extractTextFromXML)
      allText += decodedText + ' ';
    }
    
    // Normaliser le texte extrait (comme le fait extractTextFromXML)
    allText = allText.replace(/\s+/g, ' ').trim();
    
    console.log(`[CORRECT PPTX] Extracted text (${allText.length} chars): "${allText.slice(0, 200)}..."`);
    
    // Préparer les variantes de recherche
    const errorNormalized = normalizeText(errorText);
    const errorNoSpaces = removeAllSpaces(errorText);
    const allTextNormalized = normalizeText(allText);
    const allTextNoSpaces = removeAllSpaces(allText);
    
    console.log(`[CORRECT PPTX] Search variants:`);
    console.log(`  - errorNormalized: "${errorNormalized}"`);
    console.log(`  - errorNoSpaces: "${errorNoSpaces}"`);
    console.log(`  - allText contains normalized: ${allTextNormalized.toLowerCase().includes(errorNormalized.toLowerCase())}`);
    console.log(`  - allText contains noSpaces: ${allTextNoSpaces.toLowerCase().includes(errorNoSpaces.toLowerCase())}`);
    
    // Chercher avec la version normalisée (avec espaces)
    let errorIndex = allTextNormalized.toLowerCase().indexOf(errorNormalized.toLowerCase());
    let searchVariant = 'normalized';
    
    // Si pas trouvé, essayer sans espaces
    if (errorIndex === -1) {
      const noSpacesIndex = allTextNoSpaces.toLowerCase().indexOf(errorNoSpaces.toLowerCase());
      if (noSpacesIndex !== -1) {
        // Convertir l'index sans espaces vers l'index avec espaces
        // C'est approximatif mais devrait fonctionner pour la plupart des cas
        errorIndex = findIndexWithSpaces(allTextNormalized, errorNoSpaces, noSpacesIndex);
        searchVariant = 'noSpaces';
        console.log(`[CORRECT PPTX] Found with noSpaces variant at approx index ${errorIndex}`);
      }
    }
    
    if (errorIndex !== -1) {
      console.log(`[CORRECT PPTX] Found error at index ${errorIndex} (variant: ${searchVariant})`);
      
      // Trouver quels segments de texte sont concernés
      // Recalculer les positions dans le texte normalisé
      let currentTextPos = 0;
      const segmentRanges = [];
      
      for (let i = 0; i < textPositions.length; i++) {
        const pos = textPositions[i];
        const segmentText = normalizeText(pos.text);
        segmentRanges.push({
          ...pos,
          normalizedStart: currentTextPos,
          normalizedEnd: currentTextPos + segmentText.length
        });
        currentTextPos += segmentText.length + 1; // +1 pour l'espace
      }
      
      // Trouver les segments qui contiennent l'erreur
      const errorEnd = errorIndex + errorNormalized.length;
      let foundStart = false;
      let correctionApplied = false;
      
      for (let i = 0; i < segmentRanges.length; i++) {
        const seg = segmentRanges[i];
        
        // Ce segment chevauche-t-il l'erreur ?
        if (seg.normalizedEnd > errorIndex && seg.normalizedStart < errorEnd) {
          console.log(`[CORRECT PPTX] Segment ${i} overlaps error: "${seg.text}" (${seg.normalizedStart}-${seg.normalizedEnd})`);
          
          if (!foundStart) {
            // Premier segment - appliquer la correction complète ici
            foundStart = true;
            
            const errorStartInSegment = Math.max(0, errorIndex - seg.normalizedStart);
            const errorEndInSegment = Math.min(seg.text.length, errorEnd - seg.normalizedStart);
            
            // Construire le nouveau texte
            const beforeError = seg.text.substring(0, errorStartInSegment);
            const afterError = seg.text.substring(errorEndInSegment);
            const newText = beforeError + correctionText + afterError;
            
            console.log(`[CORRECT PPTX] Replacing in segment: "${seg.text}" → "${newText}"`);
            
            // Mettre à jour le XML
            const newTag = `<a:t>${encode(newText)}</a:t>`;
            modifiedXml = modifiedXml.substring(0, seg.start) + newTag + modifiedXml.substring(seg.end);
            
            // Ajuster les positions pour les segments suivants
            const lengthDiff = newTag.length - seg.fullMatch.length;
            for (let j = i + 1; j < segmentRanges.length; j++) {
              segmentRanges[j].start += lengthDiff;
              segmentRanges[j].end += lengthDiff;
            }
            
            stats.changedTextNodes++;
            pushExample(stats.examples, errorText, correctionText);
            correctionApplied = true;
            console.log(`[CORRECT PPTX] ✅ Method 2 - Applied fragmented: "${errorText}" → "${correctionText}"`);
            
          } else if (correctionApplied) {
            // Segments suivants qui faisaient partie de l'erreur - les vider
            console.log(`[CORRECT PPTX] Clearing subsequent segment: "${seg.text}"`);
            const newTag = `<a:t></a:t>`;
            modifiedXml = modifiedXml.substring(0, seg.start) + newTag + modifiedXml.substring(seg.end);
            
            const lengthDiff = newTag.length - seg.fullMatch.length;
            for (let j = i + 1; j < segmentRanges.length; j++) {
              segmentRanges[j].start += lengthDiff;
              segmentRanges[j].end += lengthDiff;
            }
          }
        }
      }
      
      if (!correctionApplied) {
        console.log(`[CORRECT PPTX] ⚠️ Could not apply correction for: "${errorText}"`);
      }
    } else {
      console.log(`[CORRECT PPTX] ⚠️ Error not found in text: "${errorText}"`);
    }
  }
  
  return modifiedXml;
}

/**
 * Trouve l'index approximatif dans le texte avec espaces
 * à partir d'un index dans le texte sans espaces
 */
function findIndexWithSpaces(textWithSpaces, searchNoSpaces, indexNoSpaces) {
  let noSpaceCount = 0;
  for (let i = 0; i < textWithSpaces.length; i++) {
    if (textWithSpaces[i] !== ' ') {
      if (noSpaceCount === indexNoSpaces) {
        return i;
      }
      noSpaceCount++;
    }
  }
  return -1;
}

/**
 * Applique les corrections de spellingErrors au XML DOCX
 */
function applySpellingCorrectionsToDocxXML(xml, spellingErrors, stats) {
  if (!spellingErrors || spellingErrors.length === 0) {
    console.log('[CORRECT DOCX] No spelling errors to apply');
    return xml;
  }
  
  console.log(`[CORRECT DOCX] Starting correction with ${spellingErrors.length} errors`);
  
  let modifiedXml = xml;
  
  const sortedErrors = [...spellingErrors].sort(
    (a, b) => (b.error?.length || 0) - (a.error?.length || 0)
  );
  
  for (const err of sortedErrors) {
    const errorText = err.error;
    const correctionText = err.correction;
    
    if (!errorText || !correctionText || errorText === correctionText) continue;
    
    console.log(`[CORRECT DOCX] Processing: "${errorText}" → "${correctionText}"`);
    
    // Méthode 1 : Recherche directe
    const simplePattern = escapeRegExp(errorText);
    const simpleRegex = new RegExp(simplePattern, 'gi');
    
    if (simpleRegex.test(modifiedXml)) {
      const beforeCount = (modifiedXml.match(simpleRegex) || []).length;
      modifiedXml = modifiedXml.replace(simpleRegex, correctionText);
      if (beforeCount > 0) {
        stats.changedTextNodes += beforeCount;
        pushExample(stats.examples, errorText, correctionText);
        console.log(`[CORRECT DOCX] ✅ Applied: "${errorText}" → "${correctionText}" (${beforeCount}x)`);
      }
      continue;
    }
    
    // Méthode 2 : Recherche dans le texte extrait (pour DOCX, similaire à PPTX)
    const textContentRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let allText = '';
    let match;
    const textPositions = [];
    
    textContentRegex.lastIndex = 0;
    
    while ((match = textContentRegex.exec(modifiedXml)) !== null) {
      const decodedText = decode(match[1]);
      textPositions.push({
        start: match.index,
        end: match.index + match[0].length,
        text: decodedText,
        fullMatch: match[0],
        textStart: allText.length,
        textEnd: allText.length + decodedText.length
      });
      allText += decodedText + ' ';
    }
    
    allText = allText.replace(/\s+/g, ' ').trim();
    
    const errorNormalized = normalizeText(errorText);
    const allTextNormalized = normalizeText(allText);
    
    const errorIndex = allTextNormalized.toLowerCase().indexOf(errorNormalized.toLowerCase());
    
    if (errorIndex !== -1) {
      // Recalculer les positions
      let currentTextPos = 0;
      const segmentRanges = [];
      
      for (let i = 0; i < textPositions.length; i++) {
        const pos = textPositions[i];
        const segmentText = normalizeText(pos.text);
        segmentRanges.push({
          ...pos,
          normalizedStart: currentTextPos,
          normalizedEnd: currentTextPos + segmentText.length
        });
        currentTextPos += segmentText.length + 1;
      }
      
      const errorEnd = errorIndex + errorNormalized.length;
      let foundStart = false;
      let correctionApplied = false;
      
      for (let i = 0; i < segmentRanges.length; i++) {
        const seg = segmentRanges[i];
        
        if (seg.normalizedEnd > errorIndex && seg.normalizedStart < errorEnd) {
          if (!foundStart) {
            foundStart = true;
            
            const errorStartInSegment = Math.max(0, errorIndex - seg.normalizedStart);
            const errorEndInSegment = Math.min(seg.text.length, errorEnd - seg.normalizedStart);
            
            const beforeError = seg.text.substring(0, errorStartInSegment);
            const afterError = seg.text.substring(errorEndInSegment);
            const newText = beforeError + correctionText + afterError;
            
            // Extraire les attributs du tag original
            const attrMatch = seg.fullMatch.match(/<w:t([^>]*)>/);
            const attrs = attrMatch ? attrMatch[1] : '';
            
            const newTag = `<w:t${attrs}>${encode(newText)}</w:t>`;
            modifiedXml = modifiedXml.substring(0, seg.start) + newTag + modifiedXml.substring(seg.end);
            
            const lengthDiff = newTag.length - seg.fullMatch.length;
            for (let j = i + 1; j < segmentRanges.length; j++) {
              segmentRanges[j].start += lengthDiff;
              segmentRanges[j].end += lengthDiff;
            }
            
            stats.changedTextNodes++;
            pushExample(stats.examples, errorText, correctionText);
            correctionApplied = true;
            console.log(`[CORRECT DOCX] ✅ Applied fragmented: "${errorText}" → "${correctionText}"`);
            
          } else if (correctionApplied) {
            const attrMatch = seg.fullMatch.match(/<w:t([^>]*)>/);
            const attrs = attrMatch ? attrMatch[1] : '';
            const newTag = `<w:t${attrs}></w:t>`;
            modifiedXml = modifiedXml.substring(0, seg.start) + newTag + modifiedXml.substring(seg.end);
            
            const lengthDiff = newTag.length - seg.fullMatch.length;
            for (let j = i + 1; j < segmentRanges.length; j++) {
              segmentRanges[j].start += lengthDiff;
              segmentRanges[j].end += lengthDiff;
            }
          }
        }
      }
      
      if (!correctionApplied) {
        console.log(`[CORRECT DOCX] ⚠️ Could not apply correction for: "${errorText}"`);
      }
    } else {
      console.log(`[CORRECT DOCX] ⚠️ Error not found in text: "${errorText}"`);
    }
  }
  
  return modifiedXml;
}

// ===================================================================
// DOCX TEXT CORRECTION
// ===================================================================
export async function correctDOCXText(buffer, correctFn, options = {}) {
  console.log('[DOCX] Starting text correction...');
  
  const zip = await JSZip.loadAsync(buffer);
  const targets = ['word/document.xml', ...Object.keys(zip.files).filter(k => /word\/(header|footer)\d+\.xml$/.test(k))];
  const stats = { totalTextNodes: 0, changedTextNodes: 0, examples: [] };
  
  const spellingErrors = options.spellingErrors || [];
  console.log(`[DOCX] Received ${spellingErrors.length} spelling errors to apply`);

  for (const p of targets) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async('string');

    // ÉTAPE 1 : Appliquer les spellingErrors détectés (prioritaire)
    if (spellingErrors.length > 0) {
      console.log(`[DOCX] Applying spelling corrections to ${p}`);
      xml = applySpellingCorrectionsToDocxXML(xml, spellingErrors, stats);
    }

    // ÉTAPE 2 : Correction IA node par node (si pas de spellingErrors ou en complément)
    if (!spellingErrors.length || options.alsoRunAI) {
      console.log(`[DOCX] Running AI correction on ${p}`);
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
  
  console.log(`[DOCX] Correction complete: ${stats.changedTextNodes} changes made`);
  return { outBuffer: await zip.generateAsync({ type: 'nodebuffer' }), stats };
}

// ===================================================================
// PPTX TEXT CORRECTION
// ===================================================================
export async function correctPPTXText(buffer, correctFn, options = {}) {
  console.log('[PPTX] Starting text correction...');
  
  const zip = await JSZip.loadAsync(buffer);
  const slides = Object.keys(zip.files).filter(k => /^ppt\/slides\/slide\d+\.xml$/.test(k));
  const stats = { totalTextNodes: 0, changedTextNodes: 0, examples: [] };
  
  const spellingErrors = options.spellingErrors || [];
  console.log(`[PPTX] Received ${spellingErrors.length} spelling errors to apply`);
  
  // Log les erreurs reçues pour debug
  if (spellingErrors.length > 0) {
    console.log('[PPTX] Spelling errors to apply:');
    spellingErrors.forEach((err, i) => {
      console.log(`  ${i + 1}. "${err.error}" → "${err.correction}" (type: ${err.type})`);
    });
  }

  for (const sp of slides) {
    let xml = await zip.file(sp).async('string');

    // ÉTAPE 1 : Appliquer les spellingErrors détectés (prioritaire)
    if (spellingErrors.length > 0) {
      console.log(`[PPTX] Applying spelling corrections to ${sp}`);
      xml = applySpellingCorrectionsToXML(xml, spellingErrors, stats);
    }

    // ÉTAPE 2 : Correction IA node par node (si pas de spellingErrors ou en complément)
    if (!spellingErrors.length || options.alsoRunAI) {
      console.log(`[PPTX] Running AI correction on ${sp}`);
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
  
  console.log(`[PPTX] Correction complete: ${stats.changedTextNodes} changes made`);
  return { outBuffer: await zip.generateAsync({ type: 'nodebuffer' }), stats };
}

export default { correctDOCXText, correctPPTXText };
