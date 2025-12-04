// lib/officeCorrect.js - VERSION 3.2 CORRIGÉE
// Correction des mots fragmentés dans PPTX/DOCX
// FIX v3.2: Préservation correcte du texte après l'erreur dans le dernier segment
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

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeText(text) {
  return (text || '').replace(/\s+/g, ' ').trim();
}

function removeAllSpaces(text) {
  return (text || '').replace(/\s+/g, '');
}

/**
 * Convertit un index dans le texte sans espaces vers l'index dans le texte original
 */
function convertNoSpacesIndexToOriginal(originalText, noSpacesIndex) {
  let noSpaceCount = 0;
  for (let i = 0; i < originalText.length; i++) {
    if (!/\s/.test(originalText[i])) {
      if (noSpaceCount === noSpacesIndex) {
        return i;
      }
      noSpaceCount++;
    }
  }
  return -1;
}

/**
 * Trouve la longueur de l'erreur dans le texte original (avec espaces)
 */
function findErrorLengthInOriginal(originalText, startIndex, errorNoSpaces) {
  let matched = 0;
  let i = startIndex;
  while (i < originalText.length && matched < errorNoSpaces.length) {
    if (/\s/.test(originalText[i])) {
      i++;
      continue;
    }
    if (originalText[i].toLowerCase() === errorNoSpaces[matched].toLowerCase()) {
      matched++;
    }
    i++;
  }
  return i - startIndex;
}

/**
 * Calcule la fin réelle de l'erreur en tenant compte des caractères à conserver
 * FIX v3.2: Retourne aussi l'index du dernier caractère de l'erreur dans chaque segment
 */
function calculateErrorEndDetails(textPositions, errorIndex, errorEnd) {
  const details = {
    lastSegmentIndex: -1,
    errorEndInLastSegment: 0,
    segmentsToEmpty: []
  };
  
  for (let i = 0; i < textPositions.length; i++) {
    const seg = textPositions[i];
    const overlaps = seg.allTextEnd > errorIndex && seg.allTextStart < errorEnd;
    
    if (overlaps) {
      if (errorEnd <= seg.allTextEnd) {
        // Ce segment contient la fin de l'erreur
        details.lastSegmentIndex = i;
        details.errorEndInLastSegment = errorEnd - seg.allTextStart;
      } else if (seg.allTextEnd <= errorEnd) {
        // Ce segment est entièrement dans l'erreur
        details.segmentsToEmpty.push(i);
      }
    }
  }
  
  return details;
}

/**
 * Applique les corrections de spellingErrors directement au XML PPTX
 * VERSION 3.2 - FIX: Préservation correcte du texte après l'erreur
 */
function applySpellingCorrectionsToXML(xml, spellingErrors, stats) {
  if (!spellingErrors || spellingErrors.length === 0) {
    console.log('[CORRECT PPTX] No spelling errors to apply');
    return xml;
  }
  
  console.log(`[CORRECT PPTX] Starting correction with ${spellingErrors.length} errors`);
  
  let modifiedXml = xml;
  
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
    // MÉTHODE 2 : Recherche dans le texte extrait (mots fragmentés)
    // FIX v3.2: Gestion correcte des segments multiples
    // ========================================
    const textContentRegex = /<a:t>([^<]*)<\/a:t>/g;
    let allText = '';
    let match;
    const textPositions = [];
    
    textContentRegex.lastIndex = 0;
    
    while ((match = textContentRegex.exec(modifiedXml)) !== null) {
      const decodedText = decode(match[1]);
      const startInAllText = allText.length;
      allText += decodedText;
      
      textPositions.push({
        xmlStart: match.index,
        xmlEnd: match.index + match[0].length,
        text: decodedText,
        fullMatch: match[0],
        allTextStart: startInAllText,
        allTextEnd: allText.length
      });
    }
    
    console.log(`[CORRECT PPTX] Extracted text (${allText.length} chars): "${allText.slice(0, 200)}..."`);
    console.log(`[CORRECT PPTX] Found ${textPositions.length} text segments`);
    
    // Préparer les variantes de recherche
    const errorNormalized = normalizeText(errorText);
    const errorNoSpaces = removeAllSpaces(errorText);
    const allTextNoSpaces = removeAllSpaces(allText);
    
    // Chercher l'erreur dans allText
    let errorIndex = allText.toLowerCase().indexOf(errorNormalized.toLowerCase());
    let searchVariant = 'direct';
    let effectiveErrorLength = errorNormalized.length;
    
    // Si pas trouvé directement, essayer sans espaces
    if (errorIndex === -1) {
      const noSpacesIndex = allTextNoSpaces.toLowerCase().indexOf(errorNoSpaces.toLowerCase());
      if (noSpacesIndex !== -1) {
        errorIndex = convertNoSpacesIndexToOriginal(allText, noSpacesIndex);
        effectiveErrorLength = findErrorLengthInOriginal(allText, errorIndex, errorNoSpaces);
        searchVariant = 'noSpaces';
        console.log(`[CORRECT PPTX] Found with noSpaces variant at index ${errorIndex}, length ${effectiveErrorLength}`);
      }
    }
    
    console.log(`[CORRECT PPTX] Search: errorIndex=${errorIndex}, variant=${searchVariant}`);
    
    if (errorIndex !== -1) {
      const errorEnd = errorIndex + effectiveErrorLength;
      console.log(`[CORRECT PPTX] Error range in allText: ${errorIndex}-${errorEnd}`);
      
      // FIX v3.2: Identifier tous les segments impliqués et leur rôle
      let firstSegmentIndex = -1;
      let correctionApplied = false;
      
      for (let i = 0; i < textPositions.length; i++) {
        const seg = textPositions[i];
        
        // Ce segment chevauche-t-il l'erreur ?
        const overlaps = seg.allTextEnd > errorIndex && seg.allTextStart < errorEnd;
        
        if (!overlaps) continue;
        
        console.log(`[CORRECT PPTX] Segment ${i} overlaps: "${seg.text}" (range: ${seg.allTextStart}-${seg.allTextEnd})`);
        
        if (firstSegmentIndex === -1) {
          // Premier segment - appliquer la correction ici
          firstSegmentIndex = i;
          
          const errorStartInSegment = Math.max(0, errorIndex - seg.allTextStart);
          const beforeError = seg.text.substring(0, errorStartInSegment);
          
          // FIX v3.2: Vérifier si l'erreur se termine dans CE segment
          const errorEndsInThisSegment = errorEnd <= seg.allTextEnd;
          
          let newText;
          if (errorEndsInThisSegment) {
            // L'erreur commence ET finit dans ce segment
            const errorEndInSegment = errorEnd - seg.allTextStart;
            const afterError = seg.text.substring(errorEndInSegment);
            newText = beforeError + correctionText + afterError;
            console.log(`[CORRECT PPTX] Error contained in single segment: before="${beforeError}", after="${afterError}"`);
          } else {
            // L'erreur commence ici mais continue dans d'autres segments
            newText = beforeError + correctionText;
            console.log(`[CORRECT PPTX] Error spans multiple segments: before="${beforeError}", correction applied here`);
          }
          
          console.log(`[CORRECT PPTX] First segment: "${seg.text}" → "${newText}"`);
          
          const newTag = `<a:t>${encode(newText)}</a:t>`;
          modifiedXml = modifiedXml.substring(0, seg.xmlStart) + newTag + modifiedXml.substring(seg.xmlEnd);
          
          // Ajuster les positions XML pour les segments suivants
          const lengthDiff = newTag.length - seg.fullMatch.length;
          for (let j = i + 1; j < textPositions.length; j++) {
            textPositions[j].xmlStart += lengthDiff;
            textPositions[j].xmlEnd += lengthDiff;
          }
          
          stats.changedTextNodes++;
          pushExample(stats.examples, errorText, correctionText);
          correctionApplied = true;
          console.log(`[CORRECT PPTX] ✅ Method 2 - Applied fragmented: "${errorText}" → "${correctionText}"`);
          
        } else {
          // Segments suivants qui faisaient partie de l'erreur
          
          // FIX v3.2: Vérifier si ce segment contient la FIN de l'erreur
          const errorEndsInThisSegment = errorEnd > seg.allTextStart && errorEnd < seg.allTextEnd;
          const segmentEntirelyInError = seg.allTextEnd <= errorEnd;
          
          if (errorEndsInThisSegment) {
            // Ce segment contient la fin de l'erreur - PRÉSERVER le texte après
            const errorEndInSegment = errorEnd - seg.allTextStart;
            const afterError = seg.text.substring(errorEndInSegment);
            
            console.log(`[CORRECT PPTX] Last error segment: preserving afterError="${afterError}"`);
            
            const newTag = `<a:t>${encode(afterError)}</a:t>`;
            modifiedXml = modifiedXml.substring(0, seg.xmlStart) + newTag + modifiedXml.substring(seg.xmlEnd);
            
            const lengthDiff = newTag.length - seg.fullMatch.length;
            for (let j = i + 1; j < textPositions.length; j++) {
              textPositions[j].xmlStart += lengthDiff;
              textPositions[j].xmlEnd += lengthDiff;
            }
            
          } else if (segmentEntirelyInError) {
            // Ce segment est ENTIÈREMENT dans l'erreur - le vider
            console.log(`[CORRECT PPTX] Clearing segment entirely in error: "${seg.text}"`);
            
            const newTag = `<a:t></a:t>`;
            modifiedXml = modifiedXml.substring(0, seg.xmlStart) + newTag + modifiedXml.substring(seg.xmlEnd);
            
            const lengthDiff = newTag.length - seg.fullMatch.length;
            for (let j = i + 1; j < textPositions.length; j++) {
              textPositions[j].xmlStart += lengthDiff;
              textPositions[j].xmlEnd += lengthDiff;
            }
          }
          // Si seg.allTextStart >= errorEnd, ce segment est APRÈS l'erreur - ne rien faire
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
 * Applique les corrections de spellingErrors au XML DOCX
 * VERSION 3.2 - Même fix que PPTX
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
        stats.changedTextNodes++;
        pushExample(stats.examples, errorText, correctionText);
        console.log(`[CORRECT DOCX] ✅ Applied: "${errorText}" → "${correctionText}" (${beforeCount}x)`);
      }
      continue;
    }
    
    // Méthode 2 : Recherche dans le texte extrait (FIX v3.2)
    const textContentRegex = /<w:t[^>]*>([^<]*)<\/w:t>/g;
    let allText = '';
    let match;
    const textPositions = [];
    
    textContentRegex.lastIndex = 0;
    
    while ((match = textContentRegex.exec(modifiedXml)) !== null) {
      const decodedText = decode(match[1]);
      const startInAllText = allText.length;
      allText += decodedText;
      
      textPositions.push({
        xmlStart: match.index,
        xmlEnd: match.index + match[0].length,
        text: decodedText,
        fullMatch: match[0],
        allTextStart: startInAllText,
        allTextEnd: allText.length
      });
    }
    
    const errorNormalized = normalizeText(errorText);
    const errorNoSpaces = removeAllSpaces(errorText);
    const allTextNoSpaces = removeAllSpaces(allText);
    
    let errorIndex = allText.toLowerCase().indexOf(errorNormalized.toLowerCase());
    let effectiveErrorLength = errorNormalized.length;
    
    if (errorIndex === -1) {
      const noSpacesIndex = allTextNoSpaces.toLowerCase().indexOf(errorNoSpaces.toLowerCase());
      if (noSpacesIndex !== -1) {
        errorIndex = convertNoSpacesIndexToOriginal(allText, noSpacesIndex);
        effectiveErrorLength = findErrorLengthInOriginal(allText, errorIndex, errorNoSpaces);
      }
    }
    
    if (errorIndex !== -1) {
      const errorEnd = errorIndex + effectiveErrorLength;
      let firstSegmentIndex = -1;
      let correctionApplied = false;
      
      for (let i = 0; i < textPositions.length; i++) {
        const seg = textPositions[i];
        const overlaps = seg.allTextEnd > errorIndex && seg.allTextStart < errorEnd;
        
        if (!overlaps) continue;
        
        if (firstSegmentIndex === -1) {
          firstSegmentIndex = i;
          
          const errorStartInSegment = Math.max(0, errorIndex - seg.allTextStart);
          const beforeError = seg.text.substring(0, errorStartInSegment);
          
          const errorEndsInThisSegment = errorEnd <= seg.allTextEnd;
          
          let newText;
          if (errorEndsInThisSegment) {
            const errorEndInSegment = errorEnd - seg.allTextStart;
            const afterError = seg.text.substring(errorEndInSegment);
            newText = beforeError + correctionText + afterError;
          } else {
            newText = beforeError + correctionText;
          }
          
          const attrMatch = seg.fullMatch.match(/<w:t([^>]*)>/);
          const attrs = attrMatch ? attrMatch[1] : '';
          
          const newTag = `<w:t${attrs}>${encode(newText)}</w:t>`;
          modifiedXml = modifiedXml.substring(0, seg.xmlStart) + newTag + modifiedXml.substring(seg.xmlEnd);
          
          const lengthDiff = newTag.length - seg.fullMatch.length;
          for (let j = i + 1; j < textPositions.length; j++) {
            textPositions[j].xmlStart += lengthDiff;
            textPositions[j].xmlEnd += lengthDiff;
          }
          
          stats.changedTextNodes++;
          pushExample(stats.examples, errorText, correctionText);
          correctionApplied = true;
          console.log(`[CORRECT DOCX] ✅ Applied fragmented: "${errorText}" → "${correctionText}"`);
          
        } else {
          // FIX v3.2: Gérer les segments suivants correctement
          const errorEndsInThisSegment = errorEnd > seg.allTextStart && errorEnd < seg.allTextEnd;
          const segmentEntirelyInError = seg.allTextEnd <= errorEnd;
          
          const attrMatch = seg.fullMatch.match(/<w:t([^>]*)>/);
          const attrs = attrMatch ? attrMatch[1] : '';
          
          if (errorEndsInThisSegment) {
            const errorEndInSegment = errorEnd - seg.allTextStart;
            const afterError = seg.text.substring(errorEndInSegment);
            
            const newTag = `<w:t${attrs}>${encode(afterError)}</w:t>`;
            modifiedXml = modifiedXml.substring(0, seg.xmlStart) + newTag + modifiedXml.substring(seg.xmlEnd);
            
            const lengthDiff = newTag.length - seg.fullMatch.length;
            for (let j = i + 1; j < textPositions.length; j++) {
              textPositions[j].xmlStart += lengthDiff;
              textPositions[j].xmlEnd += lengthDiff;
            }
            
          } else if (segmentEntirelyInError) {
            const newTag = `<w:t${attrs}></w:t>`;
            modifiedXml = modifiedXml.substring(0, seg.xmlStart) + newTag + modifiedXml.substring(seg.xmlEnd);
            
            const lengthDiff = newTag.length - seg.fullMatch.length;
            for (let j = i + 1; j < textPositions.length; j++) {
              textPositions[j].xmlStart += lengthDiff;
              textPositions[j].xmlEnd += lengthDiff;
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

    if (spellingErrors.length > 0) {
      console.log(`[DOCX] Applying spelling corrections to ${p}`);
      xml = applySpellingCorrectionsToDocxXML(xml, spellingErrors, stats);
    }

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
  
  if (spellingErrors.length > 0) {
    console.log('[PPTX] Spelling errors to apply:');
    spellingErrors.forEach((err, i) => {
      console.log(`  ${i + 1}. "${err.error}" → "${err.correction}" (type: ${err.type})`);
    });
  }

  for (const sp of slides) {
    let xml = await zip.file(sp).async('string');

    if (spellingErrors.length > 0) {
      console.log(`[PPTX] Applying spelling corrections to ${sp}`);
      xml = applySpellingCorrectionsToXML(xml, spellingErrors, stats);
    }

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
