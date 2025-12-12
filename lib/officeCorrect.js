// lib/officeCorrect.js
// VERSION 5.2 — REVERSE-ORDER SAFE (DOCX + PPTX + XLSX)
// ✔ Corrections applied in REVERSE order (prevents index shifting)
// ✔ XML validation before write
// ✔ Failsafe mode (keeps original if corruption detected)
// ✔ Enhanced logging for debugging
// ✔ Protection against empty/invalid corrections

import JSZip from "jszip";

// ---------------- Utils ----------------
const decode = (s) =>
  (s || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");

const encode = (s) =>
  (s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");

const normalize = (s) => (s || "").replace(/\s+/g, " ").trim();

function pushExample(stats, before, after) {
  if (!stats.examples) stats.examples = [];
  if (stats.examples.length < 10) {
    stats.examples.push({
      before: String(before || "").slice(0, 140),
      after: String(after || "").slice(0, 140),
    });
  }
}

// ---------------- XML Validation ----------------
function isXmlSafe(text) {
  if (!text || typeof text !== "string") return false;
  // After encoding, should not contain raw < or >
  const encoded = encode(text);
  return !encoded.includes('<') && !encoded.includes('>');
}

function validateDocxXml(xml) {
  // Basic structure validation for DOCX
  if (!xml || typeof xml !== "string") return false;
  if (xml.length < 100) return false;
  
  // Check for essential DOCX XML structure
  const hasXmlDecl = xml.includes('<?xml');
  const hasClosingTags = xml.includes('</w:t>') || xml.includes('</w:p>');
  
  // Count opening vs closing w:t tags
  const openCount = (xml.match(/<w:t[^>]*>/g) || []).length;
  const closeCount = (xml.match(/<\/w:t>/g) || []).length;
  
  if (openCount !== closeCount) {
    console.error(`[VALIDATE] Tag mismatch: ${openCount} open vs ${closeCount} close w:t tags`);
    return false;
  }
  
  return hasXmlDecl || hasClosingTags;
}

function validatePptxXml(xml) {
  if (!xml || typeof xml !== "string") return false;
  if (xml.length < 100) return false;
  
  const openCount = (xml.match(/<a:t>/g) || []).length;
  const closeCount = (xml.match(/<\/a:t>/g) || []).length;
  
  if (openCount !== closeCount) {
    console.error(`[VALIDATE] Tag mismatch: ${openCount} open vs ${closeCount} close a:t tags`);
    return false;
  }
  
  return true;
}

function validateXlsxXml(xml) {
  if (!xml || typeof xml !== "string") return false;
  if (xml.length < 50) return false;
  return true;
}

// ---------------- Context match ----------------
function scoreContext(allText, idx, err) {
  const e = err.error;
  const before = normalize(err.contextBefore || "");
  const after = normalize(err.contextAfter || "");
  let score = 0;

  const winBefore = normalize(allText.slice(Math.max(0, idx - 60), idx));
  const winAfter = normalize(allText.slice(idx + e.length, idx + e.length + 60));

  if (before && winBefore.endsWith(before)) score += 5;
  if (after && winAfter.startsWith(after)) score += 5;
  if (allText.slice(idx, idx + e.length) === e) score += 3;

  return score;
}

function findBestOccurrence(allText, err) {
  const e = err.error;
  if (!e) return -1;

  const lcText = allText.toLowerCase();
  const lcErr = e.toLowerCase();

  const positions = [];
  let i = 0;
  while (true) {
    const idx = lcText.indexOf(lcErr, i);
    if (idx === -1) break;
    positions.push(idx);
    i = idx + Math.max(1, lcErr.length);
  }
  if (!positions.length) return -1;

  let bestIdx = -1;
  let bestScore = -1;
  for (const p of positions) {
    const sc = scoreContext(allText, p, err);
    if (sc > bestScore) {
      bestScore = sc;
      bestIdx = p;
    }
  }

  const hasContext =
    (err.contextBefore && err.contextBefore.trim()) ||
    (err.contextAfter && err.contextAfter.trim());

  if (hasContext && bestScore <= 0) return -1;
  return bestIdx;
}

// ---------------- Prepare and sort corrections ----------------
function prepareCorrections(allText, segments, spellingErrors) {
  const prepared = [];
  
  for (const err of spellingErrors || []) {
    if (!err?.error || typeof err?.correction !== "string") {
      console.warn(`[SKIP] Invalid correction object:`, err);
      continue;
    }
    
    // Skip empty corrections
    if (!err.error.trim()) {
      console.warn(`[SKIP] Empty error string`);
      continue;
    }
    
    // Safety check: skip if correction would break XML
    if (!isXmlSafe(err.correction)) {
      console.warn(`[SKIP] Unsafe correction: "${err.error}" → "${err.correction}"`);
      continue;
    }

    const idx = findBestOccurrence(allText, err);
    if (idx === -1) {
      console.log(`[SKIP] Not found in text: "${err.error.slice(0, 50)}"`);
      continue;
    }

    const end = idx + err.error.length;

    // Find overlapping segments
    const overlapped = [];
    for (let i = 0; i < segments.length; i++) {
      const s = segments[i];
      if (s.allTextEnd <= idx || s.allTextStart >= end) continue;
      overlapped.push({ i, s: { ...s } }); // Clone segment
    }
    
    if (!overlapped.length) {
      console.warn(`[SKIP] No overlapping segments for: "${err.error.slice(0, 50)}"`);
      continue;
    }

    prepared.push({
      err,
      idx,
      end,
      overlapped,
    });
  }

  // ✅ KEY FIX: Sort by position DESCENDING (last first)
  // This way, corrections at the end don't affect indices of earlier corrections
  prepared.sort((a, b) => b.idx - a.idx);
  
  console.log(`[PREPARE] ${prepared.length} corrections ready (sorted reverse order)`);
  return prepared;
}

// ---------------- Apply a single correction ----------------
function applySingleCorrection(modified, segments, allText, correction, stats) {
  const { err, idx, end, overlapped } = correction;

  // Re-find overlapping segments (indices may have shifted)
  // Since we apply in reverse order, we need fresh segment data
  const currentOverlapped = [];
  for (let i = 0; i < segments.length; i++) {
    const s = segments[i];
    if (s.allTextEnd <= idx || s.allTextStart >= end) continue;
    currentOverlapped.push({ i, s });
  }
  
  if (!currentOverlapped.length) {
    console.warn(`[SKIP] Segments no longer overlap: "${err.error.slice(0, 30)}"`);
    return { modified, allText, applied: false };
  }

  // ---------- Single segment case ----------
  if (currentOverlapped.length === 1) {
    const { i: segIndex, s } = currentOverlapped[0];

    const startInSeg = Math.max(0, idx - s.allTextStart);
    const endInSeg = Math.min(s.text.length, end - s.allTextStart);

    const inside = s.text.slice(startInSeg, endInSeg);
    if (inside.toLowerCase() !== String(err.error).toLowerCase()) {
      console.warn(`[SKIP] Text mismatch: expected "${err.error}", found "${inside}"`);
      return { modified, allText, applied: false };
    }

    const newText = s.text.slice(0, startInSeg) + err.correction + s.text.slice(endInSeg);
    const newTag = s.full.replace(s.raw, encode(newText));

    // Apply modification
    const before = modified.slice(0, s.xmlStart);
    const after = modified.slice(s.xmlEnd);
    modified = before + newTag + after;

    const xmlDiff = newTag.length - s.full.length;
    const textDiff = newText.length - s.text.length;

    // Update this segment
    s.full = newTag;
    s.raw = encode(newText);
    s.text = newText;
    s.xmlEnd += xmlDiff;
    s.allTextEnd += textDiff;

    // Update allText
    allText = allText.slice(0, idx) + err.correction + allText.slice(end);

    // Update subsequent segments (those with higher indices in XML)
    for (let j = segIndex + 1; j < segments.length; j++) {
      segments[j].xmlStart += xmlDiff;
      segments[j].xmlEnd += xmlDiff;
      segments[j].allTextStart += textDiff;
      segments[j].allTextEnd += textDiff;
    }

    stats.changedTextNodes = (stats.changedTextNodes || 0) + 1;
    pushExample(stats, err.error, err.correction);
    
    return { modified, allText, applied: true };
  }

  // ---------- Multi-segment case ----------
  const first = currentOverlapped[0];
  const mergedText = currentOverlapped.map((o) => o.s.text).join("");

  const mergedStart = idx - first.s.allTextStart;
  const mergedEnd = mergedStart + err.error.length;

  const insideMerged = mergedText.slice(mergedStart, mergedEnd);
  if (insideMerged.toLowerCase() !== String(err.error).toLowerCase()) {
    console.warn(`[SKIP] Multi-segment text mismatch`);
    return { modified, allText, applied: false };
  }

  const newMerged = mergedText.slice(0, mergedStart) + err.correction + mergedText.slice(mergedEnd);

  const lengths = currentOverlapped.map((o) => o.s.text.length);
  const totalLen = lengths.reduce((a, b) => a + b, 0);
  if (totalLen <= 0) {
    return { modified, allText, applied: false };
  }

  const totalTextDiff = newMerged.length - mergedText.length;

  // Distribute new text across segments
  let cursor = 0;
  const splitTexts = lengths.map((len, k) => {
    if (k === lengths.length - 1) return newMerged.slice(cursor);
    const part = newMerged.slice(cursor, cursor + len);
    cursor += len;
    return part;
  });

  // Apply changes to each overlapping segment (in reverse to preserve indices)
  let cumulativeXmlDiff = 0;
  
  for (let n = currentOverlapped.length - 1; n >= 0; n--) {
    const { i: segIndex, s } = currentOverlapped[n];
    const nextText = splitTexts[n] ?? "";
    const newTag = s.full.replace(s.raw, encode(nextText));

    const before = modified.slice(0, s.xmlStart);
    const after = modified.slice(s.xmlEnd);
    modified = before + newTag + after;

    const xmlDiff = newTag.length - s.full.length;

    s.full = newTag;
    s.raw = encode(nextText);
    s.text = nextText;
    s.xmlEnd = s.xmlStart + newTag.length;
    
    cumulativeXmlDiff += xmlDiff;
  }

  // Update segments after the last overlapped one
  const lastOverlappedIndex = currentOverlapped[currentOverlapped.length - 1].i;
  for (let j = lastOverlappedIndex + 1; j < segments.length; j++) {
    segments[j].xmlStart += cumulativeXmlDiff;
    segments[j].xmlEnd += cumulativeXmlDiff;
    segments[j].allTextStart += totalTextDiff;
    segments[j].allTextEnd += totalTextDiff;
  }

  // Update allText
  allText = allText.slice(0, idx) + err.correction + allText.slice(end);

  stats.changedTextNodes = (stats.changedTextNodes || 0) + currentOverlapped.length;
  pushExample(stats, err.error, err.correction);

  return { modified, allText, applied: true };
}

// ---------------- Apply corrections in segmented XML (V5.2) ----------------
function applyCorrectionsToSegments(xml, tagRegex, spellingErrors, stats, validateFn) {
  const originalXml = xml; // Keep original for failsafe
  let modified = xml;

  // Extract all segments
  const segments = [];
  let allText = "";
  let m;

  tagRegex.lastIndex = 0;
  while ((m = tagRegex.exec(xml)) !== null) {
    const raw = m[1] ?? "";
    const decoded = decode(raw);
    const start = allText.length;
    allText += decoded;

    segments.push({
      xmlStart: m.index,
      xmlEnd: m.index + m[0].length,
      raw,
      text: decoded,
      allTextStart: start,
      allTextEnd: allText.length,
      full: m[0],
    });
  }

  if (!segments.length) {
    console.log(`[CORRECT] No segments found`);
    return modified;
  }

  console.log(`[CORRECT] Found ${segments.length} text segments, ${allText.length} chars total`);

  // Prepare corrections sorted in reverse order
  const preparedCorrections = prepareCorrections(allText, segments, spellingErrors);
  
  if (!preparedCorrections.length) {
    console.log(`[CORRECT] No valid corrections to apply`);
    return modified;
  }

  let appliedCount = 0;
  let failedCount = 0;

  // Apply each correction (already sorted in reverse order)
  for (const correction of preparedCorrections) {
    const result = applySingleCorrection(modified, segments, allText, correction, stats);
    modified = result.modified;
    allText = result.allText;
    
    if (result.applied) {
      appliedCount++;
    } else {
      failedCount++;
    }
  }

  console.log(`[CORRECT] Applied: ${appliedCount}, Failed: ${failedCount}`);

  // ✅ FAILSAFE: Validate XML structure before returning
  if (validateFn && !validateFn(modified)) {
    console.error(`[FAILSAFE] XML validation failed! Returning original XML`);
    stats.failsafeTriggered = true;
    return originalXml;
  }

  return modified;
}

// ---------------- DOCX ----------------
export async function correctDOCXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);

  const targets = [
    "word/document.xml",
    ...Object.keys(zip.files).filter((k) => /word\/(header|footer)\d+\.xml$/.test(k)),
    "word/footnotes.xml",
    "word/endnotes.xml",
  ];

  const stats = { changedTextNodes: 0, examples: [], failsafeTriggered: false };
  const spellingErrors = options.spellingErrors || [];

  const existingTargets = targets.filter(t => zip.file(t));
  console.log(`[DOCX] Processing ${existingTargets.length} files, ${spellingErrors.length} corrections`);

  for (const p of existingTargets) {
    let xml = await zip.file(p).async("string");
    const originalLength = xml.length;

    if (spellingErrors.length) {
      xml = applyCorrectionsToSegments(
        xml, 
        /<w:t[^>]*>([\s\S]*?)<\/w:t>/g, 
        spellingErrors, 
        stats,
        validateDocxXml
      );
    }

    console.log(`[DOCX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  if (stats.failsafeTriggered) {
    console.warn(`[DOCX] ⚠️ Failsafe was triggered - some files kept original content`);
  }

  console.log(`[DOCX] Total changes: ${stats.changedTextNodes}`);
  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

// ---------------- PPTX ----------------
export async function correctPPTXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);

  const targets = Object.keys(zip.files).filter((k) =>
    (/^ppt\/slides\/slide\d+\.xml$/.test(k) || /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(k))
  );

  const stats = { changedTextNodes: 0, examples: [], failsafeTriggered: false };
  const spellingErrors = options.spellingErrors || [];

  console.log(`[PPTX] Processing ${targets.length} files, ${spellingErrors.length} corrections`);

  for (const p of targets) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async("string");
    const originalLength = xml.length;

    if (spellingErrors.length) {
      xml = applyCorrectionsToSegments(
        xml, 
        /<a:t>([\s\S]*?)<\/a:t>/g, 
        spellingErrors, 
        stats,
        validatePptxXml
      );
    }

    console.log(`[PPTX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  if (stats.failsafeTriggered) {
    console.warn(`[PPTX] ⚠️ Failsafe was triggered - some files kept original content`);
  }

  console.log(`[PPTX] Total changes: ${stats.changedTextNodes}`);
  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

// ---------------- XLSX ----------------
export async function correctXLSXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);

  const targets = [
    "xl/sharedStrings.xml",
    ...Object.keys(zip.files).filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k)),
  ];

  const stats = { changedTextNodes: 0, examples: [], failsafeTriggered: false };
  const spellingErrors = options.spellingErrors || [];

  const existingTargets = targets.filter(t => zip.file(t));
  console.log(`[XLSX] Processing ${existingTargets.length} files, ${spellingErrors.length} corrections`);

  for (const p of existingTargets) {
    let xml = await zip.file(p).async("string");
    const originalLength = xml.length;

    if (spellingErrors.length) {
      xml = applyCorrectionsToSegments(
        xml, 
        /<t[^>]*>([\s\S]*?)<\/t>/g, 
        spellingErrors, 
        stats,
        validateXlsxXml
      );
    }

    console.log(`[XLSX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  if (stats.failsafeTriggered) {
    console.warn(`[XLSX] ⚠️ Failsafe was triggered - some files kept original content`);
  }

  console.log(`[XLSX] Total changes: ${stats.changedTextNodes}`);
  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

export default { correctDOCXText, correctPPTXText, correctXLSXText };
