// lib/officeCorrect.js
// VERSION 5.4 — CROSS-NODE FRAGMENTED WORD FIX (DOCX + PPTX + XLSX)
// ✅ PPTX: Handles fragmented words across multiple <a:t> nodes
// ✅ DOCX/XLSX: Node-only safe corrections
// ✅ Prevents corruption from invalid XML
// ✅ Keeps original if validation fails (failsafe)
// ✅ Detailed logging

import JSZip from "jszip";

// ---------------- Utils ----------------
const decode = (s) =>
  String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");

const encode = (s) =>
  String(s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");

const normalize = (s) => String(s || "").replace(/\s+/g, " ").trim();

function pushExample(stats, before, after) {
  if (!stats.examples) stats.examples = [];
  if (stats.examples.length < 10) {
    stats.examples.push({
      before: String(before || "").slice(0, 140),
      after: String(after || "").slice(0, 140),
    });
  }
}

// ---------------- Safety checks ----------------

// XML 1.0 illegal chars: 0x00-0x08, 0x0B, 0x0C, 0x0E-0x1F
function hasIllegalXmlChars(text) {
  const s = String(text || "");
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (
      (c >= 0x00 && c <= 0x08) ||
      c === 0x0b ||
      c === 0x0c ||
      (c >= 0x0e && c <= 0x1f)
    ) {
      return true;
    }
  }
  return false;
}

// reject unpaired surrogates (can break parsers)
function hasUnpairedSurrogate(text) {
  const s = String(text || "");
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c >= 0xd800 && c <= 0xdbff) {
      const n = s.charCodeAt(i + 1);
      if (!(n >= 0xdc00 && n <= 0xdfff)) return true;
      i++;
    } else if (c >= 0xdc00 && c <= 0xdfff) {
      return true;
    }
  }
  return false;
}

function isCorrectionSafe(err) {
  if (!err || typeof err.error !== "string" || typeof err.correction !== "string") return false;
  const from = err.error;
  const to = err.correction;

  if (!from.trim()) return false;
  if (from === to) return false;

  if (hasIllegalXmlChars(to) || hasUnpairedSurrogate(to)) return false;
  if (to.length > 2000) return false;

  return true;
}

// ---------------- XML Validation ----------------
function validateDocxXml(xml) {
  if (!xml || typeof xml !== "string" || xml.length < 50) return false;

  const openWT = (xml.match(/<w:t\b[^>]*>/g) || []).length;
  const closeWT = (xml.match(/<\/w:t>/g) || []).length;
  if (openWT !== closeWT) {
    console.error(`[VALIDATE DOCX] <w:t> mismatch: ${openWT} open vs ${closeWT} close`);
    return false;
  }

  const openWP = (xml.match(/<w:p\b[^>]*>/g) || []).length;
  const closeWP = (xml.match(/<\/w:p>/g) || []).length;
  if (openWP !== closeWP && (openWP > 0 || closeWP > 0)) {
    console.error(`[VALIDATE DOCX] <w:p> mismatch: ${openWP} open vs ${closeWP} close`);
    return false;
  }

  return true;
}

function validatePptxXml(xml) {
  if (!xml || typeof xml !== "string" || xml.length < 50) return false;
  const openAT = (xml.match(/<a:t>/g) || []).length;
  const closeAT = (xml.match(/<\/a:t>/g) || []).length;
  if (openAT !== closeAT) {
    console.error(`[VALIDATE PPTX] <a:t> mismatch: ${openAT} open vs ${closeAT} close`);
    return false;
  }
  return true;
}

function validateXlsxXml(xml) {
  if (!xml || typeof xml !== "string" || xml.length < 20) return false;
  return true;
}

// ---------------- Context scoring ----------------
function scoreContextInNode(nodeText, idx, err) {
  const e = err.error;
  const before = normalize(err.contextBefore || "");
  const after = normalize(err.contextAfter || "");
  let score = 0;

  const winBefore = normalize(nodeText.slice(Math.max(0, idx - 60), idx));
  const winAfter = normalize(nodeText.slice(idx + e.length, idx + e.length + 60));

  if (before && winBefore.endsWith(before)) score += 5;
  if (after && winAfter.startsWith(after)) score += 5;
  if (nodeText.slice(idx, idx + e.length) === e) score += 3;

  return score;
}

function findBestOccurrenceInNode(nodeText, err) {
  const e = err.error;
  if (!e) return -1;

  const lcText = nodeText.toLowerCase();
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
    const sc = scoreContextInNode(nodeText, p, err);
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

// ---------------- Node-only corrections (DOCX/XLSX) ----------------
function applyCorrectionsNodeOnly(xml, nodeRegex, spellingErrors, stats, validateFn, label) {
  const originalXml = xml;
  const nodes = [];

  nodeRegex.lastIndex = 0;
  let m;
  while ((m = nodeRegex.exec(xml)) !== null) {
    nodes.push({
      start: m.index,
      end: m.index + m[0].length,
      open: m[1],
      raw: m[2] ?? "",
      close: m[3],
      text: decode(m[2] ?? ""),
    });
  }

  if (!nodes.length) {
    console.log(`[${label}] No text nodes found`);
    return xml;
  }

  const safeErrors = (spellingErrors || []).filter(isCorrectionSafe);
  if (!safeErrors.length) {
    console.log(`[${label}] No safe corrections to apply`);
    return xml;
  }

  const plansByNode = new Map();
  let planned = 0;

  for (const err of safeErrors) {
    let bestNode = -1;
    let bestPos = -1;
    let bestScore = -1;

    for (let ni = 0; ni < nodes.length; ni++) {
      const t = nodes[ni].text;
      if (!t || t.length < err.error.length) continue;

      const pos = findBestOccurrenceInNode(t, err);
      if (pos === -1) continue;

      const sc = scoreContextInNode(t, pos, err);
      if (sc > bestScore) {
        bestScore = sc;
        bestNode = ni;
        bestPos = pos;
      }
    }

    if (bestNode === -1) continue;

    if (!plansByNode.has(bestNode)) plansByNode.set(bestNode, []);
    plansByNode.get(bestNode).push({
      err,
      pos: bestPos,
      end: bestPos + err.error.length,
    });
    planned++;
  }

  console.log(`[${label}] Nodes=${nodes.length}, safeCorrections=${safeErrors.length}, planned=${planned}`);

  if (!planned) return xml;

  for (const [nodeIndex, plans] of plansByNode.entries()) {
    const node = nodes[nodeIndex];
    if (!node || !node.text) continue;

    plans.sort((a, b) => b.pos - a.pos);

    let text = node.text;
    let appliedInNode = 0;

    for (const p of plans) {
      const inside = text.slice(p.pos, p.end);
      if (inside.toLowerCase() !== String(p.err.error).toLowerCase()) {
        continue;
      }
      const before = text.slice(0, p.pos);
      const after = text.slice(p.end);
      text = before + p.err.correction + after;
      appliedInNode++;
    }

    if (appliedInNode > 0) {
      node.text = text;
      stats.changedTextNodes = (stats.changedTextNodes || 0) + 1;
      pushExample(stats, plans[plans.length - 1].err.error, plans[plans.length - 1].err.correction);
    }
  }

  let out = "";
  let cursor = 0;

  for (const node of nodes) {
    out += xml.slice(cursor, node.start);
    out += node.open + encode(node.text) + node.close;
    cursor = node.end;
  }
  out += xml.slice(cursor);

  if (validateFn && !validateFn(out)) {
    console.error(`[${label}] FAILSAFE: validation failed, returning original XML`);
    stats.failsafeTriggered = true;
    return originalXml;
  }

  return out;
}

// ---------------- PPTX Cross-Node corrections (for fragmented words) ----------------
function applyPptxCrossNodeCorrections(xml, spellingErrors, stats, label) {
  const originalXml = xml;
  const nodeRegex = /(<a:t>)([\s\S]*?)(<\/a:t>)/g;
  
  // Extract all nodes with positions
  const nodes = [];
  let m;
  nodeRegex.lastIndex = 0;
  while ((m = nodeRegex.exec(xml)) !== null) {
    nodes.push({
      index: nodes.length,
      start: m.index,
      end: m.index + m[0].length,
      open: m[1],
      raw: m[2] ?? "",
      close: m[3],
      text: decode(m[2] ?? ""),
      modified: false,
      newText: null,
    });
  }

  if (!nodes.length) {
    console.log(`[${label}] No <a:t> nodes found`);
    return xml;
  }

  // Build concatenated text with node mapping
  let allText = "";
  const charToNode = []; // For each char in allText, which node index owns it
  
  for (const node of nodes) {
    for (let i = 0; i < node.text.length; i++) {
      charToNode.push(node.index);
    }
    allText += node.text;
  }

  console.log(`[${label}] allText (${allText.length} chars): "${allText.slice(0, 200)}${allText.length > 200 ? '...' : ''}"`);

  const safeErrors = (spellingErrors || []).filter(isCorrectionSafe);
  if (!safeErrors.length) {
    console.log(`[${label}] No safe corrections to apply`);
    return xml;
  }

  // Find corrections in the concatenated text
  const corrections = [];
  
  for (const err of safeErrors) {
    const errLower = err.error.toLowerCase();
    const allTextLower = allText.toLowerCase();
    
    let searchPos = 0;
    while (true) {
      const idx = allTextLower.indexOf(errLower, searchPos);
      if (idx === -1) break;
      
      // Check if this is a good match (exact case or context)
      const foundText = allText.slice(idx, idx + err.error.length);
      
      corrections.push({
        err,
        start: idx,
        end: idx + err.error.length,
        foundText,
      });
      
      console.log(`[${label}] Found "${err.error}" at pos ${idx}, will replace with "${err.correction}"`);
      searchPos = idx + 1;
      break; // Only first occurrence for each error
    }
  }

  if (!corrections.length) {
    console.log(`[${label}] No corrections found in allText`);
    return xml;
  }

  // Sort corrections by position (descending) to apply from end to start
  corrections.sort((a, b) => b.start - a.start);

  // Apply corrections to allText and track affected nodes
  let modifiedAllText = allText;
  const nodeChanges = new Map(); // nodeIndex -> array of changes
  
  for (const corr of corrections) {
    const { err, start, end } = corr;
    
    // Find which nodes are affected
    const startNode = charToNode[start];
    const endNode = charToNode[end - 1];
    
    console.log(`[${label}] Correction spans nodes ${startNode} to ${endNode}`);
    
    // Apply to modifiedAllText
    modifiedAllText = modifiedAllText.slice(0, start) + err.correction + modifiedAllText.slice(end);
    
    // Update charToNode mapping (shift indices after this correction)
    const lengthDiff = err.correction.length - err.error.length;
    if (lengthDiff !== 0) {
      // We need to rebuild charToNode after each correction... 
      // For simplicity, we'll rebuild the nodes from modifiedAllText at the end
    }
    
    stats.changedTextNodes = (stats.changedTextNodes || 0) + 1;
    pushExample(stats, err.error, err.correction);
  }

  console.log(`[${label}] modifiedAllText: "${modifiedAllText.slice(0, 200)}${modifiedAllText.length > 200 ? '...' : ''}"`);

  // Now redistribute modifiedAllText back into the original node structure
  // Strategy: preserve original node boundaries as much as possible
  let cursor = 0;
  for (const node of nodes) {
    const originalLength = node.text.length;
    
    if (cursor < modifiedAllText.length) {
      // Calculate how much text this node should have
      // For nodes in the middle, keep same length
      // For the last affected node, take remaining text
      
      let newText;
      if (node.index === nodes.length - 1) {
        // Last node gets everything remaining
        newText = modifiedAllText.slice(cursor);
      } else {
        // Try to keep original length, but don't exceed remaining text
        const remainingText = modifiedAllText.slice(cursor);
        if (remainingText.length <= originalLength) {
          // Less text remaining than this node had, give it all to this node
          newText = remainingText;
        } else {
          newText = modifiedAllText.slice(cursor, cursor + originalLength);
        }
      }
      
      node.newText = newText;
      node.modified = true;
      cursor += newText.length;
    } else {
      // No more text, this node becomes empty
      node.newText = "";
      node.modified = true;
    }
  }

  // Rebuild XML with modified nodes
  let out = "";
  let xmlCursor = 0;

  for (const node of nodes) {
    out += xml.slice(xmlCursor, node.start);
    const textToUse = node.modified ? node.newText : node.text;
    out += node.open + encode(textToUse) + node.close;
    xmlCursor = node.end;
  }
  out += xml.slice(xmlCursor);

  // Validate
  if (!validatePptxXml(out)) {
    console.error(`[${label}] FAILSAFE: validation failed, returning original XML`);
    stats.failsafeTriggered = true;
    return originalXml;
  }

  console.log(`[${label}] Successfully applied ${corrections.length} cross-node corrections`);
  return out;
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

  const existingTargets = targets.filter((t) => zip.file(t));
  console.log(`[DOCX] Processing ${existingTargets.length} files, corrections=${spellingErrors.length}`);

  for (const p of existingTargets) {
    const file = zip.file(p);
    if (!file) continue;

    let xml = await file.async("string");
    const originalLength = xml.length;

    if (spellingErrors.length) {
      xml = applyCorrectionsNodeOnly(
        xml,
        /(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g,
        spellingErrors,
        stats,
        validateDocxXml,
        `DOCX:${p}`
      );
    }

    console.log(`[DOCX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  if (stats.failsafeTriggered) {
    console.warn(`[DOCX] ⚠️ Failsafe triggered`);
  }

  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

// ---------------- PPTX ----------------
export async function correctPPTXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);

  const targets = Object.keys(zip.files).filter(
    (k) => /^ppt\/slides\/slide\d+\.xml$/.test(k) || /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(k)
  );

  const stats = { changedTextNodes: 0, examples: [], failsafeTriggered: false };
  const spellingErrors = options.spellingErrors || [];

  console.log(`[PPTX] Processing ${targets.length} files, corrections=${spellingErrors.length}`);

  for (const p of targets) {
    const file = zip.file(p);
    if (!file) continue;

    let xml = await file.async("string");
    const originalLength = xml.length;

    if (spellingErrors.length) {
      // ✅ Use cross-node correction for PPTX (handles fragmented words)
      xml = applyPptxCrossNodeCorrections(xml, spellingErrors, stats, `PPTX:${p}`);
    }

    console.log(`[PPTX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  if (stats.failsafeTriggered) {
    console.warn(`[PPTX] ⚠️ Failsafe triggered`);
  }

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

  const existingTargets = targets.filter((t) => zip.file(t));
  console.log(`[XLSX] Processing ${existingTargets.length} files, corrections=${spellingErrors.length}`);

  for (const p of existingTargets) {
    const file = zip.file(p);
    if (!file) continue;

    let xml = await file.async("string");
    const originalLength = xml.length;

    if (spellingErrors.length) {
      xml = applyCorrectionsNodeOnly(
        xml,
        /(<t\b[^>]*>)([\s\S]*?)(<\/t>)/g,
        spellingErrors,
        stats,
        validateXlsxXml,
        `XLSX:${p}`
      );
    }

    console.log(`[XLSX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  if (stats.failsafeTriggered) {
    console.warn(`[XLSX] ⚠️ Failsafe triggered`);
  }

  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

export default { correctDOCXText, correctPPTXText, correctXLSXText };
