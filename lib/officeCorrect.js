// lib/officeCorrect.js
// VERSION 5.5 — ROBUST APPLY + SANITIZE ARTIFACTS (DOCX + PPTX + XLSX)
// ✅ DOCX/XLSX: Node-only safe corrections (no XML corruption)
// ✅ PPTX: Cross-node corrections WITHOUT naive redistribution (fixes "not applied" cases)
// ✅ Removes HTML-like marker artifacts from final text (e.g. class="correction-marker"...)
// ✅ Failsafe: returns original XML if validation fails
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

// ---------------- Artifact Sanitizer ----------------
// Removes any accidental HTML-ish fragments that should never exist in Office text
// Example seen in your preview: class="correction-marker" data-index="13">
function stripMarkerArtifacts(text) {
  let t = String(text || "");
  if (!t) return t;

  // remove whole attribute blocks if present
  t = t.replace(/\s*class="correction-marker"\s*/gi, " ");
  t = t.replace(/\s*data-index="[^"]*"\s*/gi, " ");

  // remove leftover marker fragments like: class=... data-index=... >
  t = t.replace(/\bclass\s*=\s*["']correction-marker["']\b/gi, "");
  t = t.replace(/\bdata-index\s*=\s*["'][^"']*["']\b/gi, "");

  // remove stray tag-like endings introduced by bugs
  t = t.replace(/\s*>+\s*/g, " ");

  // cleanup spacing
  t = t.replace(/\s{2,}/g, " ").trim();
  return t;
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

  // sanitize artifacts BEFORE applying corrections
  for (const n of nodes) {
    const sanitized = stripMarkerArtifacts(n.text);
    if (sanitized !== n.text) {
      n.text = sanitized;
      stats.sanitizedNodes = (stats.sanitizedNodes || 0) + 1;
    }
  }

  const safeErrors = (spellingErrors || []).filter(isCorrectionSafe);
  if (!safeErrors.length) {
    console.log(`[${label}] No safe corrections to apply`);
    // still rebuild if sanitize happened
    const anySan = (stats.sanitizedNodes || 0) > 0;
    if (!anySan) return xml;
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

  console.log(`[${label}] Nodes=${nodes.length}, safeCorrections=${safeErrors.length}, planned=${planned}, sanitized=${stats.sanitizedNodes || 0}`);

  for (const [nodeIndex, plans] of plansByNode.entries()) {
    const node = nodes[nodeIndex];
    if (!node || !node.text) continue;

    plans.sort((a, b) => b.pos - a.pos);

    let text = node.text;
    let appliedInNode = 0;

    for (const p of plans) {
      const inside = text.slice(p.pos, p.end);
      if (inside.toLowerCase() !== String(p.err.error).toLowerCase()) continue;

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

// ---------------- PPTX Cross-Node corrections (robust) ----------------
// Key fix vs v5.4: do NOT "redistribute" whole text across nodes (breaks runs).
// Instead:
// 1) Build concatenated text with mapping of (nodeIndex, localCharIndex) per global char.
// 2) Find each correction occurrence (first match only).
// 3) Apply each correction by updating ONLY the affected nodes (split across them if needed).
function applyPptxCrossNodeCorrections(xml, spellingErrors, stats, label) {
  const originalXml = xml;
  const nodeRegex = /(<a:t>)([\s\S]*?)(<\/a:t>)/g;

  const nodes = [];
  let m;
  nodeRegex.lastIndex = 0;
  while ((m = nodeRegex.exec(xml)) !== null) {
    const decoded = decode(m[2] ?? "");
    const sanitized = stripMarkerArtifacts(decoded);
    nodes.push({
      index: nodes.length,
      start: m.index,
      end: m.index + m[0].length,
      open: m[1],
      close: m[3],
      text: sanitized,
    });
    if (sanitized !== decoded) {
      stats.sanitizedNodes = (stats.sanitizedNodes || 0) + 1;
    }
  }

  if (!nodes.length) {
    console.log(`[${label}] No <a:t> nodes found`);
    return xml;
  }

  // Build allText + mapping: globalCharPos -> { nodeIndex, localPos }
  let allText = "";
  const map = []; // map[g] = { nodeIndex, localPos }
  for (const node of nodes) {
    for (let i = 0; i < node.text.length; i++) {
      map.push({ nodeIndex: node.index, localPos: i });
    }
    allText += node.text;
  }

  const safeErrors = (spellingErrors || []).filter(isCorrectionSafe);
  if (!safeErrors.length && !(stats.sanitizedNodes > 0)) {
    console.log(`[${label}] No safe corrections to apply`);
    return xml;
  }

  const allLower = allText.toLowerCase();
  const corrections = [];

  for (const err of safeErrors) {
    const needle = err.error.toLowerCase();
    const idx = allLower.indexOf(needle);
    if (idx === -1) continue;

    corrections.push({
      err,
      start: idx,
      end: idx + err.error.length,
    });

    console.log(`[${label}] Found "${err.error}" at ${idx} → "${err.correction}"`);
  }

  if (!corrections.length) {
    console.log(`[${label}] No corrections found in concatenated text (sanitized=${stats.sanitizedNodes || 0})`);
    // still rebuild if sanitized occurred
  }

  // Apply from end to start to avoid index shifts in mapping calculations
  corrections.sort((a, b) => b.start - a.start);

  // We will mutate node.text by slicing around impacted ranges
  for (const corr of corrections) {
    const { err, start, end } = corr;

    if (start < 0 || end > map.length || end <= start) continue;

    const startRef = map[start];
    const endRef = map[end - 1];
    if (!startRef || !endRef) continue;

    const startNodeIdx = startRef.nodeIndex;
    const endNodeIdx = endRef.nodeIndex;

    // If entirely within one node, easy
    if (startNodeIdx === endNodeIdx) {
      const node = nodes[startNodeIdx];
      const localStart = startRef.localPos;
      const localEnd = endRef.localPos + 1;

      const before = node.text.slice(0, localStart);
      const mid = node.text.slice(localStart, localEnd);
      const after = node.text.slice(localEnd);

      if (mid.toLowerCase() !== err.error.toLowerCase()) continue;

      node.text = before + err.correction + after;

      stats.changedTextNodes = (stats.changedTextNodes || 0) + 1;
      pushExample(stats, err.error, err.correction);
      continue;
    }

    // Spans multiple nodes:
    // Replace the cross-node segment by:
    // - truncate from start node localStart to end-of-node
    // - empty middle nodes fully
    // - replace in end node from 0..localEnd
    // Insert correction into start node (preferred) and adjust end node removal
    const startNode = nodes[startNodeIdx];
    const endNode = nodes[endNodeIdx];

    const startLocal = startRef.localPos;
    const endLocalEnd = endRef.localPos + 1;

    const startBefore = startNode.text.slice(0, startLocal);
    const endAfter = endNode.text.slice(endLocalEnd);

    // Validate what we’re replacing (best effort)
    const replaced = allText.slice(start, end);
    if (replaced.toLowerCase() !== err.error.toLowerCase()) {
      // mapping may be stale because earlier corrections changed node.text lengths
      // failsafe: skip to avoid corruption
      console.warn(`[${label}] Skip cross-node correction (mapping stale): "${err.error}"`);
      continue;
    }

    // Apply:
    // 1) start node becomes: before + correction
    startNode.text = startBefore + err.correction;

    // 2) middle nodes cleared
    for (let ni = startNodeIdx + 1; ni < endNodeIdx; ni++) {
      nodes[ni].text = "";
    }

    // 3) end node becomes: after remaining suffix
    endNode.text = endAfter;

    stats.changedTextNodes = (stats.changedTextNodes || 0) + 1;
    pushExample(stats, err.error, err.correction);

    // Important: we do NOT rebuild map/allText mid-way to keep it simple and safe.
    // Cross-node spans are rare; skipping if mapping stale avoids corruption.
  }

  // Rebuild XML with modified nodes
  let out = "";
  let xmlCursor = 0;
  for (const node of nodes) {
    out += xml.slice(xmlCursor, node.start);
    out += node.open + encode(node.text) + node.close;
    xmlCursor = node.end;
  }
  out += xml.slice(xmlCursor);

  if (!validatePptxXml(out)) {
    console.error(`[${label}] FAILSAFE: validation failed, returning original XML`);
    stats.failsafeTriggered = true;
    return originalXml;
  }

  console.log(
    `[${label}] Applied corrections. changedTextNodes=${stats.changedTextNodes || 0}, sanitized=${stats.sanitizedNodes || 0}`
  );
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

  const stats = { changedTextNodes: 0, examples: [], failsafeTriggered: false, sanitizedNodes: 0 };
  const spellingErrors = options.spellingErrors || [];

  const existingTargets = targets.filter((t) => zip.file(t));
  console.log(`[DOCX] Processing ${existingTargets.length} files, corrections=${spellingErrors.length}`);

  for (const p of existingTargets) {
    const file = zip.file(p);
    if (!file) continue;

    let xml = await file.async("string");
    const originalLength = xml.length;

    xml = applyCorrectionsNodeOnly(
      xml,
      /(<w:t\b[^>]*>)([\s\S]*?)(<\/w:t>)/g,
      spellingErrors,
      stats,
      validateDocxXml,
      `DOCX:${p}`
    );

    console.log(`[DOCX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  if (stats.failsafeTriggered) console.warn(`[DOCX] ⚠️ Failsafe triggered`);
  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

// ---------------- PPTX ----------------
export async function correctPPTXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);

  const targets = Object.keys(zip.files).filter(
    (k) => /^ppt\/slides\/slide\d+\.xml$/.test(k) || /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(k)
  );

  const stats = { changedTextNodes: 0, examples: [], failsafeTriggered: false, sanitizedNodes: 0 };
  const spellingErrors = options.spellingErrors || [];

  console.log(`[PPTX] Processing ${targets.length} files, corrections=${spellingErrors.length}`);

  for (const p of targets) {
    const file = zip.file(p);
    if (!file) continue;

    let xml = await file.async("string");
    const originalLength = xml.length;

    xml = applyPptxCrossNodeCorrections(xml, spellingErrors, stats, `PPTX:${p}`);

    console.log(`[PPTX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  if (stats.failsafeTriggered) console.warn(`[PPTX] ⚠️ Failsafe triggered`);
  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

// ---------------- XLSX ----------------
export async function correctXLSXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);

  const targets = [
    "xl/sharedStrings.xml",
    ...Object.keys(zip.files).filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k)),
  ];

  const stats = { changedTextNodes: 0, examples: [], failsafeTriggered: false, sanitizedNodes: 0 };
  const spellingErrors = options.spellingErrors || [];

  const existingTargets = targets.filter((t) => zip.file(t));
  console.log(`[XLSX] Processing ${existingTargets.length} files, corrections=${spellingErrors.length}`);

  for (const p of existingTargets) {
    const file = zip.file(p);
    if (!file) continue;

    let xml = await file.async("string");
    const originalLength = xml.length;

    xml = applyCorrectionsNodeOnly(
      xml,
      /(<t\b[^>]*>)([\s\S]*?)(<\/t>)/g,
      spellingErrors,
      stats,
      validateXlsxXml,
      `XLSX:${p}`
    );

    console.log(`[XLSX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  if (stats.failsafeTriggered) console.warn(`[XLSX] ⚠️ Failsafe triggered`);
  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

export default { correctDOCXText, correctPPTXText, correctXLSXText };
