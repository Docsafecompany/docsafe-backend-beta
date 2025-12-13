// lib/officeCorrect.js
// VERSION 5.3 — ULTRA SAFE NODE-ONLY PATCH (DOCX + PPTX + XLSX)
// ✅ Applies corrections ONLY inside a single XML text node (<w:t>, <a:t>, <t>)
// ✅ Prevents DOCX/PPTX corruption from cross-run / cross-node replacements
// ✅ Rejects invalid XML characters & unsafe corrections
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
      // high surrogate must be followed by low surrogate
      const n = s.charCodeAt(i + 1);
      if (!(n >= 0xdc00 && n <= 0xdfff)) return true;
      i++;
    } else if (c >= 0xdc00 && c <= 0xdfff) {
      // low surrogate without preceding high surrogate
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

  // prevent injecting illegal XML chars
  if (hasIllegalXmlChars(to) || hasUnpairedSurrogate(to)) return false;

  // prevent absurdly long replacement (guardrail)
  if (to.length > 2000) return false;

  return true;
}

// ---------------- XML Validation (basic but effective) ----------------
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
  // Light validation only
  return true;
}

// ---------------- Context scoring (node-local) ----------------
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

// Find best occurrence inside one node only
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

// ---------------- Core: apply corrections inside nodes only ----------------
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

  // Assign each correction to ONE best node (no cross-node)
  const plansByNode = new Map(); // nodeIndex -> array of plans
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

  console.log(
    `[${label}] Nodes=${nodes.length}, safeCorrections=${safeErrors.length}, planned(node-only)=${planned}`
  );

  if (!planned) return xml;

  // Apply replacements inside each node (reverse order inside node)
  for (const [nodeIndex, plans] of plansByNode.entries()) {
    const node = nodes[nodeIndex];
    if (!node || !node.text) continue;

    plans.sort((a, b) => b.pos - a.pos);

    let text = node.text;
    let appliedInNode = 0;

    for (const p of plans) {
      const inside = text.slice(p.pos, p.end);
      if (inside.toLowerCase() !== String(p.err.error).toLowerCase()) {
        continue; // text shifted by previous replacements, skip
      }
      const before = text.slice(0, p.pos);
      const after = text.slice(p.end);
      text = before + p.err.correction + after;
      appliedInNode++;
    }

    if (appliedInNode > 0) {
      node.text = text;
      stats.changedTextNodes = (stats.changedTextNodes || 0) + 1;
      // example: take first plan
      pushExample(stats, plans[plans.length - 1].err.error, plans[plans.length - 1].err.correction);
    }
  }

  // Rebuild xml from nodes list (stable)
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
      // ✅ node-only: <w:t ...>...</w:t>
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
    console.warn(`[DOCX] ⚠️ Failsafe triggered (some parts kept original XML)`);
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
      // ✅ node-only: <a:t>...</a:t>
      xml = applyCorrectionsNodeOnly(
        xml,
        /(<a:t>)([\s\S]*?)(<\/a:t>)/g,
        spellingErrors,
        stats,
        validatePptxXml,
        `PPTX:${p}`
      );
    }

    console.log(`[PPTX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  if (stats.failsafeTriggered) {
    console.warn(`[PPTX] ⚠️ Failsafe triggered (some parts kept original XML)`);
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
      // sharedStrings uses <t>...</t> (can have attrs sometimes)
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
    console.warn(`[XLSX] ⚠️ Failsafe triggered (some parts kept original XML)`);
  }

  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

export default { correctDOCXText, correctPPTXText, correctXLSXText };
