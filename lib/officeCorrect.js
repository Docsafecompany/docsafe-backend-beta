// lib/officeCorrect.js
// VERSION 5.8 — PPTX ZIP INTEGRITY FIX (RELATIONSHIPS + CONTENT TYPES) + FIX WORD-JOINING (DOCX + PPTX + XLSX)
// ✅ Accepts approved list as IDs OR objects
// ✅ Supports legacy {error, correction, contextBefore, contextAfter}
// ✅ Supports new {before, after, anchor:{left,right}} (and optional contextBefore/After)
// ✅ DOCX/XLSX: Node-only safe corrections (no XML corruption)
// ✅ PPTX: Cross-node corrections WITHOUT naive redistribution (fixes "not applied" cases)
// ✅ Removes HTML-like marker artifacts from final text
// ✅ CRITICAL FIX: no trim() or global whitespace collapse on node text (prevents word-joining)
// ✅ NEW: PPTX zip integrity post-clean: removes orphaned .rels + [Content_Types].xml overrides
// ✅ Failsafe: returns original XML if validation fails
// ✅ Detailed logging

import JSZip from "jszip";
import path from "path";

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

// NOTE: used only for scoring contexts (not for writing nodes)
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
// Removes accidental HTML-ish fragments that should never exist in Office text
// CRITICAL: Do NOT trim() node text. Leading/trailing spaces can be meaningful across runs.
// Example in preview: class="correction-marker" data-index="13">
function stripMarkerArtifacts(text) {
  const t = String(text ?? "");
  if (!t) return t;

  // Preserve leading/trailing whitespace EXACTLY (prevents word-joining across adjacent nodes)
  const lead = (t.match(/^\s+/) || [""])[0];
  const trail = (t.match(/\s+$/) || [""])[0];

  // Work on the core only
  let core = t.slice(lead.length, t.length - trail.length);

  // remove whole attribute blocks if present
  core = core.replace(/\s*class="correction-marker"\s*/gi, " ");
  core = core.replace(/\s*data-index="[^"]*"\s*/gi, " ");

  // remove leftover fragments like: class=... data-index=...
  core = core.replace(/\bclass\s*=\s*["']correction-marker["']\b/gi, "");
  core = core.replace(/\bdata-index\s*=\s*["'][^"']*["']\b/gi, "");

  // remove stray tag-like endings introduced by bugs
  // replace one-or-more > with a single space (do NOT trim)
  core = core.replace(/>+/g, " ");

  // Cleanup ONLY repeated plain spaces inside the core.
  // Do not collapse newlines/tabs here to avoid changing structure unexpectedly.
  core = core.replace(/ {2,}/g, " ");

  return lead + core + trail;
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

// ---------------- Correction normalization ----------------
// Accepts:
// - legacy: { id, error, correction, contextBefore, contextAfter }
// - new:    { id, before, after, anchor:{left,right}, contextBefore?, contextAfter? }
function toLegacyCorrectionObject(input) {
  if (!input) return null;

  // Already legacy
  if (typeof input.error === "string" && typeof input.correction === "string") {
    return {
      id: input.id || null,
      error: input.error,
      correction: input.correction,
      contextBefore: input.contextBefore || input.anchor?.left || "",
      contextAfter: input.contextAfter || input.anchor?.right || "",
    };
  }

  // New format
  if (typeof input.before === "string" && typeof input.after === "string") {
    return {
      id: input.id || null,
      error: input.before,
      correction: input.after,
      contextBefore: input.contextBefore || input.anchor?.left || "",
      contextAfter: input.contextAfter || input.anchor?.right || "",
    };
  }

  return null;
}

// approvedList can be: [] | ["id1","id2"] | [{...},{...}]
function normalizeApprovedList(detectedList = [], approvedList = []) {
  if (!approvedList || approvedList.length === 0) return detectedList;

  if (typeof approvedList[0] === "string") {
    const set = new Set(approvedList);
    return (detectedList || []).filter((e) => e && e.id && set.has(e.id));
  }

  return approvedList;
}

function isCorrectionSafe(input) {
  const err = toLegacyCorrectionObject(input);
  if (!err) return false;

  const from = err.error;
  const to = err.correction;

  if (!from || !String(from).trim()) return false;
  if (from === to) return false;

  if (hasIllegalXmlChars(to) || hasUnpairedSurrogate(to)) return false;
  if (String(to).length > 2000) return false;

  return true;
}

// ---------------- XML Validation ----------------
function validateDocxXml(xml) {
  if (!xml || typeof xml !== "string" || xml.length < 50) return false;

  const openWT = (xml.match(/<w:t\b[^>]*>/g) || []).length;
  const closeWT = (xml.match(/<\/w:t>/g) || []).length;
  if (openWT !== closeWT) {
    console.error(
      `[VALIDATE DOCX] <w:t> mismatch: ${openWT} open vs ${closeWT} close`
    );
    return false;
  }

  const openWP = (xml.match(/<w:p\b[^>]*>/g) || []).length;
  const closeWP = (xml.match(/<\/w:p>/g) || []).length;
  if (openWP !== closeWP && (openWP > 0 || closeWP > 0)) {
    console.error(
      `[VALIDATE DOCX] <w:p> mismatch: ${openWP} open vs ${closeWP} close`
    );
    return false;
  }

  return true;
}

function validatePptxXml(xml) {
  if (!xml || typeof xml !== "string" || xml.length < 50) return false;
  const openAT = (xml.match(/<a:t>/g) || []).length;
  const closeAT = (xml.match(/<\/a:t>/g) || []).length;
  if (openAT !== closeAT) {
    console.error(
      `[VALIDATE PPTX] <a:t> mismatch: ${openAT} open vs ${closeAT} close`
    );
    return false;
  }
  return true;
}

function validateXlsxXml(xml) {
  if (!xml || typeof xml !== "string" || xml.length < 20) return false;
  return true;
}

// ---------------- PPTX ZIP integrity helpers ----------------
// Removes relationships (.rels) pointing to missing Targets.
// Also cleans [Content_Types].xml Override PartName entries for missing parts.
//
// Why this matters:
// - If another sanitizer step deletes ppt/notesSlides/*.xml or ppt/embeddings/*
//   but leaves references in *.rels or [Content_Types].xml, PowerPoint will "repair"
//   the file or claim it's corrupted.
//
// This post-clean is safe to run even if nothing was deleted.
function resolveRelsTarget(relsFilePath, target) {
  const t = String(target || "").trim();
  if (!t) return null;

  // Absolute part inside package, usually like "/ppt/slides/slide1.xml"
  if (t.startsWith("/")) return t.slice(1);

  const relsDir = path.posix.dirname(relsFilePath);
  const resolved = path.posix.normalize(path.posix.join(relsDir, t));

  // Prevent escaping above root
  if (resolved.startsWith("..")) return null;
  return resolved;
}

function getXmlRootOpenTag(xml, rootName) {
  const re = new RegExp(`<${rootName}\\b[^>]*>`, "i");
  const m = String(xml || "").match(re);
  return m ? m[0] : `<${rootName}>`;
}

async function cleanOrphanedRelationships(zip) {
  const relsFiles = Object.keys(zip.files).filter((f) => f.endsWith(".rels"));
  if (!relsFiles.length) return zip;

  let totalRemoved = 0;

  for (const relsPath of relsFiles) {
    const file = zip.file(relsPath);
    if (!file) continue;

    const xml = await file.async("string");
    if (!xml || !xml.includes("<Relationship")) continue;

    const rootOpen = getXmlRootOpenTag(xml, "Relationships");
    const rootClose = "</Relationships>";

    const relTagRe = /<Relationship\b[^>]*\/>/gi;
    const relTags = xml.match(relTagRe) || [];

    const kept = [];
    let removedHere = 0;

    for (const tag of relTags) {
      const targetMatch = tag.match(/\bTarget\s*=\s*"([^"]+)"/i);
      const target = targetMatch ? targetMatch[1] : "";

      // If no target, keep (rare but safer)
      if (!target) {
        kept.push(tag);
        continue;
      }

      const resolved = resolveRelsTarget(relsPath, target);
      if (!resolved) {
        // Can't resolve safely -> remove
        removedHere++;
        continue;
      }

      if (zip.file(resolved)) {
        kept.push(tag);
      } else {
        removedHere++;
      }
    }

    if (removedHere > 0) {
      totalRemoved += removedHere;

      // Preserve XML declaration if present
      const decl = xml.startsWith("<?xml")
        ? xml.match(/^\<\?xml[^>]*\?>\s*/i)?.[0] || ""
        : "";

      const rebuilt = decl + rootOpen + "\n" + kept.join("\n") + "\n" + rootClose;

      zip.file(relsPath, rebuilt);
      console.log(
        `[PPTX:ZIP] Cleaned orphan rels: ${relsPath} removed=${removedHere} kept=${kept.length}`
      );
    }
  }

  if (totalRemoved > 0) {
    console.log(`[PPTX:ZIP] Total orphaned relationships removed: ${totalRemoved}`);
  }

  return zip;
}

async function cleanContentTypes(zip) {
  const ctPath = "[Content_Types].xml";
  const file = zip.file(ctPath);
  if (!file) return zip;

  const xml = await file.async("string");
  if (!xml || !xml.includes("<Override")) return zip;

  const rootOpen = getXmlRootOpenTag(xml, "Types");
  const rootClose = "</Types>";

  const overrideRe = /<Override\b[^>]*\/>/gi;
  const overrides = xml.match(overrideRe) || [];

  let removed = 0;
  const keptOverrides = [];

  for (const tag of overrides) {
    const partMatch = tag.match(/\bPartName\s*=\s*"([^"]+)"/i);
    const partName = partMatch ? partMatch[1] : "";
    if (!partName) {
      keptOverrides.push(tag);
      continue;
    }

    const partPath = partName.startsWith("/") ? partName.slice(1) : partName;
    if (zip.file(partPath)) {
      keptOverrides.push(tag);
    } else {
      removed++;
    }
  }

  if (removed === 0) return zip;

  // Keep all <Default .../> entries untouched
  const defaultRe = /<Default\b[^>]*\/>/gi;
  const defaults = xml.match(defaultRe) || [];

  const decl = xml.startsWith("<?xml")
    ? xml.match(/^\<\?xml[^>]*\?>\s*/i)?.[0] || ""
    : "";
  const rebuilt =
    decl +
    rootOpen +
    "\n" +
    defaults.join("\n") +
    (defaults.length ? "\n" : "") +
    keptOverrides.join("\n") +
    "\n" +
    rootClose;

  zip.file(ctPath, rebuilt);
  console.log(
    `[PPTX:ZIP] Cleaned [Content_Types].xml overrides removed=${removed} kept=${keptOverrides.length}`
  );

  return zip;
}

// ---------------- Context scoring ----------------
function scoreContextInNode(nodeText, idx, err) {
  const e = err.error;
  const before = normalize(err.contextBefore || "");
  const after = normalize(err.contextAfter || "");
  let score = 0;

  const winBefore = normalize(nodeText.slice(Math.max(0, idx - 60), idx));
  const winAfter = normalize(
    nodeText.slice(idx + e.length, idx + e.length + 60)
  );

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
function applyCorrectionsNodeOnly(
  xml,
  nodeRegex,
  spellingErrors,
  stats,
  validateFn,
  label
) {
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

  // Normalize + filter safe corrections
  const safeErrors = (spellingErrors || [])
    .map(toLegacyCorrectionObject)
    .filter(Boolean)
    .filter(isCorrectionSafe);

  if (!safeErrors.length) {
    console.log(`[${label}] No safe corrections to apply`);
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

  console.log(
    `[${label}] Nodes=${nodes.length}, safeCorrections=${safeErrors.length}, planned=${planned}, sanitized=${stats.sanitizedNodes || 0}`
  );

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
      pushExample(
        stats,
        plans[plans.length - 1].err.error,
        plans[plans.length - 1].err.correction
      );
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

  // Normalize + filter safe corrections
  const safeErrors = (spellingErrors || [])
    .map(toLegacyCorrectionObject)
    .filter(Boolean)
    .filter(isCorrectionSafe);

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
    console.log(
      `[${label}] No corrections found in concatenated text (sanitized=${stats.sanitizedNodes || 0})`
    );
  }

  // Apply from end to start to avoid index shifts in mapping calculations
  corrections.sort((a, b) => b.start - a.start);

  for (const corr of corrections) {
    const { err, start, end } = corr;

    if (start < 0 || end > map.length || end <= start) continue;

    const startRef = map[start];
    const endRef = map[end - 1];
    if (!startRef || !endRef) continue;

    const startNodeIdx = startRef.nodeIndex;
    const endNodeIdx = endRef.nodeIndex;

    // If entirely within one node
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

    // Spans multiple nodes (rare)
    const startNode = nodes[startNodeIdx];
    const endNode = nodes[endNodeIdx];

    const startLocal = startRef.localPos;
    const endLocalEnd = endRef.localPos + 1;

    const startBefore = startNode.text.slice(0, startLocal);
    const endAfter = endNode.text.slice(endLocalEnd);

    const replaced = allText.slice(start, end);
    if (replaced.toLowerCase() !== err.error.toLowerCase()) {
      console.warn(
        `[${label}] Skip cross-node correction (mapping stale): "${err.error}"`
      );
      continue;
    }

    startNode.text = startBefore + err.correction;

    for (let ni = startNodeIdx + 1; ni < endNodeIdx; ni++) {
      nodes[ni].text = "";
    }

    endNode.text = endAfter;

    stats.changedTextNodes = (stats.changedTextNodes || 0) + 1;
    pushExample(stats, err.error, err.correction);
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
    ...Object.keys(zip.files).filter((k) =>
      /word\/(header|footer)\d+\.xml$/.test(k)
    ),
    "word/footnotes.xml",
    "word/endnotes.xml",
  ];

  const stats = {
    changedTextNodes: 0,
    examples: [],
    failsafeTriggered: false,
    sanitizedNodes: 0,
  };

  // ✅ Support: detectedSpellingErrors + approved IDs/objects in spellingErrors
  const detected = Array.isArray(options.detectedSpellingErrors)
    ? options.detectedSpellingErrors
    : [];
  const approved = Array.isArray(options.spellingErrors)
    ? options.spellingErrors
    : [];
  const spellingErrors = normalizeApprovedList(
    detected.length ? detected : approved,
    approved
  );

  const existingTargets = targets.filter((t) => zip.file(t));
  console.log(
    `[DOCX] Processing ${existingTargets.length} files, corrections=${spellingErrors.length}`
  );

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
    (k) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(k) ||
      /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(k)
  );

  const stats = {
    changedTextNodes: 0,
    examples: [],
    failsafeTriggered: false,
    sanitizedNodes: 0,
  };

  // ✅ Support: detectedSpellingErrors + approved IDs/objects in spellingErrors
  const detected = Array.isArray(options.detectedSpellingErrors)
    ? options.detectedSpellingErrors
    : [];
  const approved = Array.isArray(options.spellingErrors)
    ? options.spellingErrors
    : [];
  const spellingErrors = normalizeApprovedList(
    detected.length ? detected : approved,
    approved
  );

  console.log(
    `[PPTX] Processing ${targets.length} files, corrections=${spellingErrors.length}`
  );

  for (const p of targets) {
    const file = zip.file(p);
    if (!file) continue;

    let xml = await file.async("string");
    const originalLength = xml.length;

    xml = applyPptxCrossNodeCorrections(xml, spellingErrors, stats, `PPTX:${p}`);

    console.log(`[PPTX] ${p}: ${originalLength} → ${xml.length} chars`);
    zip.file(p, xml);
  }

  // ✅ NEW: clean relationships + content types to avoid corrupted PPTX after other deletions in pipeline
  await cleanOrphanedRelationships(zip);
  await cleanContentTypes(zip);

  if (stats.failsafeTriggered) console.warn(`[PPTX] ⚠️ Failsafe triggered`);
  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

// ---------------- XLSX ----------------
export async function correctXLSXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);

  const targets = [
    "xl/sharedStrings.xml",
    ...Object.keys(zip.files).filter((k) =>
      /^xl\/worksheets\/sheet\d+\.xml$/.test(k)
    ),
  ];

  const stats = {
    changedTextNodes: 0,
    examples: [],
    failsafeTriggered: false,
    sanitizedNodes: 0,
  };

  // ✅ Support: detectedSpellingErrors + approved IDs/objects in spellingErrors
  const detected = Array.isArray(options.detectedSpellingErrors)
    ? options.detectedSpellingErrors
    : [];
  const approved = Array.isArray(options.spellingErrors)
    ? options.spellingErrors
    : [];
  const spellingErrors = normalizeApprovedList(
    detected.length ? detected : approved,
    approved
  );

  const existingTargets = targets.filter((t) => zip.file(t));
  console.log(
    `[XLSX] Processing ${existingTargets.length} files, corrections=${spellingErrors.length}`
  );

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
