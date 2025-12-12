// lib/officeCorrect.js
// VERSION 4.2 — ANCHORED & SAFE (multi-run capable)
// ✔ No raw XML replace
// ✔ Context-aware correction
// ✔ DOCX + PPTX stable
// ✔ Can fix errors spanning multiple <w:t>/<a:t> nodes safely

import JSZip from "jszip";

// ============================================================
// Utils
// ============================================================

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

// ============================================================
// Context-based matching (CRITICAL PART)
// ============================================================

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

// ============================================================
// Apply corrections inside segmented XML (DOCX / PPTX)
// multi-run safe strategy
// ============================================================

function applyCorrectionsToSegments(xml, tagRegex, spellingErrors, stats) {
  let modified = xml;

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

  if (!segments.length) return modified;

  for (const err of spellingErrors || []) {
    if (!err?.error || typeof err?.correction !== "string") continue;

    const idx = findBestOccurrence(allText, err);
    if (idx === -1) continue;

    const end = idx + err.error.length;

    // find segments overlapping [idx, end)
    const overlapped = [];
    for (let i = 0; i < segments.length; i++) {
      const s = segments[i];
      if (s.allTextEnd <= idx || s.allTextStart >= end) continue;
      overlapped.push({ i, s });
    }
    if (!overlapped.length) continue;

    // ---------- Single segment: easy replace ----------
    if (overlapped.length === 1) {
      const { i: segIndex, s } = overlapped[0];

      const startInSeg = Math.max(0, idx - s.allTextStart);
      const endInSeg = Math.min(s.text.length, end - s.allTextStart);

      const inside = s.text.slice(startInSeg, endInSeg);
      if (inside.toLowerCase() !== String(err.error).toLowerCase()) continue;

      const newText =
        s.text.slice(0, startInSeg) + err.correction + s.text.slice(endInSeg);

      const newTag = s.full.replace(s.raw, encode(newText));

      modified = modified.slice(0, s.xmlStart) + newTag + modified.slice(s.xmlEnd);

      const diff = newTag.length - s.full.length;
      for (let j = segIndex + 1; j < segments.length; j++) {
        segments[j].xmlStart += diff;
        segments[j].xmlEnd += diff;
      }

      s.full = newTag;
      s.raw = encode(newText);
      s.text = newText;
      s.xmlEnd += diff;

      stats.changedTextNodes = (stats.changedTextNodes || 0) + 1;
      pushExample(stats, err.error, err.correction);
      continue;
    }

    // ---------- Multi segment: merge->replace->split by original lengths ----------
    const first = overlapped[0];
    const mergedText = overlapped.map((o) => o.s.text).join("");

    const mergedStart = idx - first.s.allTextStart;
    const mergedEnd = mergedStart + err.error.length;

    const insideMerged = mergedText.slice(mergedStart, mergedEnd);
    if (insideMerged.toLowerCase() !== String(err.error).toLowerCase()) continue;

    const newMerged =
      mergedText.slice(0, mergedStart) + err.correction + mergedText.slice(mergedEnd);

    const lengths = overlapped.map((o) => o.s.text.length);
    const totalLen = lengths.reduce((a, b) => a + b, 0);
    if (totalLen <= 0) continue;

    let cursor = 0;
    const splitTexts = lengths.map((len, k) => {
      if (k === lengths.length - 1) return newMerged.slice(cursor);
      const part = newMerged.slice(cursor, cursor + len);
      cursor += len;
      return part;
    });

    // apply split parts per segment in order
    for (let n = 0; n < overlapped.length; n++) {
      const { i: segIndex, s } = overlapped[n];
      const nextText = splitTexts[n] ?? "";
      const nextTag = s.full.replace(s.raw, encode(nextText));

      modified = modified.slice(0, s.xmlStart) + nextTag + modified.slice(s.xmlEnd);

      const diff = nextTag.length - s.full.length;
      for (let j = segIndex + 1; j < segments.length; j++) {
        segments[j].xmlStart += diff;
        segments[j].xmlEnd += diff;
      }

      s.full = nextTag;
      s.raw = encode(nextText);
      s.text = nextText;
      s.xmlEnd += diff;
    }

    stats.changedTextNodes = (stats.changedTextNodes || 0) + overlapped.length;
    pushExample(stats, err.error, err.correction);
  }

  return modified;
}

// ============================================================
// DOCX
// ============================================================

export async function correctDOCXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);

  const targets = [
    "word/document.xml",
    ...Object.keys(zip.files).filter((k) => /word\/(header|footer)\d+\.xml$/.test(k)),
  ];

  const stats = { changedTextNodes: 0, examples: [] };
  const spellingErrors = options.spellingErrors || [];

  for (const p of targets) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async("string");

    if (spellingErrors.length) {
      xml = applyCorrectionsToSegments(
        xml,
        /<w:t[^>]*>([\s\S]*?)<\/w:t>/g,
        spellingErrors,
        stats
      );
    }

    zip.file(p, xml);
  }

  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

// ============================================================
// PPTX
// ============================================================

export async function correctPPTXText(buffer, correctFn, options = {}) {
  const zip = await JSZip.loadAsync(buffer);
  const slides = Object.keys(zip.files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k));

  const stats = { changedTextNodes: 0, examples: [] };
  const spellingErrors = options.spellingErrors || [];

  for (const p of slides) {
    if (!zip.file(p)) continue;
    let xml = await zip.file(p).async("string");

    if (spellingErrors.length) {
      xml = applyCorrectionsToSegments(
        xml,
        /<a:t>([\s\S]*?)<\/a:t>/g,
        spellingErrors,
        stats
      );
    }

    zip.file(p, xml);
  }

  return { outBuffer: await zip.generateAsync({ type: "nodebuffer" }), stats };
}

export default { correctDOCXText, correctPPTXText };
