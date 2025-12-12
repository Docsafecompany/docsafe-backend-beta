// lib/officeCorrect.js
// VERSION 4.0 — ANCHORED & SAFE
// ✔ No raw XML replace
// ✔ Context-aware correction
// ✔ Human-like behavior
// ✔ DOCX + PPTX stable

import JSZip from "jszip";

// ============================================================
// Utils
// ============================================================

const decode = s =>
  (s || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");

const encode = s =>
  (s || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");

const normalize = s => (s || "").replace(/\s+/g, " ").trim();
const removeSpaces = s => (s || "").replace(/\s+/g, "");

function pushExample(stats, before, after) {
  if (!stats.examples) stats.examples = [];
  if (stats.examples.length < 10) {
    stats.examples.push({
      before: before.slice(0, 140),
      after: after.slice(0, 140),
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
  const winAfter = normalize(
    allText.slice(idx + e.length, idx + e.length + 60)
  );

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
// ============================================================

function applyCorrectionsToSegments(xml, tagRegex, spellingErrors, stats) {
  let modified = xml;

  const segments = [];
  let allText = "";
  let m;

  tagRegex.lastIndex = 0;
  while ((m = tagRegex.exec(xml)) !== null) {
    const raw = m[1];
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

  for (const err of spellingErrors) {
    const idx = findBestOccurrence(allText, err);
    if (idx === -1) continue;

    const end = idx + err.error.length;
    let firstSeg = -1;

    for (let i = 0; i < segments.length; i++) {
      const s = segments[i];
      if (s.allTextEnd <= idx || s.allTextStart >= end) continue;

      if (firstSeg === -1) {
        firstSeg = i;

        const startInSeg = Math.max(0, idx - s.allTextStart);
        const before = s.text.slice(0, startInSeg);
        const after =
          end <= s.allTextEnd
            ? s.text.slice(end - s.allTextStart)
            : "";

        const newText = before + err.correction + after;
        const newTag = s.full.replace(s.raw, encode(newText));

        modified =
          modified.slice(0, s.xmlStart) +
          newTag +
          modified.slice(s.xmlEnd);

        const diff = newTag.length - s.full.length;
        for (let j = i + 1; j < segments.length; j++) {
          segments[j].xmlStart += diff;
          segments[j].xmlEnd += diff;
        }

        stats.changedTextNodes++;
        pushExample(stats, err.error, err.correction);
      } else {
        // Clear overlapping segments
        const emptyTag = s.full.replace(s.raw, "");
        modified =
          modified.slice(0, s.xmlStart) +
          emptyTag +
          modified.slice(s.xmlEnd);

        const diff = emptyTag.length - s.full.length;
        for (let j = i + 1; j < segments.length; j++) {
          segments[j].xmlStart += diff;
          segments[j].xmlEnd += diff;
        }
      }
    }
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
    ...Object.keys(zip.files).filter(k =>
      /word\/(header|footer)\d+\.xml$/.test(k)
    ),
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
  const slides = Object.keys(zip.files).filter(k =>
    /^ppt\/slides\/slide\d+\.xml$/.test(k)
  );

  const stats = { changedTextNodes: 0, examples: [] };
  const spellingErrors = options.spellingErrors || [];

  for (const p of slides) {
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
