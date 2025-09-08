// lib/officeXml.js
import JSZip from "jszip";
import { aiCorrectText, aiRephraseText } from "./ai.js";

/* ============== Utils ============== */
const esc = (s = "") =>
  String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const unesc = (s = "") =>
  String(s).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");

function conservativeSanitize(s) {
  let t = String(s);
  t = t.replace(/[ \t]{2,}/g, " ");
  t = t.replace(/\s+,/g, ",");
  t = t.replace(/,([^\s\n])/g, ", $1");
  t = t.replace(/\s+([:;!?)\]\}])/g, "$1");
  t = t.replace(/([.:;!?])([A-Za-z0-9])/g, "$1 $2");
  t = t.replace(/([.,;:!?@])\1+/g, "$1");
  return t.trim();
}

/* ============== Pré-correction déterministe (anti “soc ial”, “enablin g”, “gdigital”…) ============== */
function preFixIntraWordSplits(s) {
  let t = String(s);

  // A) Liste blanche vue dans tes fichiers
  t = t.replace(/\bsoc\s+ial\b/gi, "social");
  t = t.replace(/\bcommu\s+nication\b/gi, "communication");
  t = t.replace(/\benablin\s+g\b/gi, "enabling");
  t = t.replace(/\bth\s+e\b/gi, "the");
  t = t.replace(/\bp\s+otential\b/gi, "potential");
  t = t.replace(/\bdis\s+connection\b/gi, "disconnection");
  t = t.replace(/\brig\s+hts\b/gi, "rights");
  t = t.replace(/\bco\s+mmunication\b/gi, "communication");
  t = t.replace(/\bcorpo,\s*rations\b/gi, "corporations");
  t = t.replace(/\bo\s+f\b/gi, "of");
  t = t.replace(/\bc\s+an\b/gi, "can");

  // B) Non-mots récurrents -> vrais mots
  t = t.replace(/\bgggdigital\b/gi, "digital");
  t = t.replace(/\bgdigital\b/gi, "digital");

  // C) Bruit: lettre répétée en tête (gggdigital -> digital, rrights -> rights)
  t = t.replace(/\b([a-z])\1{2,}([a-z]+)/gi, (_m, a, rest) => a + rest);

  // D) Heuristiques suffixes très fréquents
  t = t.replace(
    /\b([A-Za-z]{3,})\s+(tion|sion|ment|ness|ship|ance|ence|hood|ward|tial|cial|ing|ed|ly|al|ary|able|ible|ism|ist|ize|ise|ized|ised|ization|isation|tions|ments)\b/gi,
    (_m, a, suf) => a + suf
  );

  // E) 1 lettre + espace + reste (th e -> the, p otential -> potential) — prudence
  t = t.replace(/\b([A-Za-z])\s+([A-Za-z]{2,})\b/g, (_m, a, rest) => {
    if (/^(a|i|à|y)$/i.test(a)) return _m; // ne pas coller de vrais mots uniques
    return a + rest;
  });

  // F) Espace multiple au milieu d’un mot
  t = t.replace(/\b([A-Za-z])\s{2,}([A-Za-z]+)\b/g, (_m, a, rest) => a + rest);

  // G) Normalisation espaces/ponctuation légère
  t = t.replace(/[ \t]{2,}/g, " ");
  t = t.replace(/\s+,/g, ",");
  t = t.replace(/,([^\s\n])/g, ", $1");
  t = t.replace(/\s+([:;!?)\]\}])/g, "$1");
  t = t.replace(/([.:;!?])([A-Za-z0-9])/g, "$1 $2");
  t = t.replace(/([.,;:!?@])\1+/g, "$1");

  return t;
}

/* ============== Helpers re-split ============== */
function splitToRuns(corrected, runTexts) {
  const lens = runTexts.map((r) => String(r).length);
  const out = [];
  let i = 0;
  for (const L of lens) {
    out.push(corrected.slice(i, i + L));
    i += L;
  }
  if (i < corrected.length && out.length) out[out.length - 1] += corrected.slice(i);
  return out;
}

function rebuildParagraph(paraXml, runs, pieces) {
  let rebuilt = paraXml;
  for (let i = 0; i < runs.length; i++) {
    const { open, close } = runs[i];
    const seg = `${open}${esc(pieces[i] ?? "")}${close}`;
    const rx = new RegExp(
      `(${open.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")})([\\s\\S]*?)(${close.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")})`
    );
    rebuilt = rebuilt.replace(rx, seg);
  }
  return rebuilt;
}

/* ============== DOCX ============== */
async function processDocxXml(xml, mode, stats) {
  const paraRx = /<w:p\b[\s\S]*?<\/w:p>/g;
  const runRx = /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g;

  const parts = [];
  let last = 0, m, changed = 0, total = 0;

  while ((m = paraRx.exec(xml))) {
    parts.push(xml.slice(last, m.index));
    let para = m[0];

    const runs = [];
    let rm;
    while ((rm = runRx.exec(para))) {
      runs.push({ open: rm[1], text: unesc(rm[2]), close: rm[3] });
    }
    if (!runs.length) { parts.push(para); last = paraRx.lastIndex; continue; }

    total++;
    const original = runs.map(r => r.text).join("");

    const pre = preFixIntraWordSplits(original);
    let aiOut = null;
    try {
      aiOut = mode === "V2" ? await aiRephraseText(pre) : await aiCorrectText(pre);
    } catch {}
    const corrected = String(aiOut || conservativeSanitize(pre));
    if (corrected !== original) changed++;

    const pieces = splitToRuns(corrected, runs.map(r => r.text));
    para = rebuildParagraph(para, runs, pieces);

    parts.push(para);
    last = paraRx.lastIndex;
  }
  parts.push(xml.slice(last));

  if (stats) stats.push({ kind: "docx", total, changed });
  return parts.join("");
}

/* ============== PPTX ============== */
async function processPptxXml(xml, mode, stats) {
  const paraRx = /<a:p\b[\s\S]*?<\/a:p>/g;
  const runRx = /(<a:t[^>]*>)([\s\S]*?)(<\/a:t>)/g;

  const parts = [];
  let last = 0, m, changed = 0, total = 0;

  while ((m = paraRx.exec(xml))) {
    parts.push(xml.slice(last, m.index));
    let para = m[0];

    const runs = [];
    let rm;
    while ((rm = runRx.exec(para))) {
      runs.push({ open: rm[1], text: unesc(rm[2]), close: rm[3] });
    }
    if (!runs.length) { parts.push(para); last = paraRx.lastIndex; continue; }

    total++;
    const original = runs.map(r => r.text).join("");

    const pre = preFixIntraWordSplits(original);
    let aiOut = null;
    try {
      aiOut = mode === "V2" ? await aiRephraseText(pre) : await aiCorrectText(pre);
    } catch {}
    const corrected = String(aiOut || conservativeSanitize(pre));
    if (corrected !== original) changed++;

    const pieces = splitToRuns(corrected, runs.map(r => r.text));
    para = rebuildParagraph(para, runs, pieces);

    parts.push(para);
    last = paraRx.lastIndex;
  }
  parts.push(xml.slice(last));

  if (stats) stats.push({ kind: "pptx", total, changed });
  return parts.join("");
}

/* ============== Public API ============== */
export async function processDocxBuffer(buffer, mode = "V1") {
  const stats = [];
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) =>
    /^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i.test(k)
  );
  for (const f of targets) {
    const xml = await zip.file(f).async("string");
    const updated = await processDocxXml(xml, mode, stats);
    zip.file(f, updated);
  }
  const out = await zip.generateAsync({ type: "nodebuffer" });
  out.__stats = stats;
  return out;
}

export async function processPptxBuffer(buffer, mode = "V1") {
  const stats = [];
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) =>
    /^(ppt\/slides\/slide\d+\.xml|ppt\/notesSlides\/notesSlide\d+\.xml)$/i.test(k)
  );
  for (const f of targets) {
    const xml = await zip.file(f).async("string");
    const updated = await processPptxXml(xml, mode, stats);
    zip.file(f, updated);
  }
  const out = await zip.generateAsync({ type: "nodebuffer" });
  out.__stats = stats;
  return out;
}

/* ============== Analyse sans écriture (debug) ============== */
export async function analyzeDocxBuffer(buffer, mode = "V1") {
  const stats = [];
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) =>
    /^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i.test(k)
  );
  for (const f of targets) {
    const xml = await zip.file(f).async("string");
    await processDocxXml(xml, mode, stats);
  }
  const total = stats.reduce((a, s) => a + s.total, 0);
  const changed = stats.reduce((a, s) => a + s.changed, 0);
  return { totalParagraphs: total, changedParagraphs: changed, perPart: stats };
}
