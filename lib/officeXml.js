// lib/officeXml.js
import JSZip from "jszip";
import { aiCorrectText, aiRephraseText } from "./ai.js";

const esc = (s="") => String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const unesc = (s="") => String(s).replace(/&lt;/g,"<").replace(/&gt;/g,">").replace(/&amp;/g,"&");

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

function splitToRuns(corrected, runTexts) {
  const lens = runTexts.map(t => String(t).length);
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
    const rx = new RegExp(`(${open.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")})([\\s\\S]*?)(${close.replace(/[.*+?^${}()|[\]\\]/g,"\\$&")})`);
    rebuilt = rebuilt.replace(rx, seg);
  }
  return rebuilt;
}

async function processDocxXml(xml, mode) {
  const paraRx = /<w:p\b[\s\S]*?<\/w:p>/g;
  const runRx = /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g;
  const parts = [];
  let last = 0, m;

  while ((m = paraRx.exec(xml))) {
    parts.push(xml.slice(last, m.index));
    let para = m[0];
    const runs = [];
    let rm;
    while ((rm = runRx.exec(para))) {
      runs.push({ open: rm[1], text: unesc(rm[2]), close: rm[3] });
    }
    if (!runs.length) { parts.push(para); last = paraRx.lastIndex; continue; }

    const original = runs.map(r => r.text).join("");
    let corrected = null;
    try {
      corrected = (mode==="V2") ? await aiRephraseText(original) : await aiCorrectText(original);
    } catch {}
    if (!corrected) corrected = conservativeSanitize(original);

    const pieces = splitToRuns(String(corrected), runs.map(r=>r.text));
    para = rebuildParagraph(para, runs, pieces);
    parts.push(para);
    last = paraRx.lastIndex;
  }
  parts.push(xml.slice(last));
  return parts.join("");
}

async function processPptxXml(xml, mode) {
  const paraRx = /<a:p\b[\s\S]*?<\/a:p>/g;
  const runRx = /(<a:t[^>]*>)([\s\S]*?)(<\/a:t>)/g; // FIX: [\s\S] correct
  const parts = [];
  let last = 0, m;

  while ((m = paraRx.exec(xml))) {
    parts.push(xml.slice(last, m.index));
    let para = m[0];
    const runs = [];
    let rm;
    while ((rm = runRx.exec(para))) {
      runs.push({ open: rm[1], text: unesc(rm[2]), close: rm[3] });
    }
    if (!runs.length) { parts.push(para); last = paraRx.lastIndex; continue; }

    const original = runs.map(r => r.text).join("");
    let corrected = null;
    try {
      corrected = (mode==="V2") ? await aiRephraseText(original) : await aiCorrectText(original);
    } catch {}
    if (!corrected) corrected = conservativeSanitize(original);

    const pieces = splitToRuns(String(corrected), runs.map(r=>r.text));
    para = rebuildParagraph(para, runs, pieces);
    parts.push(para);
    last = paraRx.lastIndex;
  }
  parts.push(xml.slice(last));
  return parts.join("");
}

export async function processDocxBuffer(buffer, mode="V1") {
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter(k =>
    /^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i.test(k)
  );
  for (const f of targets) {
    const xml = await zip.file(f).async("string");
    const updated = await processDocxXml(xml, mode);
    zip.file(f, updated);
  }
  return await zip.generateAsync({ type: "nodebuffer" });
}

export async function processPptxBuffer(buffer, mode="V1") {
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter(k =>
    /^(ppt\/slides\/slide\d+\.xml|ppt\/notesSlides\/notesSlide\d+\.xml)$/i.test(k)
  );
  for (const f of targets) {
    const xml = await zip.file(f).async("string");
    const updated = await processPptxXml(xml, mode);
    zip.file(f, updated);
  }
  return await zip.generateAsync({ type: "nodebuffer" });
}

