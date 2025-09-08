// lib/officeXml.js
import JSZip from "jszip";
import { aiCorrectText, aiRephraseText } from "./ai.js";

// -------- utilitaires XML --------
function xmlEscape(s = "") {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function unescapeXml(s = "") {
  return String(s).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
}

// Découpe sécurisée par regex avec groupes capturés (sélection + remplacement)
function replaceTextNodes(xml, openCloseRegex, transformRunGroup) {
  return xml.replace(openCloseRegex, (full, openTag, inner, closeTag) => {
    const newText = transformRunGroup(inner);
    return `${openTag}${xmlEscape(newText)}${closeTag}`;
  });
}

// Re-découpe une chaîne corrigée pour “remplir” la liste des longueurs de runs d’origine
function splitByRunLengths(corrected, runTexts) {
  const lengths = runTexts.map((t) => String(t).length);
  const out = [];
  let i = 0;
  for (let k = 0; k < lengths.length; k++) {
    const len = lengths[k];
    out.push(corrected.slice(i, i + len));
    i += len;
  }
  // Si la longueur globale a changé (ex. espaces corrigés), on équilibre sur le dernier run
  if (i < corrected.length && out.length) {
    out[out.length - 1] += corrected.slice(i);
  }
  return out;
}

// Transforme un bloc <p> (ou équivalent) :
// - on concatène le contenu des <t> enfants,
// - on passe à l’IA (correction ou reformulation),
// - on re-split le résultat suivant la “forme” d’origine, run par run,
// - on remplace le texte dans chaque <t> sans toucher aux styles.
function transformParagraphXml(xml, tTag, paragraphRegex, tRegex, aiMode /* "V1" | "V2" */, fallbackFn) {
  return xml.replace(paragraphRegex, (fullPara) => {
    // Extraire tous les <t> du paragraphe
    const runs = [];
    const runOpenClose = new RegExp(`(<${tTag}[^>]*>)([\\s\\S]*?)(</${tTag}>)`, "g");
    let m;
    while ((m = runOpenClose.exec(fullPara))) {
      runs.push({ open: m[1], text: unescapeXml(m[2]), close: m[3], start: m.index, len: m[0].length });
    }
    if (runs.length === 0) return fullPara;

    const originalConcat = runs.map(r => r.text).join("");

    // Appel IA (ou fallback si pas de clé)
    const process = async () => {
      try {
        if (aiMode === "V2") {
          const out = await aiRephraseText(originalConcat);
          if (out) return String(out);
        } else {
          const out = await aiCorrectText(originalConcat);
          if (out) return String(out);
        }
      } catch (e) {
        console.warn("AI paragraph call failed -> fallback", e?.message || e);
      }
      return fallbackFn(originalConcat);
    };

    // Comme on est dans un replacer sync, on ne peut pas await.
    // On “marque” le paragraphe et on post-traitera plus bas.
    const TOKEN = `__AI_PARA_TOKEN_${Math.random().toString(36).slice(2)}__`;
    pendingParas.push({ token: TOKEN, runs, fullPara });
    return TOKEN;
  });
}

// Stocke les paragraphes marqués entre deux passes
let pendingParas = [];

// Remplace les tokens par le XML final après appel IA
async function resolvePending(xml, aiMode, fallbackFn) {
  for (const p of pendingParas) {
    const originalConcat = p.runs.map(r => r.text).join("");
    let corrected = null;
    try {
      if (aiMode === "V2") corrected = await aiRephraseText(originalConcat);
      else corrected = await aiCorrectText(originalConcat);
    } catch (e) {
      corrected = null;
    }
    if (!corrected) corrected = fallbackFn(originalConcat);

    const pieces = splitByRunLengths(corrected, p.runs.map(r => r.text));
    let rebuilt = p.fullPara;
    // Remplacement run par run (texte uniquement)
    for (let i = 0; i < p.runs.length; i++) {
      const seg = `${p.runs[i].open}${xmlEscape(pieces[i] ?? "")}${p.runs[i].close}`;
      // on remplace le premier match séquentiellement
      rebuilt = rebuilt.replace(new RegExp(`(<${p.runs[i].open.split("<")[1]}[\\s\\S]*?${p.runs[i].close.replace(/([.*+?^${}()|\[\]\\])/g,"\\$1")})`), seg);
    }

    xml = xml.replace(p.token, rebuilt);
  }
  pendingParas = [];
  return xml;
}

// -------- Traitements Office --------

// Fallback (sans IA) : sanitize conservateur (mini fix espaces/ponctuation)
function conservativeSanitize(s) {
  let t = String(s);
  t = t.replace(/[ \t]{2,}/g, " ");
  t = t.replace(/\s+,/g, ",");
  t = t.replace(/,([^\s\n])/g, ", $1");
  t = t.replace(/\s+([:;!?)\]\}])/g, "$1");
  t = t.replace(/([.:;!?])([A-Za-z0-9])/g, "$1 $2");
  t = t.replace(/([.,;:!?@])\1+/g, "$1");
  return t;
}

export async function processDocxBuffer(buffer, aiTransform /* (text)->Promise<string> handled inside */, mode = "V1") {
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) =>
    /^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i.test(k)
  );

  // On va traiter bloc par bloc <w:p> pour garder la mise en forme des runs.
  for (const f of targets) {
    let xml = await zip.file(f).async("string");

    // 1) marquage des paragraphes avec tokens (sync), collecte des runs
    pendingParas = [];
    const paraRegex = /<w:p\b[\s\S]*?<\/w:p>/g;
    xml = transformParagraphXml(
      xml,
      "w:t",
      paraRegex,
      /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g,
      mode === "V2" ? "V2" : "V1",
      conservativeSanitize
    );

    // 2) résolution async : appel IA, re-split sur les runs, réinjection XML
    xml = await resolvePending(xml, mode === "V2" ? "V2" : "V1", conservativeSanitize);

    zip.file(f, xml);
  }

  return await zip.generateAsync({ type: "nodebuffer" });
}

export async function processPptxBuffer(buffer, mode = "V1") {
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) =>
    /^(ppt\/slides\/slide\d+\.xml|ppt\/notesSlides\/notesSlide\d+\.xml)$/i.test(k)
  );

  for (const f of targets) {
    let xml = await zip.file(f).async("string");

    pendingParas = [];
    // Un "paragraphe" PPTX est souvent une <a:p> ; on traite pareil
    const paraRegex = /<a:p\b[\s\S]*?<\/a:p>/g;
    xml = transformParagraphXml(
      xml,
      "a:t",
      paraRegex,
      /(<a:t[^>]*>)([\s\S]*?)(<\/a:t>)/g,
      mode === "V2" ? "V2" : "V1",
      conservativeSanitize
    );

    xml = await resolvePending(xml, mode === "V2" ? "V2" : "V1", conservativeSanitize);

    zip.file(f, xml);
  }

  return await zip.generateAsync({ type: "nodebuffer" });
}

