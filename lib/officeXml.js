// lib/officeXml.js
import JSZip from "jszip";

// --- utils ---
function xmlEscape(s = "") {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function replaceTextNodes(xml, openCloseRegex, transformFn) {
  return xml.replace(openCloseRegex, (full, openTag, inner, closeTag) => {
    const newText = transformFn(inner);
    return `${openTag}${xmlEscape(newText)}${closeTag}`;
  });
}
function escRe(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

/**
 * Paires d’artefacts réellement observées dans tes fichiers :
 * "soc ial", "commu nication", "enablin g", "th e", "p otential",
 * "dis connection", "rig hts", "co mmunication"
 * (cf. fichiers fournis). On reste volontairement en liste blanche.  */
const CROSS_RUN_PAIRS = [
  ["soc", "ial"],
  ["commu", "nication"],
  ["enablin", "g"],
  ["th", "e"],
  ["p", "otential"],
  ["dis", "connection"],
  ["rig", "hts"],
  ["co", "mmunication"],
];

/**
 * Supprime UNIQUEMENT l’espace “cassant” au bord des runs
 * pour les paires ci-dessus, sans toucher aux balises ni aux attributs.
 * On gère DOCX (<w:r><w:t>) et PPTX (<a:r><a:t>).
 */
function fixCrossRunArtifacts(xml, { isDocx }) {
  const r = isDocx ? "w:r" : "a:r";
  const t = isDocx ? "w:t" : "a:t";

  for (const [L, R] of CROSS_RUN_PAIRS) {
    // Cas 1 : l’espace est au DÉBUT du <t> suivant :  "...>L</w:t></w:r><w:r...><w:t...>  R"
    const rxLeading = new RegExp(
      `(${escRe(L)}<\/${t}>\\s*<\/${r}>\\s*<${r}[^>]*>\\s*<${t}[^>]*>)\\s+(${escRe(R)})`,
      "g"
    );
    xml = xml.replace(rxLeading, "$1$2");

    // Cas 2 : l’espace est à la FIN du <t> courant :  "...>L  </w:t></w:r><w:r...><w:t...>R"
    const rxTrailing = new RegExp(
      `(${escRe(L)})\\s+(<\/${t}>\\s*<\/${r}>\\s*<${r}[^>]*>\\s*<${t}[^>]*>${escRe(R)})`,
      "g"
    );
    xml = xml.replace(rxTrailing, "$1$2");
  }
  return xml;
}

// --- DOCX : document + headers/footers + notes + commentaires ---
export async function processDocxBuffer(buffer, transformFn) {
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) =>
    /^word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i.test(k)
  );

  for (const f of targets) {
    let xml = await zip.file(f).async("string");

    // 1) Corriger les cassures inter-runs ciblées (sans perdre le style)
    xml = fixCrossRunArtifacts(xml, { isDocx: true });

    // 2) Appliquer la transformation sur CHAQUE nœud texte
    xml = replaceTextNodes(xml, /(<w:t[^>]*>)([\s\S]*?)(<\/w:t>)/g, transformFn);

    zip.file(f, xml);
  }

  return await zip.generateAsync({ type: "nodebuffer" });
}

// --- PPTX : slides + notes ---
export async function processPptxBuffer(buffer, transformFn) {
  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) =>
    /^(ppt\/slides\/slide\d+\.xml|ppt\/notesSlides\/notesSlide\d+\.xml)$/i.test(k)
  );

  for (const f of targets) {
    let xml = await zip.file(f).async("string");

    // 1) Corriger les cassures inter-runs ciblées
    xml = fixCrossRunArtifacts(xml, { isDocx: false });

    // 2) Transformation par nœud texte
    xml = replaceTextNodes(xml, /(<a:t[^>]*>)([\s\S]*?)(<\/a:t>)/g, transformFn);

    zip.file(f, xml);
  }

  return await zip.generateAsync({ type: "nodebuffer" });
}

