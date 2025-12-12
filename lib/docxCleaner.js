// lib/docxCleaner.js
import JSZip from "jszip";

// ============================================================
// Small helpers
// ============================================================

const removeAllCount = (xml, re) => {
  let count = 0;
  const out = String(xml || "").replace(re, () => {
    count++;
    return "";
  });
  return { xml: out, count };
};

const unwrapTagCount = (xml, tag) => {
  let count = 0;
  const open = new RegExp(`<${tag}[^>]*>`, "g");
  const close = new RegExp(`</${tag}>`, "g");
  xml = String(xml || "").replace(open, () => {
    count++;
    return "";
  });
  xml = String(xml || "").replace(close, () => {
    count++;
    return "";
  });
  return { xml, countOpenClose: count };
};

function escapeRegExp(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// ============================================================
// Hidden / white text detection (DOCX) — unchanged logic, safer
// ============================================================

function decodeXmlText(s = "") {
  return String(s)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCharCode(parseInt(d, 10)));
}

function extractRunText(runXml) {
  const texts = [];
  const re = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;
  let m;
  while ((m = re.exec(runXml))) texts.push(decodeXmlText(m[1] || ""));
  return texts.join("");
}

function getRunPr(runXml) {
  const m = String(runXml || "").match(/<w:rPr[\s\S]*?<\/w:rPr>/);
  return m ? m[0] : "";
}

function isVanished(runPrXml) {
  return /<w:(vanish|specVanish)\b[^\/]*\/?>/i.test(runPrXml || "");
}

function isWhiteText(runPrXml) {
  const m =
    String(runPrXml || "").match(/<w:color\b[^>]*w:val="([^"]+)"/i) ||
    String(runPrXml || "").match(/<w:color\b[^>]*val="([^"]+)"/i);
  if (!m) return false;
  const val = String(m[1] || "").replace("#", "").toUpperCase();
  return val === "FFFFFF";
}

function scanHiddenTextInPart(xml, partName) {
  const detections = [];
  const paraRe = /<w:p\b[\s\S]*?<\/w:p>/g;

  let pIndex = 0;
  let pMatch;

  while ((pMatch = paraRe.exec(xml))) {
    const pXml = pMatch[0];
    const runRe = /<w:r\b[\s\S]*?<\/w:r>/g;

    let rIndex = 0;
    let rMatch;

    while ((rMatch = runRe.exec(pXml))) {
      const rXml = rMatch[0];
      const rPr = getRunPr(rXml);

      const vanished = isVanished(rPr);
      const white = isWhiteText(rPr);

      if (!vanished && !white) {
        rIndex++;
        continue;
      }

      const text = extractRunText(rXml).trim();
      if (!text) {
        rIndex++;
        continue;
      }

      detections.push({
        type: white ? "white_text" : "vanished_text",
        reason: white ? "white_color" : "vanish_property",
        content: text,
        location: {
          part: partName,
          paragraph: pIndex,
          run: rIndex,
        },
      });

      rIndex++;
    }

    pIndex++;
  }

  return detections;
}

// ============================================================
// Content_Types + rels cleanup (anti-corruption)
// ============================================================

function removeContentTypeOverrides(ctXml, partNames = []) {
  let xml = String(ctXml || "");
  let removed = 0;

  for (const pn of partNames) {
    const partEsc = escapeRegExp(String(pn).replace(/^\//, ""));
    // PartName="/word/comments.xml" OR PartName="/customXml/item1.xml"
    const re = new RegExp(`<Override[^>]*PartName="\\/${partEsc}"[^>]*/>`, "g");
    const out = removeAllCount(xml, re);
    xml = out.xml;
    removed += out.count;
  }
  return { xml, removed };
}

function listRelsFiles(zip) {
  return Object.keys(zip.files).filter((k) => k.endsWith(".rels"));
}

function removeRelationshipsToTargets(relsXml, targetBasenames = []) {
  // Remove any Relationship where Target="comments.xml" or Target="../customXml/item1.xml" etc.
  // We do "contains" match on basename to be robust across relative paths.
  let xml = String(relsXml || "");
  let removed = 0;

  for (const base of targetBasenames) {
    const b = escapeRegExp(base);
    // capture any <Relationship ... Target="...base..." ... />
    const re = new RegExp(`<Relationship\\b[^>]*Target="[^"]*${b}[^"]*"[^>]*/>`, "g");
    const out = removeAllCount(xml, re);
    xml = out.xml;
    removed += out.count;
  }
  return { xml, removed };
}

function removeRelationshipsByType(relsXml, typeContainsList = []) {
  // Remove Relationship entries with Type containing keywords (image/oleObject/embeddedObject/hyperlink etc. if needed)
  let xml = String(relsXml || "");
  let removed = 0;

  for (const kw of typeContainsList) {
    const k = escapeRegExp(kw);
    const re = new RegExp(`<Relationship\\b[^>]*Type="[^"]*${k}[^"]*"[^>]*/>`, "g");
    const out = removeAllCount(xml, re);
    xml = out.xml;
    removed += out.count;
  }
  return { xml, removed };
}

// ============================================================
// Main
// ============================================================

/**
 * drawPolicy:
 *  - "auto" (default): supprime ink/doodles + vieux VML ; conserve <w:drawing> (logos/images)
 *  - "all": supprime TOUS dessins/images + rels image/ole/embeddings => texte-only visuel
 *  - "none": ne supprime rien côté dessins/images
 */
export async function cleanDOCX(buffer, { drawPolicy = "auto" } = {}) {
  const stats = {
    metaRemoved: 0,
    commentsXmlRemoved: 0,
    commentMarkersRemoved: 0,
    revisionsAccepted: { deletionsRemoved: 0, insertionsUnwrapped: 0 },
    drawingsRemoved: 0,
    vmlRemoved: 0,
    inkRemoved: 0,
    mediaDeleted: 0,

    // NEW
    hiddenTextFound: 0,

    // NEW (stability)
    contentTypesOverridesRemoved: 0,
    relsRemoved: 0,
  };

  const detections = {
    hiddenContent: [],
  };

  const zip = await JSZip.loadAsync(buffer);

  // ----------------------------
  // 0) Define parts to remove
  // ----------------------------
  const removeParts = [
    "docProps/core.xml",
    "docProps/app.xml",
    "docProps/custom.xml",

    "word/comments.xml",
    "word/commentsExtended.xml",

    "customXml/item1.xml",
    "customXml/itemProps1.xml",
  ];

  // Also remove any extra customXml items if present (safer)
  const extraCustomXml = Object.keys(zip.files).filter(
    (k) => /^customXml\/item\d+\.xml$/i.test(k) || /^customXml\/itemProps\d+\.xml$/i.test(k)
  );
  for (const k of extraCustomXml) {
    if (!removeParts.includes(k)) removeParts.push(k);
  }

  // ----------------------------
  // 1) Remove files (physical)
  // ----------------------------
  for (const p of removeParts) {
    if (zip.file(p)) {
      zip.remove(p);
      stats.metaRemoved++;
      if (p.includes("comments")) stats.commentsXmlRemoved++;
    }
  }

  // ----------------------------
  // 2) Clean [Content_Types].xml overrides pointing to removed parts
  // ----------------------------
  if (zip.file("[Content_Types].xml")) {
    const ctRaw = await zip.file("[Content_Types].xml").async("string");
    const { xml: ctClean, removed } = removeContentTypeOverrides(
      ctRaw,
      removeParts.map((p) => "/" + p) // Override uses leading /
    );
    stats.contentTypesOverridesRemoved += removed;
    zip.file("[Content_Types].xml", ctClean);
  }

  // ----------------------------
  // 3) Remove relationships pointing to removed parts (ALL .rels)
  // ----------------------------
  const relsFiles = listRelsFiles(zip);
  const targetBasenames = removeParts.map((p) => p.split("/").pop());

  for (const relPath of relsFiles) {
    const rf = zip.file(relPath);
    if (!rf) continue;

    const raw = await rf.async("string");
    const { xml: cleaned, removed } = removeRelationshipsToTargets(raw, targetBasenames);

    if (removed > 0) {
      stats.relsRemoved += removed;
      zip.file(relPath, cleaned);
    }
  }

  // ----------------------------
  // 4) XML targets to clean (document + headers/footers)
  // ----------------------------
  const targets = [
    "word/document.xml",
    ...Object.keys(zip.files).filter((k) => /word\/(header|footer)\d+\.xml$/i.test(k)),
  ];

  for (const p of targets) {
    const part = zip.file(p);
    if (!part) continue;

    let xml = await part.async("string");

    // Detect hidden/white text BEFORE cleaning
    const hiddenFound = scanHiddenTextInPart(xml, p);
    if (hiddenFound.length) {
      detections.hiddenContent.push(...hiddenFound);
      stats.hiddenTextFound += hiddenFound.length;
    }

    // a) Comment markers
    let out = removeAllCount(xml, /<w:commentRangeStart[^>]*\/>/g);
    xml = out.xml;
    stats.commentMarkersRemoved += out.count;

    out = removeAllCount(xml, /<w:commentRangeEnd[^>]*\/>/g);
    xml = out.xml;
    stats.commentMarkersRemoved += out.count;

    out = removeAllCount(xml, /<w:commentReference[^>]*\/>/g);
    xml = out.xml;
    stats.commentMarkersRemoved += out.count;

    // b) Track changes (accept)
    xml = xml.replace(/<w:del\b[\s\S]*?<\/w:del>/g, () => {
      stats.revisionsAccepted.deletionsRemoved++;
      return "";
    });

    const unwrap = unwrapTagCount(xml, "w:ins");
    xml = unwrap.xml;
    stats.revisionsAccepted.insertionsUnwrapped += unwrap.countOpenClose
      ? Math.ceil(unwrap.countOpenClose / 2)
      : 0;

    // c) Drawings policy
    if (drawPolicy === "all") {
      out = removeAllCount(xml, /<w:drawing[\s\S]*?<\/w:drawing>/g);
      xml = out.xml;
      stats.drawingsRemoved += out.count;

      out = removeAllCount(xml, /<w:pict[\s\S]*?<\/w:pict>/g);
      xml = out.xml;
      stats.vmlRemoved += out.count;

      out = removeAllCount(xml, /<v:shape[\s\S]*?<\/v:shape>/g);
      xml = out.xml;
      stats.vmlRemoved += out.count;

      out = removeAllCount(
        xml,
        /<a:graphicData[^>]*uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/ink"[\s\S]*?<\/a:graphicData>/g
      );
      xml = out.xml;
      stats.inkRemoved += out.count;

      out = removeAllCount(xml, /<a14:ink[\s\S]*?<\/a14:ink>/g);
      xml = out.xml;
      stats.inkRemoved += out.count;
    } else if (drawPolicy === "auto") {
      // remove ink/doodles + old VML; keep <w:drawing>
      out = removeAllCount(
        xml,
        /<a:graphicData[^>]*uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2010\/ink"[\s\S]*?<\/a:graphicData>/g
      );
      xml = out.xml;
      stats.inkRemoved += out.count;

      out = removeAllCount(xml, /<a14:ink[\s\S]*?<\/a14:ink>/g);
      xml = out.xml;
      stats.inkRemoved += out.count;

      out = removeAllCount(xml, /<w:pict[\s\S]*?<\/w:pict>/g);
      xml = out.xml;
      stats.vmlRemoved += out.count;

      out = removeAllCount(xml, /<v:shape[\s\S]*?<\/v:shape>/g);
      xml = out.xml;
      stats.vmlRemoved += out.count;
    }
    // "none": do nothing

    zip.file(p, xml);

    // If drawPolicy=all, also remove relationships for this part (images/embeddings/ole) to avoid broken targets
    if (drawPolicy === "all") {
      const relsPath = p.replace(/^word\//, "word/_rels/") + ".rels"; // word/document.xml -> word/_rels/document.xml.rels
      if (zip.file(relsPath)) {
        const relRaw = await zip.file(relsPath).async("string");
        const { xml: relClean1, removed: r1 } = removeRelationshipsByType(relRaw, [
          "image",
          "oleObject",
          "embeddedObject",
          "package",
        ]);
        if (r1 > 0) {
          stats.relsRemoved += r1;
          zip.file(relsPath, relClean1);
        }
      }
    }
  }

  // ----------------------------
  // 5) If drawPolicy=all: remove /word/media + purge any rels still pointing to /media/
  // ----------------------------
  if (drawPolicy === "all") {
    // delete media files
    for (const k of Object.keys(zip.files)) {
      if (/^word\/media\//i.test(k)) {
        zip.remove(k);
        stats.mediaDeleted++;
      }
    }

    // remove remaining relationships targeting media
    const relsFiles2 = listRelsFiles(zip);
    for (const relPath of relsFiles2) {
      const rf = zip.file(relPath);
      if (!rf) continue;
      const raw = await rf.async("string");
      const { xml: cleaned, removed } = removeAllCount(
        raw,
        /<Relationship\b[^>]*Target="[^"]*word\/media\/[^"]*"[^>]*/g
      );
      // the regex above removes only start; safer to remove full tag:
      let xml = cleaned;
      let extra = 0;
      const fullTagRe = /<Relationship\b[^>]*Target="[^"]*word\/media\/[^"]*"[^>]*\/>/g;
      const out2 = removeAllCount(raw, fullTagRe);
      xml = out2.xml;
      extra = out2.count;

      if (extra > 0) {
        stats.relsRemoved += extra;
        zip.file(relPath, xml);
      }
    }
  }

  // ----------------------------
  // 6) Generate docx (with compression)
  // ----------------------------
  const outBuffer = await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  return { outBuffer, stats, detections };
}

export default { cleanDOCX };
