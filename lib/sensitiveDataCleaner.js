// lib/sensitiveDataCleaner.js
// VERSION 1.1 — Sensitive data redaction aligned with removeSensitiveData(buffer, ext, items)

import JSZip from "jszip";

// ============================================================
// Escape special regex characters
// ============================================================
function escapeRegex(str) {
  return String(str || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// ============================================================
// Internal: collect XML target files for a given ext + zip
// ============================================================
function collectXmlTargets(zip, ext) {
  const targets = [];

  if (ext === "docx") {
    // Same spirit as your snippet, but safer (auto-detect all headers/footers)
    targets.push("word/document.xml");
    targets.push(
      ...Object.keys(zip.files).filter((k) => /word\/header\d+\.xml$/.test(k))
    );
    targets.push(
      ...Object.keys(zip.files).filter((k) => /word\/footer\d+\.xml$/.test(k))
    );
    targets.push("word/footnotes.xml");
    targets.push("word/endnotes.xml");
  }

  if (ext === "pptx") {
    const slideFiles = Object.keys(zip.files).filter((k) =>
      /^ppt\/slides\/slide\d+\.xml$/.test(k)
    );
    const noteFiles = Object.keys(zip.files).filter((k) =>
      /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(k)
    );
    targets.push(...slideFiles, ...noteFiles);
  }

  if (ext === "xlsx") {
    targets.push("xl/sharedStrings.xml");
    targets.push(
      ...Object.keys(zip.files).filter((k) =>
        /^xl\/worksheets\/sheet\d+\.xml$/.test(k)
      )
    );
  }

  // Remove duplicates + keep only existing files
  return [...new Set(targets)].filter((p) => zip.file(p));
}

// ============================================================
// Internal: redact values in XML files
// returns { outBuffer, stats: { removed, filesTouched, examples } }
// ============================================================
async function redactInZipXml(zip, targets, items, label = "SENSITIVE") {
  let totalReplacements = 0;
  let filesTouched = 0;
  const examples = [];

  for (const xmlPath of targets) {
    const file = zip.file(xmlPath);
    if (!file) continue;

    let content = await file.async("string");
    const original = content;

    for (const item of items || []) {
      const value = item?.value;
      if (!value || String(value).trim().length < 1) continue;

      const escaped = escapeRegex(value);
      const regex = new RegExp(escaped, "g"); // EXACT match (same as your snippet)

      // Count occurrences before replace
      const matches = content.match(regex);
      const count = matches ? matches.length : 0;

      if (count > 0) {
        content = content.replace(regex, "[REDACTED]");
        totalReplacements += count;

        if (examples.length < 5) {
          examples.push({
            type: item?.type || "unknown",
            value: String(value).slice(0, 80),
            file: xmlPath,
            count,
          });
        }

        console.log(
          `[${label}] Replaced "${String(value).slice(0, 120)}" -> [REDACTED] in ${xmlPath} (${count}x)`
        );
      }
    }

    if (content !== original) {
      zip.file(xmlPath, content);
      filesTouched += 1;
    }
  }

  return { totalReplacements, filesTouched, examples };
}

// ============================================================
// DOCX: Remove sensitive data (exact replacement of item.value)
// ============================================================
export async function removeSensitiveDataFromDOCX(buffer, sensitiveItems) {
  if (!sensitiveItems || sensitiveItems.length === 0) {
    return { outBuffer: buffer, stats: { removed: 0, filesTouched: 0, examples: [] } };
  }

  const zip = await JSZip.loadAsync(buffer);
  const targets = collectXmlTargets(zip, "docx");

  const { totalReplacements, filesTouched, examples } = await redactInZipXml(
    zip,
    targets,
    sensitiveItems,
    "SENSITIVE DOCX"
  );

  console.log(`[SENSITIVE DOCX] Total replacements: ${totalReplacements} across ${filesTouched} file(s)`);

  return {
    outBuffer: await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
    }),
    stats: { removed: totalReplacements, filesTouched, examples },
  };
}

// ============================================================
// PPTX: Remove sensitive data (exact replacement of item.value)
// ============================================================
export async function removeSensitiveDataFromPPTX(buffer, sensitiveItems) {
  if (!sensitiveItems || sensitiveItems.length === 0) {
    return { outBuffer: buffer, stats: { removed: 0, filesTouched: 0, examples: [] } };
  }

  const zip = await JSZip.loadAsync(buffer);
  const targets = collectXmlTargets(zip, "pptx");

  const { totalReplacements, filesTouched, examples } = await redactInZipXml(
    zip,
    targets,
    sensitiveItems,
    "SENSITIVE PPTX"
  );

  console.log(`[SENSITIVE PPTX] Total replacements: ${totalReplacements} across ${filesTouched} file(s)`);

  return {
    outBuffer: await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
    }),
    stats: { removed: totalReplacements, filesTouched, examples },
  };
}

// ============================================================
// XLSX: Remove sensitive data (exact replacement of item.value)
// ============================================================
export async function removeSensitiveDataFromXLSX(buffer, sensitiveItems) {
  if (!sensitiveItems || sensitiveItems.length === 0) {
    return { outBuffer: buffer, stats: { removed: 0, filesTouched: 0, examples: [] } };
  }

  const zip = await JSZip.loadAsync(buffer);
  const targets = collectXmlTargets(zip, "xlsx");

  const { totalReplacements, filesTouched, examples } = await redactInZipXml(
    zip,
    targets,
    sensitiveItems,
    "SENSITIVE XLSX"
  );

  console.log(`[SENSITIVE XLSX] Total replacements: ${totalReplacements} across ${filesTouched} file(s)`);

  return {
    outBuffer: await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
    }),
    stats: { removed: totalReplacements, filesTouched, examples },
  };
}

// ============================================================
// DOCX: Remove white/hidden text (basic)
// ============================================================
export async function removeHiddenContentFromDOCX(buffer, hiddenItems) {
  if (!hiddenItems || !hiddenItems.length) return { outBuffer: buffer, stats: { removed: 0 } };

  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("word/document.xml");
  if (!file) return { outBuffer: buffer, stats: { removed: 0 } };

  let xml = await file.async("string");
  const originalXml = xml;

  let totalRemoved = 0;

  for (const item of hiddenItems) {
    if (item.type === "white_text" || item.type === "invisible_text") {
      const text = item.content || item.text || "";
      if (text && text.length > 2) {
        const escapedText = escapeRegex(text);
        const regex = new RegExp(`>\\s*${escapedText}\\s*<`, "gi");
        const matches = xml.match(regex);
        const count = matches ? matches.length : 0;

        if (count > 0) {
          xml = xml.replace(regex, "><");
          totalRemoved += count;
        }
      }
    }
  }

  if (xml !== originalXml) {
    zip.file("word/document.xml", xml);
  }

  console.log(`[HIDDEN DOCX] Removed ${totalRemoved} hidden content occurrences`);
  return {
    outBuffer: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
    stats: { removed: totalRemoved },
  };
}

// ============================================================
// PPTX: Remove white/hidden text (basic)
// ============================================================
export async function removeHiddenContentFromPPTX(buffer, hiddenItems) {
  if (!hiddenItems || !hiddenItems.length) return { outBuffer: buffer, stats: { removed: 0 } };

  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k));

  let totalRemoved = 0;

  for (const p of targets) {
    const file = zip.file(p);
    if (!file) continue;

    let xml = await file.async("string");
    const original = xml;

    for (const item of hiddenItems) {
      if (item.type === "white_text" || item.type === "invisible_text") {
        const text = item.content || item.text || "";
        if (text && text.length > 2) {
          const escapedText = escapeRegex(text);
          const regex = new RegExp(`<a:t>${escapedText}</a:t>`, "gi");

          const matches = xml.match(regex);
          const count = matches ? matches.length : 0;

          if (count > 0) {
            xml = xml.replace(regex, "<a:t></a:t>");
            totalRemoved += count;
          }
        }
      }
    }

    if (xml !== original) {
      zip.file(p, xml);
    }
  }

  console.log(`[HIDDEN PPTX] Removed ${totalRemoved} hidden content occurrences`);
  return {
    outBuffer: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
    stats: { removed: totalRemoved },
  };
}

// ============================================================
// PPTX: Remove visual objects (covering shapes) — unchanged
// ============================================================
export async function removeVisualObjectsFromPPTX(buffer, visualObjects) {
  if (!visualObjects || !visualObjects.length) return { outBuffer: buffer, stats: { removed: 0 } };

  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k));

  let totalRemoved = 0;

  for (const path of targets) {
    const file = zip.file(path);
    if (!file) continue;

    let xml = await file.async("string");
    const originalLength = xml.length;

    for (const obj of visualObjects) {
      if (obj.type === "covering_shape" || obj.type === "shape_covering_text") {
        const shapeName = obj.name || obj.description || "";
        if (shapeName) {
          const shapeRegex = new RegExp(
            `<p:sp[^>]*>(?:(?!<p:sp).)*?<p:nvSpPr>(?:(?!<p:sp).)*?${escapeRegex(
              shapeName
            )}(?:(?!<p:sp).)*?</p:sp>`,
            "gis"
          );
          xml = xml.replace(shapeRegex, "");
        }
      }
    }

    if (xml.length !== originalLength) {
      totalRemoved++;
      zip.file(path, xml);
    }
  }

  console.log(`[VISUAL PPTX] Removed ${totalRemoved} visual objects`);
  return {
    outBuffer: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
    stats: { removed: totalRemoved },
  };
}

export default {
  removeSensitiveDataFromDOCX,
  removeSensitiveDataFromPPTX,
  removeSensitiveDataFromXLSX,
  removeHiddenContentFromDOCX,
  removeHiddenContentFromPPTX,
  removeVisualObjectsFromPPTX,
};
