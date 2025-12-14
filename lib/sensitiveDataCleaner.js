// lib/sensitiveDataCleaner.js
// VERSION 1.0 — Removes sensitive data and hidden content from documents

import JSZip from "jszip";

// ============================================================
// Escape special regex characters
// ============================================================
function escapeRegex(str) {
  return String(str || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

// ============================================================
// DOCX: Remove sensitive data (emails, phones, etc.)
// ============================================================
export async function removeSensitiveDataFromDOCX(buffer, sensitiveItems) {
  if (!sensitiveItems || !sensitiveItems.length) return { outBuffer: buffer, stats: { removed: 0 } };

  const zip = await JSZip.loadAsync(buffer);
  const targets = [
    "word/document.xml",
    ...Object.keys(zip.files).filter((k) => /word\/(header|footer)\d+\.xml$/.test(k)),
    "word/footnotes.xml",
    "word/endnotes.xml",
  ];

  let totalRemoved = 0;
  const examples = [];

  for (const path of targets) {
    const file = zip.file(path);
    if (!file) continue;

    let xml = await file.async("string");
    const originalXml = xml;

    for (const item of sensitiveItems) {
      const value = item.value || item.context || "";
      if (!value || value.length < 3) continue;

      // Try exact match first
      const exactRegex = new RegExp(escapeRegex(value), "gi");
      const beforeLength = xml.length;
      xml = xml.replace(exactRegex, "[REDACTED]");

      if (xml.length !== beforeLength) {
        totalRemoved++;
        if (examples.length < 5) {
          examples.push({ type: item.type, value: value.slice(0, 50), redacted: true });
        }
      }
    }

    if (xml !== originalXml) {
      zip.file(path, xml);
    }
  }

  console.log(`[SENSITIVE DOCX] Removed ${totalRemoved} sensitive items`);
  return {
    outBuffer: await zip.generateAsync({ type: "nodebuffer" }),
    stats: { removed: totalRemoved, examples },
  };
}

// ============================================================
// PPTX: Remove sensitive data
// ============================================================
export async function removeSensitiveDataFromPPTX(buffer, sensitiveItems) {
  if (!sensitiveItems || !sensitiveItems.length) return { outBuffer: buffer, stats: { removed: 0 } };

  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter(
    (k) => /^ppt\/slides\/slide\d+\.xml$/.test(k) || /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(k)
  );

  let totalRemoved = 0;
  const examples = [];

  for (const path of targets) {
    const file = zip.file(path);
    if (!file) continue;

    let xml = await file.async("string");
    const originalXml = xml;

    for (const item of sensitiveItems) {
      const value = item.value || item.context || "";
      if (!value || value.length < 3) continue;

      const exactRegex = new RegExp(escapeRegex(value), "gi");
      const beforeLength = xml.length;
      xml = xml.replace(exactRegex, "[REDACTED]");

      if (xml.length !== beforeLength) {
        totalRemoved++;
        if (examples.length < 5) {
          examples.push({ type: item.type, value: value.slice(0, 50), redacted: true });
        }
      }
    }

    if (xml !== originalXml) {
      zip.file(path, xml);
    }
  }

  console.log(`[SENSITIVE PPTX] Removed ${totalRemoved} sensitive items`);
  return {
    outBuffer: await zip.generateAsync({ type: "nodebuffer" }),
    stats: { removed: totalRemoved, examples },
  };
}

// ============================================================
// XLSX: Remove sensitive data
// ============================================================
export async function removeSensitiveDataFromXLSX(buffer, sensitiveItems) {
  if (!sensitiveItems || !sensitiveItems.length) return { outBuffer: buffer, stats: { removed: 0 } };

  const zip = await JSZip.loadAsync(buffer);
  const targets = [
    "xl/sharedStrings.xml",
    ...Object.keys(zip.files).filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k)),
  ];

  let totalRemoved = 0;
  const examples = [];

  for (const path of targets) {
    const file = zip.file(path);
    if (!file) continue;

    let xml = await file.async("string");
    const originalXml = xml;

    for (const item of sensitiveItems) {
      const value = item.value || item.context || "";
      if (!value || value.length < 3) continue;

      const exactRegex = new RegExp(escapeRegex(value), "gi");
      const beforeLength = xml.length;
      xml = xml.replace(exactRegex, "[REDACTED]");

      if (xml.length !== beforeLength) {
        totalRemoved++;
        if (examples.length < 5) {
          examples.push({ type: item.type, value: value.slice(0, 50), redacted: true });
        }
      }
    }

    if (xml !== originalXml) {
      zip.file(path, xml);
    }
  }

  console.log(`[SENSITIVE XLSX] Removed ${totalRemoved} sensitive items`);
  return {
    outBuffer: await zip.generateAsync({ type: "nodebuffer" }),
    stats: { removed: totalRemoved, examples },
  };
}

// ============================================================
// DOCX: Remove white/hidden text
// ============================================================
export async function removeHiddenContentFromDOCX(buffer, hiddenItems) {
  if (!hiddenItems || !hiddenItems.length) return { outBuffer: buffer, stats: { removed: 0 } };

  const zip = await JSZip.loadAsync(buffer);
  const file = zip.file("word/document.xml");
  if (!file) return { outBuffer: buffer, stats: { removed: 0 } };

  let xml = await file.async("string");
  let totalRemoved = 0;

  for (const item of hiddenItems) {
    if (item.type === "white_text" || item.type === "invisible_text") {
      // Find and remove white text (color FFFFFF or similar)
      // Pattern: <w:r>...<w:color w:val="FFFFFF"/>...<w:t>TEXT</w:t>...</w:r>
      const text = item.content || item.text || "";
      if (text && text.length > 2) {
        const escapedText = escapeRegex(text);
        // Remove the text content
        const regex = new RegExp(`>\\s*${escapedText}\\s*<`, "gi");
        const before = xml.length;
        xml = xml.replace(regex, "><");
        if (xml.length !== before) {
          totalRemoved++;
        }
      }
    }
  }

  if (totalRemoved > 0) {
    zip.file("word/document.xml", xml);
  }

  console.log(`[HIDDEN DOCX] Removed ${totalRemoved} hidden content items`);
  return {
    outBuffer: await zip.generateAsync({ type: "nodebuffer" }),
    stats: { removed: totalRemoved },
  };
}

// ============================================================
// PPTX: Remove white/hidden text and covering shapes
// ============================================================
export async function removeHiddenContentFromPPTX(buffer, hiddenItems) {
  if (!hiddenItems || !hiddenItems.length) return { outBuffer: buffer, stats: { removed: 0 } };

  const zip = await JSZip.loadAsync(buffer);
  const targets = Object.keys(zip.files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k));

  let totalRemoved = 0;

  for (const path of targets) {
    const file = zip.file(path);
    if (!file) continue;

    let xml = await file.async("string");
    const originalXml = xml;

    for (const item of hiddenItems) {
      if (item.type === "white_text" || item.type === "invisible_text") {
        const text = item.content || item.text || "";
        if (text && text.length > 2) {
          const escapedText = escapeRegex(text);
          const regex = new RegExp(`<a:t>${escapedText}</a:t>`, "gi");
          const before = xml.length;
          xml = xml.replace(regex, "<a:t></a:t>");
          if (xml.length !== before) {
            totalRemoved++;
          }
        }
      }
    }

    if (xml !== originalXml) {
      zip.file(path, xml);
    }
  }

  console.log(`[HIDDEN PPTX] Removed ${totalRemoved} hidden content items`);
  return {
    outBuffer: await zip.generateAsync({ type: "nodebuffer" }),
    stats: { removed: totalRemoved },
  };
}

// ============================================================
// PPTX: Remove visual objects (covering shapes)
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
        // Try to find and remove the shape by its name or description
        const shapeName = obj.name || obj.description || "";
        if (shapeName) {
          // Remove <p:sp> elements that match
          const shapeRegex = new RegExp(
            `<p:sp[^>]*>(?:(?!<p:sp).)*?<p:nvSpPr>(?:(?!<p:sp).)*?${escapeRegex(shapeName)}(?:(?!<p:sp).)*?</p:sp>`,
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
    outBuffer: await zip.generateAsync({ type: "nodebuffer" }),
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
