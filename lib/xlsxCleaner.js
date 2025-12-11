// lib/xlsxCleaner.js
// Nettoyage des fichiers Excel (XLSX) - version "product ready"

import JSZip from "jszip";
import xml2js from "xml2js";

const parseStringPromise = xml2js.parseStringPromise;
const Builder = xml2js.Builder;

/**
 * Nettoie un fichier Excel selon les options
 * @param {Buffer} buffer - Buffer du fichier XLSX
 * @param {Object} options - Options de nettoyage
 * @returns {Promise<{outBuffer: Buffer, stats: Object}>}
 */
export async function cleanXLSX(buffer, options = {}) {
  const {
    removeMetadata = true,
    removeComments = true,
    removeHiddenSheets = true,
    removeEmbeddings = true,
    removeMacros = true,
    removeFormulas = false, // Convertir formules en valeurs
  } = options;

  const stats = {
    // aligné avec calculateAfterScore
    metaRemoved: 0,
    metadataFields: [],

    commentsXmlRemoved: 0,

    hiddenSheetsRemoved: 0,
    hiddenRemoved: 0,

    embeddingsRemoved: 0,

    macrosRemoved: 0, // nombre de fichiers macros supprimés

    formulasConverted: 0,
  };

  const zip = await JSZip.loadAsync(buffer);

  // =========================================================
  // 1. Clean metadata (docProps/app.xml + docProps/core.xml)
  // =========================================================
  if (removeMetadata) {
    // app.xml
    const appXml = zip.file("docProps/app.xml");
    if (appXml) {
      const content = await appXml.async("text");
      const parsed = await parseStringPromise(content);
      if (parsed?.Properties) {
        const fields = ["Company", "Manager", "Application"];
        fields.forEach((field) => {
          if (parsed.Properties[field]) {
            stats.metaRemoved++;
            stats.metadataFields.push(field.toLowerCase());
            delete parsed.Properties[field];
          }
        });

        const builder = new Builder();
        zip.file("docProps/app.xml", builder.buildObject(parsed));
      }
    }

    // core.xml
    const coreXml = zip.file("docProps/core.xml");
    if (coreXml) {
      const content = await coreXml.async("text");
      const parsed = await parseStringPromise(content);
      if (parsed?.["cp:coreProperties"]) {
        const props = parsed["cp:coreProperties"];
        const fieldsToClean = [
          "dc:creator",
          "dc:title",
          "dc:subject",
          "cp:keywords",
          "cp:lastModifiedBy",
        ];

        fieldsToClean.forEach((field) => {
          if (props[field]) {
            stats.metaRemoved++;
            stats.metadataFields.push(field.split(":")[1]);
            props[field] = [""];
          }
        });

        // Reset révision
        props["cp:revision"] = ["1"];

        const builder = new Builder();
        zip.file("docProps/core.xml", builder.buildObject(parsed));
      }
    }
  }

  // =========================================================
  // 2. Remove comments (xl/comments*.xml + relations)
  // =========================================================
  if (removeComments) {
    const commentFiles = Object.keys(zip.files).filter((name) =>
      name.match(/xl\/comments\d*\.xml/)
    );

    for (const commentFile of commentFiles) {
      const file = zip.file(commentFile);
      if (file) {
        const content = await file.async("text");
        const countMatch = content.match(/<comment /g);
        stats.commentsXmlRemoved += countMatch ? countMatch.length : 0;
      }
      zip.remove(commentFile);
    }

    // Remove comment relationships from worksheets rels
    const relsFiles = Object.keys(zip.files).filter(
      (name) =>
        name.includes("worksheets/_rels/") && name.endsWith(".rels")
    );

    for (const relsFile of relsFiles) {
      const file = zip.file(relsFile);
      if (file) {
        let content = await file.async("text");
        content = content.replace(
          /<Relationship[^>]*Target="[^"]*comments[^"]*"[^>]*\/>/g,
          ""
        );
        zip.file(relsFile, content);
      }
    }
  }

  // =========================================================
  // 3. Remove hidden sheets (workbook.xml + workbook.xml.rels + sheetX.xml)
  // =========================================================
  if (removeHiddenSheets) {
    const workbookXml = zip.file("xl/workbook.xml");
    const relsXml = zip.file("xl/_rels/workbook.xml.rels");

    if (workbookXml) {
      const workbookContent = await workbookXml.async("text");
      const workbookParsed = await parseStringPromise(workbookContent);

      let relsParsed = null;
      if (relsXml) {
        const relsContent = await relsXml.async("text");
        relsParsed = await parseStringPromise(relsContent);
      }

      if (workbookParsed?.workbook?.sheets?.[0]?.sheet) {
        const sheets = workbookParsed.workbook.sheets[0].sheet;
        const sheetIdsToRemove = new Set();
        const sheetTargetsToRemove = [];

        sheets.forEach((sheet) => {
          const state = sheet.$?.state;
          const rId = sheet.$?.["r:id"];
          if (state === "hidden" || state === "veryHidden") {
            stats.hiddenSheetsRemoved++;
            stats.hiddenRemoved++;
            if (rId) sheetIdsToRemove.add(rId);
          }
        });

        // Filter visible sheets in workbook.xml
        workbookParsed.workbook.sheets[0].sheet = sheets.filter(
          (sheet) =>
            sheet.$?.state !== "hidden" && sheet.$?.state !== "veryHidden"
        );

        const builder = new Builder();
        zip.file("xl/workbook.xml", builder.buildObject(workbookParsed));

        // Remove sheet relationships + collect targets
        if (relsParsed?.Relationships?.Relationship && sheetIdsToRemove.size) {
          relsParsed.Relationships.Relationship =
            relsParsed.Relationships.Relationship.filter((rel) => {
              const id = rel.$.Id;
              if (sheetIdsToRemove.has(id)) {
                const target = rel.$.Target; // e.g. "worksheets/sheet3.xml"
                if (target) {
                  let path = target.replace(/^\/?/, "");
                  if (!path.startsWith("xl/")) path = "xl/" + path;
                  sheetTargetsToRemove.push(path);
                }
                return false; // remove this relationship
              }
              return true;
            });

          const relsBuilder = new Builder();
          zip.file(
            "xl/_rels/workbook.xml.rels",
            relsBuilder.buildObject(relsParsed)
          );
        }

        // Remove the actual sheet XML files
        sheetTargetsToRemove.forEach((sheetPath) => {
          if (zip.file(sheetPath)) {
            console.log(`[XLSX] Removing hidden sheet file: ${sheetPath}`);
            zip.remove(sheetPath);
          }
        });
      }
    }
  }

  // =========================================================
  // 4. Remove embedded objects (xl/embeddings/*)
  // =========================================================
  if (removeEmbeddings) {
    const embedFiles = Object.keys(zip.files).filter((name) =>
      name.startsWith("xl/embeddings/")
    );

    stats.embeddingsRemoved = embedFiles.length;
    embedFiles.forEach((name) => zip.remove(name));
  }

  // =========================================================
  // 5. Remove macros (VBA) - vbaProject*.bin etc.
  // =========================================================
  if (removeMacros) {
    const macroFiles = Object.keys(zip.files).filter(
      (name) =>
        name.includes("vbaProject") || name.endsWith(".bin")
    );

    if (macroFiles.length > 0) {
      stats.macrosRemoved = macroFiles.length;
      macroFiles.forEach((name) => zip.remove(name));
    }
  }

  // =========================================================
  // 6. Convert formulas to values (remove <f>, keep <v>)
  // =========================================================
  if (removeFormulas) {
    const sheetFiles = Object.keys(zip.files).filter(
      (name) =>
        name.startsWith("xl/worksheets/") && name.endsWith(".xml")
    );

    for (const sheetFile of sheetFiles) {
      const file = zip.file(sheetFile);
      if (!file) continue;

      const xml = await file.async("text");
      let parsed;
      try {
        parsed = await parseStringPromise(xml);
      } catch (e) {
        console.warn(
          `[XLSX] Failed to parse worksheet ${sheetFile} for formulas conversion`,
          e
        );
        continue;
      }

      const sheetData = parsed?.worksheet?.sheetData?.[0]?.row || [];
      sheetData.forEach((row) => {
        const cells = row.c || [];
        cells.forEach((cell) => {
          if (cell.f) {
            // il y a une formule
            stats.formulasConverted++;
            delete cell.f; // on enlève la formule, on garde la valeur <v>
          }
        });
      });

      const builder = new Builder();
      zip.file(sheetFile, builder.buildObject(parsed));
    }
  }

  // =========================================================
  // Génération du fichier final
  // =========================================================
  const outBuffer = await zip.generateAsync({ type: "nodebuffer" });

  return { outBuffer, stats };
}

export default { cleanXLSX };
