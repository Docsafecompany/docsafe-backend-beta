// lib/xlsxCleaner.js
// Nettoyage des fichiers Excel

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
    metadataRemoved: [],
    commentsRemoved: 0,
    hiddenSheetsRemoved: 0,
    embeddingsRemoved: 0,
    macrosRemoved: false,
    formulasConverted: 0
  };
  
  const zip = await JSZip.loadAsync(buffer);
  
  // 1. Clean metadata
  if (removeMetadata) {
    // app.xml
    const appXml = zip.file("docProps/app.xml");
    if (appXml) {
      const content = await appXml.async("text");
      const parsed = await parseStringPromise(content);
      if (parsed?.Properties) {
        if (parsed.Properties.Company) {
          stats.metadataRemoved.push('company');
          delete parsed.Properties.Company;
        }
        if (parsed.Properties.Manager) {
          stats.metadataRemoved.push('manager');
          delete parsed.Properties.Manager;
        }
        if (parsed.Properties.Application) {
          stats.metadataRemoved.push('application');
          delete parsed.Properties.Application;
        }
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
        const fieldsToClean = ['dc:creator', 'dc:title', 'dc:subject', 'cp:keywords', 'cp:lastModifiedBy'];
        fieldsToClean.forEach(field => {
          if (props[field]) {
            stats.metadataRemoved.push(field.split(':')[1]);
            props[field] = [''];
          }
        });
        props['cp:revision'] = ['1'];
        const builder = new Builder();
        zip.file("docProps/core.xml", builder.buildObject(parsed));
      }
    }
  }
  
  // 2. Remove comments
  if (removeComments) {
    const commentFiles = Object.keys(zip.files).filter(name =>
      name.match(/xl\/comments\d*\.xml/)
    );
    
    for (const commentFile of commentFiles) {
      const file = zip.file(commentFile);
      if (file) {
        const content = await file.async("text");
        const countMatch = content.match(/<comment /g);
        stats.commentsRemoved += countMatch ? countMatch.length : 0;
      }
      zip.remove(commentFile);
    }
    
    // Also remove comment relationships
    const relsFiles = Object.keys(zip.files).filter(name =>
      name.includes('worksheets/_rels/') && name.endsWith('.rels')
    );
    
    for (const relsFile of relsFiles) {
      const file = zip.file(relsFile);
      if (file) {
        let content = await file.async("text");
        content = content.replace(/<Relationship[^>]*Target="[^"]*comments[^"]*"[^>]*\/>/g, '');
        zip.file(relsFile, content);
      }
    }
  }
  
  // 3. Remove or unhide hidden sheets
  if (removeHiddenSheets) {
    const workbookXml = zip.file("xl/workbook.xml");
    if (workbookXml) {
      const content = await workbookXml.async("text");
      const parsed = await parseStringPromise(content);
      
      if (parsed?.workbook?.sheets?.[0]?.sheet) {
        const sheets = parsed.workbook.sheets[0].sheet;
        const sheetsToRemove = [];
        
        sheets.forEach((sheet, index) => {
          const state = sheet.$?.state;
          if (state === 'hidden' || state === 'veryHidden') {
            sheetsToRemove.push({ index, name: sheet.$?.name });
            stats.hiddenSheetsRemoved++;
          }
        });
        
        // Remove hidden sheets from workbook.xml
        parsed.workbook.sheets[0].sheet = sheets.filter((sheet) => {
          const state = sheet.$?.state;
          return state !== 'hidden' && state !== 'veryHidden';
        });
        
        const builder = new Builder();
        zip.file("xl/workbook.xml", builder.buildObject(parsed));
        
        // Remove actual sheet files (simplified - would need proper sheet ID mapping)
        sheetsToRemove.forEach(({ name }) => {
          console.log(`[XLSX] Removing hidden sheet: ${name}`);
        });
      }
    }
  }
  
  // 4. Remove embedded objects
  if (removeEmbeddings) {
    const embedFiles = Object.keys(zip.files).filter(name =>
      name.startsWith("xl/embeddings/")
    );
    
    stats.embeddingsRemoved = embedFiles.length;
    embedFiles.forEach(name => zip.remove(name));
  }
  
  // 5. Remove macros (VBA)
  if (removeMacros) {
    const macroFiles = Object.keys(zip.files).filter(name =>
      name.includes("vbaProject") || name.endsWith(".bin")
    );
    
    if (macroFiles.length > 0) {
      stats.macrosRemoved = true;
      macroFiles.forEach(name => zip.remove(name));
    }
  }
  
  // Generate output
  const outBuffer = await zip.generateAsync({ type: "nodebuffer" });
  
  return { outBuffer, stats };
}

export default { cleanXLSX };
