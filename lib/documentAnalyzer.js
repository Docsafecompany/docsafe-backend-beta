// lib/documentAnalyzer.js
// Analyse les documents AVANT nettoyage pour le preview interactif

import JSZip from "jszip";
import xml2js from "xml2js";
import { PDFDocument } from "pdf-lib";

const parseStringPromise = xml2js.parseStringPromise;

/**
 * Analyse un document et retourne les détections pour le preview
 * @param {Buffer} fileBuffer - Le buffer du fichier
 * @param {string} fileType - Type MIME du fichier
 * @returns {Promise<Object>} - Détections structurées
 */
export async function analyzeDocument(fileBuffer, fileType) {
  const ext = getExtFromMime(fileType);
  
  switch (ext) {
    case 'docx':
      return await analyzeDOCX(fileBuffer);
    case 'pptx':
      return await analyzePPTX(fileBuffer);
    case 'xlsx':
      return await analyzeXLSX(fileBuffer);
    case 'pdf':
      return await analyzePDF(fileBuffer);
    default:
      throw new Error(`Unsupported file type: ${ext}`);
  }
}

function getExtFromMime(mime) {
  const map = {
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
    'application/pdf': 'pdf'
  };
  return map[mime] || null;
}

// ============= DOCX ANALYSIS =============
async function analyzeDOCX(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  
  return {
    metadata: await analyzeOfficeMetadata(zip, 'word'),
    comments: await analyzeDOCXComments(zip),
    trackChanges: await analyzeDOCXTrackChanges(zip),
    hiddenContent: await analyzeDOCXHiddenContent(zip),
    embeddedObjects: await analyzeEmbeddedObjects(zip, 'word'),
    macros: await analyzeMacros(zip, 'word'),
    sensitiveData: await analyzeSensitiveData(zip, 'docx'),
    spellingErrors: [], // Sera rempli par l'IA si demandé
    brokenLinks: [],
    businessInconsistencies: [],
    complianceRisks: []
  };
}

// ============= PPTX ANALYSIS =============
async function analyzePPTX(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  
  return {
    metadata: await analyzeOfficeMetadata(zip, 'ppt'),
    comments: await analyzePPTXComments(zip),
    trackChanges: [],
    hiddenContent: await analyzePPTXHiddenContent(zip),
    embeddedObjects: await analyzeEmbeddedObjects(zip, 'ppt'),
    macros: await analyzeMacros(zip, 'ppt'),
    sensitiveData: await analyzeSensitiveData(zip, 'pptx'),
    spellingErrors: [],
    brokenLinks: [],
    businessInconsistencies: [],
    complianceRisks: []
  };
}

// ============= XLSX ANALYSIS =============
async function analyzeXLSX(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  
  return {
    metadata: await analyzeOfficeMetadata(zip, 'xl'),
    comments: await analyzeExcelComments(zip),
    trackChanges: [],
    hiddenContent: [],
    hiddenSheets: await analyzeExcelHiddenSheets(zip),
    hiddenColumns: [],
    sensitiveFormulas: await analyzeExcelSensitiveFormulas(zip),
    embeddedObjects: await analyzeEmbeddedObjects(zip, 'xl'),
    macros: await analyzeMacros(zip, 'xl'),
    sensitiveData: await analyzeSensitiveData(zip, 'xlsx'),
    spellingErrors: [],
    brokenLinks: [],
    businessInconsistencies: [],
    complianceRisks: []
  };
}

// ============= PDF ANALYSIS =============
async function analyzePDF(buffer) {
  const pdfDoc = await PDFDocument.load(buffer);
  
  const metadata = [];
  const author = pdfDoc.getAuthor();
  const title = pdfDoc.getTitle();
  const subject = pdfDoc.getSubject();
  const creator = pdfDoc.getCreator();
  const producer = pdfDoc.getProducer();
  const keywords = pdfDoc.getKeywords();
  
  if (author) metadata.push({ category: 'author', value: author, canRemove: true });
  if (title) metadata.push({ category: 'title', value: title, canRemove: true });
  if (subject) metadata.push({ category: 'subject', value: subject, canRemove: true });
  if (creator) metadata.push({ category: 'creator', value: creator, canRemove: true });
  if (producer) metadata.push({ category: 'producer', value: producer, canRemove: true });
  if (keywords) metadata.push({ category: 'keywords', value: keywords, canRemove: true });
  
  return {
    metadata,
    comments: [],
    trackChanges: [],
    hiddenContent: [],
    embeddedObjects: [],
    macros: [],
    sensitiveData: [],
    spellingErrors: [],
    brokenLinks: [],
    businessInconsistencies: [],
    complianceRisks: []
  };
}

// ============= HELPERS =============

async function analyzeOfficeMetadata(zip, prefix) {
  const metadata = [];
  
  // app.xml
  const appXml = zip.file("docProps/app.xml");
  if (appXml) {
    const content = await appXml.async("text");
    const parsed = await parseStringPromise(content);
    if (parsed?.Properties) {
      if (parsed.Properties.Company?.[0]) 
        metadata.push({ category: 'company', value: parsed.Properties.Company[0], canRemove: true });
      if (parsed.Properties.Manager?.[0]) 
        metadata.push({ category: 'manager', value: parsed.Properties.Manager[0], canRemove: true });
      if (parsed.Properties.Application?.[0]) 
        metadata.push({ category: 'application', value: parsed.Properties.Application[0], canRemove: true });
    }
  }
  
  // core.xml
  const coreXml = zip.file("docProps/core.xml");
  if (coreXml) {
    const content = await coreXml.async("text");
    const parsed = await parseStringPromise(content);
    if (parsed?.["cp:coreProperties"]) {
      const props = parsed["cp:coreProperties"];
      if (props["dc:creator"]?.[0]) 
        metadata.push({ category: 'author', value: props["dc:creator"][0], canRemove: true });
      if (props["dc:title"]?.[0]) 
        metadata.push({ category: 'title', value: props["dc:title"][0], canRemove: true });
      if (props["dc:subject"]?.[0]) 
        metadata.push({ category: 'subject', value: props["dc:subject"][0], canRemove: true });
      if (props["cp:keywords"]?.[0]) 
        metadata.push({ category: 'keywords', value: props["cp:keywords"][0], canRemove: true });
      if (props["cp:lastModifiedBy"]?.[0]) 
        metadata.push({ category: 'lastModifiedBy', value: props["cp:lastModifiedBy"][0], canRemove: true });
    }
  }
  
  return metadata;
}

async function analyzeDOCXComments(zip) {
  const comments = [];
  const commentsXml = zip.file("word/comments.xml");
  
  if (commentsXml) {
    const content = await commentsXml.async("text");
    const parsed = await parseStringPromise(content);
    
    if (parsed?.["w:comments"]?.["w:comment"]) {
      const commentsList = parsed["w:comments"]["w:comment"];
      commentsList.forEach((comment, index) => {
        const author = comment.$?.["w:author"] || "Unknown";
        const text = extractTextFromXML(JSON.stringify(comment));
        
        comments.push({
          id: `comment_${index}`,
          author,
          text: text.slice(0, 200),
          location: `Comment ${index + 1}`,
          severity: determineSeverity(text)
        });
      });
    }
  }
  
  return comments;
}

async function analyzeDOCXTrackChanges(zip) {
  const trackChanges = [];
  const documentXml = zip.file("word/document.xml");
  
  if (documentXml) {
    const content = await documentXml.async("text");
    
    // Detect insertions
    const insertMatches = content.matchAll(/<w:ins[^>]*w:author="([^"]*)"[^>]*>([\s\S]*?)<\/w:ins>/g);
    for (const match of insertMatches) {
      trackChanges.push({
        id: `ins_${trackChanges.length}`,
        type: 'insertion',
        author: match[1],
        text: extractTextFromXML(match[2]).slice(0, 100),
        severity: 'medium'
      });
    }
    
    // Detect deletions
    const deleteMatches = content.matchAll(/<w:del[^>]*w:author="([^"]*)"[^>]*>([\s\S]*?)<\/w:del>/g);
    for (const match of deleteMatches) {
      trackChanges.push({
        id: `del_${trackChanges.length}`,
        type: 'deletion',
        author: match[1],
        text: extractTextFromXML(match[2]).slice(0, 100),
        severity: 'medium'
      });
    }
  }
  
  return trackChanges;
}

async function analyzeDOCXHiddenContent(zip) {
  const hidden = [];
  const documentXml = zip.file("word/document.xml");
  
  if (documentXml) {
    const content = await documentXml.async("text");
    
    // Hidden text (w:vanish)
    const hiddenMatches = content.matchAll(/<w:vanish[^>]*\/>/g);
    let count = 0;
    for (const _ of hiddenMatches) count++;
    
    if (count > 0) {
      hidden.push({
        id: 'hidden_text',
        type: 'hidden_text',
        description: `${count} hidden text element(s) detected`,
        severity: 'high'
      });
    }
    
    // White text on white background
    const whiteTextMatches = content.matchAll(/<w:color[^>]*w:val="FFFFFF"[^>]*\/>/g);
    let whiteCount = 0;
    for (const _ of whiteTextMatches) whiteCount++;
    
    if (whiteCount > 0) {
      hidden.push({
        id: 'white_text',
        type: 'white_text',
        description: `${whiteCount} white text element(s) detected (potentially hidden)`,
        severity: 'high'
      });
    }
  }
  
  return hidden;
}

async function analyzePPTXComments(zip) {
  const comments = [];
  const commentFiles = Object.keys(zip.files).filter(name => 
    name.startsWith("ppt/comments/comment") && name.endsWith(".xml")
  );
  
  for (const commentFile of commentFiles) {
    const file = zip.file(commentFile);
    if (file) {
      const content = await file.async("text");
      const parsed = await parseStringPromise(content);
      
      if (parsed?.["p:cmLst"]?.["p:cm"]) {
        parsed["p:cmLst"]["p:cm"].forEach((cm, index) => {
          comments.push({
            id: `ppt_comment_${comments.length}`,
            author: cm.$?.authorId || "Unknown",
            text: cm["p:text"]?.[0] || "",
            location: commentFile,
            severity: 'medium'
          });
        });
      }
    }
  }
  
  return comments;
}

async function analyzePPTXHiddenContent(zip) {
  const hidden = [];
  const presentationXml = zip.file("ppt/presentation.xml");
  
  if (presentationXml) {
    const content = await presentationXml.async("text");
    const parsed = await parseStringPromise(content);
    
    if (parsed?.["p:presentation"]?.["p:sldIdLst"]?.[0]?.["p:sldId"]) {
      const slides = parsed["p:presentation"]["p:sldIdLst"][0]["p:sldId"];
      slides.forEach((slide, index) => {
        if (slide.$?.show === "0") {
          hidden.push({
            id: `hidden_slide_${index}`,
            type: 'hidden_slide',
            description: `Slide ${index + 1} is hidden`,
            slideNumber: index + 1,
            severity: 'high'
          });
        }
      });
    }
  }
  
  // Check for speaker notes
  const notesFiles = Object.keys(zip.files).filter(name => 
    name.startsWith("ppt/notesSlides/") && name.endsWith(".xml")
  );
  
  if (notesFiles.length > 0) {
    hidden.push({
      id: 'speaker_notes',
      type: 'speaker_notes',
      description: `${notesFiles.length} slide(s) contain speaker notes`,
      severity: 'medium'
    });
  }
  
  return hidden;
}

async function analyzeExcelComments(zip) {
  const comments = [];
  const commentFiles = Object.keys(zip.files).filter(name =>
    name.startsWith("xl/comments") && name.endsWith(".xml")
  );

  for (const commentFile of commentFiles) {
    const file = zip.file(commentFile);
    if (file) {
      const content = await file.async("text");
      const parsed = await parseStringPromise(content);
      
      if (parsed?.comments?.commentList?.[0]?.comment) {
        parsed.comments.commentList[0].comment.forEach((comment, index) => {
          const author = comment.$?.authorId || "Unknown";
          const text = comment.text?.[0]?.t?.[0] || comment.text?.[0] || "";
          
          comments.push({
            id: `excel_comment_${comments.length}`,
            author,
            text: String(text).slice(0, 200),
            location: comment.$?.ref || `Comment ${index + 1}`,
            severity: determineSeverity(String(text))
          });
        });
      }
    }
  }
  
  return comments;
}

async function analyzeExcelHiddenSheets(zip) {
  const hiddenSheets = [];
  const workbookXml = zip.file("xl/workbook.xml");
  
  if (workbookXml) {
    const content = await workbookXml.async("text");
    const parsed = await parseStringPromise(content);
    
    if (parsed?.workbook?.sheets?.[0]?.sheet) {
      parsed.workbook.sheets[0].sheet.forEach((sheet, index) => {
        const state = sheet.$?.state;
        if (state === 'hidden' || state === 'veryHidden') {
          hiddenSheets.push({
            id: `hidden_sheet_${index}`,
            sheetName: sheet.$?.name || `Sheet ${index + 1}`,
            type: state === 'veryHidden' ? 'very_hidden' : 'hidden',
            hasData: true,
            severity: 'high'
          });
        }
      });
    }
  }
  
  return hiddenSheets;
}

async function analyzeExcelSensitiveFormulas(zip) {
  const sensitiveFormulas = [];
  const sheetFiles = Object.keys(zip.files).filter(name =>
    name.startsWith("xl/worksheets/sheet") && name.endsWith(".xml")
  );
  
  for (const sheetFile of sheetFiles) {
    const file = zip.file(sheetFile);
    if (file) {
      const content = await file.async("text");
      const sheetName = sheetFile.replace('xl/worksheets/', '').replace('.xml', '');
      
      const formulaMatches = content.matchAll(/<f[^>]*>(.*?)<\/f>/gs);
      for (const match of formulaMatches) {
        const formula = match[1];
        let risk = 'low';
        let reason = '';
        
        // External references
        if (formula.includes('[') && formula.includes(']')) {
          risk = 'high';
          reason = 'External file reference';
        }
        // Database connections
        else if (/SQL|ODBC/i.test(formula)) {
          risk = 'high';
          reason = 'Database connection';
        }
        // Web queries
        else if (/WEBSERVICE|FILTERXML/i.test(formula)) {
          risk = 'high';
          reason = 'External web query';
        }
        // File paths
        else if (formula.includes(':\\') || formula.includes('/Users/')) {
          risk = 'medium';
          reason = 'File path in formula';
        }
        
        if (risk !== 'low') {
          sensitiveFormulas.push({
            id: `formula_${sensitiveFormulas.length}`,
            sheet: sheetName,
            formula: formula.slice(0, 100),
            risk,
            reason,
            severity: risk
          });
        }
      }
    }
  }
  
  return sensitiveFormulas;
}

async function analyzeEmbeddedObjects(zip, prefix) {
  const embeddings = Object.keys(zip.files).filter(name =>
    name.startsWith(`${prefix}/embeddings/`)
  );
  
  return embeddings.map((name, index) => ({
    id: `embed_${index}`,
    filename: name.split('/').pop(),
    type: 'embedded_object',
    path: name,
    severity: 'medium'
  }));
}

async function analyzeMacros(zip, prefix) {
  const macros = [];
  const hasMacros = Object.keys(zip.files).some(name =>
    name.startsWith(`${prefix}/vbaProject`) || name.endsWith(".bin")
  );
  
  if (hasMacros) {
    macros.push({
      id: 'vba_macros',
      name: 'VBA Macros detected',
      description: 'Document contains executable macro code',
      risk: 'high',
      severity: 'high'
    });
  }
  
  return macros;
}

async function analyzeSensitiveData(zip, fileType) {
  const sensitiveData = [];
  let text = '';
  
  // Extract all text from the document
  if (fileType === 'docx') {
    const docXml = zip.file("word/document.xml");
    if (docXml) text = await docXml.async("text");
  } else if (fileType === 'pptx') {
    const slideFiles = Object.keys(zip.files).filter(name => 
      name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
    );
    for (const sf of slideFiles) {
      const file = zip.file(sf);
      if (file) text += " " + await file.async("text");
    }
  } else if (fileType === 'xlsx') {
    const sharedStrings = zip.file("xl/sharedStrings.xml");
    if (sharedStrings) text = await sharedStrings.async("text");
  }
  
  // Extract readable text
  const readableText = extractTextFromXML(text);
  
  // Email detection
  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
  const emails = readableText.match(emailRegex) || [];
  emails.slice(0, 10).forEach((email, index) => {
    sensitiveData.push({
      id: `email_${index}`,
      type: 'email',
      value: email,
      category: 'personal',
      severity: 'medium'
    });
  });
  
  // Phone detection
  const phoneRegex = /(\+?\d{1,3}[-.\s]?)?(\(?\d{2,4}\)?[-.\s]?)?\d{3,4}[-.\s]?\d{3,4}/g;
  const phones = readableText.match(phoneRegex) || [];
  phones.filter(p => p.replace(/\D/g, '').length >= 8).slice(0, 5).forEach((phone, index) => {
    sensitiveData.push({
      id: `phone_${index}`,
      type: 'phone',
      value: phone.trim(),
      category: 'personal',
      severity: 'medium'
    });
  });
  
  // IBAN detection
  const ibanRegex = /[A-Z]{2}\d{2}[A-Z0-9]{10,30}/g;
  const ibans = readableText.match(ibanRegex) || [];
  ibans.slice(0, 5).forEach((iban, index) => {
    sensitiveData.push({
      id: `iban_${index}`,
      type: 'iban',
      value: iban,
      category: 'financial',
      severity: 'high'
    });
  });
  
  // Pricing detection
  const pricingRegex = /(\d{1,3}(?:[,.\s]\d{3})*(?:[.,]\d{2})?[\s]?[€$£])|([€$£][\s]?\d{1,3}(?:[,.\s]\d{3})*)/g;
  const prices = readableText.match(pricingRegex) || [];
  prices.slice(0, 5).forEach((price, index) => {
    sensitiveData.push({
      id: `pricing_${index}`,
      type: 'pricing',
      value: price,
      category: 'internal',
      severity: 'high'
    });
  });
  
  // Project codes
  const projectCodeRegex = /(PROJ[-_]?\d+)|(\#\d{4,})|([A-Z]{2,4}[-_]\d{3,})/g;
  const codes = readableText.match(projectCodeRegex) || [];
  codes.slice(0, 5).forEach((code, index) => {
    sensitiveData.push({
      id: `project_${index}`,
      type: 'project_code',
      value: code,
      category: 'internal',
      severity: 'medium'
    });
  });
  
  return sensitiveData;
}

function extractTextFromXML(xml) {
  // Remove XML tags and extract text content
  return xml
    .replace(/<[^>]+>/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/\s+/g, ' ')
    .trim();
}

function determineSeverity(text) {
  const lower = text.toLowerCase();
  if (/confidentiel|secret|todo|fixme|urgent|attention/i.test(lower)) return 'high';
  if (/draft|preliminary|internal|review/i.test(lower)) return 'medium';
  return 'low';
}

/**
 * Calcule un résumé des détections
 */
export function calculateSummary(detections) {
  const summary = {
    totalIssues: 0,
    critical: 0,
    high: 0,
    medium: 0,
    low: 0,
    categories: {}
  };
  
  const countItems = (arr, category) => {
    if (!arr || !Array.isArray(arr)) return;
    summary.categories[category] = arr.length;
    summary.totalIssues += arr.length;
    
    arr.forEach(item => {
      const sev = item.severity || item.risk || 'medium';
      if (sev === 'critical') summary.critical++;
      else if (sev === 'high') summary.high++;
      else if (sev === 'medium') summary.medium++;
      else summary.low++;
    });
  };
  
  countItems(detections.metadata, 'metadata');
  countItems(detections.comments, 'comments');
  countItems(detections.trackChanges, 'trackChanges');
  countItems(detections.hiddenContent, 'hiddenContent');
  countItems(detections.hiddenSheets, 'hiddenSheets');
  countItems(detections.sensitiveFormulas, 'sensitiveFormulas');
  countItems(detections.embeddedObjects, 'embeddedObjects');
  countItems(detections.macros, 'macros');
  countItems(detections.sensitiveData, 'sensitiveData');
  
  return summary;
}

export default { analyzeDocument, calculateSummary };
