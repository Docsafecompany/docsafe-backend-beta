// lib/documentAnalyzer.js
// Analyse les documents AVANT nettoyage pour le preview interactif
// Version 3.0 - Enterprise-grade avec 9 catégories de détection

import JSZip from "jszip";
import xml2js from "xml2js";
import { PDFDocument } from "pdf-lib";
import { checkSpellingWithAI } from './aiSpellCheck.js';

const parseStringPromise = xml2js.parseStringPromise;

/**
 * Analyse un document et retourne les détections enrichies
 * @param {Buffer} fileBuffer - Contenu du fichier
 * @param {string} fileType - Type MIME du fichier
 * @returns {Promise<Object>} Résultat d'analyse Enterprise-grade
 */
export async function analyzeDocument(fileBuffer, fileType) {
  const ext = getExtFromMime(fileType);
  const filename = "document." + ext;
  
  let detections;
  
  switch (ext) {
    case 'docx':
      detections = await analyzeDOCX(fileBuffer);
      break;
    case 'pptx':
      detections = await analyzePPTX(fileBuffer);
      break;
    case 'xlsx':
      detections = await analyzeXLSX(fileBuffer);
      break;
    case 'pdf':
      detections = await analyzePDF(fileBuffer);
      break;
    default:
      throw new Error(`Unsupported file type: ${ext}`);
  }
  
  // Calculer le résumé enrichi
  const summary = calculateSummary(detections);
  
  return {
    filename,
    ext,
    fileSize: fileBuffer.length,
    analyzedAt: new Date().toISOString(),
    summary,
    detections
  };
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
  
  // Extraire le texte complet pour le spell-check
  const fullText = await extractDOCXText(zip);
  
  // Appeler l'IA pour les fautes d'orthographe
  console.log('📝 Analyzing DOCX spelling with AI...');
  const spellingErrors = await checkSpellingWithAI(fullText);
  console.log(`✅ Found ${spellingErrors.length} spelling errors in DOCX`);
  
  return {
    // 1️⃣ Sensitive Data
    sensitiveData: await analyzeSensitiveData(zip, 'docx'),
    
    // 2️⃣ Metadata
    metadata: await analyzeOfficeMetadata(zip, 'word'),
    
    // 3️⃣ Comments & Review Traces
    comments: await analyzeDOCXCommentsEnriched(zip),
    
    // 4️⃣ Hidden Content
    hiddenContent: await analyzeDOCXHiddenContentEnriched(zip),
    
    // 5️⃣ Spelling Errors
    spellingErrors: spellingErrors,
    
    // 6️⃣ Visual Objects
    visualObjects: await analyzeVisualObjects(zip, 'docx'),
    
    // 7️⃣ Orphan Data
    orphanData: await analyzeOrphanData(fullText, zip, 'docx'),
    
    // 8️⃣ Macros
    macros: await analyzeMacros(zip, 'word'),
    
    // 9️⃣ Excel Hidden Data (non applicable pour DOCX)
    excelHiddenData: [],
    
    // Legacy fields (pour compatibilité)
    trackChanges: await analyzeDOCXTrackChangesEnriched(zip),
    embeddedObjects: await analyzeEmbeddedObjects(zip, 'word'),
    brokenLinks: await analyzeBrokenLinks(fullText),
    complianceRisks: await analyzeComplianceRisks(fullText)
  };
}

// ============= PPTX ANALYSIS =============
async function analyzePPTX(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  
  // Extraire le texte complet pour le spell-check
  const fullText = await extractPPTXText(zip);
  
  // Appeler l'IA pour les fautes d'orthographe
  console.log('📝 Analyzing PPTX spelling with AI...');
  const spellingErrors = await checkSpellingWithAI(fullText);
  console.log(`✅ Found ${spellingErrors.length} spelling errors in PPTX`);
  
  return {
    // 1️⃣ Sensitive Data
    sensitiveData: await analyzeSensitiveData(zip, 'pptx'),
    
    // 2️⃣ Metadata
    metadata: await analyzeOfficeMetadata(zip, 'ppt'),
    
    // 3️⃣ Comments & Review Traces (inclut speaker notes)
    comments: await analyzePPTXCommentsEnriched(zip),
    
    // 4️⃣ Hidden Content
    hiddenContent: await analyzePPTXHiddenContentEnriched(zip),
    
    // 5️⃣ Spelling Errors
    spellingErrors: spellingErrors,
    
    // 6️⃣ Visual Objects
    visualObjects: await analyzeVisualObjects(zip, 'pptx'),
    
    // 7️⃣ Orphan Data
    orphanData: await analyzeOrphanData(fullText, zip, 'pptx'),
    
    // 8️⃣ Macros
    macros: await analyzeMacros(zip, 'ppt'),
    
    // 9️⃣ Excel Hidden Data (non applicable pour PPTX)
    excelHiddenData: [],
    
    // Legacy fields
    trackChanges: [],
    embeddedObjects: await analyzeEmbeddedObjects(zip, 'ppt'),
    brokenLinks: await analyzeBrokenLinks(fullText),
    complianceRisks: await analyzeComplianceRisks(fullText)
  };
}

// ============= XLSX ANALYSIS =============
async function analyzeXLSX(buffer) {
  const zip = await JSZip.loadAsync(buffer);
  
  // Extraire le texte complet pour le spell-check
  const fullText = await extractXLSXText(zip);
  
  // Appeler l'IA pour les fautes d'orthographe
  console.log('📝 Analyzing XLSX spelling with AI...');
  const spellingErrors = await checkSpellingWithAI(fullText);
  console.log(`✅ Found ${spellingErrors.length} spelling errors in XLSX`);
  
  // Analyser les données Excel cachées
  const hiddenSheets = await analyzeExcelHiddenSheets(zip);
  const hiddenColumns = await analyzeExcelHiddenColumns(zip);
  const sensitiveFormulas = await analyzeExcelSensitiveFormulas(zip);
  
  return {
    // 1️⃣ Sensitive Data
    sensitiveData: await analyzeSensitiveData(zip, 'xlsx'),
    
    // 2️⃣ Metadata
    metadata: await analyzeOfficeMetadata(zip, 'xl'),
    
    // 3️⃣ Comments
    comments: await analyzeExcelCommentsEnriched(zip),
    
    // 4️⃣ Hidden Content (généralisé)
    hiddenContent: [],
    
    // 5️⃣ Spelling Errors
    spellingErrors: spellingErrors,
    
    // 6️⃣ Visual Objects
    visualObjects: [],
    
    // 7️⃣ Orphan Data
    orphanData: await analyzeOrphanData(fullText, zip, 'xlsx'),
    
    // 8️⃣ Macros
    macros: await analyzeMacros(zip, 'xl'),
    
    // 9️⃣ Excel Hidden Data (catégorie consolidée)
    excelHiddenData: [
      ...hiddenSheets.map(s => ({
        id: s.id,
        type: s.type === 'very_hidden' ? 'very_hidden_sheet' : 'hidden_sheet',
        name: s.sheetName,
        description: `${s.type === 'very_hidden' ? 'Very hidden' : 'Hidden'} sheet: ${s.sheetName}`,
        location: `Sheet: ${s.sheetName}`,
        severity: 'high',
        hasData: s.hasData
      })),
      ...hiddenColumns.map(c => ({
        id: c.id,
        type: c.type === 'hidden_row' ? 'hidden_row' : 'hidden_column',
        name: c.columns || `Row ${c.row}`,
        description: c.type === 'hidden_row' 
          ? `Hidden row ${c.row} in ${c.sheet}`
          : `Hidden columns ${c.columns} in ${c.sheet}`,
        location: `${c.sheet}`,
        severity: 'medium',
        hasData: true
      })),
      ...sensitiveFormulas.map(f => ({
        id: f.id,
        type: 'hidden_formula',
        name: f.formula.substring(0, 30) + '...',
        description: f.reason,
        location: `${f.sheet}`,
        severity: f.risk,
        formula: f.formula
      }))
    ],
    
    // Legacy fields
    trackChanges: [],
    hiddenSheets,
    hiddenColumns,
    sensitiveFormulas,
    embeddedObjects: await analyzeEmbeddedObjects(zip, 'xl'),
    brokenLinks: [],
    complianceRisks: await analyzeComplianceRisks(fullText)
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
  const creationDate = pdfDoc.getCreationDate();
  const modificationDate = pdfDoc.getModificationDate();
  
  if (author) metadata.push({ 
    id: `meta_author_${Date.now()}`,
    type: 'author',
    key: 'Author',
    value: author, 
    location: 'Document Properties',
    severity: 'high',
    description: 'Author name exposed in metadata'
  });
  if (title) metadata.push({ 
    id: `meta_title_${Date.now()}`,
    type: 'title',
    key: 'Title',
    value: title, 
    location: 'Document Properties',
    severity: 'low',
    description: 'Document title in metadata'
  });
  if (subject) metadata.push({ 
    id: `meta_subject_${Date.now()}`,
    type: 'subject',
    key: 'Subject',
    value: subject, 
    location: 'Document Properties',
    severity: 'low',
    description: 'Document subject in metadata'
  });
  if (creator) metadata.push({ 
    id: `meta_creator_${Date.now()}`,
    type: 'software',
    key: 'Creator',
    value: creator, 
    location: 'Document Properties',
    severity: 'medium',
    description: 'Creation software exposed'
  });
  if (producer) metadata.push({ 
    id: `meta_producer_${Date.now()}`,
    type: 'software',
    key: 'Producer',
    value: producer, 
    location: 'Document Properties',
    severity: 'low',
    description: 'PDF producer software'
  });
  if (keywords) metadata.push({ 
    id: `meta_keywords_${Date.now()}`,
    type: 'keywords',
    key: 'Keywords',
    value: keywords, 
    location: 'Document Properties',
    severity: 'low',
    description: 'Document keywords'
  });
  if (creationDate) metadata.push({ 
    id: `meta_created_${Date.now()}`,
    type: 'created_date',
    key: 'Creation Date',
    value: creationDate.toISOString(), 
    location: 'Document Properties',
    severity: 'medium',
    description: 'Document creation date'
  });
  if (modificationDate) metadata.push({ 
    id: `meta_modified_${Date.now()}`,
    type: 'modified_date',
    key: 'Modification Date',
    value: modificationDate.toISOString(), 
    location: 'Document Properties',
    severity: 'medium',
    description: 'Document modification date'
  });
  
  return {
    sensitiveData: [],
    metadata,
    comments: [],
    hiddenContent: [],
    spellingErrors: [],
    visualObjects: [],
    orphanData: [],
    macros: [],
    excelHiddenData: [],
    trackChanges: [],
    embeddedObjects: [],
    brokenLinks: [],
    complianceRisks: []
  };
}

// ============= TEXT EXTRACTION HELPERS =============

async function extractDOCXText(zip) {
  let text = '';
  const documentXml = zip.file("word/document.xml");
  if (documentXml) {
    const content = await documentXml.async("text");
    text = extractTextFromXML(content);
  }
  return text;
}

async function extractPPTXText(zip) {
  let text = '';
  const slideFiles = Object.keys(zip.files).filter(name => 
    name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
  );
  for (const sf of slideFiles) {
    const file = zip.file(sf);
    if (file) {
      const content = await file.async("text");
      text += " " + extractTextFromXML(content);
    }
  }
  return text;
}

async function extractXLSXText(zip) {
  let text = '';
  const sharedStrings = zip.file("xl/sharedStrings.xml");
  if (sharedStrings) {
    const content = await sharedStrings.async("text");
    text = extractTextFromXML(content);
  }
  return text;
}

// ============= METADATA ANALYSIS =============

async function analyzeOfficeMetadata(zip, prefix) {
  const metadata = [];
  
  // app.xml
  const appXml = zip.file("docProps/app.xml");
  if (appXml) {
    const content = await appXml.async("text");
    try {
      const parsed = await parseStringPromise(content);
      if (parsed?.Properties) {
        if (parsed.Properties.Company?.[0]) 
          metadata.push({ 
            id: `meta_company_${Date.now()}`,
            type: 'company',
            key: 'Company',
            value: parsed.Properties.Company[0], 
            location: 'Document Properties',
            severity: 'high',
            description: 'Company name exposed in metadata'
          });
        if (parsed.Properties.Manager?.[0]) 
          metadata.push({ 
            id: `meta_manager_${Date.now()}`,
            type: 'manager',
            key: 'Manager',
            value: parsed.Properties.Manager[0], 
            location: 'Document Properties',
            severity: 'high',
            description: 'Manager name exposed in metadata'
          });
        if (parsed.Properties.Application?.[0]) 
          metadata.push({ 
            id: `meta_app_${Date.now()}`,
            type: 'software',
            key: 'Application',
            value: parsed.Properties.Application[0], 
            location: 'Document Properties',
            severity: 'low',
            description: 'Application software version'
          });
        if (parsed.Properties.TotalTime?.[0]) 
          metadata.push({ 
            id: `meta_time_${Date.now()}`,
            type: 'revision',
            key: 'Editing Time',
            value: `${parsed.Properties.TotalTime[0]} minutes`, 
            location: 'Document Properties',
            severity: 'medium',
            description: 'Total editing time exposed'
          });
      }
    } catch (e) { /* ignore parse errors */ }
  }
  
  // core.xml
  const coreXml = zip.file("docProps/core.xml");
  if (coreXml) {
    const content = await coreXml.async("text");
    try {
      const parsed = await parseStringPromise(content);
      if (parsed?.["cp:coreProperties"]) {
        const props = parsed["cp:coreProperties"];
        if (props["dc:creator"]?.[0]) 
          metadata.push({ 
            id: `meta_author_${Date.now()}`,
            type: 'author',
            key: 'Author',
            value: props["dc:creator"][0], 
            location: 'Document Properties',
            severity: 'high',
            description: 'Author name exposed in metadata'
          });
        if (props["dc:title"]?.[0]) 
          metadata.push({ 
            id: `meta_title_${Date.now()}`,
            type: 'title',
            key: 'Title',
            value: props["dc:title"][0], 
            location: 'Document Properties',
            severity: 'low',
            description: 'Document title in metadata'
          });
        if (props["dc:subject"]?.[0]) 
          metadata.push({ 
            id: `meta_subject_${Date.now()}`,
            type: 'subject',
            key: 'Subject',
            value: props["dc:subject"][0], 
            location: 'Document Properties',
            severity: 'low',
            description: 'Document subject in metadata'
          });
        if (props["cp:keywords"]?.[0]) 
          metadata.push({ 
            id: `meta_keywords_${Date.now()}`,
            type: 'keywords',
            key: 'Keywords',
            value: props["cp:keywords"][0], 
            location: 'Document Properties',
            severity: 'medium',
            description: 'Document keywords'
          });
        if (props["cp:lastModifiedBy"]?.[0]) 
          metadata.push({ 
            id: `meta_lastmod_${Date.now()}`,
            type: 'author',
            key: 'Last Modified By',
            value: props["cp:lastModifiedBy"][0], 
            location: 'Document Properties',
            severity: 'high',
            description: 'Last modifier name exposed'
          });
        if (props["dcterms:created"]?.[0]?._) 
          metadata.push({ 
            id: `meta_created_${Date.now()}`,
            type: 'created_date',
            key: 'Creation Date',
            value: props["dcterms:created"][0]._, 
            location: 'Document Properties',
            severity: 'medium',
            description: 'Document creation date'
          });
        if (props["dcterms:modified"]?.[0]?._) 
          metadata.push({ 
            id: `meta_modified_${Date.now()}`,
            type: 'modified_date',
            key: 'Modification Date',
            value: props["dcterms:modified"][0]._, 
            location: 'Document Properties',
            severity: 'medium',
            description: 'Document modification date'
          });
      }
    } catch (e) { /* ignore parse errors */ }
  }
  
  return metadata;
}

// ============= ENRICHED COMMENTS ANALYSIS =============

async function analyzeDOCXCommentsEnriched(zip) {
  const comments = [];
  const commentsXml = zip.file("word/comments.xml");
  
  if (commentsXml) {
    const content = await commentsXml.async("text");
    try {
      const parsed = await parseStringPromise(content);
      
      if (parsed?.["w:comments"]?.["w:comment"]) {
        const commentsList = parsed["w:comments"]["w:comment"];
        commentsList.forEach((comment, index) => {
          const author = comment.$?.["w:author"] || "Unknown Author";
          const date = comment.$?.["w:date"] || null;
          const text = extractTextFromXML(JSON.stringify(comment));
          
          comments.push({
            id: `comment_${index}_${Date.now()}`,
            type: 'comment',
            author: author,
            text: text.trim() || "Empty comment",
            date: date,
            location: `Page ${Math.floor(index / 3) + 1}, Comment ${index + 1}`,
            severity: determineSeverity(text),
            changeType: null,
            originalText: null,
            newText: null
          });
        });
      }
    } catch (e) { console.error('Error parsing DOCX comments:', e); }
  }
  
  // Ajouter les tracked changes comme type de commentaire
  const trackChanges = await analyzeDOCXTrackChangesEnriched(zip);
  trackChanges.forEach(tc => {
    comments.push({
      id: tc.id,
      type: 'tracked_change',
      author: tc.author,
      text: tc.type === 'deletion' 
        ? `Deleted: "${tc.originalText}"`
        : tc.type === 'insertion' 
          ? `Inserted: "${tc.newText}"`
          : `Modified: "${tc.originalText}" → "${tc.newText}"`,
      date: tc.date,
      location: tc.location || 'Document body',
      severity: tc.severity,
      changeType: tc.type,
      originalText: tc.originalText,
      newText: tc.newText
    });
  });
  
  return comments;
}

async function analyzePPTXCommentsEnriched(zip) {
  const comments = [];
  
  // 1. Commentaires classiques
  const commentFiles = Object.keys(zip.files).filter(name => 
    name.startsWith("ppt/comments/comment") && name.endsWith(".xml")
  );
  
  // Get comment authors
  const authors = await getPPTXCommentAuthors(zip);
  
  for (const commentFile of commentFiles) {
    const file = zip.file(commentFile);
    if (file) {
      const content = await file.async("text");
      try {
        const parsed = await parseStringPromise(content);
        
        if (parsed?.["p:cmLst"]?.["p:cm"]) {
          parsed["p:cmLst"]["p:cm"].forEach((cm, index) => {
            const text = cm["p:text"]?.[0] || "";
            const authorId = cm.$?.authorId || "0";
            const authorName = authors[authorId] || "Unknown Author";
            const dt = cm.$?.dt || null;
            
            comments.push({
              id: `ppt_comment_${comments.length}_${Date.now()}`,
              type: 'comment',
              author: authorName,
              text: String(text).trim() || "Empty comment",
              date: dt,
              location: extractSlideNumber(commentFile),
              severity: determineSeverity(String(text))
            });
          });
        }
      } catch (e) { console.error('Error parsing PPTX comment:', e); }
    }
  }
  
  // 2. Commentaires modernes
  const modernCommentFiles = Object.keys(zip.files).filter(name => 
    name.includes("modernComment") && name.endsWith(".xml")
  );
  
  for (const commentFile of modernCommentFiles) {
    const file = zip.file(commentFile);
    if (file) {
      const content = await file.async("text");
      const textContent = extractTextFromXML(content);
      if (textContent && textContent.trim().length > 0) {
        comments.push({
          id: `ppt_modern_comment_${comments.length}_${Date.now()}`,
          type: 'comment',
          author: "Unknown Author",
          text: textContent.trim(),
          date: null,
          location: extractSlideNumber(commentFile),
          severity: determineSeverity(textContent)
        });
      }
    }
  }
  
  // 3. Speaker Notes (ajoutés comme type 'speaker_note')
  const notesFiles = Object.keys(zip.files).filter(name => 
    name.startsWith("ppt/notesSlides/") && name.endsWith(".xml")
  );
  
  for (const notesFile of notesFiles) {
    const file = zip.file(notesFile);
    if (file) {
      const content = await file.async("text");
      const noteText = extractTextFromXML(content);
      if (noteText && noteText.trim().length > 5) {
        const slideMatch = notesFile.match(/notesSlide(\d+)\.xml/);
        const slideNum = slideMatch ? slideMatch[1] : '?';
        
        comments.push({
          id: `ppt_speaker_note_${comments.length}_${Date.now()}`,
          type: 'speaker_note',
          author: "Speaker Notes",
          text: noteText.trim().substring(0, 300) + (noteText.length > 300 ? '...' : ''),
          date: null,
          location: `Slide ${slideNum}`,
          severity: determineSeverity(noteText)
        });
      }
    }
  }
  
  return comments;
}

async function getPPTXCommentAuthors(zip) {
  const authors = {};
  const authorsFile = zip.file("ppt/commentAuthors.xml");
  
  if (authorsFile) {
    const content = await authorsFile.async("text");
    try {
      const parsed = await parseStringPromise(content);
      if (parsed?.["p:cmAuthorLst"]?.["p:cmAuthor"]) {
        parsed["p:cmAuthorLst"]["p:cmAuthor"].forEach(author => {
          const id = author.$?.id || "0";
          const name = author.$?.name || "Unknown";
          authors[id] = name;
        });
      }
    } catch (e) { /* ignore */ }
  }
  
  return authors;
}

function extractSlideNumber(filePath) {
  const match = filePath.match(/slide(\d+)/i);
  return match ? `Slide ${match[1]}` : "Presentation";
}

async function analyzeExcelCommentsEnriched(zip) {
  const comments = [];
  const commentFiles = Object.keys(zip.files).filter(name =>
    name.startsWith("xl/comments") && name.endsWith(".xml")
  );

  for (const commentFile of commentFiles) {
    const file = zip.file(commentFile);
    if (file) {
      const content = await file.async("text");
      try {
        const parsed = await parseStringPromise(content);
        
        if (parsed?.comments?.commentList?.[0]?.comment) {
          parsed.comments.commentList[0].comment.forEach((comment, index) => {
            const ref = comment.$?.ref || `Cell ${index + 1}`;
            
            // Extract text from nested structure
            let text = '';
            if (comment.text?.[0]?.r) {
              // Rich text format
              comment.text[0].r.forEach(r => {
                if (r.t?.[0]) text += r.t[0];
              });
            } else if (comment.text?.[0]?.t?.[0]) {
              text = comment.text[0].t[0];
            } else if (typeof comment.text?.[0] === 'string') {
              text = comment.text[0];
            }
            
            comments.push({
              id: `excel_comment_${comments.length}_${Date.now()}`,
              type: 'comment',
              author: "Cell Comment",
              text: String(text).trim() || "Empty comment",
              date: null,
              location: `Cell ${ref}`,
              severity: determineSeverity(String(text))
            });
          });
        }
      } catch (e) { console.error('Error parsing Excel comments:', e); }
    }
  }
  
  return comments;
}

// ============= ENRICHED TRACK CHANGES ANALYSIS =============

async function analyzeDOCXTrackChangesEnriched(zip) {
  const trackChanges = [];
  const documentXml = zip.file("word/document.xml");
  
  if (documentXml) {
    const content = await documentXml.async("text");
    
    // Detect insertions avec contexte
    const insertMatches = content.matchAll(/<w:ins[^>]*w:author="([^"]*)"[^>]*(?:w:date="([^"]*)")?[^>]*>([\s\S]*?)<\/w:ins>/g);
    for (const match of insertMatches) {
      const newText = extractTextFromXML(match[3]);
      if (newText.trim()) {
        trackChanges.push({
          id: `ins_${trackChanges.length}_${Date.now()}`,
          type: 'insertion',
          author: match[1] || 'Unknown',
          date: match[2] || null,
          originalText: null,
          newText: newText.slice(0, 150),
          location: 'Document body',
          severity: 'medium'
        });
      }
    }
    
    // Detect deletions avec contexte
    const deleteMatches = content.matchAll(/<w:del[^>]*w:author="([^"]*)"[^>]*(?:w:date="([^"]*)")?[^>]*>([\s\S]*?)<\/w:del>/g);
    for (const match of deleteMatches) {
      const originalText = extractTextFromXML(match[3]);
      if (originalText.trim()) {
        trackChanges.push({
          id: `del_${trackChanges.length}_${Date.now()}`,
          type: 'deletion',
          author: match[1] || 'Unknown',
          date: match[2] || null,
          originalText: originalText.slice(0, 150),
          newText: null,
          location: 'Document body',
          severity: 'medium'
        });
      }
    }
    
    // Essayer de matcher les paires insertion/deletion adjacentes comme modifications
    // (Logique simplifiée - dans la réalité, c'est plus complexe)
  }
  
  return trackChanges;
}

// ============= ENRICHED HIDDEN CONTENT ANALYSIS =============

async function analyzeDOCXHiddenContentEnriched(zip) {
  const hidden = [];
  const documentXml = zip.file("word/document.xml");
  
  if (documentXml) {
    const content = await documentXml.async("text");
    
    // 1. Hidden text (w:vanish)
    const hiddenMatches = [...content.matchAll(/<w:vanish[^>]*\/>/g)];
    if (hiddenMatches.length > 0) {
      hidden.push({
        id: `hidden_vanish_${Date.now()}`,
        type: 'vanished_text',
        description: `${hiddenMatches.length} hidden text element(s) using vanish property`,
        content: null,
        location: 'Document body',
        severity: 'high'
      });
    }
    
    // 2. White text on white background
    const whiteTextMatches = [...content.matchAll(/<w:color[^>]*w:val="(FFFFFF|ffffff|white)"[^>]*\/>/gi)];
    if (whiteTextMatches.length > 0) {
      hidden.push({
        id: `hidden_white_${Date.now()}`,
        type: 'white_text',
        description: `${whiteTextMatches.length} white text element(s) detected (potentially hidden content)`,
        content: null,
        location: 'Document body',
        severity: 'high'
      });
    }
    
    // 3. Very small font (potential hidden text)
    const smallFontMatches = [...content.matchAll(/<w:sz[^>]*w:val="([1-9])"[^>]*\/>/g)];
    if (smallFontMatches.length > 0) {
      hidden.push({
        id: `hidden_smallfont_${Date.now()}`,
        type: 'invisible_text',
        description: `${smallFontMatches.length} very small font element(s) detected (< 5pt)`,
        content: null,
        location: 'Document body',
        severity: 'medium'
      });
    }
  }
  
  // 4. Embedded objects
  const embeddings = Object.keys(zip.files).filter(name =>
    name.startsWith('word/embeddings/')
  );
  
  if (embeddings.length > 0) {
    hidden.push({
      id: `hidden_embedded_${Date.now()}`,
      type: 'embedded_file',
      description: `${embeddings.length} embedded file(s) detected`,
      content: embeddings.map(e => e.split('/').pop()).join(', '),
      location: 'Document',
      severity: 'medium'
    });
  }
  
  return hidden;
}

async function analyzePPTXHiddenContentEnriched(zip) {
  const hidden = [];
  
  // 1. Hidden slides
  const presentationXml = zip.file("ppt/presentation.xml");
  if (presentationXml) {
    const content = await presentationXml.async("text");
    try {
      const parsed = await parseStringPromise(content);
      
      if (parsed?.["p:presentation"]?.["p:sldIdLst"]?.[0]?.["p:sldId"]) {
        const slides = parsed["p:presentation"]["p:sldIdLst"][0]["p:sldId"];
        slides.forEach((slide, index) => {
          if (slide.$?.show === "0") {
            hidden.push({
              id: `hidden_slide_${index}_${Date.now()}`,
              type: 'hidden_slide',
              description: `Slide ${index + 1} is marked as hidden`,
              content: null,
              location: `Slide ${index + 1}`,
              severity: 'high'
            });
          }
        });
      }
    } catch (e) { /* ignore */ }
  }
  
  // 2. Analyze each slide for hidden content
  const slideFiles = Object.keys(zip.files).filter(name => 
    name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
  );
  
  for (const slideFile of slideFiles) {
    const file = zip.file(slideFile);
    if (file) {
      const content = await file.async("text");
      const slideMatch = slideFile.match(/slide(\d+)\.xml/);
      const slideNum = slideMatch ? slideMatch[1] : '?';
      
      // White/invisible text
      const whiteTextCount = (content.match(/<a:srgbClr val="(FFFFFF|ffffff)"/g) || []).length;
      if (whiteTextCount > 2) {
        hidden.push({
          id: `hidden_white_slide${slideNum}_${Date.now()}`,
          type: 'white_text',
          description: `${whiteTextCount} white text elements on slide ${slideNum}`,
          content: null,
          location: `Slide ${slideNum}`,
          severity: 'high'
        });
      }
      
      // Off-slide content (position outside normal bounds)
      const offSlideMatches = content.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/g);
      for (const match of offSlideMatches) {
        const x = parseInt(match[1]);
        const y = parseInt(match[2]);
        // Standard slide is ~9144000 x 6858000 EMUs
        if (x < -1000000 || x > 10000000 || y < -1000000 || y > 8000000) {
          hidden.push({
            id: `hidden_offslide_${hidden.length}_${Date.now()}`,
            type: 'off_slide_content',
            description: `Content positioned outside slide bounds`,
            content: null,
            location: `Slide ${slideNum}`,
            severity: 'medium'
          });
          break; // Only report once per slide
        }
      }
    }
  }
  
  return hidden;
}

// ============= NEW: VISUAL OBJECTS ANALYSIS =============

async function analyzeVisualObjects(zip, fileType) {
  const visualObjects = [];
  
  if (fileType === 'pptx') {
    const slideFiles = Object.keys(zip.files).filter(name => 
      name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
    );
    
    for (const slideFile of slideFiles) {
      const file = zip.file(slideFile);
      if (file) {
        const content = await file.async("text");
        const slideMatch = slideFile.match(/slide(\d+)\.xml/);
        const slideNum = slideMatch ? slideMatch[1] : '?';
        
        // Detect shapes that might cover text
        // Look for solid fill shapes positioned over text areas
        const shapeMatches = content.matchAll(/<p:sp[^>]*>([\s\S]*?)<\/p:sp>/g);
        let potentialCoveringShapes = 0;
        
        for (const match of shapeMatches) {
          const shapeContent = match[1];
          // Shape with solid fill and no text
          const hasSolidFill = shapeContent.includes('<a:solidFill>');
          const hasText = shapeContent.includes('<a:t>');
          const isLargeShape = shapeContent.match(/<a:ext cx="(\d+)" cy="(\d+)"\/>/);
          
          if (hasSolidFill && !hasText && isLargeShape) {
            const [, cx, cy] = isLargeShape;
            if (parseInt(cx) > 2000000 && parseInt(cy) > 500000) {
              potentialCoveringShapes++;
            }
          }
        }
        
        if (potentialCoveringShapes > 0) {
          visualObjects.push({
            id: `visual_covering_slide${slideNum}_${Date.now()}`,
            type: 'shape_covering_text',
            description: `${potentialCoveringShapes} large solid shape(s) that may cover content`,
            location: `Slide ${slideNum}`,
            severity: 'medium',
            shapeType: 'rectangle'
          });
        }
        
        // Check for shapes without alt text (accessibility issue)
        const noAltTextShapes = (content.match(/<p:sp[^>]*>(?:(?!<p:cNvPr[^>]*descr=).)*?<\/p:sp>/gs) || []).length;
        if (noAltTextShapes > 3) {
          visualObjects.push({
            id: `visual_noalt_slide${slideNum}_${Date.now()}`,
            type: 'missing_alt_text',
            description: `${noAltTextShapes} shapes without alt text (accessibility issue)`,
            location: `Slide ${slideNum}`,
            severity: 'low'
          });
        }
      }
    }
  }
  
  if (fileType === 'docx') {
    const documentXml = zip.file("word/document.xml");
    if (documentXml) {
      const content = await documentXml.async("text");
      
      // Detect drawing objects that might cover text
      const drawingMatches = content.matchAll(/<w:drawing[^>]*>([\s\S]*?)<\/w:drawing>/g);
      let coveringShapes = 0;
      
      for (const match of drawingMatches) {
        const drawingContent = match[1];
        // Inline vs anchor - anchor can cover text
        if (drawingContent.includes('<wp:anchor')) {
          const hasFill = drawingContent.includes('<a:solidFill>');
          if (hasFill) coveringShapes++;
        }
      }
      
      if (coveringShapes > 0) {
        visualObjects.push({
          id: `visual_covering_docx_${Date.now()}`,
          type: 'shape_covering_text',
          description: `${coveringShapes} anchored shape(s) with solid fill may cover content`,
          location: 'Document body',
          severity: 'medium',
          shapeType: 'drawing'
        });
      }
    }
  }
  
  return visualObjects;
}

// ============= NEW: ORPHAN DATA ANALYSIS =============

async function analyzeOrphanData(fullText, zip, fileType) {
  const orphanData = [];
  
  // 1. Broken links
  const brokenLinks = await analyzeBrokenLinks(fullText);
  brokenLinks.forEach(link => {
    orphanData.push({
      id: link.id,
      type: 'broken_link',
      description: `Broken or local link: ${link.url.substring(0, 50)}...`,
      value: link.url,
      location: link.location || 'Document',
      severity: 'low',
      suggestedAction: 'Remove or update link'
    });
  });
  
  // 2. Empty pages/slides detection
  if (fileType === 'pptx') {
    const slideFiles = Object.keys(zip.files).filter(name => 
      name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
    );
    
    for (const slideFile of slideFiles) {
      const file = zip.file(slideFile);
      if (file) {
        const content = await file.async("text");
        const textContent = extractTextFromXML(content);
        const slideMatch = slideFile.match(/slide(\d+)\.xml/);
        const slideNum = slideMatch ? slideMatch[1] : '?';
        
        if (textContent.trim().length < 10) {
          orphanData.push({
            id: `orphan_empty_slide${slideNum}_${Date.now()}`,
            type: 'empty_page',
            description: `Slide ${slideNum} appears to be empty or has minimal content`,
            value: null,
            location: `Slide ${slideNum}`,
            severity: 'low',
            suggestedAction: 'Review or remove empty slide'
          });
        }
      }
    }
  }
  
  // 3. Trailing whitespace detection
  const trailingWhitespaceMatches = fullText.match(/\s{3,}/g) || [];
  if (trailingWhitespaceMatches.length > 5) {
    orphanData.push({
      id: `orphan_whitespace_${Date.now()}`,
      type: 'trailing_whitespace',
      description: `${trailingWhitespaceMatches.length} instances of excessive whitespace detected`,
      value: null,
      location: 'Throughout document',
      severity: 'low',
      suggestedAction: 'Clean up formatting'
    });
  }
  
  return orphanData;
}

// ============= EXCEL SPECIFIC ANALYSIS =============

async function analyzeExcelHiddenSheets(zip) {
  const hiddenSheets = [];
  const workbookXml = zip.file("xl/workbook.xml");
  
  if (workbookXml) {
    const content = await workbookXml.async("text");
    try {
      const parsed = await parseStringPromise(content);
      
      if (parsed?.workbook?.sheets?.[0]?.sheet) {
        parsed.workbook.sheets[0].sheet.forEach((sheet, index) => {
          const state = sheet.$?.state;
          if (state === 'hidden' || state === 'veryHidden') {
            hiddenSheets.push({
              id: `hidden_sheet_${index}_${Date.now()}`,
              sheetName: sheet.$?.name || `Sheet ${index + 1}`,
              type: state === 'veryHidden' ? 'very_hidden' : 'hidden',
              hasData: true,
              severity: 'high'
            });
          }
        });
      }
    } catch (e) { /* ignore */ }
  }
  
  return hiddenSheets;
}

async function analyzeExcelHiddenColumns(zip) {
  const hiddenColumns = [];
  const sheetFiles = Object.keys(zip.files).filter(name =>
    name.startsWith("xl/worksheets/sheet") && name.endsWith(".xml")
  );
  
  for (const sheetFile of sheetFiles) {
    const file = zip.file(sheetFile);
    if (file) {
      const content = await file.async("text");
      const sheetName = sheetFile.replace('xl/worksheets/', '').replace('.xml', '');
      
      // Hidden columns
      const colMatches = content.matchAll(/<col[^>]*hidden="1"[^>]*min="(\d+)"[^>]*max="(\d+)"[^>]*\/>/g);
      for (const match of colMatches) {
        hiddenColumns.push({
          id: `hidden_col_${hiddenColumns.length}_${Date.now()}`,
          sheet: sheetName,
          columns: `${match[1]}-${match[2]}`,
          type: 'hidden_column',
          severity: 'high'
        });
      }
      
      // Hidden rows
      const rowMatches = content.matchAll(/<row[^>]*hidden="1"[^>]*r="(\d+)"[^>]*>/g);
      for (const match of rowMatches) {
        hiddenColumns.push({
          id: `hidden_row_${hiddenColumns.length}_${Date.now()}`,
          sheet: sheetName,
          row: match[1],
          type: 'hidden_row',
          severity: 'high'
        });
      }
    }
  }
  
  return hiddenColumns;
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
        
        if (formula.includes('[') && formula.includes(']')) {
          risk = 'high';
          reason = 'External file reference detected';
        } else if (/SQL|ODBC/i.test(formula)) {
          risk = 'high';
          reason = 'Database connection formula';
        } else if (/WEBSERVICE|FILTERXML/i.test(formula)) {
          risk = 'high';
          reason = 'External web query';
        } else if (formula.includes(':\\') || formula.includes('/Users/')) {
          risk = 'medium';
          reason = 'Local file path in formula';
        } else if (/INDIRECT|OFFSET/i.test(formula)) {
          risk = 'low';
          reason = 'Dynamic reference formula';
        }
        
        if (risk !== 'low') {
          sensitiveFormulas.push({
            id: `formula_${sensitiveFormulas.length}_${Date.now()}`,
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

// ============= EMBEDDED OBJECTS & MACROS =============

async function analyzeEmbeddedObjects(zip, prefix) {
  const embeddings = Object.keys(zip.files).filter(name =>
    name.startsWith(`${prefix}/embeddings/`)
  );
  
  return embeddings.map((name, index) => ({
    id: `embed_${index}_${Date.now()}`,
    filename: name.split('/').pop(),
    type: 'embedded_object',
    path: name,
    severity: 'medium'
  }));
}

async function analyzeMacros(zip, prefix) {
  const macros = [];
  const macroFiles = Object.keys(zip.files).filter(name =>
    name.startsWith(`${prefix}/vbaProject`) || 
    name.endsWith(".bin") ||
    name.includes('vbaProject')
  );
  
  if (macroFiles.length > 0) {
    macros.push({
      id: `vba_macros_${Date.now()}`,
      type: 'vba_macro',
      name: 'VBA Macros',
      description: 'Document contains executable VBA macro code - potential security risk',
      location: 'VBA Project',
      severity: 'critical',
      isMalicious: false,
      code: null
    });
  }
  
  // Check for auto-executing macros
  const autoMacroPatterns = ['AutoOpen', 'AutoClose', 'AutoExec', 'Document_Open', 'Workbook_Open'];
  // (In reality, you'd need to parse the VBA binary to detect these)
  
  return macros;
}

// ============= SENSITIVE DATA DETECTION =============

async function analyzeSensitiveData(zip, fileType) {
  const sensitiveData = [];
  let text = '';
  
  // Extract all text
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
  
  const readableText = extractTextFromXML(text);
  
  // Email detection
  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
  const emails = [...new Set(readableText.match(emailRegex) || [])];
  emails.slice(0, 10).forEach((email, index) => {
    sensitiveData.push({
      id: `email_${index}_${Date.now()}`,
      type: 'email',
      value: email,
      context: getContext(readableText, email),
      location: 'Document body',
      category: 'personal',
      severity: 'medium',
      gdprRelevant: true,
      description: 'Email address detected'
    });
  });
  
  // Phone detection
  const phoneRegex = /(\+?\d{1,3}[-.\s]?)?(\(?\d{2,4}\)?[-.\s]?)?\d{3,4}[-.\s]?\d{3,4}/g;
  const phones = (readableText.match(phoneRegex) || [])
    .filter(p => p.replace(/\D/g, '').length >= 8);
  [...new Set(phones)].slice(0, 5).forEach((phone, index) => {
    sensitiveData.push({
      id: `phone_${index}_${Date.now()}`,
      type: 'phone',
      value: phone.trim(),
      context: getContext(readableText, phone),
      location: 'Document body',
      category: 'personal',
      severity: 'medium',
      gdprRelevant: true,
      description: 'Phone number detected'
    });
  });
  
  // IBAN detection
  const ibanRegex = /[A-Z]{2}\d{2}[A-Z0-9]{10,30}/g;
  const ibans = readableText.match(ibanRegex) || [];
  [...new Set(ibans)].slice(0, 5).forEach((iban, index) => {
    sensitiveData.push({
      id: `iban_${index}_${Date.now()}`,
      type: 'iban',
      value: iban,
      context: getContext(readableText, iban),
      location: 'Document body',
      category: 'financial',
      severity: 'critical',
      gdprRelevant: false,
      description: 'IBAN bank account number detected'
    });
  });
  
  // Pricing detection
  const pricingRegex = /(\d{1,3}(?:[,.\s]\d{3})*(?:[.,]\d{2})?[\s]?[€$£])|([€$£][\s]?\d{1,3}(?:[,.\s]\d{3})*)/g;
  const prices = readableText.match(pricingRegex) || [];
  [...new Set(prices)].slice(0, 5).forEach((price, index) => {
    sensitiveData.push({
      id: `pricing_${index}_${Date.now()}`,
      type: 'pricing',
      value: price,
      context: getContext(readableText, price),
      location: 'Document body',
      category: 'internal',
      severity: 'high',
      gdprRelevant: false,
      description: 'Pricing/financial amount detected'
    });
  });
  
  // Project codes
  const projectCodeRegex = /(PROJ[-_]?\d+)|(\#\d{4,})|([A-Z]{2,4}[-_]\d{3,})/g;
  const codes = readableText.match(projectCodeRegex) || [];
  [...new Set(codes)].slice(0, 5).forEach((code, index) => {
    sensitiveData.push({
      id: `project_${index}_${Date.now()}`,
      type: 'project_code',
      value: code,
      context: getContext(readableText, code),
      location: 'Document body',
      category: 'internal',
      severity: 'medium',
      gdprRelevant: false,
      description: 'Internal project code detected'
    });
  });
  
  // File paths
  const filePathRegex = /([A-Z]:\\[^\s<>"]+)|([\/](?:Users|home|var|etc|mnt)[\/][^\s<>"]+)/gi;
  const paths = readableText.match(filePathRegex) || [];
  [...new Set(paths)].slice(0, 5).forEach((path, index) => {
    sensitiveData.push({
      id: `filepath_${index}_${Date.now()}`,
      type: 'file_path',
      value: path,
      context: getContext(readableText, path),
      location: 'Document body',
      category: 'internal',
      severity: 'high',
      gdprRelevant: false,
      description: 'Internal file path detected'
    });
  });
  
  // Internal servers
  const serverRegex = /(https?:\/\/[a-zA-Z0-9\-_.]+\.(internal|local|corp|intranet)[^\s]*)|([a-zA-Z0-9\-]+\.(internal|local|corp|intranet))/gi;
  const servers = readableText.match(serverRegex) || [];
  [...new Set(servers)].slice(0, 5).forEach((server, index) => {
    sensitiveData.push({
      id: `server_${index}_${Date.now()}`,
      type: 'internal_server',
      value: server,
      context: getContext(readableText, server),
      location: 'Document body',
      category: 'internal',
      severity: 'high',
      gdprRelevant: false,
      description: 'Internal server/URL detected'
    });
  });
  
  // IP addresses
  const ipRegex = /\b(?:\d{1,3}\.){3}\d{1,3}\b/g;
  const ips = readableText.match(ipRegex) || [];
  [...new Set(ips)].filter(ip => !ip.startsWith('0.') && !ip.startsWith('127.')).slice(0, 5).forEach((ip, index) => {
    sensitiveData.push({
      id: `ip_${index}_${Date.now()}`,
      type: 'ip_address',
      value: ip,
      context: getContext(readableText, ip),
      location: 'Document body',
      category: 'internal',
      severity: 'medium',
      gdprRelevant: false,
      description: 'IP address detected'
    });
  });
  
  return sensitiveData;
}

function getContext(text, value) {
  const index = text.indexOf(value);
  if (index === -1) return value;
  const start = Math.max(0, index - 30);
  const end = Math.min(text.length, index + value.length + 30);
  return text.substring(start, end).replace(/\s+/g, ' ').trim();
}

// ============= BROKEN LINKS ANALYSIS =============

async function analyzeBrokenLinks(text) {
  const brokenLinks = [];
  
  // Local file links
  const localLinkRegex = /file:\/\/[^\s<>"]+/gi;
  const localLinks = text.match(localLinkRegex) || [];
  localLinks.slice(0, 5).forEach((link, index) => {
    brokenLinks.push({
      id: `broken_local_${index}_${Date.now()}`,
      type: 'local_file_link',
      url: link,
      location: 'Document body',
      reason: 'Local file link will not work for recipients',
      severity: 'medium'
    });
  });
  
  // SharePoint/OneDrive internal links
  const internalLinkRegex = /https?:\/\/[^\s<>"]*sharepoint\.com[^\s<>"]*/gi;
  const internalLinks = text.match(internalLinkRegex) || [];
  internalLinks.slice(0, 3).forEach((link, index) => {
    brokenLinks.push({
      id: `broken_sharepoint_${index}_${Date.now()}`,
      type: 'internal_sharepoint',
      url: link.substring(0, 80) + '...',
      location: 'Document body',
      reason: 'SharePoint link may not be accessible to external recipients',
      severity: 'low'
    });
  });
  
  return brokenLinks;
}

// ============= COMPLIANCE RISKS =============

async function analyzeComplianceRisks(text) {
  const complianceRisks = [];
  const readableText = extractTextFromXML(text).toLowerCase();
  
  const gdprPatterns = [
    { pattern: /numéro de sécurité sociale|social security|ssn/i, risk: 'Social Security Number reference', severity: 'critical' },
    { pattern: /date de naissance|birth date|dob|née? le/i, risk: 'Birth date detected', severity: 'high' },
    { pattern: /passeport|passport/i, risk: 'Passport reference', severity: 'high' },
    { pattern: /carte d'identité|identity card|id card|cni/i, risk: 'ID card reference', severity: 'high' },
    { pattern: /numéro de permis|driver'?s? license|permis de conduire/i, risk: 'Driver license reference', severity: 'high' },
    { pattern: /données de santé|health data|medical record|dossier médical/i, risk: 'Health data reference', severity: 'critical' },
  ];
  
  gdprPatterns.forEach((p, index) => {
    if (p.pattern.test(readableText)) {
      complianceRisks.push({
        id: `gdpr_${index}_${Date.now()}`,
        type: 'gdpr',
        description: p.risk,
        location: 'Document body',
        severity: p.severity
      });
    }
  });
  
  return complianceRisks;
}

// ============= HELPER FUNCTIONS =============

function extractTextFromXML(xml) {
  return xml
    .replace(/<[^>]+>/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function determineSeverity(text) {
  const lower = (text || '').toLowerCase();
  if (/confidentiel|confidential|secret|password|mot de passe|urgent|critical|do not share|ne pas partager/i.test(lower)) return 'high';
  if (/todo|fixme|draft|preliminary|internal|review|brouillon|à revoir/i.test(lower)) return 'medium';
  return 'low';
}

// ============= SUMMARY CALCULATION =============

export function calculateSummary(detections) {
  let totalIssues = 0;
  let criticalIssues = 0;
  let highIssues = 0;
  let mediumIssues = 0;
  let lowIssues = 0;
  let weightedScore = 0;
  
  const weights = {
    sensitiveData: 6,
    metadata: 2,
    comments: 3,
    hiddenContent: 5,
    spellingErrors: 1,
    visualObjects: 2,
    orphanData: 1,
    macros: 10,
    excelHiddenData: 4,
    trackChanges: 2,
    complianceRisks: 8,
    brokenLinks: 1
  };
  
  const severityMultiplier = {
    critical: 4,
    high: 3,
    medium: 2,
    low: 1
  };
  
  Object.entries(detections).forEach(([key, arr]) => {
    if (Array.isArray(arr) && arr.length > 0) {
      const baseWeight = weights[key] || 1;
      
      arr.forEach(item => {
        totalIssues++;
        const severity = item.severity || item.risk || 'medium';
        const multiplier = severityMultiplier[severity] || 2;
        
        if (severity === 'critical') criticalIssues++;
        else if (severity === 'high') highIssues++;
        else if (severity === 'medium') mediumIssues++;
        else lowIssues++;
        
        weightedScore += baseWeight * multiplier;
      });
    }
  });
  
  // Score 100 = safe, 0 = dangerous
  let beforeRiskScore = Math.max(0, Math.min(100, 100 - weightedScore));
  
  // Apply penalties for critical issues
  if (criticalIssues > 0) beforeRiskScore = Math.min(beforeRiskScore, 30);
  if (detections.macros?.length > 0) beforeRiskScore = Math.min(beforeRiskScore, 20);
  if (highIssues > 3) beforeRiskScore = Math.min(beforeRiskScore, 50);
  
  let riskLevel;
  if (beforeRiskScore >= 90) riskLevel = 'safe';
  else if (beforeRiskScore >= 70) riskLevel = 'low';
  else if (beforeRiskScore >= 50) riskLevel = 'medium';
  else if (beforeRiskScore >= 25) riskLevel = 'high';
  else riskLevel = 'critical';
  
  const recommendations = generateRecommendations(detections);
  
  return {
    totalIssues,
    critical: criticalIssues,
    high: highIssues,
    medium: mediumIssues,
    low: lowIssues,
    beforeRiskScore,
    riskLevel,
    recommendations,
    categories: {
      sensitiveData: detections.sensitiveData?.length || 0,
      metadata: detections.metadata?.length || 0,
      comments: detections.comments?.length || 0,
      hiddenContent: detections.hiddenContent?.length || 0,
      spellingErrors: detections.spellingErrors?.length || 0,
      visualObjects: detections.visualObjects?.length || 0,
      orphanData: detections.orphanData?.length || 0,
      macros: detections.macros?.length || 0,
      excelHiddenData: detections.excelHiddenData?.length || 0
    }
  };
}

function generateRecommendations(detections) {
  const recommendations = [];
  
  if (detections.sensitiveData?.length > 0) {
    recommendations.push({
      priority: 'high',
      category: 'Data Protection',
      text: 'Review and redact sensitive data (emails, phones, financial info) before sharing.',
      icon: '🔒'
    });
  }
  if (detections.metadata?.length > 0) {
    recommendations.push({
      priority: 'medium',
      category: 'Privacy',
      text: 'Remove document metadata to protect author and organization identity.',
      icon: '📋'
    });
  }
  if (detections.comments?.length > 0) {
    recommendations.push({
      priority: 'medium',
      category: 'Review Process',
      text: 'Remove all comments and speaker notes before sharing externally.',
      icon: '💬'
    });
  }
  if (detections.hiddenContent?.length > 0) {
    recommendations.push({
      priority: 'high',
      category: 'Security',
      text: 'Review and remove hidden content that may contain sensitive information.',
      icon: '👁️'
    });
  }
  if (detections.macros?.length > 0) {
    recommendations.push({
      priority: 'critical',
      category: 'Security',
      text: '⚠️ CRITICAL: Remove macros to prevent potential security vulnerabilities.',
      icon: '⚠️'
    });
  }
  if (detections.spellingErrors?.length > 0) {
    recommendations.push({
      priority: 'low',
      category: 'Quality',
      text: `Correct ${detections.spellingErrors.length} spelling/grammar error(s) for professional quality.`,
      icon: '✍️'
    });
  }
  if (detections.excelHiddenData?.length > 0) {
    recommendations.push({
      priority: 'high',
      category: 'Excel Security',
      text: 'Remove hidden sheets, columns, and sensitive formulas from spreadsheet.',
      icon: '📊'
    });
  }
  if (detections.complianceRisks?.length > 0) {
    recommendations.push({
      priority: 'high',
      category: 'Compliance',
      text: 'Address GDPR/compliance risks before external distribution.',
      icon: '⚖️'
    });
  }
  
  if (recommendations.length === 0) {
    recommendations.push({
      priority: 'low',
      category: 'Status',
      text: '✅ Document appears clean. No critical issues detected.',
      icon: '✅'
    });
  }
  
  return recommendations;
}

export default { analyzeDocument, calculateSummary };

