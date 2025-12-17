// lib/documentAnalyzer.js
// Analyse les documents AVANT nettoyage pour le preview interactif
// VERSION 3.3.0 - Universal Risk Model (4 exposure surfaces) + Executive Risk Categories
//
// ✅ Adds: riskObjects[] (universal schema) WITHOUT breaking existing detections.*
// ✅ Adds: riskSummary (clientReady + overallSeverity + executiveSignals)
// ✅ Adds: deterministic "Delivery & Commitment" signals (regex) -> detections.businessInconsistencies
//
// IMPORTANT: still returns { detections, summary, documentStats } exactly like before.

import JSZip from "jszip";
import xml2js from "xml2js";
import { PDFDocument } from "pdf-lib";
import { checkSpellingWithAI } from "./aiProofreadAnchored.js";
import { detectSensitiveData } from "./sensitiveData.js";
import { extractPdfText, filterExtractedLines } from "./pdfTools.js";

const parseStringPromise = xml2js.parseStringPromise;

// ============================================================
// PUBLIC API
// ============================================================

/**
 * Analyse un document et retourne les détections enrichies + riskObjects (universal)
 * @param {Buffer} fileBuffer - Contenu du fichier
 * @param {string} fileType - Type MIME du fichier
 * @returns {Promise<Object>} Résultat d'analyse Enterprise-grade
 */
export async function analyzeDocument(fileBuffer, fileType) {
  const ext = getExtFromMime(fileType);
  if (!ext) throw new Error(`Unsupported file type: ${fileType}`);

  const filename = "document." + ext;

  // Stats (pages/slides/sheets/tables)
  const documentStats = await computeDocumentStats(fileBuffer, ext);

  let detections;

  switch (ext) {
    case "docx":
      detections = await analyzeDOCX(fileBuffer);
      break;
    case "pptx":
      detections = await analyzePPTX(fileBuffer);
      break;
    case "xlsx":
      detections = await analyzeXLSX(fileBuffer);
      break;
    case "pdf":
      detections = await analyzePDF(fileBuffer);
      break;
    default:
      throw new Error(`Unsupported file type: ${ext}`);
  }

  // Summary (existing)
  const summary = calculateSummary(detections);

  // ✅ NEW: riskObjects + riskSummary (universal 4-surfaces)
  const { riskObjects, riskSummary } = buildUniversalRiskOutput({
    ext,
    detections,
    summary,
  });

  return {
    filename,
    ext,
    fileSize: fileBuffer.length,
    analyzedAt: new Date().toISOString(),

    documentStats,
    summary,
    detections,

    // ✅ NEW (non-breaking additions)
    riskObjects,
    riskSummary,
  };
}

function getExtFromMime(mime) {
  const map = {
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "docx",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": "pptx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "xlsx",
    "application/pdf": "pdf",
  };
  return map[mime] || null;
}

// ============================================================
// DOCUMENT STATS (pages/slides/sheets/tables)
// ============================================================

async function computeDocumentStats(fileBuffer, ext) {
  const documentStats = { pages: 0, slides: 0, sheets: 0, tables: 0 };

  try {
    if (ext === "docx") {
      const zip = await JSZip.loadAsync(fileBuffer);
      const docXml = await zip.file("word/document.xml")?.async("string");
      if (docXml) {
        const sectionBreaks = (docXml.match(/<w:sectPr/g) || []).length;
        documentStats.pages = Math.max(1, sectionBreaks);
        documentStats.tables = (docXml.match(/<w:tbl>/g) || []).length;
      }
    } else if (ext === "pptx") {
      const zip = await JSZip.loadAsync(fileBuffer);
      const slideFiles = Object.keys(zip.files).filter((f) => f.match(/ppt\/slides\/slide\d+\.xml/));
      documentStats.slides = slideFiles.length;
    } else if (ext === "xlsx") {
      const zip = await JSZip.loadAsync(fileBuffer);
      const workbookXml = await zip.file("xl/workbook.xml")?.async("string");
      if (workbookXml) {
        documentStats.sheets = (workbookXml.match(/<sheet /g) || []).length;
      }
    } else if (ext === "pdf") {
      const pdfDoc = await PDFDocument.load(fileBuffer);
      documentStats.pages = pdfDoc.getPageCount();
    }
  } catch (e) {
    console.warn("computeDocumentStats failed:", e);
  }

  return documentStats;
}

// ============================================================
// DOCX ANALYSIS
// ============================================================

async function analyzeDOCX(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  const fullText = await extractDOCXText(zip);

  // AI spell-check (detection only)
  console.log("📝 Analyzing DOCX spelling with AI...");
  const spellingErrors = await checkSpellingWithAI(fullText);
  console.log(`✅ Found ${spellingErrors.length} spelling errors in DOCX`);

  // Sensitive data (deterministic)
  console.log("🔍 Analyzing sensitive data...");
  const sensitiveDataResult = detectSensitiveData(fullText);
  console.log(`✅ Found ${sensitiveDataResult.findings.length} sensitive data items`);

  // ✅ Deterministic "Delivery & Commitment" signals (regex)
  const businessInconsistencies = detectDeliveryCommitmentSignals(fullText, "docx");

  return {
    sensitiveData: sensitiveDataResult.findings.map((f, idx) => ({
      id: `sensitive_${idx}_${Date.now()}`,
      type: f.type,
      value: f.value, // (NOTE: may contain sensitive info; consider redaction later)
      maskedValue: f.masked || null,
      context: f.context?.full || "",
      location: f.location || "Document body",
      category: getCategoryFromType(f.type),
      severity: f.severity || "medium",
      gdprRelevant: ["email", "phone", "ssn", "credit_card"].includes(f.type),
      description: f.recommendation || "",
    })),

    metadata: await analyzeOfficeMetadata(zip),

    comments: await analyzeDOCXCommentsEnriched(zip),

    hiddenContent: await analyzeDOCXHiddenContentEnriched(zip),

    spellingErrors,

    visualObjects: await analyzeVisualObjects(zip, "docx"),

    orphanData: await analyzeOrphanData(fullText, zip, "docx"),

    macros: await analyzeMacros(zip, "word"),

    excelHiddenData: [],

    // Legacy fields
    trackChanges: await analyzeDOCXTrackChangesEnriched(zip),
    embeddedObjects: await analyzeEmbeddedObjects(zip, "word"),
    brokenLinks: await analyzeBrokenLinks(fullText),
    complianceRisks: await analyzeComplianceRisks(fullText),

    // ✅ already referenced by your server.js output (keep it)
    sensitiveFormulas: [],
    hiddenSheets: [],
    hiddenColumns: [],

    // ✅ NEW (was referenced in server.js output but missing in analyzer)
    businessInconsistencies,
  };
}

// ============================================================
// PPTX ANALYSIS
// ============================================================

async function analyzePPTX(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  const fullText = await extractPPTXText(zip);

  console.log("📝 Analyzing PPTX spelling with AI...");
  const spellingErrors = await checkSpellingWithAI(fullText);
  console.log(`✅ Found ${spellingErrors.length} spelling errors in PPTX`);

  console.log("🔍 Analyzing sensitive data...");
  const sensitiveDataResult = detectSensitiveData(fullText);
  console.log(`✅ Found ${sensitiveDataResult.findings.length} sensitive data items`);

  const businessInconsistencies = detectDeliveryCommitmentSignals(fullText, "pptx");

  return {
    sensitiveData: sensitiveDataResult.findings.map((f, idx) => ({
      id: `sensitive_${idx}_${Date.now()}`,
      type: f.type,
      value: f.value,
      maskedValue: f.masked || null,
      context: f.context?.full || "",
      location: f.location || "Slides",
      category: getCategoryFromType(f.type),
      severity: f.severity || "medium",
      gdprRelevant: ["email", "phone", "ssn", "credit_card"].includes(f.type),
      description: f.recommendation || "",
    })),

    metadata: await analyzeOfficeMetadata(zip),

    comments: await analyzePPTXCommentsEnriched(zip),

    hiddenContent: await analyzePPTXHiddenContentEnriched(zip),

    spellingErrors,

    visualObjects: await analyzeVisualObjects(zip, "pptx"),

    orphanData: await analyzeOrphanData(fullText, zip, "pptx"),

    macros: await analyzeMacros(zip, "ppt"),

    excelHiddenData: [],

    // Legacy fields
    trackChanges: [],
    embeddedObjects: await analyzeEmbeddedObjects(zip, "ppt"),
    brokenLinks: await analyzeBrokenLinks(fullText),
    complianceRisks: await analyzeComplianceRisks(fullText),

    sensitiveFormulas: [],
    hiddenSheets: [],
    hiddenColumns: [],

    businessInconsistencies,
  };
}

// ============================================================
// XLSX ANALYSIS
// ============================================================

async function analyzeXLSX(buffer) {
  const zip = await JSZip.loadAsync(buffer);

  const fullText = await extractXLSXText(zip);

  console.log("📝 Analyzing XLSX spelling with AI...");
  const spellingErrors = await checkSpellingWithAI(fullText);
  console.log(`✅ Found ${spellingErrors.length} spelling errors in XLSX`);

  console.log("🔍 Analyzing sensitive data...");
  const sensitiveDataResult = detectSensitiveData(fullText);
  console.log(`✅ Found ${sensitiveDataResult.findings.length} sensitive data items`);

  const hiddenSheets = await analyzeExcelHiddenSheets(zip);
  const hiddenColumns = await analyzeExcelHiddenColumns(zip);
  const sensitiveFormulas = await analyzeExcelSensitiveFormulas(zip);

  const businessInconsistencies = detectDeliveryCommitmentSignals(fullText, "xlsx");

  return {
    sensitiveData: sensitiveDataResult.findings.map((f, idx) => ({
      id: `sensitive_${idx}_${Date.now()}`,
      type: f.type,
      value: f.value,
      maskedValue: f.masked || null,
      context: f.context?.full || "",
      location: f.location || "Workbook",
      category: getCategoryFromType(f.type),
      severity: f.severity || "medium",
      gdprRelevant: ["email", "phone", "ssn", "credit_card"].includes(f.type),
      description: f.recommendation || "",
    })),

    metadata: await analyzeOfficeMetadata(zip),

    comments: await analyzeExcelCommentsEnriched(zip),

    hiddenContent: [],

    spellingErrors,

    visualObjects: [],

    orphanData: await analyzeOrphanData(fullText, zip, "xlsx"),

    macros: await analyzeMacros(zip, "xl"),

    excelHiddenData: [
      ...hiddenSheets.map((s) => ({
        id: s.id,
        type: s.type === "very_hidden" ? "very_hidden_sheet" : "hidden_sheet",
        name: s.sheetName,
        description: `${s.type === "very_hidden" ? "Very hidden" : "Hidden"} sheet: ${s.sheetName}`,
        location: `Sheet: ${s.sheetName}`,
        severity: "high",
        hasData: s.hasData,
      })),
      ...hiddenColumns.map((c) => ({
        id: c.id,
        type: c.type === "hidden_row" ? "hidden_row" : "hidden_column",
        name: c.columns || `Row ${c.row}`,
        description:
          c.type === "hidden_row"
            ? `Hidden row ${c.row} in ${c.sheet}`
            : `Hidden columns ${c.columns} in ${c.sheet}`,
        location: `${c.sheet}`,
        severity: "medium",
        hasData: true,
      })),
      ...sensitiveFormulas.map((f) => ({
        id: f.id,
        type: "hidden_formula",
        name: (f.formula || "").substring(0, 30) + "...",
        description: f.reason,
        location: `${f.sheet}`,
        severity: f.risk,
        formula: f.formula,
      })),
    ],

    // Legacy fields
    trackChanges: [],
    hiddenSheets,
    hiddenColumns,
    sensitiveFormulas,
    embeddedObjects: await analyzeEmbeddedObjects(zip, "xl"),
    brokenLinks: [],
    complianceRisks: await analyzeComplianceRisks(fullText),

    businessInconsistencies,
  };
}

// ============================================================
// PDF ANALYSIS
// ============================================================

async function analyzePDF(buffer) {
  const pdfDoc = await PDFDocument.load(buffer);

  // Metadata
  const metadata = [];
  const author = pdfDoc.getAuthor();
  const title = pdfDoc.getTitle();
  const subject = pdfDoc.getSubject();
  const creator = pdfDoc.getCreator();
  const producer = pdfDoc.getProducer();
  const keywords = pdfDoc.getKeywords();
  const creationDate = pdfDoc.getCreationDate();
  const modificationDate = pdfDoc.getModificationDate();

  const pushMeta = (type, key, value, severity, description) => {
    if (!value) return;
    metadata.push({
      id: `meta_${type}_${Date.now()}`,
      type,
      key,
      value: String(value),
      location: "Document Properties",
      severity,
      description,
    });
  };

  pushMeta("author", "Author", author, "high", "Author name exposed in metadata");
  pushMeta("title", "Title", title, "low", "Document title in metadata");
  pushMeta("subject", "Subject", subject, "low", "Document subject in metadata");
  pushMeta("software", "Creator", creator, "medium", "Creation software exposed");
  pushMeta("software", "Producer", producer, "low", "PDF producer software");
  pushMeta("keywords", "Keywords", keywords, "low", "Document keywords");
  pushMeta("created_date", "Creation Date", creationDate?.toISOString?.() || null, "medium", "Document creation date");
  pushMeta(
    "modified_date",
    "Modification Date",
    modificationDate?.toISOString?.() || null,
    "medium",
    "Document modification date"
  );

  // Text extraction (best effort)
  console.log("📄 Extracting PDF text for analysis...");
  let text = "";
  try {
    const rawText = await extractPdfText(buffer);
    text = filterExtractedLines(rawText, { strictPdf: false });
  } catch (e) {
    console.warn("PDF text extraction failed, continuing with metadata only:", e);
    text = "";
  }

  // Spell check (AI detection only)
  let spellingErrors = [];
  if (text && text.trim().length > 0) {
    console.log("📝 Analyzing PDF spelling with AI...");
    spellingErrors = await checkSpellingWithAI(text);
    console.log(`✅ Found ${spellingErrors.length} spelling errors in PDF`);
  }

  // Sensitive data
  console.log("🔍 Analyzing sensitive data in PDF...");
  const sensitiveDataResult = text ? detectSensitiveData(text) : { findings: [] };
  console.log(`✅ Found ${sensitiveDataResult.findings.length} sensitive data items in PDF`);

  const sensitiveData = sensitiveDataResult.findings.map((f, idx) => ({
    id: `sensitive_${idx}_${Date.now()}`,
    type: f.type,
    value: f.value,
    maskedValue: f.masked || null,
    context: f.context?.full || "",
    location: f.location || "PDF text",
    category: getCategoryFromType(f.type),
    severity: f.severity || "medium",
    gdprRelevant: ["email", "phone", "ssn", "credit_card"].includes(f.type),
    description: f.recommendation || "",
  }));

  const brokenLinks = text ? await analyzeBrokenLinks(text) : [];
  const complianceRisks = text ? await analyzeComplianceRisks(text) : [];
  const orphanData = await analyzeOrphanData(text || "", null, "pdf");

  const businessInconsistencies = detectDeliveryCommitmentSignals(text || "", "pdf");

  return {
    sensitiveData,
    metadata,
    comments: [],
    hiddenContent: [],
    spellingErrors,
    visualObjects: [],
    orphanData,
    macros: [],
    excelHiddenData: [],
    trackChanges: [],
    embeddedObjects: [],
    brokenLinks,
    complianceRisks,

    sensitiveFormulas: [],
    hiddenSheets: [],
    hiddenColumns: [],

    businessInconsistencies,
  };
}

// ============================================================
// CATEGORY FROM TYPE (sensitiveData)
// ============================================================

function getCategoryFromType(type) {
  const categoryMap = {
    email: "personal",
    phone: "personal",
    ssn: "personal",
    credit_card: "financial",
    iban: "financial",
    price: "financial",
    project_code: "internal",
    file_path: "internal",
    server_path: "internal",
    ip_address: "internal",
    internal_url: "internal",
    confidential_keyword: "confidential",
  };
  return categoryMap[type] || "other";
}

// ============================================================
// TEXT EXTRACTION HELPERS (FULL COVERAGE)
// ============================================================

async function extractDOCXText(zip) {
  let text = "";

  const targets = [
    "word/document.xml",
    ...Object.keys(zip.files).filter((k) => /word\/header\d+\.xml$/.test(k)),
    ...Object.keys(zip.files).filter((k) => /word\/footer\d+\.xml$/.test(k)),
    "word/footnotes.xml",
    "word/endnotes.xml",
  ];

  const seen = new Set();
  for (const p of targets) {
    if (seen.has(p)) continue;
    seen.add(p);
    const f = zip.file(p);
    if (!f) continue;
    const xml = await f.async("text");
    const extracted = extractTextFromXML(xml);
    if (extracted) text += "\n" + extracted;

    const tbxMatches = xml.matchAll(/<w:txbxContent\b[^>]*>[\s\S]*?<\/w:txbxContent>/g);
    for (const m of tbxMatches) {
      const tbxText = extractTextFromXML(m[0]);
      if (tbxText) text += "\n" + tbxText;
    }
  }

  return text.replace(/\s+\n/g, "\n").trim();
}

async function extractPPTXText(zip) {
  let text = "";

  const files = Object.keys(zip.files).filter(
    (k) =>
      (k.startsWith("ppt/slides/slide") || k.startsWith("ppt/notesSlides/")) &&
      k.endsWith(".xml")
  );

  for (const p of files) {
    const f = zip.file(p);
    if (!f) continue;
    const xml = await f.async("text");
    const extracted = extractTextFromXML(xml);
    if (extracted) text += "\n" + extracted;
  }

  return text.replace(/\s+\n/g, "\n").trim();
}

async function extractXLSXText(zip) {
  let text = "";

  const sharedStrings = zip.file("xl/sharedStrings.xml");
  if (sharedStrings) {
    const xml = await sharedStrings.async("text");
    const extracted = extractTextFromXML(xml);
    if (extracted) text += "\n" + extracted;
  }

  if (!text || text.trim().length < 5) {
    const sheetFiles = Object.keys(zip.files).filter(
      (k) => k.startsWith("xl/worksheets/sheet") && k.endsWith(".xml")
    );

    for (const p of sheetFiles) {
      const f = zip.file(p);
      if (!f) continue;
      const xml = await f.async("text");

      const inlineStr = Array.from(xml.matchAll(/<is>[\s\S]*?<\/is>/g)).map((m) => extractTextFromXML(m[0]));
      const values = Array.from(xml.matchAll(/<v>([\s\S]*?)<\/v>/g)).map((m) => (m[1] || "").trim());

      const merged = [...inlineStr, ...values].filter(Boolean).join(" ");
      if (merged) text += "\n" + merged;
    }
  }

  return text.replace(/\s+\n/g, "\n").trim();
}

// ============================================================
// METADATA (Office)
// ============================================================

async function analyzeOfficeMetadata(zip) {
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
            type: "company",
            key: "Company",
            value: parsed.Properties.Company[0],
            location: "Document Properties",
            severity: "high",
            description: "Company name exposed in metadata",
          });
        if (parsed.Properties.Manager?.[0])
          metadata.push({
            id: `meta_manager_${Date.now()}`,
            type: "manager",
            key: "Manager",
            value: parsed.Properties.Manager[0],
            location: "Document Properties",
            severity: "high",
            description: "Manager name exposed in metadata",
          });
        if (parsed.Properties.Application?.[0])
          metadata.push({
            id: `meta_app_${Date.now()}`,
            type: "software",
            key: "Application",
            value: parsed.Properties.Application[0],
            location: "Document Properties",
            severity: "low",
            description: "Application software version",
          });
        if (parsed.Properties.TotalTime?.[0])
          metadata.push({
            id: `meta_time_${Date.now()}`,
            type: "revision",
            key: "Editing Time",
            value: `${parsed.Properties.TotalTime[0]} minutes`,
            location: "Document Properties",
            severity: "medium",
            description: "Total editing time exposed",
          });
      }
    } catch {
      /* ignore */
    }
  }

  // core.xml
  const coreXml = zip.file("docProps/core.xml");
  if (coreXml) {
    const content = await coreXml.async("text");
    try {
      const parsed = await parseStringPromise(content);
      if (parsed?.["cp:coreProperties"]) {
        const props = parsed["cp:coreProperties"];
        const pushCore = (id, type, key, value, severity, description) => {
          if (!value) return;
          metadata.push({
            id,
            type,
            key,
            value,
            location: "Document Properties",
            severity,
            description,
          });
        };

        pushCore(`meta_author_${Date.now()}`, "author", "Author", props["dc:creator"]?.[0], "high", "Author name exposed");
        pushCore(`meta_title_${Date.now()}`, "title", "Title", props["dc:title"]?.[0], "low", "Document title in metadata");
        pushCore(`meta_subject_${Date.now()}`, "subject", "Subject", props["dc:subject"]?.[0], "low", "Document subject in metadata");
        pushCore(`meta_keywords_${Date.now()}`, "keywords", "Keywords", props["cp:keywords"]?.[0], "medium", "Document keywords");
        pushCore(
          `meta_lastmod_${Date.now()}`,
          "author",
          "Last Modified By",
          props["cp:lastModifiedBy"]?.[0],
          "high",
          "Last modifier name exposed"
        );
        pushCore(
          `meta_created_${Date.now()}`,
          "created_date",
          "Creation Date",
          props["dcterms:created"]?.[0]?._,
          "medium",
          "Document creation date"
        );
        pushCore(
          `meta_modified_${Date.now()}`,
          "modified_date",
          "Modification Date",
          props["dcterms:modified"]?.[0]?._,
          "medium",
          "Document modification date"
        );
      }
    } catch {
      /* ignore */
    }
  }

  return metadata;
}

// ============================================================
// COMMENTS / TRACK CHANGES / HIDDEN CONTENT / VISUAL / ORPHAN
// (Your existing logic kept, unchanged where possible)
// ============================================================

function extractTextFromParsedXmlNode(node) {
  if (!node) return "";
  if (typeof node === "string") return node;
  if (Array.isArray(node)) return node.map(extractTextFromParsedXmlNode).join(" ");
  if (typeof node === "object") {
    let out = "";
    if (typeof node._ === "string") out += node._ + " ";
    if (node["w:t"]) out += extractTextFromParsedXmlNode(node["w:t"]) + " ";
    if (node.t) out += extractTextFromParsedXmlNode(node.t) + " ";
    for (const k of Object.keys(node)) {
      if (k === "_" || k === "$" || k === "w:t" || k === "t") continue;
      out += extractTextFromParsedXmlNode(node[k]) + " ";
    }
    return out.replace(/\s+/g, " ").trim();
  }
  return "";
}

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

          const extracted = extractTextFromParsedXmlNode(comment);
          const text = (extracted || "").trim();

          comments.push({
            id: `comment_${index}_${Date.now()}`,
            type: "comment",
            author,
            text: text || "Empty comment",
            date,
            location: `Comment ${index + 1}`,
            severity: determineSeverity(text),
            changeType: null,
            originalText: null,
            newText: null,
          });
        });
      }
    } catch (e) {
      console.error("Error parsing DOCX comments:", e);
    }
  }

  const trackChanges = await analyzeDOCXTrackChangesEnriched(zip);
  trackChanges.forEach((tc) => {
    comments.push({
      id: tc.id,
      type: "tracked_change",
      author: tc.author,
      text:
        tc.type === "deletion"
          ? `Deleted: "${tc.originalText}"`
          : tc.type === "insertion"
          ? `Inserted: "${tc.newText}"`
          : `Modified: "${tc.originalText}" → "${tc.newText}"`,
      date: tc.date,
      location: tc.location || "Document body",
      severity: tc.severity,
      changeType: tc.type,
      originalText: tc.originalText,
      newText: tc.newText,
    });
  });

  return comments;
}

async function analyzePPTXCommentsEnriched(zip) {
  const comments = [];

  const commentFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith("ppt/comments/comment") && name.endsWith(".xml")
  );

  const authors = await getPPTXCommentAuthors(zip);

  for (const commentFile of commentFiles) {
    const file = zip.file(commentFile);
    if (file) {
      const content = await file.async("text");
      try {
        const parsed = await parseStringPromise(content);

        if (parsed?.["p:cmLst"]?.["p:cm"]) {
          parsed["p:cmLst"]["p:cm"].forEach((cm) => {
            const text = cm["p:text"]?.[0] || "";
            const authorId = cm.$?.authorId || "0";
            const authorName = authors[authorId] || "Unknown Author";
            const dt = cm.$?.dt || null;

            comments.push({
              id: `ppt_comment_${comments.length}_${Date.now()}`,
              type: "comment",
              author: authorName,
              text: String(text).trim() || "Empty comment",
              date: dt,
              location: extractSlideNumber(commentFile),
              severity: determineSeverity(String(text)),
            });
          });
        }
      } catch (e) {
        console.error("Error parsing PPTX comment:", e);
      }
    }
  }

  const modernCommentFiles = Object.keys(zip.files).filter(
    (name) => name.includes("modernComment") && name.endsWith(".xml")
  );

  for (const commentFile of modernCommentFiles) {
    const file = zip.file(commentFile);
    if (file) {
      const content = await file.async("text");
      const textContent = extractTextFromXML(content);
      if (textContent && textContent.trim().length > 0) {
        comments.push({
          id: `ppt_modern_comment_${comments.length}_${Date.now()}`,
          type: "comment",
          author: "Unknown Author",
          text: textContent.trim(),
          date: null,
          location: extractSlideNumber(commentFile),
          severity: determineSeverity(textContent),
        });
      }
    }
  }

  const notesFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith("ppt/notesSlides/") && name.endsWith(".xml")
  );

  for (const notesFile of notesFiles) {
    const file = zip.file(notesFile);
    if (file) {
      const content = await file.async("text");
      const noteText = extractTextFromXML(content);
      if (noteText && noteText.trim().length > 5) {
        const slideMatch = notesFile.match(/notesSlide(\d+)\.xml/);
        const slideNum = slideMatch ? slideMatch[1] : "?";

        comments.push({
          id: `ppt_speaker_note_${comments.length}_${Date.now()}`,
          type: "speaker_note",
          author: "Speaker Notes",
          text: noteText.trim().substring(0, 300) + (noteText.length > 300 ? "..." : ""),
          date: null,
          location: `Slide ${slideNum}`,
          severity: determineSeverity(noteText),
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
        parsed["p:cmAuthorLst"]["p:cmAuthor"].forEach((author) => {
          const id = author.$?.id || "0";
          const name = author.$?.name || "Unknown";
          authors[id] = name;
        });
      }
    } catch {
      /* ignore */
    }
  }

  return authors;
}

function extractSlideNumber(filePath) {
  const match = filePath.match(/slide(\d+)/i);
  return match ? `Slide ${match[1]}` : "Presentation";
}

async function analyzeExcelCommentsEnriched(zip) {
  const comments = [];
  const commentFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith("xl/comments") && name.endsWith(".xml")
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

            let text = "";
            if (comment.text?.[0]?.r) {
              comment.text[0].r.forEach((r) => {
                if (r.t?.[0]) text += r.t[0];
              });
            } else if (comment.text?.[0]?.t?.[0]) {
              text = comment.text[0].t[0];
            } else if (typeof comment.text?.[0] === "string") {
              text = comment.text[0];
            }

            comments.push({
              id: `excel_comment_${comments.length}_${Date.now()}`,
              type: "comment",
              author: "Cell Comment",
              text: String(text).trim() || "Empty comment",
              date: null,
              location: `Cell ${ref}`,
              severity: determineSeverity(String(text)),
            });
          });
        }
      } catch (e) {
        console.error("Error parsing Excel comments:", e);
      }
    }
  }

  return comments;
}

async function analyzeDOCXTrackChangesEnriched(zip) {
  const trackChanges = [];
  const documentXml = zip.file("word/document.xml");

  if (documentXml) {
    const content = await documentXml.async("text");

    const insertMatches = content.matchAll(
      /<w:ins[^>]*w:author="([^"]*)"[^>]*(?:w:date="([^"]*)")?[^>]*>([\s\S]*?)<\/w:ins>/g
    );
    for (const match of insertMatches) {
      const newText = extractTextFromXML(match[3]);
      if (newText.trim()) {
        trackChanges.push({
          id: `ins_${trackChanges.length}_${Date.now()}`,
          type: "insertion",
          author: match[1] || "Unknown",
          date: match[2] || null,
          originalText: null,
          newText: newText.slice(0, 150),
          location: "Document body",
          severity: "medium",
        });
      }
    }

    const deleteMatches = content.matchAll(
      /<w:del[^>]*w:author="([^"]*)"[^>]*(?:w:date="([^"]*)")?[^>]*>([\s\S]*?)<\/w:del>/g
    );
    for (const match of deleteMatches) {
      const originalText = extractTextFromXML(match[3]);
      if (originalText.trim()) {
        trackChanges.push({
          id: `del_${trackChanges.length}_${Date.now()}`,
          type: "deletion",
          author: match[1] || "Unknown",
          date: match[2] || null,
          originalText: originalText.slice(0, 150),
          newText: null,
          location: "Document body",
          severity: "medium",
        });
      }
    }
  }

  return trackChanges;
}

async function analyzeDOCXHiddenContentEnriched(zip) {
  const hidden = [];
  const detailedHiddenText = [];

  const documentXml = zip.file("word/document.xml");

  if (documentXml) {
    const content = await documentXml.async("text");

    const runRegex = /<w:r\b[^>]*>[\s\S]*?<\/w:r>/g;
    let match;
    let runIndex = 0;

    while ((match = runRegex.exec(content)) !== null) {
      const runXml = match[0];

      const hasVanish = /<w:vanish\b[^\/>]*\/>/.test(runXml);
      const colorMatch = /<w:color[^>]*w:val="([0-9A-Fa-f]{6}|[a-zA-Z]+)"/.exec(runXml);
      const sizeMatch = /<w:sz[^>]*w:val="(\d+)"/.exec(runXml);

      const isWhiteColor = colorMatch && /^(FFFFFF|ffffff|white)$/i.test(colorMatch[1]);
      const isTinyFont = sizeMatch && parseInt(sizeMatch[1], 10) < 10;

      if (hasVanish || isWhiteColor || isTinyFont) {
        const text = extractTextFromXML(runXml);
        const cleanText = (text || "").trim();
        if (!cleanText) {
          runIndex++;
          continue;
        }

        let reason = "hidden_style";
        if (hasVanish) reason = "vanish";
        else if (isWhiteColor) reason = "white_color";
        else if (isTinyFont) reason = "tiny_font";

        detailedHiddenText.push({
          id: `hidden_text_${runIndex}_${Date.now()}`,
          reason,
          text: cleanText,
          preview: cleanText.length > 120 ? cleanText.slice(0, 117) + "..." : cleanText,
          color: colorMatch ? colorMatch[1] : null,
          fontSizeHalfPoints: sizeMatch ? parseInt(sizeMatch[1], 10) : null,
          location: `Run #${runIndex} in document.xml`,
        });
      }

      runIndex++;
    }

    const vanishCount = detailedHiddenText.filter((e) => e.reason === "vanish").length;
    const whiteCount = detailedHiddenText.filter((e) => e.reason === "white_color").length;
    const tinyCount = detailedHiddenText.filter((e) => e.reason === "tiny_font").length;

    if (vanishCount > 0) {
      hidden.push({
        id: `hidden_vanish_${Date.now()}`,
        type: "vanished_text",
        description: `${vanishCount} hidden text element(s) using vanish property`,
        content: null,
        location: "Document body",
        severity: "high",
      });
    }

    if (whiteCount > 0) {
      hidden.push({
        id: `hidden_white_${Date.now()}`,
        type: "white_text",
        description: `${whiteCount} white text element(s) detected (potentially hidden content)`,
        content: null,
        location: "Document body",
        severity: "high",
      });
    }

    if (tinyCount > 0) {
      hidden.push({
        id: `hidden_smallfont_${Date.now()}`,
        type: "invisible_text",
        description: `${tinyCount} very small font element(s) detected (< 5pt)`,
        content: null,
        location: "Document body",
        severity: "medium",
        items: null,
      });
    }

    if (detailedHiddenText.length > 0) {
      hidden.push({
        id: `hidden_details_${Date.now()}`,
        type: "hidden_text_details",
        description: `${detailedHiddenText.length} hidden/white text run(s) detected in document body`,
        location: "Document body",
        severity: "high",
        items: detailedHiddenText,
      });
    }
  }

  const embeddings = Object.keys(zip.files).filter((name) => name.startsWith("word/embeddings/"));

  if (embeddings.length > 0) {
    hidden.push({
      id: `hidden_embedded_${Date.now()}`,
      type: "embedded_file",
      description: `${embeddings.length} embedded file(s) detected`,
      content: embeddings.map((e) => e.split("/").pop()).join(", "),
      location: "Document",
      severity: "medium",
    });
  }

  return hidden;
}

async function analyzePPTXHiddenContentEnriched(zip) {
  const hidden = [];
  const detailedHiddenText = [];

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
              type: "hidden_slide",
              description: `Slide ${index + 1} is marked as hidden`,
              content: null,
              location: `Slide ${index + 1}`,
              severity: "high",
            });
          }
        });
      }
    } catch {
      /* ignore */
    }
  }

  const slideFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
  );

  for (const slideFile of slideFiles) {
    const file = zip.file(slideFile);
    if (!file) continue;

    const content = await file.async("text");
    const slideMatch = slideFile.match(/slide(\d+)\.xml/);
    const slideNum = slideMatch ? slideMatch[1] : "?";
    const slideLocation = `Slide ${slideNum}`;

    const runRegex = /<a:r\b[^>]*>[\s\S]*?<\/a:r>/g;
    let m;
    const localWhiteTexts = [];

    while ((m = runRegex.exec(content)) !== null) {
      const runXml = m[0];
      const hasWhiteColor = /<a:srgbClr[^>]*val="(FFFFFF|ffffff)"/.test(runXml);
      if (!hasWhiteColor) continue;

      const text = extractTextFromXML(runXml);
      const cleanText = (text || "").trim();
      if (!cleanText) continue;

      localWhiteTexts.push({
        id: `ppt_hidden_text_${slideNum}_${localWhiteTexts.length}_${Date.now()}`,
        reason: "white_color",
        text: cleanText,
        preview: cleanText.length > 120 ? cleanText.slice(0, 117) + "..." : cleanText,
        color: "FFFFFF",
        location: slideLocation,
      });
    }

    if (localWhiteTexts.length > 0) {
      detailedHiddenText.push(...localWhiteTexts);

      hidden.push({
        id: `hidden_white_slide${slideNum}_${Date.now()}`,
        type: "white_text",
        description: `${localWhiteTexts.length} white text element(s) on slide ${slideNum}`,
        content: null,
        location: slideLocation,
        severity: "high",
      });
    }

    const offSlideMatches = content.matchAll(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/g);
    for (const match of offSlideMatches) {
      const x = parseInt(match[1], 10);
      const y = parseInt(match[2], 10);
      if (x < -1000000 || x > 10000000 || y < -1000000 || y > 8000000) {
        hidden.push({
          id: `hidden_offslide_${hidden.length}_${Date.now()}`,
          type: "off_slide_content",
          description: `Content positioned outside slide bounds`,
          content: null,
          location: slideLocation,
          severity: "medium",
        });
        break;
      }
    }
  }

  if (detailedHiddenText.length > 0) {
    hidden.push({
      id: `hidden_details_ppt_${Date.now()}`,
      type: "hidden_text_details",
      description: `${detailedHiddenText.length} hidden/white text run(s) detected across slides`,
      location: "Slides",
      severity: "high",
      items: detailedHiddenText,
    });
  }

  return hidden;
}

// ================= VISUAL OBJECTS =================

function pickAttr(tag, attr) {
  if (!tag) return null;
  const re = new RegExp(`${attr}="([^"]+)"`, "i");
  const m = re.exec(tag);
  return m ? m[1] : null;
}

function detectPptShapeColor(shapeXml) {
  const m1 = /<a:srgbClr[^>]*val="([0-9A-Fa-f]{6})"/.exec(shapeXml);
  if (m1) return m1[1].toUpperCase();
  const m2 = /<a:schemeClr[^>]*val="([^"]+)"/.exec(shapeXml);
  if (m2) return m2[1];
  return null;
}

function detectPptShapeGeom(shapeXml) {
  const m = /<a:prstGeom[^>]*prst="([^"]+)"/.exec(shapeXml);
  return m ? m[1] : null;
}

function detectPptShapeExt(shapeXml) {
  const m = /<a:ext[^>]*cx="(\d+)"[^>]*cy="(\d+)"/.exec(shapeXml);
  if (!m) return null;
  return { cx: parseInt(m[1], 10), cy: parseInt(m[2], 10) };
}

async function analyzeVisualObjects(zip, fileType) {
  const visualObjects = [];

  if (fileType === "pptx") {
    const slideFiles = Object.keys(zip.files).filter(
      (name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
    );

    for (const slideFile of slideFiles) {
      const file = zip.file(slideFile);
      if (!file) continue;

      const content = await file.async("text");
      const slideMatch = slideFile.match(/slide(\d+)\.xml/);
      const slideNum = slideMatch ? slideMatch[1] : "?";
      const slideLocation = `Slide ${slideNum}`;

      const shapeMatches = content.matchAll(/<p:sp\b[^>]*>[\s\S]*?<\/p:sp>/g);

      let potentialCoveringShapes = 0;
      let missingAlt = 0;

      for (const m of shapeMatches) {
        const shapeXml = m[0];

        const cNvPrTag = /<p:cNvPr\b[^>]*\/?>/i.exec(shapeXml)?.[0] || "";
        const shapeId = pickAttr(cNvPrTag, "id") || null;
        const shapeName = pickAttr(cNvPrTag, "name") || null;
        const descr = pickAttr(cNvPrTag, "descr") || null;

        if (!descr) missingAlt++;

        const geom = detectPptShapeGeom(shapeXml);
        const ext = detectPptShapeExt(shapeXml);
        const color = detectPptShapeColor(shapeXml);

        const hasSolidFill = shapeXml.includes("<a:solidFill>");
        const hasTextBody = shapeXml.includes("<p:txBody") || shapeXml.includes("<a:t>");
        const text = hasTextBody ? extractTextFromXML(shapeXml) : "";

        const cleanText = (text || "").trim();
        if (cleanText.length > 0) {
          visualObjects.push({
            id: `visual_shape_text_slide${slideNum}_${shapeId || visualObjects.length}_${Date.now()}`,
            type: "shape_text",
            description: `Text inside shape${geom ? ` (${geom})` : ""}`,
            text: cleanText,
            location: `${slideLocation}${shapeName ? ` — ${shapeName}` : ""}${shapeId ? ` (id ${shapeId})` : ""}`,
            severity: determineSeverity(cleanText),
            shape: {
              id: shapeId,
              name: shapeName,
              kind: geom || "shape",
              fill: color ? { rgb: color } : null,
              ext: ext || null,
            },
          });
        }

        if (hasSolidFill && (!cleanText || cleanText.length === 0) && ext) {
          if (ext.cx > 2000000 && ext.cy > 500000) {
            potentialCoveringShapes++;
            visualObjects.push({
              id: `visual_covering_slide${slideNum}_${shapeId || visualObjects.length}_${Date.now()}`,
              type: "shape_covering_text",
              description: `Large solid shape may cover content${geom ? ` (${geom})` : ""}`,
              location: `${slideLocation}${shapeName ? ` — ${shapeName}` : ""}${shapeId ? ` (id ${shapeId})` : ""}`,
              severity: "medium",
              shape: { id: shapeId, name: shapeName, kind: geom || "shape", fill: color ? { rgb: color } : null, ext },
            });
          }
        }
      }

      if (potentialCoveringShapes > 0) {
        visualObjects.push({
          id: `visual_covering_summary_slide${slideNum}_${Date.now()}`,
          type: "shape_covering_text",
          description: `${potentialCoveringShapes} large solid shape(s) that may cover content`,
          location: slideLocation,
          severity: "medium",
          shapeType: "solid_fill_large",
        });
      }

      if (missingAlt > 3) {
        visualObjects.push({
          id: `visual_noalt_slide${slideNum}_${Date.now()}`,
          type: "missing_alt_text",
          description: `${missingAlt} shapes without alt text (accessibility issue)`,
          location: slideLocation,
          severity: "low",
        });
      }
    }
  }

  if (fileType === "docx") {
    const documentXml = zip.file("word/document.xml");
    if (documentXml) {
      const content = await documentXml.async("text");

      const tbxMatches = content.matchAll(/<w:txbxContent\b[^>]*>[\s\S]*?<\/w:txbxContent>/g);
      let tbxIndex = 0;

      for (const m of tbxMatches) {
        const tbxXml = m[0];
        const tbxText = extractTextFromXML(tbxXml);
        const clean = (tbxText || "").trim();
        if (clean.length > 0) {
          visualObjects.push({
            id: `visual_shape_text_docx_${tbxIndex}_${Date.now()}`,
            type: "shape_text",
            description: "Text inside textbox/shape",
            text: clean,
            location: `Document body — Textbox #${tbxIndex + 1}`,
            severity: determineSeverity(clean),
            shape: { kind: "textbox" },
          });
        }
        tbxIndex++;
      }

      const drawingMatches = content.matchAll(/<w:drawing[^>]*>([\s\S]*?)<\/w:drawing>/g);
      let coveringShapes = 0;

      for (const match of drawingMatches) {
        const drawingContent = match[1];
        if (drawingContent.includes("<wp:anchor")) {
          const hasFill = drawingContent.includes("<a:solidFill>");
          if (hasFill) coveringShapes++;
        }
      }

      if (coveringShapes > 0) {
        visualObjects.push({
          id: `visual_covering_docx_${Date.now()}`,
          type: "shape_covering_text",
          description: `${coveringShapes} anchored shape(s) with solid fill may cover content`,
          location: "Document body",
          severity: "medium",
          shapeType: "drawing",
        });
      }
    }
  }

  return visualObjects;
}

// ================= ORPHAN DATA =================

async function analyzeOrphanData(fullText, zip, fileType) {
  const orphanData = [];

  const brokenLinks = await analyzeBrokenLinks(fullText);
  brokenLinks.forEach((link) => {
    orphanData.push({
      id: link.id,
      type: "broken_link",
      description: `Broken or local link: ${String(link.url || "").substring(0, 50)}...`,
      value: link.url,
      location: link.location || "Document",
      severity: "low",
      suggestedAction: "Remove or update link",
    });
  });

  if (fileType === "pptx" && zip) {
    const slideFiles = Object.keys(zip.files).filter(
      (name) => name.startsWith("ppt/slides/slide") && name.endsWith(".xml")
    );

    for (const slideFile of slideFiles) {
      const file = zip.file(slideFile);
      if (file) {
        const content = await file.async("text");
        const textContent = extractTextFromXML(content);
        const slideMatch = slideFile.match(/slide(\d+)\.xml/);
        const slideNum = slideMatch ? slideMatch[1] : "?";

        if (textContent.trim().length < 10) {
          orphanData.push({
            id: `orphan_empty_slide${slideNum}_${Date.now()}`,
            type: "empty_page",
            description: `Slide ${slideNum} appears to be empty or has minimal content`,
            value: null,
            location: `Slide ${slideNum}`,
            severity: "low",
            suggestedAction: "Review or remove empty slide",
          });
        }
      }
    }
  }

  const trailingWhitespaceMatches = fullText.match(/\s{3,}/g) || [];
  if (trailingWhitespaceMatches.length > 5) {
    orphanData.push({
      id: `orphan_whitespace_${Date.now()}`,
      type: "trailing_whitespace",
      description: `${trailingWhitespaceMatches.length} instances of excessive whitespace detected`,
      value: null,
      location: "Throughout document",
      severity: "low",
      suggestedAction: "Clean up formatting",
    });
  }

  return orphanData;
}

// ============================================================
// EXCEL SPECIFIC
// ============================================================

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
          if (state === "hidden" || state === "veryHidden") {
            hiddenSheets.push({
              id: `hidden_sheet_${index}_${Date.now()}`,
              sheetName: sheet.$?.name || `Sheet ${index + 1}`,
              type: state === "veryHidden" ? "very_hidden" : "hidden",
              hasData: true,
              severity: "high",
            });
          }
        });
      }
    } catch {
      /* ignore */
    }
  }

  return hiddenSheets;
}

async function analyzeExcelHiddenColumns(zip) {
  const hiddenColumns = [];
  const sheetFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith("xl/worksheets/sheet") && name.endsWith(".xml")
  );

  for (const sheetFile of sheetFiles) {
    const file = zip.file(sheetFile);
    if (file) {
      const content = await file.async("text");
      const sheetName = sheetFile.replace("xl/worksheets/", "").replace(".xml", "");

      const colMatches = content.matchAll(
        /<col[^>]*hidden="1"[^>]*min="(\d+)"[^>]*max="(\d+)"[^>]*\/>/g
      );
      for (const match of colMatches) {
        hiddenColumns.push({
          id: `hidden_col_${hiddenColumns.length}_${Date.now()}`,
          sheet: sheetName,
          columns: `${match[1]}-${match[2]}`,
          type: "hidden_column",
          severity: "high",
        });
      }

      const rowMatches = content.matchAll(/<row[^>]*hidden="1"[^>]*r="(\d+)"[^>]*>/g);
      for (const match of rowMatches) {
        hiddenColumns.push({
          id: `hidden_row_${hiddenColumns.length}_${Date.now()}`,
          sheet: sheetName,
          row: match[1],
          type: "hidden_row",
          severity: "high",
        });
      }
    }
  }

  return hiddenColumns;
}

async function analyzeExcelSensitiveFormulas(zip) {
  const sensitiveFormulas = [];
  const sheetFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith("xl/worksheets/sheet") && name.endsWith(".xml")
  );

  for (const sheetFile of sheetFiles) {
    const file = zip.file(sheetFile);
    if (file) {
      const content = await file.async("text");
      const sheetName = sheetFile.replace("xl/worksheets/", "").replace(".xml", "");

      const formulaMatches = content.matchAll(/<f[^>]*>(.*?)<\/f>/gs);
      for (const match of formulaMatches) {
        const formula = match[1] || "";
        let risk = "low";
        let reason = "";

        if (formula.includes("[") && formula.includes("]")) {
          risk = "high";
          reason = "External file reference detected";
        } else if (/SQL|ODBC/i.test(formula)) {
          risk = "high";
          reason = "Database connection formula";
        } else if (/WEBSERVICE|FILTERXML/i.test(formula)) {
          risk = "high";
          reason = "External web query";
        } else if (formula.includes(":\\") || formula.includes("/Users/")) {
          risk = "medium";
          reason = "Local file path in formula";
        } else if (/INDIRECT|OFFSET/i.test(formula)) {
          risk = "low";
          reason = "Dynamic reference formula";
        }

        if (risk !== "low") {
          sensitiveFormulas.push({
            id: `formula_${sensitiveFormulas.length}_${Date.now()}`,
            sheet: sheetName,
            formula: formula.slice(0, 100),
            risk,
            reason,
            severity: risk,
          });
        }
      }
    }
  }

  return sensitiveFormulas;
}

// ============================================================
// EMBEDDED OBJECTS & MACROS
// ============================================================

async function analyzeEmbeddedObjects(zip, prefix) {
  const embeddings = Object.keys(zip.files).filter((name) => name.startsWith(`${prefix}/embeddings/`));
  return embeddings.map((name, index) => ({
    id: `embed_${index}_${Date.now()}`,
    filename: name.split("/").pop(),
    type: "embedded_object",
    path: name,
    severity: "medium",
  }));
}

async function analyzeMacros(zip, prefix) {
  const macros = [];
  const macroFiles = Object.keys(zip.files).filter(
    (name) => name.startsWith(`${prefix}/vbaProject`) || name.endsWith(".bin") || name.includes("vbaProject")
  );

  if (macroFiles.length > 0) {
    macros.push({
      id: `vba_macros_${Date.now()}`,
      type: "vba_macro",
      name: "VBA Macros",
      description: "Document contains executable VBA macro code - potential security risk",
      location: "VBA Project",
      severity: "critical",
      isMalicious: false,
      code: null,
    });
  }

  return macros;
}

// ============================================================
// BROKEN LINKS
// ============================================================

async function analyzeBrokenLinks(text) {
  const brokenLinks = [];

  const localLinkRegex = /file:\/\/[^\s<>"]+/gi;
  const localLinks = text.match(localLinkRegex) || [];
  localLinks.slice(0, 5).forEach((link, index) => {
    brokenLinks.push({
      id: `broken_local_${index}_${Date.now()}`,
      type: "local_file_link",
      url: link,
      location: "Document body",
      reason: "Local file link will not work for recipients",
      severity: "medium",
    });
  });

  const internalLinkRegex = /https?:\/\/[^\s<>"]*sharepoint\.com[^\s<>"]*/gi;
  const internalLinks = text.match(internalLinkRegex) || [];
  internalLinks.slice(0, 3).forEach((link, index) => {
    brokenLinks.push({
      id: `broken_sharepoint_${index}_${Date.now()}`,
      type: "internal_sharepoint",
      url: link.substring(0, 80) + "...",
      location: "Document body",
      reason: "SharePoint link may not be accessible to external recipients",
      severity: "low",
    });
  });

  return brokenLinks;
}

// ============================================================
// COMPLIANCE RISKS
// ============================================================

async function analyzeComplianceRisks(text) {
  const complianceRisks = [];
  const readableText = (text || "").toLowerCase();

  const gdprPatterns = [
    { pattern: /numéro de sécurité sociale|social security|ssn/i, risk: "Social Security Number reference", severity: "critical" },
    { pattern: /date de naissance|birth date|dob|née? le/i, risk: "Birth date detected", severity: "high" },
    { pattern: /passeport|passport/i, risk: "Passport reference", severity: "high" },
    { pattern: /carte d'identité|identity card|id card|cni/i, risk: "ID card reference", severity: "high" },
    { pattern: /numéro de permis|driver'?s? license|permis de conduire/i, risk: "Driver license reference", severity: "high" },
    { pattern: /données de santé|health data|medical record|dossier médical/i, risk: "Health data reference", severity: "critical" },
  ];

  gdprPatterns.forEach((p, index) => {
    if (p.pattern.test(readableText)) {
      complianceRisks.push({
        id: `gdpr_${index}_${Date.now()}`,
        type: "gdpr",
        description: p.risk,
        location: "Document body",
        severity: p.severity,
      });
    }
  });

  return complianceRisks;
}

// ============================================================
// ✅ NEW: Delivery & Commitment signals (deterministic)
// ============================================================

function detectDeliveryCommitmentSignals(text, fileType) {
  const t = String(text || "");
  if (!t.trim()) return [];

  const signals = [];

  // Conservative: only obvious commitments / open-ended scope / strong guarantees.
  const patterns = [
    {
      id: "DELIV_STRONG_COMMITMENT",
      re: /\b(we will|we'll|we commit|we guarantee|we guarantee that|we ensure|we will deliver|we deliver|we take full responsibility)\b/gi,
      severity: "high",
      reason: "Strong commitment language may create delivery/contract risk",
    },
    {
      id: "DELIV_OPEN_ENDED",
      re: /\b(asap|as soon as possible|at no extra cost|free of charge|unlimited|no limit|whenever needed|until completion|end-to-end)\b/gi,
      severity: "high",
      reason: "Open-ended or unlimited language may create scope exposure",
    },
    {
      id: "DELIV_TIMELINE_GUARANTEE",
      re: /\b(by end of (week|month|quarter)|within \d+\s?(days|weeks)|guaranteed timeline|hard deadline)\b/gi,
      severity: "medium",
      reason: "Timeline guarantees may reduce delivery control",
    },
    {
      id: "DELIV_SCOPE_AMBIGUITY",
      re: /\b(all inclusive|everything included|full scope|complete scope|any request|all requests)\b/gi,
      severity: "medium",
      reason: "Scope phrasing may be interpreted as broader than intended",
    },
  ];

  for (const p of patterns) {
    const matches = t.match(p.re);
    if (matches && matches.length) {
      signals.push({
        id: `biz_${p.id}_${Date.now()}`,
        type: "commitment_language",
        description: p.reason,
        location: `Text (${fileType})`,
        severity: p.severity,
        ruleId: p.id,
        count: Math.min(matches.length, 20),
      });
    }
  }

  return signals;
}

// ============================================================
// ✅ UNIVERSAL RISK OUTPUT (riskObjects + riskSummary)
// ============================================================

const EXEC_RISK_CATEGORIES = {
  margin: "Margin Exposure",
  delivery: "Delivery & Commitment Risk",
  negotiation: "Negotiation Leakage",
  credibility: "Professional Credibility Risk",
};

const SURFACES = {
  visible: "Visible",
  hidden: "Hidden",
  structural: "Structural",
  residual: "Residual",
};

function buildUniversalRiskOutput({ ext, detections, summary }) {
  const riskObjects = [];

  const pushFrom = (items, meta) => {
    if (!Array.isArray(items) || items.length === 0) return;
    for (const item of items) {
      const ro = mapDetectionToRiskObject(item, meta, ext);
      if (ro) riskObjects.push(ro);
    }
  };

  // ===== map detections -> risk objects (deterministic) =====

  // Visible
  pushFrom(detections.sensitiveData, { surface: SURFACES.visible, source: "sensitiveData" });
  pushFrom(detections.spellingErrors, { surface: SURFACES.visible, source: "spellingErrors" });
  pushFrom(detections.brokenLinks, { surface: SURFACES.visible, source: "brokenLinks" });
  pushFrom(detections.orphanData, { surface: SURFACES.visible, source: "orphanData" });
  pushFrom(detections.businessInconsistencies, { surface: SURFACES.visible, source: "businessInconsistencies" });

  // Hidden
  pushFrom(detections.comments, { surface: SURFACES.hidden, source: "comments" });
  pushFrom(detections.trackChanges, { surface: SURFACES.hidden, source: "trackChanges" });
  pushFrom(detections.hiddenContent, { surface: SURFACES.hidden, source: "hiddenContent" });
  pushFrom(detections.hiddenSheets, { surface: SURFACES.hidden, source: "hiddenSheets" });
  pushFrom(detections.hiddenColumns, { surface: SURFACES.hidden, source: "hiddenColumns" });
  pushFrom(detections.excelHiddenData, { surface: SURFACES.hidden, source: "excelHiddenData" });

  // Structural
  pushFrom(detections.sensitiveFormulas, { surface: SURFACES.structural, source: "sensitiveFormulas" });
  pushFrom(detections.embeddedObjects, { surface: SURFACES.structural, source: "embeddedObjects" });
  pushFrom(detections.macros, { surface: SURFACES.structural, source: "macros" });

  // Residual / metadata
  pushFrom(detections.metadata, { surface: SURFACES.residual, source: "metadata" });

  // ===== risk summary =====
  const riskSummary = buildRiskSummary(riskObjects);

  // Keep compatibility: do not override your existing summary riskScore.
  // But we can provide a universal "overallSeverity" + "clientReady".
  return { riskObjects, riskSummary };
}

function mapDetectionToRiskObject(item, meta, ext) {
  const id = item?.id || `risk_${meta.source}_${Date.now()}_${Math.random().toString(16).slice(2)}`;

  const { riskCategoryKey, reason, ruleId, fixability } = mapToExecutiveCategoryAndReason(item, meta);

  const points = scoreRiskPoints({
    surface: meta.surface,
    riskCategoryKey,
    source: meta.source,
    item,
  });

  const severity = pointsToSeverity(points, item?.severity);

  return {
    id,
    fileType: ext,
    surface: meta.surface, // Visible / Hidden / Structural / Residual
    riskCategory: riskCategoryKey, // margin / delivery / negotiation / credibility
    riskCategoryLabel: EXEC_RISK_CATEGORIES[riskCategoryKey] || "Professional Credibility Risk",
    severity, // Low / Medium / High / Critical
    points, // explainable numeric
    ruleId: ruleId || `${meta.source}`,

    // short neutral reason (no technical mechanics)
    reason,

    // optional metadata for UI grouping
    source: meta.source,

    // fixability: auto-fix / manual / not-fixable
    fixability,

    // minimal location (safe)
    location: item?.location || null,
  };
}

function mapToExecutiveCategoryAndReason(item, meta) {
  const src = meta.source;

  // Defaults
  let riskCategoryKey = "credibility";
  let ruleId = null;

  // Fixability defaults
  let fixability = "manual";

  // ===== CATEGORY MAPPING (deterministic) =====

  // Margin Exposure
  if (src === "sensitiveFormulas") {
    riskCategoryKey = "margin";
    ruleId = "MARGIN_FORMULA";
    fixability = "manual"; // flattening formulas may be possible later; keep manual here
  }
  if (src === "excelHiddenData") {
    // hidden_formula -> margin, hidden_sheet/row/col -> negotiation
    const t = String(item?.type || "");
    if (t === "hidden_formula") {
      riskCategoryKey = "margin";
      ruleId = "MARGIN_HIDDEN_FORMULA";
      fixability = "manual";
    } else {
      riskCategoryKey = "negotiation";
      ruleId = "NEGOTIATION_HIDDEN_EXCEL_DATA";
      fixability = "auto-fix";
    }
  }

  // Delivery & Commitment Risk
  if (src === "businessInconsistencies") {
    riskCategoryKey = "delivery";
    ruleId = item?.ruleId || "DELIVERY_COMMITMENT";
    fixability = "manual";
  }

  // Negotiation Leakage
  if (src === "metadata") {
    riskCategoryKey = "negotiation";
    ruleId = "NEGOTIATION_METADATA";
    fixability = "auto-fix";
  }
  if (src === "hiddenSheets" || src === "hiddenColumns") {
    riskCategoryKey = "negotiation";
    ruleId = "NEGOTIATION_HIDDEN_EXCEL_STRUCTURE";
    fixability = "auto-fix";
  }
  if (src === "embeddedObjects") {
    riskCategoryKey = "negotiation";
    ruleId = "NEGOTIATION_EMBEDDED_OBJECT";
    fixability = "auto-fix";
  }

  // Professional Credibility
  if (src === "comments" || src === "trackChanges" || src === "hiddenContent") {
    riskCategoryKey = "credibility";
    ruleId = src === "comments" ? "CREDIBILITY_COMMENTS" : src === "trackChanges" ? "CREDIBILITY_TRACK_CHANGES" : "CREDIBILITY_HIDDEN_CONTENT";
    fixability = "auto-fix";
  }
  if (src === "spellingErrors") {
    riskCategoryKey = "credibility";
    ruleId = "CREDIBILITY_SPELLING";
    fixability = "manual"; // correction exists but is AI-based; keep manual by default
  }
  if (src === "brokenLinks" || src === "orphanData") {
    riskCategoryKey = "credibility";
    ruleId = "CREDIBILITY_ORPHAN_OR_LINK";
    fixability = "manual";
  }
  if (src === "macros") {
    riskCategoryKey = "credibility";
    ruleId = "CREDIBILITY_MACROS";
    fixability = "auto-fix";
  }

  // SensitiveData: can be margin or negotiation depending on type/category
  if (src === "sensitiveData") {
    const t = String(item?.type || "");
    const cat = String(item?.category || "");
    if (t === "price" || cat === "financial") {
      riskCategoryKey = "margin";
      ruleId = "MARGIN_SENSITIVE_VALUE";
    } else if (cat === "internal" || t === "project_code" || t === "internal_url" || t === "file_path") {
      riskCategoryKey = "negotiation";
      ruleId = "NEGOTIATION_INTERNAL_REFERENCE";
    } else {
      riskCategoryKey = "credibility";
      ruleId = "CREDIBILITY_SENSITIVE_DATA";
    }
    fixability = "manual"; // you can auto-remove in cleaner, but executive gate should default to review
  }

  // Compliance risks: map to credibility (exec view)
  if (src === "complianceRisks") {
    riskCategoryKey = "credibility";
    ruleId = "CREDIBILITY_COMPLIANCE";
    fixability = "manual";
  }

  const reason = buildNeutralReason(riskCategoryKey, meta.surface, src, item);

  return { riskCategoryKey, reason, ruleId, fixability };
}

function buildNeutralReason(riskCategoryKey, surface, source, item) {
  // Neutral, non-technical, executive phrasing
  const catLabel = EXEC_RISK_CATEGORIES[riskCategoryKey] || "Professional Credibility Risk";

  if (surface === SURFACES.hidden) {
    return `Hidden content detected that may impact ${catLabel.toLowerCase()}.`;
  }
  if (surface === SURFACES.structural) {
    return `Derived or linked content detected that may impact ${catLabel.toLowerCase()}.`;
  }
  if (surface === SURFACES.residual) {
    return `Metadata detected that may reveal internal information (impacting ${catLabel.toLowerCase()}).`;
  }
  // visible
  if (source === "businessInconsistencies") {
    return `Commitment language detected that may reduce delivery control.`;
  }
  return `Client-visible content detected that may impact ${catLabel.toLowerCase()}.`;
}

function scoreRiskPoints({ surface, riskCategoryKey, source, item }) {
  // Spec scoring (simple & explainable):
  // Hidden but accessible: +3
  // Generated / derived: +2
  // Cross-object dependency: +2
  // Metadata revealing intent: +2
  // Financial proximity: +2

  let pts = 0;

  // surface-based
  if (surface === SURFACES.hidden) pts += 3;
  if (surface === SURFACES.structural) pts += 2;
  if (surface === SURFACES.residual) pts += 2;

  // cross-object dependency (structural / embedded / external refs)
  if (source === "embeddedObjects") pts += 2;
  if (source === "sensitiveFormulas") pts += 2;
  if (source === "excelHiddenData" && String(item?.type || "") === "hidden_formula") pts += 2;

  // metadata intent / identity
  if (source === "metadata") pts += 2;

  // financial proximity
  if (riskCategoryKey === "margin") pts += 2;

  // macros are always a big credibility hit
  if (source === "macros") pts += 3;

  return pts;
}

function pointsToSeverity(points, originalSeverity) {
  // If originalSeverity is "critical", keep it critical.
  if (String(originalSeverity || "").toLowerCase() === "critical") return "Critical";

  if (points >= 9) return "Critical";
  if (points >= 7) return "High";
  if (points >= 4) return "Medium";
  return "Low";
}

function buildRiskSummary(riskObjects) {
  const byCategory = { margin: 0, delivery: 0, negotiation: 0, credibility: 0 };
  const maxSeverity = { level: "Low", score: 1 };

  const severityScore = (s) => {
    const v = String(s || "Low").toLowerCase();
    if (v === "critical") return 4;
    if (v === "high") return 3;
    if (v === "medium") return 2;
    return 1;
  };

  for (const ro of riskObjects) {
    const k = ro.riskCategory;
    if (byCategory[k] !== undefined) byCategory[k]++;

    const s = severityScore(ro.severity);
    if (s > maxSeverity.score) {
      maxSeverity.score = s;
      maxSeverity.level =
        s === 4 ? "Critical" : s === 3 ? "High" : s === 2 ? "Medium" : "Low";
    }
  }

  // clientReady rule (strict):
  // Any High/Critical => NO
  const hasBlocking = riskObjects.some((r) => ["High", "Critical"].includes(r.severity));
  const clientReady = !hasBlocking;

  const executiveSignals = [
    byCategory.margin > 0 ? EXEC_RISK_CATEGORIES.margin : null,
    byCategory.delivery > 0 ? EXEC_RISK_CATEGORIES.delivery : null,
    byCategory.negotiation > 0 ? EXEC_RISK_CATEGORIES.negotiation : null,
    byCategory.credibility > 0 ? EXEC_RISK_CATEGORIES.credibility : null,
  ].filter(Boolean);

  // Blocking issues list (short, neutral)
  const blockingIssues = riskObjects
    .filter((r) => ["High", "Critical"].includes(r.severity))
    .slice(0, 12)
    .map((r) => ({
      id: r.id,
      severity: r.severity,
      riskCategory: r.riskCategory,
      label: r.riskCategoryLabel,
      reason: r.reason,
      fixability: r.fixability,
    }));

  return {
    clientReady: clientReady ? "YES" : "NO",
    overallSeverity: maxSeverity.level,
    executiveSignals,
    byCategory,
    blockingIssues,
    totalRiskObjects: riskObjects.length,
  };
}

// ============================================================
// TEXT EXTRACTION (TOKEN-BASED) + SAFE SPACING
// ============================================================

function decodeXmlEntities(s) {
  return String(s || "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#39;/g, "'")
    .replace(/&#34;/g, '"');
}

function shouldInsertSpace(prevText, nextToken) {
  if (!prevText || !nextToken) return false;

  const prev = String(prevText);
  const next = String(nextToken);

  const prevTrim = prev.replace(/\s+$/g, "");
  if (!prevTrim) return false;

  const p = prevTrim[prevTrim.length - 1];
  const n = next[0];

  if (/\s/.test(p) || /\s/.test(n)) return false;

  const isLetter = (ch) => /[A-Za-zÀ-ÖØ-öø-ÿ]/.test(ch);
  const isDigit = (ch) => /[0-9]/.test(ch);

  const hardNoSpaceBefore = /[,\.;:\)\]\}]/;
  const hardNoSpaceAfter = /[\(\[\{]/;

  if (hardNoSpaceBefore.test(n)) return false;
  if (hardNoSpaceAfter.test(p)) return false;

  if (isDigit(p) && next.startsWith(".")) return false;
  if (p === "." || n === ".") return false;

  if (isLetter(p) && isLetter(n)) return true;
  if (isLetter(p) && isDigit(n)) return true;
  if (isDigit(p) && isLetter(n)) return true;

  return false;
}

function extractTextFromXML(xml) {
  const raw = String(xml || "");
  if (!raw.trim()) return "";

  let x = raw
    .replace(/<w:tab\b[^\/>]*\/>/gi, " \t ")
    .replace(/<w:br\b[^\/>]*\/>/gi, " \n ")
    .replace(/<\/w:p>/gi, " \n ")
    .replace(/<\/w:tr>/gi, " \n ")
    .replace(/<\/w:tc>/gi, " \t ");

  x = x
    .replace(/<a:br\b[^\/>]*\/>/gi, " \n ")
    .replace(/<\/a:p>/gi, " \n ")
    .replace(/<\/p:sp>/gi, " \n ");

  x = x.replace(/<\/row>/gi, " \n ").replace(/<\/c>/gi, " \t ");

  const tokens = [];
  const re = />([^<]+)</g;
  let m;
  while ((m = re.exec(x)) !== null) {
    let t = decodeXmlEntities(m[1]);
    if (!t) continue;

    t = t.replace(/\r/g, "");

    const parts = t.split(/(\n|\t)/g);
    for (const part of parts) {
      if (!part) continue;
      if (part === "\n" || part === "\t") {
        tokens.push(part);
        continue;
      }
      const cleaned = part.replace(/\s+/g, " ").trim();
      if (cleaned) tokens.push(cleaned);
    }
  }

  if (!tokens.length) return "";

  let out = "";

  const appendToken = (tok) => {
    if (!tok) return;
    if (tok === "\n") {
      out = out.replace(/[ \t]+$/g, "");
      out += "\n";
      return;
    }
    if (tok === "\t") {
      if (out && !out.endsWith("\n") && !out.endsWith("\t")) out += "\t";
      else out += "\t";
      return;
    }

    if (!out) {
      out = tok;
      return;
    }

    const needsSpace = shouldInsertSpace(out, tok);
    out += (needsSpace ? " " : "") + tok;
  };

  for (const tok of tokens) appendToken(tok);

  out = out
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .replace(/[ \t]{2,}/g, " ")
    .trim();

  return out;
}

function determineSeverity(text) {
  const lower = (text || "").toLowerCase();
  if (/confidentiel|confidential|secret|password|mot de passe|urgent|critical|do not share|ne pas partager/i.test(lower))
    return "high";
  if (/todo|fixme|draft|preliminary|internal|review|brouillon|à revoir/i.test(lower)) return "medium";
  return "low";
}

// ============================================================
// SUMMARY CALCULATION (kept as-is, no breaking changes)
// ============================================================

export function calculateSummary(detections) {
  let totalIssues = 0;
  let criticalIssues = 0;
  let highIssues = 0;
  let mediumIssues = 0;
  let lowIssues = 0;

  const categoryCounts = {
    sensitiveData: 0,
    metadata: 0,
    comments: 0,
    hiddenContent: 0,
    spellingErrors: 0,
    visualObjects: 0,
    orphanData: 0,
    macros: 0,
    excelHiddenData: 0,
    trackChanges: 0,
    embeddedObjects: 0,
    brokenLinks: 0,
    complianceRisks: 0,
    businessInconsistencies: 0,
  };

  const countBySeverity = (items, categoryKey) => {
    if (!items || !Array.isArray(items)) return;

    items.forEach((item) => {
      totalIssues++;
      categoryCounts[categoryKey]++;

      const severity = item.severity || "low";
      switch (severity) {
        case "critical":
          criticalIssues++;
          break;
        case "high":
          highIssues++;
          break;
        case "medium":
          mediumIssues++;
          break;
        default:
          lowIssues++;
      }
    });
  };

  countBySeverity(detections.sensitiveData, "sensitiveData");
  countBySeverity(detections.metadata, "metadata");
  countBySeverity(detections.comments, "comments");
  countBySeverity(detections.hiddenContent, "hiddenContent");
  countBySeverity(detections.spellingErrors, "spellingErrors");
  countBySeverity(detections.visualObjects, "visualObjects");
  countBySeverity(detections.orphanData, "orphanData");
  countBySeverity(detections.macros, "macros");
  countBySeverity(detections.excelHiddenData, "excelHiddenData");
  countBySeverity(detections.trackChanges, "trackChanges");
  countBySeverity(detections.embeddedObjects, "embeddedObjects");
  countBySeverity(detections.brokenLinks, "brokenLinks");
  countBySeverity(detections.complianceRisks, "complianceRisks");
  countBySeverity(detections.businessInconsistencies, "businessInconsistencies");

  let riskScore = 100;

  const categoryPenalties = {
    sensitiveData: { perItem: 20, maxPenalty: 50 },
    macros: { perItem: 30, maxPenalty: 30 },
    hiddenContent: { perItem: 10, maxPenalty: 30 },
    comments: { perItem: 3, maxPenalty: 15 },
    trackChanges: { perItem: 3, maxPenalty: 15 },
    metadata: { perItem: 2, maxPenalty: 10 },
    spellingErrors: { perItem: 1, maxPenalty: 10 },
    complianceRisks: { perItem: 25, maxPenalty: 50 },
    excelHiddenData: { perItem: 10, maxPenalty: 30 },
    visualObjects: { perItem: 5, maxPenalty: 15 },
    orphanData: { perItem: 2, maxPenalty: 10 },
    embeddedObjects: { perItem: 5, maxPenalty: 15 },
    brokenLinks: { perItem: 2, maxPenalty: 10 },
    businessInconsistencies: { perItem: 8, maxPenalty: 24 },
  };

  Object.keys(categoryPenalties).forEach((category) => {
    const count = categoryCounts[category] || 0;
    const config = categoryPenalties[category];
    const penalty = Math.min(count * config.perItem, config.maxPenalty);
    riskScore -= penalty;
  });

  riskScore -= criticalIssues * 10;
  riskScore -= highIssues * 5;

  if (totalIssues > 20) riskScore -= 10;
  if (totalIssues > 50) riskScore -= 15;

  riskScore = Math.max(0, Math.min(100, Math.round(riskScore)));

  let riskLevel;
  let riskColor;
  if (riskScore >= 90) {
    riskLevel = "safe";
    riskColor = "green";
  } else if (riskScore >= 70) {
    riskLevel = "low";
    riskColor = "lime";
  } else if (riskScore >= 50) {
    riskLevel = "medium";
    riskColor = "yellow";
  } else if (riskScore >= 25) {
    riskLevel = "high";
    riskColor = "orange";
  } else {
    riskLevel = "critical";
    riskColor = "red";
  }

  return {
    totalIssues,
    criticalIssues,
    highIssues,
    mediumIssues,
    lowIssues,
    riskScore,
    riskLevel,
    riskColor,
    categoryCounts,
    recommendations: [],
    hasCriticalIssues: criticalIssues > 0,
    hasSensitiveData: categoryCounts.sensitiveData > 0,
    hasMacros: categoryCounts.macros > 0,
    hasHiddenContent: categoryCounts.hiddenContent > 0 || categoryCounts.excelHiddenData > 0,
    hasComments: categoryCounts.comments > 0,
    needsReview: riskScore < 70,
  };
}

export default {
  analyzeDocument,
  calculateSummary,
};
