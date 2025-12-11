// lib/sensitiveData.js
// VERSION 3.2 — ENTERPRISE SENSITIVE DATA DETECTOR
// Compatible documentAnalyzer.js & Render (pure ESM)

// ============================================================================
// PATTERNS
// ============================================================================

const PATTERNS = {
  email: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/gi,
  phone: /(?:\+?\d{1,4}[\s.-]?)?(?:\(?\d{1,4}\)?[\s.-]?)?\d{1,4}[\s.-]?\d{1,4}[\s.-]?\d{1,9}/g,
  priceEuro: /(\d{1,3}(?:[.,\s]\d{3})*(?:[.,]\d{1,2})?)\s*[€$£]/gi,
  priceSymbolFirst: /[€$£]\s*(\d{1,3}(?:[.,\s]\d{3})*(?:[.,]\d{1,2})?)/gi,
  priceWithCurrency: /(\d{1,3}(?:[.,\s]\d{3})*(?:[.,]\d{1,2})?)\s*(EUR|USD|GBP|CHF|euros?|dollars?)/gi,
  iban: /\b[A-Z]{2}\d{2}[A-Z0-9]{4}\d{7,18}\b/gi,
  ssnFR: /\b[12]\s?\d{2}\s?\d{2}\s?\d{2}\s?\d{3}\s?\d{3}\s?\d{2}\b/g,
  projectCode: /\b(?:PRJ|PROJ|PROJECT|REF|CASE|TICKET|INC|CHG|REQ)[-_ ]?\d{3,12}\b/gi,
  windowsPath: /[A-Z]:\\(?:[^\\/:*?"<>|\r\n]+\\)*[^\\/:*?"<>|\r\n]*/gi,
  uncPath: /\\\\[a-zA-Z0-9._-]+\\[a-zA-Z0-9._$-]+(?:\\[^\\/:*?"<>|\r\n]+)*/gi,
  unixPath: /\/(?:home|var|usr|etc|opt|tmp|mnt|srv|data|backup|Users|Volumes)\/[^\s"'<>]+/gi,
  ipv4: /\b(?:(?:25[0-5]|2[0-4]\d|[01]?\d\d)\.){3}(?:25[0-5]|2[0-4]\d|[01]?\d\d)\b/g,
  creditCard: /\b(?:\d{4}[ -]?){3}\d{4}\b/g,
  internalUrl: /\bhttps?:\/\/(?:intranet|internal|dev|test|staging|local|localhost|192\.168|10\.|172\.)[^\s"'<>\]]*/gi,
};

// ============================================================================
// CONFIDENTIAL KEYWORDS
// ============================================================================
const CONFIDENTIAL_KEYWORDS = [
  "CONFIDENTIEL", "STRICTEMENT CONFIDENTIEL", "NE PAS DIFFUSER",
  "USAGE INTERNE", "INTERNE UNIQUEMENT", "RÉSERVÉ", "SECRET",
  "DIFFUSION RESTREINTE", "NON DIVULGABLE", "PRIVÉ",
  "CONFIDENTIAL", "STRICTLY CONFIDENTIAL", "DO NOT DISTRIBUTE",
  "INTERNAL USE ONLY", "RESTRICTED", "PRIVATE", "PROPRIETARY"
];

// ============================================================================
// HELPERS
// ============================================================================

function getContext(text, start, end, span = 50) {
  const before = text.substring(Math.max(0, start - span), start);
  const match = text.substring(start, end);
  const after = text.substring(end, Math.min(text.length, end + span));

  return {
    before: before.trim(),
    match,
    after: after.trim(),
    full: `...${before}${match}${after}...`
  };
}

function isValidPhone(phone) {
  const digits = phone.replace(/[^\d]/g, "");
  return digits.length >= 8 && digits.length <= 15;
}

function isValidCreditCard(number) {
  const digits = number.replace(/[ -]/g, "");
  if (digits.length !== 16) return false;

  let sum = 0, alt = false;
  for (let i = digits.length - 1; i >= 0; i--) {
    let n = parseInt(digits[i]);
    if (alt) { n *= 2; if (n > 9) n -= 9; }
    sum += n;
    alt = !alt;
  }
  return sum % 10 === 0;
}

function maskValue(value, type) {
  switch (type) {
    case "email": return value.replace(/(.).+(@.+)/, "$1***$2");
    case "phone": return value.replace(/\d(?=\d{2})/g, "*");
    case "credit_card": return "**** **** **** " + value.slice(-4);
    case "iban": return value.slice(0, 4) + " **** **** **** " + value.slice(-4);
    default: return value.slice(0, 3) + "***";
  }
}

function addFinding(arr, type, match, index, severity, text, recommendation) {
  arr.push({
    type,
    value: match,
    masked: maskValue(match, type),
    context: getContext(text, index, index + match.length),
    location: index,
    severity,
    recommendation
  });
}

// ============================================================================
// DETECTORS
// ============================================================================

function detectEmails(t, out) {
  PATTERNS.email.lastIndex = 0;
  let m;
  while ((m = PATTERNS.email.exec(t)) !== null) {
    addFinding(out, "email", m[0], m.index, "high", t, "Remove or anonymize email addresses.");
  }
}

function detectPhones(t, out) {
  PATTERNS.phone.lastIndex = 0;
  let m;
  while ((m = PATTERNS.phone.exec(t)) !== null) {
    if (!isValidPhone(m[0])) continue;
    addFinding(out, "phone", m[0], m.index, "high", t, "Phone numbers must not appear in shared documents.");
  }
}

function detectPrices(t, out) {
  const pricePatterns = [
    PATTERNS.priceEuro,
    PATTERNS.priceSymbolFirst,
    PATTERNS.priceWithCurrency,
  ];
  for (const p of pricePatterns) {
    p.lastIndex = 0;
    let m;
    while ((m = p.exec(t)) !== null) {
      addFinding(out, "price", m[0], m.index, "medium", t, "Internal pricing should be removed.");
    }
  }
}

function detectIBANs(t, out) {
  PATTERNS.iban.lastIndex = 0;
  let m;
  while ((m = PATTERNS.iban.exec(t)) !== null) {
    addFinding(out, "iban", m[0], m.index, "critical", t, "IBAN numbers must not be shared.");
  }
}

function detectSSN(t, out) {
  PATTERNS.ssnFR.lastIndex = 0;
  let m;
  while ((m = PATTERNS.ssnFR.exec(t)) !== null) {
    addFinding(out, "ssn", m[0], m.index, "critical", t, "Social security numbers are strictly confidential.");
  }
}

function detectProjectCodes(t, out) {
  PATTERNS.projectCode.lastIndex = 0;
  let m;
  while ((m = PATTERNS.projectCode.exec(t)) !== null) {
    addFinding(out, "project_code", m[0], m.index, "medium", t, "Project codes reveal internal structure.");
  }
}

function detectPaths(t, out) {
  const paths = [
    { regex: PATTERNS.windowsPath, type: "file_path" },
    { regex: PATTERNS.uncPath, type: "file_path" },
    { regex: PATTERNS.unixPath, type: "file_path" },
  ];
  for (const p of paths) {
    p.regex.lastIndex = 0;
    let m;
    while ((m = p.regex.exec(t)) !== null) {
      addFinding(out, "file_path", m[0], m.index, "high", t, "File paths can leak internal server architecture.");
    }
  }
}

function detectIPs(t, out) {
  PATTERNS.ipv4.lastIndex = 0;
  let m;
  while ((m = PATTERNS.ipv4.exec(t)) !== null) {
    addFinding(out, "ip_address", m[0], m.index, "high", t, "Internal IP addresses should be hidden.");
  }
}

function detectConfidential(t, out) {
  const upper = t.toUpperCase();
  for (const kw of CONFIDENTIAL_KEYWORDS) {
    let idx = upper.indexOf(kw);
    while (idx !== -1) {
      addFinding(out, "confidential_keyword", kw, idx, "high", t, "Confidentiality markers detected.");
      idx = upper.indexOf(kw, idx + kw.length);
    }
  }
}

function detectCreditCards(t, out) {
  PATTERNS.creditCard.lastIndex = 0;
  let m;
  while ((m = PATTERNS.creditCard.exec(t)) !== null) {
    if (!isValidCreditCard(m[0])) continue;
    addFinding(out, "credit_card", m[0], m.index, "critical", t, "Credit card numbers must be removed immediately.");
  }
}

function detectInternalURLs(t, out) {
  PATTERNS.internalUrl.lastIndex = 0;
  let m;
  while ((m = PATTERNS.internalUrl.exec(t)) !== null) {
    addFinding(out, "internal_url", m[0], m.index, "medium", t, "Internal URLs should not be shared externally.");
  }
}

// ============================================================================
// MAIN EXPORTED FUNCTION
// ============================================================================

export function detectSensitiveData(text = "") {
  if (!text || typeof text !== "string") {
    return { findings: [] };
  }

  const findings = [];

  detectEmails(text, findings);
  detectPhones(text, findings);
  detectPrices(text, findings);
  detectIBANs(text, findings);
  detectSSN(text, findings);
  detectProjectCodes(text, findings);
  detectPaths(text, findings);
  detectIPs(text, findings);
  detectConfidential(text, findings);
  detectCreditCards(text, findings);
  detectInternalURLs(text, findings);

  findings.sort((a, b) => a.location - b.location);

  return { findings };
}

// ============================================================================
// NAMED EXPORTS (OPTIONAL)
// ============================================================================

export {
  PATTERNS,
  CONFIDENTIAL_KEYWORDS,
  maskValue,
  getContext
};
