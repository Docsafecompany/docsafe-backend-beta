// lib/sensitiveData.js
// VERSION 2.0 - Détection avancée des données sensibles (ESM VERSION)
// Compatible Render & import { detectSensitiveData }

// ============================================================================
// PATTERNS - Regex pour chaque type de données sensibles
// ============================================================================

const PATTERNS = {
  email: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/gi,
  phone: /(?:\+?\d{1,4}[\s.-]?)?(?:\(?\d{1,4}\)?[\s.-]?)?\d{1,4}[\s.-]?\d{1,4}[\s.-]?\d{1,9}/g,
  priceEuro: /(?:(?:prix|tarif|coût|budget|montant|total|facture|devis)\s*[:=]?\s*)?(\d{1,3}(?:[\s.,]\d{3})*(?:[.,]\d{1,2})?)\s*[€$£¥CHF]{1,3}/gi,
  priceSymbolFirst: /[€$£¥]\s*(\d{1,3}(?:[,.\s]\d{3})*(?:[.,]\d{1,2})?)/gi,
  priceWithCurrency: /(\d{1,3}(?:[\s.,]\d{3})*(?:[.,]\d{1,2})?)\s*(?:euros?|dollars?|USD|EUR|GBP|CHF)/gi,
  iban: /[A-Z]{2}\d{2}[\s]?(?:\d{4}[\s]?){2,7}\d{1,4}/gi,
  ssnFR: /[12]\s?\d{2}\s?\d{2}\s?\d{2}\s?\d{3}\s?\d{3}\s?\d{2}/g,
  projectCode: /(?:PRJ|PROJ|PROJECT|REF|DOSSIER|CASE|TICKET|INC|CHG|REQ)[-_]?\d{2,4}[-_]?\d{2,6}/gi,
  windowsPath: /[A-Z]:\\(?:[^\\/:*?"<>|\r\n]+\\)*[^\\/:*?"<>|\r\n]*/gi,
  uncPath: /\\\\[a-zA-Z0-9._-]+\\[a-zA-Z0-9._$-]+(?:\\[^\\/:*?"<>|\r\n]+)*/gi,
  unixPath: /(?:\/(?:home|var|usr|etc|opt|tmp|mnt|srv|data|backup|Users|Volumes)\/[^\s:*?"<>|]+)/gi,
  ipv4: /(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)/g,
  creditCard: /(?:\d{4}[\s-]?){3}\d{4}/g,
  internalUrl: /https?:\/\/(?:intranet|internal|dev|staging|preprod|test|localhost|192\.168|10\.0|172\.(?:1[6-9]|2[0-9]|3[01]))[^\s<>"']*/gi
};

// ============================================================================
// CONFIDENTIAL KEYWORDS
// ============================================================================

const CONFIDENTIAL_KEYWORDS = [
  'CONFIDENTIEL', 'STRICTEMENT CONFIDENTIEL', 'NE PAS DIFFUSER',
  'USAGE INTERNE', 'INTERNE UNIQUEMENT', 'RÉSERVÉ', 'SECRET',
  'DIFFUSION RESTREINTE', 'NON DIVULGABLE', 'PRIVÉ',
  'CONFIDENTIAL', 'STRICTLY CONFIDENTIAL', 'DO NOT DISTRIBUTE',
  'INTERNAL USE ONLY', 'RESTRICTED', 'PRIVATE', 'PROPRIETARY',
  'VERTRAULICH', 'GEHEIM', 'CONFIDENCIAL'
];

// ============================================================================
// PRICE CONTEXT KEYWORDS
// ============================================================================

const PRICE_CONTEXT_KEYWORDS = [
  'prix', 'tarif', 'coût', 'budget', 'montant', 'total', 'facture',
  'devis', 'honoraires', 'taux', 'tjm', 'remise', 'réduction', 'marge',
  'price', 'cost', 'budget', 'amount', 'invoice', 'quote', 'fee',
  'rate', 'discount', 'margin', 'revenue'
];

// ============================================================================
// HELPERS
// ============================================================================

function getContext(text, startIndex, endIndex, contextLength = 50) {
  const beforeStart = Math.max(0, startIndex - contextLength);
  const afterEnd = Math.min(text.length, endIndex + contextLength);

  const before = text.substring(beforeStart, startIndex);
  const match = text.substring(startIndex, endIndex);
  const after = text.substring(endIndex, afterEnd);

  return {
    before: before.replace(/[\r\n]+/g, ' ').trim(),
    match,
    after: after.replace(/[\r\n]+/g, ' ').trim(),
    full: `...${before}[${match}]${after}...`
  };
}

function isValidPhone(phone) {
  const digits = phone.replace(/[\s.-]/g, '');
  if (digits.length < 8 || digits.length > 15) return false;
  if (/^(19|20)\d{2}$/.test(digits)) return false;
  return true;
}

function isValidCreditCard(number) {
  const digits = number.replace(/[\s-]/g, '');
  if (digits.length < 13 || digits.length > 19) return false;

  let sum = 0, isEven = false;
  for (let i = digits.length - 1; i >= 0; i--) {
    let d = parseInt(digits[i], 10);
    if (isEven) { d *= 2; if (d > 9) d -= 9; }
    sum += d;
    isEven = !isEven;
  }
  return sum % 10 === 0;
}

function maskValue(value, type) {
  switch (type) {
    case 'email': return value.slice(0, 2) + "***@" + value.split("@")[1];
    case 'phone': return value.slice(0, 4) + "****" + value.slice(-2);
    case 'credit_card': return "****-****-****-" + value.slice(-4);
    case 'iban': return value.slice(0, 4) + " **** **** " + value.slice(-4);
    default: return value.slice(0, 3) + "***";
  }
}

// ============================================================================
// DETECTION FUNCTIONS
// ============================================================================

function detectEmails(text) {
  const results = [];
  let match;
  PATTERNS.email.lastIndex = 0;

  while ((match = PATTERNS.email.exec(text)) !== null) {
    const ctx = getContext(text, match.index, match.index + match[0].length);
    results.push({
      type: "email",
      value: match[0],
      masked: maskValue(match[0], "email"),
      context: ctx,
      severity: "high",
      location: match.index
    });
  }
  return results;
}

function detectPhones(text) {
  const results = [];
  let m;
  PATTERNS.phone.lastIndex = 0;

  while ((m = PATTERNS.phone.exec(text)) !== null) {
    if (!isValidPhone(m[0])) continue;
    const ctx = getContext(text, m.index, m.index + m[0].length);

    results.push({
      type: "phone",
      value: m[0],
      masked: maskValue(m[0], "phone"),
      context: ctx,
      severity: "high",
      location: m.index
    });
  }
  return results;
}

function detectPrices(text) {
  const results = [];
  const patterns = [
    PATTERNS.priceEuro,
    PATTERNS.priceSymbolFirst,
    PATTERNS.priceWithCurrency
  ];

  for (const regex of patterns) {
    regex.lastIndex = 0;
    let m;
    while ((m = regex.exec(text)) !== null) {
      const ctx = getContext(text, m.index, m.index + m[0].length);
      results.push({
        type: "price",
        value: m[0],
        context: ctx,
        severity: "medium",
        location: m.index
      });
    }
  }
  return results;
}

function detectIBANs(text) {
  const results = [];
  let m;
  PATTERNS.iban.lastIndex = 0;

  while ((m = PATTERNS.iban.exec(text)) !== null) {
    const ctx = getContext(text, m.index, m.index + m[0].length);
    results.push({
      type: "iban",
      value: m[0],
      masked: maskValue(m[0], "iban"),
      severity: "critical",
      context: ctx,
      location: m.index
    });
  }
  return results;
}

function detectSSN(text) {
  const results = [];
  let m;
  PATTERNS.ssnFR.lastIndex = 0;

  while ((m = PATTERNS.ssnFR.exec(text)) !== null) {
    const ctx = getContext(text, m.index, m.index + m[0].length);

    results.push({
      type: "ssn",
      value: m[0],
      masked: maskValue(m[0], "ssn"),
      severity: "critical",
      context: ctx,
      location: m.index
    });
  }
  return results;
}

function detectProjectCodes(text) {
  const results = [];
  PATTERNS.projectCode.lastIndex = 0;

  let m;
  while ((m = PATTERNS.projectCode.exec(text)) !== null) {
    const ctx = getContext(text, m.index, m.index + m[0].length);
    results.push({
      type: "project_code",
      value: m[0],
      context: ctx,
      severity: "medium",
      location: m.index
    });
  }
  return results;
}

function detectFilePaths(text) {
  const results = [];
  
  const patterns = [
    { regex: PATTERNS.windowsPath, type: "file_path" },
    { regex: PATTERNS.uncPath, type: "file_path" },
    { regex: PATTERNS.unixPath, type: "file_path" }
  ];

  for (const { regex, type } of patterns) {
    regex.lastIndex = 0;
    let m;
    while ((m = regex.exec(text)) !== null) {
      const ctx = getContext(text, m.index, m.index + m[0].length);
      results.push({
        type,
        value: m[0],
        context: ctx,
        severity: "high",
        location: m.index
      });
    }
  }
  return results;
}

function detectIPAddresses(text) {
  const results = [];
  let m;
  PATTERNS.ipv4.lastIndex = 0;

  while ((m = PATTERNS.ipv4.exec(text)) !== null) {
    const ctx = getContext(text, m.index, m.index + m[0].length);
    results.push({
      type: "ip_address",
      value: m[0],
      severity: "high",
      context: ctx,
      location: m.index
    });
  }
  return results;
}

function detectConfidentialKeywords(text) {
  const results = [];
  const upper = text.toUpperCase();

  CONFIDENTIAL_KEYWORDS.forEach(keyword => {
    let idx = upper.indexOf(keyword);
    while (idx !== -1) {
      const ctx = getContext(text, idx, idx + keyword.length);
      results.push({
        type: "confidential_keyword",
        value: keyword,
        severity: "high",
        context: ctx,
        location: idx
      });
      idx = upper.indexOf(keyword, idx + keyword.length);
    }
  });

  return results;
}

function detectCreditCards(text) {
  const results = [];
  let m;

  PATTERNS.creditCard.lastIndex = 0;
  while ((m = PATTERNS.creditCard.exec(text)) !== null) {
    if (!isValidCreditCard(m[0])) continue;

    const ctx = getContext(text, m.index, m.index + m[0].length);
    results.push({
      type: "credit_card",
      value: maskValue(m[0], "credit_card"),
      severity: "critical",
      context: ctx,
      location: m.index
    });
  }
  return results;
}

function detectInternalURLs(text) {
  const results = [];
  let m;

  PATTERNS.internalUrl.lastIndex = 0;
  while ((m = PATTERNS.internalUrl.exec(text)) !== null) {
    const ctx = getContext(text, m.index, m.index + m[0].length);
    results.push({
      type: "internal_url",
      value: m[0],
      context: ctx,
      severity: "high",
      location: m.index
    });
  }
  return results;
}

// ============================================================================
// MAIN ENTRYPOINT
// ============================================================================

function detectSensitiveData(text) {
  if (!text || typeof text !== "string") {
    return { findings: [], summary: { total: 0 } };
  }

  const findings = [
    ...detectEmails(text),
    ...detectPhones(text),
    ...detectPrices(text),
    ...detectIBANs(text),
    ...detectSSN(text),
    ...detectProjectCodes(text),
    ...detectFilePaths(text),
    ...detectIPAddresses(text),
    ...detectConfidentialKeywords(text),
    ...detectCreditCards(text),
    ...detectInternalURLs(text)
  ];

  findings.sort((a, b) => a.location - b.location);

  return {
    findings,
    summary: {
      total: findings.length,
      bySeverity: {
        critical: findings.filter(f => f.severity === "critical").length,
        high: findings.filter(f => f.severity === "high").length,
        medium: findings.filter(f => f.severity === "medium").length,
        low: findings.filter(f => f.severity === "low").length
      }
    }
  };
}

// ============================================================================
// EXPORT (ESM FORMAT) — compatible import { detectSensitiveData }
// ============================================================================

export {
  detectSensitiveData,
  detectEmails,
  detectPhones,
  detectPrices,
  detectIBANs,
  detectSSN,
  detectProjectCodes,
  detectFilePaths,
  detectIPAddresses,
  detectConfidentialKeywords,
  detectCreditCards,
  detectInternalURLs,
  getContext,
  maskValue,
  isValidPhone,
  isValidCreditCard,
  PATTERNS,
  CONFIDENTIAL_KEYWORDS,
  PRICE_CONTEXT_KEYWORDS
};
