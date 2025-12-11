// lib/sensitiveData.js
// VERSION 2.0 - Détection avancée des données sensibles avec contexte élargi
// Deploy to Render backend

// ============================================================================
// PATTERNS - Regex pour chaque type de données sensibles
// ============================================================================

const PATTERNS = {
  // Emails
  email: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/gi,
  
  // Téléphones internationaux (FR, US, UK, DE, etc.)
  phone: /(?:\+?\d{1,4}[\s.-]?)?(?:\(?\d{1,4}\)?[\s.-]?)?\d{1,4}[\s.-]?\d{1,4}[\s.-]?\d{1,9}/g,
  
  // Prix avec formats européens et américains
  // Capture: 2 500€, 2500€, €2,500.00, $4,500, 3.500,00 €, prix : 1500$
  priceEuro: /(?:(?:prix|tarif|coût|budget|montant|total|facture|devis)\s*[:=]?\s*)?(\d{1,3}(?:[\s.,]\d{3})*(?:[.,]\d{1,2})?)\s*[€$£¥CHF]{1,3}/gi,
  priceSymbolFirst: /[€$£¥]\s*(\d{1,3}(?:[,.\s]\d{3})*(?:[.,]\d{1,2})?)/gi,
  priceWithCurrency: /(\d{1,3}(?:[\s.,]\d{3})*(?:[.,]\d{1,2})?)\s*(?:euros?|dollars?|USD|EUR|GBP|CHF)/gi,
  
  // IBANs multi-pays
  iban: /[A-Z]{2}\d{2}[\s]?(?:\d{4}[\s]?){2,7}\d{1,4}/gi,
  
  // Numéros de sécurité sociale (FR)
  ssnFR: /[12]\s?\d{2}\s?\d{2}\s?\d{2}\s?\d{3}\s?\d{3}\s?\d{2}/g,
  
  // Codes projet internes
  projectCode: /(?:PRJ|PROJ|PROJECT|REF|DOSSIER|CASE|TICKET|INC|CHG|REQ)[-_]?\d{2,4}[-_]?\d{2,6}/gi,
  
  // Chemins fichiers Windows
  windowsPath: /[A-Z]:\\(?:[^\\/:*?"<>|\r\n]+\\)*[^\\/:*?"<>|\r\n]*/gi,
  
  // Chemins serveurs UNC
  uncPath: /\\\\[a-zA-Z0-9._-]+\\[a-zA-Z0-9._$-]+(?:\\[^\\/:*?"<>|\r\n]+)*/gi,
  
  // Chemins Unix/Linux/Mac
  unixPath: /(?:\/(?:home|var|usr|etc|opt|tmp|mnt|srv|data|backup|Users|Volumes)\/[^\s:*?"<>|]+)/gi,
  
  // Adresses IP v4
  ipv4: /(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)/g,
  
  // Numéros de carte de crédit
  creditCard: /(?:\d{4}[\s-]?){3}\d{4}/g,
  
  // URLs internes
  internalUrl: /https?:\/\/(?:intranet|internal|dev|staging|preprod|test|localhost|192\.168|10\.0|172\.(?:1[6-9]|2[0-9]|3[01]))[^\s<>"']*/gi
};

// ============================================================================
// MOTS-CLÉS CONFIDENTIELS - Multilingue
// ============================================================================

const CONFIDENTIAL_KEYWORDS = [
  // Français
  'CONFIDENTIEL', 'STRICTEMENT CONFIDENTIEL', 'NE PAS DIFFUSER',
  'USAGE INTERNE', 'INTERNE UNIQUEMENT', 'RÉSERVÉ', 'SECRET',
  'DIFFUSION RESTREINTE', 'NON DIVULGABLE', 'PRIVÉ',
  
  // Anglais
  'CONFIDENTIAL', 'STRICTLY CONFIDENTIAL', 'DO NOT DISTRIBUTE',
  'INTERNAL USE ONLY', 'INTERNAL ONLY', 'RESTRICTED', 'SECRET',
  'PRIVATE', 'NOT FOR DISTRIBUTION', 'PROPRIETARY',
  'FOR INTERNAL USE', 'DO NOT SHARE', 'DO NOT COPY',
  
  // Allemand
  'VERTRAULICH', 'STRENG VERTRAULICH', 'NUR FÜR INTERNEN GEBRAUCH',
  'INTERN', 'GEHEIM',
  
  // Espagnol
  'CONFIDENCIAL', 'USO INTERNO', 'NO DISTRIBUIR', 'PRIVADO', 'SECRETO'
];

// ============================================================================
// MOTS-CLÉS DE CONTEXTE PRIX
// ============================================================================

const PRICE_CONTEXT_KEYWORDS = [
  // Français
  'prix', 'tarif', 'coût', 'budget', 'montant', 'total', 'facture',
  'devis', 'honoraires', 'forfait', 'taux', 'journalier', 'tjm',
  'remise', 'réduction', 'marge', 'bénéfice', 'chiffre',
  
  // Anglais
  'price', 'cost', 'budget', 'amount', 'total', 'invoice', 'quote',
  'fee', 'rate', 'daily', 'discount', 'margin', 'profit', 'revenue',
  'billing', 'charge', 'payment', 'salary', 'wage'
];

// ============================================================================
// HELPER FUNCTIONS
// ============================================================================

/**
 * Récupère le contexte autour d'une position (50 chars avant/après)
 */
function getContext(text, startIndex, endIndex, contextLength = 50) {
  const beforeStart = Math.max(0, startIndex - contextLength);
  const afterEnd = Math.min(text.length, endIndex + contextLength);
  
  const before = text.substring(beforeStart, startIndex);
  const match = text.substring(startIndex, endIndex);
  const after = text.substring(endIndex, afterEnd);
  
  // Nettoyer les retours à la ligne pour l'affichage
  const cleanBefore = before.replace(/[\r\n]+/g, ' ').trim();
  const cleanAfter = after.replace(/[\r\n]+/g, ' ').trim();
  
  return {
    before: cleanBefore,
    match: match,
    after: cleanAfter,
    full: `...${cleanBefore}[${match}]${cleanAfter}...`
  };
}

/**
 * Vérifie si un numéro de téléphone est valide
 */
function isValidPhone(phone) {
  // Supprimer les espaces et caractères de formatage
  const digits = phone.replace(/[\s.-]/g, '');
  
  // Doit avoir entre 8 et 15 chiffres
  if (digits.length < 8 || digits.length > 15) return false;
  
  // Ne pas matcher les années (1990, 2024, etc.)
  if (/^(19|20)\d{2}$/.test(digits)) return false;
  
  // Ne pas matcher les codes postaux courts
  if (digits.length === 5 && /^\d{5}$/.test(digits)) return false;
  
  return true;
}

/**
 * Vérifie si un numéro ressemble à une carte de crédit valide (Luhn check)
 */
function isValidCreditCard(number) {
  const digits = number.replace(/[\s-]/g, '');
  if (digits.length < 13 || digits.length > 19) return false;
  
  // Vérification Luhn
  let sum = 0;
  let isEven = false;
  
  for (let i = digits.length - 1; i >= 0; i--) {
    let digit = parseInt(digits[i], 10);
    
    if (isEven) {
      digit *= 2;
      if (digit > 9) digit -= 9;
    }
    
    sum += digit;
    isEven = !isEven;
  }
  
  return sum % 10 === 0;
}

/**
 * Masque partiellement une valeur sensible
 */
function maskValue(value, type) {
  switch (type) {
    case 'email':
      const [local, domain] = value.split('@');
      return local.substring(0, 2) + '***@' + domain;
    case 'phone':
      return value.substring(0, 4) + '****' + value.slice(-2);
    case 'credit_card':
      return '****-****-****-' + value.slice(-4);
    case 'iban':
      return value.substring(0, 4) + ' **** **** ' + value.slice(-4);
    case 'ssn':
      return '***-**-' + value.slice(-4);
    default:
      return value.substring(0, 3) + '***';
  }
}

// ============================================================================
// DETECTION FUNCTIONS
// ============================================================================

/**
 * Détecte les adresses email
 */
function detectEmails(text) {
  const results = [];
  let match;
  
  while ((match = PATTERNS.email.exec(text)) !== null) {
    const context = getContext(text, match.index, match.index + match[0].length);
    
    results.push({
      type: 'email',
      value: match[0],
      masked: maskValue(match[0], 'email'),
      context: context.full,
      contextBefore: context.before,
      contextAfter: context.after,
      severity: 'high',
      location: `Position ${match.index}`,
      recommendation: 'Supprimer ou anonymiser cette adresse email avant partage externe'
    });
  }
  
  return results;
}

/**
 * Détecte les numéros de téléphone
 */
function detectPhones(text) {
  const results = [];
  let match;
  
  // Reset regex
  PATTERNS.phone.lastIndex = 0;
  
  while ((match = PATTERNS.phone.exec(text)) !== null) {
    const phone = match[0].trim();
    
    if (!isValidPhone(phone)) continue;
    
    const context = getContext(text, match.index, match.index + match[0].length);
    
    results.push({
      type: 'phone',
      value: phone,
      masked: maskValue(phone, 'phone'),
      context: context.full,
      contextBefore: context.before,
      contextAfter: context.after,
      severity: 'high',
      location: `Position ${match.index}`,
      recommendation: 'Supprimer ou masquer ce numéro de téléphone'
    });
  }
  
  return results;
}

/**
 * Détecte les prix et montants avec contexte élargi
 */
function detectPrices(text) {
  const results = [];
  const seenPositions = new Set();
  
  // Fonction helper pour ajouter un résultat
  const addResult = (match, startIndex) => {
    const endIndex = startIndex + match.length;
    const posKey = `${startIndex}-${endIndex}`;
    
    if (seenPositions.has(posKey)) return;
    seenPositions.add(posKey);
    
    const context = getContext(text, startIndex, endIndex, 60);
    
    // Vérifier le contexte pour la sévérité
    const contextLower = (context.before + context.after).toLowerCase();
    const hasContextKeyword = PRICE_CONTEXT_KEYWORDS.some(kw => contextLower.includes(kw));
    
    results.push({
      type: 'price',
      value: match,
      context: context.full,
      contextBefore: context.before,
      contextAfter: context.after,
      severity: hasContextKeyword ? 'high' : 'medium',
      location: `Position ${startIndex}`,
      recommendation: 'Vérifier si ce montant doit être visible dans le document partagé'
    });
  };
  
  // Pattern 1: Prix avec symbole après (2 500€, 3500$)
  let match;
  PATTERNS.priceEuro.lastIndex = 0;
  while ((match = PATTERNS.priceEuro.exec(text)) !== null) {
    addResult(match[0], match.index);
  }
  
  // Pattern 2: Prix avec symbole avant (€2,500, $4500)
  PATTERNS.priceSymbolFirst.lastIndex = 0;
  while ((match = PATTERNS.priceSymbolFirst.exec(text)) !== null) {
    addResult(match[0], match.index);
  }
  
  // Pattern 3: Prix avec devise textuelle (500 euros, 1000 USD)
  PATTERNS.priceWithCurrency.lastIndex = 0;
  while ((match = PATTERNS.priceWithCurrency.exec(text)) !== null) {
    addResult(match[0], match.index);
  }
  
  return results;
}

/**
 * Détecte les IBANs
 */
function detectIBANs(text) {
  const results = [];
  let match;
  
  PATTERNS.iban.lastIndex = 0;
  
  while ((match = PATTERNS.iban.exec(text)) !== null) {
    const context = getContext(text, match.index, match.index + match[0].length);
    
    results.push({
      type: 'iban',
      value: match[0],
      masked: maskValue(match[0], 'iban'),
      context: context.full,
      contextBefore: context.before,
      contextAfter: context.after,
      severity: 'critical',
      location: `Position ${match.index}`,
      recommendation: 'CRITIQUE: Supprimer immédiatement ce numéro de compte bancaire'
    });
  }
  
  return results;
}

/**
 * Détecte les numéros de sécurité sociale
 */
function detectSSN(text) {
  const results = [];
  let match;
  
  PATTERNS.ssnFR.lastIndex = 0;
  
  while ((match = PATTERNS.ssnFR.exec(text)) !== null) {
    const context = getContext(text, match.index, match.index + match[0].length);
    
    results.push({
      type: 'ssn',
      value: match[0],
      masked: maskValue(match[0], 'ssn'),
      context: context.full,
      contextBefore: context.before,
      contextAfter: context.after,
      severity: 'critical',
      location: `Position ${match.index}`,
      recommendation: 'CRITIQUE: Supprimer immédiatement ce numéro de sécurité sociale (données personnelles RGPD)'
    });
  }
  
  return results;
}

/**
 * Détecte les codes projet internes
 */
function detectProjectCodes(text) {
  const results = [];
  let match;
  
  PATTERNS.projectCode.lastIndex = 0;
  
  while ((match = PATTERNS.projectCode.exec(text)) !== null) {
    const context = getContext(text, match.index, match.index + match[0].length);
    
    results.push({
      type: 'project_code',
      value: match[0],
      context: context.full,
      contextBefore: context.before,
      contextAfter: context.after,
      severity: 'medium',
      location: `Position ${match.index}`,
      recommendation: 'Vérifier si ce code projet interne doit être visible'
    });
  }
  
  return results;
}

/**
 * Détecte les chemins de fichiers et serveurs
 */
function detectFilePaths(text) {
  const results = [];
  const seenPaths = new Set();
  
  const addPath = (match, startIndex, pathType) => {
    const path = match.trim();
    if (seenPaths.has(path)) return;
    seenPaths.add(path);
    
    const context = getContext(text, startIndex, startIndex + match.length);
    
    results.push({
      type: 'file_path',
      subtype: pathType,
      value: path,
      context: context.full,
      contextBefore: context.before,
      contextAfter: context.after,
      severity: 'high',
      location: `Position ${startIndex}`,
      recommendation: 'Supprimer ce chemin de fichier/serveur interne'
    });
  };
  
  let match;
  
  // Windows paths
  PATTERNS.windowsPath.lastIndex = 0;
  while ((match = PATTERNS.windowsPath.exec(text)) !== null) {
    addPath(match[0], match.index, 'windows');
  }
  
  // UNC paths
  PATTERNS.uncPath.lastIndex = 0;
  while ((match = PATTERNS.uncPath.exec(text)) !== null) {
    addPath(match[0], match.index, 'unc_server');
  }
  
  // Unix paths
  PATTERNS.unixPath.lastIndex = 0;
  while ((match = PATTERNS.unixPath.exec(text)) !== null) {
    addPath(match[0], match.index, 'unix');
  }
  
  return results;
}

/**
 * Détecte les adresses IP
 */
function detectIPAddresses(text) {
  const results = [];
  let match;
  
  PATTERNS.ipv4.lastIndex = 0;
  
  while ((match = PATTERNS.ipv4.exec(text)) !== null) {
    const ip = match[0];
    const context = getContext(text, match.index, match.index + ip.length);
    
    // Déterminer si c'est une IP interne
    const isInternal = ip.startsWith('192.168.') || 
                       ip.startsWith('10.') || 
                       ip.startsWith('172.16.') ||
                       ip.startsWith('172.17.') ||
                       ip.startsWith('172.18.') ||
                       ip.startsWith('172.19.') ||
                       ip.startsWith('172.2') ||
                       ip.startsWith('172.30.') ||
                       ip.startsWith('172.31.');
    
    results.push({
      type: 'ip_address',
      value: ip,
      isInternal: isInternal,
      context: context.full,
      contextBefore: context.before,
      contextAfter: context.after,
      severity: isInternal ? 'high' : 'medium',
      location: `Position ${match.index}`,
      recommendation: isInternal 
        ? 'Supprimer cette adresse IP interne (infrastructure sensible)'
        : 'Vérifier si cette adresse IP doit être visible'
    });
  }
  
  return results;
}

/**
 * Détecte les mots-clés de confidentialité
 */
function detectConfidentialKeywords(text) {
  const results = [];
  const textUpper = text.toUpperCase();
  
  for (const keyword of CONFIDENTIAL_KEYWORDS) {
    let index = 0;
    const keywordUpper = keyword.toUpperCase();
    
    while ((index = textUpper.indexOf(keywordUpper, index)) !== -1) {
      const context = getContext(text, index, index + keyword.length);
      
      results.push({
        type: 'confidential_keyword',
        value: text.substring(index, index + keyword.length),
        context: context.full,
        contextBefore: context.before,
        contextAfter: context.after,
        severity: keyword.includes('STRICT') || keyword.includes('SECRET') ? 'critical' : 'high',
        location: `Position ${index}`,
        recommendation: 'Ce document contient des marqueurs de confidentialité - vérifier les autorisations de diffusion'
      });
      
      index += keyword.length;
    }
  }
  
  return results;
}

/**
 * Détecte les numéros de carte de crédit
 */
function detectCreditCards(text) {
  const results = [];
  let match;
  
  PATTERNS.creditCard.lastIndex = 0;
  
  while ((match = PATTERNS.creditCard.exec(text)) !== null) {
    const cardNumber = match[0];
    
    if (!isValidCreditCard(cardNumber)) continue;
    
    const context = getContext(text, match.index, match.index + cardNumber.length);
    
    results.push({
      type: 'credit_card',
      value: maskValue(cardNumber, 'credit_card'),
      context: context.full,
      contextBefore: context.before,
      contextAfter: context.after,
      severity: 'critical',
      location: `Position ${match.index}`,
      recommendation: 'CRITIQUE: Supprimer immédiatement ce numéro de carte de crédit (PCI-DSS)'
    });
  }
  
  return results;
}

/**
 * Détecte les URLs internes
 */
function detectInternalURLs(text) {
  const results = [];
  let match;
  
  PATTERNS.internalUrl.lastIndex = 0;
  
  while ((match = PATTERNS.internalUrl.exec(text)) !== null) {
    const context = getContext(text, match.index, match.index + match[0].length);
    
    results.push({
      type: 'internal_url',
      value: match[0],
      context: context.full,
      contextBefore: context.before,
      contextAfter: context.after,
      severity: 'high',
      location: `Position ${match.index}`,
      recommendation: 'Supprimer cette URL interne avant partage externe'
    });
  }
  
  return results;
}

// ============================================================================
// MAIN FUNCTION
// ============================================================================

/**
 * Détecte toutes les données sensibles dans un texte
 * @param {string} text - Le texte à analyser
 * @returns {Object} - Résultats de la détection avec catégories et statistiques
 */
function detectSensitiveData(text) {
  if (!text || typeof text !== 'string') {
    return {
      findings: [],
      summary: {
        total: 0,
        bySeverity: { critical: 0, high: 0, medium: 0, low: 0 },
        byType: {}
      }
    };
  }
  
  console.log(`[SENSITIVE DATA v2.0] Analyzing ${text.length} characters...`);
  
  // Collecter toutes les détections
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
  
  // Trier par position
  findings.sort((a, b) => {
    const posA = parseInt(a.location.replace('Position ', ''));
    const posB = parseInt(b.location.replace('Position ', ''));
    return posA - posB;
  });
  
  // Calculer les statistiques
  const summary = {
    total: findings.length,
    bySeverity: {
      critical: findings.filter(f => f.severity === 'critical').length,
      high: findings.filter(f => f.severity === 'high').length,
      medium: findings.filter(f => f.severity === 'medium').length,
      low: findings.filter(f => f.severity === 'low').length
    },
    byType: {}
  };
  
  // Compter par type
  for (const finding of findings) {
    summary.byType[finding.type] = (summary.byType[finding.type] || 0) + 1;
  }
  
  console.log(`[SENSITIVE DATA v2.0] Found ${findings.length} sensitive items:`, summary.byType);
  
  return {
    findings,
    summary
  };
}

// ============================================================================
// EXPORTS
// ============================================================================

module.exports = {
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
  // Helpers
  getContext,
  maskValue,
  isValidPhone,
  isValidCreditCard,
  // Constants (for testing)
  PATTERNS,
  CONFIDENTIAL_KEYWORDS,
  PRICE_CONTEXT_KEYWORDS
};
