// lib/riskDetector.js
// Détection des risques dans le texte extrait

const riskPatterns = [
  // Personal data (GDPR)
  { pattern: /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, type: 'email', severity: 'medium', category: 'personal' },
  { pattern: /(\+?\d{1,3}[-.\s]?)?(\(?\d{2,4}\)?[-.\s]?)?\d{3,4}[-.\s]?\d{3,4}/g, type: 'phone', severity: 'medium', category: 'personal' },
  { pattern: /[A-Z]{2}\d{2}[A-Z0-9]{10,30}/g, type: 'iban', severity: 'high', category: 'financial' },
  
  // Internal sensitive data
  { pattern: /(\d{1,3}(?:[,.\s]\d{3})*(?:[.,]\d{2})?[\s]?[€$£])|([€$£][\s]?\d{1,3}(?:[,.\s]\d{3})*)/g, type: 'pricing', severity: 'high', category: 'internal' },
  { pattern: /(PROJ[-_]?\d+)|(\#\d{4,})|([A-Z]{2,4}[-_]\d{3,})/g, type: 'project_code', severity: 'medium', category: 'internal' },
  { pattern: /\\\\[A-Za-z0-9_-]+\\[A-Za-z0-9_\-\\]+/g, type: 'server_path', severity: 'high', category: 'internal' },
  { pattern: /[A-Z]:\\[A-Za-z0-9_\-\\]+/g, type: 'file_path', severity: 'medium', category: 'internal' },
  
  // Confidential keywords
  { pattern: /\b(confidentiel|confidential|secret|interne|internal only|ne pas diffuser|do not distribute)\b/gi, type: 'confidential_keyword', severity: 'high', category: 'compliance' },
  { pattern: /\b(draft|brouillon|preliminary|work in progress|wip)\b/gi, type: 'draft_indicator', severity: 'medium', category: 'quality' },
  
  // French specific
  { pattern: /\b\d{1,2}\s?(rue|avenue|boulevard|place|chemin)\s[A-Za-zÀ-ÿ\s-]+,?\s?\d{5}\s?[A-Za-zÀ-ÿ\s-]+/gi, type: 'french_address', severity: 'high', category: 'personal' },
  { pattern: /\b[12]\d{2}[01]\d[0-3]\d\d{5}[0-9]{2}\b/g, type: 'french_ssn', severity: 'high', category: 'personal' },
  
  // Technical risks
  { pattern: /https?:\/\/[^\s<>"{}|\\^`\[\]]+/g, type: 'url', severity: 'low', category: 'technical' },
  { pattern: /\b(password|mot de passe|pwd|passwd)\s*[:=]\s*\S+/gi, type: 'password_leak', severity: 'critical', category: 'security' },
  { pattern: /\b(api[_-]?key|apikey|secret[_-]?key)\s*[:=]\s*\S+/gi, type: 'api_key_leak', severity: 'critical', category: 'security' },
];

/**
 * Détecte les risques dans un texte
 * @param {string} text - Texte extrait du document
 * @returns {Array} - Liste des risques détectés
 */
export function detectRisks(text) {
  const risks = [];
  
  for (const { pattern, type, severity, category } of riskPatterns) {
    const matches = text.matchAll(new RegExp(pattern.source, pattern.flags));
    
    for (const match of matches) {
      // Avoid duplicates
      const value = match[0].trim();
      if (value.length < 3) continue;
      if (risks.some(r => r.value === value && r.type === type)) continue;
      
      // Get context (50 chars before and after)
      const startIdx = Math.max(0, match.index - 50);
      const endIdx = Math.min(text.length, match.index + value.length + 50);
      const context = text.slice(startIdx, endIdx).replace(/\s+/g, ' ').trim();
      
      risks.push({
        type,
        severity,
        category,
        value: value.slice(0, 100),
        description: getDescriptionForType(type, value),
        location: `Character ${match.index}`,
        contextSnippet: `...${context}...`
      });
      
      // Limit per type to avoid spam
      if (risks.filter(r => r.type === type).length >= 10) break;
    }
  }
  
  return risks.sort((a, b) => {
    const severityOrder = { critical: 0, high: 1, medium: 2, low: 3 };
    return severityOrder[a.severity] - severityOrder[b.severity];
  });
}

function getDescriptionForType(type, value) {
  const descriptions = {
    email: `Email address detected: ${value.slice(0, 30)}...`,
    phone: `Phone number detected`,
    iban: `IBAN bank account detected`,
    pricing: `Pricing/amount information: ${value}`,
    project_code: `Internal project code: ${value}`,
    server_path: `Internal server path exposed`,
    file_path: `Local file path exposed`,
    confidential_keyword: `Confidential marking found: "${value}"`,
    draft_indicator: `Draft/preliminary indicator: "${value}"`,
    french_address: `French postal address detected`,
    french_ssn: `French social security number detected`,
    url: `URL reference found`,
    password_leak: `Potential password exposure!`,
    api_key_leak: `Potential API key exposure!`
  };
  
  return descriptions[type] || `${type} detected`;
}

/**
 * Calcule un score de sécurité
 */
export function calculateSecurityScore(cleaningSummary, risks) {
  let score = 100;
  
  // Deduct for cleaning actions needed
  if (cleaningSummary.metadataRemoved?.length > 0) score -= cleaningSummary.metadataRemoved.length * 2;
  if (cleaningSummary.commentsRemoved > 0) score -= cleaningSummary.commentsRemoved * 3;
  if (cleaningSummary.trackChangesAccepted) score -= 10;
  if (cleaningSummary.embeddedObjectsRemoved > 0) score -= cleaningSummary.embeddedObjectsRemoved * 5;
  if (cleaningSummary.macrosRemoved > 0) score -= 15;
  
  // Deduct for risks
  for (const risk of risks) {
    switch (risk.severity) {
      case 'critical': score -= 20; break;
      case 'high': score -= 10; break;
      case 'medium': score -= 5; break;
      case 'low': score -= 2; break;
    }
  }
  
  score = Math.max(0, Math.min(100, score));
  
  let level;
  if (score >= 90) level = 'clean';
  else if (score >= 70) level = 'high';
  else if (score >= 50) level = 'medium';
  else if (score >= 30) level = 'low';
  else level = 'critical';
  
  return {
    score,
    maxScore: 100,
    level,
    breakdown: {
      technicalThreats: risks.filter(r => ['security', 'technical'].includes(r.category)).length,
      businessRisks: risks.filter(r => ['internal', 'compliance', 'personal', 'financial'].includes(r.category)).length
    },
    explanation: getScoreExplanation(level, score)
  };
}

function getScoreExplanation(level, score) {
  const explanations = {
    clean: 'Document is well sanitized and ready for external sharing.',
    high: 'Document has minor issues that should be reviewed before sharing.',
    medium: 'Document contains several elements that need attention before sharing.',
    low: 'Document has significant issues that must be addressed before sharing.',
    critical: 'Document contains critical security or privacy issues. Do not share externally.'
  };
  return explanations[level];
}

export default { detectRisks, calculateSecurityScore };
