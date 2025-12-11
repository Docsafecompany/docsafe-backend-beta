// lib/aiSpellCheck.js
// VERSION 5.0 - Approche stricte avec dictionnaire obligatoire
// Corrige : "Confi dential" → "Confidential", "sl ide" → "slide"

// ============= DICTIONNAIRE DE MOTS VALIDES =============
const COMMON_WORDS = new Set([
  // Basic words
  'the', 'be', 'to', 'of', 'and', 'a', 'in', 'that', 'have', 'i',
  'it', 'for', 'not', 'on', 'with', 'he', 'as', 'you', 'do', 'at',
  'this', 'but', 'his', 'by', 'from', 'they', 'we', 'say', 'her', 'she',
  'or', 'an', 'will', 'my', 'one', 'all', 'would', 'there', 'their', 'what',
  'so', 'up', 'out', 'if', 'about', 'who', 'get', 'which', 'go', 'me',
  'when', 'make', 'can', 'like', 'time', 'no', 'just', 'him', 'know', 'take',
  'people', 'into', 'year', 'your', 'good', 'some', 'could', 'them', 'see', 'other',
  'than', 'then', 'now', 'look', 'only', 'come', 'its', 'over', 'think', 'also',
  'back', 'after', 'use', 'two', 'how', 'our', 'work', 'first', 'well', 'way',
  'even', 'new', 'want', 'because', 'any', 'these', 'give', 'day', 'most', 'us',
  
  // Common verbs
  'become', 'becomes', 'became', 'becoming',
  'provide', 'provides', 'provided', 'providing',
  'foster', 'fosters', 'fostered', 'fostering',
  'remain', 'remains', 'remained', 'remaining',
  'strike', 'strikes', 'struck', 'striking',
  'guide', 'guides', 'guided', 'guiding',
  'shape', 'shapes', 'shaped', 'shaping',
  'lie', 'lies', 'lay', 'lying',
  'revolutionize', 'revolutionizes', 'revolutionized', 'revolutionizing',
  'enable', 'enables', 'enabled', 'enabling',
  'redefine', 'redefines', 'redefined', 'redefining',
  'connect', 'connects', 'connected', 'connecting',
  'message', 'messages', 'messaging',
  'process', 'processes', 'processed', 'processing',
  'analyze', 'analyzes', 'analyzed', 'analyzing',
  'create', 'creates', 'created', 'creating',
  'develop', 'develops', 'developed', 'developing',
  'implement', 'implements', 'implemented', 'implementing',
  'manage', 'manages', 'managed', 'managing',
  'support', 'supports', 'supported', 'supporting',
  'include', 'includes', 'included', 'including',
  'require', 'requires', 'required', 'requiring',
  'consider', 'considers', 'considered', 'considering',
  'continue', 'continues', 'continued', 'continuing',
  'increase', 'increases', 'increased', 'increasing',
  'decrease', 'decreases', 'decreased', 'decreasing',
  'establish', 'establishes', 'established', 'establishing',
  'maintain', 'maintains', 'maintained', 'maintaining',
  'determine', 'determines', 'determined', 'determining',
  'occur', 'occurs', 'occurred', 'occurring',
  'appear', 'appears', 'appeared', 'appearing',
  'represent', 'represents', 'represented', 'representing',
  
  // Business/Professional terms
  'business', 'company', 'corporate', 'corporation', 'corporations',
  'professional', 'management', 'strategy', 'strategic',
  'development', 'implementation', 'performance', 'organization',
  'communication', 'communications', 'information', 'technology',
  'consultant', 'consulting', 'analysis', 'assessment',
  'project', 'projects', 'process', 'processes',
  'service', 'services', 'solution', 'solutions',
  'client', 'clients', 'customer', 'customers',
  'market', 'markets', 'marketing', 'industry', 'industries',
  'financial', 'finance', 'economic', 'economics',
  'investment', 'investments', 'revenue', 'revenues',
  'growth', 'profit', 'profits', 'cost', 'costs',
  'budget', 'budgets', 'resource', 'resources',
  'team', 'teams', 'leader', 'leaders', 'leadership',
  'employee', 'employees', 'staff', 'personnel',
  'department', 'departments', 'division', 'divisions',
  'executive', 'executives', 'director', 'directors',
  'manager', 'managers', 'supervisor', 'supervisors',
  'partner', 'partners', 'partnership', 'partnerships',
  'stakeholder', 'stakeholders', 'shareholder', 'shareholders',
  
  // Document/Report terms
  'document', 'documents', 'report', 'reports', 'reporting',
  'presentation', 'presentations', 'slide', 'slides',
  'content', 'contents', 'page', 'pages', 'section', 'sections',
  'chapter', 'chapters', 'appendix', 'summary', 'summaries',
  'introduction', 'conclusion', 'conclusions', 'recommendation', 'recommendations',
  'finding', 'findings', 'result', 'results', 'outcome', 'outcomes',
  'objective', 'objectives', 'goal', 'goals', 'target', 'targets',
  'scope', 'methodology', 'approach', 'framework', 'frameworks',
  'overview', 'background', 'context', 'description', 'descriptions',
  'detail', 'details', 'detailed', 'example', 'examples',
  'figure', 'figures', 'table', 'tables', 'chart', 'charts',
  'graph', 'graphs', 'diagram', 'diagrams', 'image', 'images',
  'reference', 'references', 'source', 'sources', 'citation', 'citations',
  'note', 'notes', 'comment', 'comments', 'annotation', 'annotations',
  'draft', 'drafts', 'version', 'versions', 'revision', 'revisions',
  'review', 'reviews', 'feedback', 'approval', 'approvals',
  
  // Technical terms
  'system', 'systems', 'software', 'hardware', 'network', 'networks',
  'database', 'databases', 'server', 'servers', 'application', 'applications',
  'platform', 'platforms', 'interface', 'interfaces', 'module', 'modules',
  'component', 'components', 'feature', 'features', 'function', 'functions',
  'functionality', 'configuration', 'integration', 'integrations',
  'security', 'privacy', 'compliance', 'regulation', 'regulations',
  'standard', 'standards', 'protocol', 'protocols', 'procedure', 'procedures',
  'requirement', 'requirements', 'specification', 'specifications',
  'architecture', 'infrastructure', 'environment', 'environments',
  'testing', 'deployment', 'maintenance', 'operation', 'operations',
  
  // Common adjectives
  'social', 'digital', 'global', 'local', 'national', 'international',
  'public', 'private', 'personal', 'individual', 'individuals',
  'general', 'specific', 'particular', 'special', 'unique',
  'major', 'minor', 'significant', 'important', 'critical', 'essential',
  'primary', 'secondary', 'main', 'key', 'core', 'central',
  'current', 'previous', 'next', 'future', 'past', 'present',
  'new', 'old', 'modern', 'traditional', 'conventional',
  'high', 'low', 'large', 'small', 'big', 'great',
  'long', 'short', 'full', 'complete', 'partial', 'total',
  'available', 'accessible', 'effective', 'efficient', 'successful',
  'positive', 'negative', 'potential', 'possible', 'probable',
  'different', 'similar', 'same', 'other', 'additional', 'further',
  'necessary', 'required', 'optional', 'recommended', 'preferred',
  'appropriate', 'suitable', 'relevant', 'applicable', 'related',
  'dark', 'darker', 'light', 'lighter', 'bright', 'brighter',
  'right', 'rights', 'left', 'correct', 'incorrect',
  'true', 'false', 'real', 'actual', 'virtual',
  
  // Common nouns
  'way', 'ways', 'part', 'parts', 'place', 'places',
  'case', 'cases', 'point', 'points', 'fact', 'facts',
  'issue', 'issues', 'problem', 'problems', 'challenge', 'challenges',
  'opportunity', 'opportunities', 'risk', 'risks', 'benefit', 'benefits',
  'advantage', 'advantages', 'disadvantage', 'disadvantages',
  'factor', 'factors', 'element', 'elements', 'aspect', 'aspects',
  'level', 'levels', 'degree', 'degrees', 'rate', 'rates',
  'type', 'types', 'kind', 'kinds', 'form', 'forms',
  'area', 'areas', 'region', 'regions', 'zone', 'zones',
  'period', 'periods', 'phase', 'phases', 'stage', 'stages',
  'step', 'steps', 'action', 'actions', 'activity', 'activities',
  'event', 'events', 'situation', 'situations', 'condition', 'conditions',
  'state', 'states', 'status', 'position', 'positions',
  'role', 'roles', 'responsibility', 'responsibilities',
  'relationship', 'relationships', 'connection', 'connections',
  'impact', 'impacts', 'effect', 'effects', 'influence', 'influences',
  'change', 'changes', 'difference', 'differences', 'variation', 'variations',
  'trend', 'trends', 'pattern', 'patterns', 'model', 'models',
  'method', 'methods', 'technique', 'techniques', 'tool', 'tools',
  'measure', 'measures', 'indicator', 'indicators', 'metric', 'metrics',
  'value', 'values', 'quality', 'qualities', 'property', 'properties',
  'characteristic', 'characteristics', 'attribute', 'attributes',
  'category', 'categories', 'class', 'classes', 'group', 'groups',
  'set', 'sets', 'series', 'range', 'ranges', 'spectrum',
  'structure', 'structures', 'format', 'formats', 'layout', 'layouts',
  'design', 'designs', 'style', 'styles', 'theme', 'themes',
  
  // Connectivity/Communication terms
  'connectivity', 'connection', 'connections', 'connected',
  'media', 'internet', 'online', 'web', 'website', 'websites',
  'email', 'emails', 'phone', 'phones', 'mobile', 'mobiles',
  'continent', 'continents', 'country', 'countries', 'city', 'cities',
  'world', 'worldwide', 'landscape', 'landscapes',
  'misinformation', 'disinformation', 'information',
  'news', 'article', 'articles', 'post', 'posts',
  'user', 'users', 'audience', 'audiences', 'reader', 'readers',
  'viewer', 'viewers', 'visitor', 'visitors', 'follower', 'followers',
  'subscriber', 'subscribers', 'member', 'members',
  
  // Academic/Research terms
  'additionally', 'furthermore', 'however', 'therefore', 'consequently',
  'moreover', 'nevertheless', 'nonetheless', 'meanwhile', 'otherwise',
  'emphasized', 'collaborated', 'coordinated', 'accessibility', 'empowerment',
  'marginalized', 'momentum', 'amplify', 'oversight', 'verification',
  'circulation', 'polarization', 'exacerbated', 'commodification', 'monetized',
  'multifaceted', 'geographical', 'interpersonal', 'essence',
  'largely', 'due', 'not', 'yet', 'still', 'already',
  
  // Security/Confidential terms
  'confidential', 'confidentiality', 'secret', 'secrets', 'private',
  'sensitive', 'restricted', 'classified', 'internal', 'external',
  'secure', 'secured', 'security', 'protection', 'protected',
  'authorized', 'unauthorized', 'permission', 'permissions',
  'access', 'accessed', 'accessible', 'inaccessible',
  'encrypt', 'encrypted', 'encryption', 'decrypt', 'decrypted',
  'authenticate', 'authenticated', 'authentication',
  'verify', 'verified', 'verification', 'validate', 'validated', 'validation',
  
  // Correction/News terms  
  'correction', 'corrections', 'corrected', 'correcting',
  'news', 'newspaper', 'newspapers', 'newsletter', 'newsletters',
  'price', 'prices', 'pricing', 'priced',
  'dollar', 'dollars', 'euro', 'euros', 'pound', 'pounds',
  
  // Common prepositions and connectors
  'about', 'above', 'across', 'after', 'against', 'along', 'among', 'around',
  'before', 'behind', 'below', 'beneath', 'beside', 'between', 'beyond',
  'during', 'except', 'inside', 'outside', 'through', 'throughout',
  'toward', 'towards', 'under', 'underneath', 'until', 'upon', 'within', 'without',
  
  // French common words (for multilingual documents)
  'le', 'la', 'les', 'un', 'une', 'des', 'du', 'de', 'et', 'ou', 'mais',
  'pour', 'dans', 'sur', 'avec', 'sans', 'sous', 'entre', 'vers', 'chez',
  'faire', 'plat', 'être', 'avoir', 'pouvoir', 'vouloir', 'devoir', 'savoir',
  'prix', 'coût', 'tarif', 'euro', 'euros', 'dollar', 'dollars',
  
  // Numbers written as words
  'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten',
  'first', 'second', 'third', 'fourth', 'fifth', 'sixth', 'seventh', 'eighth', 'ninth', 'tenth',
  'hundred', 'hundreds', 'thousand', 'thousands', 'million', 'millions', 'billion', 'billions'
]);

// ============= PHRASES VALIDES À NE PAS FUSIONNER =============
const VALID_PHRASES = new Set([
  // Common word combinations that should NOT be merged
  'become one', 'become a', 'become the', 'become an',
  'message can', 'message a', 'message the',
  'fostered a', 'fostered the', 'fostered an',
  'provides a', 'provides the', 'provides an',
  'individuals can', 'individuals are', 'individuals have',
  'continents can', 'continents are',
  'connectivity can', 'connectivity is',
  'media can', 'media is', 'media are',
  'misinformation can', 'misinformation is',
  'revolutionized how', 'revolutionized the',
  'shaping how', 'shaping the', 'shaping a',
  'largely due', 'largely to', 'largely the',
  'remains a', 'remains the', 'remains an',
  'strike a', 'strike the', 'strike an',
  'lies not', 'lies in', 'lies the', 'lies a',
  'than a', 'than the', 'than an',
  'and x', 'and a', 'and the', 'and an',
  'one c', 'one a', 'one the', // Fragment réel
  's are', 'e use', 'e the', // Fragments réels
  'a message', 'a the', 'a an',
  'the price', 'the cost', 'the value',
  'in the', 'on the', 'at the', 'to the', 'for the', 'by the', 'with the',
  'of the', 'from the', 'into the', 'onto the', 'upon the',
  'is the', 'was the', 'are the', 'were the', 'be the',
  'has the', 'have the', 'had the',
  'can the', 'could the', 'would the', 'should the', 'will the',
  'not the', 'not a', 'not an',
  'due to', 'due the',
  // Verb + preposition/article combinations
  'guide the', 'guide a', 'guide an',
  'enable the', 'enable a', 'enable an',
  'shape the', 'shape a',
  'strike the', 'strike a',
  // Word + common short words
  'how the', 'how a', 'how to', 'how we', 'how it',
  'one of', 'one the', 'one a',
  'use the', 'use a', 'use an',
  // Additional patterns
  'faire un', 'faire le', 'faire la', 'faire une', 'faire des',
  'un plat', 'le plat', 'la plat'
]);

// ============= FONCTIONS DE VALIDATION =============

/**
 * Vérifie si un mot existe dans le dictionnaire
 * STRICT: retourne true UNIQUEMENT si le mot est dans COMMON_WORDS
 */
function isRealWord(word) {
  if (!word || word.length < 2) return false;
  const normalized = word.toLowerCase().trim();
  return COMMON_WORDS.has(normalized);
}

/**
 * Vérifie si une phrase est valide et ne doit pas être fusionnée
 */
function isValidPhrase(phrase) {
  const normalized = phrase.toLowerCase().trim();
  
  // Check direct match
  if (VALID_PHRASES.has(normalized)) return true;
  
  // Check pattern: word + common short word
  const parts = normalized.split(/\s+/);
  if (parts.length === 2) {
    const [first, second] = parts;
    
    // Common short words that often follow other words
    const shortWords = ['a', 'an', 'the', 'to', 'of', 'in', 'on', 'at', 'by', 'for', 'with', 'is', 'are', 'was', 'were', 'be', 'can', 'could', 'would', 'should', 'will', 'may', 'might', 'must', 'have', 'has', 'had', 'do', 'does', 'did', 'not', 'how', 'one', 'due'];
    
    // If second word is a common short word and first word is valid
    if (shortWords.includes(second) && isRealWord(first)) {
      return true;
    }
    
    // If both parts are valid standalone words
    if (isRealWord(first) && isRealWord(second)) {
      // Check if combining them would create a valid word
      const combined = first + second;
      // Only return true (valid phrase) if combined is NOT a real word
      // This means "Confi dential" won't be blocked because "confidential" IS a real word
      if (!isRealWord(combined)) {
        return true; // Keep as two words
      }
    }
  }
  
  return false;
}

/**
 * Pré-détection des mots fragmentés par regex AVANT l'appel IA
 * VERSION 5.0 - Patterns élargis pour capturer "Confi dential", "sl ide", etc.
 */
function preDetectFragmentedWords(text) {
  const issues = [];
  
  // Pattern 1: Mot coupé par un espace (ex: "Confi dential", "soc ial", "sl ide")
  // Élargi pour capturer des fragments de 2 à 12 lettres de chaque côté
  const fragmentPattern = /\b([a-zA-Z]{2,})[ ]+([a-zA-Z]{2,12})\b/g;
  let match;
  
  while ((match = fragmentPattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const part1 = match[1];
    const part2 = match[2];
    const combined = part1 + part2;
    
    // STRICT: Vérifier que le mot combiné existe dans le dictionnaire
    // ET que ce n'est pas une phrase valide
    if (isRealWord(combined) && !isValidPhrase(fullMatch.toLowerCase())) {
      issues.push({
        error: fullMatch,
        correction: combined.charAt(0).toUpperCase() + combined.slice(1).toLowerCase(),
        type: 'fragmented_word',
        severity: 'high',
        message: `Word incorrectly split by space: "${fullMatch}" → "${combined}"`
      });
    }
  }
  
  // Pattern 2: Lettre isolée suivie d'un mot (ex: "o f" → "of", "th e" → "the")
  const singleLetterPattern = /\b([a-zA-Z])[ ]+([a-zA-Z]{1,15})\b/g;
  
  while ((match = singleLetterPattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const letter = match[1];
    const rest = match[2];
    const combined = letter + rest;
    
    // Vérifier que le mot combiné existe
    if (isRealWord(combined) && !isValidPhrase(fullMatch.toLowerCase())) {
      issues.push({
        error: fullMatch,
        correction: combined.toLowerCase(),
        type: 'fragmented_word',
        severity: 'high',
        message: `Single letter fragment: "${fullMatch}" → "${combined}"`
      });
    }
  }
  
  // Pattern 3: Mot suivi d'une lettre isolée (ex: "essenc e" → "essence")
  const trailingLetterPattern = /\b([a-zA-Z]{2,})[ ]+([a-zA-Z])\b/g;
  
  while ((match = trailingLetterPattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const word = match[1];
    const letter = match[2];
    const combined = word + letter;
    
    // Vérifier que le mot combiné existe
    if (isRealWord(combined) && !isValidPhrase(fullMatch.toLowerCase())) {
      issues.push({
        error: fullMatch,
        correction: combined.toLowerCase(),
        type: 'fragmented_word',
        severity: 'high',
        message: `Trailing letter fragment: "${fullMatch}" → "${combined}"`
      });
    }
  }
  
  // Pattern 4: Espaces multiples dans un mot (ex: "p  otential" → "potential")
  const multiSpacePattern = /\b([a-zA-Z]+)[ ]{2,}([a-zA-Z]+)\b/g;
  
  while ((match = multiSpacePattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const part1 = match[1];
    const part2 = match[2];
    const combined = part1 + part2;
    
    if (isRealWord(combined)) {
      issues.push({
        error: fullMatch,
        correction: combined.toLowerCase(),
        type: 'multiple_spaces',
        severity: 'high',
        message: `Multiple spaces in word: "${fullMatch}" → "${combined}"`
      });
    }
  }
  
  return issues;
}

/**
 * Filtre les corrections invalides de l'IA
 * VERSION 5.0 - Approche stricte avec validation dictionnaire
 */
function filterInvalidCorrections(corrections) {
  return corrections.filter(c => {
    const error = (c.error || '').trim();
    const correction = (c.correction || '').trim();
    
    // Rejeter si pas de changement réel
    if (error.toLowerCase() === correction.toLowerCase()) {
      console.log(`❌ REJECTED (no change): "${error}" → "${correction}"`);
      return false;
    }
    
    // Rejeter si c'est une phrase valide qu'on essaie de fusionner
    if (isValidPhrase(error.toLowerCase())) {
      console.log(`❌ REJECTED (valid phrase): "${error}"`);
      return false;
    }
    
    // Pour les fusions (correction sans espace, erreur avec espace)
    const errorHasSpace = /\s/.test(error);
    const correctionHasNoSpace = !/\s/.test(correction);
    
    if (errorHasSpace && correctionHasNoSpace) {
      // C'est une tentative de fusion - vérifier que le mot résultant existe
      if (!isRealWord(correction)) {
        console.log(`❌ REJECTED (not a real word): "${error}" → "${correction}"`);
        return false;
      }
      
      // Vérifier que les deux parties ne sont pas des mots valides séparés
      const parts = error.split(/\s+/);
      if (parts.length === 2 && isRealWord(parts[0]) && isRealWord(parts[1])) {
        // Les deux parties sont des mots valides - ne pas fusionner sauf si la fusion donne un mot plus courant
        const combinedExists = isRealWord(correction);
        if (!combinedExists) {
          console.log(`❌ REJECTED (both parts are valid words): "${error}"`);
          return false;
        }
        // Si le mot combiné existe ET l'erreur ressemble à un mot fragmenté (pas une phrase)
        // Vérifier que ce n'est pas un pattern comme "word can", "word the", etc.
        const [first, second] = parts;
        const commonConnectors = ['a', 'an', 'the', 'can', 'is', 'are', 'was', 'were', 'has', 'have', 'had', 'will', 'would', 'could', 'should', 'may', 'might', 'to', 'of', 'in', 'on', 'at', 'by', 'for', 'with', 'not', 'how', 'one', 'due'];
        if (commonConnectors.includes(second.toLowerCase())) {
          console.log(`❌ REJECTED (word + connector pattern): "${error}"`);
          return false;
        }
      }
    }
    
    // Accepter la correction
    console.log(`✅ ACCEPTED: "${error}" → "${correction}"`);
    return true;
  });
}

// ============= PROMPT IA OPTIMISÉ =============

const PROFESSIONAL_PROOFREADER_PROMPT = `You are a professional proofreader. Analyze the text and identify ONLY genuine spelling, grammar, and fragmentation errors.

CRITICAL RULES - READ CAREFULLY:
1. FRAGMENTED WORDS (priority): Detect words incorrectly split by spaces:
   - "Confi dential" → "Confidential" ✅
   - "sl ide" → "slide" ✅  
   - "soc ial" → "social" ✅
   - "p otential" → "potential" ✅
   - "th e" → "the" ✅
   - "commu nication" → "communication" ✅

2. NEVER MERGE VALID SEPARATE WORDS:
   - "become one" → KEEP AS IS ❌ (two valid words)
   - "message can" → KEEP AS IS ❌ (two valid words)
   - "fostered a" → KEEP AS IS ❌ (verb + article)
   - "than a" → KEEP AS IS ❌ (two valid words)
   - "strike a" → KEEP AS IS ❌ (two valid words)
   - "lies not" → KEEP AS IS ❌ (two valid words)
   - "largely due" → KEEP AS IS ❌ (two valid words)

3. HOW TO DECIDE:
   - If BOTH parts are real English words AND combining them doesn't create a MORE COMMON word → DO NOT merge
   - If combining creates a real word that makes more sense in context → merge it
   - Examples:
     - "Confi" is NOT a word, "dential" is NOT a word → merge to "Confidential" ✅
     - "sl" is NOT a word, "ide" is NOT a word → merge to "slide" ✅
     - "become" IS a word, "one" IS a word → DO NOT merge ❌
     - "than" IS a word, "a" IS a word → DO NOT merge ❌

4. OTHER ERRORS TO DETECT:
   - Real typos: "gggdigital" → "digital"
   - Punctuation errors: "corpo,rations" → "corporations"
   - Currency formatting: "4500$" → "$4,500" (optional)

Return a JSON array with this exact structure:
[
  {
    "error": "exact text with error",
    "correction": "corrected text",
    "type": "fragmented_word|spelling|grammar|punctuation",
    "severity": "high|medium|low",
    "message": "brief explanation"
  }
]

Return ONLY the JSON array, no other text.`;

// ============= FONCTION PRINCIPALE =============

/**
 * Analyse orthographique avec IA + pré-détection regex
 * @param {string} text - Texte à analyser
 * @returns {Promise<Array>} Liste des erreurs détectées
 */
export async function checkSpellingWithAI(text) {
  console.log('📝 AI spell-check VERSION 5.0 starting...');
  
  if (!text || text.length < 10) {
    return [];
  }
  
  // Limite de caractères pour éviter les tokens excessifs
  const maxChars = 15000;
  const truncatedText = text.length > maxChars ? text.substring(0, maxChars) : text;
  console.log(`[SPELLCHECK] Text length: ${truncatedText.length} characters`);
  
  // 1. Pré-détection par regex (garantie de détection des patterns évidents)
  const regexIssues = preDetectFragmentedWords(truncatedText);
  console.log(`[REGEX] Found ${regexIssues.length} fragmented word patterns`);
  
  // 2. Appel à l'IA pour les cas plus subtils
  let aiIssues = [];
  
  try {
    const OPENAI_API_KEY = process.env.OPENAI_API_KEY;
    
    if (!OPENAI_API_KEY) {
      console.log('[SPELLCHECK] No OpenAI API key, using regex-only detection');
      return regexIssues;
    }
    
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${OPENAI_API_KEY}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [
          { role: 'system', content: PROFESSIONAL_PROOFREADER_PROMPT },
          { role: 'user', content: `Analyze this text for errors:\n\n${truncatedText}` }
        ],
        temperature: 0.1,
        max_tokens: 2000
      })
    });
    
    if (!response.ok) {
      console.error(`[SPELLCHECK] OpenAI API error: ${response.status}`);
      return regexIssues;
    }
    
    const data = await response.json();
    const content = data.choices?.[0]?.message?.content || '[]';
    
    // Parse JSON response
    try {
      const jsonMatch = content.match(/\[[\s\S]*\]/);
      if (jsonMatch) {
        const rawIssues = JSON.parse(jsonMatch[0]);
        console.log(`[AI RAW] Detected ${rawIssues.length} errors`);
        
        // Filtrer les corrections invalides
        aiIssues = filterInvalidCorrections(rawIssues);
      }
    } catch (parseError) {
      console.error('[SPELLCHECK] Failed to parse AI response:', parseError.message);
    }
    
  } catch (error) {
    console.error('[SPELLCHECK] AI call failed:', error.message);
  }
  
  // 3. Fusionner regex + IA, dédupliquer
  const allIssues = [...regexIssues];
  const seenErrors = new Set(regexIssues.map(i => i.error.toLowerCase()));
  
  for (const issue of aiIssues) {
    const errorKey = issue.error.toLowerCase();
    if (!seenErrors.has(errorKey)) {
      allIssues.push({
        error: issue.error,
        correction: issue.correction,
        type: issue.type || 'spelling',
        severity: issue.severity || 'medium',
        message: issue.message || 'AI-detected error'
      });
      seenErrors.add(errorKey);
    }
  }
  
  console.log(`✅ AI spell-check complete: ${allIssues.length} valid errors found`);
  
  // Log des corrections pour debug
  allIssues.slice(0, 20).forEach((issue, i) => {
    console.log(`  ${i + 1}. [${issue.type}] "${issue.error}" → "${issue.correction}" (${issue.message})`);
  });
  
  return allIssues;
}

export default { checkSpellingWithAI };
