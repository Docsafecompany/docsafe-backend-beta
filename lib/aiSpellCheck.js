// lib/aiSpellCheck.js
// VERSION 5.1 - Strict + Context-aware reconstruction (safe-merge)
// Corrige : "Confi dential" → "Confidential", "sl ide" → "slide"
// Ajoute : correction mots évidents même si NON dans COMMON_WORDS (ex: "gen uine" → "genuine")

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

// ============= HELPERS =============

/**
 * STRICT: true UNIQUEMENT si le mot est dans COMMON_WORDS
 */
function isRealWord(word) {
  if (!word || word.length < 2) return false;
  const normalized = word.toLowerCase().trim();
  return COMMON_WORDS.has(normalized);
}

/**
 * Heuristique "safe" pour accepter un mot reconstruit même s'il n'est pas dans COMMON_WORDS
 * (sinon tu dois maintenir un dictionnaire infini)
 */
function isPlausibleWord(word) {
  if (!word) return false;
  const w = word.trim();

  // uniquement lettres (pas de chiffres/punct)
  if (!/^[A-Za-z]+$/.test(w)) return false;

  // longueur raisonnable
  if (w.length < 3 || w.length > 30) return false;

  // éviter trucs genre "aaaaaa" / "xxxxxx"
  if (/^(.)\1{4,}$/.test(w.toLowerCase())) return false;

  return true;
}

/**
 * Vérifie si une phrase est valide et ne doit pas être fusionnée
 */
function isValidPhrase(phrase) {
  const normalized = phrase.toLowerCase().trim();
  if (VALID_PHRASES.has(normalized)) return true;

  const parts = normalized.split(/\s+/);
  if (parts.length === 2) {
    const [first, second] = parts;

    const shortWords = [
      'a', 'an', 'the', 'to', 'of', 'in', 'on', 'at', 'by', 'for', 'with',
      'is', 'are', 'was', 'were', 'be', 'can', 'could', 'would', 'should', 'will',
      'may', 'might', 'must', 'have', 'has', 'had', 'do', 'does', 'did',
      'not', 'how', 'one', 'due'
    ];

    if (shortWords.includes(second) && isRealWord(first)) {
      return true;
    }

    if (isRealWord(first) && isRealWord(second)) {
      const combined = first + second;
      if (!isRealWord(combined)) {
        return true;
      }
    }
  }

  return false;
}

/**
 * Détermine si on peut "safe-merge" une erreur du type "gen uine" → "genuine"
 * sans risquer de casser des phrases valides.
 */
function canSafeMerge(error, correction) {
  const err = (error || '').trim();
  const corr = (correction || '').trim();

  const errorHasSpace = /\s/.test(err);
  const correctionHasNoSpace = !/\s/.test(corr);
  if (!errorHasSpace || !correctionHasNoSpace) return false;

  // jamais si c'est une phrase valide
  if (isValidPhrase(err.toLowerCase())) return false;

  const parts = err.split(/\s+/).filter(Boolean);
  if (parts.length !== 2) return false;

  const [p1, p2] = parts;

  // si les 2 parties sont des mots réels, on évite (risque de merger une vraie phrase)
  if (isRealWord(p1) && isRealWord(p2)) return false;

  // correction plausible
  if (!isPlausibleWord(corr)) return false;

  return true;
}

// ============= PRÉ-DÉTECTION REGEX =============

/**
 * Pré-détection des mots fragmentés par regex AVANT l'appel IA
 * VERSION 5.1 - Accepte les reconstructions plausibles même hors dictionnaire
 */
function preDetectFragmentedWords(text) {
  const issues = [];

  // Pattern 1: Mot coupé par un espace (ex: "soc ial", "enablin g")
  const fragmentPattern = /\b([a-zA-Z]{2,})[ ]+([a-zA-Z]{2,12})\b/g;
  let match;

  while ((match = fragmentPattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const part1 = match[1];
    const part2 = match[2];
    const combined = part1 + part2;

    // 1) si mot existe dans dict -> OK
    // 2) sinon, si c'est un mot plausible ET pas une phrase valide -> OK
    const okByDict = isRealWord(combined);
    const okBySafe = canSafeMerge(fullMatch, combined);

    if ((okByDict || okBySafe) && !isValidPhrase(fullMatch.toLowerCase())) {
      issues.push({
        error: fullMatch,
        correction: combined.charAt(0).toUpperCase() + combined.slice(1).toLowerCase(),
        type: 'fragmented_word',
        severity: 'high',
        message: `Word incorrectly split by space: "${fullMatch}" → "${combined}"`
      });
    }
  }

  // Pattern 2: Lettre isolée + mot (ex: "th e" → "the", "o f" → "of")
  const singleLetterPattern = /\b([a-zA-Z])[ ]+([a-zA-Z]{1,15})\b/g;

  while ((match = singleLetterPattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const letter = match[1];
    const rest = match[2];
    const combined = letter + rest;

    const okByDict = isRealWord(combined);
    const okBySafe = canSafeMerge(fullMatch, combined);

    if ((okByDict || okBySafe) && !isValidPhrase(fullMatch.toLowerCase())) {
      issues.push({
        error: fullMatch,
        correction: combined.toLowerCase(),
        type: 'fragmented_word',
        severity: 'high',
        message: `Single letter fragment: "${fullMatch}" → "${combined}"`
      });
    }
  }

  // Pattern 3: Mot + lettre isolée (ex: "essenc e" → "essence")
  const trailingLetterPattern = /\b([a-zA-Z]{2,})[ ]+([a-zA-Z])\b/g;

  while ((match = trailingLetterPattern.exec(text)) !== null) {
    const fullMatch = match[0];
    const word = match[1];
    const letter = match[2];
    const combined = word + letter;

    const okByDict = isRealWord(combined);
    const okBySafe = canSafeMerge(fullMatch, combined);

    if ((okByDict || okBySafe) && !isValidPhrase(fullMatch.toLowerCase())) {
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

    const okByDict = isRealWord(combined);
    const okBySafe = canSafeMerge(fullMatch, combined);

    if (okByDict || okBySafe) {
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

// ============= FILTRAGE DES CORRECTIONS IA =============

/**
 * Filtre les corrections invalides de l'IA
 * VERSION 5.1 - Autorise safe-merge pour mots fragmentés (même hors dictionnaire)
 */
function filterInvalidCorrections(corrections) {
  return (corrections || []).filter(c => {
    const error = (c.error || '').trim();
    const correction = (c.correction || '').trim();
    const type = (c.type || '').trim();

    if (!error || !correction) return false;

    // Rejeter si pas de changement réel
    if (error.toLowerCase() === correction.toLowerCase()) {
      console.log(`❌ REJECTED (no change): "${error}" → "${correction}"`);
      return false;
    }

    // Ne jamais merger une phrase valide
    if (isValidPhrase(error.toLowerCase())) {
      console.log(`❌ REJECTED (valid phrase): "${error}"`);
      return false;
    }

    const errorHasSpace = /\s/.test(error);
    const correctionHasNoSpace = !/\s/.test(correction);

    // Cas de fusion mot cassé
    if (errorHasSpace && correctionHasNoSpace) {
      // 1) OK si dict
      if (isRealWord(correction)) {
        console.log(`✅ ACCEPTED (dict merge): "${error}" → "${correction}"`);
        return true;
      }

      // 2) OK si safe-merge (mots fragmentés) même si hors dict
      if (type === 'fragmented_word' || type === 'multiple_spaces') {
        if (canSafeMerge(error, correction)) {
          console.log(`✅ ACCEPTED (safe merge): "${error}" → "${correction}"`);
          return true;
        }
        console.log(`❌ REJECTED (unsafe merge): "${error}" → "${correction}"`);
        return false;
      }

      // sinon rejet
      console.log(`❌ REJECTED (merge not allowed): "${error}" → "${correction}"`);
      return false;
    }

    // Capitalization / punctuation / spelling : on accepte (pas besoin dict)
    // mais on évite corrections "bizarres"
    if (type === 'capitalization') {
      if (!/[A-Za-z]/.test(correction)) return false;
      console.log(`✅ ACCEPTED: "${error}" → "${correction}"`);
      return true;
    }

    if (type === 'punctuation' || type === 'spelling' || type === 'grammar') {
      // petite sécurité : ne pas transformer en chaîne vide
      if (correction.length < 1) return false;
      console.log(`✅ ACCEPTED: "${error}" → "${correction}"`);
      return true;
    }

    // fallback: accepter si raisonnable
    console.log(`✅ ACCEPTED (fallback): "${error}" → "${correction}"`);
    return true;
  });
}

// ============= PROMPT IA (CONTEXT-AWARE) =============

const PROFESSIONAL_PROOFREADER_PROMPT = `
You are an expert English academic proofreader.

Your job is to detect and propose corrections for REAL errors only, using full sentence context.
This is NOT a simple spellchecker: you must reconstruct intended words when the text is clearly corrupted.

CRITICAL PRIORITIES (must catch these):
A) Fragmented words (spaces inside a single intended word)
- Examples: "soc ial"→"social", "enablin g"→"enabling", "gen uine"→"genuine", "dar ker"→"darker"
- Also multiple spaces inside a word: "p  otential"→"potential", "dis   connection"→"disconnection"
- Single-letter fragments: "th e"→"the", "o f"→"of"

B) Corrupted typos where the intended word is obvious in context
- Extra letters: "correctionnns"→"corrections" or "correction"
- Keyboard/near typos: "searcrh"→"search"
- Repeated garbage prefix/suffix: "gggdigital"→"digital"
- Punctuation inserted inside words: "corpo,rations"→"corporations"

C) Capitalization that is clearly incorrect
- Sentence start must be capitalized.
- Proper nouns must be capitalized (names, brands, countries).
- Titles should be in Title Case (e.g., "The Impact of Social Media on Modern Communication").

STRICT RULES:
1) Do NOT invent content, do NOT paraphrase, do NOT rewrite sentences.
   Only fix spelling/spacing/capitalization/punctuation errors.

2) Do NOT merge valid separate words.
   If it is clearly two intended words, keep them separate.
   Example: "become one", "message can", "strike a", "largely due" must remain two words.

3) Context-aware reconstruction:
   If a word is split or corrupted but the intended word is obvious from meaning, you MUST correct it.
   Do not leave obvious errors uncorrected.

4) Return corrections as small exact spans (shortest possible "error" string) so replacement is safe.
   The "error" must match the exact substring in the input.

OUTPUT FORMAT:
Return ONLY a valid JSON array of objects with exactly this schema:
[
  {
    "error": "exact substring from input",
    "correction": "corrected substring",
    "type": "fragmented_word|spelling|capitalization|punctuation|grammar",
    "severity": "high|medium|low",
    "message": "short reason"
  }
]

QUALITY GATE:
Before outputting, re-check your own corrections:
- No obvious split-word errors should remain unfixed if present in the text.
- Do not propose merges of valid phrases.

Return ONLY the JSON array, no other text.
`;

// ============= FONCTION PRINCIPALE =============

/**
 * Analyse orthographique avec IA + pré-détection regex
 * @param {string} text - Texte à analyser
 * @returns {Promise<Array>} Liste des erreurs détectées
 */
export async function checkSpellingWithAI(text) {
  console.log('📝 AI spell-check VERSION 5.1 starting...');

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
  const seenErrors = new Set(regexIssues.map(i => (i.error || '').toLowerCase()));

  for (const issue of aiIssues) {
    const errorKey = (issue.error || '').toLowerCase();
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
  allIssues.slice(0, 25).forEach((issue, i) => {
    console.log(`  ${i + 1}. [${issue.type}] "${issue.error}" → "${issue.correction}" (${issue.message})`);
  });

  return allIssues;
}

export default { checkSpellingWithAI };
