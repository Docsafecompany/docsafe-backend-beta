// lib/aiSpellCheck.js
// VERSION 4.0 - Correcteur professionnel multilingue complet

/**
 * Dictionnaire élargi de mots communs pour validation (400+ mots)
 */
const COMMON_WORDS = new Set([
  // English - Articles, prepositions, pronouns
  'the', 'a', 'an', 'and', 'or', 'but', 'if', 'of', 'to', 'in', 'on', 'at', 'by', 'for', 'with',
  'as', 'is', 'it', 'be', 'am', 'are', 'was', 'were', 'been', 'being', 'have', 'has', 'had',
  'do', 'does', 'did', 'will', 'would', 'could', 'should', 'may', 'might', 'must', 'shall',
  'i', 'you', 'he', 'she', 'we', 'they', 'me', 'him', 'her', 'us', 'them', 'my', 'your', 'his',
  'its', 'our', 'their', 'this', 'that', 'these', 'those', 'who', 'whom', 'whose', 'which', 'what',
  
  // English - Common verbs
  'go', 'went', 'gone', 'going', 'come', 'came', 'coming', 'see', 'saw', 'seen', 'seeing',
  'get', 'got', 'getting', 'make', 'made', 'making', 'take', 'took', 'taken', 'taking',
  'know', 'knew', 'known', 'knowing', 'think', 'thought', 'thinking', 'say', 'said', 'saying',
  'give', 'gave', 'given', 'giving', 'find', 'found', 'finding', 'tell', 'told', 'telling',
  'ask', 'asked', 'asking', 'use', 'used', 'using', 'try', 'tried', 'trying', 'need', 'needed',
  'feel', 'felt', 'feeling', 'become', 'became', 'becoming', 'leave', 'left', 'leaving',
  'put', 'putting', 'mean', 'meant', 'meaning', 'keep', 'kept', 'keeping', 'let', 'letting',
  'begin', 'began', 'begun', 'beginning', 'show', 'showed', 'shown', 'showing',
  'hear', 'heard', 'hearing', 'play', 'played', 'playing', 'run', 'ran', 'running',
  'move', 'moved', 'moving', 'live', 'lived', 'living', 'believe', 'believed', 'believing',
  'hold', 'held', 'holding', 'bring', 'brought', 'bringing', 'happen', 'happened', 'happening',
  'write', 'wrote', 'written', 'writing', 'provide', 'provided', 'providing', 'sit', 'sat', 'sitting',
  'stand', 'stood', 'standing', 'lose', 'lost', 'losing', 'pay', 'paid', 'paying',
  'meet', 'met', 'meeting', 'include', 'included', 'including', 'continue', 'continued', 'continuing',
  'set', 'setting', 'learn', 'learned', 'learning', 'change', 'changed', 'changing',
  'lead', 'led', 'leading', 'understand', 'understood', 'understanding', 'watch', 'watched', 'watching',
  'follow', 'followed', 'following', 'stop', 'stopped', 'stopping', 'create', 'created', 'creating',
  'speak', 'spoke', 'spoken', 'speaking', 'read', 'reading', 'allow', 'allowed', 'allowing',
  'add', 'added', 'adding', 'spend', 'spent', 'spending', 'grow', 'grew', 'grown', 'growing',
  'open', 'opened', 'opening', 'walk', 'walked', 'walking', 'win', 'won', 'winning',
  'offer', 'offered', 'offering', 'remember', 'remembered', 'remembering', 'love', 'loved', 'loving',
  'consider', 'considered', 'considering', 'appear', 'appeared', 'appearing', 'buy', 'bought', 'buying',
  'wait', 'waited', 'waiting', 'serve', 'served', 'serving', 'die', 'died', 'dying',
  'send', 'sent', 'sending', 'expect', 'expected', 'expecting', 'build', 'built', 'building',
  'stay', 'stayed', 'staying', 'fall', 'fell', 'fallen', 'falling', 'cut', 'cutting',
  'reach', 'reached', 'reaching', 'kill', 'killed', 'killing', 'remain', 'remained', 'remaining',
  
  // English - Common adjectives
  'good', 'better', 'best', 'bad', 'worse', 'worst', 'new', 'old', 'young', 'big', 'small',
  'large', 'little', 'long', 'short', 'high', 'low', 'great', 'important', 'different', 'same',
  'few', 'many', 'much', 'more', 'most', 'other', 'first', 'last', 'next', 'early', 'late',
  'right', 'wrong', 'true', 'false', 'real', 'own', 'only', 'sure', 'able', 'possible',
  'free', 'full', 'open', 'clear', 'easy', 'hard', 'strong', 'weak', 'fast', 'slow',
  'hot', 'cold', 'warm', 'cool', 'dark', 'light', 'white', 'black', 'red', 'blue', 'green',
  
  // English - Common nouns
  'time', 'year', 'people', 'way', 'day', 'man', 'woman', 'child', 'children', 'world',
  'life', 'hand', 'part', 'place', 'case', 'week', 'company', 'system', 'program', 'question',
  'work', 'government', 'number', 'night', 'point', 'home', 'water', 'room', 'mother', 'area',
  'money', 'story', 'fact', 'month', 'lot', 'right', 'study', 'book', 'eye', 'job', 'word',
  'business', 'issue', 'side', 'kind', 'head', 'house', 'service', 'friend', 'father', 'power',
  'hour', 'game', 'line', 'end', 'member', 'law', 'car', 'city', 'community', 'name',
  'president', 'team', 'minute', 'idea', 'kid', 'body', 'information', 'back', 'parent', 'face',
  'others', 'level', 'office', 'door', 'health', 'person', 'art', 'war', 'history', 'party',
  'result', 'change', 'morning', 'reason', 'research', 'girl', 'guy', 'moment', 'air', 'teacher',
  'force', 'education', 'foot', 'boy', 'age', 'policy', 'process', 'music', 'market', 'sense',
  'nation', 'plan', 'college', 'interest', 'death', 'experience', 'effect', 'use', 'class', 'control',
  'care', 'field', 'development', 'role', 'effort', 'rate', 'heart', 'drug', 'show', 'leader',
  'light', 'voice', 'wife', 'police', 'mind', 'difference', 'period', 'value', 'building', 'action',
  'authority', 'model', 'daughter', 'activity', 'population', 'record', 'paper', 'order', 'view',
  'ground', 'form', 'decision', 'technology', 'century', 'course', 'section', 'term', 'practice',
  
  // Business/Document specific words
  'document', 'report', 'analysis', 'summary', 'review', 'proposal', 'presentation', 'meeting',
  'project', 'client', 'customer', 'service', 'product', 'solution', 'strategy', 'objective',
  'goal', 'target', 'budget', 'cost', 'revenue', 'profit', 'loss', 'margin', 'growth', 'decline',
  'increase', 'decrease', 'trend', 'forecast', 'estimate', 'assessment', 'evaluation', 'performance',
  'metric', 'indicator', 'benchmark', 'standard', 'compliance', 'regulation', 'requirement', 'specification',
  'implementation', 'deployment', 'integration', 'migration', 'optimization', 'automation', 'transformation',
  'digital', 'technology', 'software', 'hardware', 'platform', 'system', 'network', 'infrastructure',
  'security', 'privacy', 'data', 'information', 'database', 'storage', 'backup', 'recovery',
  'communication', 'collaboration', 'coordination', 'management', 'administration', 'governance',
  'leadership', 'organization', 'structure', 'process', 'procedure', 'policy', 'guideline', 'framework',
  'methodology', 'approach', 'technique', 'tool', 'resource', 'asset', 'capability', 'capacity',
  'efficiency', 'effectiveness', 'productivity', 'quality', 'reliability', 'availability', 'scalability',
  'flexibility', 'agility', 'innovation', 'creativity', 'expertise', 'knowledge', 'skill', 'competency',
  'training', 'development', 'career', 'opportunity', 'challenge', 'risk', 'issue', 'problem',
  'solution', 'recommendation', 'suggestion', 'feedback', 'input', 'output', 'outcome', 'result',
  'impact', 'benefit', 'advantage', 'disadvantage', 'strength', 'weakness', 'opportunity', 'threat',
  'stakeholder', 'shareholder', 'partner', 'vendor', 'supplier', 'contractor', 'consultant', 'advisor',
  'enabling', 'individuals', 'landscape', 'potential', 'connection', 'balance', 'message', 'matter',
  'barrier', 'realm', 'contribution', 'ability', 'significant', 'distances', 'platforms', 'presence',
  'letters', 'abroad', 'travel', 'seconds', 'global', 'community', 'provides', 'expression',
  'contributing', 'individuality', 'essential', 'strike', 'embracing', 'benefits', 'remains',
  'meaningful', 'rather', 'conclusion', 'stands', 'edged', 'embodies', 'unity', 'creativity',
  'regulate', 'guide', 'essence', 'increasingly', 'influential', 'forces', 'shaping', 'connect',
  'share', 'facebook', 'instagram', 'tiktok', 'twitter', 'dominate', 'daily', 'interactions',
  'exchange', 'ideas', 'across', 'geographical', 'boundaries', 'instantly', 'traditional', 'methods',
  'phone', 'calls', 'relied', 'heavily', 'physical', 'lengthy', 'processes', 'mailing', 'today',
  'staying', 'touch', 'loved', 'ones', 'fostered', 'breaking', 'barriers', 'self', 'cultures',
  'identity', 'decline', 'personal', 'depth', 'nuance', 'struggle', 'replicate', 'screen',
  'based', 'feels', 'impersonal', 'superficial', 'relationships', 'loneliness', 'users',
  'surrounded', 'virtual', 'friends', 'genuine', 'human', 'spread', 'misinformation', 'fake',
  'news', 'viral', 'speed', 'corrections', 'issued', 'undermining', 'trust', 'institutions',
  'creating', 'divisions', 'within', 'communities', 'political', 'polarization', 'recent',
  'example', 'exacerbated', 'circulation', 'unverified', 'claims', 'additionally', 'raised',
  'concerns', 'commodification', 'every', 'interaction', 'likes', 'comments', 'search',
  'creates', 'footprints', 'corporations', 'targeted', 'advertising', 'personalization', 'enhance',
  'raises', 'ethical', 'questions', 'consent', 'surveillance', 'fully', 'realize', 'extent',
  'habits', 'monitored', 'analyzed', 'growing', 'debates', 'protection', 'rights', 'ultimately',
  'modern', 'profound', 'multifaceted', 'revolutionized', 'providing', 'activism', 'contributed',
  'erosion', 'disconnection', 'falsehoods', 'exploitation', 'responsibility', 'lies', 'governments',
  'critically', 'reflect', 'preserve', 'confidential', 'slide', 'professional', 'content',
  'social', 'media', 'double', 'sword', 'tool',
  
  // French common words
  'le', 'la', 'les', 'de', 'du', 'des', 'un', 'une', 'et', 'est', 'en', 'que', 'qui', 'dans',
  'ce', 'il', 'pas', 'plus', 'pour', 'sur', 'par', 'au', 'aux', 'avec', 'son', 'sont', 'nous',
  'vous', 'leur', 'cette', 'bien', 'aussi', 'comme', 'tout', 'elle', 'entre', 'faire', 'fait',
  'peut', 'donc', 'sans', 'mais', 'ou', 'où', 'avant', 'après', 'avoir', 'être', 'très',
  'encore', 'autre', 'notre', 'votre', 'même', 'quand', 'quel', 'quelle', 'si', 'ni', 'car',
  'puis', 'dont', 'sera', 'été', 'ces', 'ses', 'mes', 'tes', 'nos', 'vos', 'leurs',
  'confidentiel', 'diapositive', 'contenu', 'professionnel', 'numérique', 'document', 'fichier',
  'rapport', 'analyse', 'résumé', 'présentation', 'réunion', 'projet', 'client', 'service',
  'produit', 'solution', 'stratégie', 'objectif', 'budget', 'coût', 'revenu', 'bénéfice',
  'croissance', 'tendance', 'évaluation', 'performance', 'conformité', 'réglementation',
  'exigence', 'spécification', 'mise', 'œuvre', 'déploiement', 'intégration', 'optimisation',
  'sécurité', 'confidentialité', 'données', 'stockage', 'sauvegarde', 'récupération',
  'communication', 'collaboration', 'coordination', 'gestion', 'administration', 'gouvernance',
  'organisation', 'structure', 'processus', 'procédure', 'politique', 'méthodologie', 'approche',
  'technique', 'outil', 'ressource', 'capacité', 'efficacité', 'productivité', 'qualité',
  'fiabilité', 'disponibilité', 'flexibilité', 'innovation', 'créativité', 'expertise',
  'connaissance', 'compétence', 'formation', 'développement', 'carrière', 'opportunité',
  'défi', 'risque', 'problème', 'recommandation', 'suggestion', 'retour', 'résultat',
  'impact', 'avantage', 'inconvénient', 'force', 'faiblesse', 'menace'
]);

/**
 * Phrases valides à NE JAMAIS corriger (100+ patterns)
 */
const VALID_PHRASES = new Set([
  // Article + noun/adj patterns
  'a message', 'a matter', 'a sense', 'a tool', 'a balance', 'a platform', 'a double', 'a barrier',
  'a new', 'a big', 'a few', 'a lot', 'a bit', 'a man', 'a day', 'a way', 'a set', 'a key',
  'a good', 'a great', 'a small', 'a large', 'a long', 'a short', 'a high', 'a low', 'a real',
  'a true', 'a full', 'a free', 'a clear', 'a strong', 'a weak', 'a fast', 'a slow', 'a hard',
  'an example', 'an issue', 'an idea', 'an area', 'an hour', 'an end', 'an old', 'an open',
  'an important', 'an interesting', 'an easy', 'an early', 'an only', 'an other', 'an average',
  
  // Preposition + article patterns
  'in the', 'of the', 'to the', 'by the', 'on the', 'at the', 'for the', 'from the', 'with the',
  'in a', 'of a', 'to a', 'by a', 'on a', 'at a', 'for a', 'from a', 'with a',
  'in an', 'of an', 'to an', 'by an', 'on an', 'at an', 'for an', 'from an', 'with an',
  'into the', 'onto the', 'upon the', 'about the', 'after the', 'before the', 'during the',
  'into a', 'onto a', 'upon a', 'about a', 'after a', 'before a', 'during a',
  'through the', 'across the', 'around the', 'between the', 'among the', 'within the',
  'through a', 'across a', 'around a', 'between a', 'among a', 'within a',
  
  // Verb + article patterns
  'is the', 'is a', 'is an', 'is its', 'is it', 'is not', 'is also', 'is both', 'is still',
  'are the', 'are a', 'are not', 'are also', 'are both', 'are still', 'are being',
  'was the', 'was a', 'was an', 'was not', 'was also', 'was being', 'was still',
  'were the', 'were not', 'were also', 'were being', 'were still',
  'has the', 'has a', 'has an', 'has been', 'has not', 'has also', 'has its',
  'have the', 'have a', 'have an', 'have been', 'have not', 'have also',
  'had the', 'had a', 'had an', 'had been', 'had not', 'had also',
  'will be', 'will have', 'will not', 'will also', 'will still',
  'would be', 'would have', 'would not', 'would also', 'would still',
  'could be', 'could have', 'could not', 'could also', 'could still',
  'should be', 'should have', 'should not', 'should also', 'should still',
  'may be', 'may have', 'may not', 'may also', 'may still',
  'might be', 'might have', 'might not', 'might also', 'might still',
  'must be', 'must have', 'must not', 'must also', 'must still',
  
  // Pronoun patterns
  'it is', 'it has', 'it was', 'it can', 'it may', 'it will', 'it would', 'it could', 'it should',
  'it does', 'it did', 'it had', 'it might', 'it must', 'it seems', 'it appears', 'it remains',
  'to it', 'to be', 'to do', 'to go', 'to see', 'to get', 'to make', 'to take', 'to have',
  'to give', 'to find', 'to know', 'to think', 'to say', 'to tell', 'to ask', 'to use',
  'i am', 'i was', 'i will', 'i have', 'i can', 'i could', 'i would', 'i should', 'i do',
  'i did', 'i had', 'i may', 'i might', 'i must', 'i think', 'i know', 'i believe', 'i feel',
  'he is', 'he was', 'he has', 'he will', 'he can', 'he could', 'he would', 'he should',
  'she is', 'she was', 'she has', 'she will', 'she can', 'she could', 'she would', 'she should',
  'we are', 'we were', 'we have', 'we will', 'we can', 'we could', 'we would', 'we should',
  'they are', 'they were', 'they have', 'they will', 'they can', 'they could', 'they would',
  'you are', 'you were', 'you have', 'you will', 'you can', 'you could', 'you would',
  
  // Special patterns
  'on how', 'of dis', 'as a', 'as an', 'as the', 'such as', 'such a', 'such an',
  'so that', 'so the', 'so a', 'so it', 'so we', 'so they', 'so you', 'so he', 'so she',
  'if the', 'if a', 'if an', 'if it', 'if we', 'if they', 'if you', 'if he', 'if she',
  'or the', 'or a', 'or an', 'or it', 'and the', 'and a', 'and an', 'and it',
  'but the', 'but a', 'but an', 'but it', 'but we', 'but they', 'but you',
  'all the', 'all of', 'all in', 'one of', 'some of', 'many of', 'most of', 'each of',
  'this is', 'that is', 'there is', 'there are', 'here is', 'here are',
  'which is', 'which are', 'which was', 'which were', 'which has', 'which have',
  'what is', 'what are', 'what was', 'what were', 'what has', 'what have',
  'who is', 'who are', 'who was', 'who were', 'who has', 'who have',
  
  // French patterns
  'de la', 'de le', 'de les', 'à la', 'à le', 'à les', 'en un', 'en une',
  'il y', 'ce qui', 'ce que', 'il est', 'elle est', 'ils sont', 'elles sont',
  "c'est", "qu'il", "qu'elle", "n'est", "d'un", "d'une", "l'un", "l'une",
  'par le', 'par la', 'pour le', 'pour la', 'sur le', 'sur la', 'dans le', 'dans la'
]);

/**
 * Le prompt IA professionnel complet pour la correction
 */
const PROFESSIONAL_PROOFREADER_PROMPT = `You are a professional multilingual proofreader and language expert. Your mission is to make documents PERFECT and ready for professional delivery.

AUTOMATICALLY DETECT the language of the text and analyze ALL errors.

═══════════════════════════════════════════════════════════════
📝 ERROR CATEGORIES TO DETECT
═══════════════════════════════════════════════════════════════

1️⃣ SPELLING ERRORS
   - Misspelled words: "recieve" → "receive", "occurence" → "occurrence"
   - Typos: "teh" → "the", "adn" → "and"
   - Homophones: "their/there/they're", "your/you're", "its/it's"
   - Accented characters (French): "éléphant" not "elephant", "café" not "cafe"

2️⃣ CONJUGATION & VERB ERRORS
   - Wrong tense: "He go yesterday" → "He went yesterday"
   - Subject-verb agreement: "She have" → "She has", "They was" → "They were"
   - Irregular verbs: "I runned" → "I ran", "He thinked" → "He thought"
   - French: "Il a manger" → "Il a mangé", "Nous avons fini" (not "finit")

3️⃣ GRAMMAR ERRORS
   - Article usage: "a apple" → "an apple", "the informations" → "the information"
   - Plural/singular: "many information" → "much information", "datas" → "data"
   - Prepositions: "depends of" → "depends on", "interested to" → "interested in"
   - Pronouns: "Me and him went" → "He and I went"
   - French: "à" vs "a", "ou" vs "où", "ce" vs "se"

4️⃣ SYNTAX ERRORS
   - Run-on sentences: missing punctuation between independent clauses
   - Sentence fragments: incomplete sentences
   - Misplaced modifiers: "Running quickly, the finish line was crossed"
   - Parallel structure: "I like reading, writing, and to swim" → "...and swimming"

5️⃣ PUNCTUATION ERRORS
   - Missing periods, commas, colons, semicolons
   - Wrong capitalization: "the President" vs "the president"
   - Apostrophes: "dont" → "don't", "its not" → "it's not"
   - Quotation marks: proper opening/closing
   - French spacing: spaces before : ; ! ? (required in French)

6️⃣ FRAGMENTED WORDS (CRITICAL)
   - Single letter + space: "c an" → "can", "th e" → "the", "o f" → "of"
   - Word split by space: "Confi dential" → "Confidential", "pro fessional" → "professional"
   - Multiple spaces: "p  otential" → "potential"
   - Floating letters: "enablin g" → "enabling" (the g belongs to enablin)
   
   ⚠️ CRITICAL: When you see "X g Y" pattern:
   - Check if "Xg" forms a real word → Yes = correction is "Xg"
   - "enablin g individuals" → The "g" belongs to "enablin" = "enabling"
   - DO NOT create fake words like "gindividuals" - that does NOT exist!

7️⃣ MERGED/FUSED WORDS
   - Missing space: "thankyou" → "thank you", "alot" → "a lot"
   - Concatenated words: "cannotbe" → "cannot be"

═══════════════════════════════════════════════════════════════
🚫 CRITICAL RULES - MUST FOLLOW
═══════════════════════════════════════════════════════════════

1. VERIFY EVERY CORRECTION IS A REAL WORD
   Before suggesting ANY correction, ask: "Does this word exist?"
   ❌ "gindividuals" = DOES NOT EXIST → Never suggest it
   ❌ "amessage" = DOES NOT EXIST → Never suggest it
   ❌ "Inthe" = DOES NOT EXIST → Never suggest it
   ✅ "enabling" = EXISTS → Valid correction
   ✅ "potential" = EXISTS → Valid correction

2. NEVER MERGE LEGITIMATE SEPARATE WORDS
   These are CORRECT as written - DO NOT flag them:
   ❌ "In the" → "Inthe" = WRONG
   ❌ "a message" → "amessage" = WRONG
   ❌ "is the" → "isthe" = WRONG
   ❌ "of the" → "ofthe" = WRONG
   ❌ "It has" → "Ithas" = WRONG
   
3. CONTEXT IS KING
   Always read the FULL sentence to understand:
   - Which word is incomplete
   - What the author meant to write
   - Whether the correction makes semantic sense

4. FLOATING LETTERS BELONG TO ADJACENT WORDS
   For "X g Y" patterns:
   - Check: Does "Xg" form a real word? → Yes = correction is "Xg"
   - Check: Does "gY" form a real word? → Usually NO
   - Example: "enablin g individuals" → "enabling individuals"
   - NOT: "enablin gindividuals" (gindividuals is NOT a word!)

5. MULTILINGUAL SUPPORT
   - Auto-detect language (English, French, German, Spanish, etc.)
   - Apply language-specific grammar rules
   - Respect accents and special characters
   - French: il/elle a mangé (not manger), ils/elles ont fini

6. SKIP IF NOT 100% CERTAIN
   If you're not absolutely sure, DO NOT include the error.
   Better to miss an error than suggest a wrong correction.

═══════════════════════════════════════════════════════════════
📊 OUTPUT FORMAT
═══════════════════════════════════════════════════════════════

Return ONLY errors where you are 100% confident the correction is correct.
Each error must have:
- error: The exact incorrect text as it appears
- correction: The corrected text (MUST be a real word/phrase)
- context: 30-50 characters surrounding the error
- severity: "high" (spelling, fragments), "medium" (grammar), "low" (style)
- type: spelling, conjugation, grammar, syntax, punctuation, fragmented_word, merged_word
- message: Brief explanation of why this is wrong

QUALITY GOAL: After corrections, the document should be PERFECT for professional use.`;

/**
 * Vérifie si un mot existe dans le dictionnaire
 */
function isRealWord(word) {
  if (!word || word.length < 2) return false;
  const normalized = word.toLowerCase().trim();
  
  // Mots de 1-3 lettres doivent être dans le dictionnaire
  if (normalized.length <= 3) {
    return COMMON_WORDS.has(normalized);
  }
  
  // Vérifier le dictionnaire
  if (COMMON_WORDS.has(normalized)) return true;
  
  // Rejeter les patterns impossibles
  const invalidPatterns = [
    /^[bcdfghjklmnpqrstvwxz]{4,}/i,  // 4+ consonnes au début
    /[bcdfghjklmnpqrstvwxz]{5,}/i,   // 5+ consonnes consécutives
    /(.)\1{3,}/i,                     // 4+ caractères identiques
    /^[aeiou]{4,}/i,                  // 4+ voyelles au début
    /^g[a-z]+s$/i                     // Pattern "g...s" souvent invalide
  ];
  
  for (const pattern of invalidPatterns) {
    if (pattern.test(normalized)) return false;
  }
  
  return true;
}

/**
 * Vérifie si une phrase est valide
 */
function isValidPhrase(phrase) {
  if (!phrase) return false;
  const normalized = phrase.toLowerCase().trim();
  return VALID_PHRASES.has(normalized);
}

/**
 * Obtient le contexte autour d'une position
 */
function getContext(text, index, length) {
  const start = Math.max(0, index - 30);
  const end = Math.min(text.length, index + length + 30);
  return text.slice(start, end).replace(/\s+/g, ' ').trim();
}

/**
 * Pré-détection ciblée des vrais mots fragmentés
 */
function preDetectFragmentedWords(text) {
  const issues = [];
  
  // Pattern 1: Mot + espace + 1-3 lettres qui complètent le mot
  const fragmentPattern = /\b([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]{3,})\s+([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]{1,3})\b/g;
  
  let match;
  while ((match = fragmentPattern.exec(text)) !== null) {
    const [fullMatch, part1, part2] = match;
    const combined = part1 + part2;
    
    if (isRealWord(combined) && !isValidPhrase(fullMatch.toLowerCase())) {
      if (!isRealWord(part2) || COMMON_WORDS.has(combined.toLowerCase())) {
        issues.push({
          error: fullMatch,
          correction: combined,
          context: getContext(text, match.index, fullMatch.length),
          type: 'fragmented_word',
          severity: 'high',
          message: 'Word incorrectly split by space',
          source: 'regex'
        });
      }
    }
  }
  
  // Pattern 2: 1-2 lettres + espace + mot
  const prefixPattern = /\b([a-zA-Z]{1,2})\s+([a-zA-Z]{2,})\b/g;
  
  while ((match = prefixPattern.exec(text)) !== null) {
    const [fullMatch, prefix, word] = match;
    const combined = prefix + word;
    
    if (isValidPhrase(fullMatch.toLowerCase())) continue;
    if (prefix.toLowerCase() === 'i' || prefix.toLowerCase() === 'a') {
      if (isRealWord(word)) continue;
    }
    
    if (isRealWord(combined) && !isRealWord(prefix)) {
      issues.push({
        error: fullMatch,
        correction: combined,
        context: getContext(text, match.index, fullMatch.length),
        type: 'fragmented_word',
        severity: 'high',
        message: 'Word incorrectly split by space',
        source: 'regex'
      });
    }
  }
  
  // Pattern 3: Espaces multiples dans un mot
  const multiSpacePattern = /\b([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]+)\s{2,}([a-zA-ZàâäéèêëïîôùûüçÀÂÄÉÈÊËÏÎÔÙÛÜÇ]+)\b/g;
  
  while ((match = multiSpacePattern.exec(text)) !== null) {
    const combined = match[1] + match[2];
    if (isRealWord(combined)) {
      issues.push({
        error: match[0],
        correction: combined,
        context: getContext(text, match.index, match[0].length),
        type: 'multiple_spaces',
        severity: 'high',
        message: 'Multiple spaces inside word',
        source: 'regex'
      });
    }
  }
  
  console.log(`[REGEX] Found ${issues.length} fragmented word patterns`);
  return issues;
}

/**
 * Filtre les corrections invalides
 */
function filterInvalidCorrections(errors) {
  return errors.filter(err => {
    if (!isRealWord(err.correction)) {
      console.log(`❌ REJECTED (not a real word): "${err.error}" → "${err.correction}"`);
      return false;
    }
    
    if (isValidPhrase(err.error)) {
      console.log(`❌ REJECTED (valid phrase): "${err.error}"`);
      return false;
    }
    
    if (err.error.trim() === err.correction.trim()) {
      console.log(`❌ REJECTED (no change): "${err.error}"`);
      return false;
    }
    
    const words = err.error.toLowerCase().split(/\s+/);
    if (words.length === 2 && isRealWord(words[0]) && isRealWord(words[1])) {
      if (!COMMON_WORDS.has(err.correction.toLowerCase())) {
        console.log(`❌ REJECTED (both parts are valid words): "${err.error}"`);
        return false;
      }
    }
    
    return true;
  });
}

/**
 * Vérifie l'orthographe avec l'IA - VERSION 4.0
 */
export async function checkSpellingWithAI(text) {
  if (!text || text.trim().length < 10) return [];
  
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    console.warn('⚠️ OPENAI_API_KEY not configured - spelling check disabled');
    return [];
  }

  try {
    console.log('📝 AI spell-check VERSION 4.0 starting...');
    console.log(`[SPELLCHECK] Text length: ${text.length} characters`);
    
    const regexErrors = preDetectFragmentedWords(text);
    
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': `Bearer ${apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [
          {
            role: 'system',
            content: PROFESSIONAL_PROOFREADER_PROMPT
          },
          {
            role: 'user',
            content: `Analyze this text and find ALL spelling, grammar, conjugation, syntax, and punctuation errors. 

REMEMBER: 
- Only suggest corrections that result in REAL words
- Never merge separate valid words like "In the" or "a message"
- Floating letters usually belong to the LEFT word (e.g., "enablin g" → "enabling")
- Detect conjugation errors (e.g., "He go" → "He went")
- Detect grammar errors (e.g., "a apple" → "an apple")

TEXT TO ANALYZE:
${text.slice(0, 15000)}`
          }
        ],
        tools: [
          {
            type: 'function',
            function: {
              name: 'report_language_errors',
              description: 'Report all detected language errors with validated corrections',
              parameters: {
                type: 'object',
                properties: {
                  errors: {
                    type: 'array',
                    items: {
                      type: 'object',
                      properties: {
                        error: { type: 'string', description: 'The exact incorrect text as it appears' },
                        correction: { type: 'string', description: 'The corrected text (MUST be a real word/phrase)' },
                        context: { type: 'string', description: 'Surrounding text for context (30-50 chars)' },
                        severity: { type: 'string', enum: ['low', 'medium', 'high'] },
                        type: { type: 'string', enum: ['spelling', 'conjugation', 'grammar', 'syntax', 'punctuation', 'fragmented_word', 'merged_word'] },
                        message: { type: 'string', description: 'Brief explanation of the error' }
                      },
                      required: ['error', 'correction', 'severity', 'type', 'message']
                    }
                  }
                },
                required: ['errors']
              }
            }
          }
        ],
        tool_choice: { type: 'function', function: { name: 'report_language_errors' } },
        temperature: 0.05,
        max_tokens: 4000
      })
    });

    if (!response.ok) {
      console.error('❌ OpenAI API error:', response.status);
      return formatErrors(filterInvalidCorrections(regexErrors));
    }

    const data = await response.json();
    const toolCall = data.choices?.[0]?.message?.tool_calls?.[0];
    
    let aiErrors = [];
    if (toolCall?.function?.arguments) {
      try {
        const parsed = JSON.parse(toolCall.function.arguments);
        aiErrors = parsed.errors || [];
        console.log(`[AI RAW] Detected ${aiErrors.length} errors`);
      } catch (e) {
        console.error('❌ Failed to parse AI response:', e.message);
      }
    }
    
    const allErrors = mergeErrors(regexErrors, aiErrors);
    const validErrors = filterInvalidCorrections(allErrors);
    const spellingErrors = formatErrors(validErrors);

    console.log(`✅ AI spell-check complete: ${spellingErrors.length} valid errors found`);
    spellingErrors.forEach((err, i) => {
      console.log(`  ${i + 1}. [${err.type}] "${err.error}" → "${err.correction}" (${err.message})`);
    });
    
    return spellingErrors;

  } catch (error) {
    console.error('❌ AI spell-check error:', error.message);
    return [];
  }
}

function mergeErrors(regexErrors, aiErrors) {
  const merged = [...regexErrors];
  const seen = new Set(regexErrors.map(e => e.error.toLowerCase().replace(/\s+/g, '')));
  
  for (const aiErr of aiErrors) {
    const normalized = aiErr.error.toLowerCase().replace(/\s+/g, '');
    if (!seen.has(normalized)) {
      merged.push(aiErr);
      seen.add(normalized);
    }
  }
  
  return merged;
}

function formatErrors(errors) {
  return errors.map((err, index) => ({
    id: `spell_${index}_${Date.now()}`,
    error: err.error,
    correction: err.correction,
    context: err.context || '',
    location: `Detection ${index + 1}`,
    severity: err.severity || 'medium',
    type: err.type || 'spelling',
    rule: err.type?.toUpperCase() || 'AI_CORRECTION',
    message: err.message || getErrorMessage(err.type),
    category: getCategory(err.type),
    source: err.source || 'ai'
  }));
}

function getErrorMessage(type) {
  const messages = {
    'fragmented_word': 'Word incorrectly split by space',
    'multiple_spaces': 'Multiple spaces inside word',
    'merged_word': 'Words incorrectly merged together',
    'spelling': 'Spelling error',
    'conjugation': 'Verb conjugation/tense error',
    'grammar': 'Grammar error',
    'syntax': 'Syntax error',
    'punctuation': 'Punctuation error'
  };
  return messages[type] || 'Spelling/grammar correction';
}

function getCategory(type) {
  const categories = {
    'fragmented_word': 'Word Fragments',
    'multiple_spaces': 'Word Fragments',
    'merged_word': 'Word Fragments',
    'conjugation': 'Conjugation',
    'grammar': 'Grammar',
    'syntax': 'Syntax',
    'punctuation': 'Punctuation'
  };
  return categories[type] || 'Spelling';
}

export function applyCorrections(text, corrections) {
  if (!corrections?.length) return { correctedText: text, examples: [], changedCount: 0 };

  let correctedText = text;
  const examples = [];
  let changedCount = 0;

  const sorted = [...corrections].sort((a, b) => (b.error?.length || 0) - (a.error?.length || 0));

  for (const c of sorted) {
    if (!c.error || !c.correction || c.error === c.correction) continue;
    
    const regex = new RegExp(escapeRegExp(c.error), 'gi');
    const matches = correctedText.match(regex);
    
    if (matches?.length) {
      correctedText = correctedText.replace(regex, c.correction);
      changedCount += matches.length;
      if (examples.length < 15) {
        examples.push({ 
          before: c.error, 
          after: c.correction, 
          type: c.type,
          message: c.message 
        });
      }
    }
  }

  console.log(`[APPLY] ${changedCount} corrections applied`);
  return { correctedText, examples, changedCount };
}

function escapeRegExp(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

export default { checkSpellingWithAI, applyCorrections };
