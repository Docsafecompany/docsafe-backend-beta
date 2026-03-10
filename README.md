# Mindorion — Spécification Produit & Architecture Backend

> **Version** : 1.0 — Mars 2026
> **Produits** : Qualion For All · Qualion Proposal
> **Stack** : Node.js · Express · Render · Supabase · OpenAI

---

## Table des matières

1. [Vue d'ensemble](#1-vue-densemble)
2. [Qualion For All — Détections](#2-qualion-for-all--détections)
3. [Qualion Proposal — Détections](#3-qualion-proposal--détections)
4. [Taxonomie Qualion Proposal](#4-taxonomie-qualion-proposal)
5. [Format affiché — Qualion For All](#5-format-affiché--qualion-for-all)
6. [Format affiché — Qualion Proposal](#6-format-affiché--qualion-proposal)
7. [Format de téléchargement — Qualion For All](#7-format-de-téléchargement--qualion-for-all)
8. [Format de téléchargement — Qualion Proposal](#8-format-de-téléchargement--qualion-proposal)
9. [Architecture backend](#9-architecture-backend)
10. [Modèle de données Supabase](#10-modèle-de-données-supabase)
11. [Priorités d'implémentation](#11-priorités-dimplémentation)

---

## 1. Vue d'ensemble

Mindorion est une plateforme SaaS d'analyse de documents professionnels avant envoi client. Elle expose deux outils distincts qui partagent une infrastructure commune.

| Outil | Objectif | Moteur |
|---|---|---|
| **Qualion For All** | Hygiène documentaire | 100% déterministe (règles) |
| **Qualion Proposal** | Hygiène + analyse business d'une proposition | Déterministe + IA structurée |

**Formats supportés** : `.docx` · `.pptx` · `.xlsx` · `.pdf`

**Principe fondamental** : l'IA n'intervient jamais là où une règle déterministe suffit. Les règles sont la source de vérité. L'IA ajoute de l'interprétation et de la profondeur business.

---

## 2. Qualion For All — Détections

### Objectif

Vérifier l'hygiène documentaire d'un fichier avant envoi externe. Résultat attendu : déterministe, rapide, fiable.

### A. Métadonnées exposées

- Nom de l'auteur
- Nom de l'entreprise
- Dernier éditeur
- Date de création / modification
- Logiciel utilisé
- Propriétés cachées du fichier

### B. Commentaires oubliés

- Commentaires encore présents
- Fils de commentaires
- Remarques internes
- Annotations non supprimées

### C. Track changes / révisions

- Suivi des modifications activé
- Insertions / suppressions visibles
- Texte supprimé encore accessible
- Versions intermédiaires visibles

### D. Données personnelles exposées

- Adresses email
- Numéros de téléphone
- Noms de personnes
- Coordonnées internes
- Signatures internes
- Informations identifiantes dans le texte

### E. Références internes

- Chemins de fichiers locaux
- Chemins réseau internes
- Noms de dossiers internes
- URLs internes
- Noms d'équipes ou de projets internes non souhaités

### F. Placeholders / brouillons

- TODO · TBD · FIXME · WIP
- "insert client name" · "à compléter" · "draft"
- Sections incomplètes
- Textes oubliés en surbrillance ou entre crochets

### G. Hygiène / cohérence du document

- Incohérences de formatage
- Doubles espaces
- Sauts de ligne étranges
- Titres incohérents
- Sections vides
- Headers / footers incohérents
- Anciennes références de société ou de client

### H. Pièces jointes / éléments embarqués

- Images ou objets embarqués non désirés
- Liens cassés
- Feuilles / slides cachées (si applicable)
- Notes de présentation (PPTX)

---

## 3. Qualion Proposal — Détections

### Objectif

Reprendre **tout Qualion For All**, puis ajouter une analyse business approfondie de la proposition.

### Couche 1 — Tout Qualion For All

Qualion Proposal inclut automatiquement toutes les détections de Qualion For All : métadonnées, commentaires, track changes, données personnelles, références internes, placeholders, hygiène documentaire.

### Couche 2 — Analyse business / proposal

#### A. Erreurs client

- Mauvais nom client
- Mauvais logo / mauvaise référence entreprise
- Confusion entre plusieurs clients
- Références copiées d'une ancienne proposition

#### B. Incohérences de scope

- Scope contradictoire d'une section à l'autre
- Livrables mal définis
- Périmètre flou
- Promesses qui ne matchent pas le contenu détaillé
- Sections manquantes par rapport au besoin exprimé

#### C. Risques delivery

- Engagements trop ambitieux
- Planning peu crédible
- Dépendances non mentionnées
- Moyens non cohérents avec le scope
- Responsabilités peu claires

#### D. Risques de wording

- Langage trop vague
- Formulations juridiquement ou commercialement risquées
- Promesses trop fortes
- Formulations qui affaiblissent la position de négociation
- Tournures peu professionnelles ou défensives

#### E. Contenu trop générique

- Texte passe-partout
- Manque de spécificité client
- Manque d'ancrage métier
- Contenu réutilisé sans adaptation
- Absence de différenciation claire

#### F. Faiblesse de la proposition de valeur

- Bénéfices peu explicites
- Valeur business peu claire
- Manque d'impact client
- Pas d'arguments différenciants
- Pas de lien clair entre besoin et solution

#### G. Références / preuves manquantes

- Pas de références projet
- Pas de cas concrets
- Peu de preuves de crédibilité
- Manque d'éléments rassurants

#### H. Détail technique insuffisant

- Approche trop superficielle
- Méthodologie peu détaillée
- Ressources / profils non décrits
- Hypothèses non explicitées
- Solution pas assez tangible

#### I. Signaux pricing / marge

- Indices de marge internes
- Commentaires liés au pricing
- Termes internes de négociation
- Tableaux ou notes qui ne devraient pas être visibles

#### J. Signaux de confidentialité / gouvernance

- Informations stratégiques exposées
- Langage interne non destiné au client
- Informations non validées
- Incohérences de conformité ou de validation

#### K. Alignement RFP (phase suivante)

- Sections requises absentes
- Réponses incomplètes à la demande
- Annexes manquantes
- Mismatch entre RFP et proposal

---

## 4. Taxonomie Qualion Proposal

Chaque problème détecté dans Qualion Proposal est classé dans une catégorie fixe.

| Catégorie | Description |
|---|---|
| `client_error` | Erreur sur le nom, le logo ou les références client |
| `scope_mismatch` | Incohérence ou flou sur le périmètre |
| `delivery_risk` | Risque sur les engagements ou le planning |
| `risky_commitment` | Formulation trop engageante ou juridiquement risquée |
| `generic_content` | Contenu générique sans ancrage client |
| `confidentiality_exposure` | Information interne ou confidentielle visible |
| `pricing_signal` | Indice de marge ou de pricing visible |
| `missing_reference` | Absence de preuves ou de références projet |
| `weak_value_proposition` | Proposition de valeur floue ou peu convaincante |
| `insufficient_detail` | Détail technique ou méthodologique insuffisant |

---

## 5. Format affiché — Qualion For All

**En-tête**

- Document Cleanliness Score : 0–100
- Statut : Safe to send / Needs review / Not safe to send

**Section 1 — Summary**

- Nombre total de problèmes détectés
- Niveau de risque global
- 3 actions prioritaires

**Section 2 — Issues detected**

Par catégorie : Metadata exposure · Comments · Track changes · Personal data · Internal references · Hygiene issues

Pour chaque item : titre · sévérité (Low / Medium / High) · extrait concerné · recommandation simple

**Section 3 — Suggested fixes**

- Supprimer les métadonnées
- Supprimer les commentaires
- Accepter / rejeter les track changes
- Remplacer les données personnelles
- Supprimer les placeholders

**Section 4 — Final recommendation**

- Ready to send
- ou Fix before sending

---

## 6. Format affiché — Qualion Proposal

**En-tête**

- Client-Ready Score : 0–100
- Risk Score : 0–100
- Statut : Client-ready / Needs revision / High risk before sending

**Section 1 — Executive Summary**

- Assessment global de la proposition
- Top 3 risques identifiés
- Top 3 actions prioritaires

**Section 2 — Document Hygiene**

Reprend tous les résultats de Qualion For All.

**Section 3 — Business Risk Analysis**

Liste des risques classés par catégorie (client_error, scope_mismatch, delivery_risk, risky_commitment, generic_content, missing_reference, insufficient_detail). Pour chaque risque : catégorie · sévérité · description · extrait · recommandation.

**Section 4 — Strength Signals**

Points forts de la proposition : technical clarity · scope definition · references · client alignment.

**Section 5 — Proposal Dimension Scores**

| Dimension | Score (0-100) |
|---|---|
| Client alignment | — |
| Technical depth | — |
| Clarity | — |
| Risk exposure | — |
| Evidence strength | — |

**Section 6 — Final Recommendation**

- Ready to send
- Revise before submission
- High-risk proposal — do not send

---

## 7. Format de téléchargement — Qualion For All

### A. Document nettoyé

**Nom** : `{filename}_cleaned.docx` ou `{filename}_cleaned.pdf`

Contenu : commentaires supprimés · track changes acceptés / retirés · métadonnées nettoyées · placeholders supprimés · données sensibles masquées ou surlignées pour correction manuelle.

### B. Rapport d'analyse

**Nom** : `{filename}_qualion_report.pdf`

Contenu : score global · statut safe / not safe · liste des problèmes détectés avec sévérité · recommandations · résumé des corrections automatiques · éléments restant à corriger manuellement.

---

## 8. Format de téléchargement — Qualion Proposal

### A. Document modifié

**Nom** : `{proposalname}_client_ready_version.docx`

Contenu : nettoyage hygiène documentaire complet · commentaires supprimés · track changes retirés · zones à risque marquées ou annotées.

### B. Rapport exécutif PDF

**Nom** : `{proposalname}_executive_risk_report.pdf`

| Page | Contenu |
|---|---|
| Page 1 | Client-Ready Score · Risk Score · Verdict global |
| Page 2 | Top business risks · Top actions prioritaires |
| Page 3 | Document hygiene findings |
| Page 4 | Business risk categories détaillées |
| Page 5 | Strengths · Dimension scores · Final recommendation |

### Trois options de téléchargement

| Option | Contenu |
|---|---|
| Clean Version | Document nettoyé, prêt à correction finale |
| Analysis Report | Rapport PDF lisible et partageable |
| Annotated Version (Proposal only) | Version avec highlights sur les zones à corriger |

---

## 9. Architecture backend

### État actuel (points critiques)

- `server.js` = 72 KB — toute la logique dans un seul fichier
- `documentAnalyzer.js` = 77 KB — idem
- `aiBusinessRisk.js` = corps vide, non implémenté
- Pas de base de données connectée (Supabase non branché)
- Pas d'authentification
- Pas de rate limiting
- Cache in-memory uniquement (perdu à chaque restart)

### Structure cible

```
/server.js                    <- bootstrap uniquement (20 lignes)
/app.js                       <- middleware + enregistrement des routes
/routes/
  analyze.js                  <- POST /analyze
  clean.js                    <- POST /clean
  rephrase.js                 <- POST /rephrase
  health.js                   <- GET /health
/modules/
  qualion-clean/
    parser.js                 <- extraction de texte par format
    detectors.js              <- toutes les détections déterministes
    businessRisk.js           <- moteur de flags business
    scoring.js                <- calcul du score
    report.js                 <- générateur de rapport
  qualion-proposal/
    parser.js                 <- extraction approfondie pour proposals
    aiOrchestrator.js         <- orchestration des appels IA
    riskCategories.js         <- taxonomie des risques Proposal
    dimensionScorer.js        <- scores par dimension
    validator.js              <- validation du JSON retourné par l'IA
/lib/
  ai.js                       <- client OpenAI brut
  cache.js                    <- adaptateur cache (Map -> Redis)
  storage.js                  <- adaptateur Supabase
  validators.js               <- validation des inputs fichier
  errors.js                   <- classes d'erreurs typées
/middleware/
  auth.js                     <- validation API key
  rateLimit.js                <- limites par IP
  upload.js                   <- multer + vérification MIME
/db/
  migrations/
  schema.sql
```

### Endpoints

| Méthode | Route | Description |
|---|---|---|
| GET | /health | Statut du service |
| POST | /analyze | Analyse rapide ou complète |
| POST | /analyze-stream | Analyse avec progression SSE |
| POST | /clean | Nettoyage du document |
| POST | /rephrase | Reformulation du document |

---

## 10. Modèle de données Supabase

### Table `analyses`

```sql
CREATE TABLE analyses (
  id               UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  created_at       TIMESTAMPTZ DEFAULT now(),
  user_id          UUID REFERENCES auth.users,
  organization_id  UUID,
  product          TEXT NOT NULL,
  file_name        TEXT,
  file_type        TEXT,
  file_size        INTEGER,
  file_hash        TEXT,
  mode             TEXT,
  duration_ms      INTEGER,
  risk_score       INTEGER,
  risk_level       TEXT,
  client_ready     BOOLEAN
);
```

### Table `detection_signals`

```sql
CREATE TABLE detection_signals (
  id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  analysis_id  UUID REFERENCES analyses(id),
  category     TEXT,
  severity     TEXT,
  rule_id      TEXT,
  evidence     JSONB,
  was_cleaned  BOOLEAN
);
```

### Table `business_risk_flags`

```sql
CREATE TABLE business_risk_flags (
  id           UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  analysis_id  UUID REFERENCES analyses(id),
  category     TEXT,
  level        TEXT,
  rule_id      TEXT,
  reason       TEXT,
  evidence     JSONB,
  source       JSONB
);
```

### Table `proposal_scores`

```sql
CREATE TABLE proposal_scores (
  id                  UUID PRIMARY KEY DEFAULT gen_random_uuid(),
  analysis_id         UUID REFERENCES analyses(id),
  dimension           TEXT,
  score               INTEGER,
  strengths           JSONB,
  weaknesses          JSONB,
  ai_recommendations  JSONB
);
```

---

## 11. Priorités d'implémentation

### Critique — Immédiat

| # | Action | Effort |
|---|---|---|
| 1 | Connecter Supabase — persister toutes les analyses | 1 jour |
| 2 | Middleware d'authentification (API key) | 0.5 jour |
| 3 | Implémenter aiBusinessRisk.js (corps vide actuellement) | 3 jours |
| 4 | Découper server.js en routes + controllers | 2 jours |
| 5 | Rate limiting + validation MIME réelle | 0.5 jour |

### Important — Phase 2

| # | Action | Effort |
|---|---|---|
| 6 | Validation JSON retourné par l'IA (schema + ajv) | 1 jour |
| 7 | Timeout OpenAI + fallback déterministe | 0.5 jour |
| 8 | Cache Redis (Upstash) pour remplacer Map in-memory | 1 jour |
| 9 | Tables detection_signals et proposal_scores | 1 jour |

### Roadmap — Phase 3

| # | Action | Effort |
|---|---|---|
| 10 | Queue asynchrone pour traitement lourd (BullMQ) | 3 jours |
| 11 | Séparation complète modules Qualion Clean / Proposal | 1 semaine |
| 12 | Profils de risque par organisation | 1 semaine |
| 13 | Moteur d'apprentissage basé sur les signaux accumulés | Continu |

---

*Document maintenu par l'équipe Mindorion — dernière mise à jour : Mars 2026*
