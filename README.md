# 🏢 SIRET Fournisseurs ABRAPA

> Outil de recherche, vérification et audit SIRET/SIREN pour la comptabilité ABRAPA.  
> Basé sur l'API SIRENE officielle INSEE et l'API recherche-entreprises (data.gouv.fr).

![Version](https://img.shields.io/badge/version-5.1-orange)
![Licence](https://img.shields.io/badge/licence-MIT-blue)
![API](https://img.shields.io/badge/API-SIRENE%20officielle-green)
![Sage](https://img.shields.io/badge/Sage-FRP%201000-purple)
![Standalone](https://img.shields.io/badge/mode-autonome-lightgrey)

---

## 📁 Fichiers

```
📁 dossier-installation/
├── RECHERCHE_SIRET_FOURNISSEURS.html     ← Application principale (version autonome)
├── GUIDE_EMPLOI_RECHERCHE_SIRET.html     ← Guide d'emploi interactif
├── xlsx.full.min.js                      ← Librairie exports Excel (fallback local)
└── README.md                             ← Ce fichier
```

> ⚠️ Les quatre fichiers doivent être placés dans le **même dossier**.

---

## 🚀 Utilisation

Ouvrir `RECHERCHE_SIRET_FOURNISSEURS.html` directement dans le navigateur (Chrome, Edge, Firefox).  
**Aucune installation, aucun serveur, aucun Python requis.**

---

## ✨ Nouveautés v5.1

| # | Fonctionnalité | Détail |
|---|----------------|--------|
| 🆕 | **Comparaison champ par champ (Audit)** | Bouton 📊 Comparer sur chaque ligne : modale Sage ↔ SIRENE avec badge Diff/OK par champ, Accepter individuel et Tout accepter |
| 🆕 | **Résumé visuel batch** | Donut SVG + 4 KPI cards automatiques après chaque traitement par lot (taux d'enrichissement, durée, ventilation) |
| 🆕 | **Import Excel dans le Batch** | Le traitement par lot accepte désormais `.xlsx`, `.xls`, `.csv`, `.txt` — détection automatique du format |
| 🆕 | **Indicateur statut API** | Voyant 🟢/🟡/🔴 dans le header — test live des deux APIs SIRENE toutes les 5 min, tooltip avec latence et bouton ↻ Tester |

---

## ✨ Nouveautés v5.0

| # | Fonctionnalité | Détail |
|---|----------------|--------|
| 🆕 | **Sélection multiple + export groupé** | Cochez plusieurs cartes résultat et exportez uniquement votre sélection |
| 🆕 | **Historique persistant** | 50 dernières recherches conservées entre les sessions (localStorage) avec horodatage relatif |
| 🆕 | **Badge « Nom ✕ » cliquable** | Réinitialise la recherche en un clic, accessible sur mobile |
| 🆕 | **Export établissements filtré** | L'export de la modale établissements respecte le filtre actif (Actifs / Fermés / Tous) |

---

## ✨ Fonctionnalités

### Onglet 1 — Recherche & Vérification

Champ unique intelligent qui détecte automatiquement le type de saisie :

| Saisie | Mode | Action |
|--------|------|--------|
| Texte libre | 🔤 Nom | Recherche par nom + code postal optionnel |
| 9 chiffres | 🏢 SIREN | Fiche entreprise complète + bouton liste établissements |
| 14 chiffres | 🔢 SIRET | Vérification directe de l'établissement |

**Bouton "Tous les établissements du SIREN"** → Modale avec liste complète via API INSEE, filtres actifs/fermés, recherche, export Excel **respectant le filtre actif**.

**Actions disponibles sur chaque résultat :**
- 📋 Bouton sur chaque champ `CT_*` pour copier individuellement
- **Tout copier** — copie tous les champs en tableau `champ[TAB]valeur` (collable dans Excel)
- **☑️ Sélection multiple** — cliquer sur une carte (ou sa case) la sélectionne
- **📥 Exporter la sélection** — fichier Excel avec uniquement les cartes cochées
- **📥 Exporter tout** — fichier `.xlsx` complet prêt à importer dans Sage
- **Badge Nom ✕** — réinitialise le formulaire et les résultats

**Sélection multiple :**  
Un bandeau sticky apparaît en bas dès qu'une carte est sélectionnée avec : compteur, *Tout sélectionner / Tout désélectionner*, *📥 Exporter la sélection* et *✕ Effacer la sélection*.

**Historique des recherches :**  
Les **50 derniers** termes sont affichés en badges horodatés. Clic → relance la recherche. Persistant **entre les sessions** grâce au localStorage.

---

### Onglet 2 — Traitement par lot

1. Exporter la liste tiers depuis Sage (Fichier → Export)
2. Glisser-déposer le fichier **CSV, TXT ou Excel (.xlsx/.xls)** — détection automatique du format
3. Associer les colonnes (détection automatique `CT_*`)
4. Lancer — pause/reprise possible
5. Consulter le **résumé visuel** (donut SVG + KPI cards) puis exporter

**Indicateur statut API :**  
Un voyant 🟢/🟡/🔴 dans le header vérifie automatiquement les deux APIs SIRENE toutes les 5 minutes. Clic → tooltip détaillé (latence par API, heure de vérification, bouton ↻ Tester).

**Résumé visuel automatique :**  
Après chaque traitement, un panneau s'affiche avec un graphique donut SVG (trouvés / conservés / introuvables), le taux d'enrichissement et la durée totale.

---

### Onglet 3 — Audit base fournisseurs

1. Exporter depuis Sage : `CT_Num`, `CT_Intitule`, `CT_Siret` (+ `CT_CodePostal` recommandé)
2. Vérification dual mode : par **SIRET** (confirmation directe) ou par **nom** (fallback si SIRET absent)
3. Filtrer par statut, consulter les KPI
4. **Comparer champ par champ** via le bouton 📊 Comparer sur chaque ligne
5. Exporter rapport / mises à jour / actions par priorité

**Statuts de l'audit :**

| Statut | Signification | Priorité | Action Sage FRP 1000 |
|--------|---------------|----------|----------------------|
| ✅ OK | SIRET actif, nom concordant | — | Aucune action |
| 🔴 Fermé | Établissement radié | URGENT | Bloquer le tiers (`CT_Sommeil = 1`) |
| 🔴 Fermé + Nom ≠ | Clos et raison sociale différente | URGENT | Bloquer + rechercher successeur |
| ⚠️ Format invalide | SIRET ≠ 14 chiffres dans Sage | CRITIQUE | Corriger `CT_Siret` |
| 📝 Nom différent | Raison sociale modifiée | NORMAL | Mettre à jour `CT_Intitule` |
| 🔍 Trouvé par nom | Identifié par recherche nom (pas SIRET) | À VÉRIFIER | Confirmer et saisir le SIRET |
| ❓ Introuvable | SIRET absent de SIRENE | À VÉRIFIER | Vérification manuelle |
| — Sans SIRET | Champ `CT_Siret` vide | À VÉRIFIER | SIRET à saisir |

**Comparaison champ par champ :**  
Sur chaque ligne avec données SIRENE disponibles (statuts OK, Nom ≠, Fermé, Trouvé/nom), le bouton **📊 Comparer** ouvre une modale affichant côte à côte les valeurs Sage et SIRENE pour 10 champs (`CT_Intitule`, `CT_Siret`, `CT_Adresse`, `CT_CodePostal`, `CT_Ville`, `CT_NatureJuridique`, `CT_NumTVAIntracomm`, `CT_Telephone`, `CT_Email`, Code NAF). Chaque champ peut être accepté individuellement ou via **Tout accepter**.

---

## 📖 Guide d'emploi

Un guide interactif `GUIDE_EMPLOI_RECHERCHE_SIRET.html` est fourni avec l'application.

- Accessible depuis l'application via le bouton **📚 Guide complet** dans l'en-tête
- Couvre les 4 onglets, la sélection multiple, l'historique, la comparaison champ par champ, les exports et les conseils pratiques
- Mis à jour pour la **v5.1**

---

## 📦 Champs enrichis depuis l'API SIRENE

| Champ API SIRENE | Colonne Sage FRP 1000 | Disponibilité |
|---|---|---|
| `siret` | `CT_Siret` | ✅ Toujours |
| `siren` | `CT_Siren` | ✅ Toujours |
| `nom_complet` | `CT_Intitule` | ✅ Toujours |
| `adresse` | `CT_Adresse` | ✅ Très souvent |
| `code_postal` | `CT_CodePostal` | ✅ Très souvent |
| `libelle_commune` | `CT_Ville` | ✅ Très souvent |
| `nature_juridique_label` | `CT_NatureJuridique` | ✅ Très souvent |
| `numero_tva_intra` | `CT_NumTVAIntracomm` | ⚡ Souvent |
| `telephone` | `CT_Telephone` | ⚠️ Minorité |
| `email` | `CT_Email` | ⚠️ Minorité |
| `site_internet` | `CT_Site` | ⚠️ Minorité |

> **Note IBAN :** Stockés dans `F_REGLEMENTT`, jamais affectés par cet import.

---

## ⚙️ Import dans Sage FRP 1000

1. **Fichier → Import → Tiers**
2. Sélectionner le CSV (ISO-8859-1, séparateur `;`) ou Excel
3. En-têtes `CT_*` reconnus automatiquement
4. Tester sur 5–10 tiers avant l'import complet

| Situation | Comportement Sage |
|---|---|
| CT_Num existe + champ renseigné | ✅ Mise à jour |
| CT_Num existe + champ **vide** | ⚠️ Écrasé à blanc |
| CT_Num n'existe pas | ✅ Création |

> Tables **non affectées** : `F_REGLEMENTT` (IBAN/RIB), `F_CONTACTT` (contacts), `F_ECRITUREC` (écritures).

---

## 🌐 APIs utilisées

| API | Usage | Auth |
|-----|-------|------|
| `recherche-entreprises.api.gouv.fr` | Recherche par nom, traitement par lot, audit | Aucune (publique) — ~150 req/min |
| `api.insee.fr/api-sirene/3.11` | Liste complète des établissements d'un SIREN | Clé API intégrée |

---

## 🔑 Clé API INSEE

- Clé intégrée dans le code (header `X-INSEE-Api-Key-Integration`)
- Valable **indéfiniment** — 30 requêtes/minute maximum
- En cas d'expiration : obtenir une nouvelle clé sur **[portail-api.insee.fr](https://portail-api.insee.fr/)** et mettre à jour la constante `INSEE_KEY` dans le fichier HTML (rechercher `ecad76f8`)

---

## 📦 Librairie xlsx.js

L'outil charge `xlsx.full.min.js` selon la priorité suivante :
1. **CDN Cloudflare** (si internet accessible) — chargement automatique
2. **Fichier local** `xlsx.full.min.js` — fallback automatique si CDN bloqué (réseau proxy)

---

## 🔒 Sécurité & Stockage local

| Donnée | Stockage | Durée |
|--------|----------|-------|
| Historique des recherches | `localStorage` | Permanent (jusqu'à effacement manuel) |
| Session traitement par lot | `sessionStorage` | Durée de la session navigateur |
| Session audit | `sessionStorage` | Durée de la session navigateur |

Aucune donnée n'est transmise hormis les appels aux APIs SIRENE officielles.

---

## 🔄 Mise à jour

Pour remplacer la version installée :

1. Télécharger les nouveaux fichiers depuis ce dépôt
2. Remplacer `RECHERCHE_SIRET_FOURNISSEURS.html` et `GUIDE_EMPLOI_RECHERCHE_SIRET.html`
3. **Ne pas remplacer** `xlsx.full.min.js` sauf si une nouvelle version est explicitement fournie

---

## 📋 Compatibilité

- Chrome / Edge / Firefox — versions récentes
- Fonctionne sur réseau avec proxy corporate (pas de dépendance réseau hors APIs SIRENE)
- Aucune dépendance serveur — fichier HTML autonome

---

## 📝 Historique des versions

| Version | Date | Nouveautés |
|---------|------|------------|
| v5.1 | 03/2026 | Comparaison champ par champ dans l'Audit · Résumé visuel batch (donut SVG + KPI) · Import Excel (.xlsx/.xls) dans le Batch · Indicateur statut API (voyant 🟢/🟡/🔴, tooltip latence) |
| v5.0 | 03/2026 | Sélection multiple + export groupé · Historique persistant localStorage · Badge Nom ✕ cliquable · Export établissements respectant le filtre |
| v4.2 | 03/2026 | Fusion onglets Recherche + Vérification · Liste établissements API INSEE · Audit dual mode SIRET + nom · CDN + fallback local xlsx · Guide d'emploi interactif |
| v4.0 | 03/2026 | Audit base fournisseurs · Export 3 formats · Traitement par lot amélioré |
| v3.0 | 02/2026 | Recherche unitaire · Traitement par lot · Vérification SIRET/SIREN |

---

## 🏗️ Structure du projet

```
Recherche-SIRET/
├── RECHERCHE_SIRET_FOURNISSEURS.html    # Application principale (version autonome)
├── GUIDE_EMPLOI_RECHERCHE_SIRET.html   # Guide d'emploi interactif
├── xlsx.full.min.js                     # Librairie Excel (fallback local)
└── README.md
```

---

## 📄 Licence

MIT — Libre d'utilisation, de modification et de distribution.

---

*Développé pour le département Comptabilité & Finance — ABRAPA — v5.1*
