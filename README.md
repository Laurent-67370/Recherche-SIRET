# 🏢 SIRET Fournisseurs ABRAPA

> Outil de recherche, vérification et audit SIRET/SIREN pour la comptabilité ABRAPA.  
> Basé sur l'API SIRENE officielle INSEE et l'API recherche-entreprises (data.gouv.fr).

![Version](https://img.shields.io/badge/version-5.0-orange)
![Licence](https://img.shields.io/badge/licence-MIT-blue)
![API](https://img.shields.io/badge/API-SIRENE%20officielle-green)
![Sage](https://img.shields.io/badge/Sage-FRP%201000-purple)

---

## 📁 Fichiers

```
📁 dossier-installation/
├── RECHERCHE_SIRET_FOURNISSEURS.html     ← Application principale
├── GUIDE_EMPLOI_RECHERCHE_SIRET.html     ← Guide d'emploi interactif
├── xlsx.full.min.js                      ← Librairie exports Excel (fallback local)
└── README.md                             ← Ce fichier
```

> ⚠️ Les quatre fichiers doivent être placés dans le **même dossier**.  
> 📂 Chemin réseau : `L:\COMPTE\FOURNISSEURS\HUB OUTILS COMPTA\`

---

## 🚀 Utilisation

Ouvrir `RECHERCHE_SIRET_FOURNISSEURS.html` directement dans le navigateur (Chrome, Edge, Firefox).  
**Aucune installation, aucun serveur, aucun Python requis.**

---

## ✨ Fonctionnalités

### Onglet 1 — Recherche & Vérification

Champ unique intelligent qui détecte automatiquement le type de saisie :

| Saisie | Mode | Action |
|--------|------|--------|
| Texte libre | 🔤 Nom | Recherche par nom + code postal optionnel |
| 9 chiffres | 🏢 SIREN | Fiche entreprise complète + bouton liste établissements |
| 14 chiffres | 🔢 SIRET | Vérification directe de l'établissement |

**Bouton "Tous les établissements du SIREN"** → Modale avec liste complète via API INSEE, filtres actifs/fermés, recherche, export Excel.

**Actions disponibles sur chaque résultat :**
- 📋 Bouton sur chaque champ `CT_*` pour copier individuellement
- **Tout copier** — copie tous les champs en tableau `champ[TAB]valeur` (collable dans Excel)
- **📥 Export Excel Sage FRP 1000** — fichier `.xlsx` prêt à importer
- **✖ Nouvelle recherche** — réinitialise le formulaire et les résultats

**Historique des recherches :**  
Les 20 derniers termes sont affichés en badges. Clic → relance la recherche. Persistant durant la session navigateur.

---

### Onglet 2 — Traitement par lot

1. Exporter la liste tiers depuis Sage (Fichier → Export)
2. Glisser-déposer le fichier CSV ou Excel
3. Associer les colonnes (détection automatique `CT_*`)
4. Lancer — pause/reprise possible
5. Exporter CSV compatible Sage + export Excel coloré

---

### Onglet 3 — Audit base fournisseurs

1. Exporter depuis Sage : `CT_Num`, `CT_Intitule`, `CT_Siret` (+ `CT_CodePostal` recommandé)
2. Vérification dual mode : par **SIRET** (confirmation directe) ou par **nom** (fallback si SIRET absent)
3. Filtrer par statut, consulter les KPI
4. Exporter rapport / mises à jour / actions par priorité

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

---

## 📖 Guide d'emploi

Un guide interactif `GUIDE_EMPLOI_RECHERCHE_SIRET.html` est fourni avec l'application.

- Accessible depuis l'application via le bouton **📚 Guide complet** dans l'en-tête
- Bouton **← Retour à l'application** pour naviguer facilement entre les deux
- Couvre les 3 onglets, la liste établissements, les exports et les conseils pratiques

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
2. **Fichier local** `xlsx.full.min.js` — fallback automatique si CDN bloqué (réseau proxy ABRAPA)

---

## 🔒 Sécurité

- Aucune donnée transmise hormis les APIs SIRENE officielles
- Fonctionne entièrement côté navigateur, sans serveur
- Session stockée uniquement dans le `sessionStorage` local

---

## 🔄 Mise à jour

Pour remplacer la version installée :

1. Télécharger les nouveaux fichiers depuis ce dépôt
2. Remplacer `RECHERCHE_SIRET_FOURNISSEURS.html` et `GUIDE_EMPLOI_RECHERCHE_SIRET.html`
3. **Ne pas remplacer** `xlsx.full.min.js` sauf si une nouvelle version est explicitement fournie

---

## 📋 Compatibilité

- Chrome / Edge / Firefox — versions récentes
- Réseau ABRAPA (proxy corporate) — testé et fonctionnel
- Aucune dépendance serveur — fichiers HTML autonomes

---

## 📝 Historique des versions

| Version | Date | Nouveautés |
|---------|------|------------|
| v5.0 | 06/03/2026 | Fusion onglets Recherche + Vérification · Liste établissements API INSEE · Audit dual mode SIRET + nom · CDN + fallback local xlsx · Guide d'emploi interactif connecté |
| v4.2 | 03/2026 | Historique 20 recherches · Copie par champ · Recherche SIREN · Grille CT_* · Export Excel · Réinitialisation |
| v4.0 | 03/2026 | Audit base fournisseurs · Export 3 formats · Traitement par lot amélioré |
| v3.0 | 02/2026 | Recherche unitaire · Traitement par lot · Vérification SIRET/SIREN |

---

## 🏗️ Structure du projet

```
Recherche-SIRET/
├── RECHERCHE_SIRET_FOURNISSEURS.html    # Application principale
├── GUIDE_EMPLOI_RECHERCHE_SIRET.html   # Guide d'emploi interactif
├── xlsx.full.min.js                     # Librairie Excel (fallback local)
└── README.md
```

---

## 📄 Licence

MIT — Libre d'utilisation, de modification et de distribution.

---

*Développé pour le département Comptabilité & Finance — ABRAPA — v5.0*
