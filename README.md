# 🏢 SIRET Fournisseurs ABRAPA

> Outil d'enrichissement et d'audit des SIRET fournisseurs via l'API SIRENE officielle (INSEE / data.gouv.fr)

![Version](https://img.shields.io/badge/version-4.0-orange)
![Licence](https://img.shields.io/badge/licence-MIT-blue)
![API](https://img.shields.io/badge/API-SIRENE%20officielle-green)
![Sage](https://img.shields.io/badge/Sage-FRP%201000-purple)

---

## 📋 Présentation

Ce projet fournit deux outils complémentaires pour la gestion des SIRET fournisseurs dans **Sage FRP 1000** :

| Outil | Fichier | Usage |
|---|---|---|
| Interface web | `siret-search-v4.html` | Utilisation graphique dans le navigateur |
| Script Python | `completer_siret.py` | Automatisation en ligne de commande |

Les deux interrogent l'**API SIRENE officielle** en temps réel pour enrichir les fiches tiers avec les données officielles du registre des entreprises françaises.

---

## ✨ Fonctionnalités

### Interface HTML (navigateur)

| Onglet | Description |
|---|---|
| 🔍 **Recherche unitaire** | Trouver le SIRET d'un fournisseur par son nom |
| 📋 **Traitement par lot** | Compléter les SIRET manquants depuis un fichier CSV |
| ✅ **Vérifier un SIRET** | Valider un SIRET 14 chiffres contre SIRENE |
| 📊 **Audit base fournisseurs** | Comparer en masse la base Sage avec SIRENE |

**Fonctionnalités transverses :**
- Pause / reprise de la recherche
- Sauvegarde automatique de session (sessionStorage)
- Export CSV encodé ISO-8859-1 (compatible Excel FR)
- Export Sage FRP 1000 avec colonnes `CT_*` prêtes à importer
- Lecture des fichiers Excel `.xlsx` / `.xls` en plus du CSV
- Guide d'utilisation intégré

### Script Python

```
Mode complete  →  Recherche les SIRET manquants
Mode verify    →  Audite les SIRET existants contre SIRENE
```

Génère un **fichier Excel coloré** par statut + un **CSV d'actions** trié par priorité pour Sage FRP 1000.

---

## 📦 Champs enrichis depuis l'API SIRENE

| Champ API | Colonne Sage FRP 1000 | Disponibilité |
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

> **Note :** Les IBAN / RIB sont stockés dans la table `F_REGLEMENTT` de Sage, distincte de la fiche tiers. Ils ne sont jamais touchés par cet import.

---

## 🚀 Utilisation

### Interface HTML

Aucune installation requise. Ouvrir `siret-search-v4.html` directement dans un navigateur moderne (Chrome, Firefox, Edge).

> ⚠️ Une connexion internet est nécessaire pour interroger l'API SIRENE.

---

### Script Python

#### Prérequis

```bash
pip install pandas requests openpyxl
```

#### Mode `complete` — Compléter les SIRET manquants

```bash
python completer_siret.py base_fournisseurs.xlsx
```

```bash
# Avec colonnes explicites
python completer_siret.py base.xlsx --col-nom CT_Intitule --col-siret CT_Siret --col-cp CT_CodePostal
```

```bash
# Aperçu sans modification
python completer_siret.py base.xlsx --dry-run
```

#### Mode `verify` — Auditer les SIRET existants

```bash
python completer_siret.py export_sage.xlsx --mode verify
```

```bash
# Avec référence Sage (CT_Num) pour l'export
python completer_siret.py export_sage.xlsx --mode verify --col-siret CT_Siret --col-id CT_Num
```

```bash
# Aperçu sans modification
python completer_siret.py export_sage.xlsx --mode verify --dry-run
```

#### Paramètres disponibles

| Paramètre | Description | Défaut |
|---|---|---|
| `--mode` | `complete` ou `verify` | `complete` |
| `--col-nom` | Colonne nom fournisseur | Auto-détection |
| `--col-siret` | Colonne SIRET | Auto-détection |
| `--col-cp` | Colonne code postal | Auto-détection |
| `--col-id` | Colonne référence Sage (`CT_Num`) | Auto-détection |
| `--delay` | Délai entre requêtes (secondes) | `0.4` |
| `--dry-run` | Analyse sans modification | — |

---

## 📊 Statuts de l'audit (`--mode verify`)

| Statut | Signification | Action Sage FRP 1000 |
|---|---|---|
| ✅ OK | SIRET actif, raison sociale concordante | Aucune action |
| 🔴 Fermé | Établissement radié dans SIRENE | Bloquer le tiers (`CT_Sommeil = 1`) |
| 📝 Nom différent | Raison sociale modifiée (fusion, changement) | Mettre à jour `CT_Intitule` |
| 🔴 Fermé + Nom ≠ | Clos et raison sociale différente | Bloquer + rechercher successeur |
| ❓ Introuvable | SIRET absent de SIRENE | Vérification manuelle |
| ⚠️ Format invalide | SIRET ≠ 14 chiffres dans Sage | Corriger `CT_Siret` |
| — Sans SIRET | Champ vide dans Sage | SIRET à saisir |

---

## 📁 Fichiers générés

### Mode `complete`

| Fichier | Contenu |
|---|---|
| `*_siret_enrichi.xlsx` | Fichier source enrichi, cellules colorées par statut |
| `*_siret_introuvables.csv` | Liste des fournisseurs sans SIRET trouvé |

### Mode `verify`

| Fichier | Contenu |
|---|---|
| `*_audit_sirene.xlsx` | Rapport complet coloré par statut + onglet récapitulatif |
| `*_audit_actions_sage.csv` | Actions à réaliser, triées par priorité, format `CT_*` pour import Sage |

---

## 🔌 API utilisée

**[recherche-entreprises.api.gouv.fr](https://recherche-entreprises.api.gouv.fr)**

- Gratuite, sans clé API
- Données officielles INSEE / SIRENE
- Rate limit : ~150 requêtes/minute (délai de 380 ms entre chaque appel géré automatiquement)
- En cas de dépassement (HTTP 429) : pause automatique de 2 secondes

---

## 🏗️ Structure du projet

```
siret-fournisseurs-abrapa/
├── siret-search-v4.html       # Interface web complète (autonome)
├── completer_siret.py         # Script Python CLI
└── README.md
```

---

## ⚙️ Import dans Sage FRP 1000

1. **Fichier → Import → Tiers**
2. Sélectionner le CSV exporté (encodage **ISO-8859-1**, séparateur **`;`**)
3. Les en-têtes `CT_*` sont reconnus automatiquement
4. Sage identifie chaque tiers par `CT_Num` :
   - Code existant → **mise à jour** des champs renseignés
   - Code inexistant → **création** d'un nouveau tiers
   - Champ **vide** dans le CSV → ⚠️ champ existant **écrasé à blanc**

> 💡 **Recommandation** : pour un premier import, n'inclure que les colonnes `CT_Siret`, `CT_Intitule`, `CT_Adresse`, `CT_CodePostal`, `CT_Ville` afin de préserver les valeurs saisies manuellement (téléphone, email).

---

## 🔒 Sécurité des données

- Aucune donnée n'est transmise à un serveur tiers hormis l'API SIRENE officielle
- Le fichier HTML fonctionne entièrement côté navigateur
- Aucune clé API, aucun compte requis
- La session en cours est stockée uniquement dans le `sessionStorage` local du navigateur

---

## 📄 Licence

MIT — Libre d'utilisation, de modification et de distribution.

---

*Développé pour le département Comptabilité & Finance — ABRAPA*
