# 🏢 SIRET Fournisseurs ABRAPA

> Outil d'enrichissement et d'audit des SIRET fournisseurs via l'API SIRENE officielle (INSEE / data.gouv.fr)

![Version](https://img.shields.io/badge/version-4.1-orange)
![Licence](https://img.shields.io/badge/licence-MIT-blue)
![API](https://img.shields.io/badge/API-SIRENE%20officielle-green)
![Sage](https://img.shields.io/badge/Sage-FRP%201000-purple)

---

## 📋 Présentation

Ce projet fournit deux outils complémentaires pour la gestion des SIRET fournisseurs dans **Sage FRP 1000** :

| Outil | Fichier | Usage |
|---|---|---|
| Interface web | `*.html` (renommable librement) | Utilisation graphique dans le navigateur |
| Script Python | `completer_siret.py` | Automatisation en ligne de commande |

Les deux interrogent l'**API SIRENE officielle** en temps réel pour enrichir les fiches tiers avec les données officielles du registre des entreprises françaises.

> 💡 Le fichier HTML peut être renommé à volonté sans impact sur son fonctionnement — il ne contient aucune référence à son propre nom.

---

## ✨ Fonctionnalités

### Interface HTML (navigateur)

| Onglet | Description | Export disponible |
|---|---|---|
| 🔍 **Recherche unitaire** | Trouver le SIRET d'un fournisseur par son nom | 📥 Excel Sage FRP 1000 |
| 📋 **Traitement par lot** | Compléter les SIRET manquants depuis un fichier CSV/Excel | 📥 CSV complet + CSV Sage |
| ✅ **Vérifier un SIRET** | Valider un SIRET 14 chiffres contre SIRENE | — |
| 📊 **Audit base fournisseurs** | Comparer en masse la base Sage avec SIRENE | 📥 3 formats d'export |

**Fonctionnalités transverses :**
- Guide d'utilisation intégré (bouton ❓ dans le header)
- Pause / reprise de la recherche par lot et de l'audit
- Sauvegarde automatique de session (sessionStorage) avec restauration au rechargement
- Export CSV encodé ISO-8859-1 (compatible Excel FR, séparateur `;`)
- Export Excel `.xlsx` avec colonnes `CT_*` prêtes à importer dans Sage
- Lecture des fichiers Excel `.xlsx` / `.xls` en plus du CSV
- Détection automatique du séparateur CSV (`;` ou `,`)

### Script Python

```
Mode complete  →  Recherche les SIRET manquants
Mode verify    →  Audite les SIRET existants contre SIRENE
```

Génère un **fichier Excel coloré** par statut + un **CSV d'actions** trié par priorité pour Sage FRP 1000.

---

## 📦 Champs enrichis depuis l'API SIRENE

Les deux outils (HTML et Python) remontent les mêmes champs, mappés directement sur les colonnes Sage FRP 1000 :

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

> **Note IBAN :** Les coordonnées bancaires sont stockées dans la table `F_REGLEMENTT` de Sage, distincte de la fiche tiers `F_COMPTET`. Elles ne sont **jamais touchées** par cet import.

---

## 🚀 Utilisation

### Interface HTML

Aucune installation requise. Ouvrir le fichier HTML directement dans un navigateur moderne (Chrome, Firefox, Edge).

> ⚠️ Une connexion internet est nécessaire pour interroger l'API SIRENE.

#### 🔍 Recherche unitaire

1. Saisir une partie du nom du fournisseur (ex : `PHARMAT`, `Boehringer`)
2. Ajouter le code postal pour affiner en cas d'enseigne nationale
3. Appuyer sur **Entrée** ou cliquer **Rechercher**
4. Utiliser le bouton **Copier** sur le SIRET souhaité
5. Cliquer **📥 Export Excel Sage FRP 1000** pour télécharger un fichier `.xlsx` avec tous les résultats au format `CT_*`

#### 📋 Traitement par lot

1. Exporter la liste tiers depuis Sage FRP 1000 (menu *Fichier → Export*)
2. Glisser-déposer le fichier CSV ou Excel dans la zone de dépôt
3. Associer les colonnes (détection automatique des colonnes Sage `CT_*`)
4. Cliquer **Lancer** — pause/reprise possible à tout moment
5. Exporter : CSV complet ou **Export Sage FRP 1000** (colonnes `CT_*`)

#### 📊 Audit base fournisseurs

1. Exporter depuis Sage : colonnes `CT_Num`, `CT_Intitule`, `CT_Siret` (+ `CT_CodePostal` recommandé)
2. Importer le fichier (Excel ou CSV)
3. Associer les colonnes
4. Lancer la vérification
5. Consulter le tableau de bord KPI et filtrer par statut
6. Exporter : rapport complet / mises à jour Sage / liste d'actions par priorité

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

## 📊 Statuts de l'audit

Utilisés dans l'onglet **Audit base fournisseurs** (HTML) et le mode `--mode verify` (Python) :

| Statut | Signification | Priorité | Action Sage FRP 1000 |
|---|---|---|---|
| ✅ OK | SIRET actif, raison sociale concordante | — | Aucune action |
| 🔴 Fermé | Établissement radié dans SIRENE | 🔺 URGENT | Bloquer le tiers (`CT_Sommeil = 1`) |
| 🔴 Fermé + Nom ≠ | Clos et raison sociale différente | 🔺 URGENT | Bloquer + rechercher successeur |
| ⚠️ Format invalide | SIRET ≠ 14 chiffres dans Sage | 🔶 CRITIQUE | Corriger `CT_Siret` |
| 📝 Nom différent | Raison sociale modifiée (fusion, changement) | 🔷 NORMAL | Mettre à jour `CT_Intitule` |
| ❓ Introuvable | SIRET absent de SIRENE | 🔹 À VÉRIFIER | Vérification manuelle |
| — Sans SIRET | Champ `CT_Siret` vide dans Sage | 🔹 À VÉRIFIER | SIRET à saisir |

---

## 📁 Fichiers générés

### Interface HTML — Recherche unitaire

| Fichier | Contenu |
|---|---|
| `sage_frp1000_<nom>_<date>.xlsx` | Résultats au format `CT_*`, 2 onglets : *Tiers SIRENE* + *Notice import Sage* |

### Interface HTML — Traitement par lot

| Fichier | Contenu |
|---|---|
| `fournisseurs_siret_<date>.csv` | Rapport complet tous fournisseurs |
| `sage_frp1000_enrichi_<date>.csv` | Colonnes `CT_*` pour import Sage (SIRET trouvés + conservés) |
| `sage_frp1000_a_traiter_manuellement_<date>.csv` | Fournisseurs introuvables |

### Interface HTML — Audit base fournisseurs

| Fichier | Contenu |
|---|---|
| Rapport complet CSV | Tous les fournisseurs avec statut et données SIRENE |
| Mises à jour Sage CSV | Colonnes `CT_*` — uniquement les lignes avec action requise |
| Actions CSV | Trié par priorité (1-URGENT → 4-À_VÉRIFIER) |

### Script Python — Mode `complete`

| Fichier | Contenu |
|---|---|
| `*_siret_enrichi.xlsx` | Fichier source enrichi, cellules colorées par statut |
| `*_siret_introuvables.csv` | Liste des fournisseurs sans SIRET trouvé |

### Script Python — Mode `verify`

| Fichier | Contenu |
|---|---|
| `*_audit_sirene.xlsx` | Rapport complet coloré par statut |
| `*_audit_actions_sage.csv` | Actions triées par priorité, colonnes `CT_*` pour import Sage |

---

## ⚙️ Import dans Sage FRP 1000

### Procédure

1. **Fichier → Import → Tiers**
2. Sélectionner le fichier CSV (encodage **ISO-8859-1**, séparateur **`;`**) ou Excel
3. Les en-têtes `CT_*` sont reconnus automatiquement
4. Vérifier la correspondance des colonnes
5. **Tester sur un échantillon de 5–10 tiers** avant l'import complet

### Comportement de Sage selon la situation

| Situation | Comportement |
|---|---|
| `CT_Num` existe + champ renseigné | ✅ Mise à jour du champ |
| `CT_Num` existe + champ **vide** | ⚠️ Champ existant **écrasé à blanc** |
| `CT_Num` n'existe pas | ✅ Création d'un nouveau tiers |

### Tables Sage non affectées par l'import

| Donnée | Table Sage | Impact |
|---|---|---|
| IBAN / RIB / BIC | `F_REGLEMENTT` | ✅ Aucun |
| Contacts associés | `F_CONTACTT` | ✅ Aucun |
| Conditions de règlement | `F_REGLEMENTT` | ✅ Aucun |
| Historique écritures | `F_ECRITUREC` | ✅ Aucun |

> 💡 **Recommandation premier import** : n'inclure que `CT_Siret`, `CT_Intitule`, `CT_Adresse`, `CT_CodePostal`, `CT_Ville` pour préserver les valeurs saisies manuellement (téléphone, email, site).

---

## 🔌 API utilisée

**[recherche-entreprises.api.gouv.fr](https://recherche-entreprises.api.gouv.fr)**

- Gratuite, sans clé API, sans compte requis
- Données officielles INSEE / SIRENE, mises à jour en temps réel
- Rate limit : ~150 requêtes/minute — délai de 380 ms entre chaque appel géré automatiquement
- En cas de dépassement (HTTP 429) : pause automatique de 2 secondes

---

## 🔒 Sécurité des données

- Aucune donnée n'est transmise à un serveur tiers hormis l'API SIRENE officielle
- Le fichier HTML fonctionne entièrement côté navigateur (aucun serveur requis)
- La session en cours est stockée uniquement dans le `sessionStorage` local du navigateur
- Le fichier HTML peut être renommé librement sans impact fonctionnel

---

## 🏗️ Structure du projet

```
siret-fournisseurs-abrapa/
├── siret-search-v4.html       # Interface web complète (autonome, renommable)
├── completer_siret.py         # Script Python CLI (modes complete + verify)
└── README.md
```

---

## 📄 Licence

MIT — Libre d'utilisation, de modification et de distribution.

---

*Développé pour le département Comptabilité & Finance — ABRAPA*
