# 🏢 SIRET Fournisseurs ABRAPA

> Outil d'enrichissement et d'audit des SIRET fournisseurs via l'API SIRENE officielle (INSEE / data.gouv.fr)

![Version](https://img.shields.io/badge/version-4.2-orange)
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

> 💡 Le fichier HTML peut être renommé à volonté sans impact sur son fonctionnement.  
> 📂 Chemin réseau : `L:\COMPTE\FOURNISSEURS\HUB OUTILS COMPTA\RECHERCHE SIRET FOURNISSEURS.html`

---

## ✨ Fonctionnalités

### Interface HTML (navigateur)

| Onglet | Description | Nouveautés v4.2 |
|---|---|---|
| 🔍 **Recherche unitaire** | Trouver le SIRET d'un fournisseur par son nom | Historique 20 recherches, copie par champ, Tout copier, réinitialisation |
| 📋 **Traitement par lot** | Compléter les SIRET manquants depuis un fichier CSV/Excel | — |
| ✅ **Vérifier un SIRET / SIREN** | Valider un SIRET (14) ou lister les établissements via SIREN (9) | Recherche SIREN, grille CT_*, copie par champ, export Excel, réinitialisation |
| 📊 **Audit base fournisseurs** | Comparer en masse la base Sage avec SIRENE | — |

**Fonctionnalités transverses :**
- Guide d'utilisation intégré (bouton ❓ dans le header)
- Pause / reprise sur traitement par lot et audit
- Sauvegarde automatique de session (sessionStorage)
- Export CSV ISO-8859-1 et Excel `.xlsx` compatibles Sage FRP 1000
- Lecture Excel `.xlsx` / `.xls` + CSV avec détection automatique du séparateur

---

## 🔍 Recherche unitaire — Détail

La fiche de chaque résultat affiche tous les champs Sage en **grille 2 colonnes** :

| Identification | Adresse & contact |
|---|---|
| CT_Intitule, CT_Siret, CT_Siren | CT_Adresse, CT_CodePostal, CT_Ville |
| CT_NatureJuridique, CT_NumTVAIntracomm | CT_Telephone, CT_Email, CT_Site (si disponibles) |
| Code NAF + libellé | — |

**Actions disponibles sur chaque résultat :**
- 📋 Bouton sur chaque champ CT_* pour copier individuellement
- **Tout copier** — copie tous les champs en tableau `champ[TAB]valeur` (collable dans Excel)
- **📥 Export Excel Sage FRP 1000** — fichier `.xlsx` prêt à importer
- **✖ Nouvelle recherche** — réinitialise le formulaire et les résultats

**Historique des recherches :**  
Les 20 derniers termes sont affichés en badges. Clic → relance la recherche. Persistant durant la session navigateur.

---

## ✅ Vérifier un SIRET / SIREN — Détail

| Saisie | Comportement | Cas d'usage |
|---|---|---|
| 14 chiffres (SIRET) | Vérifie l'établissement exact | Contrôle avant saisie dans Sage |
| 9 chiffres (SIREN) | Liste tous les établissements de l'entreprise | Fournisseur multi-sites, trouver le bon établissement |

Même fiche détaillée et mêmes actions que la Recherche unitaire (copie champ par champ, Tout copier, Export Excel, réinitialisation).

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

## 🚀 Utilisation

### Interface HTML

Ouvrir le fichier directement dans Chrome, Firefox ou Edge. Connexion internet requise.

#### 🔍 Recherche unitaire

1. Saisir le nom (partiel suffit) + code postal optionnel
2. Entrée ou clic **Rechercher**
3. Copier un champ 📋 ou cliquer **Tout copier** pour coller dans Sage
4. **📥 Export Excel** pour importer plusieurs résultats dans Sage
5. **✖ Nouvelle recherche** pour remettre à zéro

#### ✅ Vérifier un SIRET / SIREN

1. Saisir 14 chiffres (SIRET) ou 9 chiffres (SIREN)
2. Cliquer **Vérifier**
3. Avec un SIREN : tous les établissements de l'entreprise s'affichent
4. Mêmes boutons copie/export que la Recherche unitaire

#### 📋 Traitement par lot

1. Exporter la liste tiers depuis Sage (Fichier → Export)
2. Glisser-déposer le fichier CSV ou Excel
3. Associer les colonnes (détection automatique CT_*)
4. Lancer — pause/reprise possible
5. Exporter CSV complet ou **Export Sage FRP 1000**

#### 📊 Audit base fournisseurs

1. Exporter depuis Sage : CT_Num, CT_Intitule, CT_Siret (+ CT_CodePostal recommandé)
2. Importer, mapper, lancer
3. Filtrer par statut, consulter les KPI
4. Exporter rapport / mises à jour / actions par priorité

---

### Script Python

#### Prérequis

```bash
pip install pandas requests openpyxl
```

#### Mode `complete` — Compléter les SIRET manquants

```bash
python completer_siret.py base_fournisseurs.xlsx
python completer_siret.py base.xlsx --col-nom CT_Intitule --col-siret CT_Siret --col-cp CT_CodePostal
python completer_siret.py base.xlsx --dry-run
```

#### Mode `verify` — Auditer les SIRET existants

```bash
python completer_siret.py export_sage.xlsx --mode verify
python completer_siret.py export_sage.xlsx --mode verify --col-siret CT_Siret --col-id CT_Num
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

| Statut | Signification | Priorité | Action Sage FRP 1000 |
|---|---|---|---|
| ✅ OK | SIRET actif, nom concordant | — | Aucune action |
| 🔴 Fermé | Établissement radié | URGENT | Bloquer le tiers (`CT_Sommeil = 1`) |
| 🔴 Fermé + Nom ≠ | Clos et raison sociale différente | URGENT | Bloquer + rechercher successeur |
| ⚠️ Format invalide | SIRET ≠ 14 chiffres dans Sage | CRITIQUE | Corriger `CT_Siret` |
| 📝 Nom différent | Raison sociale modifiée | NORMAL | Mettre à jour `CT_Intitule` |
| ❓ Introuvable | SIRET absent de SIRENE | À VÉRIFIER | Vérification manuelle |
| — Sans SIRET | Champ `CT_Siret` vide | À VÉRIFIER | SIRET à saisir |

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

## 🔌 API utilisée

**[recherche-entreprises.api.gouv.fr](https://recherche-entreprises.api.gouv.fr)** — Gratuite, sans clé API, données INSEE en temps réel.  
Rate limit ~150 req/min — délai 380 ms géré automatiquement. Pause 2 s sur HTTP 429.

---

## 🔒 Sécurité

- Aucune donnée transmise hormis l'API SIRENE officielle
- Fonctionne entièrement côté navigateur, sans serveur
- Session stockée uniquement dans le `sessionStorage` local
- Fichier HTML renommable librement

---

## 🏗️ Structure du projet

```
siret-fournisseurs-abrapa/
├── RECHERCHE SIRET FOURNISSEURS.html   # Interface web (renommable)
├── completer_siret.py                  # Script Python CLI
└── README.md
```

---

## 📄 Licence

MIT — Libre d'utilisation, de modification et de distribution.

---

*Développé pour le département Comptabilité & Finance — ABRAPA — v4.2*
