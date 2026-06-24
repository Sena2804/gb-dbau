# CNBAU — Application de Gestion des Bourses du Maroc

> Application web développée pour la **Commission Nationale des Bourses et Aides Universitaires (CNBAU)** du Bénin, dans le cadre de la session d'attribution des bourses de coopération marocaine.

---

## Table des matières

1. [Contexte et objectif](#1-contexte-et-objectif)
2. [Méthodologie technique](#2-méthodologie-technique)
3. [Résultats et livrables](#3-résultats-et-livrables)
4. [Difficultés rencontrées et solutions](#4-difficultés-rencontrées-et-solutions)
5. [Conclusion](#5-conclusion)
6. [Annexe](#6-annexe)

---

## 1. Contexte et objectif

### Contexte

Pour la session 2026, la CNBAU organise les délibérations des bourses de coopération marocaine pour deux programmes :

- **Formation universitaire (FU)** : 417 dossiers et 80 places, aux niveaux Licence, Master, Doctorat et Spécialité médicale.
- **Formation professionnelle (FP)** : 70 dossiers au niveau Bac + 2 ans, avec 50 places contingentées par filière et 10 places sans quota de filière.

Avant ce projet, la commission gérait les délibérations manuellement : annotation de tableaux Excel, calcul manuel des quotas, rédaction manuelle des documents officiels. Ce processus était lent, source d'erreurs et difficile à tracer.

### Objectif

Développer une application web **légère, déployable localement**, permettant à la commission de :

- **Charger** les dossiers de candidature depuis un fichier Excel
- **Parcourir et filtrer** les candidatures en temps réel
- **Attribuer** les avis (Favorable, Défavorable, Suppléant, En attente) avec contrôle automatique des quotas
- **Réallouer** les places non utilisées entre filières sans modifier le total de 80 bourses
- **Exporter** les décisions sous forme de documents officiels (Word et Excel)

---

## 2. Méthodologie technique

### Stack technologique

| Composant | Technologie |
|-----------|-------------|
| Interface utilisateur | [Streamlit](https://streamlit.io/) ≥ 1.33 |
| Base de données | SQLite (via `sqlite3` natif Python) |
| Traitement des données | Pandas, OpenPyXL |
| Export Word | python-docx |
| Export Excel | openpyxl |

### Architecture

L'application suit une architecture modulaire en 5 fichiers principaux :

```
gb-dbau/
├── app.py          # Interface principale — 5 onglets Streamlit
├── database.py     # Couche de données — SQLite, imports, exports
├── ui_helper.py    # Composants HTML réutilisables
├── style.py        # CSS, thème clair/sombre, header sticky (JS)
└── quotas.json     # Configuration des places par filière et niveau
```

### Base de données

Deux tables SQLite gèrent la session en cours :

**`candidatures`** — stocke les dossiers chargés depuis l'Excel :

| Colonne | Type | Description |
|---------|------|-------------|
| `id_demande` | TEXT (PK) | Identifiant unique du dossier |
| `numero` | INTEGER | Numéro d'ordre |
| `name` | TEXT | Nom et prénom |
| `filiere` | TEXT | Filière d'études |
| `niveau_etudes` | TEXT | Niveau (Bac + 2 ans / Licence / Master / Doctorat / Spécialité médicale) |
| `moyenne` | TEXT | Moyenne générale (virgule ou point décimal) |
| `avis` | TEXT | Décision : `Favorable` / `Défavorable` / `Suppléant` / `En attente` |

**`quotas`** — reflète la configuration des places par (niveau, filière) :

| Colonne | Type | Description |
|---------|------|-------------|
| `niveau_etudes` | TEXT (PK) | Niveau d'études |
| `filiere` | TEXT (PK) | Filière |
| `nb_places` | INTEGER | Nombre de places allouées |

### Flux de données

```
Fichier Excel CNBAU
       │
       ▼
_parse_real_excel()     ← normalisation niveau + filière
       │
       ▼
Table SQLite candidatures
       │
       ├──► Onglet Liste       (filtres, pagination, actions)
       ├──► Onglet Quotas      (barres de progression temps réel)
       ├──► Onglet Examen      (recherche, décision individuelle)
       ├──► Onglet Réallocation (transfert de places inter-filières)
       └──► Onglet Export      (Word, Excel)
```

---

## 3. Résultats et livrables

### Fonctionnalités livrées

#### Onglet 1 — Liste des candidatures
- Filtres croisés par niveau, filière et avis
- Tableau paginé (15 lignes/page) trié par niveau → filière → moyenne
- Boutons d'action rapide par candidat (Favorable / Défavorable / Suppléant / En attente)
- Bouton « Favorable » automatiquement désactivé si le quota de la filière est atteint

#### Onglet 2 — Suivi des quotas
- Grille de cartes organisée par niveau d'études
- Barre de progression par filière avec compteur en temps réel (X/N places)
- Indicateur visuel : bleu = places disponibles, vert = complet

#### Onglet 3 — Examen individuel
- Recherche par numéro d'ordre ou nom (avec recherche approximative)
- Fiche détaillée du candidat (état civil, diplôme, observation)
- Mini-indicateur de quota contextualisé à la filière du candidat
- Attribution d'avis avec protection contre le dépassement de quota

#### Onglet 4 — Réallocation des quotas
- Formulaire de transfert de places d'une filière source vers une filière destination
- Sélection réactive : les filières se filtrent dynamiquement selon le niveau choisi
- Validations : places disponibles suffisantes, source ≠ destination, total = 80
- Historique des transferts de la session (source, destination, nombre de places, heure)

#### Onglet 5 — Export
Quatre formats d'export disponibles :

| Format | Contenu |
|--------|---------|
| Word — Titulaires & Suppléants | Document officiel avec les Favorables et Suppléants, organisé par niveau et filière |
| Word — Toutes les décisions | Document complet : Titulaires + Suppléants + Défavorables |
| Excel — Candidatures par avis | Classeur avec une feuille par catégorie d'avis |
| Excel — Grille des quotas | Tableau de synthèse : filière / places / favorables / restantes |

#### Sidebar
- Indicateur de progression global : `X / 80 bourses accordées`
- Barre de progression colorée (jaune en cours, verte à 100%)
- Bouton de réinitialisation de session

---

## 4. Difficultés rencontrées et solutions

### 4.1 — Parsing du fichier Excel CNBAU

**Problème** : Le format Excel fourni par la CNBAU n'est pas tabulaire standard. Il contient des cellules fusionnées, des en-têtes `NIVEAU: ...` et `FILIERE: ...` intercalés entre les lignes de données, et des variantes orthographiques multiples (`LICENCE` / `Licence` / `licence`).

**Solution** : Développement d'un parseur dédié `_parse_real_excel()` qui :
- Détecte les marqueurs de niveau et filière à la volée
- Normalise les niveaux via `_normalize_niveau()` (expressions régulières)
- Mappe les noms de filières vers les noms canoniques de `quotas.json` via `_normalize_filiere()`

---

### 4.2 — Moyenne avec séparateur décimal variable

**Problème** : Les moyennes dans le fichier Excel peuvent utiliser une virgule (`14,5`) ou un point (`14.5`) comme séparateur décimal. Python's `float()` échoue silencieusement sur les virgules et affichait `0.00`.

**Solution** : Remplacement systématique de la virgule avant la conversion :
```python
raw_moy = str(row.get("moyenne") or 0).replace(",", ".")
moyenne = float(raw_moy)
```

---

### 4.3 — Réactivité des filtres dans un formulaire Streamlit

**Problème** : À l'intérieur d'un `st.form()`, Streamlit ne déclenche pas de rerun lors du changement d'un `selectbox`. Les filières ne se mettaient donc pas à jour dynamiquement lors du changement de niveau dans l'onglet Réallocation.

**Solution** : Déplacement des selectboxes niveau/filière **hors du formulaire** pour qu'ils déclenchent un rerun immédiat. Seuls le `number_input` et le bouton de soumission restent dans le `st.form()`.

---

### 4.4 — Header sticky avec synchronisation des KPIs

**Problème** : Streamlit ne supporte pas nativement les éléments CSS `position: sticky` sur ses composants. Il fallait que le bandeau de KPIs reste visible lors du scroll.

**Solution** : Injection d'un script JavaScript (via `st.components.v1.html`) utilisant un `IntersectionObserver` pour détecter le scroll et cloner dynamiquement la ligne de KPIs dans un header fixe. Un debounce de 150 ms évite les boucles infinies causées par les mutations DOM de Streamlit.

---

### 4.5 — Sidebar désynchronisée après réallocation

**Problème** : La sidebar lisait les quotas directement depuis `quotas.json` (fichier statique), ignorant les réallocations effectuées en session dans la BDD. Le total affiché restait figé à 150 même après transfert.

**Solution** : Remplacement de la lecture du fichier JSON par l'utilisation des quotas déjà chargés depuis la BDD :
```python
# Avant
total = sum(sum(cat.values()) for cat in json.load(open("quotas.json")))

# Après
total = sum(quotas.values())  # quotas déjà lu depuis db.get_quotas()
```

---

### 4.6 — Performance avec plusieurs centaines de candidats

**Problème** : Chaque interaction Streamlit déclenche un rerun complet du script Python, causant des requêtes SQL répétitives et un affichage lent.

**Solution** : Mise en cache avec TTL de 2 secondes via `@st.cache_data(ttl=2)` sur les fonctions de lecture, avec invalidation manuelle après chaque écriture (`invalidate_cache()`). Une seule requête SQL agrégée (`get_stats()`) remplace plusieurs `COUNT` séparés.

---

## 5. Conclusion

L'application CNBAU répond intégralement aux besoins de la commission : chargement du fichier Excel existant, délibération assistée et contrôlée par quota, réallocation traçable des places non utilisées, et génération automatisée des documents officiels.

L'architecture légère (Streamlit + SQLite, sans serveur dédié) permet un déploiement immédiat sur un poste local sans infrastructure particulière. Les données de session persistent dans un fichier SQLite et ne sont jamais perdues lors d'un redémarrage ou d'une mise à jour du code — seul un rafraîchissement de page suffit pour appliquer les correctifs en cours de session.

Le projet a été livré avec **43 commits**, couvrant le développement initial, les optimisations de performance, les corrections de bugs identifiés en session réelle, et les évolutions fonctionnelles (réallocation, exports multiples).

---

## 6. Annexe

### Installation et lancement

```bash
# Cloner le dépôt
git clone https://github.com/Sena2804/gb-dbau.git
cd gb-dbau

# Installer les dépendances
pip install -r requirements.txt

# Lancer l'application
streamlit run app.py
```

### Premier démarrage

1. L'application s'ouvre sur une page de chargement
2. Importer le fichier Excel des candidatures (`.xlsx`)
3. Le fichier `quotas.json` est chargé automatiquement
4. La session est prête

### Mise à jour en cours de session

Une mise à jour du code (`git pull`) ne nécessite pas de recharger les données. Un simple **rafraîchissement de la page (F5)** suffit — la BDD SQLite est préservée.

### Réinitialiser la session

Cliquer sur **"Réinitialiser la session"** dans la sidebar pour effacer toutes les données et recommencer depuis zéro.

### Structure du fichier `quotas.json`

```json
{
  "Licence": {
    "Nom de la filière": nombre_de_places
  },
  "Master": { ... },
  "Doctorat": { ... },
  "Spécialité médicale": { ... }
}
```

Le fichier `quotas.json` sert de configuration de référence pour la formation universitaire. Lors d'un import FU ou FP, les quotas inscrits dans le classeur sont chargés automatiquement et remplacent cette configuration pour la session courante.

### Format du fichier Excel attendu

Le parseur supporte le format Excel structuré de la CNBAU avec :
- Des lignes d'en-tête `NIVEAU: ...` et `FILIERE: ...`
- Des colonnes : N°, Sexe, Nom, Date/Lieu de naissance, Diplôme/Filière/Année, Moyenne/Mention, Observations, Avis CNaBAU
- Des quotas indiqués dans chaque ligne de filière, par exemple `(11 places)`
- Les fichiers `Tableau_FU_MAROC2026.xlsx` et `Tableau_FP MAROC2026.xlsx`
- La suppression sécurisée des lignes portant un numéro de dossier dupliqué, en conservant la première occurrence

### Dépendances

```
streamlit>=1.33
pandas
openpyxl
python-docx
```
