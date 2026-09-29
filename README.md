# T480 — marché des voitures d'occasion (avito.ma)

Scraper Python qui transforme les pages de recherche d'avito.ma en données
structurées, puis en un classeur Excel prêt à analyser. Il lit le payload
JSON `__NEXT_DATA__` embarqué dans chaque page — pas de navigateur headless.

Le dépôt contient :

- le scraper (`scraper.py`),
- un workflow GitHub Actions qui fait tourner la collecte et publie le
  classeur (`T480-avito-cars.xlsx` + CSV),
- le cache brut des pages (`data/cache/*.tar.gz`) pour pouvoir reprendre ou
  recalculer les données à tout moment.

**Le dépôt est privé** : les annonces contiennent des numéros de téléphone de
vendeurs.

## Install

```bash
python3 -m pip install -r requirements.txt
```

## Usage

```bash
# tout crawler (≈ 28 000 annonces)
python3 scraper.py

# re-publier le classeur depuis le cache, sans toucher au réseau
python3 scraper.py --export-only --skip-details

# rafraîchir les 40 premières pages (nouvelles annonces) et garder le reste
python3 scraper.py --refresh-top 40

# repartir de zéro
python3 scraper.py --fresh
```

Le crawl est interruptible (Ctrl-C) et reprend à la dernière page terminée.

## Options

| Flag | Défaut | Rôle |
| --- | --- | --- |
| `--url` | `.../fr/maroc/voitures_a_vendre` | recherche de départ |
| `--output-dir` | `data` | cache + exports |
| `--max-pages` | `0` | s'arrêter après N pages (`0` = toutes) |
| `--delay` | `1.5` | secondes entre deux requêtes (politesse) |
| `--timeout` | `30` | timeout HTTP |
| `--retries` | `5` | tentatives par page |
| `--refresh-top` | `0` | re-télécharger les N premières pages |
| `--fresh` | — | effacer le cache et repartir de zéro |
| `--export-only` | — | régénérer les exports depuis le cache |
| `--pending-details` | — | compter les fiches détaillées manquantes |
| `--all-categories` | — | garder aussi location et leasing |

## Le classeur

`data/avito_cars.xlsx` contient trois onglets :

1. **Résumé** — chiffres clés (volume, prix moyen/médian, min/max, marques,
   villes), répartition des prix et des millésimes, complétude de chaque
   champ, et la note de méthode.
2. **Annonces** — une ligne par voiture, filtrable, avec `Marque`, `Modèle`,
   `Année`, `Prix` en tête puis tous les champs supplémentaires.
3. **À vérifier** — uniquement les lignes incomplètes (champ(s) manquant(s),
   ville, vendeur, lien) pour une reprise manuelle rapide.

## Nettoyage

`clean_records()` est appliqué avant chaque export :

- **Prix aberrants écartés** — une voiture ne se vend pas 300 DH, et le site
  remplit le champ prix avec un nombre placeholder (ou un numéro de
  téléphone mal lu) quand le vendeur dit « prix à discuter ». Les valeurs hors
  **5 000 – 5 000 000 DH** sont supprimées.
- **Doublons regroupés** — un même vendeur ré-publie souvent la même voiture
  plusieurs fois. Les annonces partageant marque, modèle, année, prix, ville et
  vendeur sont fusionnées en une ligne, la plus complète étant conservée.
  La colonne **Annonces regroupées** indique combien d'annonces
  représentent chaque ligne.

Les données brutes ne sont jamais perdues : `data/cache/*.tar.gz` conserve les
pages et les fiches telles que téléchargées.

## Champs

Une ligne de base : Marque, Modèle, Année, Prix, Kilométrage, Carburant, Boîte
de vitesses, Ville, Vendeur, Url, Titre, Description, puis tous les paramètres
optionnels de la fiche (puissance, portes, origine, état, première main,
options, photos…).

Les quatre champs essentiels sont renseignés dans cet ordre :

1. la fiche annonce (fiable),
2. déduits du titre (`infer_brand_model`),
3. extraits de la description (`price_from_text`) si le vendeur cache le prix.

`Origine du prix` indique la source : `annonce`, `detail` ou `description` ;
vide si l'annonce n'a réellement pas de prix. Un prix n'est jamais inventé.

## GitHub Actions

`.github/workflows/scrape.yml` fait tourner la collecte sur les runners
GitHub. Seuls les caches compressés `data/cache/*.tar.gz` sont versionnés ; le
job les décompresse et reprend où le run précédent s'était arrêté.

- **Lancer maintenant :** Actions → *T480 — avito.ma* → *Run workflow*
  (`detail_workers`, `detail_delay`, `max_details`, `refresh_top`).
- **Nuit :** chaque nuit à 4h17, le run rafraîchit les 40 premières pages du
  flux (trié du plus récent au plus ancien) pour récupérer les nouvelles
  annonces, puis complète les fiches manquantes.
- Chaque run publie `avito_cars.xlsx` + `avito_cars.csv` en artifact et
  committe le cache rafraîchi.

Le débit est volontairement bas (3 workers, 0,5 s d'écart) avec backoff
`Retry-After` et jitter, car avito.ma répond aux rafales par des HTTP 429.

## Notes

- Avito tourne sur Next.js ; si la structure change, seul `parse_page` /
  `extract_ad` est à mettre à jour.
- Dépasser ~1 requête toutes les 1,5 s expose au blocage.
- Respecter les conditions d'utilisation d'avito.ma et le RGPD : les données
  contiennent des informations personnelles.
