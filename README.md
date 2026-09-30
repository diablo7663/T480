# T480 — marché des voitures d'occasion (Maroc)

Scraper Python qui transforme les pages de recherche des petites annonces
marocaines en données structurées, puis en un classeur Excel prêt à analyser.
Le dépôt couvre trois sites :

| Site | Annonces propres | Mécanisme | robots.txt |
| --- | --- | --- | --- |
| **avito.ma** | ~26 900 | payload JSON `__NEXT_DATA__` embarqué dans chaque page | autorisé |
| **oneclickdrive.ma** | ~6 100 | JSON-LD `ItemList` serveur, 8 villes, 20 annonces/page | autorisé (`Allow: /*?page=`) |
| **occasion.kifal.ma** | ~400 | cartes serveur + tableau de fiche détaillée | autorisé (`Disallow:` vide) |

Aucun navigateur headless : `requests` + `BeautifulSoup`, poli par hôtes
(1,5 s minimum entre deux requêtes, backoff `Retry-After`).

Le dépôt contient :

- le scraper avito (`scraper.py`),
- un paquet de sources additionnelles (`sources/`) avec un module par site,
- le fusionneur multi-sites (`merge.py`),
- un workflow GitHub Actions qui fait tourner la collecte et publie les
  classeurs (`T480_maroc.xlsx` + `avito_cars.xlsx` + CSV),
- le cache brut des pages (`data/cache/*.tar.gz`) pour reprendre ou
  recalculer les données à tout moment.

**Le dépôt est privé** : les annonces contiennent des numéros de téléphone de
vendeurs.

## Install

```bash
python3 -m pip install -r requirements.txt
```

## Usage

```bash
# 1. crawler avito (≈ 28 000 annonces brutes)
python3 scraper.py

# 2. crawler les autres sites (kifal + oneclickdrive, ≈ 6 600 annonces)
python3 -c "
from sources.kifal import crawl as k
from sources.oneclickdrive import crawl as o
k('data'); o('data')
"

# 3. fusionner, nettoyer et publier le classeur unifié
python3 merge.py
```

Sorties :

| Fichier | Contenu |
| --- | --- |
| `data/T480_maroc.xlsx` | classeur unifié (Résumé, Annonces, À vérifier, Doublons entre sites) |
| `data/T480_maroc.csv` / `.json` | les mêmes lignes, plates |
| `data/avito_cars.xlsx` / `.csv` | export avito seul (inchangé) |

```bash
# re-publier les classeurs depuis le cache, sans toucher au réseau
python3 scraper.py --export-only --skip-details
python3 merge.py

# rafraîchir les 40 premières pages avito (nouvelles annonces)
python3 scraper.py --refresh-top 40
```

Les crawls sont interruptibles (Ctrl-C) et reprennent à la dernière page
terminée : chaque page et chaque fiche sont mises en cache une par une.

## Options

| Flag | Défaut | Rôle |
| --- | --- | --- |
| `--url` | `.../fr/maroc/voitures_a_vendre` | recherche de départ (avito) |
| `--output-dir` | `data` | cache + exports |
| `--max-pages` | `0` | s'arrêter après N pages (`0` = toutes) |
| `--delay` | `1.5` | secondes entre deux requêtes (politesse) |
| `--refresh-top` | `0` | re-télécharger les N premières pages |
| `--fresh` | — | effacer le cache avito et repartir de zéro |
| `--export-only` | — | régénérer les exports avito depuis le cache |
| `--pending-details` | — | compter les fiches détaillées manquantes |
| `--all-categories` | — | garder aussi location et leasing |

`merge.py` : `--skip-avito` (fusionner uniquement `sources/`), `--out NOM`
(base de sortie), `--no-excel`.

## Le classeur unifié

`data/T480_maroc.xlsx` contient quatre onglets :

1. **Résumé** — chiffres clés (volume, prix moyen/médian, min/max, marques,
   villes), répartition **par site**, des prix, des millésimes, carburant,
   boîte, complétude de chaque champ, et la note de méthode.
2. **Annonces** — une ligne par voiture, filtrable, avec `Marque`, `Modèle`,
   `Année`, `Prix` en tête, puis l'identité (`Site`, `Lien`, `Titre`) et tous
   les champs supplémentaires.
3. **À vérifier** — uniquement les lignes incomplètes (champ(s) manquant(s),
   ville, site, vendeur, lien) pour une reprise manuelle rapide.
4. **Doublons entre sites** — les voitures trouvées sur deux sites à la fois
   (même marque, modèle, année, prix, ville **et** kilométrage). Elles ne
   sont **jamais** fusionnées entre sites : le rapport les liste pour qu'un
   humain décide.

## Nettoyage

`clean_records()` est appliqué avant chaque export :

- **Prix aberrants écartés** — une voiture ne se vend pas 300 DH, et le site
  remplit le champ prix avec un nombre placeholder (ou un numéro de
  téléphone mal lu) quand le vendeur dit « prix à discuter ». Les valeurs hors
  **5 000 – 5 000 000 DH** sont supprimées.
- **Doublons regroupés, site par site** — les annonces partageant source,
  marque, modèle, année, kilométrage, ville, vendeur et prix sont fusionnées
  en une ligne, la plus complète étant conservée. Le kilométrage exact
  empêche qu'un concessionnaire qui aligne trois fois le même modèle au même
  prix soit réduit à une seule ligne, et la source dans la clé garantit
  qu'aucune fusion ne traverse deux sites. La colonne **Annonces regroupées**
  indique combien d'annonces représentent chaque ligne ; un export déjà
  nettoyé peut être re-nettoyé sans perdre ce compteur.

Les données brutes ne sont jamais perdues : `data/cache/*.tar.gz` conserve les
pages et les fiches telles que téléchargées.

## Champs

Une ligne de base : Marque, Modèle, Année, Prix, Kilométrage, Carburant, Boîte
de vitesses, Ville, Vendeur, Url, Titre, Description, Site, puis tous les
paramètres optionnels (puissance, portes, origine, état, première main,
carrosserie, couleur, options, photos…). Les colonnes propres à un site
restent vides sur les autres lignes — jamais inventées.

Les quatre champs essentiels sont renseignés dans cet ordre :

1. la fiche annonce (fiable),
2. déduits du titre (`infer_brand_model`),
3. extraits de la description (`price_from_text`) si le vendeur cache le prix.

`Origine du prix` indique la source : `annonce`, `detail` ou `description` ;
vide si l'annonce n'a réellement pas de prix. Un prix n'est jamais inventé.

## Sites écartés

Triage des autres petites annonces automobiles marocaines — chaque site a été
testé (robots.txt + structure) avant décision :

| Site | Décision | Raison |
| --- | --- | --- |
| `moteur.ma` | écarté | republie le flux d'avito.ma (images `content.avito.ma/...?t=moteur_feed`), ~96 % de doublons directs |
| `marocannonces.com` | écarté | `robots.txt` : `User-agent: * Disallow: /` |
| `voiturenet.ma` | écarté | challenge Cloudflare (« Just a moment »), pas de contournement |
| `auto24.ma` | écarté | SPA pure : toutes les URLs renvoient la même coquille de 2 Ko, aucune API découverte |
| `simmo.ma` | écarté | annonces rendues côté client, uniquement via `/api` (interdit par `robots.txt`) |
| `occaro.com` | écarté | `robots.txt` interdit `/car/*` et `/cars*` |
| `opensooq.com` (MA) | écarté | `robots.txt` : `Disallow: /*`, et 16 annonces seulement |
| `marodrive.ma` | écarté | serveur qui ne répond pas (timeout) |
| `chad.ma` | écarté | domaine qui ne résout pas |
| `agenz.ma` | écarté | immobilier uniquement |

Ajouter un site = créer `sources/<site>.py` qui émet des lignes au format
exact d'avito via `sources/common.make_record()`, puis le laisser tomber dans
`data/sources/<site>/annonces.json` : `merge.py` le découvre tout seul.

## GitHub Actions

`.github/workflows/scrape.yml` fait tourner la collecte sur les runners
GitHub. Seuls les caches compressés `data/cache/*.tar.gz` sont versionnés ; le
job les décompresse et reprend où le run précédent s'était arrêté.

- **Lancer maintenant :** Actions → *T480 — voitures d'occasion (Maroc)* →
  *Run workflow* (`detail_workers`, `detail_delay`, `max_details`,
  `refresh_top`, `refresh_sources`).
- **Nuit :** chaque nuit à 4h17, le run rafraîchit la tête des quatre flux
  (triés du plus récent au plus ancien) pour récupérer les nouvelles
  annonces, complète les fiches manquantes, puis refait la fusion.
- Chaque run publie deux artefacts — `T480-avito-cars` et
  `T480-maroc-cars` — et committe le cache rafraîchi.

Le débit est volontairement bas (3 workers, 0,5 s d'écart sur avito ;
1,5 s entre deux requêtes par hôtes ailleurs) avec backoff `Retry-After`,
car avito.ma répond aux rafales par des HTTP 429.

## Notes

- Avito tourne sur Next.js ; si la structure change, seul `parse_page` /
  `extract_ad` est à mettre à jour. Les autres sites suivent la même règle :
  la structure d'un site ne touche que son module dans `sources/`.
- Dépasser ~1 requête toutes les 1,5 s expose au blocage.
- Respecter les conditions d'utilisation des sites et le RGPD : les données
  contiennent des informations personnelles.
