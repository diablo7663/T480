# avito.ma car listings scraper

Python scraper that dumps the car listings database on avito.ma as structured
data. It reads the `__NEXT_DATA__` JSON payload embedded in each search page —
no headless browser needed. Listings are deduplicated, cached page-by-page for
resumability, and exported as CSV + JSON.

## Install

```bash
python3 -m pip install -r requirements.txt
```

## Usage

```bash
# crawl everything (all ~34k listings on avito.ma)
python3 scraper.py

# only the first 10 result pages, fast
python3 scraper.py --max-pages 10 --delay 0.5

# scrape a different search (e.g. brand-new cars)
python3 scraper.py --url 'https://www.avito.ma/fr/maroc/voitures_neuves'

# re-build the CSV/JSON exports from already-cached pages only
python3 scraper.py --export-only

# start over from scratch
python3 scraper.py --fresh
```

You can stop the crawl at any time (Ctrl-C). Re-running resumes from the last
completed page.

## Options

| Flag | Default | Meaning |
| --- | --- | --- |
| `--url` | `.../fr/maroc/voitures_a_vendre` | Base search URL to scrape |
| `--output-dir` | `data` | Where cached pages + exports live |
| `--max-pages` | `0` | Stop after N pages (`0` = all) |
| `--delay` | `1.5` | Seconds between page requests (be polite) |
| `--timeout` | `30` | HTTP timeout per request |
| `--retries` | `5` | Retries per page before giving up |
| `--fresh` | — | Clear cached pages and restart |
| `--export-only` | — | Only regenerate exports from cache |
| `--verbose` | — | Debug logging |

## Output

- `data/pages/page_XXXX.json` — one file per result page (raw extracted ads)
- `data/progress.json` — resume state
- `data/avito_cars.json` — all unique listings, with `images[]` lists
- `data/avito_cars.csv` — flat table (UTF-8 BOM, Excel-friendly)

Fields per listing: id, list_id, url, title, description, category,
ad_type, price, currency, monthly_payment, old_price, year, mileage_km,
fuel, gearbox, extra params, location, city/area ids, date posted,
seller (name/type/phone/verified), and flags like `is_professional`,
`is_premium`, `is_urgent`, `is_car_checked`, plus all photo URLs.

## Notes

- Avito runs on a Next.js stack; if it starts serving a different layout the
  `__NEXT_DATA__` payload may change. The extraction lives in `parse_page` /
  `extract_ad` and is the only place that needs updating.
- Increasing volume beyond ~1 request every 1.5 s risks rate-limiting/blocking.
- Respect avito.ma's terms of service and robots.txt when using scraped data.