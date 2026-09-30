#!/usr/bin/env python3
"""Build the unified multi-source workbook for T480.

    python3 merge.py                  # avito.ma + every crawled sources/ site
    python3 merge.py --skip-avito      # only the other Moroccan sites
    python3 merge.py --out T480_maroc  # change the output basename

avito.ma is read from its raw page cache when it exists (one clean pass over
the original rows); otherwise the already-cleaned avito_cars.json is reused,
which is safe because clean_records carries each row's merged count forward.
Every data/sources/<site>/annonces.json joins the same union.

Outputs under data/:
    <out>.json   one row per car (union of every site)
    <out>.csv    same, flat
    <out>.xlsx   French workbook: Résumé, Annonces, À vérifier and — when more
                 than one site is present — Doublons entre sites
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from collections import Counter, defaultdict
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from scraper import (clean_records, export_excel, fold, load_raw_records,
                     source_label, write_csv)

log = logging.getLogger("t480.merge")


def load_sources(data_dir):
    """Every data/sources/<site>/annonces.json next to its page cache."""
    records = []
    root = data_dir / "sources"
    if not root.exists():
        return records
    for path in sorted(root.glob("*/annonces.json")):
        rows = json.loads(path.read_text(encoding="utf-8"))
        site = path.parent.name
        for row in rows:
            row.setdefault("source", site)
        records.extend(rows)
        log.info("source %s: %s listings", site, len(rows))
    return records


def _cross_key(rec):
    """Same-car signature ignoring the site: price and odometer included."""
    params = rec.get("other_params") or {}
    parts = []
    for field in ("brand", "model", "year", "price", "location"):
        value = params.get(field) if field in ("brand", "model") else rec.get(field)
        if value in (None, ""):
            return None
        parts.append(fold(value) if isinstance(value, str) else value)
    parts.append(rec.get("mileage_km"))  # exact odometer: same car, twice
    return tuple(parts)


def cross_site_duplicates(records):
    """Cars present on two or more sites — reported, never merged.

    Matching requires the same marque/modèle/année/prix/ville/kilométrage so
    that a popular model listed by many different sellers does not flood the
    sheet with false pairs; what remains is almost always one dealer's stock
    posted on two marketplaces at once.
    """
    groups = defaultdict(list)
    for rec in records:
        key = _cross_key(rec)
        if key is not None:
            groups[key].append(rec)

    report = []
    for group in groups.values():
        sites = {row.get("source") or "avito" for row in group}
        if len(sites) < 2:
            continue
        first = max(group, key=lambda row: row.get("merged_ads") or 1)
        params = first.get("other_params") or {}
        report.append({
            "brand": params.get("brand"),
            "model": params.get("model"),
            "year": first.get("year"),
            "price": first.get("price"),
            "location": first.get("location"),
            "sites": ", ".join(source_label(site) for site in sorted(sites)),
            "count": len(group),
            "sellers": " · ".join(
                f"{source_label(row.get('source'))}: "
                f"{row.get('seller_name') or '—'}"
                for row in sorted(group,
                                  key=lambda row: str(row.get("source")))),
            "urls": " | ".join(str(row.get("url")) for row in group),
        })
    report.sort(key=lambda row: (-row["count"], str(row.get("brand") or ""),
                                 str(row.get("model") or "")))
    return report


def _by_source(records):
    return Counter(row.get("source") or "avito" for row in records)


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Merge avito.ma with the other Moroccan car sources")
    parser.add_argument("--data-dir", default="data",
                        help="cache/export directory (default: data)")
    parser.add_argument("--out", default="T480_maroc",
                        help="output basename (default: T480_maroc)")
    parser.add_argument("--skip-avito", action="store_true",
                        help="merge only the sources/ caches")
    parser.add_argument("--allow-all-categories", action="store_true",
                        help="keep non-voiture avito ads too")
    parser.add_argument("--no-excel", action="store_true",
                        help="skip the xlsx workbook")
    args = parser.parse_args(argv)

    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s")
    data_dir = Path(args.data_dir)

    records = []
    if not args.skip_avito:
        if (data_dir / "pages").exists():
            avito, duplicates, enriched = load_raw_records(
                data_dir, args.allow_all_categories)
            log.info("avito: %s listings from the page cache (%s exact "
                     "duplicates skipped, %s enriched from detail cache)",
                     len(avito), duplicates, enriched)
            records.extend(avito)
        elif (data_dir / "avito_cars.json").exists():
            avito = json.loads((data_dir / "avito_cars.json")
                               .read_text(encoding="utf-8"))
            log.warning("avito: no page cache, reusing the cleaned export "
                        "(%s rows)", len(avito))
            records.extend(avito)
        else:
            log.warning("avito: neither data/pages nor avito_cars.json found")

    records.extend(load_sources(data_dir))
    if not records:
        log.error("nothing to merge under %s", data_dir)
        return 1

    raw_counts = _by_source(records)
    cleaned, stats = clean_records(records)
    final_counts = _by_source(cleaned)

    log.info("raw: %s listings", sum(raw_counts.values()))
    for source, count in sorted(raw_counts.items()):
        log.info("  %-14s %6s -> %6s", source_label(source), count,
                 final_counts.get(source, 0))
    log.info("cleaning: dropped %s rows with an impossible price, collapsed "
             "%s re-posted ads into %s rows (%s ads represented in total)",
             stats.get("dropped_bad_price", 0),
             stats.get("merged_ads", 0), stats.get("merged_groups", 0),
             len(cleaned) + stats.get("merged_total", 0))

    json_out = data_dir / f"{args.out}.json"
    json_out.write_text(json.dumps(cleaned, ensure_ascii=False, indent=1),
                        encoding="utf-8")
    csv_path = data_dir / f"{args.out}.csv"
    columns = write_csv(cleaned, csv_path)

    cross = cross_site_duplicates(cleaned)
    if cross:
        log.info("cross-site duplicates: %s cars listed on 2+ sites "
                 "(sheet « Doublons entre sites »)", len(cross))
    else:
        log.info("cross-site duplicates: none")

    if not args.no_excel:
        xlsx_path = data_dir / f"{args.out}.xlsx"
        export_excel(cleaned, xlsx_path, stats, cross)
        log.info("wrote %s", xlsx_path)
    log.info("wrote %s (%s rows, %s columns)", json_out, len(cleaned),
             len(columns))
    log.info("wrote %s", csv_path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
