"""www.publimaroc.com — Publimaroc's «Voitures occasion» classifieds.

robots.txt explicitly allows this crawl (``User-agent: *`` → ``Allow: /``)
and only keeps crawlers out of account, contact and search endpoints, so the
category pages and the ad pages are fair game. The category is
server-rendered — 18 ads per page behind ``?page=N`` — and each ad page adds
the structured spec cards (marque, modèle, année, carburant, boîte, état),
the city, the seller's description and name. Phone numbers are reserved for
logged-in members, so the phone columns stay empty rather than guessed.

Cache layout (mirrors scraper.py):
    <output>/sources/publimaroc/pages/page_0001.json   parsed search cards
    <output>/sources/publimaroc/details/<list_id>.json parsed detail specs
    <output>/sources/publimaroc/annonces.json          merged export
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

from .common import (fetch, first_year, make_record, money, new_session,
                     read_json, slug_key, soup, to_int, write_json)

log = logging.getLogger("t480.publimaroc")

BASE = "https://www.publimaroc.com"
LIST_URL = BASE + "/categorie/21/voitures-occasion?page={page}"
SOURCE = "publimaroc"
MAX_PAGES = 40

# Detail-page label -> the field name scraper.py already uses (avito parity).
FIELD_MAP = {
    "marque": "brand",
    "modèle": "model",
    "modele": "model",
    "année": "year",
    "annee": "year",
    "carburant": "fuel",
    "boîte de vitesses": "gearbox",
    "boite de vitesses": "gearbox",
    "kilométrage": "mileage_km",
    "kilometrage": "mileage_km",
    "ville": "location",
    "état": ("params", "auto_condition"),
    "etat": ("params", "auto_condition"),
}

# The "Type" card is the poster's intent, i.e. avito's ad_type.
AD_TYPES = {"je vends": "à vendre", "je cherche": "recherche"}


def _absolute(href):
    if not href:
        return None
    return BASE + href if href.startswith("/") else href


def parse_search(html):
    """One category page -> list of records (18 ads per page)."""
    doc = soup(html)
    records = []
    seen = set()
    for card in doc.select(".ccat-card"):
        url = _absolute(card.get("href"))
        match = re.search(r"/annonce/(\d+)/", url or "")
        if not url or not match:
            continue
        list_id = f"pm_{match.group(1)}"
        if list_id in seen:
            continue
        seen.add(list_id)

        title_el = card.select_one(".ccat-title")
        price_el = card.select_one(".ccat-price")
        image = card.find("img")
        image_url = _absolute(image.get("src") or image.get("data-src")) if image else None

        price = money(price_el.get_text(" ", strip=True)) if price_el else None
        records.append(make_record(
            SOURCE, list_id, url,
            title=" ".join(title_el.get_text(" ", strip=True).split())
            if title_el else None,
            price=price,
            price_source="annonce" if price else None,
            images=[image_url] if image_url else [],
        ))
    return records


def parse_detail(html):
    """Ad page -> structured specs, city, description, poster name."""
    doc = soup(html)
    out = {"params": {}, "labels": {}}

    title = doc.select_one("h1.detail-title")
    if title:
        out["title"] = " ".join(title.get_text(" ", strip=True).split())

    price = doc.select_one(".detail-hero-price")
    if price:
        out["price"] = money(price.get_text(" ", strip=True))

    city = doc.select_one(".detail-location-text")
    if city:
        out["location"] = " ".join(city.get_text(" ", strip=True).split())

    for spec in doc.select(".spec-card"):
        label_el = spec.select_one(".spec-card-label")
        value_el = spec.select_one(".spec-card-value")
        if label_el is None or value_el is None:
            continue
        label = " ".join(label_el.get_text(" ", strip=True).split())
        value = " ".join(value_el.get_text(" ", strip=True).split())
        if not label or not value:
            continue

        key = label.lower()
        if key == "type":
            ad_type = AD_TYPES.get(value.lower())
            if ad_type:
                out["ad_type"] = ad_type
            continue

        target = FIELD_MAP.get(key)
        if target == "year":
            out["year"] = first_year(value)
        elif target == "mileage_km":
            out["mileage_km"] = to_int(value)
        elif isinstance(target, tuple) and target[0] == "params":
            out["params"][target[1]] = value
            out["labels"][target[1]] = label
        elif target in ("brand", "model"):
            out[target] = value
        elif target:
            out[target] = value
        else:
            param = slug_key(label)
            out["params"][param] = value
            out["labels"][param] = label

    seller = doc.select_one(".contact-name")
    if seller:
        text = " ".join(seller.get_text(" ", strip=True).split())
        out["seller_name"] = text or None

    description = doc.select_one(".description-content")
    if description:
        text = " ".join(description.get_text(" ", strip=True).split())
        out["description"] = text or None

    images = []
    for img in doc.select(".detail-gallery-wrap img, .detail-gallery img, "
                          ".detail-card img"):
        src = _absolute(img.get("src") or img.get("data-src"))
        if src and src not in images:
            images.append(src)
    out["images"] = images
    return out


def enrich(record, details_dir, session):
    """Merge one ad's detail page into its record (cached per listing)."""
    cache = details_dir / f'{record["list_id"]}.json'
    detail = read_json(cache)
    if detail is None:
        html = fetch(session, record["url"])
        if html is None:
            write_json(cache, {"failed": True})
            return record
        detail = parse_detail(html)
        write_json(cache, detail)
    if detail.get("failed"):
        return record

    params = record.setdefault("other_params", {})
    params.update({k: v for k, v in (detail.get("params") or {}).items()
                   if v not in (None, "", [])})
    labels = record.setdefault("detail_labels", {})
    labels.update(detail.get("labels") or {})

    for key in ("year", "mileage_km", "fuel", "gearbox", "location",
                "description", "seller_name", "ad_type", "title"):
        value = detail.get(key)
        if value not in (None, "", []):
            record[key] = value
    for key, label in (("brand", "Marque"), ("model", "Modèle")):
        value = detail.get(key)
        if value:
            params[key] = value
            labels.setdefault(key, label)
            record[key] = value

    if detail.get("images"):
        record["images"] = detail["images"]
        record["image_count"] = len(detail["images"])
        record["default_image"] = detail["images"][0]
    if not record.get("price") and detail.get("price"):
        record["price"] = detail["price"]
        record["currency"] = record.get("currency") or "DH"
        record["price_source"] = "detail"
    return record


def crawl(output_dir, *, max_pages=0, with_details=True, session=None,
          refresh_top=0):
    """Crawl the «Voitures occasion» category. Returns the record list.

    ``refresh_top`` re-fetches the first N category pages (newest-first) so a
    scheduled run picks up new ads without re-walking the whole pagination.
    """
    root = Path(output_dir) / "sources" / "publimaroc"
    pages_dir = root / "pages"
    details_dir = root / "details"
    pages_dir.mkdir(parents=True, exist_ok=True)
    details_dir.mkdir(parents=True, exist_ok=True)

    session = session or new_session()
    unique: dict[str, dict] = {}
    page = 1
    limit = max_pages or MAX_PAGES

    while page <= limit:
        cache = pages_dir / f"page_{page:04d}.json"
        cached = None if (refresh_top and page <= refresh_top) else read_json(cache)
        if cached is None:
            html = fetch(session, LIST_URL.format(page=page))
            if html is None:
                log.warning("page %s could not be fetched", page)
                break
            parsed = parse_search(html)
            write_json(cache, parsed)
        else:
            parsed = cached if isinstance(cached, list) else (cached.get("listings") or [])

        if not parsed:
            log.info("end of pagination at page %s", page)
            break

        before = len(unique)
        for rec in parsed:
            unique.setdefault(rec["list_id"], rec)
        log.info("page %s: %s ads (unique %s)", page, len(parsed), len(unique))
        if len(unique) == before:
            log.info("no new ads at page %s, stopping", page)
            break
        page += 1

    records = list(unique.values())
    if with_details:
        records = [enrich(rec, details_dir, session) for rec in records]

    write_json(root / "annonces.json", records)
    log.info("publimaroc: %s listings", len(records))
    return records
