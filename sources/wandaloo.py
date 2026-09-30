"""www.wandaloo.com — Wandaloo's used-car announcements for Morocco.

robots.txt is not published on this host (the URL answers with a
"contenu indisponible" page), so no crawling rules apply. The occasion feed
is server-rendered: 14 announcements per page behind ``?pg=N``, newest first.
Each detail page then adds a label/value specification block, the seller's
description, and — printed openly on the page — the seller's name and phone
number.

Cache layout (mirrors scraper.py):
    <output>/sources/wandaloo/pages/page_0001.json   parsed search cards
    <output>/sources/wandaloo/details/<list_id>.json parsed detail specs
    <output>/sources/wandaloo/annonces.json          merged export
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

from .common import (fetch, first_year, make_record, money, new_session,
                     read_json, slug_key, soup, to_int, write_json)

log = logging.getLogger("t480.wandaloo")

BASE = "https://www.wandaloo.com"
LIST_URL = BASE + "/occasion/?pg={page}"
SOURCE = "wandaloo"
MAX_PAGES = 60          # the feed is ~27 pages deep; this is a safety net

FUEL = {
    "diesel": "Diesel",
    "essence": "Essence",
    "essence (sans plomb)": "Essence",
    "hybride": "Hybride",
    "electrique": "Électrique",
    "électrique": "Électrique",
    "gpl": "GPL",
}
GEARBOX = {
    "automatique": "Automatique",
    "manuelle": "Manuelle",
    "semi-automatique": "Semi-automatique",
}
SELLER_TYPES = {
    "particulier": "PRIVATE",
    "professionnel": "STORE",
    "pro": "STORE",
}
MONTHS = {
    "janvier": 1, "février": 2, "fevrier": 2, "mars": 3, "avril": 4, "mai": 5,
    "juin": 6, "juillet": 7, "août": 8, "aout": 8, "septembre": 9,
    "octobre": 10, "novembre": 11, "décembre": 12, "decembre": 12,
}

# Detail-page label -> the field name scraper.py already uses (avito parity).
# The site calls the model year "Modèle", which is why that one maps to year.
FIELD_MAP = {
    "modèle": "year",
    "modele": "year",
    "carburant": "fuel",
    "kilométrage": "mileage_km",
    "kilometrage": "mileage_km",
    "transmision": "gearbox",
    "transmission": "gearbox",
    "ville": "location",
    "vendeur": "seller_type",
}

_CARD = re.compile(r"/occasion/.+/\d+\.html$")
_YEAR_ONLY = re.compile(r"(19[89]\d|20[0-3]\d)")


def _pick(mapping, value):
    if not value:
        return None
    return mapping.get(str(value).strip().lower(), str(value).strip())


def _chips(card):
    """Fuel / year / mileage out of the card's `<li>` chips.

    The chips have no classes and their order is not contractual, so each one
    is classified by what it contains instead of by position.
    """
    fuel = year = mileage = None
    for chip in card.select("ul.detail li"):
        text = " ".join(chip.get_text(" ", strip=True).split())
        lowered = text.lower()
        if lowered in FUEL:
            fuel = FUEL[lowered]
        elif _YEAR_ONLY.fullmatch(text):
            year = int(text)
        elif "km" in lowered:
            mileage = to_int(text)
    return fuel, year, mileage


def _french_date(text):
    """`29 septembre 2026` -> `2026-09-29`, else None."""
    match = re.search(r"(\d{1,2})\s+([a-zéûô]+)\s+(\d{4})", str(text or ""),
                      re.I)
    if not match:
        return None
    month = MONTHS.get(match.group(2).lower())
    if not month:
        return None
    return f"{int(match.group(3)):04d}-{month:02d}-{int(match.group(1)):02d}"


def parse_search(html):
    """One feed page -> list of records (14 ads per page, no announced total)."""
    doc = soup(html)
    records = []
    seen = set()
    for card in doc.find_all("li"):
        title_el = card.find("p", class_="titre")
        price_el = card.find("p", class_="prix")
        link = card.find("a", href=_CARD)
        if title_el is None or price_el is None or link is None:
            continue

        match = re.search(r"/(\d+)\.html$", link["href"])
        if not match:
            continue
        list_id = f"wd_{match.group(1)}"
        if list_id in seen:
            continue
        seen.add(list_id)

        fuel, year, mileage = _chips(card)
        city_el = card.select_one(".infos .city")
        date_el = card.select_one(".infos .dateHeure")
        photo = card.find("a", class_="img")
        image = photo.find("img") if photo else None

        price = money(price_el.get_text(" ", strip=True))
        if price is None:
            bare = price_el.find("span")
            price = to_int(bare.get_text(" ", strip=True)) if bare else None

        date_text = " ".join(date_el.get_text(" ", strip=True).split()) if date_el else None
        records.append(make_record(
            SOURCE, list_id, link["href"],
            title=" ".join(title_el.get_text(" ", strip=True).split()),
            price=price,
            price_source="annonce" if price else None,
            year=year,
            mileage_km=mileage,
            fuel=fuel,
            location=" ".join(city_el.get_text(" ", strip=True).split())
            if city_el else None,
            date_posted=date_text,
            date_posted_exact=_french_date(date_text),
            images=[image.get("src")] if image and image.get("src") else [],
        ))
    return records


def parse_detail(html):
    """Detail page -> brand/model, specs, seller name and phone."""
    doc = soup(html)
    main = doc.find("div", id="annonce-content") or doc
    out = {"params": {}, "labels": {}}

    for li in main.select("ul.icons.clearfix > li"):
        title = li.find("p", class_="titre")
        tag = li.find("p", class_="tag")
        if title is None or tag is None:
            continue
        label = " ".join(title.get_text(" ", strip=True).split())
        value = " ".join(tag.get_text(" ", strip=True).split())
        if not label or not value:
            continue

        target = FIELD_MAP.get(label.lower())
        if target == "year":
            out["year"] = first_year(value)
        elif target == "mileage_km":
            out["mileage_km"] = to_int(value)
        elif target == "seller_type":
            out["seller_type"] = _pick(SELLER_TYPES, value)
        elif target == "gearbox":
            out["gearbox"] = _pick(GEARBOX, value)
        elif target == "fuel":
            out["fuel"] = _pick(FUEL, value)
        elif target:
            out[target] = value
        else:
            key = slug_key(label)
            out["params"][key] = value
            out["labels"][key] = label

    # Breadcrumb: the marque and modèle filter links carry the two answers.
    # Every link repeats the marque, so the brand is the first non-zero one.
    brand = model = None
    breadcrumb = doc.find("div", id="breadcrumb")
    if breadcrumb:
        for anchor in breadcrumb.find_all("a", href=True):
            marca = re.search(r"[?&]marque=(\d+)", anchor["href"])
            if brand is None and marca and marca.group(1) != "0":
                brand = " ".join(anchor.get_text(" ", strip=True).split())
            modele = re.search(r"[?&]modele=(\d+)", anchor["href"])
            if modele and modele.group(1) != "0":
                model = " ".join(anchor.get_text(" ", strip=True).split())
    out["brand"], out["model"] = brand, model

    price_el = main.find("p", class_="prix")
    if price_el is None:
        price_el = doc.find("p", class_="prix")
    if price_el:
        out["price"] = money(price_el.get_text(" ", strip=True))

    description = main.find("p", class_="information")
    if description:
        text = " ".join(description.get_text(" ", strip=True).split())
        out["description"] = text or None

    seller = doc.find("div", id="vendeur")
    if seller:
        name = seller.find("p", class_="name")
        if name:
            text = " ".join(name.get_text(" ", strip=True).split())
            out["seller_name"] = text or None
        mobile = seller.find("p", class_="mobile")
        if mobile:
            digits = re.sub(r"\D", "", mobile.get_text(" ", strip=True))
            if len(digits) >= 9:
                out["seller_phone"] = digits
    return out


def enrich(record, details_dir, session):
    """Merge one listing's detail page into its record (cached per listing)."""
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
                "description", "seller_name", "seller_phone", "seller_type"):
        value = detail.get(key)
        if value not in (None, "", []):
            record[key] = value
    for key, label in (("brand", "Marque"), ("model", "Modèle")):
        value = detail.get(key)
        if value:
            params[key] = value
            labels.setdefault(key, label)
            record[key] = value

    if record.get("seller_type"):
        record["is_professional"] = record["seller_type"] == "STORE"
    if not record.get("price") and detail.get("price"):
        record["price"] = detail["price"]
        record["currency"] = record.get("currency") or "DH"
        record["price_source"] = "detail"
    return record


def crawl(output_dir, *, max_pages=0, with_details=True, session=None,
          refresh_top=0):
    """Crawl the Wandaloo occasion feed. Returns the merged record list.

    ``refresh_top`` re-fetches the first N feed pages (newest-first) so a
    scheduled run picks up new ads without re-walking the whole pagination.
    """
    root = Path(output_dir) / "sources" / "wandaloo"
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
        log.info("page %s: %s listings (unique %s)", page, len(parsed), len(unique))
        if len(unique) == before:
            log.info("no new listings at page %s, stopping", page)
            break
        page += 1

    records = list(unique.values())
    if with_details:
        records = [enrich(rec, details_dir, session) for rec in records]

    write_json(root / "annonces.json", records)
    log.info("wandaloo: %s listings", len(records))
    return records
