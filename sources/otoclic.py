"""otoclic.com — OTOCLIC's own used-car stock (Casa / Agadir / Fès showrooms).

robots.txt allows the crawl (``User-agent: *`` only blocks ``/wp-admin/``).
The stock archive is a server-rendered WordPress listing: 12 cars per page
behind ``/page/N/``, and every card already carries make, model, year,
mileage, fuel, gearbox, price and the monthly instalment through its own
taxonomy classes. The detail page is fetched anyway for the showroom, the
seller's name and the phone number printed on the page.

Cache layout (mirrors scraper.py):
    <output>/sources/otoclic/pages/page_0001.json   parsed stock cards
    <output>/sources/otoclic/details/<list_id>.json parsed detail specs
    <output>/sources/otoclic/annonces.json          merged export
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

from .common import (fetch, first_year, make_record, money, new_session,
                     read_json, slug_key, soup, to_int, write_json)

log = logging.getLogger("t480.otoclic")

BASE = "https://otoclic.com"
LIST_ROOT = BASE + "/acheter-voiture-occasion-maroc"
LIST_URLS = [LIST_ROOT + "/", LIST_ROOT + "/page/{page}/"]
SOURCE = "otoclic"
MAX_PAGES = 30

GEARBOX = {
    "auto": "Automatique",
    "automatique": "Automatique",
    "manuel": "Manuelle",
    "manuelle": "Manuelle",
}
FUEL = {
    "essence": "Essence",
    "diesel": "Diesel",
    "hybride": "Hybride",
    "electrique": "Électrique",
    "électrique": "Électrique",
    "gpl": "GPL",
}
SHOWROOM_CITIES = (
    ("casa", "Casablanca"),
    ("agadir", "Agadir"),
    ("fès", "Fès"),
    ("fes", "Fès"),
    ("marrakech", "Marrakech"),
    ("rabat", "Rabat"),
    ("tanger", "Tanger"),
)

# Detail-page label -> the field name scraper.py already uses (avito parity).
FIELD_MAP = {
    "marque": "brand",
    "modéle": "model",
    "modèle": "model",
    "modele": "model",
    "année": "year",
    "annee": "year",
    "kilométrage": "mileage_km",
    "kilometrage": "mileage_km",
    "boite à vitesse": "gearbox",
    "boîte à vitesse": "gearbox",
    "energie": "fuel",
    "énergie": "fuel",
    "localisation": "location",
    "reference": ("params", "reference"),
}


def _pick(mapping, value):
    if not value:
        return None
    return mapping.get(str(value).strip().lower(), str(value).strip())


def _titlecase(slug):
    if not slug:
        return None
    return slug.replace("-", " ").title()


def _taxonomy(classes, prefix):
    for name in classes:
        if name.startswith(prefix):
            return name[len(prefix):]
    return None


def _city(text):
    """`Showroom Casa Siége` -> `Casablanca`; anything else passes through."""
    if not text:
        return None
    lowered = text.lower()
    for needle, label in SHOWROOM_CITIES:
        if needle in lowered:
            return label
    return text


def parse_search(html):
    """One stock page -> list of records (12 cars per page)."""
    doc = soup(html)
    records = []
    seen = set()
    for card in doc.select(".listing-item"):
        classes = card.get("class") or []
        post = _taxonomy(classes, "post-")
        link_el = card.find("a", class_="listing-image") or card.find("a", href=True)
        href = link_el.get("href") if link_el else None
        if not href:
            continue
        list_id = f"oto_{post}" if post else f"oto_{href.rstrip('/').rsplit('/', 1)[-1]}"
        if list_id in seen:
            continue
        seen.add(list_id)

        title_el = card.select_one(".listing-title")
        price_el = card.select_one(".listing-price .price-text")
        month_el = card.select_one(".mensualite-price")
        year_el = card.select_one(".listing-meta.year .value-suffix")
        km_el = card.select_one(".listing-meta.mileage .value-suffix")
        fuel_el = card.select_one(".listing-tax.fuel_type .value-suffix")
        gear_el = card.select_one(".listing-tax.transmission .value-suffix")
        image = card.find("a", class_="listing-image")
        image = image.find("img") if image else card.find("img")

        brand = _titlecase(_taxonomy(classes, "listing_make-"))
        title = " ".join(title_el.get_text(" ", strip=True).split()) if title_el else None
        model = _titlecase(_taxonomy(classes, "listing_model-"))
        if brand and title and title.lower().startswith(brand.lower()):
            tail = title[len(brand):].strip()
            if tail:
                model = tail

        price = to_int(price_el.get_text(" ", strip=True)) if price_el else None
        records.append(make_record(
            SOURCE, list_id, href,
            title=title,
            brand=brand,
            model=model,
            year=first_year(year_el) if year_el else None,
            price=price,
            price_source="annonce" if price else None,
            monthly_payment=money(month_el.get_text(" ", strip=True))
            if month_el else None,
            mileage_km=to_int(km_el.get_text(" ", strip=True)) if km_el else None,
            fuel=_pick(FUEL, fuel_el.get_text(" ", strip=True)) if fuel_el else None,
            gearbox=_pick(GEARBOX, gear_el.get_text(" ", strip=True))
            if gear_el else None,
            seller_type="STORE",
            is_professional=True,
            images=[image.get("src")] if image and image.get("src") else [],
        ))
    return records


def parse_detail(html):
    """Stock page -> specs, showroom, seller name and phone."""
    doc = soup(html)
    out = {"params": {}, "labels": {}}

    for li in doc.select("li.meta-overview"):
        label_el = li.select_one(".field-title")
        value_el = li.select_one(".content-value")
        if label_el is None or value_el is None:
            continue
        label = " ".join(label_el.get_text(" ", strip=True).split()).rstrip(": ")
        value = " ".join(value_el.get_text(" ", strip=True).split())
        if not label or not value:
            continue

        target = FIELD_MAP.get(label.lower())
        if target == "year":
            out["year"] = first_year(value)
        elif target == "mileage_km":
            out["mileage_km"] = to_int(value)
        elif target == "gearbox":
            out["gearbox"] = _pick(GEARBOX, value)
        elif target == "fuel":
            out["fuel"] = _pick(FUEL, value)
        elif target == "location":
            out["location"] = _city(value)
            out["params"]["showroom"] = value
            out["labels"]["showroom"] = "Showroom"
        elif isinstance(target, tuple) and target[0] == "params":
            out["params"][target[1]] = value
            out["labels"][target[1]] = label
        elif target in ("brand", "model"):
            out[target] = value.title() if target == "brand" else value
        elif target:
            out[target] = value

    seller = doc.select_one(".listing-detail-author-info .title-user")
    if seller:
        text = " ".join(seller.get_text(" ", strip=True).split())
        out["seller_name"] = text or None

    phone = doc.select_one(".agent-phone .phone") or doc.select_one("a[href^='tel:']")
    if phone:
        digits = re.sub(r"\D", "", phone.get_text(" ", strip=True) or
                        (phone.get("href") or ""))
        if len(digits) >= 9:
            out["seller_phone"] = digits

    price = doc.select_one(".listing-price .price-text")
    if price:
        out["price"] = to_int(price.get_text(" ", strip=True))
    return out


def enrich(record, details_dir, session):
    """Merge one car's detail page into its record (cached per listing)."""
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
                "seller_name", "seller_phone"):
        value = detail.get(key)
        if value not in (None, "", []):
            record[key] = value
    for key, label in (("brand", "Marque"), ("model", "Modèle")):
        value = detail.get(key)
        if value:
            params[key] = value
            labels.setdefault(key, label)
            record[key] = value

    if not record.get("price") and detail.get("price"):
        record["price"] = detail["price"]
        record["currency"] = record.get("currency") or "DH"
        record["price_source"] = "detail"
    return record


def crawl(output_dir, *, max_pages=0, with_details=True, session=None,
          refresh_top=0):
    """Crawl the stock archive. Returns the merged record list.

    ``refresh_top`` re-fetches the first N stock pages (newest-first) so a
    scheduled run picks up new arrivals without re-walking the whole archive.
    """
    root = Path(output_dir) / "sources" / "otoclic"
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
            url = LIST_URLS[0] if page == 1 else LIST_URLS[1].format(page=page)
            html = fetch(session, url)
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
        log.info("page %s: %s cars (unique %s)", page, len(parsed), len(unique))
        if len(unique) == before:
            log.info("no new cars at page %s, stopping", page)
            break
        page += 1

    records = list(unique.values())
    if with_details:
        records = [enrich(rec, details_dir, session) for rec in records]

    write_json(root / "annonces.json", records)
    log.info("otoclic: %s listings", len(records))
    return records
