"""www.oneclickdrive.ma — OneClickDrive's Moroccan used-car listings.

robots.txt permits this crawl: the blanket ``Disallow: /*?`` is overridden by
the longer ``Allow: /*?page=`` and ``Allow: /*?id=`` rules, so search
pagination and detail links are explicitly allowed. Listings are
server-rendered as an ItemList inside JSON-LD (20 cars per page), one paginated
page per city, and the city list comes from ``/sitemap/morocco.xml``.

Only the search pages are fetched: they already carry the structured car data
(price, brand, model, year, mileage, fuel, gearbox, body, colour, photo,
description). Seller contact details are not published on these pages, so the
seller columns stay empty rather than guessed.

Cache layout (mirrors scraper.py):
    <output>/sources/oneclickdrive/pages/<city>_p0001.json
    <output>/sources/oneclickdrive/annonces.json
"""

from __future__ import annotations

import json
import logging
import re
from pathlib import Path

from .common import (fetch, first_year, make_record, new_session, read_json,
                     to_int, write_json)

log = logging.getLogger("t480.oneclickdrive")

BASE = "https://www.oneclickdrive.ma"
LIST_URL = BASE + "/buy-used-cars-{city}?page={page}"
SOURCE = "oneclickdrive"

# City slug (as used in the site's own URLs) -> French label.
CITIES = {
    "casablanca": "Casablanca",
    "agadir": "Agadir",
    "fes": "Fès",
    "marrakech": "Marrakech",
    "nador": "Nador",
    "oujda": "Oujda",
    "rabat": "Rabat",
    "tangier": "Tanger",
}

FUEL = {
    "petrol": "Essence",
    "diesel": "Diesel",
    "hybrid": "Hybride",
    "electric": "Électrique",
    "lpg": "GPL",
    "other": "Autre",
}
GEARBOX = {
    "auto": "Automatique",
    "automatic": "Automatique",
    "manual": "Manuelle",
    "manuelle": "Manuelle",
}

_PAGER = re.compile(
    r'showing-listings">\s*[\d\s]*-\s*[\d\s]+\s*of\s*(?:<[^>]*>\s*)*(\d[\d\s]*)')


def _pick(mapping, value):
    if not value:
        return None
    return mapping.get(str(value).strip().lower(), str(value).strip())


def _walk(node, out):
    if isinstance(node, dict):
        if node.get("@type") == "ItemList":
            for entry in node.get("itemListElement") or []:
                item = entry.get("item") if isinstance(entry, dict) else None
                if isinstance(item, dict):
                    out.append(item)
        for value in node.values():
            _walk(value, out)
    elif isinstance(node, list):
        for value in node:
            _walk(value, out)


def _items(html):
    """Every car embedded in the page's JSON-LD ItemList blocks."""
    out = []
    for raw in re.findall(r'<script type="application/ld\+json">(.*?)</script>',
                          html or "", re.S):
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            continue
        _walk(data, out)
    return out


def _total(html):
    match = _PAGER.search(html or "")
    return to_int(match.group(1)) if match else None


def parse_search(html, city):
    """One search page -> (records, announced total for the city or None)."""
    city_label = CITIES.get(city, city.replace("-", " ").title())
    records = []
    for item in _items(html):
        url = item.get("@id") or (item.get("offers") or {}).get("url")
        if not url:
            continue
        match = re.search(r"[?&]id=(\d+)", url)
        if match:
            list_id = f"ocd_{match.group(1)}"
        else:
            slug = re.sub(r"/+$", "", url).rsplit("/", 1)[-1]
            list_id = f"ocd_{slug}"

        offer = item.get("offers") or {}
        odometer = item.get("mileageFromOdometer") or {}
        images = item.get("image") or []
        if isinstance(images, str):
            images = [images]

        price = to_int(offer.get("price"))
        brand = (item.get("brand") or {}).get("name")
        model = item.get("model")
        color = (item.get("color") or "").strip() or None
        body = (item.get("bodyType") or "").strip() or None

        params = {"color": color} if color else {}
        labels = {}
        if body:
            params["body_type"] = body
            labels["body_type"] = "Carrosserie"

        records.append(make_record(
            SOURCE, list_id, url,
            title=item.get("name"),
            brand=brand,
            model=model,
            year=first_year(item.get("productionDate")),
            price=price,
            price_source="annonce" if price else None,
            mileage_km=to_int(odometer.get("value")),
            fuel=_pick(FUEL, item.get("fuelType")),
            gearbox=_pick(GEARBOX, item.get("vehicleTransmission")),
            location=city_label,
            description=(item.get("description") or "").strip() or None,
            images=[src for src in images if src],
            params=params,
            detail_labels=labels,
        ))
    return records, _total(html)


def crawl(output_dir, *, max_pages=0, session=None, refresh_top=0):
    """Crawl every city's listings. Returns the merged record list.

    ``refresh_top`` re-fetches the first N pages of every city (newest-first
    feed) so a scheduled run picks up new ads without re-walking each city.
    """
    root = Path(output_dir) / "sources" / "oneclickdrive"
    pages_dir = root / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)

    session = session or new_session()
    unique: dict[str, dict] = {}

    for city in CITIES:
        page = 1
        announced = None
        city_seen: dict[str, dict] = {}
        while True:
            if max_pages and page > max_pages:
                break
            cache = pages_dir / f"{city}_p{page:04d}.json"
            cached = (None if (refresh_top and page <= refresh_top)
                      else read_json(cache))
            if cached is None:
                html = fetch(session, LIST_URL.format(city=city, page=page))
                if html is None:
                    log.warning("%s page %s could not be fetched", city, page)
                    break
                parsed, total = parse_search(html, city)
                write_json(cache, {"total": total, "listings": parsed})
            else:
                parsed = cached.get("listings") or []
                total = cached.get("total")
            if total and not announced:
                announced = total

            if not parsed:
                log.info("%s: end of pagination at page %s", city, page)
                break

            before = len(city_seen)
            for rec in parsed:
                city_seen.setdefault(rec["list_id"], rec)
            log.info("%s page %s: %s listings (city unique %s / total %s)",
                     city, page, len(parsed), len(city_seen),
                     announced or "?")
            if len(city_seen) == before:
                log.info("%s: no new listings, stopping", city)
                break
            if announced and len(city_seen) >= announced:
                break
            page += 1
        unique.update(city_seen)

    records = list(unique.values())
    write_json(root / "annonces.json", records)
    log.info("oneclickdrive: %s listings", len(records))
    return records
