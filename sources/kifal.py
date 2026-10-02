"""occasion.kifal.ma — Kifal Auto's used-car marketplace (Morocco).

robots.txt allows crawling (``Disallow:`` is empty). The search pages are
server-rendered Bootstrap cards; the detail page adds a clean label/value
specification table. Seller phone numbers are masked by the site, so they are
recorded as hidden rather than guessed.

Cache layout (mirrors scraper.py):
    <output>/sources/kifal/pages/page_0001.json   parsed search cards
    <output>/sources/kifal/details/<list_id>.json parsed detail specs
    <output>/sources/kifal/annonces.json          merged export
"""

from __future__ import annotations

import logging
import re
from pathlib import Path

from .common import (first_year, fetch, label_for, make_record, money, new_session,
                     read_json, slug_key, soup, to_int, write_json)

log = logging.getLogger("t480.kifal")

BASE = "https://occasion.kifal.ma"
LIST_URL = BASE + "/annonces?page={page}"
SOURCE = "kifal"
MAX_PAGES = 150      # the feed is ~40 pages deep; this is a safety net

# Detail-page label -> the field name scraper.py already uses (avito parity).
FIELD_MAP = {
    "ville": ("location", None),
    "marque": ("brand", None),
    "modèle": ("model", None),
    "modele": ("model", None),
    "la finition": ("params", "trim"),
    "carburant": ("fuel", None),
    "transmission": ("gearbox", None),
    "année": ("year", None),
    "annee": ("year", None),
    "kilométrage": ("mileage_km", None),
    "kilometrage": ("mileage_km", None),
    "puissance fiscale": ("params", "pfiscale"),
    "origine": ("params", "v_origin"),
    "première main": ("params", "first_owner"),
    "premiere main": ("params", "first_owner"),
    "type de voiture": ("params", "v_type"),
    "état": ("params", "auto_condition"),
    "etat": ("params", "auto_condition"),
    "couleur": ("params", "color"),
}


def _datalayer(text):
    """`pushDataLayer({..., product_brand:"BMW", product_model:"X5"})` -> dict."""
    out = {}
    for key in ("product_brand", "product_model", "product_year"):
        match = re.search(rf'{key}\s*:\s*"([^"]*)"', text or "")
        if match and match.group(1):
            out[key] = match.group(1)
    return out


def _icon_text(card, icon_class, root=None):
    """Text of the chip that carries a given FontAwesome icon.

    Kifal renders the same footer twice (a featured strip and the main list)
    with different wrapper classes, so the icon itself is the only stable
    anchor: take the text of the icon's parent chip.
    """
    scope = (root if root is not None else card)
    icon = scope.find("i", class_=lambda c: c and icon_class in c)
    if icon is None:
        return None
    chip = icon.parent
    text = " ".join(chip.get_text(" ", strip=True).split())
    return text or None


def _price(anchor) -> int | None:
    """Price from the price anchor: digits plus the trailing `Dh` word."""
    if anchor is None:
        return None
    amount = money(anchor.get_text(" ", strip=True))
    if amount is not None:
        return amount
    # Some cards write the amount with no currency word; the anchor only ever
    # wraps the price, so its digits are safe to use.
    digits = anchor.find("b")
    return to_int(digits.get_text(" ", strip=True)) if digits else None


def parse_search(html):
    """One search page -> (records, announced total or None)."""
    doc = soup(html)
    total = None
    counter = doc.select_one(".total-card-annonces")
    if counter:
        total = to_int(counter.get_text(" ", strip=True))

    records = []
    for card in doc.select("div.card-annonce"):
        data_url = (card.get("data-url") or "").strip()
        link = card.select_one('a[href*="/annonce/"]')
        href = (link.get("href") if link else None) or (
            f"{BASE}/annonce/{data_url}" if data_url else None)
        if not href:
            continue

        list_id = data_url[:-4] if data_url.endswith(".htm") else data_url
        list_id = list_id or href.rsplit("/", 1)[-1]

        onclick = " ".join(a.get("onclick") or "" for a in card.select("a[onclick]"))
        meta = _datalayer(onclick)

        title_el = card.select_one(".title-card-full-annonce")
        title = " ".join(title_el.get_text(" ", strip=True).split()) if title_el else None

        year = first_year(meta.get("product_year")) or first_year(
            _icon_text(card, "fa-calendar") or "")

        city = None
        for span in card.select("span.text-muted-dark"):
            if span.find("i", class_=lambda c: c and "fa-map-marker" in c):
                city = " ".join(span.get_text(" ", strip=True).split())
                break

        trim_el = card.select_one("div.leading-tight")
        trim = " ".join(trim_el.get_text(" ", strip=True).split()) if trim_el else None

        footer = card.select_one(".card-footer") or card
        fuel = _icon_text(card, "fa-gas-pump", root=footer)
        gearbox = _icon_text(card, "fa-cogs", root=footer)
        mileage = to_int(_icon_text(card, "fa-road", root=footer))

        images = [img.get("src") for img in card.select("img.cover-image")
                  if img.get("src")]

        energy = card.select_one(".blockHead .blocktext")
        price = _price(card.select_one("a.price-font-size"))

        record = make_record(
            SOURCE, list_id, href,
            title=title,
            brand=meta.get("product_brand"),
            model=meta.get("product_model"),
            year=year,
            price=price,
            price_source="annonce" if price else None,
            monthly_payment=_price(card.select_one("a.price-font-size-credit")),
            mileage_km=mileage,
            fuel=fuel,
            gearbox=gearbox,
            location=city,
            images=images,
            params={"trim": trim, "energy_class": energy.get_text(strip=True)
                    if energy else None},
            detail_labels={"trim": label_for("trim"), "energy_class": "Étiquette énergie"},
        )
        record["other_params"] = {k: v for k, v in record["other_params"].items()
                                  if v not in (None, "", [])}
        records.append(record)
    return records, total


def parse_detail(html):
    """Detail page -> dict of spec fields plus seller info."""
    doc = soup(html)
    out = {"params": {}, "labels": {}}

    for cell in doc.select("table td"):
        spans = cell.find_all("span")
        if len(spans) < 2:
            continue
        label = " ".join(spans[0].get_text(" ", strip=True).split()).rstrip(": ")
        value = " ".join(spans[1].get_text(" ", strip=True).split())
        if not label or not value:
            continue
        target, param_key = FIELD_MAP.get(label.lower(), (None, slug_key(label)))
        if target == "params" and param_key:
            out["params"][param_key] = value
            out["labels"][param_key] = label
        elif target in ("year", "mileage_km"):
            out[target] = first_year(value) if target == "year" else to_int(value)
        elif target:
            out[target] = value

    seller = doc.select_one(".profile-pic") or doc.select_one(".item-user")
    if seller:
        out["seller_name"] = " ".join(seller.get_text(" ", strip=True).split())

    # The site only reveals the last digits; do not fabricate the rest.
    masked = doc.select_one("div.click-phone-annonce")
    if masked:
        out["phone_hidden"] = True

    meta = doc.find("meta", attrs={"name": "description"})
    if meta and meta.get("content"):
        text = re.sub(r"^description_detail_annonce_start", "", meta["content"])
        out["description"] = re.sub(r"description_detail_annonce_end$", "",
                                    text).strip() or None

    return out


def crawl(output_dir, *, max_pages=0, with_details=True, session=None,
          detail_delay=1.5, refresh_top=0):
    """Crawl the Kifal listings. Returns the merged record list.

    ``refresh_top`` re-fetches the first N search pages (the feed is
    newest-first) so a scheduled run picks up new ads without re-walking
    the whole pagination.
    """
    root = Path(output_dir) / "sources" / "kifal"
    pages_dir = root / "pages"
    details_dir = root / "details"
    pages_dir.mkdir(parents=True, exist_ok=True)
    details_dir.mkdir(parents=True, exist_ok=True)

    session = session or new_session()
    unique: dict[str, dict] = {}
    page = 1
    announced_total = None
    limit = max_pages or MAX_PAGES
    stale = 0

    while page <= limit:
        cache = pages_dir / f"page_{page:04d}.json"
        cached = None if (refresh_top and page <= refresh_top) else read_json(cache)
        if cached is None:
            html = fetch(session, LIST_URL.format(page=page))
            if html is None:
                log.warning("page %s could not be fetched", page)
                break
            parsed, total = parse_search(html)
            write_json(cache, {"total": total, "listings": parsed})
        else:
            if isinstance(cached, dict):
                parsed = cached.get("listings") or []
                total = cached.get("total")
            else:
                parsed = cached
                total = None
        if total and not announced_total:
            announced_total = total

        if not parsed:
            log.info("end of pagination at page %s", page)
            break

        before = len(unique)
        for rec in parsed:
            unique.setdefault(rec["list_id"], rec)
        log.info("page %s: %s listings (unique %s / total %s)", page,
                 len(parsed), len(unique), announced_total or "?")
        if len(unique) == before:
            # The site serves pages long after the feed ends (the same ads
            # over and over, while its own counter claims more exist), so
            # two dead pages in a row mean there is nothing left to collect.
            stale += 1
            if stale >= 2:
                log.info("no new listings at page %s, stopping", page)
                break
        else:
            stale = 0
        if announced_total and len(unique) >= announced_total:
            break
        page += 1

    records = list(unique.values())
    if with_details:
        records = [enrich(rec, details_dir, session, detail_delay) for rec in records]

    write_json(root / "annonces.json", records)
    log.info("kifal: %s listings", len(records))
    return records


def enrich(record, details_dir, session, delay=1.5):
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
    for key in ("year", "mileage_km", "fuel", "gearbox", "location", "description"):
        value = detail.get(key)
        if value not in (None, "", []):
            record[key] = value
    if detail.get("seller_name"):
        record["seller_name"] = detail["seller_name"]
        record["is_professional"] = True
    if detail.get("phone_hidden"):
        record["phone_hidden"] = True
    return record
