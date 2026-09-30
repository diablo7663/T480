"""Shared plumbing for the Moroccan car-marketplace sources.

Every source module in this package fetches one site and emits records in
exactly the shape scraper.py already writes for avito.ma (same top-level keys,
same other_params / detail_labels convention). merge.py can then union every
source and reuse the cleaning and Excel export untouched.

Everything here is deliberately polite: one request per host every
MIN_INTERVAL seconds, Retry-After-aware backoff, no thread pool.
"""

from __future__ import annotations

import json
import logging
import re
import threading
import time
from pathlib import Path
from urllib.parse import urlsplit

import requests
from bs4 import BeautifulSoup

log = logging.getLogger("t480.sources")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)

MIN_INTERVAL = 1.5      # seconds between two hits on the same host
RETRY_STATUS = {429, 500, 502, 503, 504}

_last_hit: dict[str, float] = {}
_lock = threading.Lock()


def new_session() -> requests.Session:
    """A session that looks like a normal browser and speaks French."""
    session = requests.Session()
    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept-Language": "fr-FR,fr;q=0.9,ar;q=0.8,en;q=0.6",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    })
    return session


def _throttle(host: str) -> None:
    """Keep at least MIN_INTERVAL seconds between two hits on the same host."""
    with _lock:
        last = _last_hit.get(host)
        if last is not None:
            wait = MIN_INTERVAL - (time.monotonic() - last)
            if wait > 0:
                time.sleep(wait)
        _last_hit[host] = time.monotonic()


def fetch(session, url, *, timeout=30.0, retries=5):
    """GET with per-host politeness and Retry-After-aware backoff.

    Returns the body as text, or None when the page never comes back (4xx
    other than 429 are treated as final: the page genuinely isn't there).
    """
    host = urlsplit(url).netloc
    last_error = None
    for attempt in range(retries):
        _throttle(host)
        try:
            response = session.get(url, timeout=timeout)
        except requests.RequestException as exc:
            last_error = exc
            time.sleep(min(2 ** attempt, 30))
            continue
        if response.status_code == 200:
            return response.text
        if response.status_code in RETRY_STATUS:
            retry_after = str(response.headers.get("Retry-After") or "").strip()
            delay = float(retry_after) if retry_after.isdigit() else float(2 ** attempt)
            delay = min(delay, 120)
            log.info("%s -> HTTP %s, retrying in %.0fs", host, response.status_code, delay)
            time.sleep(delay)
            last_error = f"HTTP {response.status_code}"
            continue
        log.warning("HTTP %s for %s", response.status_code, url)
        return None
    log.error("giving up on %s (%s)", url, last_error)
    return None


def soup(html) -> BeautifulSoup:
    return BeautifulSoup(html or "", "html.parser")


# --------------------------------------------------------------------------
# Text parsers
# --------------------------------------------------------------------------

_NO_DIGITS = re.compile(r"[^\d]")
_MONEY = re.compile(
    r"(\d[\d\s .,]{1,16})\s*(?:dh|dhs|mad|dirhams?)\b", re.I)
_YEAR = re.compile(r"\b(19[89]\d|20[0-3]\d)\b")


def to_int(value) -> int | None:
    """`222 091` / `222091 km` / `8` -> int. Only for fields that are numbers."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return int(value)
    digits = _NO_DIGITS.sub("", str(value or ""))
    return int(digits) if digits else None


def money(text) -> int | None:
    """`242 000 Dh`, `1 385 000 MAD`, `27 900 DH` -> int, else None.

    Requires a currency word so a year or a mileage is never mistaken for a
    price.
    """
    if isinstance(text, bool):
        return None
    if isinstance(text, (int, float)):
        return int(text)
    match = _MONEY.search(str(text or ""))
    if not match:
        return None
    digits = _NO_DIGITS.sub("", match.group(1))
    return int(digits) if digits else None


def first_year(text) -> int | None:
    """First plausible registration year in free text (`2016-09-30` -> 2016)."""
    match = _YEAR.search(str(text or ""))
    return int(match.group(1)) if match else None


def slug_key(label: str) -> str:
    """`Puissance fiscale` -> `puissance_fiscale`, for ad-hoc detail params."""
    key = re.sub(r"[^0-9a-zA-Z]+", "_", str(label or "")).strip("_").lower()
    return key or "param"


# French labels for the params sources add that avito.ma does not use.
PARAM_LABELS = {
    "brand": "Marque",
    "model": "Modèle",
    "trim": "Finition",
    "v_type": "Type de voiture",
    "v_origin": "Origine",
    "first_owner": "Première main",
    "auto_condition": "État",
    "pfiscale": "Puissance fiscale (CV)",
    "doors": "Nombre de portes",
    "color": "Couleur",
    "energy_class": "Étiquette énergie",
    "insurance_valid": "Assurance à jour",
    "technical_inspection": "Visite technique",
    "warranty": "Garantie",
    "owner_count": "Nombre de propriétaires",
    "emission": "Émissions CO2",
}


def label_for(key: str, fallback: str | None = None) -> str:
    """French header for a detail param, preferring what avito already calls it."""
    return PARAM_LABELS.get(key) or fallback or key


# --------------------------------------------------------------------------
# Record factory
# --------------------------------------------------------------------------

def make_record(source: str, list_id: str, url: str, **fields):
    """Build one row shaped exactly like scraper.py's avito rows.

    Anything not supplied keeps the same key as avito so the union of all
    sources has a single, stable column set.
    """
    params = dict(fields.pop("params", None) or {})
    brand = fields.pop("brand", None)
    model = fields.pop("model", None)
    if brand:
        params.setdefault("brand", brand)
    if model:
        params.setdefault("model", model)

    images = fields.pop("images", None) or []
    price = fields.get("price")

    labels = {key: label_for(key, label)
              for key, label in (fields.pop("detail_labels", None) or {}).items()}
    for key in params:
        labels.setdefault(key, label_for(key))

    record = {
        "id": list_id,
        "list_id": list_id,
        "url": url,
        "title": None,
        "description": None,
        "category": "Voitures - Voitures d'occasion",
        "ad_type": "à vendre",
        "price": None,
        "currency": None,
        "monthly_payment": None,
        "old_price": None,
        "year": None,
        "mileage_km": None,
        "fuel": None,
        "gearbox": None,
        "other_params": params,
        "location": None,
        "city_id": None,
        "area_id": None,
        "date_posted": None,
        "seller_id": None,
        "seller_type": None,
        "seller_name": None,
        "seller_phone": None,
        "seller_phone_verified": False,
        "seller_verified": False,
        "is_professional": None,
        "is_premium": None,
        "is_urgent": False,
        "is_hot_deal": False,
        "is_shop": None,
        "is_car_checked": None,
        "is_delivery": False,
        "is_highlighted": False,
        "is_immoneuf": None,
        "discount": None,
        "has_shipping": False,
        "is_ecommerce": False,
        "default_image": images[0] if images else None,
        "image_count": len(images),
        "images": images,
        "price_source": None,
        "date_posted_exact": None,
        "phone_hidden": False,
        "phone_verified": False,
        "seller_address": None,
        "seller_badges": None,
        "seller_listings": None,
        "detail_labels": labels,
        "source": source,
    }
    record.update(fields)
    if price:
        record["currency"] = record["currency"] or "DH"
    if brand:
        record["brand"] = params["brand"]
    return record


# --------------------------------------------------------------------------
# Small cache helpers (one JSON file per page, like scraper.py)
# --------------------------------------------------------------------------

def write_json(path, payload) -> None:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=1),
                    encoding="utf-8")


def read_json(path, default=None):
    path = Path(path)
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError) as exc:
        log.warning("could not read %s (%s)", path, exc)
        return default
