#!/usr/bin/env python3
"""Resumable scraper for the car listings database on avito.ma.

Extracts structured data from the __NEXT_DATA__ JSON embedded in each
search-result page, stores one JSON file per fetched page under an output
directory, and can consolidate everything into a single CSV + JSON export.

Usage:
    python3 scraper.py                         # crawl everything
    python3 scraper.py --max-pages 10 --delay 1
    python3 scraper.py --export-only           # re-export from cached pages
    python3 scraper.py --url 'https://www.avito.ma/fr/maroc/voitures_neuves'
"""

import argparse
import csv
import json
import logging
import math
import random
import re
import statistics
import threading
import time
import unicodedata
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import requests
from bs4 import BeautifulSoup

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    HAS_OPENPYXL = True
except ImportError:  # pragma: no cover
    HAS_OPENPYXL = False

log = logging.getLogger("avito-scraper")

DEFAULT_URL = "https://www.avito.ma/fr/maroc/voitures_a_vendre"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    ),
    "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.5,ar;q=0.4",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

CSV_COLUMNS = [
    "id", "list_id", "url", "title", "source", "description", "category",
    "ad_type",
    "price", "currency", "monthly_payment", "old_price", "year", "mileage_km",
    "fuel", "gearbox", "other_params_json", "location", "city_id", "area_id",
    "date_posted", "seller_id", "seller_type", "seller_name", "seller_phone",
    "seller_phone_verified", "seller_verified", "is_professional", "is_premium",
    "is_urgent", "is_hot_deal", "is_shop", "is_car_checked", "is_delivery",
    "is_highlighted", "is_immoneuf", "discount", "has_shipping", "is_ecommerce",
    "default_image", "image_count", "images_json", "price_source", "merged_ads",
]

# Detail-page params already exposed by the search page under their own column.
DETAIL_SKIP_KEYS = {"regdate", "mileage_exact", "bv", "fuel"}

# Detail-page fields flattened next to the base columns (order = sheet order).
DETAIL_EXTRA_FIELDS = [
    "date_posted_exact", "video_count", "phone_hidden", "phone_verified",
    "seller_address", "seller_website", "seller_badges", "seller_listings",
]

# Preferred order for the per-car detail columns.
DETAIL_PREFERRED_ORDER = [
    "brand", "model", "pfiscale", "doors", "v_origin", "auto_condition",
    "first_owner", "sector",
]

# Fields that must always be resolvable, even when a detail page is missing.
ESSENTIAL_FIELDS = ["brand", "model", "year", "price"]

# Titles that start with a bare model name need the marque added back, so keep
# a lookup of the marques actually listed on avito.ma (accent-free spelling).
BRAND_KEYWORDS = [
    "abarth", "alfa romeo", "aston martin", "audi", "bentley", "bmw",
    "bugatti", "cadillac", "chevrolet", "chrysler", "citroen", "citroën",
    "cupra", "dacia", "daewoo", "daihatsu", "datsun", "dodge", "ds",
    "fiat", "ford", "genesis", "honda", "hummer", "hyundai", "infiniti",
    "isuzu", "jaguar", "jeep", "kia", "lada", "lancia", "land rover",
    "landrover", "lexus", "maserati", "mazda", "mclaren", "mercedes",
    "mg", "mitsubishi", "nissan", "opel", "peugeot", "porsche", "ram",
    "renault", "rolls royce", "rover", "seat", "skoda", "smart", "ssangyong",
    "subaru", "suzuki", "tesla", "toyota", "volkswagen", "volvo",
]

# Marques whose canonical spelling is not just the title-cased keyword.
BRAND_DISPLAY = {
    "bmw": "BMW", "ds": "DS", "mg": "MG", "ram": "RAM", "vw": "Volkswagen",
}

# Model names made of two words, so "Range Rover" is not cut down to "Range".
MODEL_PHRASES = ["range rover", "land cruiser", "al mercedes", "series 3"]

# Roughly 4 800 dealers hide the price behind "prix à discuter". Most of those
# ads still spell the number out in the description, so recover it from there.
PRICE_MIN = 5_000
PRICE_MAX = 20_000_000
_PRICE_CURRENCY = re.compile(
    r"(\d[\d\s .,]{2,14})\s*(?:dh|dhmd|dhs|mad|dirhams?)\b", re.I)
_PRICE_WORDED = re.compile(
    r"(?:prix|prix[st]*\s*:|vendu|vendre|le\s+vend)\D{0,15}?"
    r"(\d[\d\s .,]{2,14})", re.I)


def _clean_amount(match):
    digits = re.sub(r"[^\d]", "", match)
    return int(digits) if digits else None


def price_from_text(text):
    """Pull a price out of free text, ignoring years / mileages / phone numbers."""
    if not text:
        return None
    for pattern in (_PRICE_CURRENCY, _PRICE_WORDED):
        for match in pattern.finditer(text):
            amount = _clean_amount(match.group(1))
            if amount is None or not PRICE_MIN <= amount <= PRICE_MAX:
                continue
            if 1900 <= amount <= 2035:  # a year, not a price
                continue
            return amount
    return None



def numeric(value):
    if isinstance(value, (int, float)):
        return value
    digits = re.sub(r"[^0-9]", "", str(value or ""))
    return int(digits) if digits else None


def secondary_params(ad):
    params = {}
    for group in (ad.get("params") or {}).values():
        if not isinstance(group, list):
            continue
        for item in group:
            if isinstance(item, dict) and item.get("key"):
                params[item["key"]] = item
    return params


def extract_ad(ad):
    seller = ad.get("seller") or {}
    phone = seller.get("phone") or {}
    price = ad.get("price") or {}
    monthly = ad.get("monthlyPayment") or {}
    old = ad.get("oldPrice") or {}
    category = ad.get("category") or {}
    secondary = secondary_params(ad)

    known = {"regdate", "mileage_exact", "bv", "fuel"}
    other_params = {
        key: item.get("fullValue", item.get("value"))
        for key, item in secondary.items()
        if key not in known
    }

    def sv(key):
        item = secondary.get(key)
        if item:
            return item.get("fullValue", item.get("value"))
        return None

    images = ad.get("images") or []

    return {
        "id": ad.get("id"),
        "list_id": ad.get("listId"),
        "url": ad.get("href"),
        "title": ad.get("subject"),
        "description": ad.get("description"),
        "category": category.get("formatted"),
        "ad_type": (ad.get("adType") or {}).get("label"),
        "price": price.get("value"),
        "currency": price.get("currency"),
        "monthly_payment": monthly.get("value"),
        "old_price": old.get("value"),
        "year": numeric(sv("regdate")),
        "mileage_km": numeric(sv("mileage_exact")),
        "fuel": sv("fuel"),
        "gearbox": sv("bv"),
        "other_params": other_params,
        "location": ad.get("location"),
        "city_id": ad.get("cityId"),
        "area_id": ad.get("areaId"),
        "date_posted": ad.get("date"),
        "seller_id": seller.get("id"),
        "seller_type": seller.get("type"),
        "seller_name": seller.get("name"),
        "seller_phone": phone.get("number"),
        "seller_phone_verified": phone.get("verified"),
        "seller_verified": seller.get("isVerifiedSeller"),
        "is_professional": seller.get("type") == "STORE",
        "is_premium": ad.get("isPremium"),
        "is_urgent": ad.get("isUrgent"),
        "is_hot_deal": ad.get("isHotDeal"),
        "is_shop": ad.get("isShop"),
        "is_car_checked": ad.get("isCarChecked"),
        "is_delivery": ad.get("isDelivery"),
        "is_highlighted": ad.get("isHighlighted"),
        "is_immoneuf": ad.get("isImmoneuf"),
        "discount": ad.get("discount"),
        "has_shipping": ad.get("hasShipping"),
        "is_ecommerce": ad.get("isEcommerce"),
        "default_image": ad.get("defaultImage"),
        "image_count": len(images),
        "images": images,
    }


def parse_page(html):
    soup = BeautifulSoup(html, "html.parser")
    script = soup.select_one("script#__NEXT_DATA__")
    if not script or not script.string:
        return None, None, None
    data = json.loads(script.string)
    component = (
        data.get("props", {}).get("pageProps", {}).get("componentProps") or {}
    )
    ads_wrapper = component.get("ads") or {}
    ads = [extract_ad(a) for a in ads_wrapper.get("ads") or []]
    total = ads_wrapper.get("totalListingAds")
    next_link = soup.select_one('link[rel="next"]')
    next_href = next_link["href"] if next_link else None
    return ads, total, next_href


def fetch_page(session, url, timeout, retries, backoff=5):
    last_error = None
    for attempt in range(1, retries + 1):
        retry_after = None
        try:
            resp = session.get(url, timeout=timeout)
            if resp.status_code == 200:
                ads, total, next_href = parse_page(resp.text)
                if ads is not None:
                    return ads, total, next_href
                last_error = "no __NEXT_DATA__ found in response"
            else:
                last_error = f"HTTP {resp.status_code}"
                if resp.status_code in RETRYABLE_STATUS:
                    retry_after = resp.headers.get("Retry-After")
        except (requests.RequestException, json.JSONDecodeError) as exc:
            last_error = f"{type(exc).__name__}: {exc}"
        wait = backoff * attempt
        if retry_after:
            try:
                wait = max(wait, int(float(retry_after)))
            except ValueError:
                pass
        wait += random.uniform(0, wait / 2)
        if attempt < retries:
            log.warning("page fetch failed (%s) on attempt %s/%s — retrying in %.0fs",
                        last_error, attempt, retries, wait)
            time.sleep(wait)
    raise RuntimeError(f"failed to fetch {url}: {last_error}")


def build_page_url(base, page):
    if page <= 1:
        return base
    sep = "&" if "?" in base else "?"
    return f"{base}{sep}o={page}"


def extract_detail(ad):
    """Flatten every structured param of a listing's detail page."""
    details = {}
    labels = {}
    for group in (ad.get("params") or {}).values():
        if not isinstance(group, list):
            continue
        for item in group:
            if not isinstance(item, dict) or not item.get("key"):
                continue
            key = item["key"]
            if key in DETAIL_SKIP_KEYS:
                continue
            if key.isdigit():
                key = "sector"
            value = item.get("fullValue", item.get("value"))
            if value in ("1", "0") and (
                key.startswith("car_") or key in {"cd_mp3_bt", "first_owner"}
            ):
                value = "Oui" if value == "1" else "Non"
            details[key] = value
            labels[key] = item.get("label") or key

    seller = ad.get("seller") or {}
    price = ad.get("price") or {}
    badges = [b.get("key") for b in (seller.get("badges") or [])
              if isinstance(b, dict) and b.get("key")]
    return {
        "details": details,
        "labels": labels,
        "detail_price": price.get("value"),
        "phone": ad.get("phone") or None,
        "date_posted_exact": ad.get("listTime"),
        "video_count": len(ad.get("videos") or []) or None,
        "phone_hidden": ad.get("isPhoneHidden"),
        "phone_verified": ad.get("isPhoneVerified"),
        "seller_address": seller.get("address"),
        "seller_website": seller.get("website"),
        "seller_badges": ", ".join(badges) or None,
        "seller_listings": seller.get("activeListingsCount"),
    }


def parse_detail_page(html):
    soup = BeautifulSoup(html, "html.parser")
    script = soup.select_one("script#__NEXT_DATA__")
    if not script or not script.string:
        return None
    data = json.loads(script.string)
    page_props = (data.get("props") or {}).get("pageProps") or {}
    ad = ((page_props.get("componentProps") or {}).get("adInfo") or {}).get("ad")
    if isinstance(ad, dict) and ad.get("params"):
        return extract_detail(ad)
    for value in (page_props.get("apolloState") or {}).values():
        if isinstance(value, dict) and isinstance(value.get("ad"), dict):
            if value["ad"].get("params"):
                return extract_detail(value["ad"])
    return None


RETRYABLE_STATUS = {429, 500, 502, 503, 504}

# A 200 whose page carries no ad is a dead/deleted listing, not a transient
# fault: give up early and cache the failure so it is never fetched again.
NO_AD_ATTEMPTS = 2


def fetch_detail(session, url, timeout, retries, backoff=3):
    last_error = None
    no_ad_attempts = 0
    for attempt in range(1, retries + 1):
        retry_after = None
        try:
            resp = session.get(url, timeout=timeout)
            if resp.status_code in (404, 410):
                return {"error": f"HTTP {resp.status_code}", "details": {}}
            if resp.status_code == 200:
                detail = parse_detail_page(resp.text)
                if detail is not None:
                    return detail
                last_error = "no ad data in __NEXT_DATA__"
                no_ad_attempts += 1
                if no_ad_attempts >= NO_AD_ATTEMPTS:
                    return {"error": last_error, "details": {}}
            else:
                last_error = f"HTTP {resp.status_code}"
                if resp.status_code in RETRYABLE_STATUS:
                    retry_after = resp.headers.get("Retry-After")
        except (requests.RequestException, json.JSONDecodeError) as exc:
            last_error = f"{type(exc).__name__}: {exc}"
        wait = backoff * attempt
        if retry_after:
            try:
                wait = max(wait, int(float(retry_after)))
            except ValueError:
                pass
        # Jitter keeps parallel workers from retrying in lockstep, which is
        # what turns a soft rate limit into a hard 429 ban.
        wait += random.uniform(0, wait / 2)
        if attempt < retries:
            log.warning("detail fetch failed (%s) on attempt %s/%s — retrying in %.0fs",
                        last_error, attempt, retries, wait)
            time.sleep(wait)
    raise RuntimeError(f"failed to fetch {url}: {last_error}")


_thread_local = threading.local()


def thread_session():
    session = getattr(_thread_local, "session", None)
    if session is None:
        session = requests.Session()
        session.headers.update(HEADERS)
        _thread_local.session = session
    return session


def crawl_details(records, output_dir, timeout, retries, workers, delay,
                  max_details=0):
    """Fetch every listing's detail page, one resumable JSON file per ad."""
    details_dir = output_dir / "details"
    details_dir.mkdir(parents=True, exist_ok=True)

    targets = [r for r in records if r.get("url") and r.get("list_id")]
    pending = [r for r in targets
               if not (details_dir / f"{r['list_id']}.json").exists()]
    if max_details:
        pending = pending[:max_details]
    log.info("details: %s/%s listings missing a cached detail page",
             len(pending), len(targets))
    if not pending:
        return 0

    done = 0
    failed = 0
    lock = threading.Lock()

    def work(rec):
        if delay:
            time.sleep(delay)
        session = thread_session()
        try:
            detail = fetch_detail(session, rec["url"], timeout, retries)
        except RuntimeError as exc:
            log.warning("%s", exc)
            return False
        path = details_dir / f"{rec['list_id']}.json"
        tmp = details_dir / f".{rec['list_id']}.tmp"
        tmp.write_text(json.dumps(detail, ensure_ascii=False), encoding="utf-8")
        tmp.replace(path)
        return True

    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futures = {pool.submit(work, rec): rec for rec in pending}
        for future in as_completed(futures):
            ok = future.result()
            with lock:
                done += 1
                if not ok:
                    failed += 1
                elif done % 250 == 0 or done == len(pending):
                    log.info("details: %s/%s fetched (%s failed)",
                             done, len(pending), failed)

    log.info("details: finished %s fetches (%s failed)", done, failed)
    return done


def _merge_details(output_dir, records):
    """Fold cached detail pages into their listing record."""
    details_dir = output_dir / "details"
    merged = 0
    for rec in records:
        labels = dict(rec.get("detail_labels") or {})
        params = dict(rec.get("other_params") or {})
        # "annonce" = listed on the search card, "detail" = only on the ad page.
        rec["price_source"] = "annonce" if rec.get("price") else None
        if details_dir.exists() and rec.get("list_id"):
            path = details_dir / f"{rec['list_id']}.json"
            if path.exists():
                try:
                    blob = json.loads(path.read_text(encoding="utf-8"))
                except json.JSONDecodeError:
                    blob = {}
                details = blob.get("details")
                if isinstance(details, dict):
                    params.update(details)
                    labels.update(blob.get("labels") or {})
                    merged += 1
                    if not rec.get("price") and blob.get("detail_price"):
                        rec["price"] = blob["detail_price"]
                        rec["currency"] = rec.get("currency") or "DH"
                        rec["price_source"] = "detail"
                    if not rec.get("seller_phone") and blob.get("phone"):
                        rec["seller_phone"] = blob["phone"]
                    for field in DETAIL_EXTRA_FIELDS:
                        if blob.get(field) is not None:
                            rec[field] = blob[field]

        # Essentials (brand / model) come from the detail page when available,
        # otherwise fall back to what the title lets us infer.
        guessed_brand, guessed_model = infer_brand_model(rec.get("title"))
        if not params.get("brand"):
            params["brand"] = rec.get("brand") or guessed_brand
        labels.setdefault("brand", "Marque")
        if not params.get("model") and guessed_model:
            params["model"] = guessed_model
        if params.get("model"):
            labels.setdefault("model", "Modèle")

        # Dealers often withhold the price; recover it from the description.
        if not rec.get("price"):
            text_price = price_from_text(rec.get("description"))
            if text_price:
                rec["price"] = text_price
                rec["currency"] = rec.get("currency") or "DH"
                rec["price_source"] = "description"

        rec["brand"] = params["brand"]
        rec["other_params"] = params
        rec["detail_labels"] = labels
    return merged



def load_progress(output_dir):
    progress_file = output_dir / "progress.json"
    if progress_file.exists():
        return json.loads(progress_file.read_text(encoding="utf-8"))
    return {}


def save_progress(output_dir, progress):
    progress["last_updated"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    (output_dir / "progress.json").write_text(
        json.dumps(progress, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def crawl(base_url, output_dir, max_pages, delay, timeout, retries, session,
          first_page=0, last_page=0):
    """Walk the search pages, caching each one.

    first_page/last_page let a scheduled run re-fetch the head of the feed
    (newest ads first) to pick up new listings without throwing away the rest
    of the cache.
    """
    pages_dir = output_dir / "pages"
    pages_dir.mkdir(parents=True, exist_ok=True)

    progress = load_progress(output_dir)
    completed = set(progress.get("completed") or [])
    if first_page:
        start_page = first_page
    else:
        start_page = (max(completed) if completed else 0) + 1
    total_listings = progress.get("total_listings")
    ads_per_page = progress.get("ads_per_page")

    if last_page:
        upper = last_page + 1
    else:
        upper = max_pages + 1 if max_pages else 1_000_000
    for page in range(start_page, upper):
        url = build_page_url(base_url, page)
        log.info("page %s: fetching %s", page, url)
        try:
            ads, total, next_href = fetch_page(session, url, timeout, retries)
        except RuntimeError:
            log.exception("stopping — page %s could not be fetched", page)
            break

        if total_listings is None:
            total_listings = total
        if ads_per_page is None and ads:
            ads_per_page = len(ads)

        if ads:
            (pages_dir / f"page_{page:04d}.json").write_text(
                json.dumps(ads, ensure_ascii=False, indent=1), encoding="utf-8"
            )

        completed.add(page)
        progress.update({
            "url": base_url,
            "total_listings": total_listings,
            "ads_per_page": ads_per_page,
            "completed": sorted(completed),
        })
        save_progress(output_dir, progress)

        estimated = (
            math.ceil(total_listings / ads_per_page)
            if total_listings and ads_per_page else None
        )
        if estimated:
            log.info("page %s/%s | %s ads cached",
                     page, estimated, len(completed))

        if not next_href:
            log.info("end of pagination reached at page %s", page)
            progress["finished"] = True
            save_progress(output_dir, progress)
            break

        if estimated and page >= estimated:
            break

        if delay:
            time.sleep(delay)

    return completed


# A car in Morocco is never sold for 300 DH, and the feed fills the price slot
# with a placeholder (or a phone number: 0666259012 -> 666259012) whenever a
# seller says "prix à discuter". Inspecting the extremes showed junk ads
# ("Khouribga", "Kia Picanto 2010") carrying millions, so anything outside
# 5 000 - 5 000 000 DH is a placeholder rather than a price.
PRICE_PLAUSIBLE_MIN = 5_000
PRICE_PLAUSIBLE_MAX = 5_000_000

def _dedup_key(rec):
    """Key meaning "same car, same seller, same site".

    The site is part of the key so two marketplaces that both list one car
    stay as two rows (a cross-site report flags them instead of merging).
    The seller falls back from id to name because only avito exposes an id,
    and price joins the key only when present: a re-post that hides its price
    still merges with its priced twin, but two rows with neither seller nor
    price carry too little signal to call them the same car.

    Mileage is compared exactly: a dealer stocking three identical cars at
    the same price would otherwise collapse into one, while a re-post copies
    the same odometer figure verbatim.
    """
    params = rec.get("other_params") or {}
    parts = [fold(str(rec.get("source") or "avito"))]
    for field in ("brand", "model", "year", "location"):
        value = params.get(field) if field in ("brand", "model") else rec.get(field)
        if value is None or (isinstance(value, str) and not value.strip()):
            return None  # too little signal to call two ads the same car
        parts.append(fold(value) if isinstance(value, str) else value)
    parts.append(rec.get("mileage_km"))  # None matches only None
    seller = rec.get("seller_id") or rec.get("seller_name") or ""
    price = rec.get("price")
    if not seller and not isinstance(price, (int, float)):
        return None
    parts.append(fold(str(seller)) if seller else "")
    parts.append(price if isinstance(price, (int, float)) else None)
    return tuple(parts)


def _base_ads(rec):
    """Ads a row already stands for (re-cleaning must not reset the count)."""
    try:
        return max(1, int(rec.get("merged_ads") or 1))
    except (TypeError, ValueError):
        return 1


def _richness(rec):
    """How much usable data a row carries, so the best row survives a merge."""
    params = rec.get("other_params") or {}
    filled = sum(1 for v in params.values() if v not in (None, "", []))
    filled += sum(1 for k, v in rec.items()
                  if v not in (None, "", [], {}) and k != "other_params")
    return (filled, len(rec.get("images") or []), rec.get("image_count") or 0)


def clean_records(records):
    """Make the export presentable: drop impossible prices, merge re-posts.

    Returns (records, stats). The raw page/detail caches under data/ are never
    modified, so a dropped row can always be recovered from the source.
    """
    stats = {"dropped_bad_price": 0, "merged_ads": 0, "merged_groups": 0}

    plausible = []
    for rec in records:
        price = rec.get("price")
        if isinstance(price, (int, float)) and not (
                PRICE_PLAUSIBLE_MIN < price < PRICE_PLAUSIBLE_MAX):
            stats["dropped_bad_price"] += 1
            continue
        plausible.append(rec)

    groups = defaultdict(list)
    ordered = []
    for rec in plausible:
        key = _dedup_key(rec)
        if key is None:
            rec["merged_ads"] = _base_ads(rec)
            ordered.append(rec)
            continue
        groups[key].append(rec)
        ordered.append(rec)

    dropped = set()
    for group in groups.values():
        if len(group) < 2:
            group[0]["merged_ads"] = _base_ads(group[0])
            continue
        keeper = max(group, key=_richness)
        keeper["merged_ads"] = sum(_base_ads(rec) for rec in group)
        stats["merged_ads"] += len(group) - 1
        stats["merged_groups"] += 1
        for rec in group:
            if rec is not keeper:
                dropped.add(id(rec))

    cleaned = [rec for rec in ordered
               if id(rec) not in dropped and rec.get("merged_ads")]
    stats["total_after_clean"] = len(cleaned)
    # Rows collapsed by earlier clean passes count too, so one re-clean of an
    # already-cleaned export still reports the true number of grouped ads.
    stats["merged_total"] = sum(_base_ads(rec) for rec in cleaned) - len(cleaned)
    return cleaned, stats


def load_raw_records(output_dir, allow_all_categories=False):
    """Cached avito search pages, deduplicated and enriched from the cache.

    Returns (records, exact_duplicates, detail_enriched). Every record is
    tagged with its source, so the multi-source merge can tell sites apart.
    The raw caches under data/ are never modified.
    """
    pages_dir = output_dir / "pages"
    records = []
    seen = set()
    duplicates = 0
    if pages_dir.exists():
        for path in sorted(pages_dir.glob("page_*.json")):
            for ad in json.loads(path.read_text(encoding="utf-8")):
                if not allow_all_categories and not (ad.get("category") or "").startswith("Voitures"):
                    continue
                key = ad.get("list_id") or ad.get("url")
                if not key:
                    continue
                if key in seen:
                    duplicates += 1
                    continue
                seen.add(key)
                ad.setdefault("source", "avito")
                records.append(ad)

    merged = _merge_details(output_dir, records)
    return records, duplicates, merged


def write_csv(records, path):
    """Write the flat CSV export (one row per car). Returns the column list."""
    columns = export_columns(records)
    with open(path, "w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns)
        writer.writeheader()
        for ad in records:
            params = ad.get("other_params") or {}
            row = {col: (params[col] if col in params else ad.get(col))
                   for col in columns}
            row["images_json"] = json.dumps(ad.get("images") or [],
                                            ensure_ascii=False)
            row["other_params_json"] = json.dumps(
                ad.get("other_params") or {}, ensure_ascii=False
            )
            writer.writerow(row)
    return columns


def consolidate(output_dir, allow_all_categories=False, stats=None):
    records, duplicates, merged = load_raw_records(output_dir,
                                                   allow_all_categories)
    records, clean_stats = clean_records(records)
    if stats is not None:
        stats.update(clean_stats)
    stats = stats if stats is not None else clean_stats

    json_out = output_dir / "avito_cars.json"
    json_out.write_text(
        json.dumps(records, ensure_ascii=False, indent=1), encoding="utf-8"
    )

    csv_path = output_dir / "avito_cars.csv"
    columns = write_csv(records, csv_path)

    log.info("consolidated %s listings (%s exact duplicates skipped, "
             "%s enriched with a detail page) -> %s after cleaning",
             len(records) + stats.get("dropped_bad_price", 0)
             + stats.get("merged_ads", 0),
             duplicates, merged, len(records))
    log.info("cleaned: dropped %s rows with an impossible price, "
             "collapsed %s re-posted ads into %s rows",
             stats.get("dropped_bad_price", 0),
             stats.get("merged_ads", 0), stats.get("merged_groups", 0))
    log.info("wrote %s (%s columns)", csv_path, len(columns))
    log.info("wrote %s", json_out)
    return records


def detail_keys(records):
    """Detail-param keys in display order, from every record that has one."""
    keys = {k for r in records for k in (r.get("other_params") or {})}
    ordered = [k for k in DETAIL_PREFERRED_ORDER if k in keys]
    ordered += sorted(keys - set(ordered))
    return ordered


def export_columns(records):
    columns = list(CSV_COLUMNS)
    for key in detail_keys(records) + DETAIL_EXTRA_FIELDS:
        if key not in columns:
            columns.append(key)
    return columns


# Essentials first, so the sheet is readable at a glance: one row per car with
# marque / modèle / année / prix up front, everything else after them.
EXCEL_ESSENTIAL_COLUMNS = ["brand", "model", "year", "price"]

EXCEL_ESSENTIAL_LABELS = {
    "brand": "Marque", "model": "Modèle", "year": "Année", "price": "Prix",
}

# French headers for the fixed columns, so the whole sheet reads consistently.
EXCEL_COLUMN_LABELS = {
    "id": "ID", "list_id": "ID annonce", "url": "Lien", "title": "Titre",
    "source": "Site",
    "description": "Description", "category": "Catégorie",
    "ad_type": "Type d'annonce", "currency": "Devise",
    "monthly_payment": "Mensualité", "old_price": "Ancien prix",
    "mileage_km": "Kilométrage (km)", "fuel": "Carburant",
    "gearbox": "Boîte de vitesses", "location": "Ville", "city_id": "ID ville",
    "area_id": "ID région", "date_posted": "Publication",
    "price_source": "Origine du prix", "seller_id": "ID vendeur",
    "seller_type": "Type de vendeur", "seller_name": "Vendeur",
    "seller_phone": "Téléphone", "seller_phone_verified": "Téléphone validé",
    "seller_verified": "Vendeur validé", "is_professional": "Professionnel",
    "is_premium": "Premium", "is_urgent": "Urgent", "is_hot_deal": "Bon plan",
    "is_shop": "Boutique", "is_car_checked": "Voiture vérifiée",
    "is_delivery": "Livraison", "is_highlighted": "Mis en avant",
    "is_immoneuf": "Immoneuf", "discount": "Remise", "has_shipping": "Expédition",
    "is_ecommerce": "E-commerce", "default_image": "Image principale",
    "image_count": "Nombre de photos", "images_json": "Toutes les photos",
    "date_posted_exact": "Date exacte", "video_count": "Nombre de vidéos",
    "phone_hidden": "Téléphone masqué", "phone_verified": "Téléphone validé (fiche)",
    "seller_address": "Adresse", "seller_website": "Site web",
    "seller_badges": "Badges", "seller_listings": "Annonces du vendeur",
}

EXCEL_IDENTITY_COLUMNS = ["id", "list_id", "source", "url", "title"]

EXCEL_CORE_COLUMNS = [
    "description", "category", "ad_type", "currency", "monthly_payment",
    "old_price", "mileage_km", "fuel", "gearbox",
]

EXCEL_TAIL_COLUMNS = [
    "location", "city_id", "area_id", "date_posted", "price_source",
    "merged_ads",
    "seller_id", "seller_type",
    "seller_name", "seller_phone", "seller_phone_verified", "seller_verified",
    "is_professional", "is_premium", "is_urgent", "is_hot_deal", "is_shop",
    "is_car_checked", "is_delivery", "is_highlighted", "is_immoneuf",
    "discount", "has_shipping", "is_ecommerce", "default_image", "image_count",
    "images_json",
]

# (label, key, read from other_params, higher is better)
COMPLETENESS_FIELDS = [
    ("Marque", "brand", True, True),
    ("Modèle", "model", True, True),
    ("Année", "year", False, True),
    ("Prix", "price", False, True),
    ("Titre", "title", False, True),
    ("Description", "description", False, True),
    ("Kilométrage", "mileage_km", False, True),
    ("Carburant", "fuel", False, True),
    ("Boîte de vitesses", "gearbox", False, True),
    ("Puissance fiscale (CV)", "pfiscale", True, True),
    ("Nombre de portes", "doors", True, True),
    ("Origine", "v_origin", True, True),
    ("État", "auto_condition", True, True),
    ("Première main", "first_owner", True, True),
    ("Ville", "location", False, True),
    ("Téléphone vendeur", "seller_phone", False, True),
    ("Photos", "image_count", False, True),
]


def _completeness(records):
    total = len(records) or 1
    rows = []
    for label, key, from_params, _ in COMPLETENESS_FIELDS:
        filled = 0
        for rec in records:
            value = ((rec.get("other_params") or {}).get(key)
                     if from_params else rec.get(key))
            if value in (None, "", [], {}) or value == 0 and key == "image_count":
                continue
            filled += 1
        rows.append([label, filled, f"{100 * filled / total:.1f} %"])
    return rows


PRICE_BANDS = [
    (0, 50_000, "Moins de 50 000 DH"),
    (50_000, 100_000, "50 000 – 100 000 DH"),
    (100_000, 150_000, "100 000 – 150 000 DH"),
    (150_000, 200_000, "150 000 – 200 000 DH"),
    (200_000, 300_000, "200 000 – 300 000 DH"),
    (300_000, 500_000, "300 000 – 500 000 DH"),
    (500_000, None, "Plus de 500 000 DH"),
]

YEAR_BANDS = [
    (0, 2000, "Avant 2000"),
    (2000, 2005, "2000 – 2004"),
    (2005, 2010, "2005 – 2009"),
    (2010, 2015, "2010 – 2014"),
    (2015, 2020, "2015 – 2019"),
    (2020, None, "2020 et plus"),
]


def _band_counts(records, key, bands):
    rows = []
    for low, high, label in bands:
        count = 0
        for rec in records:
            value = rec.get(key)
            if not isinstance(value, (int, float)):
                continue
            if value < low or (high is not None and value >= high):
                continue
            count += 1
        rows.append([label, count])
    return rows


def missing_essentials(rec):
    """Which of Marque / Modèle / Année / Prix are absent from a row."""
    params = rec.get("other_params") or {}
    missing = []
    for key in ("brand", "model", "year", "price"):
        value = params.get(key) if key in ("brand", "model") else rec.get(key)
        if value in (None, "", []):
            missing.append(EXCEL_ESSENTIAL_LABELS[key])
    return missing




def fold(text):
    """Lowercase and strip accents so 'Citroën' matches 'citroen'."""
    decomposed = unicodedata.normalize("NFKD", str(text or ""))
    return "".join(c for c in decomposed if not unicodedata.combining(c)).lower()


def infer_brand_model(title):
    """Best-effort (brand, model) from a listing title, accents aside.

    Detail pages are authoritative; this only fills the essentials for the
    handful of listings whose detail page could never be fetched.
    """
    clean = re.sub(r"\s+", " ", str(title or "")).strip()
    if not clean:
        return "Inconnu", None
    folded = fold(clean)
    for keyword in sorted(BRAND_KEYWORDS, key=len, reverse=True):
        match = re.search(rf"(?<!\w){re.escape(fold(keyword))}(?!\w)", folded)
        if not match:
            continue
        brand = BRAND_DISPLAY.get(fold(keyword), clean[match.start():match.end()].title())
        rest = clean[match.end():]
        # The model is just the next token(s): everything after it is trim,
        # engine or sales boilerplate ("occasion", "toutes options", year...).
        model = re.split(
            r"\s*[\s–\-—,:;|/()]|\s+(?:\d{4}|diesel|essence|hybride|"
            r"automatique|manuelle|manuel|occasion|neuf|etat|état)\b",
            rest.strip(), maxsplit=1, flags=re.I)[0].strip()
        for phrase in MODEL_PHRASES:
            if fold(model).startswith(phrase):
                model = phrase.title()
                break
        return brand, model or None
    return clean.split()[0], " ".join(clean.split()[1:2]) or None


def _top_counts(records, key, limit=20):
    counts = defaultdict(lambda: [0, []])
    for rec in records:
        value = rec.get(key)
        if value in (None, ""):
            value = "Non spécifié"
        bucket = counts[value]
        bucket[0] += 1
        price = rec.get("price")
        if isinstance(price, (int, float)):
            bucket[1].append(price)
    ordered = sorted(counts.items(), key=lambda kv: kv[1][0], reverse=True)[:limit]
    return [(label, stats[0], int(sum(stats[1]) / len(stats[1])) if stats[1] else None)
            for label, stats in ordered]


SOURCE_LABELS = {
    "avito": "avito.ma",
    "kifal": "occasion.kifal.ma",
    "oneclickdrive": "oneclickdrive.ma",
}


def source_label(source):
    return SOURCE_LABELS.get(source or "avito", source or "avito")


def export_excel(records, path, clean_stats=None, cross_dupes=None):
    if not HAS_OPENPYXL:
        log.warning("openpyxl not installed — skipping Excel export "
                    "(pip install openpyxl)")
        return

    clean_stats = clean_stats or {}
    wb = Workbook()

    header_fill = PatternFill("solid", fgColor="29A160")
    header_font = Font(color="FFFFFF", bold=True)
    table_fill = PatternFill("solid", fgColor="212B36")
    kpi_font = Font(size=12, bold=True)
    note_font = Font(size=9, italic=True, color="666666")

    summary = wb.active
    summary.title = "Résumé"
    prices = [r["price"] for r in records
              if isinstance(r.get("price"), (int, float))]
    brands = {(r.get("other_params") or {}).get("brand") for r in records}
    brands.discard(None)
    sources = sorted({r.get("source") or "avito" for r in records})
    multi_source = len(sources) > 1

    kpis = [
        ("Voitures annoncées", len(records)),
        ("Sites sources", len(sources)),
        ("Annonces regroupées (doublons)",
         clean_stats.get("merged_total", clean_stats.get("merged_ads", 0))),
        ("Lignes écartées (prix aberrant)", clean_stats.get("dropped_bad_price", 0)),
        ("Prix moyen (DH)", int(statistics.mean(prices)) if prices else None),
        ("Prix médian (DH)", int(statistics.median(prices)) if prices else None),
        ("Prix mini / maxi (DH)",
         f"{min(prices):,} / {max(prices):,}".replace(",", " ") if prices else None),
        ("Annonces avec prix", len(prices)),
        ("Annonces professionnelles", sum(1 for r in records if r.get("is_professional"))),
        ("Annonces particuliers", sum(1 for r in records if not r.get("is_professional"))),
        ("Annonces premium", sum(1 for r in records if r.get("is_premium"))),
        ("Marques couvertes", len(brands)),
        ("Villes couvertes", len({r.get("location") for r in records
                                  if r.get("location")})),
        ("Date d'export", time.strftime("%Y-%m-%d %H:%M:%S")),
    ]

    title = (f"T480 — Marché des voitures d'occasion ({source_label(sources[0])})"
             if not multi_source else
             "T480 — Marché des voitures d'occasion (Maroc, sources multiples)")
    summary.cell(1, 1, title).font = Font(size=16, bold=True)
    summary.cell(2, 1, "Une ligne par voiture, après nettoyage").font = kpi_font
    for i, (label, value) in enumerate(kpis, start=4):
        summary.cell(i, 1, label).font = kpi_font
        summary.cell(i, 2, value)
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 24

    def add_table(ws, row_start, title, headers, rows):
        ws.cell(row_start, 4, title).font = Font(size=13, bold=True)
        row_start += 1
        for col, header in enumerate(headers, start=4):
            cell = ws.cell(row_start, col, header)
            cell.fill = table_fill
            cell.font = header_font
        for i, row in enumerate(rows, start=row_start + 1):
            for j, value in enumerate(row, start=4):
                ws.cell(i, j, value)
        ws.column_dimensions["E"].width = 28
        ws.column_dimensions["F"].width = 12
        ws.column_dimensions["G"].width = 14
        return row_start + len(rows)

    rows = 4
    if multi_source:
        rows = add_table(summary, rows, "Par site",
                         ["Site", "Annonces", "Prix moyen (DH)"],
                         [(source_label(source), count, avg)
                          for source, count, avg
                          in _top_counts(records, "source", limit=len(sources))])
        rows += 2
    rows = add_table(summary, rows, "Villes les plus actives",
                     ["Ville", "Annonces", "Prix moyen (DH)"],
                     _top_counts(records, "location", limit=15))
    rows = add_table(summary, rows + 2, "Marques les plus annoncées",
                     ["Marque", "Annonces", "Prix moyen (DH)"],
                     _top_counts(records, "brand", limit=20))
    rows = add_table(summary, rows + 2, "Répartition des prix",
                     ["Tranche de prix", "Annonces"],
                     _band_counts(records, "price", PRICE_BANDS))
    rows = add_table(summary, rows + 2, "Répartition des millésimes",
                     ["Année", "Annonces"],
                     _band_counts(records, "year", YEAR_BANDS))
    rows = add_table(summary, rows + 2, "Carburant",
                     ["Carburant", "Annonces", "Prix moyen (DH)"],
                     _top_counts(records, "fuel"))
    rows = add_table(summary, rows + 2, "Boîte de vitesses",
                     ["Boîte", "Annonces", "Prix moyen (DH)"],
                     _top_counts(records, "gearbox"))
    rows = add_table(summary, rows + 2, "Complétude des champs",
                     ["Champ", "Remplis", "Complétude"],
                     _completeness(records))

    note = summary.cell(rows + 3, 4, "Note methodology")
    note.font = Font(size=12, bold=True)
    source_names = ", ".join(source_label(source) for source in sources)
    notes = [
        f"Données extraites des pages de recherche et de fiche annonce de "
        f"{source_names}.",
        "Nettoyage : prix hors 5 000 – 5 000 000 DH ecarte "
        f"({clean_stats.get('dropped_bad_price', 0)} lignes : prix "
        "placeholder « a discuter » ou numero de telephone lu comme un prix).",
        f"Doublons : annonces re-publiees par un meme vendeur pour la meme "
        f"voiture fusionnees ({clean_stats.get('merged_ads', 0)} annonces "
        f"reparties sur {clean_stats.get('merged_groups', 0)} groupes). "
        "La colonne « Annonces regroupees » indique combien d'annonces "
        "representent chaque ligne.",
        "Marque, modele et annee proviennent de la fiche annonce ; a defaut, "
        "ils sont deduits du titre. Le prix est celui affiche par le vendeur ; "
        "« Origine du prix » precise s'il a ete releve dans le texte.",
        "Les lignes incompletes sont listees dans l'onglet « A verifier ».",
        "Les donnees brutes restent disponibles dans le depot (data/cache).",
    ]
    if multi_source:
        notes.append(
            "Les doublons ne sont jamais fusionnes d'un site a l'autre : une "
            "meme voiture annoncee sur deux sites reste deux lignes, signalees "
            "dans l'onglet « Doublons entre sites ».")
    for offset, text in enumerate(notes, start=1):
        cell = summary.cell(rows + 3 + offset, 4, text)
        cell.font = note_font

    data = wb.create_sheet("Annonces")
    taken = set(EXCEL_ESSENTIAL_COLUMNS + EXCEL_IDENTITY_COLUMNS
                + EXCEL_CORE_COLUMNS + EXCEL_TAIL_COLUMNS)
    param_cols = [k for k in detail_keys(records)
                  if k not in taken and k not in DETAIL_EXTRA_FIELDS]
    extra_cols = [k for k in DETAIL_EXTRA_FIELDS if k not in taken]
    columns = (EXCEL_ESSENTIAL_COLUMNS + EXCEL_IDENTITY_COLUMNS
               + EXCEL_CORE_COLUMNS + param_cols + EXCEL_TAIL_COLUMNS
               + extra_cols)

    labels = {}
    for rec in records:
        for key, label in (rec.get("detail_labels") or {}).items():
            labels.setdefault(key, label)
    param_set = set(param_cols) | {"brand", "model"}
    headers = []
    for name in columns:
        if name in param_set:
            # brand/model live in other_params but are essentials, so they are
            # read from there too and carry the same French header.
            header = labels.get(name) or name
        else:
            header = name
        header = (EXCEL_ESSENTIAL_LABELS.get(name)
                  or EXCEL_COLUMN_LABELS.get(name, header))
        candidate = header
        suffix = 2
        while candidate in headers:
            candidate = f"{header} ({suffix})"
            suffix += 1
        headers.append(candidate)

    widths = {
        "title": 40, "description": 60, "seller_name": 20, "location": 18,
        "default_image": 50, "images_json": 60, "url": 55,
        "seller_address": 34, "seller_website": 30, "seller_badges": 18,
        "date_posted_exact": 20, "other_params_json": 40,
        "brand": 16, "model": 20, "price": 14, "year": 10,
    }
    for col, (name, header) in enumerate(zip(columns, headers), start=1):
        cell = data.cell(1, col, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")
        width = widths.get(name, max(14, min(26, len(header) + 3))
                           if name in param_set else 14)
        data.column_dimensions[get_column_letter(col)].width = width

    for i, rec in enumerate(records, start=2):
        params = rec.get("other_params") or {}
        for col, name in enumerate(columns, start=1):
            if name == "images_json":
                value = json.dumps(rec.get("images") or [], ensure_ascii=False)
            elif name in param_set:
                value = params.get(name)
            else:
                value = rec.get(name)
            cell = data.cell(i, col, value)
            if name in ("price", "monthly_payment", "old_price") and isinstance(value, (int, float)):
                cell.number_format = '#,##0 "DH"'
            if name == "mileage_km" and isinstance(value, (int, float)):
                cell.number_format = '#,##0'
            if name == "description":
                cell.alignment = Alignment(wrap_text=True)

    data.freeze_panes = "A2"
    data.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(records) + 1}"

    # Everything a human still has to look at, in one short list.
    check = wb.create_sheet("À vérifier")
    check_cols = ["Marque", "Modèle", "Année", "Prix", "Champ(s) manquant(s)",
                  "Ville", "Site", "Vendeur", "Type de vendeur",
                  "Annonces regroupées", "Lien", "Titre"]
    for col, header in enumerate(check_cols, start=1):
        cell = check.cell(1, col, header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(vertical="center")
        check.column_dimensions[get_column_letter(col)].width = (
            55 if header == "Lien" else 28 if header == "Titre" else 16)

    flagged = 0
    for i, rec in enumerate(records, start=2):
        missing = missing_essentials(rec)
        if not missing:
            continue
        flagged += 1
        params = rec.get("other_params") or {}
        values = [params.get("brand"), params.get("model"), rec.get("year"),
                  rec.get("price"), ", ".join(missing), rec.get("location"),
                  source_label(rec.get("source")),
                  rec.get("seller_name"), rec.get("seller_type"),
                  rec.get("merged_ads"), rec.get("url"), rec.get("title")]
        for col, value in enumerate(values, start=1):
            cell = check.cell(flagged + 1, col, value)
            if col == 4 and isinstance(value, (int, float)):
                cell.number_format = '#,##0 "DH"'
    check.freeze_panes = "A2"
    if flagged:
        check.auto_filter.ref = f"A1:{get_column_letter(len(check_cols))}{flagged + 1}"

    # One car sold on two sites stays two rows; this sheet is where a human
    # sees those pairs (never merged automatically).
    if cross_dupes:
        cross = wb.create_sheet("Doublons entre sites")
        cross_cols = ["Marque", "Modèle", "Année", "Prix", "Ville", "Sites",
                      "Nb annonces", "Vendeurs", "Liens"]
        widths_cross = {"Liens": 55, "Vendeurs": 34, "Sites": 30}
        for col, header in enumerate(cross_cols, start=1):
            cell = cross.cell(1, col, header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(vertical="center")
            cross.column_dimensions[get_column_letter(col)].width = widths_cross.get(header, 16)
        for i, row in enumerate(cross_dupes, start=2):
            values = [row.get(key) for key in
                      ("brand", "model", "year", "price", "location",
                       "sites", "count", "sellers", "urls")]
            for col, value in enumerate(values, start=1):
                cell = cross.cell(i, col, value)
                if col == 4 and isinstance(value, (int, float)):
                    cell.number_format = '#,##0 "DH"'
        cross.freeze_panes = "A2"
        cross.auto_filter.ref = (f"A1:{get_column_letter(len(cross_cols))}"
                                 f"{len(cross_dupes) + 1}")

    wb.save(path)
    log.info("wrote %s (%s rows x %s columns, %s rows to review)",
             path, len(records), len(columns), flagged)


def main():
    parser = argparse.ArgumentParser(
        description="Scrape the car listings database on avito.ma"
    )
    parser.add_argument("--url", default=DEFAULT_URL,
                        help="base search URL (default: voitures d'occasion)")
    parser.add_argument("--output-dir", default="data",
                        help="directory for cached pages and exports")
    parser.add_argument("--max-pages", type=int, default=0,
                        help="stop after this many pages (0 = all)")
    parser.add_argument("--delay", type=float, default=1.5,
                        help="seconds to sleep between page requests")
    parser.add_argument("--timeout", type=float, default=30.0,
                        help="HTTP request timeout in seconds")
    parser.add_argument("--retries", type=int, default=5,
                        help="retries per page before giving up")
    parser.add_argument("--fresh", action="store_true",
                        help="clear cached pages and start from scratch")
    parser.add_argument("--refresh-top", type=int, default=0, metavar="N",
                        help="re-fetch the first N search pages (newest ads) "
                             "and keep the rest of the cache")
    parser.add_argument("--export-only", action="store_true",
                        help="only rebuild CSV/JSON/Excel from cached pages")
    parser.add_argument("--excel-out", default="avito_cars.xlsx",
                        help="name of the generated Excel workbook")
    parser.add_argument("--no-excel", action="store_true",
                        help="skip Excel export")
    parser.add_argument("--skip-details", action="store_true",
                        help="do not fetch the per-listing detail pages")
    parser.add_argument("--detail-workers", type=int, default=6,
                        help="parallel detail-page requests (default 6)")
    parser.add_argument("--detail-delay", type=float, default=0.0,
                        help="seconds to sleep before each detail request")
    parser.add_argument("--max-details", type=int, default=0,
                        help="fetch at most N uncached detail pages (0 = all)")
    parser.add_argument("--all-categories", action="store_true",
                        help="include non-car ads leaked into the car feed")
    parser.add_argument("--pending-details", action="store_true",
                        help="print how many detail pages are still missing, "
                             "then exit without scraping or exporting")
    parser.add_argument("--verbose", "-v", action="store_true",
                        help="verbose logging")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    if args.pending_details:
        missing, total = pending_details(output_dir)
        print(f"{missing} detail pages missing out of {total} listings")
        return

    if args.fresh:
        pages_dir = output_dir / "pages"
        if pages_dir.exists():
            for path in pages_dir.glob("page_*.json"):
                path.unlink()
        (output_dir / "progress.json").unlink(missing_ok=True)
        log.info("cleared cached data in %s", output_dir)

    if args.export_only:
        finish(args, output_dir)
        return

    session = requests.Session()
    session.headers.update(HEADERS)

    if args.refresh_top:
        log.info("refreshing the %s first search pages (newest listings)",
                 args.refresh_top)
        crawl(args.url, output_dir, 0, args.delay, args.timeout, args.retries,
              session, first_page=1, last_page=args.refresh_top)
        finish(args, output_dir)
        return

    progress = load_progress(output_dir)
    if progress.get("finished") and not args.max_pages and not args.fresh:
        log.info("crawl already finished (%s pages cached) — consolidating",
                 len(progress.get("completed", [])))
        log.info("use --fresh or --max-pages to fetch more pages")
        finish(args, output_dir)
        return

    if progress.get("completed"):
        log.info("resuming with %s cached pages from previous run",
                 len(progress["completed"]))

    crawl(args.url, output_dir, args.max_pages, args.delay,
          args.timeout, args.retries, session)

    finish(args, output_dir)


def pending_details(output_dir):
    """How many listings still lack a cached detail page."""
    details_dir = output_dir / "details"
    records = consolidate_readonly(output_dir)
    missing = 0
    for rec in records:
        list_id = rec.get("list_id")
        if not rec.get("url") or not list_id:
            continue
        if not (details_dir / f"{list_id}.json").exists():
            missing += 1
    return missing, len(records)


def consolidate_readonly(output_dir, allow_all_categories=False):
    """Deduplicated listing records without writing any export file."""
    records, _, _ = load_raw_records(output_dir, allow_all_categories)
    return records


def finish(args, output_dir):
    """Consolidate, clean, enrich with detail pages, then write the exports."""
    stats = {}
    records = consolidate(output_dir,
                          allow_all_categories=args.all_categories, stats=stats)
    if not args.skip_details:
        crawl_details(records, output_dir, args.timeout, args.retries,
                      args.detail_workers, args.detail_delay,
                      args.max_details)
        records = consolidate(output_dir,
                              allow_all_categories=args.all_categories,
                              stats=stats)
    if not args.no_excel:
        export_excel(records, output_dir / args.excel_out, stats)
    return records


if __name__ == "__main__":
    main()