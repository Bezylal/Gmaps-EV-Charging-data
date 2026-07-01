#!/usr/bin/env python3
"""
EV Charging Station Scraper — Chennai
Uses Google Places API v1 (New) for EV connector-level data.

Search strategy:
  • 1 km grid across bounding box (N:13.415, S:12.689, E:80.362, W:79.904)
  • 800 m radius per grid point  →  ~4,050 overlapping circles, zero coverage gaps
  • Supplemental text-search sweep to catch any edge-case listings
  • Deduplication by place_id before saving

Cost note (Places API Preferred SKU = $0.065/call):
  ~4,050 grid calls × $0.065 ≈ $263  (Google gives $200/month free credit)
  To stay within free credit, raise GRID_SPACING_KM to 1.3  (~$160 est.)

Run:
  Windows  : set GOOGLE_API_KEY=AIzaSyCgZqlzMKgu29C6TM13gfGCjBZlfzWHBLg && python ev_charger_scraper.py
  Mac/Linux: export GOOGLE_API_KEY=your_key_here && python ev_charger_scraper.py

Install dependencies first:
  pip install requests pandas openpyxl
"""

import os
import time
import math
import logging
import requests
import pandas as pd
from typing import Optional

# ── CONFIG — tweak these if needed ─────────────────────────────────────────────
API_KEY = "AIzaSyCgZqlzMKgu29C6TM13gfGCjBZlfzWHBLg"   # your Google API key

BOUNDING_BOX = {
    "lat_south": 12.689,   # South boundary
    "lat_north": 13.415,   # North boundary
    "lng_west":  79.904,   # West boundary
    "lng_east":  80.362,   # East boundary
}
GRID_SPACING_KM  = 1.3    # Distance between grid points in km
SEARCH_RADIUS_M  = 1000   # Radius per grid point in metres
                          # 1000 m > 919 m (half-diagonal of a 1.3 km cell) = complete coverage
REQUESTS_PER_SEC = 3      # Safe rate for free tier (max ~10 is fine too)
OUTPUT_FILE      = "ev_chargers_chennai.xlsx"

# ── API ENDPOINTS ───────────────────────────────────────────────────────────────
NEARBY_URL = "https://places.googleapis.com/v1/places:searchNearby"
TEXT_URL   = "https://places.googleapis.com/v1/places:searchText"

# Field mask — requesting evChargeOptions triggers "Preferred" billing SKU
FIELD_MASK = (
    "places.id,"
    "places.displayName,"
    "places.formattedAddress,"
    "places.location,"
    "places.businessStatus,"
    "places.rating,"
    "places.userRatingCount,"
    "places.regularOpeningHours,"
    "places.evChargeOptions"
)

# ── LOGGING ─────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ── LOOKUP TABLES ───────────────────────────────────────────────────────────────
# Human-readable connector labels (strip the long API prefix)
CONNECTOR_LABELS = {
    "EV_CONNECTOR_TYPE_CCS_COMBO_1":  "CCS1",
    "EV_CONNECTOR_TYPE_CCS_COMBO_2":  "CCS2",
    "EV_CONNECTOR_TYPE_TYPE_2":       "Type 2 (AC)",
    "EV_CONNECTOR_TYPE_J1772":        "Type 1 (J1772)",
    "EV_CONNECTOR_TYPE_CHAdeMO":      "CHAdeMO",
    "EV_CONNECTOR_TYPE_TESLA":        "Tesla / NACS",
    "EV_CONNECTOR_TYPE_GB_T":         "GB/T",
    "EV_CONNECTOR_TYPE_UNSPECIFIED":  "Unspecified",
}

BUSINESS_STATUS_MAP = {
    "OPERATIONAL":        "Open",
    "CLOSED_TEMPORARILY": "Temporarily Closed",
    "CLOSED_PERMANENTLY": "Permanently Closed",
}


# ── SCRAPER CLASS ───────────────────────────────────────────────────────────────
class EVChargingScraper:
    """
    Reusable scraper for EV charging station data via Google Places API v1.

    Extend by subclassing and overriding parse_place() or scrape().

    Usage:
        scraper = EVChargingScraper(api_key="YOUR_KEY")
        records = scraper.scrape()
    """

    def __init__(self, api_key: str):
        if not api_key:
            raise ValueError("API key is required.")
        self.api_key = api_key
        self._last_call_time = 0.0

        # Reuse a single session for connection pooling
        self.session = requests.Session()
        self.session.headers.update({
            "X-Goog-Api-Key":   api_key,
            "X-Goog-FieldMask": FIELD_MASK,
            "Content-Type":     "application/json",
        })

    # ── Rate limiting ─────────────────────────────────────────────────────────
    def _throttle(self):
        """Sleep just enough to stay within REQUESTS_PER_SEC."""
        min_gap = 1.0 / REQUESTS_PER_SEC
        elapsed = time.time() - self._last_call_time
        if elapsed < min_gap:
            time.sleep(min_gap - elapsed)
        self._last_call_time = time.time()

    # ── HTTP POST with exponential-backoff retry ───────────────────────────────
    def _post(self, url: str, body: dict, retries: int = 4) -> Optional[dict]:
        """POST to Places API. Returns parsed JSON or None on permanent failure."""
        for attempt in range(1, retries + 1):
            self._throttle()
            try:
                r = self.session.post(url, json=body, timeout=20)

                if r.status_code == 200:
                    return r.json()

                if r.status_code == 429:
                    # Quota exceeded — back off exponentially
                    wait = 2 ** attempt
                    log.warning(f"Rate limited (attempt {attempt}). Waiting {wait}s …")
                    time.sleep(wait)

                elif r.status_code in (500, 502, 503, 504):
                    # Transient server errors — safe to retry
                    wait = 2 ** attempt
                    log.warning(f"Server error {r.status_code} (attempt {attempt}). Retrying in {wait}s …")
                    time.sleep(wait)

                else:
                    # Permanent error (400 bad request, 401 auth, etc.) — don't retry
                    log.error(f"API error {r.status_code}: {r.text[:400]}")
                    return None

            except requests.exceptions.Timeout:
                log.warning(f"Timeout on attempt {attempt}/{retries}")
                time.sleep(2 ** attempt)
            except requests.exceptions.RequestException as exc:
                log.warning(f"Network error on attempt {attempt}/{retries}: {exc}")
                time.sleep(2 ** attempt)

        log.error("All retries exhausted for this request.")
        return None

    # ── Single grid-point nearby search ───────────────────────────────────────
    def nearby_search(self, lat: float, lng: float) -> list[dict]:
        """
        Search within SEARCH_RADIUS_M of (lat, lng) for EV charging stations.
        Returns up to 20 results (Places API v1 nearby search limit).
        """
        body = {
            "includedTypes":  ["electric_vehicle_charging_station"],
            "maxResultCount": 20,
            "locationRestriction": {
                "circle": {
                    "center": {"latitude": lat, "longitude": lng},
                    "radius": float(SEARCH_RADIUS_M),
                }
            },
        }
        data = self._post(NEARBY_URL, body)
        return data.get("places", []) if data else []

    # ── Text search with pagination ────────────────────────────────────────────
    def text_search(self, query: str) -> list[dict]:
        """
        Text-based search returning up to ~60 results across 3 paginated pages.
        Used as a supplemental sweep after the grid search.
        """
        results = []
        page = 1
        body = {
            "textQuery":      query,
            "includedType":   "electric_vehicle_charging_station",
            "maxResultCount": 20,
        }

        while True:
            data = self._post(TEXT_URL, body)
            if not data:
                break

            batch = data.get("places", [])
            results.extend(batch)
            log.info(f"  Text search page {page}: {len(batch)} result(s)")

            next_token = data.get("nextPageToken")
            if not next_token or not batch:
                break

            # Google requires a short pause before the next-page token becomes valid
            time.sleep(2)
            body = {
                "textQuery":      query,
                "includedType":   "electric_vehicle_charging_station",
                "maxResultCount": 20,
                "pageToken":      next_token,
            }
            page += 1

        return results

    # ── Grid generation ────────────────────────────────────────────────────────
    @staticmethod
    def build_grid(bbox: dict, spacing_km: float) -> list[tuple[float, float]]:
        """
        Create a regular lat/lng grid covering the bounding box.

        Longitude step is scaled by cos(lat) so each cell is ~square on the ground
        (earth's longitude degrees get shorter as you move away from the equator).
        """
        mid_lat  = (bbox["lat_south"] + bbox["lat_north"]) / 2.0
        lat_step = spacing_km / 111.0
        lng_step = spacing_km / (111.0 * math.cos(math.radians(mid_lat)))

        points = []
        lat = bbox["lat_south"]
        while lat <= bbox["lat_north"] + lat_step * 0.5:   # 0.5× tolerance handles float drift
            lng = bbox["lng_west"]
            while lng <= bbox["lng_east"] + lng_step * 0.5:
                points.append((round(lat, 6), round(lng, 6)))
                lng += lng_step
            lat += lat_step

        log.info(
            f"Grid built: {len(points)} points  "
            f"(Δlat={lat_step:.5f}°  Δlng={lng_step:.5f}°  "
            f"radius={SEARCH_RADIUS_M} m)"
        )
        return points

    # ── Parse one API place object into a flat dict ────────────────────────────
    @staticmethod
    def parse_place(place: dict) -> dict:
        """
        Flatten a Places API v1 place object.

        Connector info is stored as pipe-separated strings so each station
        stays on a single Excel row even when it has multiple connector types.

        Example:   connector_types  = "CCS2 | Type 2 (AC)"
                   max_kw_per_type  = "50.0 | 22.0"
                   count_per_type   = "2 | 4"
        """
        loc  = place.get("location", {})
        ev   = place.get("evChargeOptions", {})
        aggs = ev.get("connectorAggregation", [])

        # Build per-connector lists
        connector_types, max_kws, counts = [], [], []
        for agg in aggs:
            raw_type = agg.get("type", "EV_CONNECTOR_TYPE_UNSPECIFIED")
            connector_types.append(CONNECTOR_LABELS.get(raw_type, raw_type))
            max_kws.append(agg.get("maxChargeRateKw", ""))
            counts.append(agg.get("count", ""))

        # Highest kW across all connectors at this station
        numeric_kws = [float(k) for k in max_kws if k != ""]
        overall_max_kw = max(numeric_kws) if numeric_kws else ""

        # Map raw business status to readable label
        raw_status = place.get("businessStatus", "")
        status = BUSINESS_STATUS_MAP.get(raw_status, raw_status)

        # Opening hours as a single readable string
        hours_list = place.get("regularOpeningHours", {}).get("weekdayDescriptions", [])
        hours_str  = " | ".join(hours_list)

        return {
            "place_id":          place.get("id", ""),
            "name":              place.get("displayName", {}).get("text", ""),
            "address":           place.get("formattedAddress", ""),
            "lat":               loc.get("latitude", ""),
            "lng":               loc.get("longitude", ""),
            "status":            status,
            "rating":            place.get("rating", ""),
            "user_rating_count": place.get("userRatingCount", ""),
            "total_connectors":  ev.get("connectorCount", ""),
            "connector_types":   " | ".join(connector_types) if connector_types else "",
            "max_kw_per_type":   " | ".join(str(k) for k in max_kws) if max_kws else "",
            "count_per_type":    " | ".join(str(c) for c in counts) if counts else "",
            "overall_max_kw":    overall_max_kw,
            "opening_hours":     hours_str,
        }

    # ── Deduplication ─────────────────────────────────────────────────────────
    @staticmethod
    def deduplicate(records: list[dict]) -> tuple[list[dict], int]:
        """
        Remove duplicate stations by place_id.
        The grid search intentionally overlaps, so duplicates are expected and normal.
        Returns (unique_records, count_removed).
        """
        seen: dict[str, dict] = {}
        for r in records:
            pid = r.get("place_id")
            if pid and pid not in seen:
                seen[pid] = r
        dupes_removed = len(records) - len(seen)
        return list(seen.values()), dupes_removed

    # ── Full scrape pipeline ───────────────────────────────────────────────────
    def scrape(self) -> list[dict]:
        """
        Run the full scrape:
          1. Grid search  — systematic 1 km coverage of bounding box
          2. Text search  — supplemental sweep for edge cases
          3. Parse        — flatten API objects to dicts
          4. Deduplicate  — remove overlapping grid hits
        Returns a list of unique station dicts ready for DataFrame conversion.
        """
        all_raw: list[dict] = []

        # ① Grid search
        grid = self.build_grid(BOUNDING_BOX, GRID_SPACING_KM)
        log.info(f"Starting grid search across {len(grid)} points …")

        for i, (lat, lng) in enumerate(grid, 1):
            batch = self.nearby_search(lat, lng)
            all_raw.extend(batch)

            # Progress log every 200 points
            if i % 200 == 0 or i == len(grid):
                log.info(f"  Grid progress: {i}/{len(grid)} points  |  raw results so far: {len(all_raw)}")

        log.info(f"Grid search complete: {len(all_raw)} raw results from {len(grid)} grid points")

        # ② Text search sweep
        log.info("Running supplemental text search …")
        text_results = self.text_search("EV charging station Chennai")
        all_raw.extend(text_results)
        log.info(f"Text search added {len(text_results)} additional raw results")

        # ③ Parse all raw place objects
        log.info("Parsing place data …")
        parsed = [self.parse_place(p) for p in all_raw]

        # ④ Deduplicate
        unique, dupes_removed = self.deduplicate(parsed)
        log.info(
            f"Deduplication: {len(all_raw)} raw  →  {len(unique)} unique  "
            f"({dupes_removed} duplicates removed)"
        )

        return unique


# ── ENTRY POINT ─────────────────────────────────────────────────────────────────
def main():
    # Use env var if set, otherwise fall back to the hardcoded key above
    api_key = os.environ.get("GOOGLE_API_KEY", API_KEY)
    if not api_key:
        raise SystemExit("\n[ERROR] No API key found. Set API_KEY at the top of this file.\n")

    log.info("=" * 60)
    log.info("EV Charging Station Scraper — Chennai")
    log.info(f"Bounding box : {BOUNDING_BOX}")
    log.info(f"Grid spacing : {GRID_SPACING_KM} km  |  Radius: {SEARCH_RADIUS_M} m")
    log.info(f"Rate limit   : {REQUESTS_PER_SEC} requests/second")
    log.info("=" * 60)

    scraper = EVChargingScraper(api_key)
    records = scraper.scrape()

    if not records:
        log.warning("No records returned. Verify your API key, billing, and quota.")
        return

    # Build DataFrame with a fixed column order
    df = pd.DataFrame(records)
    col_order = [
        "place_id", "name", "address", "lat", "lng",
        "status", "rating", "user_rating_count",
        "total_connectors", "connector_types",
        "max_kw_per_type", "count_per_type", "overall_max_kw",
        "opening_hours",
    ]
    df = df[[c for c in col_order if c in df.columns]]

    # Save Excel to the same folder as this script
    out_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), OUTPUT_FILE)
    df.to_excel(out_path, index=False, engine="openpyxl")

    log.info("=" * 60)
    log.info(f"DONE")
    log.info(f"  Total unique EV stations : {len(df)}")
    log.info(f"  Saved to                 : {out_path}")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
