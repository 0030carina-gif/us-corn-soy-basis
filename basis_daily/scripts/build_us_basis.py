#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""US soybean & corn basis daily tracker from public USDA AMS grain PDFs."""
from __future__ import annotations

import csv
import json
import re
import subprocess
import urllib.request
from collections import defaultdict
from datetime import datetime, timezone, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent
CACHE = ROOT / "_cache"
DATA = ROOT / "data"
CHARTS = ROOT / "charts"
UA = "Mozilla/5.0 us-basis/1.0 (public AMS PDF reader; contact via GitHub)"

# Report catalog: slug -> (state_code, default_region_label, notes)
REPORTS = {
    "3192": {"state": "IL", "label": "Illinois", "channels": "all"},
    "2850": {"state": "IA", "label": "Iowa", "channels": "country/mills/terminal"},
    "3043": {"state": "IA-MN", "label": "Iowa-S.Minnesota Barge", "channels": "barge"},
    "3147": {"state": "LA", "label": "Louisiana/Texas Gulf Export", "channels": "export"},
    "3225": {"state": "NE", "label": "Nebraska", "channels": "country"},
    "3463": {"state": "IN", "label": "Indiana", "channels": "all"},
    "2851": {"state": "OH", "label": "Ohio", "channels": "all"},
    "2932": {"state": "MO", "label": "Missouri", "channels": "all"},
    "3049": {"state": "MN", "label": "Southern Minnesota", "channels": "country"},
    "2886": {"state": "KS", "label": "Kansas", "channels": "country"},
    "2892": {"state": "KY", "label": "Kentucky", "channels": "all"},
    "2960": {"state": "AR", "label": "Arkansas", "channels": "all"},
    "3100": {"state": "OK", "label": "Oklahoma", "channels": "all"},
    "2912": {"state": "CO", "label": "Colorado", "channels": "country"},
    "3088": {"state": "TN", "label": "Tennessee", "channels": "all"},
    "3156": {"state": "NC", "label": "North Carolina", "channels": "all"},
    "3167": {"state": "VA", "label": "Virginia", "channels": "all"},
    "3186": {"state": "SD", "label": "South Dakota", "channels": "country"},
    "3239": {"state": "WY", "label": "Wyoming", "channels": "country"},
    "2928": {"state": "MS", "label": "Mississippi", "channels": "country"},
    "3148": {"state": "OR", "label": "Portland/Oregon", "channels": "export"},
    "3878": {"state": "ND", "label": "North Dakota", "channels": "country"},
    "3046": {"state": "MN", "label": "Minneapolis Terminal", "channels": "terminal"},
}

CHANNEL_MAP = [
    (re.compile(r"Country Elevators", re.I), "Country Elevators"),
    (re.compile(r"Barge Loading Elevators|Barge Terminal", re.I), "Barge"),
    (re.compile(r"Export Elevators|Gulf Coast Ports", re.I), "Gulf/Export"),
    (re.compile(r"Mills and Processors|Terminals/Mills/Processors", re.I), "Mills & Processors"),
    (re.compile(r"Terminal Elevators", re.I), "Terminal Elevators"),
]

COMMODITY_RE = re.compile(
    r"US\s*#?\s*\d+(?:/\d+)?\s+(?:Yellow\s+|White\s+)?(Corn|Soybeans)\b",
    re.I,
)
# Any other US #N … (Bulk) grade header clears corn/soy context (wheat, oats, sorghum, …)
SKIP_COMMODITY = re.compile(
    r"US\s*#?\s*\d+(?:/\d+)?\s+(?:(?!Yellow\s+Corn|White\s+Corn|Corn\b|Soybeans).)+",
    re.I,
)
US_BULK_HEADER_RE = re.compile(r"US\s*#?\s*\d+(?:/\d+)?\s+.+$", re.I)

REPORT_DATE_RE = re.compile(
    r"Grain Report for\s+(\d{1,2}/\d{1,2}/\d{4})",
    re.I,
)
# Fallback: "October 19, 2026" in header
HEADER_DATE_RE = re.compile(
    r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),\s+(\d{4})",
    re.I,
)

# Basis: -43.00Z to -30.00Z  OR  74.00Z to 82.00Z  OR  single 0.00Z / -25.00Z
BASIS_RANGE_RE = re.compile(
    r"(-?\d+\.\d{2})([FGHJKMNQUVXZ])\s+to\s+(-?\d+\.\d{2})([FGHJKMNQUVXZ])",
    re.I,
)
BASIS_SINGLE_RE = re.compile(
    r"(?<![.\d])(-?\d+\.\d{2})([FGHJKMNQUVXZ])(?!\s+to)",
    re.I,
)
PRICE_AVG_RE = re.compile(r"(\d+\.\d{4})\s+(\d+\.\d{4})?")  # average year-ago col often two floats
DELIVERY_TOKENS = ("Current", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

# Approximate state centroids for SVG map (lon, lat) — used as x,y in simple projection
STATE_XY = {
    "MN": (220, 95),
    "IA": (250, 155),
    "NE": (175, 160),
    "IL": (295, 175),
    "IN": (340, 175),
    "OH": (385, 170),
    "MO": (255, 215),
    "KS": (185, 210),
    "LA": (270, 310),  # Gulf marker
}


def shanghai_now() -> datetime:
    return datetime.now(timezone(timedelta(hours=8)))


def _http_get(url: str, timeout: int = 60) -> bytes:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _filerepo_latest_pdf_url(slug: str) -> str | None:
    """Parse MyMarketNews filerepo listing for the newest PDF for this slug."""
    list_url = f"https://mymarketnews.ams.usda.gov/filerepo/reports?field_slug_id_value={slug}"
    try:
        html = _http_get(list_url, timeout=90).decode("utf-8", errors="replace")
    except Exception as e:
        print(f"  filerepo list failed for {slug}: {e}")
        return None
    # Collect candidate PDF / download hrefs (prefer absolute or site-relative)
    hrefs = re.findall(r'href=["\']([^"\']+)["\']', html, flags=re.I)
    candidates: list[str] = []
    for h in hrefs:
        low = h.lower()
        if ".pdf" in low or "/filerepo/download" in low or "/download/" in low:
            if h.startswith("//"):
                h = "https:" + h
            elif h.startswith("/"):
                h = "https://mymarketnews.ams.usda.gov" + h
            elif not h.startswith("http"):
                continue
            candidates.append(h)
    # Also try AMS-style links embedded in the page
    for m in re.finditer(rf"https?://[^\"\s]+ams_{slug}\.pdf", html, flags=re.I):
        candidates.insert(0, m.group(0))
    # De-dupe preserve order
    seen = set()
    out = []
    for c in candidates:
        if c not in seen:
            seen.add(c)
            out.append(c)
    return out[0] if out else None


def fetch_pdf(slug: str, force: bool = False) -> Path:
    CACHE.mkdir(parents=True, exist_ok=True)
    dest = CACHE / f"ams_{slug}.pdf"
    if not force and dest.exists() and dest.stat().st_size > 5000:
        # Re-fetch if older than 6 hours (daily update)
        age_h = (datetime.now().timestamp() - dest.stat().st_mtime) / 3600
        if age_h < 6:
            return dest
    primary = f"https://www.ams.usda.gov/mnreports/ams_{slug}.pdf"
    data = b""
    err = None
    try:
        data = _http_get(primary, timeout=60)
    except Exception as e:
        err = e
        data = b""
    if len(data) < 5000:
        print(f"  primary PDF failed/small for {slug} ({len(data)} bytes; err={err}); trying filerepo…")
        alt = _filerepo_latest_pdf_url(slug)
        if not alt:
            raise RuntimeError(f"PDF fetch failed for {slug}: primary err={err}; no filerepo link")
        data = _http_get(alt, timeout=90)
        if len(data) < 5000:
            raise RuntimeError(f"PDF too small for {slug} via filerepo: {len(data)} bytes ({alt})")
        print(f"  filerepo OK for {slug}: {alt}")
    dest.write_bytes(data)
    return dest


def pdf_to_text(pdf_path: Path) -> str:
    txt_path = pdf_path.with_suffix(".txt")
    subprocess.run(
        ["pdftotext", "-layout", str(pdf_path), str(txt_path)],
        check=True,
        capture_output=True,
    )
    return txt_path.read_text(encoding="utf-8", errors="replace")


def parse_report_date(text: str) -> str | None:
    m = REPORT_DATE_RE.search(text)
    if m:
        dt = datetime.strptime(m.group(1), "%m/%d/%Y")
        return dt.strftime("%Y-%m-%d")
    m = HEADER_DATE_RE.search(text)
    if m:
        dt = datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%B %d %Y")
        return dt.strftime("%Y-%m-%d")
    return None


def map_channel(header: str) -> str | None:
    for pat, name in CHANNEL_MAP:
        if pat.search(header):
            return name
    return None


def normalize_location(parts: list[str]) -> str:
    s = " ".join(p.strip() for p in parts if p.strip())
    s = re.sub(r"\s+", " ", s).strip(" -")
    # Drop stray state-code fragments left by page headers (e.g. "KY Ohio River - Lower")
    s = re.sub(r"^[A-Z]{2}\s+(?=[A-Z])", "", s)
    return s


def is_continuation_line(line: str) -> bool:
    """True if line looks like a wrapped location fragment (no Bid, no Basis)."""
    s = line.strip()
    if not s:
        return False
    if re.search(r"(?<![A-Za-z/])Bid(?![A-Za-z/])", s):
        return False
    if BASIS_RANGE_RE.search(s) or BASIS_SINGLE_RE.search(s):
        return False
    if re.search(r"Source:|Explanatory|Futures Settlements|Region/Location|Email us", s, re.I):
        return False
    if COMMODITY_RE.search(s) or SKIP_COMMODITY.search(s):
        return False
    if map_channel(s):
        return False
    # short-ish location remnants
    if len(s) < 40 and not re.search(r"\d+\.\d{2}", s):
        return True
    return False


def extract_delivery(line: str) -> str:
    # Prefer token near end of line
    for tok in ("Current", "Nov¹", "Nov²", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct"):
        if re.search(rf"\b{re.escape(tok)}\b", line):
            return tok.replace("¹", "").replace("²", "")
    # month ranges like "Jun - Jul"
    m = re.search(r"\b([A-Z][a-z]{2})\s*-\s*([A-Z][a-z]{2})\b", line)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return ""


def parse_basis_from_line(line: str) -> tuple[float | None, float | None, str]:
    m = BASIS_RANGE_RE.search(line)
    if m:
        lo, mo1, hi, mo2 = float(m.group(1)), m.group(2).upper(), float(m.group(3)), m.group(4).upper()
        return lo, hi, mo1  # use first month letter
    # single value — find first basis-like after Bid
    # Avoid matching price change like DN 0.0725
    after_bid = line
    if "Bid" in line:
        after_bid = line.split("Bid", 1)[1]
    # Take first occurrence of N.NN + month letter
    m = re.search(r"(-?\d+\.\d{2})([FGHJKMNQUVXZ])\b", after_bid)
    if m:
        v = float(m.group(1))
        return v, v, m.group(2).upper()
    return None, None, ""


# Basis Change column tokens (¢/bu). Uses exactly 2 decimals so Price Change
# values like "DN 0.0725" ($/bu, 4 decimals) are never matched.
BASIS_CHANGE_RE = re.compile(
    r"UNCH-(?:UP|DN)\s+\d+\.\d{2}(?!\d)"
    r"|(?:UP|DN)\s+\d+\.\d{2}(?!\d)-UNCH"
    r"|(?:UP|DN)\s+\d+\.\d{2}(?!\d)-(?:UP|DN)\s+\d+\.\d{2}(?!\d)"
    r"|(?:UP|DN)\s+\d+\.\d{2}(?!\d)"
    r"|UNCH",
    re.I,
)


def _signed_change(direction: str, amount: str) -> float:
    v = float(amount)
    return v if direction.upper() == "UP" else -v


def parse_basis_change_token(token: str) -> tuple[float | None, float | None, float | None]:
    """Return (low, high, mid) in ¢/bu for an AMS Basis Change token."""
    token = re.sub(r"\s+", " ", token.strip().upper())
    if token == "UNCH":
        return 0.0, 0.0, 0.0
    m = re.fullmatch(r"UNCH-(UP|DN) (\d+\.\d{2})", token)
    if m:
        v = _signed_change(m.group(1), m.group(2))
        lo, hi = min(0.0, v), max(0.0, v)
        return lo, hi, round((lo + hi) / 2.0, 2)
    m = re.fullmatch(r"(UP|DN) (\d+\.\d{2})-UNCH", token)
    if m:
        v = _signed_change(m.group(1), m.group(2))
        lo, hi = min(0.0, v), max(0.0, v)
        return lo, hi, round((lo + hi) / 2.0, 2)
    m = re.fullmatch(r"(UP|DN) (\d+\.\d{2})-(UP|DN) (\d+\.\d{2})", token)
    if m:
        v1 = _signed_change(m.group(1), m.group(2))
        v2 = _signed_change(m.group(3), m.group(4))
        lo, hi = min(v1, v2), max(v1, v2)
        return lo, hi, round((lo + hi) / 2.0, 2)
    m = re.fullmatch(r"(UP|DN) (\d+\.\d{2})", token)
    if m:
        v = _signed_change(m.group(1), m.group(2))
        return v, v, v
    return None, None, None


def parse_basis_change_from_line(line: str) -> tuple[float | None, float | None, float | None, str]:
    """Extract AMS Basis Change (not Price Change) from a Bid line.

    Returns (change_low, change_high, change_mid, raw_token).
    """
    after = line
    if "Bid" in line:
        after = line.split("Bid", 1)[1]
    # Prefer the token that sits after the basis month-letter cluster and before $/Bu prices.
    # Search the whole after-Bid segment; 2-decimal UP/DN/UNCH never equals 4-decimal Price Change.
    m = BASIS_CHANGE_RE.search(after)
    if not m:
        return None, None, None, ""
    raw = m.group(0)
    lo, hi, mid = parse_basis_change_token(raw)
    return lo, hi, mid, raw


def parse_price_avg(line: str) -> float | None:
    """Pick Average column roughly: after Price Change, before Freight/Delivery."""
    # Look for pattern of X.XXXX that appears in Average Year Ago area
    # Safer: find all X.XXXX floats; average of cash price range is often earlier.
    # Prefer the float just before year-ago and freight.
    floats = re.findall(r"\b(\d+\.\d{4})\b", line)
    if not floats:
        return None
    # Heuristic: if price range exists (two early floats), Average is often the next
    # For single price, Average is often that same or next.
    # Common layout: price_low-price_high  change  average yearago
    # We'll take the float that sits after price change tokens.
    after = line
    for marker in ("DN ", "UP ", "UNCH"):
        if marker in after:
            idx = after.rfind(marker) if marker != "UNCH" else after.find("UNCH")
            # better: find last price-change cluster then next float
    # Simpler: if >= 3 floats of form d.dddd, take the one at index -2 or mid
    if len(floats) >= 3:
        # typically: lo, hi, avg, yearago  OR price, avg, yearago
        return float(floats[-2]) if len(floats) >= 4 else float(floats[1] if len(floats) == 3 else floats[0])
    if len(floats) == 2:
        return float(floats[0])  # single price + year ago → price is avg
    if len(floats) == 1:
        return float(floats[0])
    return None


def parse_pdf_text(text: str, state: str, report_id: str) -> list[dict]:
    report_date = parse_report_date(text)
    if not report_date:
        return []

    lines = text.splitlines()
    commodity = None
    channel = None
    rows: list[dict] = []
    pending_loc: list[str] = []

    i = 0
    while i < len(lines):
        line = lines[i]
        raw = line.rstrip()
        stripped = raw.strip()

        # page noise
        if re.search(r"^Source:|Email us with accessibility|General inquiries|http://|https://|Page \d+ of", stripped):
            i += 1
            continue
        if stripped.startswith("AMS Livestock") or "Dept. of Ag Market News" in stripped:
            i += 1
            continue
        # title lines like "Illinois Grain Bids"
        if re.match(r"^[A-Za-z].*Grain Bids", stripped) and "Bid" not in stripped:
            i += 1
            continue

        cm = COMMODITY_RE.search(stripped)
        if cm:
            commodity = "Corn" if cm.group(1).lower() == "corn" else "Soybeans"
            pending_loc = []
            i += 1
            continue
        if SKIP_COMMODITY.search(stripped):
            commodity = None  # pause until next corn/soy
            pending_loc = []
            i += 1
            continue

        ch = map_channel(stripped)
        if ch and "Region/Location" not in stripped and not re.search(r"(?<![A-Za-z/])Bid(?![A-Za-z/])", stripped):
            channel = ch
            pending_loc = []
            i += 1
            continue

        if commodity is None or channel is None:
            i += 1
            continue

        # Skip header row
        if "Region/Location" in stripped or "Basis (¢/Bu)" in stripped:
            i += 1
            continue

        # Accumulate location continuation before a Bid line
        if is_continuation_line(raw) and "Bid" not in raw:
            pending_loc.append(stripped)
            i += 1
            continue

        # Require standalone sale-type Bid (not Bid/Ask/Tra)
        if not re.search(r"(?<![A-Za-z/])Bid(?![A-Za-z/])", stripped):
            i += 1
            continue

        # Possibly merge previous pending location with this line's location prefix
        # Line form: "  North Central          Bid           -43.00Z ..."
        m_loc = re.match(r"^(.+?)\s+Bid(?![A-Za-z/])", stripped)
        loc_here = m_loc.group(1).strip() if m_loc else ""
        location = normalize_location(pending_loc + ([loc_here] if loc_here else []))
        pending_loc = []

        # If location empty but next lines are continuations that already passed — rare
        if not location or location == "Unknown":
            location = "(statewide)"

        lo, hi, month = parse_basis_from_line(stripped)
        if lo is None:
            i += 1
            continue

        delivery = extract_delivery(stripped)
        price_avg = parse_price_avg(stripped)
        mid = round((lo + hi) / 2.0, 2)
        ch_lo, ch_hi, ch_mid, ch_raw = parse_basis_change_from_line(stripped)

        # Minneapolis Terminal rows often have blank Region/Location
        if (not location or location == "(statewide)") and report_id == "3046":
            location = "Minneapolis"

        rows.append(
            {
                "report_date": report_date,
                "report_id": report_id,
                "state": state,
                "location": location,
                "commodity": commodity,
                "channel": channel,
                "basis_low": lo,
                "basis_high": hi,
                "basis_mid": mid,
                "basis_change_low": ch_lo if ch_lo is not None else "",
                "basis_change_high": ch_hi if ch_hi is not None else "",
                "basis_change_mid": ch_mid if ch_mid is not None else "",
                "basis_change_raw": ch_raw,
                "basis_unit": "cents/bu",
                "futures_month": month,
                "price_avg": price_avg if price_avg is not None else "",
                "delivery": delivery or "Current",
            }
        )
        i += 1

    return rows


def download_and_parse(force: bool = False) -> list[dict]:
    all_rows: list[dict] = []
    for slug, meta in REPORTS.items():
        try:
            pdf = fetch_pdf(slug, force=force)
            text = pdf_to_text(pdf)
            rows = parse_pdf_text(text, meta["state"], slug)
            print(f"  {slug} {meta['label']}: {len(rows)} rows, date={rows[0]['report_date'] if rows else '?'}")
            all_rows.extend(rows)
        except Exception as e:
            print(f"  WARN {slug} {meta['label']}: {e}")
    return all_rows


def row_key(r: dict) -> tuple:
    return (
        r["report_date"],
        r["state"],
        r["location"],
        r["commodity"],
        r["channel"],
        r["delivery"],
        r["futures_month"],
    )


FIELDNAMES = [
    "report_date",
    "report_id",
    "state",
    "location",
    "commodity",
    "channel",
    "basis_low",
    "basis_high",
    "basis_mid",
    "basis_change_low",
    "basis_change_high",
    "basis_change_mid",
    "basis_change_raw",
    "basis_unit",
    "futures_month",
    "price_avg",
    "delivery",
]


def read_csv(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def write_csv(path: Path, rows: list[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=FIELDNAMES)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in FIELDNAMES})


def merge_history(new_rows: list[dict]) -> list[dict]:
    hist_path = DATA / "basis_history.csv"
    existing = read_csv(hist_path)
    by_key = {row_key(r): r for r in existing}
    for r in new_rows:
        by_key[row_key(r)] = r
    merged = list(by_key.values())
    merged.sort(key=lambda r: (r["report_date"], r["state"], r["commodity"], r["channel"], r["location"]))
    return merged


def filter_current(rows: list[dict]) -> list[dict]:
    return [r for r in rows if (r.get("delivery") or "Current") == "Current"]


def avg_mid(rows: list[dict]) -> float | None:
    if not rows:
        return None
    return round(sum(float(r["basis_mid"]) for r in rows) / len(rows), 2)


def build_summary(latest_date: str, current_rows: list[dict], history: list[dict]) -> dict:
    # prior date with any current rows
    dates = sorted({r["report_date"] for r in history if (r.get("delivery") or "Current") == "Current"})
    prior = None
    for d in reversed(dates):
        if d < latest_date:
            prior = d
            break

    def slice_rows(rows, date, state=None, commodity=None, channel=None):
        out = []
        for r in rows:
            if r["report_date"] != date:
                continue
            if (r.get("delivery") or "Current") != "Current":
                continue
            if state and r["state"] != state:
                continue
            if commodity and r["commodity"] != commodity:
                continue
            if channel and r["channel"] != channel:
                continue
            out.append(r)
        return out

    keys = [
        ("IL", "Corn", "Country Elevators", "il_country_corn"),
        ("IL", "Soybeans", "Country Elevators", "il_country_soy"),
        ("IL", "Corn", "Mills & Processors", "il_mills_corn"),
        ("IL", "Soybeans", "Mills & Processors", "il_mills_soy"),
        ("IL", "Corn", "Barge", "il_barge_corn"),
        ("IL", "Soybeans", "Barge", "il_barge_soy"),
        ("IA", "Corn", "Country Elevators", "ia_country_corn"),
        ("IA", "Soybeans", "Country Elevators", "ia_country_soy"),
        ("IA", "Soybeans", "Mills & Processors", "ia_mills_soy"),
        ("IA-MN", "Corn", "Barge", "ia_mn_barge_corn"),
        ("IA-MN", "Soybeans", "Barge", "ia_mn_barge_soy"),
        ("LA", "Corn", "Gulf/Export", "gulf_export_corn"),
        ("LA", "Soybeans", "Gulf/Export", "gulf_export_soy"),
        ("NE", "Corn", "Country Elevators", "ne_country_corn"),
        ("NE", "Soybeans", "Country Elevators", "ne_country_soy"),
        ("IN", "Corn", "Country Elevators", "in_country_corn"),
        ("IN", "Soybeans", "Country Elevators", "in_country_soy"),
        ("OH", "Corn", "Country Elevators", "oh_country_corn"),
        ("OH", "Soybeans", "Country Elevators", "oh_country_soy"),
        ("MO", "Corn", "Country Elevators", "mo_country_corn"),
        ("MO", "Soybeans", "Country Elevators", "mo_country_soy"),
        ("MN", "Corn", "Country Elevators", "mn_country_corn"),
        ("MN", "Soybeans", "Country Elevators", "mn_country_soy"),
        ("KS", "Corn", "Country Elevators", "ks_country_corn"),
        ("KS", "Soybeans", "Country Elevators", "ks_country_soy"),
    ]

    series = {}
    for state, commodity, channel, key in keys:
        cur = avg_mid(slice_rows(history, latest_date, state, commodity, channel))
        prv = avg_mid(slice_rows(history, prior, state, commodity, channel)) if prior else None
        delta = round(cur - prv, 2) if (cur is not None and prv is not None) else None
        series[key] = {
            "state": state,
            "commodity": commodity,
            "channel": channel,
            "basis_mid": cur,
            "prior_mid": prv,
            "delta": delta,
            "unit": "cents/bu",
        }

    # Channel × commodity overview (all states pooled mid of mids — prefer IL+IA+Gulf)
    overview = {}
    for commodity in ("Corn", "Soybeans"):
        for channel in ("Country Elevators", "Barge", "Gulf/Export", "Mills & Processors"):
            rs = slice_rows(history, latest_date, None, commodity, channel)
            overview[f"{commodity}|{channel}"] = avg_mid(rs)

    # State country-elevator mids for map
    map_states = {}
    for st in ("IL", "IA", "NE", "MN", "IN", "OH", "MO", "KS", "ND", "SD", "KY", "AR"):
        map_states[st] = {
            "corn": avg_mid(slice_rows(history, latest_date, st, "Corn", "Country Elevators")),
            "soy": avg_mid(slice_rows(history, latest_date, st, "Soybeans", "Country Elevators")),
        }
    map_states["LA"] = {
        "corn": avg_mid(slice_rows(history, latest_date, "LA", "Corn", "Gulf/Export")),
        "soy": avg_mid(slice_rows(history, latest_date, "LA", "Soybeans", "Gulf/Export")),
    }

    # History series for charts (daily avg of key series)
    chart_keys = [
        "il_country_corn",
        "il_country_soy",
        "gulf_export_corn",
        "gulf_export_soy",
    ]
    chart_history = {k: [] for k in chart_keys}
    for d in dates:
        for state, commodity, channel, key in keys:
            if key not in chart_keys:
                continue
            v = avg_mid(slice_rows(history, d, state, commodity, channel))
            if v is not None:
                chart_history[key].append({"date": d, "value": v})

    counts: dict[str, int] = defaultdict(int)
    for r in current_rows:
        if r["report_date"] == latest_date:
            counts[f"{r['commodity']}|{r['channel']}"] += 1

    return {
        "report_date": latest_date,
        "prior_date": prior,
        "updated_at": shanghai_now().strftime("%Y-%m-%d %H:%M:%S CST"),
        "series": series,
        "overview": overview,
        "map_states": map_states,
        "chart_history": chart_history,
        "counts": dict(counts),
        "row_count_current": sum(1 for r in current_rows if r["report_date"] == latest_date),
    }


def color_for_basis(v: float | None, commodity: str) -> str:
    """Map basis mid (cents) to a fill color. Corn country typically negative; gulf positive."""
    if v is None:
        return "#d0d5dd"
    # Diverging: more negative = cooler blue, higher = warmer amber/green
    # Typical country: -80 .. 0 ; gulf: +50 .. +120
    if commodity == "soy":
        # soy country often -90..-20
        t = (v + 100) / 120  # -100→0, +20→1
    else:
        t = (v + 80) / 120  # -80→0, +40→1
    t = max(0.0, min(1.0, t))
    # blue (#3b82f6) -> gray -> amber (#f59e0b)
    def lerp(a, b, x):
        return int(a + (b - a) * x)

    if t < 0.5:
        u = t * 2
        r, g, b = lerp(59, 160, u), lerp(130, 170, u), lerp(246, 180, u)
    else:
        u = (t - 0.5) * 2
        r, g, b = lerp(160, 245, u), lerp(170, 158, u), lerp(180, 11, u)
    return f"#{r:02x}{g:02x}{b:02x}"


def fmt_basis(v) -> str:
    if v is None or v == "":
        return "—"
    try:
        x = float(v)
    except (TypeError, ValueError):
        return "—"
    sign = "+" if x > 0 else ""
    return f"{sign}{x:.1f}"



def build_map_svg(summary: dict) -> str:
    """Homepage overview map SVG (480×380 viewBox)."""
    map_states = summary["map_states"]
    circles = []
    for st, xy in STATE_XY.items():
        ms = map_states.get(st, {})
        corn_c = color_for_basis(ms.get("corn"), "corn")
        soy_c = color_for_basis(ms.get("soy"), "soy")
        cx, cy = xy
        tip = f"{st} corn {fmt_basis(ms.get('corn'))} / soy {fmt_basis(ms.get('soy'))}"
        circles.append(
            f'<g transform="translate({cx},{cy})">'
            f'<title>{tip}</title>'
            f'<path d="M0,-18 A18,18 0 0,0 0,18 Z" fill="{corn_c}" stroke="#334" stroke-width="1"/>'
            f'<path d="M0,-18 A18,18 0 0,1 0,18 Z" fill="{soy_c}" stroke="#334" stroke-width="1"/>'
            f'<text y="32" text-anchor="middle" font-size="12" font-weight="700" fill="#1f2937">{st}</text>'
            f'<text y="46" text-anchor="middle" font-size="9" fill="#4b5563">C {fmt_basis(ms.get("corn"))}</text>'
            f'<text y="58" text-anchor="middle" font-size="9" fill="#4b5563">S {fmt_basis(ms.get("soy"))}</text>'
            f"</g>"
        )
    return f"""
<svg viewBox="0 0 480 380" width="100%" style="max-width:560px;background:#f8fafc;border-radius:12px;border:1px solid #e5e7eb">
  <text x="12" y="22" font-size="13" fill="#374151">产地基差地图 Country Elevators mid · 左半=玉米 右半=大豆 · LA=Gulf Export</text>
  {''.join(circles)}
  <text x="12" y="370" font-size="10" fill="#6b7280">颜色越蓝越弱(更贴水)；越黄越高(升水)。单位 ¢/bu vs CBOT。</text>
</svg>"""


def export_map_image(summary: dict) -> Path:
    """Screenshot real index.html homepage → charts/homepage.png + basis_map.png.

    Optionally also write charts/basis_map.svg (map-only vector). The chat
    attachment must be the full homepage screenshot, not a map-only render page.
    """
    import shutil

    CHARTS.mkdir(parents=True, exist_ok=True)

    # Optional SVG export of the map widget only (not used for chat attachment)
    map_svg = build_map_svg(summary)
    svg_path = CHARTS / "basis_map.svg"
    inner = map_svg.strip()
    inner_body = inner[inner.find(">") + 1 :].rsplit("</svg>", 1)[0] if inner.startswith("<svg") else inner
    standalone = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 480 380" '
        'width="480" height="380" style="background:#ffffff">\n'
        f"{inner_body}\n</svg>\n"
    )
    svg_path.write_text(standalone, encoding="utf-8")

    # Drop legacy map-only render page if present
    legacy = CHARTS / "_map_render.html"
    if legacy.exists():
        try:
            legacy.unlink()
        except OSError:
            pass

    index_path = ROOT / "index.html"
    if not index_path.exists():
        print("WARN: index.html missing — cannot screenshot homepage")
        return CHARTS / "basis_map.png"

    homepage_png = CHARTS / "homepage.png"
    basis_png = CHARTS / "basis_map.png"
    chrome = Path("/usr/bin/google-chrome")
    if not chrome.exists():
        chrome = Path("/usr/bin/chromium")
    cmd = [
        str(chrome),
        "--headless=new",
        "--disable-gpu",
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--hide-scrollbars",
        "--window-size=1100,1400",
        "--virtual-time-budget=8000",
        f"--screenshot={homepage_png}",
        index_path.resolve().as_uri(),
    ]
    try:
        subprocess.run(cmd, check=True, capture_output=True, timeout=90)
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired) as e:
        print(f"WARN: homepage PNG export failed: {e}")
        return basis_png

    if homepage_png.exists():
        shutil.copy2(homepage_png, basis_png)
        print(
            f"Wrote homepage screenshot {homepage_png} "
            f"({homepage_png.stat().st_size} bytes) and copied to {basis_png}"
        )
    else:
        print("WARN: Chrome ran but homepage.png was not created")
    return basis_png


def generate_html(summary: dict) -> str:
    rd = summary["report_date"]
    series = summary["series"]
    charts = summary["chart_history"]
    overview = summary["overview"]

    # Summary table rows
    table_specs = [
        ("Country Elevators 产地", "Country Elevators"),
        ("Barge 驳船", "Barge"),
        ("Gulf/Export 港口", "Gulf/Export"),
        ("Mills & Processors 压榨厂", "Mills & Processors"),
    ]
    table_rows = []
    for label, ch in table_specs:
        corn = overview.get(f"Corn|{ch}")
        soy = overview.get(f"Soybeans|{ch}")
        table_rows.append(
            f"<tr><td>{label}</td><td class='num'>{fmt_basis(corn)}</td>"
            f"<td class='num'>{fmt_basis(soy)}</td></tr>"
        )

    # Key regional highlight cards
    cards = [
        ("IL 产地玉米", "il_country_corn"),
        ("IL 产地大豆", "il_country_soy"),
        ("IL 压榨厂玉米", "il_mills_corn"),
        ("IL 压榨厂大豆", "il_mills_soy"),
        ("IA-MN 驳船玉米", "ia_mn_barge_corn"),
        ("IA-MN 驳船大豆", "ia_mn_barge_soy"),
        ("Gulf 出口玉米", "gulf_export_corn"),
        ("Gulf 出口大豆", "gulf_export_soy"),
    ]
    card_html = []
    for title, key in cards:
        s = series.get(key, {})
        mid = s.get("basis_mid")
        delta = s.get("delta")
        dstr = ""
        if delta is not None:
            arrow = "▲" if delta > 0 else ("▼" if delta < 0 else "→")
            dstr = f"<span class='delta {'up' if delta>0 else 'down' if delta<0 else ''}'>{arrow} {fmt_basis(delta)}</span>"
        card_html.append(
            f"<div class='card'><div class='card-t'>{title}</div>"
            f"<div class='card-v'>{fmt_basis(mid)} <small>¢/bu</small></div>{dstr}</div>"
        )

    map_svg = build_map_svg(summary)

    # Chart.js datasets
    def chart_payload(key, label, color):
        pts = charts.get(key) or []
        return {
            "label": label,
            "color": color,
            "labels": [p["date"] for p in pts],
            "data": [p["value"] for p in pts],
        }

    chart_cfg = [
        chart_payload("il_country_corn", "IL Country Corn", "#2563eb"),
        chart_payload("il_country_soy", "IL Country Soy", "#16a34a"),
        chart_payload("gulf_export_corn", "Gulf Export Corn", "#d97706"),
        chart_payload("gulf_export_soy", "Gulf Export Soy", "#dc2626"),
    ]
    chart_json = json.dumps(chart_cfg, ensure_ascii=False)

    placeholder_note = ""
    if max((len(c["data"]) for c in chart_cfg), default=0) < 2:
        placeholder_note = (
            "<p class='note'>季节性曲线将在多日积累后显现；当前仅有最新一日数据点。"
            " Seasonal charts will fill in as daily history grows.</p>"
        )

    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>美国大豆/玉米基差 US Corn &amp; Soybean Basis</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
:root {{ --bg:#f5f7fb; --card:#fff; --ink:#1f2937; --muted:#6b7280; --line:#e5e7eb; --accent:#0f766e; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; font-family:"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif; background:var(--bg); color:var(--ink); }}
.wrap {{ max-width:1100px; margin:0 auto; padding:24px 18px 48px; }}
h1 {{ font-size:1.6rem; margin:0 0 6px; }}
.sub {{ color:var(--muted); margin:0 0 18px; font-size:0.95rem; }}
.grid {{ display:grid; grid-template-columns:repeat(auto-fill,minmax(160px,1fr)); gap:10px; margin:16px 0 22px; }}
.card {{ background:var(--card); border:1px solid var(--line); border-radius:12px; padding:12px 14px; }}
.card-t {{ font-size:0.78rem; color:var(--muted); }}
.card-v {{ font-size:1.35rem; font-weight:700; margin-top:4px; }}
.card-v small {{ font-size:0.75rem; font-weight:500; color:var(--muted); }}
.delta {{ font-size:0.8rem; color:var(--muted); }}
.delta.up {{ color:#16a34a; }}
.delta.down {{ color:#dc2626; }}
section {{ background:var(--card); border:1px solid var(--line); border-radius:14px; padding:16px 18px; margin:14px 0; }}
h2 {{ font-size:1.05rem; margin:0 0 10px; color:var(--accent); }}
table {{ width:100%; border-collapse:collapse; font-size:0.92rem; }}
th,td {{ border-bottom:1px solid var(--line); padding:8px 6px; text-align:left; }}
th {{ color:var(--muted); font-weight:600; }}
td.num {{ font-variant-numeric:tabular-nums; font-weight:600; }}
.note {{ color:var(--muted); font-size:0.85rem; }}
.charts {{ display:grid; grid-template-columns:1fr 1fr; gap:14px; }}
@media (max-width:800px) {{ .charts {{ grid-template-columns:1fr; }} }}
canvas {{ width:100% !important; max-height:260px; }}
footer {{ margin-top:20px; color:var(--muted); font-size:0.8rem; }}
a {{ color:var(--accent); }}
</style>
</head>
<body>
<div class="wrap">
  <h1>美国大豆 / 玉米基差</h1>
  <p class="sub">US Corn &amp; Soybean Basis · USDA AMS public grain bids · 报告日 Report date: <strong>{rd}</strong>
  · 更新 Updated: {summary.get('updated_at','')}</p>

  <div class="grid">
    {''.join(card_html)}
  </div>

  <section>
    <h2>渠道总览 Summary by channel（当日各区中位基差均值 ¢/bu）</h2>
    <table>
      <thead><tr><th>渠道 Channel</th><th>玉米 Corn</th><th>大豆 Soybeans</th></tr></thead>
      <tbody>
        {''.join(table_rows)}
      </tbody>
    </table>
    <p class="note">单位：美分/蒲式耳 vs CBOT 对应合约月。Current delivery only. Source: AMS mnreports PDFs.</p>
  </section>

  <section>
    <h2>总览地图 Overview map</h2>
    {map_svg}
  </section>

  <section>
    <h2>季节性 / 走势 Seasonal &amp; recent history</h2>
    {placeholder_note}
    <div class="charts">
      <div><canvas id="c1"></canvas></div>
      <div><canvas id="c2"></canvas></div>
      <div><canvas id="c3"></canvas></div>
      <div><canvas id="c4"></canvas></div>
    </div>
  </section>

  <footer>
    数据来源 USDA AMS Livestock, Poultry &amp; Grain Market News 公开 PDF
    （IL 3192 / IA 2850 / IA-MN barge 3043 / Gulf 3147 / NE 3225 / IN 3463 / OH 2851 / MO 2932 / MN 3049 / Minneapolis 3046 / KS 2886 / KY 2892 / ND 3878 / SD 3186 / AR 2960）。
    仅爬取公开页面，无需登录。脚本 <code>build_us_basis.py</code> ·
    <a href="data/basis_latest.csv">basis_latest.csv</a> ·
    <a href="data/summary_latest.json">summary_latest.json</a>
  </footer>
</div>
<script>
const CFGS = {chart_json};
function mkChart(id, cfg) {{
  const el = document.getElementById(id);
  const labels = cfg.labels.length ? cfg.labels : ['{rd}'];
  const data = cfg.data.length ? cfg.data : [null];
  new Chart(el, {{
    type: 'line',
    data: {{
      labels,
      datasets: [{{
        label: cfg.label + ' (¢/bu)',
        data,
        borderColor: cfg.color,
        backgroundColor: cfg.color + '33',
        tension: 0.25,
        spanGaps: true,
        pointRadius: 3,
      }}]
    }},
    options: {{
      plugins: {{ legend: {{ display: true }} }},
      scales: {{
        y: {{ title: {{ display: true, text: 'Basis ¢/bu' }} }}
      }}
    }}
  }});
}}
mkChart('c1', CFGS[0]);
mkChart('c2', CFGS[1]);
mkChart('c3', CFGS[2]);
mkChart('c4', CFGS[3]);
</script>
</body>
</html>
"""


def main(force: bool = False) -> int:
    DATA.mkdir(parents=True, exist_ok=True)
    CACHE.mkdir(parents=True, exist_ok=True)
    print("Downloading / parsing AMS PDFs ...")
    rows = download_and_parse(force=force)
    if not rows:
        print("No rows parsed — aborting without overwrite.")
        return 1

    # Prefer Current delivery for latest snapshot file, but keep all in history
    current = filter_current(rows)
    if not current:
        current = rows  # fallback

    dates = sorted({r["report_date"] for r in current})
    latest_date = dates[-1]
    latest_rows = [r for r in current if r["report_date"] == latest_date]

    write_csv(DATA / "basis_latest.csv", latest_rows)
    history = merge_history(rows)  # all deliveries into history
    write_csv(DATA / "basis_history.csv", history)

    summary = build_summary(latest_date, current, history)
    (DATA / "summary_latest.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    html = generate_html(summary)
    (ROOT / "index.html").write_text(html, encoding="utf-8")
    export_map_image(summary)

    print(f"Latest report_date: {latest_date}")
    print(f"Current-delivery rows on latest day: {len(latest_rows)}")
    print("Counts commodity×channel:")
    for k, v in sorted(summary["counts"].items()):
        print(f"  {k}: {v}")
    print(f"Wrote {DATA/'basis_latest.csv'}, history ({len(history)} rows), summary, index.html")
    return 0


if __name__ == "__main__":
    import sys

    force = "--force" in sys.argv
    raise SystemExit(main(force=force))
