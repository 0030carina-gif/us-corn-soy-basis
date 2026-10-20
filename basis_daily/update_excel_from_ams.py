# -*- coding: utf-8 -*-
"""Append AMS PDF Current-basis mids into local soy/corn Basis.xlsx sheets."""
from __future__ import annotations

import csv
import re
import shutil
from collections import defaultdict
import statistics
from datetime import datetime, date
from pathlib import Path

import openpyxl

# Paths — override via env or defaults for laptop
ROOT = Path(r"D:\grok知识库\excel_s\玉米大豆历史数据")
CSV_PATH = Path(r"C:\Users\98728\basis_daily_sync\basis_latest.csv")
BACKUP_DIR = Path(r"C:\Users\98728\basis_daily_sync\xlsx_backup")
LOG_PATH = Path(r"C:\Users\98728\basis_daily_sync\update_log.txt")

STATE_NAME = {
    "AR": "Arkansas", "CA": "California", "CO": "Colorado", "IL": "Illinois",
    "IN": "Indiana", "IA": "Iowa", "KS": "Kansas", "KY": "Kentucky",
    "LA": "Louisiana", "MD": "Maryland", "MN": "Minnesota", "MS": "Mississippi",
    "MO": "Missouri", "NE": "Nebraska", "NC": "North Carolina", "ND": "North Dakota",
    "OH": "Ohio", "OK": "Oklahoma", "OR": "Oregon", "PA": "Pennsylvania",
    "SC": "South Carolina", "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas",
    "VA": "Virginia", "WY": "Wyoming",
}

# PDF location -> alternate alnum keys for sheet suffix matching
LOC_ALIASES = {
    "northcentral": ["northcentral"],
    "southwest": ["southwest"],
    "southcentral": ["southcentral"],
    "northwest": ["northwest"],
    "northeast": ["northeast"],
    "southeast": ["southeast"],
    "westcentral": ["westcentral"],
    "eastcentral": ["eastcentral"],
    "north": ["north"],
    "south": ["south"],
    "east": ["east"],
    "west": ["west"],
    "central": ["central"],
    "littleegypt": ["littleegypt"],
    "wabash": ["wabash"],
    "chicago": ["chicago"],
    "mississippiriver": ["mississippiriver", "mississip", "mississippi"],
    "northillinoisriver": ["northillinoisriver", "northilli"],
    "southillinoisriver": ["southillinoisriver", "southilli"],
    "gulfcoastports": ["gulfcoastportsla", "gulfcoastports"],
    "ohioriver": ["ohioriver"],
    "toledoonriver": ["toledoonriver"],
    "toledooffriver": ["toledooffriver"],
    "northeastmissouri": ["northeastmissourimsriver", "northeastmissouri"],
    "msriverseutheastmissouri": ["southeastmissourimsriver", "southeastmissouri"],
    "msriversoutheastmissouri": ["southeastmissourimsriver", "southeastmissouri"],
    "stlouismsriver": ["stlouismsriver"],
    "stjoseph": ["stjoseph"],
    "kansascity": ["kansascity"],
    "northiowams": ["northiowamsriver", "northiowams"],
    "riversouthiowams": ["southiowamsriver", "southiowams"],
    "riversouthminnesota": ["southminnesotamsriver", "southminnesota"],
    "statewide": [""],  # empty suffix sheets like MillsProc
    "minneapolis": ["", "minneapolis", "terminals"],
    "ohioriverupper": ["ohioriverupper", "ohioriverupperky"],
    "ohioriverlower": ["ohioriverlower", "ohioriverlowerky"],
    "ohioriver-upper": ["ohioriverupper", "ohioriverupperky"],
    "ohioriver-lower": ["ohioriverlower", "ohioriverlowerky"],
    "pennyrile": ["pennyrile"],
    "bluegrass": ["bluegrass"],
    "greenriver": ["greenriver"],
    "louisville": ["louisville"],
    "purchase": ["purchase"],
}


def alnum(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


def normalize_loc_key(loc: str) -> str:
    s = (loc or "").strip().lower()
    s = s.replace("ms river", "msriver").replace("m.s. river", "msriver")
    s = s.replace("st.", "st").replace("st ", "st")
    s = s.replace("–", "-").replace("—", "-")
    # drop leading "River " artifact from parser
    if s.startswith("river "):
        s = s  # keep for alias table
    return alnum(s)


def loc_keys(location: str) -> list[str]:
    key = normalize_loc_key(location)
    keys = []
    # direct aliases
    for alias_key, vals in LOC_ALIASES.items():
        if key == alias_key or key.endswith(alias_key) or alias_key in key:
            keys.extend(vals)
    keys.append(key)
    # also stripped versions
    for drop in ("msrivers", "msriver", "river", "area"):
        if key.endswith(drop) and len(key) > len(drop) + 2:
            keys.append(key[: -len(drop)])
    # unique preserve order
    out = []
    for k in keys:
        if k not in out:
            out.append(k)
    return out


def channel_prefixes(channel: str) -> list[str]:
    c = (channel or "").lower()
    if "country" in c:
        return ["countryelevators", "country"]
    if "barge" in c:
        return ["bargeloadingelevators", "barge"]
    if "gulf" in c or "export" in c:
        return ["soybeans", "corn", "export", "exportelevators", "export elevators"]
    if "terminal elevator" in c or ( "terminal" in c and "mill" not in c and "processor" not in c):
        # Minneapolis Daily Grain Report (3046) Terminal Elevators
        return ["terminals", "terminalelevators", "millsproc", "millsandprocessors"]
    if "mill" in c or "terminal" in c or "processor" in c:
        return [
            "millsandprocessors",
            "millsproc",
            "terminalsmillsprocessors",
            "terminals",
            "terminalelevators",
            "millsandprocessors",
            "mills and processors",
        ]
    return []


def find_basis_dirs(root: Path) -> tuple[Path | None, Path | None]:
    soy = corn = None
    for p in root.rglob("*"):
        if not p.is_dir():
            continue
        name = p.name
        # match by presence of Illinois Basis.xlsx and parent folder hint
        if (p / "Illinois Basis.xlsx").exists():
            # distinguish soy vs corn by sibling folder names or file set
            files = [f.name for f in p.glob("*Basis.xlsx")]
            has_la_soy = (p / "Louisiana Basis.xlsx").exists()
            # corn folder has Texas/Wyoming/California; soy typically not
            has_tx = (p / "Texas Basis.xlsx").exists()
            has_wy = (p / "Wyoming Basis.xlsx").exists()
            has_ca = (p / "California Basis.xlsx").exists()
            if has_tx or has_wy or has_ca:
                corn = p
            else:
                soy = p
    return soy, corn


def sheet_parts(name: str) -> tuple[str, str]:
    """Return (prefix_alnum, suffix_alnum)."""
    if "_" in name:
        pre, suf = name.split("_", 1)
        return alnum(pre), alnum(suf)
    return alnum(name), ""


def last_date_in_sheet(ws) -> date | None:
    last = None
    for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
        v = row[0]
        if isinstance(v, datetime):
            last = v.date()
        elif isinstance(v, date):
            last = v
        elif isinstance(v, str) and len(v) >= 10 and v[4] == "-":
            try:
                last = datetime.strptime(v[:10], "%Y-%m-%d").date()
            except ValueError:
                pass
    return last


def date_already_present(ws, target: date) -> int | None:
    """Return 1-based row index if date exists, else None."""
    for i, row in enumerate(ws.iter_rows(min_row=2, max_col=1, values_only=True), start=2):
        v = row[0]
        d = None
        if isinstance(v, datetime):
            d = v.date()
        elif isinstance(v, date):
            d = v
        elif isinstance(v, str) and len(v) >= 10:
            try:
                d = datetime.strptime(v[:10], "%Y-%m-%d").date()
            except ValueError:
                d = None
        if d == target:
            return i
    return None


def score_sheet(sheet_name: str, location: str, channel: str, last: date | None, active_cutoff: date) -> float:
    pre, suf = sheet_parts(sheet_name)
    keys = loc_keys(location)
    prefixes = channel_prefixes(channel)
    score = -1.0

    # suffix match
    suf_hit = False
    for k in keys:
        if k == "":
            # statewide: prefer sheets with empty-ish suffix or bare MillsProc
            if suf in ("",) or sheet_name.lower() in ("millsproc", "mills and processors", "millsandprocessors_"):
                suf_hit = True
                score = max(score, 50)
            elif pre.startswith("mills") and suf == "":
                suf_hit = True
                score = max(score, 55)
            continue
        if suf == k:
            suf_hit = True
            score = max(score, 100)
        elif suf.startswith(k) or k.startswith(suf):
            # truncated sheet names like Mississip
            if min(len(suf), len(k)) >= 6:
                suf_hit = True
                score = max(score, 85)
        elif k in suf or suf in k:
            if min(len(suf), len(k)) >= 6:
                suf_hit = True
                score = max(score, 70)

    if not suf_hit and not (normalize_loc_key(location) in ("statewide", "") or location.strip() == "(statewide)"):
        return -1.0

    # channel prefix bonus
    pref_hit = False
    for p in prefixes:
        pa = alnum(p)
        if pre == pa or pre.startswith(pa) or pa.startswith(pre):
            pref_hit = True
            score += 20
            # prefer longer descriptive prefixes (CountryElevators > Country)
            score += min(len(pre), 30) * 0.15
            break
    if not pref_hit and prefixes:
        # allow Terminals for mills channel etc. already in prefixes
        score -= 5

    # prefer actively maintained sheets
    if last and last >= active_cutoff:
        score += 25
    elif last and last >= date(2026, 1, 1):
        score += 5
    else:
        score -= 30

    # slight preference for longer sheet names (canonical long form)
    score += min(len(sheet_name), 40) * 0.05

    # Minneapolis AMS 3046 / Terminal Elevators → prefer Terminals sheet over MillsProc
    loc_l = (location or "").strip().lower()
    ch_l = (channel or "").strip().lower()
    sn_l = sheet_name.lower()
    if "minneapolis" in loc_l or "terminal elevator" in ch_l:
        if sn_l in ("terminals", "terminalelevators") or pre in ("terminals", "terminalelevators"):
            score += 50
        elif "mills" in sn_l or pre.startswith("mills"):
            score -= 25
    return score


def load_csv(path: Path) -> list[dict]:
    rows = []
    with path.open(encoding="utf-8", newline="") as f:
        for r in csv.DictReader(f):
            if (r.get("delivery") or "Current").strip() != "Current":
                continue
            rows.append(r)
    return rows


def workbook_for(state_code: str, commodity: str, soy_dir: Path, corn_dir: Path) -> Path | None:
    # IA-MN barge lives in Minnesota workbooks
    code = state_code
    if code == "IA-MN":
        code = "MN"
    name = STATE_NAME.get(code)
    if not name:
        return None
    folder = soy_dir if commodity == "Soybeans" else corn_dir
    if folder is None:
        return None
    p = folder / f"{name} Basis.xlsx"
    return p if p.exists() else None


def ensure_change_column(ws) -> tuple[int, int, int]:
    """Ensure header has Change right after Current. Returns (date_col, current_col, change_col)."""
    headers = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    date_col = 1
    current_col = None
    change_col = None
    for i, h in enumerate(headers, start=1):
        hs = str(h or "").strip().lower()
        if hs in ("report_date", "reportdate", "date"):
            date_col = i
        if hs == "current":
            current_col = i
        if hs in ("change", "basis_change", "basischange", "变动", "基差变动"):
            change_col = i
    if current_col is None:
        current_col = 2
    if change_col is None:
        # Insert a new column immediately after Current
        insert_at = current_col + 1
        ws.insert_cols(insert_at)
        ws.cell(1, insert_at).value = "Change"
        change_col = insert_at
    return date_col, current_col, change_col


def _close(a, b) -> bool:
    try:
        if a is None and b is None:
            return True
        if a is None or b is None:
            return False
        return abs(float(a) - float(b)) < 1e-9
    except (TypeError, ValueError):
        return False


def append_or_update(ws, target: date, current_mid: float, change_mid=None) -> str:
    """Return 'append' or 'update' or 'skip'. Writes Current mid and optional AMS Basis Change mid."""
    date_col, current_col, change_col = ensure_change_column(ws)

    existing = date_already_present(ws, target)
    if existing:
        cur_cell = ws.cell(existing, current_col)
        ch_cell = ws.cell(existing, change_col)
        same_cur = _close(cur_cell.value, current_mid)
        same_ch = True
        if change_mid is not None and change_mid != "":
            same_ch = _close(ch_cell.value, change_mid)
        elif change_mid in (None, ""):
            same_ch = ch_cell.value in (None, "")
        if same_cur and same_ch:
            return "skip"
        cur_cell.value = float(current_mid)
        if change_mid is not None and change_mid != "":
            ch_cell.value = float(change_mid)
        return "update"

    # append after last non-empty date row
    last_r = 1
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, date_col).value is not None:
            last_r = r
    new_row = last_r + 1
    ws.cell(new_row, date_col).value = datetime(target.year, target.month, target.day)
    ws.cell(new_row, current_col).value = float(current_mid)
    if change_mid is not None and change_mid != "":
        ws.cell(new_row, change_col).value = float(change_mid)
    return "append"


def main() -> None:
    log = []
    def L(msg: str):
        print(msg)
        log.append(msg)

    soy_dir, corn_dir = find_basis_dirs(ROOT)
    L(f"soy_dir={soy_dir}")
    L(f"corn_dir={corn_dir}")
    if not soy_dir or not corn_dir:
        raise SystemExit("Could not locate soy/corn basis folders")

    rows = load_csv(CSV_PATH)
    L(f"csv_rows_current={len(rows)}")

    # group by workbook path for fewer open/close
    by_wb: dict[Path, list[dict]] = defaultdict(list)
    unmatched_state = []
    for r in rows:
        wb = workbook_for(r["state"], r["commodity"], soy_dir, corn_dir)
        if wb is None:
            unmatched_state.append(f"{r['state']}/{r['commodity']}/{r['location']}")
            continue
        by_wb[wb].append(r)

    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stats = defaultdict(int)
    unmatched_sheets = []
    active_cutoff = date(2026, 9, 1)

    for wb_path, items in sorted(by_wb.items(), key=lambda x: str(x[0])):
        # backup once
        bak = BACKUP_DIR / f"{wb_path.parent.name}__{wb_path.name}"
        if not bak.exists():
            shutil.copy2(wb_path, bak)

        wb = openpyxl.load_workbook(wb_path)
        sheet_names = [s for s in wb.sheetnames if s != "Index"]
        # precompute last dates
        lasts = {}
        for sn in sheet_names:
            lasts[sn] = last_date_in_sheet(wb[sn])

        # Map each CSV row to a sheet, then COMBINE duplicates that hit the same
        # sheet+date (e.g. IN soy Ohio River has two Current bids: DLVD-T and
        # DLVD-T/R). Last-write-wins previously caused fake day jumps like -14.5→-25.5.
        pending: dict[tuple[str, date], list] = defaultdict(list)
        # pending[(sheet, date)] -> list of (mid, change_mid, meta_str, score)
        for r in items:
            try:
                target = datetime.strptime(r["report_date"], "%Y-%m-%d").date()
            except ValueError:
                unmatched_sheets.append(("bad_date", r))
                continue
            mid = float(r["basis_mid"])
            ch_raw = r.get("basis_change_mid", "")
            try:
                change_mid = float(ch_raw) if ch_raw not in (None, "") else None
            except (TypeError, ValueError):
                change_mid = None
            location = r["location"]
            channel = r["channel"]

            best = None
            best_score = -1.0
            for sn in sheet_names:
                sc = score_sheet(sn, location, channel, lasts[sn], active_cutoff)
                if sc > best_score:
                    best_score = sc
                    best = sn

            if best is None or best_score < 40:
                unmatched_sheets.append(
                    (f"no_sheet score={best_score}", r["commodity"], r["state"], channel, location, r["report_date"])
                )
                stats["unmatched"] += 1
                continue

            meta = (
                f"{r['state']} {r['commodity']} {channel} {location} "
                f"mid={mid} change={change_mid} raw={r.get('basis_change_raw','')} "
                f"(score={best_score:.1f})"
            )
            pending[(best, target)].append((mid, change_mid, meta, best_score))

        for (best, target), vals in sorted(pending.items(), key=lambda x: (x[0][0], x[0][1])):
            mids = [v[0] for v in vals]
            chs = [v[1] for v in vals if v[1] is not None]
            mid = float(statistics.median(mids))
            change_mid = float(statistics.median(chs)) if chs else None
            if len(vals) > 1:
                stats["combined_dup_rows"] += len(vals)
                L(
                    f"combine x{len(vals)}: {wb_path.name} :: {best} {target} "
                    f"mids={mids} -> {mid}; changes={[v[1] for v in vals]} -> {change_mid}"
                )
                for v in vals:
                    L(f"    src: {v[2]}")
            action = append_or_update(wb[best], target, mid, change_mid)
            stats[action] += 1
            stats["sheets_touched"] += 1 if action != "skip" else 0
            if action != "skip":
                lasts[best] = target  # update cache
                L(
                    f"{action}: {wb_path.name} :: {best} <- {target} mid={mid} change={change_mid}"
                )

        wb.save(wb_path)
        wb.close()
        L(f"saved {wb_path}")

    L("--- summary ---")
    for k, v in sorted(stats.items()):
        L(f"  {k}: {v}")
    L(f"unmatched_sheet_count={len(unmatched_sheets)}")
    for u in unmatched_sheets[:40]:
        L(f"  unmatched: {u}")
    if unmatched_state:
        L(f"unmatched_state_count={len(unmatched_state)}")
        for u in unmatched_state[:20]:
            L(f"  no_wb: {u}")

    LOG_PATH.write_text("\n".join(log), encoding="utf-8")
    L(f"wrote log {LOG_PATH}")


if __name__ == "__main__":
    main()
