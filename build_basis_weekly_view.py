# -*- coding: utf-8 -*-
"""Weekly basis-change view: this Friday's Current minus last Friday's Current.

Usage: PYTHONPATH=/workspace python3 build_basis_weekly_view.py [YYYY-MM-DD 本周五]
Writes basis_daily/weekly.html and basis_daily/weekly_summary.json.
Per sheet: value on the Friday, or the latest earlier trading day with data
(noted as stale); a side more than 3 trading days stale drops the sheet.
State/layer change = median of per-sheet weekly changes.
"""
from __future__ import annotations

import json
import sys
from datetime import date, timedelta

import openpyxl

import build_basis_daily_view as B

MAX_STALE = 3


def trading_gap(d0: date, d1: date) -> int:
    """Weekdays strictly after d0 up to and including d1."""
    n, d = 0, d0
    while d < d1:
        d += timedelta(days=1)
        if d.weekday() < 5:
            n += 1
    return n


def value_at(series, target: date):
    pts = [(d, v) for d, v in series if d <= target]
    if not pts:
        return None
    d, v = pts[-1]
    if trading_gap(d, target) > MAX_STALE:
        return None
    return d, v


def load_weekly(folder: str, fri: date, prev: date) -> dict:
    layers = {"country": {}, "barge": {}, "export": {}, "mills": {}}
    for path in sorted((B.DATA / folder).glob("*.xlsx")):
        state = B.FILE_STATE.get(path.stem.replace(" Basis", "").replace(" Prices", ""))
        if state is None:
            continue
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        names = [s for s in wb.sheetnames if s != "Index"]
        wb.close()
        grouped = {}
        for sheet in names:
            if "white corn" in sheet.lower() or "whitecorn" in sheet.lower():
                continue
            kind = B.sheet_kind(sheet)
            if kind is None:
                continue
            series = B.read_sheet_current(path, sheet)
            if not series:
                continue
            key = (kind, B.location_key(sheet))
            if key not in grouped or len(series) > len(grouped[key][1]):
                grouped[key] = (sheet, series)
        per_kind = {}
        for (kind, _loc), (sheet, series) in grouped.items():
            a = value_at(series, fri)
            b = value_at(series, prev)
            if a is None or b is None:
                continue
            per_kind.setdefault(kind, []).append({
                "sheet": sheet, "cur": a[1], "cur_d": a[0].isoformat(),
                "old": b[1], "old_d": b[0].isoformat(), "chg": round(a[1] - b[1], 2),
            })
        for kind, rows in per_kind.items():
            # Require at least one sheet with a fresh quote this week.
            if not any(r["cur_d"] > prev.isoformat() for r in rows):
                continue
            layers[kind][state] = {
                "cur": B.median([r["cur"] for r in rows]),
                "chg": B.median([r["chg"] for r in rows]),
                "sheets": rows,
            }
    return layers


def build(fri: date) -> None:
    prev = fri - timedelta(days=7)
    soy = load_weekly("大豆基差", fri, prev)
    corn = load_weekly("玉米基差", fri, prev)
    asof = fri.isoformat()

    def pack(layers):
        out = {}
        for kind, states in layers.items():
            out[kind] = {}
            for code, s in states.items():
                if kind == "export" and code in B.EXPORT_GEO:
                    lat, lon, label = B.EXPORT_GEO[code]
                else:
                    lat, lon = B.STATE_GEO[code]
                    label = code
                x, y = B.xy(lon, lat)
                out[kind][code] = {"x": x, "y": y, "label": label, "dates": [asof],
                                   "values": [s["cur"]], "changes": [s["chg"]]}
        return out

    payload = {
        "asof": asof, "year": fri.year, "basemap": B.load_basemap(),
        "basemap_note": f"周变动 = {asof}（本周五）Current − {prev.isoformat()}（上周五）Current；某表当天无报价则取之前最近有数交易日，任一端缺失超过 3 个交易日的表不计",
        "states": [], "rivers": [{"name": n, "pts": [list(B.xy(lo, la)) for lo, la in p]} for n, p in B.RIVERS],
        "rails": [{"name": n, "x": B.xy(lo, la)[0], "y": B.xy(lo, la)[1]} for n, la, lo in B.RAILS],
        "ports": [{"code": c, "x": B.xy(lo, la)[0], "y": B.xy(lo, la)[1], "label": lb} for c, (la, lo, lb) in B.EXPORT_GEO.items()],
        "crops": {"soy": {"name": "大豆", "layers": pack(soy)}, "corn": {"name": "玉米", "layers": pack(corn)}},
        "season_states": {"soy": [], "corn": []},
        "season": {"soy": {}, "corn": {}},
    }
    html = B.render_html(payload)
    html = (html
            .replace("<title>美国大豆与玉米基差</title>", "<title>美国大豆与玉米基差 · 周度涨跌</title>")
            .replace("<h1>美国产地大豆与玉米基差</h1>", f"<h1>美国产地大豆与玉米基差 · 周度涨跌（{prev.isoformat()} → {asof}）</h1>")
            .replace("点的大小是相对上一个交易日的变动幅度", "点的大小是本周五相对上周五的周变动幅度")
            .replace("<th>日变动</th>", "<th>周变动</th>")
            .replace("图形越大，日变动越大", "图形越大，周变动越大")
            .replace('<button id="tab-soy" type="button">大豆季节性</button>', "")
            .replace('<button id="tab-corn" type="button">玉米季节性</button>', "")
            .replace('document.getElementById("tab-soy").onclick = () => show("soy");', "")
            .replace('document.getElementById("tab-corn").onclick = () => show("corn");', "")
            .replace('drawSeason("soy", document.getElementById("view-soy"));', "")
            .replace('drawSeason("corn", document.getElementById("view-corn"));', "")
            .replace("rows.slice(0, 18)", "rows.slice(0, 22)"))
    html = html.replace('["map","soy","corn"].forEach', '["map"].forEach')
    B.OUT.mkdir(parents=True, exist_ok=True)
    (B.OUT / "weekly.html").write_text(html, encoding="utf-8")
    summary = {"this_fri": asof, "last_fri": prev.isoformat(), "soy": soy, "corn": corn}
    (B.OUT / "weekly_summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=1), encoding="utf-8")
    print(B.OUT / "weekly.html")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        fri = date.fromisoformat(sys.argv[1])
    else:
        t = date.today()
        fri = t - timedelta(days=(t.weekday() - 4) % 7)
    build(fri)
