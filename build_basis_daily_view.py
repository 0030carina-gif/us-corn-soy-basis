# -*- coding: utf-8 -*-
"""Build a local HTML view of US soybean and corn basis.

Reads state workbooks under 玉米大豆历史数据/. Does not modify those files.
"""
from __future__ import annotations

import json
import statistics
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl

ROOT = Path(r"D:\grok知识库")
sys.path.insert(0, str(ROOT / "excel_s"))
from chart_watermark import (  # noqa: E402
    html_watermark_css,
    html_watermark_div,
    seasonal_year_styles,
)

DATA = Path(__file__).resolve().parent / "玉米大豆历史数据"
OUT = Path(__file__).resolve().parent / "basis_daily"
STATES_GEO = Path(__file__).resolve().parent / "us-states.json"

# Equirectangular window over the lower 48.
LON0, LON1 = -125.0, -66.5
LAT0, LAT1 = 24.2, 49.6
VW, VH = 1000.0, 620.0

# Representative states for seasonal panels.
SEASON_STATES = {
    "soy": ["IA", "IL", "MN", "IN"],
    "corn": ["IA", "IL", "NE", "IN"],
}

STATE_GEO = {
    "AL": (32.7, -86.7), "AR": (34.8, -92.2), "AZ": (34.3, -111.7),
    "CA": (37.2, -119.5), "CO": (39.0, -105.5), "CT": (41.6, -72.7),
    "DE": (39.0, -75.5), "FL": (28.6, -82.4), "GA": (32.7, -83.4),
    "IA": (42.0, -93.5), "ID": (44.4, -114.6), "IL": (40.1, -89.2),
    "IN": (39.9, -86.3), "KS": (38.5, -98.3), "KY": (37.5, -85.3),
    "LA": (31.0, -92.0), "MA": (42.3, -71.8), "MD": (39.1, -76.8),
    "ME": (45.3, -69.2), "MI": (43.7, -84.6), "MN": (46.0, -94.3),
    "MO": (38.4, -92.5), "MS": (32.7, -89.7), "MT": (47.0, -109.6),
    "NC": (35.6, -79.4), "ND": (47.4, -100.3), "NE": (41.5, -99.7),
    "NH": (43.7, -71.6), "NJ": (40.2, -74.7), "NM": (34.4, -106.1),
    "NV": (39.3, -116.6), "NY": (42.9, -75.5), "OH": (40.3, -82.8),
    "OK": (35.6, -97.4), "OR": (44.1, -120.5), "PA": (40.9, -77.6),
    "SC": (33.9, -80.9), "SD": (44.4, -100.2), "TN": (35.8, -86.3),
    "TX": (31.5, -99.3), "UT": (39.3, -111.7), "VA": (37.5, -78.6),
    "WA": (47.4, -120.5), "WI": (44.6, -89.7), "WV": (38.6, -80.6),
    "WY": (43.0, -107.6),
}

# Export quotes sit on the water, not on the quoting state's centroid.
EXPORT_GEO = {
    "LA": (29.95, -89.9, "新奥尔良·海湾"),
    "SC": (32.8, -79.9, "芒特普莱森特·东岸"),
    "OR": (45.6, -122.7, "波特兰·太平洋"),
    "OK": (29.3, -89.4, "海湾（俄克拉荷马报表）"),
}

RIVERS = [
    ("密西西比河", [(-95.0, 47.2), (-93.3, 44.9), (-91.2, 43.0), (-90.6, 41.5),
                   (-90.2, 38.6), (-90.1, 35.1), (-91.1, 30.5), (-89.9, 29.2)]),
    ("伊利诺伊河", [(-88.3, 41.5), (-89.5, 40.5), (-90.4, 38.9)]),
    ("俄亥俄河", [(-80.0, 40.4), (-84.5, 39.1), (-87.5, 37.9), (-89.2, 37.0)]),
    ("密苏里河", [(-100.4, 46.8), (-96.0, 41.3), (-94.6, 39.1), (-90.2, 38.6)]),
]

RAILS = [
    ("BNSF", 38.2, -101.2),
    ("UP", 41.2, -96.4),
    ("CN", 42.5, -87.8),
    ("NS / CSX", 39.4, -81.4),
]

FILE_STATE = {
    "Arkansas": "AR", "California": "CA", "Colorado": "CO", "Illinois": "IL",
    "Indiana": "IN", "Iowa": "IA", "Kansas": "KS", "Kentucky": "KY",
    "Louisiana": "LA", "Maryland": "MD", "Minnesota": "MN", "Mississippi": "MS",
    "Missouri": "MO", "Nebraska": "NE", "North Carolina": "NC",
    "North Dakota": "ND", "Ohio": "OH", "Oklahoma": "OK", "Oregon": "OR",
    "Pennsylvania": "PA", "South Carolina": "SC", "South Dakota": "SD",
    "Tennessee": "TN", "Texas": "TX", "Virginia": "VA", "Wyoming": "WY",
}


def xy(lon: float, lat: float) -> tuple[float, float]:
    x = (lon - LON0) / (LON1 - LON0) * VW
    y = (LAT1 - lat) / (LAT1 - LAT0) * VH
    return round(x, 1), round(y, 1)


def parse_date(value) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str) and len(value) >= 10 and value[4] == "-":
        try:
            return datetime.strptime(value[:10], "%Y-%m-%d").date()
        except ValueError:
            return None
    return None


def as_float(value) -> float | None:
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return None


def sheet_kind(name: str) -> str | None:
    key = name.lower().replace(" ", "")
    if key.startswith("barge") or "bargeloading" in key:
        return "barge"
    if (
        "gulf" in key
        or key.startswith("soybeans")
        or key.startswith("corn_")
        or "export" in key
        or "pacific" in key
        or "mountpleasant" in key
    ):
        return "export"
    if "mill" in key:
        return "mills"
    if key.startswith("country"):
        return "country"
    return None


def location_key(name: str) -> str:
    raw = name
    for prefix in (
        "BargeLoadingElevators_",
        "CountryElevators_",
        "Export Elevators_",
        "MillsandProcessors_",
        "TerminalsMillsProcessors_",
        "Mills and Processors_",
        "Mills and Processors",
        "TerminalElevators_",
        "Barge_",
        "Country_",
        "Export_",
        "MillsProc_",
        "Terminals_",
    ):
        if raw.startswith(prefix):
            raw = raw[len(prefix):]
            break
    return raw.lower().strip("_") or name.lower()


def read_sheet_current(path: Path, sheet: str) -> list[tuple[date, float]]:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet]
    out: list[tuple[date, float]] = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or not row:
            continue
        d = parse_date(row[0])
        v = as_float(row[1]) if len(row) > 1 else None
        if d is not None and v is not None:
            out.append((d, v))
    wb.close()
    out.sort()
    return out


def median(vals: list[float]) -> float | None:
    if not vals:
        return None
    return round(statistics.median(vals), 2)


def load_crop(folder: str, crop: str) -> dict:
    """Return layer -> state -> {dates, values, sheets, start, end}."""
    layers: dict[str, dict] = {"country": {}, "barge": {}, "export": {}, "mills": {}}
    coverage = []
    folder_path = DATA / folder
    for path in sorted(folder_path.glob("*.xlsx")):
        state = FILE_STATE.get(path.stem.replace(" Basis", "").replace(" Prices", ""))
        if state is None:
            continue
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        names = [s for s in wb.sheetnames if s != "Index"]
        wb.close()
        grouped: dict[tuple[str, str], tuple[str, list]] = {}
        for sheet in names:
            kind = sheet_kind(sheet)
            if kind is None:
                continue
            series = read_sheet_current(path, sheet)
            if not series:
                continue
            key = (kind, location_key(sheet))
            prev = grouped.get(key)
            if prev is None or len(series) > len(prev[1]):
                grouped[key] = (sheet, series)
        by_kind_dates: dict[str, dict[date, list[float]]] = {
            "country": {}, "barge": {}, "export": {}, "mills": {}
        }
        # Keep per-sheet maps so day-change = median of per-sheet deltas
        # (sheets present on BOTH today and previous agg date). Avoids fake
        # jumps when the set of sheets with a value differs by date.
        by_kind_sheet_maps: dict[str, list[dict[date, float]]] = {
            "country": [], "barge": [], "export": [], "mills": []
        }
        used = []
        for (kind, _loc), (sheet, series) in grouped.items():
            used.append(f"{kind}:{sheet}:{series[0][0].isoformat()}→{series[-1][0].isoformat()}:{len(series)}")
            bucket = by_kind_dates[kind]
            smap = {d: v for d, v in series}
            by_kind_sheet_maps[kind].append(smap)
            for d, v in series:
                bucket.setdefault(d, []).append(v)
        for kind, bucket in by_kind_dates.items():
            if not bucket:
                continue
            days = sorted(bucket)
            values = [median(bucket[d]) for d in days]
            sheet_maps = by_kind_sheet_maps[kind]
            changes: list[float | None] = []
            for i, d in enumerate(days):
                if i == 0:
                    changes.append(None)
                    continue
                prev = days[i - 1]
                deltas = []
                for sm in sheet_maps:
                    if d in sm and prev in sm:
                        deltas.append(sm[d] - sm[prev])
                if deltas:
                    changes.append(median(deltas))
                elif values[i] is not None and values[i - 1] is not None:
                    changes.append(round(values[i] - values[i - 1], 2))
                else:
                    changes.append(None)
            layers[kind].setdefault(state, {
                "dates": [d.isoformat() for d in days],
                "values": values,
                "changes": changes,
            })
        coverage.append({
            "file": path.name,
            "state": state,
            "sheets_kept": used,
        })
    return {"layers": layers, "files": coverage}


def seasonal_payload(series: dict | None, year: int) -> dict | None:
    """Each of the last five years, the current year, and their same-day mean."""
    if not series:
        return None
    by_year: dict[int, list[dict]] = {}
    hist: dict[str, list[float]] = {}
    latest = None
    for d_s, v in zip(series["dates"], series["values"]):
        y = int(d_s[:4])
        md = d_s[5:]
        if year - 5 <= y <= year:
            by_year.setdefault(y, []).append({"md": md, "v": v, "d": d_s})
            if y == year:
                latest = {"md": md, "v": v, "d": d_s}
            else:
                hist.setdefault(md, []).append(v)
    labels = [str(y) for y in range(year - 5, year + 1)]
    styles = seasonal_year_styles(labels)
    years = []
    for y in range(year - 5, year + 1):
        pts = by_year.get(y) or []
        if len(pts) < 2:
            continue
        st = styles[str(y)]
        years.append({
            "year": y,
            "color": st["color"],
            "width": st["width"],
            "pts": pts,
        })
    mean = []
    for md in sorted(hist):
        if len(hist[md]) >= 2:
            mean.append({
                "md": md,
                "v": round(sum(hist[md]) / len(hist[md]), 2),
                "n": len(hist[md]),
            })
    if not years and not mean:
        return None
    return {"years": years, "mean": mean, "latest": latest, "year": year}


def _ring_path(ring: list) -> str:
    parts = []
    for i, pair in enumerate(ring):
        lon, lat = pair[0], pair[1]
        x, y = xy(lon, lat)
        parts.append(("M" if i == 0 else "L") + f"{x:.1f} {y:.1f}")
    if parts:
        parts.append("Z")
    return "".join(parts)


def load_basemap() -> list[dict]:
    """Lower-48 outlines from the simplified Folium us-states GeoJSON."""
    fc = json.loads(STATES_GEO.read_text(encoding="utf-8"))
    shapes = []
    for feat in fc["features"]:
        code = feat.get("id") or ""
        if code in {"AK", "HI"}:
            continue
        geom = feat["geometry"]
        polys = geom["coordinates"]
        if geom["type"] == "Polygon":
            polys = [polys]
        chunks = []
        ring0 = None
        for poly in polys:
            if not poly:
                continue
            exterior = poly[0]
            inside = [
                p for p in exterior[::8]
                if LON0 <= p[0] <= LON1 and LAT0 <= p[1] <= LAT1
            ]
            if not inside:
                continue
            if ring0 is None:
                ring0 = exterior
            chunks.append(_ring_path(exterior))
        if not chunks or ring0 is None:
            continue
        xs, ys = [], []
        for lon, lat in ring0[::6]:
            x, y = xy(lon, lat)
            xs.append(x)
            ys.append(y)
        shapes.append({
            "code": code,
            "d": "".join(chunks),
            "x": round(sum(xs) / len(xs), 1),
            "y": round(sum(ys) / len(ys), 1),
        })
    return shapes


def build() -> None:
    soy = load_crop("大豆基差", "soy")
    corn = load_crop("玉米基差", "corn")
    all_dates = []
    for crop in (soy, corn):
        for layer in crop["layers"].values():
            for st in layer.values():
                if st["dates"]:
                    all_dates.append(st["dates"][-1])
    asof = max(all_dates) if all_dates else ""
    year = int(asof[:4]) if asof else date.today().year

    states_bg = []
    for code, (lat, lon) in STATE_GEO.items():
        x, y = xy(lon, lat)
        states_bg.append({"code": code, "x": x, "y": y})

    rivers = []
    for name, pts in RIVERS:
        rivers.append({"name": name, "pts": [list(xy(lon, lat)) for lon, lat in pts]})

    rails = []
    for name, lat, lon in RAILS:
        x, y = xy(lon, lat)
        rails.append({"name": name, "x": x, "y": y})

    ports = []
    for code, (lat, lon, label) in EXPORT_GEO.items():
        x, y = xy(lon, lat)
        ports.append({"code": code, "x": x, "y": y, "label": label})

    def pack_layers(crop_obj):
        out = {}
        for kind, states in crop_obj["layers"].items():
            out[kind] = {}
            for code, series in states.items():
                if kind == "export" and code in EXPORT_GEO:
                    lat, lon, label = EXPORT_GEO[code]
                else:
                    lat, lon = STATE_GEO[code]
                    label = code
                x, y = xy(lon, lat)
                entry = {
                    "x": x,
                    "y": y,
                    "label": label,
                    "dates": series["dates"],
                    "values": series["values"],
                }
                if series.get("changes") is not None:
                    entry["changes"] = series["changes"]
                out[kind][code] = entry
        return out

    payload = {
        "asof": asof,
        "year": year,
        "basemap": load_basemap(),
        "basemap_note": "州界来自简化的美国州界 GeoJSON（Folium us-states，仅本土 48 州）",
        "states": states_bg,
        "rivers": rivers,
        "rails": rails,
        "ports": ports,
        "crops": {
            "soy": {"name": "大豆", "layers": pack_layers(soy)},
            "corn": {"name": "玉米", "layers": pack_layers(corn)},
        },
        "season_states": SEASON_STATES,
    }
    season = {"soy": {}, "corn": {}}
    for key, obj in (("soy", soy), ("corn", corn)):
        for code in SEASON_STATES[key]:
            season[key][code] = {}
            for kind in ("country", "barge", "export", "mills"):
                season[key][code][kind] = seasonal_payload(obj["layers"][kind].get(code), year)
        for code in obj["layers"]["mills"]:
            season[key].setdefault(code, {})
            season[key][code]["mills"] = seasonal_payload(obj["layers"]["mills"][code], year)
        # Gulf lives on Louisiana, not on the interior states.
        season[key]["LA"] = {
            "export": seasonal_payload(obj["layers"]["export"].get("LA"), year)
        }
        if key == "soy":
            season[key]["SC"] = {
                "export": seasonal_payload(obj["layers"]["export"].get("SC"), year)
            }
    payload["season"] = season

    OUT.mkdir(parents=True, exist_ok=True)
    html = render_html(payload)
    (OUT / "index.html").write_text(html, encoding="utf-8")
    (OUT / "数据说明.md").write_text(render_md(soy, corn, asof), encoding="utf-8")
    print(OUT / "index.html")
    print("asof", asof)


def render_md(soy, corn, asof: str) -> str:
    lines = [
        "# 美国产地大豆与玉米基差 · 数据说明",
        "",
        f"页面只读本地工作簿，最新观测日 **{asof}**。基差取各表 `Current` 列，单位 ¢/bu。",
        "同一州、同一地点若有长短两套表（例如 `CountryElevators_*` 与 `Country_*`），保留观测更多的那一套，再按日取各分区中位数作为州序列。",
        "日变动优先用各分区表「当日−上一有数日」变动的中位数（避免某张表某日缺数时，州中位数对中位数出现假跳变）；若无法按表对齐则回退为州序列中位数之差。",
        "",
        "## 读过的文件",
        "",
        "目录：`excel_s/玉米大豆历史数据/玉米大豆历史数据/`",
        "",
        "价格工作簿在同目录的 `大豆价格/`、`玉米价格历史数据/`，本页没有使用。",
        "",
        "### 大豆基差",
        "",
    ]
    lines.extend(file_lines(soy))
    lines.extend(["", "### 玉米基差", ""])
    lines.extend(file_lines(corn))
    lines.extend([
        "",
        "## 三类基差覆盖到哪里",
        "",
        "- **Country Elevators（产地仓）**：多数州有分区序列，主产州（IA、IL、MN、IN、NE 等）收到 " + asof + "。马里兰等地更早停更，地图上若所选日期 10 日内没有点，不画气泡。",
        "- **Barge Loading Elevators（驳船装载仓）**：大豆有 IL、IN、KY、MN（密西西比河河段）、MO、OH、TN。玉米有 IL、IN、KY、MN、MO、OH、TN。爱荷华、内布拉斯加工作簿里没有驳船表。",
        "- **Gulf / Export（港口/出口）**：大豆、玉米的路易斯安那 `Gulf Coast Ports` 收到 " + asof + "。大豆南卡 `Export_EastCoastPorts` 同期有数。俄克拉荷马海湾出口只有 2020-08-03 一天；俄勒冈太平洋港口大豆只到 2020-11，2026 年 7 月另有两天。玉米没有南卡出口表，俄勒冈出口表在 2023-06 后 `Current` 为空。这些不进入“当前”地图。",
        "- **Mills & Processors（压榨厂）**：表名含 Mills / MillsandProcessors / MillsProc / Mills and Processors 的 Current 序列。州："
        + "大豆 " + ", ".join(sorted(soy["layers"]["mills"]))
        + "；玉米 " + ", ".join(sorted(corn["layers"]["mills"]))
        + "。没有这些表的州不画点。终点站和饲养场仍不在图里。",
        "- 总览底图是简化的美国州界（Folium 示例 us-states.json，本土 48 州），本地文件 us-states.json。",
        "- 季节性分年线覆盖最近六个日历年，另加除当年外近 5 年的同月日均值。",
        "",
        "## 页面里标成数据缺失的位置",
        "",
        "- 季节性里，IA、NE 没有驳船序列。",
        "- 季节性里，IA、IL、MN、IN、NE 没有本州港口序列。港口图单独用路易斯安那海湾；大豆另附南卡东岸。玉米东岸与太平洋岸标数据缺失。",
        "",
        "输出：`excel_s/玉米大豆历史数据/basis_daily/index.html`。只在本地，未推送。",
        "",
    ])
    return "\n".join(lines)


def file_lines(crop_obj) -> list[str]:
    rows = []
    for item in crop_obj["files"]:
        kinds = sorted({s.split(":")[0] for s in item["sheets_kept"]})
        ends = [s.split(":")[2] for s in item["sheets_kept"]]
        end = max(ends) if ends else "—"
        rows.append(
            f"- `{item['file']}`（{item['state']}）：层 {', '.join(kinds) or '无'}；保留序列最晚到 {end}。"
        )
    return rows


def render_html(payload: dict) -> str:
    data = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    css = html_watermark_css()
    wm = html_watermark_div()
    html = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>美国大豆与玉米基差</title>
<style>
__CSS__
:root {
  --bg:#f4f1ea; --ink:#1c2430; --muted:#5c6b7a; --card:#fffdf8;
  --line:#d5dbe3; --red:#E10600; --green:#1F8A4C; --flat:#8b939c;
}
* { box-sizing:border-box; }
body { margin:0; font-family:"Microsoft YaHei","PingFang SC",sans-serif; background:var(--bg); color:var(--ink); }
header { padding:22px 28px 8px; }
h1 { font-size:22px; margin:0 0 6px; font-weight:650; }
.sub { color:var(--muted); font-size:13px; line-height:1.55; max-width:980px; }
nav { display:flex; gap:8px; padding:8px 28px 0; }
nav button, .row button, .layers label {
  border:1px solid var(--line); background:#fff; border-radius:999px;
  padding:6px 12px; font-size:13px; cursor:pointer;
}
nav button.on, .row button.on { background:#1c2430; color:#fff; border-color:#1c2430; }
.panel { padding:12px 28px 40px; }
.card { background:var(--card); border:1px solid var(--line); border-radius:12px; padding:12px 14px 16px; margin-top:12px; }
.row { display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin-bottom:8px; }
.layers label { display:flex; gap:6px; align-items:center; }
.legend { display:flex; gap:14px; font-size:12px; color:var(--muted); }
.dot { width:10px; height:10px; border-radius:50%; display:inline-block; margin-right:4px; }
svg.map { width:100%; height:auto; background:#e7eef4; border-radius:10px; }
.side { display:grid; grid-template-columns: 1fr 280px; gap:12px; }
@media (max-width: 900px) { .side { grid-template-columns:1fr; } }
table { width:100%; border-collapse:collapse; font-size:12px; }
th, td { text-align:left; padding:4px 6px; border-bottom:1px solid var(--line); }
.grid { display:grid; grid-template-columns:1fr 1fr; gap:12px; }
@media (max-width: 900px) { .grid { grid-template-columns:1fr; } }
.miss { color:var(--muted); font-size:14px; padding:28px 8px; }
.chart-t { font-size:14px; margin:0 0 4px; }
.chart-s { font-size:12px; color:var(--muted); margin:0 0 6px; }
footer { color:var(--muted); font-size:12px; padding:0 4px; }
</style>
</head>
<body>
__WM__
<header>
  <h1>美国产地大豆与玉米基差</h1>
  <p class="sub">产地仓（圆）、驳船（方）、压榨厂 Mills &amp; Processors（菱形）、港口/出口（三角）。点的大小是相对上一个交易日的变动幅度（¢/bu），红色为基差走强，绿色为走弱，灰色为变动不足 0.5¢。州序列是各分区 Current 的中位数。没有序列的不画点。水印只作底纹。</p>
</header>
<nav>
  <button id="tab-map" class="on" type="button">总览</button>
  <button id="tab-soy" type="button">大豆季节性</button>
  <button id="tab-corn" type="button">玉米季节性</button>
</nav>
<section id="view-map" class="panel">
  <div class="card">
    <div class="row">
      <button id="crop-soy" class="on" type="button">大豆</button>
      <button id="crop-corn" type="button">玉米</button>
      <label>日期 <input id="asof" type="date"/></label>
      <span class="layers" id="layer-box"></span>
    </div>
    <div class="legend">
      <span><i class="dot" style="background:#E10600"></i>红：基差走强</span>
      <span><i class="dot" style="background:#1F8A4C"></i>绿：基差走弱</span>
      <span><i class="dot" style="background:#8b939c"></i>灰：持平（变动不足 0.5¢）</span>
      <span>圆=产地仓，方=驳船，菱形=压榨厂，三角=港口/出口</span>
      <span>图形越大，日变动越大（约 2¢ / 12¢ / 25¢）</span>
    </div>
    <div class="side">
      <svg class="map" id="map" viewBox="0 0 1000 620"></svg>
      <div>
        <table id="movers"><thead><tr><th>地点</th><th>层</th><th>基差</th><th>日变动</th></tr></thead><tbody></tbody></table>
        <p id="missing" class="chart-s"></p>
      </div>
    </div>
  </div>
</section>
<section id="view-soy" class="panel" hidden></section>
<section id="view-corn" class="panel" hidden></section>
<script>
const DATA = __DATA__;
const LAYER_NAME = {
  country: "产地仓 Country Elevators",
  barge: "驳船 Barge Loading",
  export: "港口/出口 Gulf · Export",
  mills: "压榨厂 Mills & Processors"
};
const LAYER_ORDER = ["country","barge","mills","export"];
let crop = "soy";
let layersOn = {country:true, barge:true, mills:true, export:true};

const asofInput = document.getElementById("asof");
asofInput.max = DATA.asof;
asofInput.value = DATA.asof;

function show(id) {
  ["map","soy","corn"].forEach(k => {
    document.getElementById("view-"+k).hidden = k !== id;
    document.getElementById("tab-"+k).classList.toggle("on", k === id);
  });
}
document.getElementById("tab-map").onclick = () => show("map");
document.getElementById("tab-soy").onclick = () => show("soy");
document.getElementById("tab-corn").onclick = () => show("corn");
document.getElementById("crop-soy").onclick = () => {{ crop="soy"; syncCrop(); drawMap(); }};
document.getElementById("crop-corn").onclick = () => {{ crop="corn"; syncCrop(); drawMap(); }};
asofInput.onchange = drawMap;

function syncCrop() {
  document.getElementById("crop-soy").classList.toggle("on", crop==="soy");
  document.getElementById("crop-corn").classList.toggle("on", crop==="corn");
}

const layerBox = document.getElementById("layer-box");
LAYER_ORDER.forEach(k => {
  const lab = document.createElement("label");
  const cb = document.createElement("input");
  cb.type = "checkbox"; cb.checked = true;
  cb.onchange = () => { layersOn[k] = cb.checked; drawMap(); };
  lab.appendChild(cb);
  lab.appendChild(document.createTextNode(LAYER_NAME[k]));
  layerBox.appendChild(lab);
});

function prevIndex(dates, iso) {
  let i = -1;
  for (let k=0;k<dates.length;k++) if (dates[k] <= iso) i = k;
  return i;
}

function pointAt(series, iso) {
  const i = prevIndex(series.dates, iso);
  if (i < 0) return null;
  const d = series.dates[i];
  const gap = (new Date(iso) - new Date(d)) / 86400000;
  if (gap > 10) return null;
  const v = series.values[i];
  let chg = null;
  if (series.changes && series.changes[i] != null && series.changes[i] !== undefined) {
    chg = Math.round(Number(series.changes[i]) * 100) / 100;
  } else {
    const prev = i > 0 ? series.values[i-1] : null;
    chg = prev == null ? null : Math.round((v - prev) * 100) / 100;
  }
  return {d, v, chg};
}

function drawMap() {
  const svg = document.getElementById("map");
  const ns = "http://www.w3.org/2000/svg";
  svg.replaceChildren();
  function el(name, attrs) {
    const n = document.createElementNS(ns, name);
    for (const [k,v] of Object.entries(attrs)) n.setAttribute(k, v);
    return n;
  }
  svg.appendChild(el("rect", {{x:0,y:0,width:1000,height:620,fill:"#d5e3ee"}}));
  (DATA.basemap || []).forEach(s => {{
    svg.appendChild(el("path", {{d:s.d, fill:"#f6f3ec", stroke:"#6d8498", "stroke-width":"1", "stroke-linejoin":"round"}}));
  }});
  (DATA.basemap || []).forEach(s => {{
    const t = el("text", {{x:s.x, y:s.y, "text-anchor":"middle", fill:"#9aa8b5", "font-size":"10"}});
    t.textContent = s.code;
    svg.appendChild(t);
  }});
  DATA.rivers.forEach(r => {{
    const d = r.pts.map((p,i) => (i?"L":"M")+p[0]+" "+p[1]).join(" ");
    svg.appendChild(el("path", {{d, fill:"none", stroke:"#7ea4c4", "stroke-width":"2.2"}}));
    const mid = r.pts[Math.floor(r.pts.length/2)];
    const t = el("text", {{x:mid[0]+6, y:mid[1]-6, fill:"#3d6f93", "font-size":"11"}});
    t.textContent = r.name;
    svg.appendChild(t);
  }});
  DATA.rails.forEach(r => {{
    const t = el("text", {{x:r.x, y:r.y, fill:"#8a6a3b", "font-size":"11", "font-style":"italic"}});
    t.textContent = r.name;
    svg.appendChild(t);
  }});
  DATA.ports.forEach(p => {{
    svg.appendChild(el("rect", {{x:p.x-3, y:p.y-3, width:6, height:6, fill:"#1c2430"}}));
    const t = el("text", {{x:p.x+8, y:p.y+4, fill:"#1c2430", "font-size":"11"}});
    t.textContent = p.label;
    svg.appendChild(t);
  }});

  const pack = DATA.crops[crop].layers;
  const rows = [];
  const missing = [];
  const offsets = {{country:-12, barge:12, mills:0, export:0}};
  const yoff = {{country:0, barge:0, mills:14, export:0}};
  LAYER_ORDER.forEach(kind => {{
    if (!layersOn[kind]) return;
    const states = pack[kind] || {{}};
    if (!Object.keys(states).length) {{
      missing.push(LAYER_NAME[kind] + "：数据缺失");
      return;
    }}
    let any = false;
    Object.entries(states).forEach(([code, series]) => {{
      const pt = pointAt(series, asofInput.value);
      if (!pt || pt.chg == null) return;
      any = true;
      const mag = Math.abs(pt.chg);
      const r = 6 + 16 * Math.min(mag / 25, 1);
      const fill = mag < 0.5 ? "#8b939c" : (pt.chg > 0 ? "#E10600" : "#1F8A4C");
      const cx = series.x + (offsets[kind] || 0);
      const cy = series.y + (yoff[kind] || 0);
      addMarker(svg, el, kind, cx, cy, r, fill);
      const label = kind === "export" ? series.label : code;
      const t = el("text", {{x:cx, y:cy - r - 3, "text-anchor":"middle", fill:"#1c2430", "font-size":"11", "font-weight":"650"}});
      t.textContent = label;
      svg.appendChild(t);
      rows.push({{label, kind: LAYER_NAME[kind], v: pt.v, chg: pt.chg, d: pt.d}});
    }});
    if (!any) missing.push(LAYER_NAME[kind] + "：所选日期附近数据缺失");
  }});
  rows.sort((a,b) => Math.abs(b.chg) - Math.abs(a.chg));
  const tb = document.querySelector("#movers tbody");
  tb.replaceChildren();
  rows.slice(0, 18).forEach(r => {{
    const tr = document.createElement("tr");
    const sign = r.chg > 0 ? "+" : "";
    tr.innerHTML = `<td>${{r.label}}</td><td>${{r.kind.split(" ")[0]}}</td><td>${{r.v}}</td><td style="color:${{r.chg > 0 ? '#d62728' : (r.chg < 0 ? '#2ca02c' : 'inherit')}};font-weight:600">${{sign}}${{r.chg}}</td>`;
    tb.appendChild(tr);
  }});
  document.getElementById("missing").textContent =
    (DATA.basemap_note || "") + (missing.length ? "。" + missing.join("；") : "");
}}

function addMarker(svg, el, kind, cx, cy, r, fill) {{
  const common = {{fill, "fill-opacity":"0.9", stroke:"#fff", "stroke-width":"1"}};
  if (kind === "mills") {{
    const d = `M ${{cx}} ${{cy-r}} L ${{cx+r}} ${{cy}} L ${{cx}} ${{cy+r}} L ${{cx-r}} ${{cy}} Z`;
    svg.appendChild(el("path", Object.assign({{d}}, common)));
  }} else if (kind === "barge") {{
    svg.appendChild(el("rect", Object.assign({{x:cx-r*0.8, y:cy-r*0.8, width:r*1.6, height:r*1.6}}, common)));
  }} else if (kind === "export") {{
    const d = `M ${{cx}} ${{cy-r}} L ${{cx+r*0.95}} ${{cy+r*0.75}} L ${{cx-r*0.95}} ${{cy+r*0.75}} Z`;
    svg.appendChild(el("path", Object.assign({{d}}, common)));
  }} else {{
    svg.appendChild(el("circle", Object.assign({{cx, cy, r}}, common)));
  }}
}}

function mdNum(md) {{
  const [m,d] = md.split("-").map(Number);
  return m * 100 + d;
}}

function drawSeason(cropKey, host) {{
  host.replaceChildren();
  const title = document.createElement("p");
  title.className = "sub";
  title.textContent = DATA.crops[cropKey].name + "：分年线为 " + (DATA.year-5) + "–" + DATA.year + "（最新一年为红色加粗，往前依次紫、黄、绿、蓝，更早为浅灰），灰色虚线是近 5 年同月日均值。最新点标数值。没有序列的格子写数据缺失。";
  host.appendChild(title);
  const states = DATA.season_states[cropKey];
  states.forEach(code => {{
    const h = document.createElement("h2");
    h.style.fontSize = "16px";
    h.textContent = code;
    host.appendChild(h);
    const grid = document.createElement("div");
    grid.className = "grid";
    ["country","barge","mills"].forEach(kind => {{
      grid.appendChild(chartCard(code + " · " + LAYER_NAME[kind], DATA.season[cropKey][code][kind]));
    }});
    host.appendChild(grid);
  }});
  const millStates = Object.keys((DATA.crops[cropKey].layers.mills) || {{}})
    .filter(code => !DATA.season_states[cropKey].includes(code))
    .sort();
  if (millStates.length) {{
    const hm = document.createElement("h2");
    hm.style.fontSize = "16px";
    hm.textContent = "其他有压榨厂序列的州";
    host.appendChild(hm);
    const mg = document.createElement("div");
    mg.className = "grid";
    millStates.forEach(code => {{
      const block = DATA.season[cropKey][code] && DATA.season[cropKey][code].mills;
      mg.appendChild(chartCard(code + " · " + LAYER_NAME.mills, block));
    }});
    host.appendChild(mg);
  }}
  const h = document.createElement("h2");
  h.style.fontSize = "16px";
  h.textContent = "港口 / 出口";
  host.appendChild(h);
  const grid = document.createElement("div");
  grid.className = "grid";
  const la = DATA.season[cropKey].LA && DATA.season[cropKey].LA.export;
  grid.appendChild(chartCard("LA 海湾 Gulf Coast Ports", la));
  if (cropKey === "soy") {{
    grid.appendChild(chartCard("SC 东岸 Export", DATA.season.soy.SC && DATA.season.soy.SC.export));
  }} else {{
    grid.appendChild(chartCard("东岸 / 太平洋岸", null));
  }}
  host.appendChild(grid);
  const note = document.createElement("p");
  note.className = "chart-s";
  note.textContent = "主产州 IA、IL、MN、IN、NE 的工作簿没有本州港口表，港口季节性只用路易斯安那海湾"
    + (cropKey==="soy" ? "和南卡东岸。" : "。玉米东岸与太平洋岸：数据缺失。");
  host.appendChild(note);
}}

function chartCard(title, series) {{
  const card = document.createElement("div");
  card.className = "card";
  const h = document.createElement("p");
  h.className = "chart-t";
  h.textContent = title;
  card.appendChild(h);
  const yearPts = (series && series.years) ? series.years : [];
  if (!series || (!yearPts.length && (!series.mean || !series.mean.length))) {{
    const p = document.createElement("p");
    p.className = "miss";
    p.textContent = "数据缺失";
    card.appendChild(p);
    return card;
  }}
  const cap = document.createElement("p");
  cap.className = "chart-s";
  const last = series.latest;
  cap.textContent = last ? ("最新 " + last.d + "  " + last.v + " ¢/bu") : "当年尚无观测";
  card.appendChild(cap);
  const legend = document.createElement("div");
  legend.className = "legend";
  yearPts.forEach(y => {{
    const span = document.createElement("span");
    span.innerHTML = `<i class="dot" style="background:${{y.color}};border-radius:2px;width:14px;height:3px"></i>${{y.year}}`;
    legend.appendChild(span);
  }});
  if (series.mean && series.mean.length) {{
    const span = document.createElement("span");
    span.innerHTML = `<i class="dot" style="background:#5c6b7a;border-radius:2px;width:14px;height:3px"></i>5年同日均值`;
    legend.appendChild(span);
  }}
  card.appendChild(legend);
  card.appendChild(seasonSvg(series));
  return card;
}}

function seasonSvg(series) {{
  const ns = "http://www.w3.org/2000/svg";
  const W = 520, H = 280, L = 46, R = 12, T = 16, B = 28;
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", `0 0 ${{W}} ${{H}}`);
  svg.style.width = "100%";
  function el(name, attrs) {{
    const n = document.createElementNS(ns, name);
    for (const [k,v] of Object.entries(attrs)) n.setAttribute(k, String(v));
    return n;
  }}
  const pts = []
    .concat((series.years||[]).flatMap(y => y.pts.map(p => p.v)))
    .concat((series.mean||[]).map(p => p.v));
  let ymin = Math.min(...pts), ymax = Math.max(...pts);
  if (ymin === ymax) {{ ymin -= 5; ymax += 5; }}
  const pad = (ymax - ymin) * 0.12;
  ymin -= pad; ymax += pad;
  const xOf = md => {{
    const [m,d] = md.split("-").map(Number);
    const doy = (Date.UTC(2024, m-1, d) - Date.UTC(2024,0,1)) / 86400000;
    return L + doy / 365 * (W - L - R);
  }};
  const yOf = v => T + (ymax - v) / (ymax - ymin) * (H - T - B);
  svg.appendChild(el("rect", {{x:0,y:0,width:W,height:H,fill:"#fffdf8"}}));
  for (let m=1;m<=12;m++) {{
    const x = xOf(String(m).padStart(2,"0")+"-01");
    svg.appendChild(el("line", {{x1:x,y1:T,x2:x,y2:H-B,stroke:"#e6ebf0"}}));
    const t = el("text", {{x, y:H-10, fill:"#5c6b7a", "font-size":"10"}});
    t.textContent = m+"月";
    svg.appendChild(t);
  }}
  [ymin, (ymin+ymax)/2, ymax].forEach(v => {{
    const y = yOf(v);
    svg.appendChild(el("line", {{x1:L,y1:y,x2:W-R,y2:y,stroke:"#e6ebf0"}}));
    const t = el("text", {{x:4, y:y+3, fill:"#5c6b7a", "font-size":"10"}});
    t.textContent = Math.round(v);
    svg.appendChild(t);
  }});
  function path(arr, stroke, width, dash) {{
    const sorted = arr.slice().sort((a,b) => mdNum(a.md) - mdNum(b.md));
    if (sorted.length < 2) return;
    const d = sorted.map((p,i) => (i?"L":"M")+xOf(p.md).toFixed(1)+" "+yOf(p.v).toFixed(1)).join(" ");
    const attrs = {{d, fill:"none", stroke, "stroke-width":width, "stroke-linejoin":"round"}};
    if (dash) attrs["stroke-dasharray"] = dash;
    svg.appendChild(el("path", attrs));
  }}
  path(series.mean||[], "#5c6b7a", 1.5, "4 3");
  (series.years || []).forEach(y => path(y.pts, y.color, y.width, null));
  if (series.latest) {{
    const x = xOf(series.latest.md), y = yOf(series.latest.v);
    svg.appendChild(el("circle", {{cx:x, cy:y, r:3.5, fill:"#E10600"}}));
    const t = el("text", {{x:Math.min(x+6, W-48), y:y-6, fill:"#E10600", "font-size":"11", "font-weight":"700"}});
    t.textContent = series.latest.v;
    svg.appendChild(t);
  }}
  return svg;
}}

drawMap();
drawSeason("soy", document.getElementById("view-soy"));
drawSeason("corn", document.getElementById("view-corn"));
</script>
</body>
</html>
"""
    html = html.replace("{{", "{").replace("}}", "}")
    return (
        html.replace("__CSS__", css)
        .replace("__WM__", wm)
        .replace("__DATA__", data)
    )


if __name__ == "__main__":
    build()
