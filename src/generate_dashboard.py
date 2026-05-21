"""
Palissy Multi-Dataset Dashboard Generator

Reads multiple dataset tabs from INPUT/gas_model_input.xlsx and produces a
single self-contained HTML page with a Gas/Power toggle.

Datasets are declared in DATASETS below. Each dataset is one config block —
adding a new page (e.g. LNG, Storage) is a config addition, not a code change.

Gas page preserves all prior behavior; Power page is new.
"""

import openpyxl
import json
import os
import base64
from datetime import datetime
from calendar import monthrange  # noqa: F401  (kept for parity)

# ============================================================
# ADMIN CONFIGURATION
# ============================================================
DISPLAY_START_YEAR = 2020
DISPLAY_END_YEAR = 2030
# ============================================================

# Paths
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
INPUT_FILE = os.path.join(PROJECT_DIR, "INPUT", "gas_model_input.xlsx")
OUTPUT_DIR = os.path.join(PROJECT_DIR, "output")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "index.html")
LOGO_FILE = os.path.join(PROJECT_DIR, "Context", "Palissy Logo.png")
FONT_FILE = os.path.join(PROJECT_DIR, "Context", "Gotham-Book.otf")

# Palissy brand colors
COLORS = {
    "dark_blue": "#272962",
    "light_green": "#539648",
    "dark_green": "#0C5B19",
    "red": "#C00000",
    "grey": "#9395A2",
    "light_blue": "#0B5AAB",
    "blue": "#258EEB",
    "card_bg": "#f8f8fb",
    "border": "rgba(39, 41, 98, 0.15)",
    "grid": "#E0E0E8",
}

# ============================================================
# DATASETS
# Each entry is a self-contained spec for one page.
# ============================================================
GAS_UNIT_CONFIG = {
    "bcf":    {"volLabel": "bcf",  "rateLabel": "bcf/d",  "volFactor": 1.0,         "rateFactor": 1.0,           "isRate": False},
    "bcf/d":  {"volLabel": "bcf",  "rateLabel": "bcf/d",  "volFactor": 1.0,         "rateFactor": 1.0,           "isRate": True},
    "bcm":    {"volLabel": "bcm",  "rateLabel": "mmcm/d", "volFactor": 1.0/35.3,    "rateFactor": 1000.0/35.3,   "isRate": False},
    "mmcm/d": {"volLabel": "bcm",  "rateLabel": "mmcm/d", "volFactor": 1.0/35.3,    "rateFactor": 1000.0/35.3,   "isRate": True},
    "TWh":    {"volLabel": "TWh",  "rateLabel": "GWh/d",  "volFactor": 1.0/3.41,    "rateFactor": 1000.0/3.41,   "isRate": False},
    "GWh/d":  {"volLabel": "TWh",  "rateLabel": "GWh/d",  "volFactor": 1.0/3.41,    "rateFactor": 1000.0/3.41,   "isRate": True},
    "mmt":    {"volLabel": "mmt",  "rateLabel": "kt/d",   "volFactor": 1.0/48.0,    "rateFactor": 1000.0/48.0,   "isRate": False},
    "kt/d":   {"volLabel": "mmt",  "rateLabel": "kt/d",   "volFactor": 1.0/48.0,    "rateFactor": 1000.0/48.0,   "isRate": True},
}

# Gas-to-power efficiency assumption for BCFE conversions.
# BCFE = (TWh) * 3.41 / POWER_EFFICIENCY
# Industry-standard "thermal efficiency" placeholder; change here if the
# assumption updates.
POWER_EFFICIENCY = 0.39
BCFE_FACTOR = 3.41 / POWER_EFFICIENCY  # ~= 8.7436

POWER_UNIT_CONFIG = {
    "TWh":    {"volLabel": "TWh",  "rateLabel": "TWh",    "volFactor": 1.0,         "rateFactor": 1.0,           "isRate": False},
    "BCFE":   {"volLabel": "BCFE", "rateLabel": "BCFE/d", "volFactor": BCFE_FACTOR, "rateFactor": BCFE_FACTOR,   "isRate": False},
    "BCFE/d": {"volLabel": "BCFE", "rateLabel": "BCFE/d", "volFactor": BCFE_FACTOR, "rateFactor": BCFE_FACTOR,   "isRate": True},
}

# Palissy-aligned source palette for power generation.
# Tuned for visual differentiation between adjacent stack layers and
# industry-standard fuel-source associations (gas=red, coal=charcoal,
# nuclear=gold, solar=orange, wind=sky, hydro=deep blue, etc.).
# Tweakable - user said "play around with the colors".
POWER_SOURCE_COLORS = {
    # Bottom five sources match the seasonality chart's line colors so that
    # users can pattern-match between the two charts. Top four picked for
    # contrast and Palissy-palette consistency.
    # Stack order bottom -> top: Nuclear / Wind / Gas / Solar / Hydro / Coal
    # / Bioenergy / Other Fossil / Other Renewables.
    "Nuclear":                              "#272962",  # Palissy dark blue (= seasonality avg)
    "Wind":                                 "#539648",  # Palissy light green (= seasonality forecast)
    "Gas":                                  "#C00000",  # Palissy red (= seasonality current)
    "Solar":                                "#E5B83A",  # warm gold-yellow
    "Hydro":                                "#0C5B19",  # Palissy dark green (= seasonality prev)
    "Coal":                                 "#2C2C2E",  # charcoal
    "Bioenergy":                            "#708B5A",  # muted olive
    "Other Fossil":                         "#92591C",  # earth brown
    "Other Renewables":                     "#B8D67A",  # pale mint
    "Total Europe Electricity Generation":  "#272962",  # Palissy dark blue
}

DATASETS = [
    {
        "key": "gas",
        "tab_label": "Gas",
        "title": "European Gas Balance",
        "sheet": "Monthly Data",
        "date_row": 1,
        "days_row": 2,
        "data_start_row": 5,
        "base_unit": "bcf",
        "units": ["bcf", "bcf/d", "bcm", "mmcm/d", "TWh", "GWh/d", "mmt", "kt/d"],
        "default_unit": "bcf",
        "unit_config": GAS_UNIT_CONFIG,
        "stock_rows": ["Opening Storage", "Closing Storage", "Storage percentage"],
        "pct_rows": ["Storage percentage"],
        "use_hierarchy": True,
        "skip_label_rows": [],
        "charts_enabled": False,
        "total_row": None,
        "source_colors": {},
    },
    {
        "key": "power",
        "tab_label": "Power",
        "title": "European Power Generation",
        "sheet": "Power Data",
        "date_row": 6,
        "days_row": 4,
        "data_start_row": 8,
        "base_unit": "TWh",
        "units": ["TWh", "BCFE", "BCFE/d"],
        "default_unit": "TWh",
        "unit_config": POWER_UNIT_CONFIG,
        "stock_rows": [],
        "pct_rows": [],
        "use_hierarchy": False,
        # Skip the "Generation by Source - EU + UK" section header which has no numeric data.
        "skip_label_rows": ["Generation by Source - EU + UK"],
        "charts_enabled": True,
        # Row label that represents the total - used by charts to handle Total
        # as a special selection (mutex with individual sources).
        "total_row": "Total Europe Electricity Generation",
        "source_colors": POWER_SOURCE_COLORS,
        # Sort non-total rows descending by total generation in this calendar year.
        # Used to order the table rows and the buildup-chart stack so the largest
        # contributor sits at the top of the table / bottom of the stack.
        "sort_by_year": 2025,
    },
]


# ============================================================
# DATA LOADING
# ============================================================
def read_dataset(wb, config):
    """Read one dataset tab according to its config. Returns (dates, days, rows)."""
    sheet_name = config["sheet"]
    print(f"\n[{config['key']}] Reading sheet '{sheet_name}'...")
    ws = wb[sheet_name]

    date_row = config["date_row"]
    days_row = config["days_row"]
    data_start_row = config["data_start_row"]

    # Detect last column with a date in the date_row
    last_col = 1
    for c in range(2, ws.max_column + 1):
        if ws.cell(row=date_row, column=c).value is not None:
            last_col = c

    # Dates
    dates = []
    for c in range(2, last_col + 1):
        v = ws.cell(row=date_row, column=c).value
        if isinstance(v, datetime):
            dates.append(v)
        elif isinstance(v, str):
            dates.append(datetime.strptime(v, "%Y-%m-%d"))
        else:
            dates.append(None)

    # Days
    days_per_month = []
    for c in range(2, last_col + 1):
        v = ws.cell(row=days_row, column=c).value
        days_per_month.append(int(v) if v is not None else 30)

    # Data rows
    skip = set(config.get("skip_label_rows", []))
    rows = []
    r = data_start_row
    blank_streak = 0
    while r <= ws.max_row and blank_streak < 3:
        label = ws.cell(row=r, column=1).value
        if label is None:
            blank_streak += 1
            r += 1
            continue
        blank_streak = 0
        label_str = str(label).strip()
        if label_str in skip:
            r += 1
            continue
        values = []
        for c in range(2, last_col + 1):
            v = ws.cell(row=r, column=c).value
            values.append(float(v) if v is not None else 0.0)
        # If no numeric values at all (e.g. orphan label), skip
        if all(v == 0.0 for v in values) and label_str not in config.get("stock_rows", []):
            # Could be a real all-zero forecast row; only skip if also has no
            # children-style behavior. For safety, keep all-zero rows so we
            # don't accidentally drop a real (empty) source.
            pass
        rows.append({"label": label_str, "values": values})
        r += 1

    print(f"  Loaded {len(rows)} rows x {len(dates)} months ({dates[0].strftime('%b %Y')} - {dates[-1].strftime('%b %Y')})")
    return dates, days_per_month, rows


# ============================================================
# HIERARCHY DETECTION
# ============================================================
def detect_hierarchy(rows, stock_rows, use_hierarchy):
    """Detect parent/child structure.

    With use_hierarchy=True (gas pattern):
      Rows prefixed + or - = parent totals
      Un-prefixed rows between parents = children of the NEXT parent
      stock_rows = always standalone (no children)

    With use_hierarchy=False (flat pattern, e.g. power):
      All rows are standalone, in source order.
    """
    if not use_hierarchy:
        return [
            {
                "label": row["label"],
                "row_index": i,
                "children": [],
                "type": "standalone",
            }
            for i, row in enumerate(rows)
        ]

    standalone = set(stock_rows)
    classified = []
    for i, row in enumerate(rows):
        label = row["label"]
        classified.append({
            "index": i,
            "label": label,
            "is_parent": label.startswith("+") or label.startswith("-"),
            "is_standalone": label in standalone,
        })

    hierarchy = []
    pending = []
    for item in classified:
        if item["is_standalone"]:
            for child in pending:
                hierarchy.append({"label": child["label"], "row_index": child["index"], "children": [], "type": "standalone"})
            pending = []
            hierarchy.append({"label": item["label"], "row_index": item["index"], "children": [], "type": "standalone"})
        elif item["is_parent"]:
            if pending:
                children = [{"label": c["label"], "row_index": c["index"]} for c in pending]
                hierarchy.append({"label": item["label"], "row_index": item["index"], "children": children, "type": "parent"})
                pending = []
            else:
                hierarchy.append({"label": item["label"], "row_index": item["index"], "children": [], "type": "standalone"})
        else:
            pending.append(item)

    for child in pending:
        hierarchy.append({"label": child["label"], "row_index": child["index"], "children": [], "type": "standalone"})
    return hierarchy


# ============================================================
# AGGREGATION (period-level)
# ============================================================
def aggregate_monthly_to_periods(dates, days_per_month):
    """Group monthly indices into period columns. Generic across datasets."""
    results = {}

    # Monthly
    monthly_cols = [{
        "label": d.strftime("%b %Y"),
        "short": d.strftime("%b-%y"),
        "year": d.year,
        "month": d.month,
        "gas_year": d.year if d.month >= 10 else d.year - 1,
        "indices": [i],
        "days": days_per_month[i],
    } for i, d in enumerate(dates)]
    results["Monthly"] = monthly_cols

    # Quarterly
    quarters = {}
    for i, d in enumerate(dates):
        q = (d.month - 1) // 3 + 1
        key = (d.year, q)
        quarters.setdefault(key, {"indices": [], "days": 0})
        quarters[key]["indices"].append(i)
        quarters[key]["days"] += days_per_month[i]
    results["Quarterly"] = [{
        "label": f"Q{q} {y}", "short": f"Q{q}-{str(y)[2:]}",
        "year": y, "quarter": q,
        "gas_year": y if q == 4 else y - 1,
        "indices": info["indices"], "days": info["days"],
    } for (y, q), info in sorted(quarters.items()) if len(info["indices"]) == 3]

    # Annual Calendar Year
    years = {}
    for i, d in enumerate(dates):
        years.setdefault(d.year, {"indices": [], "days": 0})
        years[d.year]["indices"].append(i)
        years[d.year]["days"] += days_per_month[i]
    results["Annual CY"] = [{
        "label": str(y), "short": str(y), "year": y,
        "indices": info["indices"], "days": info["days"],
    } for y, info in sorted(years.items()) if len(info["indices"]) == 12]

    # Annual Gas Year
    gas_years = {}
    for i, d in enumerate(dates):
        gy = d.year if d.month >= 10 else d.year - 1
        gas_years.setdefault(gy, {"indices": [], "days": 0})
        gas_years[gy]["indices"].append(i)
        gas_years[gy]["days"] += days_per_month[i]
    results["Gas Year"] = [{
        "label": f"GY {str(gy)[2:]}/{str(gy+1)[2:]}",
        "short": f"{str(gy)[2:]}/{str(gy+1)[2:]}",
        "year": gy, "indices": info["indices"], "days": info["days"],
    } for gy, info in sorted(gas_years.items()) if len(info["indices"]) == 12]

    # Winters (Oct-Mar)
    winters = {}
    for i, d in enumerate(dates):
        if d.month >= 10:
            gy = d.year
        elif d.month <= 3:
            gy = d.year - 1
        else:
            continue
        winters.setdefault(gy, {"indices": [], "days": 0})
        winters[gy]["indices"].append(i)
        winters[gy]["days"] += days_per_month[i]
    results["Winter"] = [{
        "label": f"Win {str(gy)[2:]}/{str(gy+1)[2:]}",
        "short": f"Win {str(gy)[2:]}/{str(gy+1)[2:]}",
        "year": gy, "indices": info["indices"], "days": info["days"],
    } for gy, info in sorted(winters.items()) if len(info["indices"]) == 6]

    # Summers (Apr-Sep)
    summers = {}
    for i, d in enumerate(dates):
        if 4 <= d.month <= 9:
            summers.setdefault(d.year, {"indices": [], "days": 0})
            summers[d.year]["indices"].append(i)
            summers[d.year]["days"] += days_per_month[i]
    results["Summer"] = [{
        "label": f"Sum {y}", "short": f"Sum {y}",
        "year": y, "indices": info["indices"], "days": info["days"],
    } for y, info in sorted(summers.items()) if len(info["indices"]) == 6]

    return results


def compute_period_values(rows, period_cols, stock_rows, pct_rows):
    """For each period column, compute aggregated value per row.
    Stock rows: opening = first index value; closing = last index value;
    Storage percentage row: last index value.
    All other rows: sum of the period.
    """
    out = []
    for row in rows:
        label = row["label"]
        vals = row["values"]
        period_values = []
        for col in period_cols:
            idx = col["indices"]
            if not idx:
                period_values.append(0)
                continue
            if label == "Opening Storage":
                period_values.append(vals[idx[0]])
            elif label == "Closing Storage":
                period_values.append(vals[idx[-1]])
            elif label in pct_rows:
                period_values.append(vals[idx[-1]])
            elif label in stock_rows:
                # Generic stock fallback (last value)
                period_values.append(vals[idx[-1]])
            else:
                period_values.append(sum(vals[i] for i in idx))
        out.append({"label": label, "base_values": period_values})
    return out


# ============================================================
# DATASET BUILDER
# ============================================================
def build_dataset_blob(config):
    """Load + aggregate one dataset and produce its JSON-ready blob."""
    wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)
    try:
        dates, days, rows = read_dataset(wb, config)
    finally:
        wb.close()

    # Optional: sort non-total rows descending by total in a target calendar year.
    # Keeps the largest contributors at the top of the table and at the bottom
    # of the stacked buildup chart.
    sort_year = config.get("sort_by_year")
    if sort_year:
        total_label = config.get("total_row")
        target_indices = [i for i, d in enumerate(dates) if d.year == sort_year]
        def row_year_sum(r):
            return sum(r["values"][i] for i in target_indices) if target_indices else 0
        sortable = [r for r in rows if r["label"] != total_label]
        total_rows = [r for r in rows if r["label"] == total_label]
        sortable.sort(key=row_year_sum, reverse=True)
        rows = sortable + total_rows
        print(f"  Sorted rows by CY {sort_year} generation (descending):")
        for r in sortable:
            print(f"    {r['label']}: {row_year_sum(r):,.1f}")

    hierarchy = detect_hierarchy(rows, config["stock_rows"], config["use_hierarchy"])
    print(f"  Hierarchy: {len(hierarchy)} items, {sum(1 for h in hierarchy if h['children']) } with children")

    period_results = aggregate_monthly_to_periods(dates, days)
    print(f"  Periods: " + ", ".join(f"{k}={len(v)}" for k, v in period_results.items()))

    views = {}
    for view_name, period_cols in period_results.items():
        aggregated = compute_period_values(rows, period_cols, config["stock_rows"], config["pct_rows"])
        days_array = [c["days"] for c in period_cols]
        col_meta = []
        for c in period_cols:
            meta = {"label": c["label"], "short": c["short"], "year": c.get("year", 0), "days": c["days"]}
            if "month" in c: meta["month"] = c["month"]
            if "quarter" in c: meta["quarter"] = c["quarter"]
            if "gas_year" in c: meta["gas_year"] = c["gas_year"]
            col_meta.append(meta)
        views[view_name] = {
            "columns": [c["label"] for c in period_cols],
            "short_columns": [c["short"] for c in period_cols],
            "col_meta": col_meta,
            "days": days_array,
            "rows": [{"label": a["label"], "base": a["base_values"]} for a in aggregated],
        }

    ui_hierarchy = []
    for item in hierarchy:
        entry = {
            "label": item["label"],
            "index": item["row_index"],
            "type": item["type"],
            "is_stock": item["label"] in config["stock_rows"],
            "is_pct": item["label"] in config["pct_rows"],
            "children": [],
        }
        for child in item["children"]:
            entry["children"].append({
                "label": child["label"],
                "index": child["row_index"],
                "is_stock": False,
                "is_pct": False,
            })
        ui_hierarchy.append(entry)

    return {
        "key": config["key"],
        "tab_label": config["tab_label"],
        "title": config["title"],
        "base_unit": config["base_unit"],
        "units": config["units"],
        "default_unit": config["default_unit"],
        "unit_config": config["unit_config"],
        "use_hierarchy": config["use_hierarchy"],
        "charts_enabled": config.get("charts_enabled", False),
        "total_row": config.get("total_row"),
        "source_colors": config.get("source_colors", {}),
        "views": views,
        "hierarchy": ui_hierarchy,
        "selectable_start": DISPLAY_START_YEAR,
        "selectable_end": DISPLAY_END_YEAR,
    }


# ============================================================
# ASSET LOADING
# ============================================================
def load_assets():
    assets = {}
    if os.path.exists(LOGO_FILE):
        with open(LOGO_FILE, "rb") as f:
            assets["logo_b64"] = base64.b64encode(f.read()).decode("utf-8")
    else:
        assets["logo_b64"] = ""
    if os.path.exists(FONT_FILE):
        with open(FONT_FILE, "rb") as f:
            assets["font_b64"] = base64.b64encode(f.read()).decode("utf-8")
    else:
        assets["font_b64"] = ""
    return assets


# ============================================================
# HTML GENERATION
# ============================================================
def generate_html(datasets_by_key, ordered_keys, assets):
    master = {
        "datasets": datasets_by_key,
        "order": ordered_keys,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    data_json = json.dumps(master, separators=(",", ":"))
    logo_b64 = assets.get("logo_b64", "")
    font_b64 = assets.get("font_b64", "")
    generated = master["generated"]
    db = COLORS["dark_blue"]
    grey = COLORS["grey"]
    red = COLORS["red"]
    card = COLORS["card_bg"]
    border = COLORS["border"]
    grid = COLORS["grid"]

    css = """
@font-face {
    font-family: 'Gotham Book';
    src: url('data:font/opentype;base64,""" + font_b64 + """') format('opentype');
    font-weight: normal; font-style: normal;
}
* { margin:0; padding:0; box-sizing:border-box; }
body {
    font-family: 'Gotham Book', 'Segoe UI', Calibri, sans-serif;
    background: #ffffff; color: """ + db + """; font-size: 13px; line-height: 1.4;
}
.header {
    text-align: center; padding: 20px 20px 10px;
    border-bottom: 2px solid """ + db + """; margin-bottom: 12px;
}
.header img { height: 70px; object-fit: contain; margin-bottom: 8px; }
.header h1 {
    font-size: 18px; font-weight: bold; color: """ + db + """;
    letter-spacing: 1px; text-transform: uppercase;
}

.dataset-toggle-bar {
    display: flex; justify-content: center; align-items: center;
    padding: 8px 20px 14px; gap: 0;
}
.dataset-toggle-btn {
    font-family: 'Gotham Book', 'Segoe UI', Calibri, sans-serif;
    font-size: 14px; font-weight: bold; padding: 9px 30px;
    border: 2px solid """ + db + """; background: #ffffff; color: """ + db + """;
    cursor: pointer; transition: all 0.2s ease;
}
.dataset-toggle-btn:first-child { border-radius: 6px 0 0 6px; border-right: 1px solid """ + db + """; }
.dataset-toggle-btn:last-child  { border-radius: 0 6px 6px 0; border-left: 1px solid """ + db + """; }
.dataset-toggle-btn.active { background: """ + db + """; color: #ffffff; }
.dataset-toggle-btn:hover:not(.active) { background: """ + card + """; }

.controls {
    display: flex; justify-content: center; align-items: center; gap: 20px;
    padding: 14px 24px; background: """ + card + """;
    border: 1px solid """ + border + """; border-radius: 12px;
    margin: 0 20px 8px; flex-wrap: wrap;
    box-shadow: 0 1px 4px rgba(39, 41, 98, 0.06);
}
.control-group { display: flex; align-items: center; gap: 8px; }
.control-group label {
    font-size: 11px; font-weight: bold; text-transform: uppercase;
    color: """ + grey + """; letter-spacing: 0.5px;
}
.control-group select {
    font-family: 'Gotham Book', 'Segoe UI', Calibri, sans-serif;
    font-size: 12px; padding: 8px 32px 8px 14px;
    border: 1.5px solid rgba(39, 41, 98, 0.25); border-radius: 8px;
    color: """ + db + """; background: #ffffff; cursor: pointer;
    appearance: none; -webkit-appearance: none;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%23272962'/%3E%3C/svg%3E");
    background-repeat: no-repeat; background-position: right 10px center;
    transition: border-color 0.2s, box-shadow 0.2s;
}
.control-group select:hover { border-color: """ + db + """; }
.control-group select:focus {
    outline: none; border-color: """ + db + """;
    box-shadow: 0 0 0 3px rgba(39, 41, 98, 0.12);
}
.reset-btn {
    font-family: 'Gotham Book', 'Segoe UI', Calibri, sans-serif;
    font-size: 11px; padding: 7px 16px; border-radius: 8px;
    border: 1.5px solid rgba(39, 41, 98, 0.25); background: #fff;
    color: """ + db + """; cursor: pointer; transition: all 0.2s;
}
.reset-btn:hover { background: """ + card + """; border-color: """ + db + """; }
.period-note {
    text-align: center; font-size: 10px; color: """ + grey + """;
    margin: 0 20px 8px; font-style: italic;
}

.table-container {
    margin: 0 20px 10px; overflow-x: auto;
    border: 1px solid """ + border + """; border-radius: 10px;
    max-height: calc(100vh - 380px); overflow-y: auto;
    box-shadow: 0 1px 4px rgba(39, 41, 98, 0.06);
}
table { border-collapse: collapse; width: max-content; min-width: 100%; }
thead th {
    position: sticky; top: 0; z-index: 10;
    background: """ + db + """; color: #ffffff;
    font-size: 11px; font-weight: normal; padding: 10px 12px;
    text-align: right; white-space: nowrap;
    border-bottom: 2px solid """ + db + """; letter-spacing: 0.3px;
}
thead th:first-child {
    text-align: left; position: sticky; left: 0; z-index: 20;
    min-width: 250px; background: """ + db + """;
    border-top-left-radius: 9px;
}
thead th:last-child { border-top-right-radius: 9px; }
tbody td {
    padding: 6px 12px; text-align: right; font-size: 12px;
    border-bottom: 1px solid """ + grid + """;
    white-space: nowrap; font-variant-numeric: tabular-nums;
}
tbody td:first-child {
    text-align: left; position: sticky; left: 0; z-index: 5;
    background: #ffffff; border-right: 1px solid """ + grid + """;
    font-size: 12px; min-width: 250px;
}
tr.parent-row td { font-weight: bold; background: #fbfbfd; }
tr.parent-row td:first-child {
    background: #fbfbfd; cursor: pointer; user-select: none;
}
tr.parent-row td:first-child:hover { color: """ + red + """; }
tr.child-row td { font-weight: normal; font-size: 11.5px; color: #444466; }
tr.child-row td:first-child { padding-left: 32px; }
tr.child-row.hidden { display: none; }
tr.standalone-row td { font-weight: bold; background: #fbfbfd; }
tr.standalone-row td:first-child { background: #fbfbfd; }
.toggle-arrow {
    display: inline-block; width: 14px; font-size: 10px;
    color: """ + grey + """; transition: transform 0.15s ease;
}
.toggle-arrow.expanded { transform: rotate(90deg); }
tr.pct-row td { font-style: italic; color: """ + grey + """; }
tr.pct-row td:first-child { background: #fbfbfd; }
td.col-highlight { background-color: rgba(39, 41, 98, 0.07) !important; }
td.row-highlight { background-color: rgba(39, 41, 98, 0.07) !important; }
td.cell-highlight { background-color: rgba(39, 41, 98, 0.13) !important; }

/* Growth table */
.growth-section { margin: 0 20px; }
.growth-controls {
    display: flex; justify-content: center; align-items: center; gap: 16px;
    padding: 10px 20px; background: """ + card + """;
    border: 1px solid """ + border + """; border-radius: 10px 10px 0 0;
    border-bottom: none; flex-wrap: wrap;
}
.growth-controls .control-group label { font-size: 10px; }
.growth-toggle { display: flex; gap: 0; }
.growth-toggle button {
    font-family: 'Gotham Book', 'Segoe UI', Calibri, sans-serif;
    font-size: 11px; padding: 6px 14px; border: 1.5px solid rgba(39, 41, 98, 0.25);
    background: #fff; color: """ + db + """; cursor: pointer; transition: all 0.2s;
}
.growth-toggle button:first-child { border-radius: 7px 0 0 7px; }
.growth-toggle button:last-child { border-radius: 0 7px 7px 0; border-left: none; }
.growth-toggle button.active {
    background: """ + db + """; color: #fff; border-color: """ + db + """;
}
.growth-table-container {
    overflow-x: auto; border: 1px solid """ + border + """;
    border-radius: 0 0 10px 10px;
    max-height: 400px; overflow-y: auto;
    box-shadow: 0 1px 4px rgba(39, 41, 98, 0.06);
    margin-bottom: 15px;
}
.growth-table-container table { border-collapse: collapse; width: max-content; min-width: 100%; }
.growth-table-container thead th {
    position: sticky; top: 0; z-index: 10;
    background: """ + db + """; color: #ffffff;
    font-size: 11px; font-weight: normal; padding: 8px 12px;
    text-align: right; white-space: nowrap;
    border-bottom: 2px solid """ + db + """; letter-spacing: 0.3px;
}
.growth-table-container thead th:first-child {
    text-align: left; position: sticky; left: 0; z-index: 20;
    min-width: 250px; background: #3a3d70;
    border-top-left-radius: 0;
}
.growth-table-container tbody td {
    padding: 5px 12px; text-align: right; font-size: 11.5px;
    border-bottom: 1px solid """ + grid + """;
    white-space: nowrap; font-variant-numeric: tabular-nums;
}
.growth-table-container tbody td:first-child {
    text-align: left; position: sticky; left: 0; z-index: 5;
    background: #ffffff; border-right: 1px solid """ + grid + """;
    font-size: 11.5px; min-width: 250px; font-weight: bold;
}
.growth-table-container tr.child-row td:first-child {
    font-weight: normal; padding-left: 32px; color: #444466;
}
.growth-table-container tr.child-row.hidden { display: none; }
td.g-pos { color: #0C7A1E; }
td.g-neg { color: #C00000; }

.table-container::-webkit-scrollbar, .growth-table-container::-webkit-scrollbar { height: 8px; width: 8px; }
.table-container::-webkit-scrollbar-track, .growth-table-container::-webkit-scrollbar-track { background: #f0f0f4; border-radius: 4px; }
.table-container::-webkit-scrollbar-thumb, .growth-table-container::-webkit-scrollbar-thumb { background: """ + grey + """; border-radius: 4px; }
.footer { text-align: center; padding: 12px; font-size: 10px; color: """ + grey + """; }
.unit-label { font-size: 10px; color: rgba(255,255,255,0.6); font-style: italic; margin-left: 6px; }
@media (max-width: 1024px) {
    .header { padding: 16px 16px 8px; }
    .header img { height: 55px; }
    .header h1 { font-size: 16px; }
    .dataset-toggle-btn { font-size: 13px; padding: 8px 22px; }
    .controls { margin: 0 12px 8px; padding: 12px 18px; gap: 14px; }
    .table-container, .growth-section { margin-left: 12px; margin-right: 12px; }
    thead th { font-size: 10.5px; padding: 8px 10px; }
    tbody td { font-size: 11.5px; padding: 5px 10px; }
    thead th:first-child, tbody td:first-child { min-width: 200px; }
    tr.child-row td:first-child { padding-left: 26px; }
}
@media (max-width: 768px) {
    .header { padding: 14px 12px 8px; margin-bottom: 8px; }
    .header img { height: 45px; margin-bottom: 6px; }
    .header h1 { font-size: 14px; letter-spacing: 0.5px; }
    .dataset-toggle-bar { padding: 6px 12px 10px; }
    .dataset-toggle-btn { font-size: 12px; padding: 7px 18px; }
    .controls {
        flex-direction: column; gap: 10px; padding: 10px 14px;
        margin: 0 8px 8px; border-radius: 10px;
    }
    .control-group { width: 100%; justify-content: space-between; }
    .control-group select { flex: 1; font-size: 13px; padding: 10px 32px 10px 12px; }
    .control-group label { font-size: 10px; min-width: 50px; }
    .table-container, .growth-section { margin-left: 8px; margin-right: 8px; }
    .table-container { max-height: calc(100vh - 340px); -webkit-overflow-scrolling: touch; }
    table { font-size: 11px; }
    thead th { font-size: 10px; padding: 8px 8px; }
    tbody td { font-size: 11px; padding: 5px 8px; }
    thead th:first-child, tbody td:first-child { min-width: 150px; font-size: 10.5px; }
    tr.child-row td:first-child { padding-left: 22px; }
    .growth-controls { flex-direction: column; gap: 8px; padding: 8px 12px; }
    .footer { font-size: 9px; padding: 10px 8px; }
    .unit-label { display: none; }
}
@media (max-width: 480px) {
    .header { padding: 10px 8px 6px; margin-bottom: 6px; }
    .header img { height: 36px; margin-bottom: 4px; }
    .header h1 { font-size: 12px; }
    .dataset-toggle-btn { font-size: 11px; padding: 6px 14px; }
    .controls { margin: 0 6px 6px; padding: 8px 10px; gap: 8px; border-radius: 8px; }
    .control-group select { font-size: 12px; padding: 8px 28px 8px 10px; }
    .table-container, .growth-section { margin-left: 6px; margin-right: 6px; }
    thead th { font-size: 9px; padding: 6px 6px; }
    tbody td { font-size: 10px; padding: 4px 6px; }
    thead th:first-child, tbody td:first-child { min-width: 120px; font-size: 9.5px; }
    tr.child-row td:first-child { padding-left: 18px; }
    .toggle-arrow { width: 10px; font-size: 8px; }
}

/* ============================================================
   Multi-pane layout (table + charts)
   ============================================================ */
.main-grid {
    display: grid;
    grid-template-columns: 1fr;
    grid-template-areas:
        "table"
        "growth";
    gap: 16px;
    margin: 0 20px 12px;
}
body.dataset-power .main-grid {
    grid-template-columns: 1fr 1fr;
    grid-template-rows: auto auto;
    grid-template-areas:
        "table   seasonality"
        "growth  buildup";
}
.quad-table       { grid-area: table; }
.quad-growth      { grid-area: growth; }
.quad-seasonality { grid-area: seasonality; }
.quad-buildup     { grid-area: buildup; }
.grid-quad { display: flex; flex-direction: column; min-width: 0; }
.grid-quad .table-container,
.grid-quad .growth-table-container {
    margin: 0; flex: 1;
}
.grid-quad .growth-section { margin: 0; }

/* Charts are hidden unless the active dataset declares charts_enabled */
.chart-quadrant { display: none; }
body.dataset-power .chart-quadrant { display: flex; }

/* Value-mode toggle above the generation table (Power only) */
.table-value-toggle-bar { display: none; padding: 0 0 8px; }
body.dataset-power .table-value-toggle-bar { display: flex; justify-content: flex-start; }
.value-toggle { display: flex; gap: 0; }
.value-toggle button {
    font-family: 'Gotham Book', 'Segoe UI', Calibri, sans-serif;
    font-size: 11px; padding: 6px 14px;
    border: 1.5px solid rgba(39, 41, 98, 0.25);
    background: #fff; color: """ + db + """;
    cursor: pointer; transition: all 0.15s;
}
.value-toggle button:first-child { border-radius: 7px 0 0 7px; }
.value-toggle button:last-child  { border-radius: 0 7px 7px 0; border-left: none; }
.value-toggle button.active { background: """ + db + """; color: #fff; border-color: """ + db + """; }

.unit-efficiency-note {
    font-size: 10px; color: """ + grey + """; font-style: italic;
    margin-left: 2px;
}

.chart-controls {
    display: flex; align-items: center; gap: 14px;
    padding: 10px 14px; background: """ + card + """;
    border: 1px solid """ + border + """; border-radius: 10px 10px 0 0;
    border-bottom: none; flex-wrap: wrap;
    box-shadow: 0 1px 4px rgba(39, 41, 98, 0.06);
}
.chart-controls .control-group label { font-size: 10px; }
.chart-controls .control-group select {
    font-size: 11.5px; padding: 6px 28px 6px 12px;
}

.chart-title {
    text-align: center; font-size: 11px; font-weight: bold;
    color: """ + db + """; letter-spacing: 0.5px; text-transform: uppercase;
    padding: 8px 12px 4px;
    background: #ffffff;
    border-left: 1px solid """ + border + """;
    border-right: 1px solid """ + border + """;
}

.chart-canvas-wrap {
    position: relative;
    flex: 1;
    min-height: 360px;
    padding: 10px 12px 6px;
    background: #ffffff;
    border: 1px solid """ + border + """;
    border-top: none;
    border-bottom: none;
    box-shadow: 0 1px 4px rgba(39, 41, 98, 0.06);
}
.chart-canvas-wrap canvas { width: 100% !important; height: 100% !important; }

/* Custom HTML legend (replaces Chart.js native to control fade-on-hidden) */
.custom-legend {
    display: flex; flex-wrap: wrap; justify-content: center;
    gap: 6px 16px; padding: 10px 14px 12px;
    background: #ffffff;
    border: 1px solid """ + border + """;
    border-top: none;
    border-radius: 0 0 10px 10px;
    box-shadow: 0 1px 4px rgba(39, 41, 98, 0.06);
    font-family: 'Gotham Book', 'Segoe UI', Calibri, sans-serif;
    font-size: clamp(10px, 0.85vw, 13px);
    color: """ + db + """;
}
.cl-item {
    display: inline-flex; align-items: center; gap: 6px;
    cursor: pointer; user-select: none;
    opacity: 1;
    transition: opacity 0.18s;
    padding: 2px 4px;
}
.cl-item:hover { opacity: 0.85; }
.cl-item.cl-hidden { opacity: 0.28; }
.cl-marker {
    display: inline-block; flex-shrink: 0;
    width: 18px; height: 12px;
}
.cl-marker-box { border-radius: 2px; border: 1px solid rgba(39,41,98,0.18); }
.cl-marker-line {
    height: 3px; border-radius: 2px;
    align-self: center;
    background-color: var(--ml, currentColor);
}
.cl-marker-line.cl-marker-dashed {
    background-color: transparent;
    background-image: repeating-linear-gradient(
        90deg,
        var(--ml, currentColor) 0 4px,
        transparent 4px 7px
    );
}
.cl-text { white-space: nowrap; }

/* ============================================================
   Multi-select dropdown widget (Sources picker)
   ============================================================ */
.multi-select {
    position: relative; display: inline-block;
    font-family: 'Gotham Book', 'Segoe UI', Calibri, sans-serif;
}
.ms-button {
    font-family: inherit;
    font-size: 11.5px; padding: 6px 30px 6px 12px;
    border: 1.5px solid rgba(39, 41, 98, 0.25); border-radius: 8px;
    background: #ffffff; color: """ + db + """;
    cursor: pointer; min-width: 180px; text-align: left;
    appearance: none; -webkit-appearance: none;
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6'%3E%3Cpath d='M0 0l5 6 5-6z' fill='%23272962'/%3E%3C/svg%3E");
    background-repeat: no-repeat; background-position: right 10px center;
    transition: border-color 0.15s, box-shadow 0.15s;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.ms-button:hover { border-color: """ + db + """; }
.ms-button:focus { outline: none; border-color: """ + db + """; box-shadow: 0 0 0 3px rgba(39, 41, 98, 0.12); }

.ms-panel {
    display: none; position: absolute; top: calc(100% + 4px); left: 0;
    min-width: 220px; background: #ffffff;
    border: 1.5px solid rgba(39, 41, 98, 0.25); border-radius: 8px;
    box-shadow: 0 4px 14px rgba(39, 41, 98, 0.15);
    z-index: 100; padding: 6px;
    max-height: 320px; overflow-y: auto;
}
.multi-select.open .ms-panel { display: block; }
.ms-item {
    display: flex; align-items: center; gap: 8px;
    padding: 6px 10px; cursor: pointer;
    font-size: 12px; color: """ + db + """;
    border-radius: 6px; user-select: none;
}
.ms-item:hover { background: """ + card + """; }
.ms-item input[type="checkbox"] {
    width: 14px; height: 14px; cursor: pointer; accent-color: """ + db + """;
}
.ms-item.ms-total {
    border-top: 1px solid """ + border + """;
    margin-top: 4px; padding-top: 8px;
    font-weight: bold;
}
.ms-item.ms-disabled { opacity: 0.4; cursor: not-allowed; }
.ms-item.ms-disabled input { cursor: not-allowed; }

/* Responsive: stack 2x2 grid on narrow screens */
@media (max-width: 1100px) {
    body.dataset-power .main-grid {
        grid-template-columns: 1fr;
        grid-template-rows: auto auto auto auto;
        grid-template-areas:
            "table"
            "growth"
            "seasonality"
            "buildup";
    }
    .chart-canvas-wrap { min-height: 320px; }
}
@media (max-width: 768px) {
    .main-grid { margin: 0 8px 8px; gap: 10px; }
    .chart-controls { padding: 8px 10px; gap: 10px; }
    .ms-button { font-size: 12px; padding: 8px 28px 8px 10px; min-width: 150px; }
    .chart-canvas-wrap { min-height: 280px; padding: 8px; }
    .chart-title { font-size: 10px; padding: 6px 10px 3px; }
}
"""

    js = r"""
var ROOT = __DATA_PLACEHOLDER__;
var ALL_DATASETS = ROOT.datasets;
var DATASET_ORDER = ROOT.order;
var currentKey = DATASET_ORDER[0];
var DATA = ALL_DATASETS[currentKey];

var expandedMain = {};
var expandedGrowth = {};
var highlightedCol = -1, highlightedRow = -1;
var growthMode = 'pct';
var growthType = 'yoy';

var GROWTH_OPTS = {
    'Monthly':   [{k:'yoy',l:'Year-on-Year'},{k:'mom',l:'Month-on-Month'},{k:'ytd',l:'YTD Year-on-Year'},{k:'ltm',l:'Last 12 Months'},{k:'ytd_gy',l:'YTD Gas Year Year-on-Year'}],
    'Quarterly': [{k:'yoy',l:'Year-on-Year'},{k:'qoq',l:'Quarter-on-Quarter'},{k:'ytd',l:'YTD Year-on-Year'},{k:'ytd_gy',l:'YTD Gas Year Year-on-Year'}],
    'Annual CY': [{k:'yoy',l:'Year-on-Year'}],
    'Gas Year':  [{k:'yoy',l:'Year-on-Year'}],
    'Winter':    [{k:'yoy',l:'Year-on-Year'}],
    'Summer':    [{k:'yoy',l:'Year-on-Year'}]
};

function unitCfg(unitKey) { return DATA.unit_config[unitKey] || DATA.unit_config[DATA.default_unit]; }
function isVolUnit() { return !unitCfg(document.getElementById('unitSelector').value).isRate; }

function formatNum(val, isPct) {
    if (isPct) return (val*100).toFixed(1)+'%';
    var a = Math.abs(val);
    if (a < 0.005) return '0';
    if (a >= 100) return Math.round(val).toString().replace(/\B(?=(\d{3})+(?!\d))/g,',');
    if (a >= 10) return val.toFixed(1);
    return val.toFixed(2);
}

function formatGrowth(val, isPctRow) {
    if (val===null||val===undefined||!isFinite(val)) return '';
    if (growthMode==='pct') {
        var pct = val*100;
        var a = Math.abs(pct);
        if (a < 0.05) return '0.0%';
        if (a >= 100) return Math.round(pct)+'%';
        return pct.toFixed(1)+'%';
    }
    if (isPctRow) {
        var pp = val*100;
        var a = Math.abs(pp);
        if (a < 0.05) return '0.0 pp';
        return pp.toFixed(1)+' pp';
    }
    var cfg = unitCfg(document.getElementById('unitSelector').value);
    var converted = cfg.isRate ? (val * cfg.rateFactor) : (val * cfg.volFactor);
    return formatNum(converted, false);
}

function computeDisplayValue(baseVal, unitKey, isStock, isPct, days) {
    if (isPct) return baseVal;
    var cfg = unitCfg(unitKey);
    if (isStock) return baseVal * cfg.volFactor;
    if (cfg.isRate) return (baseVal / days) * cfg.rateFactor;
    return baseVal * cfg.volFactor;
}

function getStockLabel(label, unitKey) {
    var cfg = unitCfg(unitKey);
    if (cfg.isRate) return label + ' (' + cfg.volLabel + ')';
    return label;
}

function getHeaderUnitLabel(unitKey) {
    var cfg = unitCfg(unitKey);
    return cfg.isRate ? cfg.rateLabel : cfg.volLabel;
}

/* === DATE RANGE === */
function getSelectableIndices(view) {
    var m = view.col_meta, out = [];
    for (var i=0;i<m.length;i++) {
        if (m[i].year >= DATA.selectable_start && m[i].year <= DATA.selectable_end) out.push(i);
    }
    return out;
}

function getDefaultRange(viewName, selectable, view) {
    if (viewName==='Monthly') {
        var now = new Date(), cy = now.getFullYear();
        var sy = cy-1, ey = cy+1, m = view.col_meta;
        var s=null, e=null;
        for (var i=0;i<selectable.length;i++) {
            var idx=selectable[i];
            if (m[idx].year >= sy && s===null) s=i;
            if (m[idx].year <= ey) e=i;
        }
        return [s!==null?s:0, e!==null?e:selectable.length-1];
    }
    return [0, selectable.length-1];
}

function updateRangeSelectors() {
    var period = document.getElementById('periodSelector').value;
    var view = DATA.views[period];
    var sel = getSelectableIndices(view);
    var fs = document.getElementById('rangeFrom');
    var ts = document.getElementById('rangeTo');
    fs.innerHTML=''; ts.innerHTML='';
    for (var i=0;i<sel.length;i++) {
        var lbl = view.col_meta[sel[i]].label;
        fs.innerHTML += '<option value="'+i+'">'+lbl+'</option>';
        ts.innerHTML += '<option value="'+i+'">'+lbl+'</option>';
    }
    var def = getDefaultRange(period, sel, view);
    fs.value = def[0]; ts.value = def[1];
}

function getVisibleIndices() {
    var period = document.getElementById('periodSelector').value;
    var view = DATA.views[period];
    var sel = getSelectableIndices(view);
    var fi = parseInt(document.getElementById('rangeFrom').value);
    var ti = parseInt(document.getElementById('rangeTo').value);
    if (isNaN(fi)||isNaN(ti)) { var d=getDefaultRange(period,sel,view); fi=d[0]; ti=d[1]; }
    if (fi>ti) { var tmp=fi; fi=ti; ti=tmp; }
    var out = [];
    for (var i=fi;i<=ti;i++) out.push(sel[i]);
    return out;
}

function resetRange() {
    updateRangeSelectors();
    if (DATA.charts_enabled) {
        rebuildChartControls();
    }
    updateAll();
}

function updatePeriodNote() {
    var p = document.getElementById('periodSelector').value;
    var el = document.getElementById('periodNote');
    if (p==='Summer') el.textContent = 'Summer 2025 = April 2025 to September 2025';
    else if (p==='Winter') el.textContent = 'Winter 24/25 = October 2024 to March 2025';
    else el.textContent = '';
}

/* === GROWTH COMPUTATION === */
function getRate(base, days) { return base/days; }
function avgRate(rowBase, days, indices) {
    var tb=0, td=0;
    for (var i=0;i<indices.length;i++) { tb+=rowBase[indices[i]]; td+=days[indices[i]]; }
    return td===0?null:tb/td;
}
function sumVol(rowBase, indices) {
    var s=0;
    for (var i=0;i<indices.length;i++) s+=rowBase[indices[i]];
    return s;
}

function findYoY(meta, idx, vn) {
    var c=meta[idx];
    for (var i=0;i<meta.length;i++) {
        if (i===idx) continue;
        var m=meta[i];
        if (vn==='Monthly' && m.year===c.year-1 && m.month===c.month) return i;
        if (vn==='Quarterly' && m.year===c.year-1 && m.quarter===c.quarter) return i;
        if ((vn==='Annual CY'||vn==='Gas Year'||vn==='Winter'||vn==='Summer') && m.year===c.year-1) return i;
    }
    return -1;
}
function getYTDIndices(meta, idx) {
    var c=meta[idx], out=[];
    for (var i=0;i<meta.length;i++) { if (meta[i].year===c.year && meta[i].month<=c.month) out.push(i); }
    return out;
}
function getYTDQIndices(meta, idx) {
    var c=meta[idx], out=[];
    for (var i=0;i<meta.length;i++) { if (meta[i].year===c.year && meta[i].quarter<=c.quarter) out.push(i); }
    return out;
}
function getYTDGYIndices(meta, idx) {
    var c=meta[idx], gy=c.gas_year, out=[];
    for (var i=0;i<meta.length;i++) { if (meta[i].gas_year===gy && i<=idx) out.push(i); }
    return out;
}
function getLTMIndices(idx) {
    if (idx<11) return [];
    var out=[];
    for (var i=idx-11;i<=idx;i++) out.push(i);
    return out;
}

function computeGrowthCell(rowBase, days, meta, idx, vn, gt, isStock) {
    var useVol = (growthMode==='abs' && isVolUnit() && !isStock);
    function getVal(i) {
        if (isStock) return rowBase[i];
        if (growthMode==='pct') return getRate(rowBase[i],days[i]);
        return useVol ? rowBase[i] : getRate(rowBase[i],days[i]);
    }
    function getAgg(indices) {
        if (isStock) return rowBase[indices[indices.length-1]];
        if (growthMode==='pct') return avgRate(rowBase,days,indices);
        return useVol ? sumVol(rowBase,indices) : avgRate(rowBase,days,indices);
    }
    function result(cur,prev) {
        if (prev===null||cur===null) return null;
        if (growthMode==='pct') return prev===0?null:(cur/prev-1);
        return cur-prev;
    }

    if (gt==='yoy') {
        var pi=findYoY(meta,idx,vn);
        return pi<0?null:result(getVal(idx),getVal(pi));
    }
    if (gt==='mom' || gt==='qoq') {
        return idx<=0?null:result(getVal(idx),getVal(idx-1));
    }
    if (gt==='ytd') {
        var curI,prevI;
        if (vn==='Monthly') {
            curI=getYTDIndices(meta,idx); var c=meta[idx]; prevI=[];
            for(var i=0;i<meta.length;i++){if(meta[i].year===c.year-1&&meta[i].month<=c.month)prevI.push(i);}
        } else {
            curI=getYTDQIndices(meta,idx); var c=meta[idx]; prevI=[];
            for(var i=0;i<meta.length;i++){if(meta[i].year===c.year-1&&meta[i].quarter<=c.quarter)prevI.push(i);}
        }
        if(curI.length===0||prevI.length<curI.length) return null;
        return result(getAgg(curI),getAgg(prevI));
    }
    if (gt==='ltm') {
        var curI=getLTMIndices(idx),prevI=getLTMIndices(idx-12);
        if(curI.length<12||prevI.length<12) return null;
        return result(getAgg(curI),getAgg(prevI));
    }
    if (gt==='ytd_gy') {
        var curI=getYTDGYIndices(meta,idx);
        if(curI.length===0) return null;
        var prevGY=meta[idx].gas_year-1,prevAll=[];
        for(var i=0;i<meta.length;i++){if(meta[i].gas_year===prevGY)prevAll.push(i);}
        var prevI=prevAll.slice(0,curI.length);
        if(prevI.length<curI.length) return null;
        return result(getAgg(curI),getAgg(prevI));
    }
    return null;
}

function updateGrowthTypeSelector() {
    var p = document.getElementById('periodSelector').value;
    var sel = document.getElementById('growthTypeSelector');
    var opts = GROWTH_OPTS[p] || [{k:'yoy',l:'Year-on-Year'}];
    sel.innerHTML = '';
    for (var i=0;i<opts.length;i++) {
        sel.innerHTML += '<option value="'+opts[i].k+'"'+(i===0?' selected':'')+'>'+opts[i].l+'</option>';
    }
    growthType = opts[0].k;
}

function setGrowthMode(mode) {
    growthMode = mode;
    document.getElementById('btnPct').className = mode==='pct'?'active':'';
    document.getElementById('btnAbs').className = mode==='abs'?'active':'';
    updateGrowthTable();
}

function rebuildUnitSelector() {
    var sel = document.getElementById('unitSelector');
    sel.innerHTML = '';
    for (var i=0;i<DATA.units.length;i++) {
        var u = DATA.units[i];
        var selected = (u === DATA.default_unit) ? ' selected' : '';
        sel.innerHTML += '<option value="'+u+'"'+selected+'>'+u+'</option>';
    }
    updateTableValueModeLabels();
}

// Value-mode toggle for the table (Generation vs % of Total Fuel)
var tableValueMode = 'gen'; // 'gen' or 'pct'

function setTableValueMode(mode) {
    tableValueMode = mode;
    var bGen = document.getElementById('btnTblGen');
    var bPct = document.getElementById('btnTblPct');
    if (bGen) bGen.className = mode === 'gen' ? 'active' : '';
    if (bPct) bPct.className = mode === 'pct' ? 'active' : '';
    updateTable();
}

function updateTableValueModeLabels() {
    var btn = document.getElementById('btnTblGen');
    if (!btn) return;
    var unitSel = document.getElementById('unitSelector');
    if (!unitSel || !unitSel.value) return;
    var unitLabel = getHeaderUnitLabel(unitSel.value);
    btn.textContent = 'Generation (' + unitLabel + ')';
}

// Returns ", at 39% efficiency" when a BCFE-family unit is active, else "".
function efficiencySuffix() {
    var unitSel = document.getElementById('unitSelector');
    if (!unitSel || !unitSel.value) return '';
    var u = unitSel.value;
    return (u === 'BCFE' || u === 'BCFE/d') ? ', at 39% efficiency' : '';
}

function updateEfficiencyNote() {
    var noteEl = document.getElementById('unitEfficiencyNote');
    if (!noteEl) return;
    var unitSel = document.getElementById('unitSelector');
    if (!unitSel) { noteEl.textContent = ''; return; }
    var u = unitSel.value;
    if (u === 'BCFE' || u === 'BCFE/d') {
        noteEl.textContent = '(at 39% gas-to-power efficiency)';
    } else {
        noteEl.textContent = '';
    }
}

function updateDatasetUI() {
    document.getElementById('pageTitle').textContent = DATA.title;
    document.title = 'Palissy Advisors - ' + DATA.title;
    document.body.className = 'dataset-' + currentKey;
    for (var i=0;i<DATASET_ORDER.length;i++) {
        var k = DATASET_ORDER[i];
        var btn = document.getElementById('btnDataset-'+k);
        if (btn) btn.classList.toggle('active', k === currentKey);
    }
}

function switchDataset(key) {
    if (key === currentKey) return;
    currentKey = key;
    DATA = ALL_DATASETS[key];
    expandedMain = {};
    expandedGrowth = {};
    updateDatasetUI();
    rebuildUnitSelector();
    updateRangeSelectors();
    updateGrowthTypeSelector();
    updatePeriodNote();
    if (DATA.charts_enabled) {
        rebuildChartControls();
        updateAll();
        updateCharts();
    } else {
        updateAll();
    }
}

/* === MAIN TABLE === */
function updateTable() {
    var period = document.getElementById('periodSelector').value;
    var unitKey = document.getElementById('unitSelector').value;
    var view = DATA.views[period];
    if (!view) return;

    var vis = getVisibleIndices();
    var days = view.days;
    var rows = view.rows;
    var hierarchy = DATA.hierarchy;
    var unitLabel = getHeaderUnitLabel(unitKey);

    // % of Total mode: precompute totals by column index from the total row.
    var showAsPct = (tableValueMode === 'pct') && DATA.total_row;
    var totalRowData = null;
    if (showAsPct) {
        for (var r = 0; r < rows.length; r++) {
            if (rows[r].label === DATA.total_row) { totalRowData = rows[r]; break; }
        }
        if (!totalRowData) showAsPct = false;  // safety
    }

    var thead = document.getElementById('tableHead');
    var headerLabel = showAsPct ? '%' : unitLabel;
    var hHtml = '<tr><th>'+headerLabel+'<span class="unit-label">('+period+')</span></th>';
    for (var i=0;i<vis.length;i++) {
        hHtml += '<th>'+(view.short_columns[vis[i]]||view.columns[vis[i]])+'</th>';
    }
    thead.innerHTML = hHtml+'</tr>';

    var tbody = document.getElementById('tableBody');
    var bHtml = '';
    for (var h=0;h<hierarchy.length;h++) {
        var item=hierarchy[h], rowData=rows[item.index];
        var isExp = expandedMain[item.label]||false;
        var hasCh = item.children&&item.children.length>0;
        var isPct=item.is_pct, isStock=item.is_stock;
        var isTotalRow = (item.label === DATA.total_row);
        var rc = item.type==='standalone'?'standalone-row':'parent-row';
        if (isPct) rc+=' pct-row';
        bHtml+='<tr class="'+rc+'">';
        var lbl='';
        if (hasCh) lbl+='<span class="toggle-arrow'+(isExp?' expanded':'')+'">&#9654;</span> ';
        var dl = isStock?getStockLabel(item.label,unitKey):item.label;
        lbl+=dl.replace(/&/g,'&amp;').replace(/</g,'&lt;');
        bHtml+= hasCh?'<td data-toggle="'+h+'">'+lbl+'</td>':'<td>'+lbl+'</td>';
        for (var i=0;i<vis.length;i++) {
            var ci=vis[i], baseVal=rowData.base[ci];
            var dv;
            if (showAsPct && !isPct && !isStock) {
                if (isTotalRow) {
                    dv = '100%';
                } else {
                    var tot = totalRowData.base[ci];
                    if (!tot || tot === 0) dv = '';
                    else dv = (baseVal / tot * 100).toFixed(1) + '%';
                }
            } else if (isPct) {
                dv = formatNum(baseVal, true);
            } else {
                dv = formatNum(computeDisplayValue(baseVal, unitKey, isStock, false, days[ci]), false);
            }
            bHtml+='<td>'+dv+'</td>';
        }
        bHtml+='</tr>';
        if (hasCh) {
            for (var c=0;c<item.children.length;c++) {
                var ch=item.children[c], cd=rows[ch.index];
                var hid=!isExp?' hidden':'';
                bHtml+='<tr class="child-row'+hid+'">';
                bHtml+='<td>'+ch.label.replace(/&/g,'&amp;').replace(/</g,'&lt;')+'</td>';
                for (var i=0;i<vis.length;i++) {
                    var ci=vis[i];
                    var cdv;
                    if (showAsPct) {
                        var tot = totalRowData.base[ci];
                        if (!tot || tot === 0) cdv = '';
                        else cdv = (cd.base[ci] / tot * 100).toFixed(1) + '%';
                    } else {
                        cdv = formatNum(computeDisplayValue(cd.base[ci], unitKey, false, false, days[ci]), false);
                    }
                    bHtml+='<td>'+cdv+'</td>';
                }
                bHtml+='</tr>';
            }
        }
    }
    tbody.innerHTML = bHtml;
}

/* === GROWTH TABLE === */
function updateGrowthTable() {
    var period = document.getElementById('periodSelector').value;
    var view = DATA.views[period];
    if (!view) return;

    var vis = getVisibleIndices();
    var days = view.days;
    var rows = view.rows;
    var meta = view.col_meta;
    var hierarchy = DATA.hierarchy;
    var gt = document.getElementById('growthTypeSelector').value;
    growthType = gt;

    var thead = document.getElementById('growthHead');
    var hHtml = '<tr><th>Change</th>';
    for (var i=0;i<vis.length;i++) {
        hHtml+='<th>'+(view.short_columns[vis[i]]||view.columns[vis[i]])+'</th>';
    }
    thead.innerHTML = hHtml+'</tr>';

    var tbody = document.getElementById('growthBody');
    var bHtml = '';
    for (var h=0;h<hierarchy.length;h++) {
        var item=hierarchy[h], rowData=rows[item.index];
        var isExp=expandedGrowth[item.label]||false;
        var hasCh=item.children&&item.children.length>0;
        var isPct=item.is_pct, isStock=item.is_stock;
        var rc=isPct?'pct-row':(item.type==='standalone'?'standalone-row':'parent-row');
        bHtml+='<tr class="'+rc+'">';
        var dl='';
        if (hasCh) dl+='<span class="toggle-arrow'+(isExp?' expanded':'')+'">&#9654;</span> ';
        dl+=item.label.replace(/&/g,'&amp;').replace(/</g,'&lt;');
        bHtml+=hasCh?'<td data-toggle-g="'+h+'">'+dl+'</td>':'<td>'+dl+'</td>';
        for (var i=0;i<vis.length;i++) {
            var ci=vis[i];
            var gv = computeGrowthCell(rowData.base,days,meta,ci,period,gt,isStock||isPct);
            if (gv===null) { bHtml+='<td></td>'; continue; }
            var cls = gv>0.0001?'g-pos':(gv<-0.0001?'g-neg':'');
            bHtml+='<td'+(cls?' class="'+cls+'"':'')+'>'+formatGrowth(gv,isPct)+'</td>';
        }
        bHtml+='</tr>';
        if (hasCh) {
            for (var c=0;c<item.children.length;c++) {
                var ch=item.children[c], cd=rows[ch.index];
                var hid=!isExp?' hidden':'';
                bHtml+='<tr class="child-row'+hid+'">';
                bHtml+='<td>'+ch.label.replace(/&/g,'&amp;').replace(/</g,'&lt;')+'</td>';
                for (var i=0;i<vis.length;i++) {
                    var ci=vis[i];
                    var gv=computeGrowthCell(cd.base,days,meta,ci,period,gt,false);
                    if (gv===null) { bHtml+='<td></td>'; continue; }
                    var cls=gv>0.0001?'g-pos':(gv<-0.0001?'g-neg':'');
                    bHtml+='<td'+(cls?' class="'+cls+'"':'')+'>'+formatGrowth(gv,false)+'</td>';
                }
                bHtml+='</tr>';
            }
        }
    }
    tbody.innerHTML = bHtml;
}

/* ============================================================
   CHARTS (seasonality + buildup) - power dataset only
   ============================================================ */
var seasonalityChart = null;
var buildupChart = null;
var multiSelectState = {};

/* ---- Helpers ---- */
function currentGY() {
    var now = new Date();
    var m = now.getMonth() + 1;
    return m >= 10 ? now.getFullYear() : now.getFullYear() - 1;
}

function gyMonthLabels() { return ['Oct','Nov','Dec','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep']; }

function gyMonthToActual(gy, gyMonthIndex) {
    if (gyMonthIndex < 3) return { year: gy, month: 10 + gyMonthIndex };
    return { year: gy + 1, month: gyMonthIndex - 2 };
}

function findMonthlyIdx(year, month) {
    var meta = DATA.views.Monthly.col_meta;
    for (var i = 0; i < meta.length; i++) {
        if (meta[i].year === year && meta[i].month === month) return i;
    }
    return -1;
}

function getRowByLabel(label) {
    var rows = DATA.views.Monthly.rows;
    for (var i = 0; i < rows.length; i++) {
        if (rows[i].label === label) return rows[i];
    }
    return null;
}

function getSourceColor(label) { return DATA.source_colors[label] || '#888'; }

function getIndividualSourceLabels() {
    var labels = [];
    for (var i = 0; i < DATA.hierarchy.length; i++) {
        var item = DATA.hierarchy[i];
        if (item.label !== DATA.total_row) labels.push(item.label);
    }
    return labels;
}

function sumSourcesAtIndex(selectedLabels, monthlyIdx) {
    var sum = 0;
    for (var i = 0; i < selectedLabels.length; i++) {
        var row = getRowByLabel(selectedLabels[i]);
        if (row && monthlyIdx >= 0 && monthlyIdx < row.base.length) sum += row.base[monthlyIdx];
    }
    return sum;
}

function applyUnitConversion(baseValue, unitKey, days) {
    var cfg = DATA.unit_config[unitKey];
    if (!cfg) return baseValue;
    if (cfg.isRate) return (baseValue / days) * cfg.rateFactor;
    return baseValue * cfg.volFactor;
}

/* Color helper: fade a hex/rgba color to a given alpha (0-1). */
function fadeColor(color, alpha) {
    if (!color || typeof color !== 'string') return color;
    if (color.charAt(0) === '#') {
        var hex = color.slice(1);
        if (hex.length === 8) hex = hex.slice(0, 6);
        if (hex.length === 6) {
            var r = parseInt(hex.slice(0,2), 16);
            var g = parseInt(hex.slice(2,4), 16);
            var b = parseInt(hex.slice(4,6), 16);
            return 'rgba(' + r + ',' + g + ',' + b + ',' + alpha + ')';
        }
    }
    if (color.indexOf('rgba(') === 0) {
        return color.replace(/[\d.]+\)$/, alpha + ')');
    }
    if (color.indexOf('rgb(') === 0) {
        return color.replace('rgb(', 'rgba(').replace(')', ',' + alpha + ')');
    }
    return color;
}

/* Responsive font scale: charts get bigger labels when the canvas is wider. */
function chartFontScale(width) {
    if (!width || isNaN(width)) return 1.0;
    if (width < 500) return 1.0;
    if (width < 700) return 1.12;
    if (width < 1000) return 1.28;
    if (width < 1400) return 1.5;
    return 1.7;
}

function scaledFont(baseSize, fam) {
    return function(ctx) {
        var w = (ctx && ctx.chart && ctx.chart.width) ? ctx.chart.width : 600;
        return { size: Math.round(baseSize * chartFontScale(w)), family: fam };
    };
}

/* ---- Multi-select widget ---- */
function buildMultiSelect(id, items, totalLabel, defaultSelected, onChange) {
    var container = document.getElementById(id);
    if (!container) return;
    container.innerHTML = '';
    container.classList.add('multi-select');
    multiSelectState[id] = defaultSelected.slice();

    var button = document.createElement('button');
    button.type = 'button';
    button.className = 'ms-button';
    container.appendChild(button);

    var panel = document.createElement('div');
    panel.className = 'ms-panel';
    container.appendChild(panel);

    function addItem(value, isTotal) {
        var lbl = document.createElement('label');
        lbl.className = 'ms-item' + (isTotal ? ' ms-total' : '');
        var cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = value;
        var sp = document.createElement('span');
        sp.textContent = value;
        lbl.appendChild(cb);
        lbl.appendChild(sp);
        panel.appendChild(lbl);
        cb.addEventListener('change', function(e) { handleMSChange(id, totalLabel, onChange, e.target.value); });
    }

    for (var i = 0; i < items.length; i++) addItem(items[i], false);
    if (totalLabel) addItem(totalLabel, true);

    button.addEventListener('click', function(e) {
        e.stopPropagation();
        container.classList.toggle('open');
    });

    applyMSState(id, totalLabel);
    updateMSButtonLabel(id, totalLabel);
}

function handleMSChange(id, totalLabel, onChange, changedValue) {
    var container = document.getElementById(id);
    var inputs = container.querySelectorAll('input[type="checkbox"]');

    // Mutex: if the just-clicked checkbox is Total and was just turned on, clear individuals;
    // if it's an individual and was just turned on, clear Total.
    var changedIsTotal = (changedValue === totalLabel);
    var changedNowOn = false;
    for (var i = 0; i < inputs.length; i++) {
        if (inputs[i].value === changedValue) { changedNowOn = inputs[i].checked; break; }
    }
    if (changedNowOn) {
        if (changedIsTotal) {
            for (var i = 0; i < inputs.length; i++) if (inputs[i].value !== totalLabel) inputs[i].checked = false;
        } else {
            for (var i = 0; i < inputs.length; i++) if (inputs[i].value === totalLabel) inputs[i].checked = false;
        }
    }

    // Collect resulting state
    var state = [];
    for (var i = 0; i < inputs.length; i++) {
        if (inputs[i].checked) state.push(inputs[i].value);
    }
    // Don't allow zero selection — fall back to Total
    if (state.length === 0 && totalLabel) {
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].value === totalLabel) { inputs[i].checked = true; state = [totalLabel]; break; }
        }
    }
    multiSelectState[id] = state;

    updateMSButtonLabel(id, totalLabel);
    if (onChange) onChange(state);
}

function applyMSState(id, totalLabel) {
    var container = document.getElementById(id);
    var inputs = container.querySelectorAll('input[type="checkbox"]');
    var state = multiSelectState[id] || [];
    for (var i = 0; i < inputs.length; i++) {
        inputs[i].checked = state.indexOf(inputs[i].value) >= 0;
    }
}

function updateMSButtonLabel(id, totalLabel) {
    var container = document.getElementById(id);
    var btn = container.querySelector('.ms-button');
    var state = multiSelectState[id] || [];
    var txt = 'None';
    if (state.length === 1) txt = state[0];
    else if (state.length === 2) txt = state.join(' + ');
    else if (state.length > 2) txt = state[0] + ' + ' + (state.length - 1) + ' more';
    btn.textContent = txt;
    btn.title = state.join(', ');
}

function getMultiSelectState(id) { return multiSelectState[id] || []; }

// Outside-click closes any open multi-select panel
document.addEventListener('click', function(e) {
    var open = document.querySelectorAll('.multi-select.open');
    for (var i = 0; i < open.length; i++) {
        if (!open[i].contains(e.target)) open[i].classList.remove('open');
    }
});

/* ---- Seasonality chart ---- */
function updateSeasonalityChart() {
    if (!DATA || !DATA.charts_enabled) return;
    var selected = getMultiSelectState('msSeasonality');
    if (selected.length === 0) selected = [DATA.total_row];

    var unitKey = document.getElementById('unitSelector').value;
    var lookback = parseInt(document.getElementById('seasonalityLookback').value) || 5;
    var cgy = currentGY();
    var prevGY = cgy - 1, nextGY = cgy + 1;
    var labels = gyMonthLabels();

    function gyValues(gy) {
        var out = [];
        for (var i = 0; i < 12; i++) {
            var amt = gyMonthToActual(gy, i);
            var idx = findMonthlyIdx(amt.year, amt.month);
            if (idx < 0) { out.push(null); continue; }
            var b = sumSourcesAtIndex(selected, idx);
            var d = DATA.views.Monthly.days[idx];
            out.push(applyUnitConversion(b, unitKey, d));
        }
        return out;
    }
    var prevVals = gyValues(prevGY);
    var currVals = gyValues(cgy);
    var nextVals = gyValues(nextGY);

    // Lookback window: `lookback` gas years prior to current
    var lbStart = cgy - lookback, lbEnd = cgy - 1;
    var avg = [], mn = [], mx = [];
    for (var mo = 0; mo < 12; mo++) {
        var samples = [];
        for (var gy = lbStart; gy <= lbEnd; gy++) {
            var amt = gyMonthToActual(gy, mo);
            var idx = findMonthlyIdx(amt.year, amt.month);
            if (idx < 0) continue;
            var b = sumSourcesAtIndex(selected, idx);
            var d = DATA.views.Monthly.days[idx];
            samples.push(applyUnitConversion(b, unitKey, d));
        }
        if (samples.length === 0) { avg.push(null); mn.push(null); mx.push(null); continue; }
        var s = 0; for (var k = 0; k < samples.length; k++) s += samples[k];
        avg.push(s / samples.length);
        mn.push(Math.min.apply(null, samples));
        mx.push(Math.max.apply(null, samples));
    }

    var titleSrc = selected.join(' + ');
    document.getElementById('seasonalityTitle').textContent =
        'Seasonality (gas-year cycle) — ' + titleSrc + ' (' + getHeaderUnitLabel(unitKey) + efficiencySuffix() + ')';

    // Order in array drives legend order (after the max filter).
    // We want: range (grey block), average, prev, current, next.
    // The max line must come BEFORE the range line in the dataset array so the
    // range's `fill: '-1'` correctly fills back to it - but we filter it out
    // of the legend.
    var prevLabel = 'GY ' + String(prevGY).slice(-2) + '/' + String(prevGY+1).slice(-2);
    var curLabel  = 'GY ' + String(cgy).slice(-2)    + '/' + String(cgy+1).slice(-2)    + ' (current)';
    var nextLabel = 'GY ' + String(nextGY).slice(-2) + '/' + String(nextGY+1).slice(-2) + ' (forecast)';
    var datasets = [
        // Invisible max line - exists only to anchor the range fill.
        { label: lookback + 'y max',     data: mx,        borderColor: 'rgba(0,0,0,0)', backgroundColor: 'rgba(0,0,0,0)', pointRadius: 0, fill: false, order: 20 },
        // Range = filled grey block, no line. Shown in legend as a colored box.
        { label: lookback + 'y range',   data: mn,        borderColor: 'rgba(0,0,0,0)', backgroundColor: 'rgba(147,149,162,0.28)', pointRadius: 0, fill: '-1', order: 19 },
        // Average = solid dark blue.
        { label: lookback + 'y average', data: avg,       borderColor: '#272962',       backgroundColor: 'rgba(0,0,0,0)', borderWidth: 1.8, pointRadius: 0, fill: false, tension: 0.25, order: 4 },
        // Previous GY = solid dark green.
        { label: prevLabel,              data: prevVals,  borderColor: '#0C5B19',       backgroundColor: 'rgba(0,0,0,0)', borderWidth: 1.8, pointRadius: 2, fill: false, tension: 0.25, order: 3 },
        // Current GY = solid red (most prominent).
        { label: curLabel,               data: currVals,  borderColor: '#C00000',       backgroundColor: 'rgba(0,0,0,0)', borderWidth: 2.6, pointRadius: 3, fill: false, tension: 0.25, order: 1 },
        // Next GY = dashed light green (forecast).
        { label: nextLabel,              data: nextVals,  borderColor: '#539648',       backgroundColor: 'rgba(0,0,0,0)', borderWidth: 1.8, borderDash: [6,4], pointRadius: 2, fill: false, tension: 0.25, order: 2 },
    ];

    if (seasonalityChart) {
        seasonalityChart.data.labels = labels;
        seasonalityChart.data.datasets = datasets;
        seasonalityChart.options.scales.y.title.text = getHeaderUnitLabel(unitKey);
        seasonalityChart.update();
    } else {
        var canvas = document.getElementById('seasonalityCanvas');
        if (!canvas) return;
        seasonalityChart = new Chart(canvas, {
            type: 'line',
            data: { labels: labels, datasets: datasets },
            options: chartOptionsLine(unitKey)
        });
    }

    // Render HTML legend - filter out the invisible "max" anchor dataset.
    renderHtmlLegend(seasonalityChart, 'seasonalityLegend', {
        filter: function(label) { return !(label || '').endsWith(' max'); }
    });
}

function chartOptionsLine(unitKey) {
    var fam = "'Gotham Book','Segoe UI',Calibri,sans-serif";
    return {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
            // Built-in legend disabled - we render our own HTML legend below
            // the canvas so we can fade hidden items without Chart.js drawing
            // the hardcoded strike-through.
            legend: { display: false },
            tooltip: {
                mode: 'index', intersect: false,
                filter: function(ctx) { return !(ctx.dataset.label || '').endsWith(' max'); },
                titleFont: scaledFont(11, fam),
                bodyFont: scaledFont(10.5, fam),
                callbacks: {
                    label: function(ctx) { return ctx.dataset.label + ': ' + (ctx.parsed.y === null || ctx.parsed.y === undefined ? '-' : ctx.parsed.y.toFixed(1)); }
                }
            }
        },
        scales: {
            x: { grid: { display: false }, ticks: { font: scaledFont(10, fam), color: '#272962' } },
            y: { title: { display: true, text: getHeaderUnitLabel(unitKey), font: scaledFont(10.5, fam), color: '#272962' },
                 grid: { color: '#E0E0E8' },
                 ticks: { font: scaledFont(10, fam), color: '#272962' } }
        }
    };
}

/* ---- Buildup chart ---- */
function getBuildupSourcesEffective(rawSelected) {
    if (rawSelected.length === 1 && rawSelected[0] === DATA.total_row) return getIndividualSourceLabels();
    return rawSelected;
}

function updateBuildupChart() {
    if (!DATA || !DATA.charts_enabled) return;
    var rawSelected = getMultiSelectState('msBuildup');
    if (rawSelected.length === 0) rawSelected = [DATA.total_row];
    var sources = getBuildupSourcesEffective(rawSelected);
    var unitKey = document.getElementById('unitSelector').value;
    var agg = document.getElementById('buildupAgg').value;
    var viewKey = agg;
    var view = DATA.views[viewKey];
    if (!view) return;

    var fromIdx = parseInt(document.getElementById('buildupFrom').value);
    var toIdx = parseInt(document.getElementById('buildupTo').value);
    if (isNaN(fromIdx) || isNaN(toIdx)) { fromIdx = 0; toIdx = view.col_meta.length - 1; }
    if (fromIdx > toIdx) { var t = fromIdx; fromIdx = toIdx; toIdx = t; }

    var labels = [];
    for (var i = fromIdx; i <= toIdx; i++) labels.push(view.short_columns[i] || view.columns[i]);

    var datasets = [];
    for (var s = 0; s < sources.length; s++) {
        var label = sources[s];
        var row = null;
        for (var r = 0; r < view.rows.length; r++) { if (view.rows[r].label === label) { row = view.rows[r]; break; } }
        if (!row) continue;
        var data = [];
        for (var i = fromIdx; i <= toIdx; i++) {
            // Clamp negatives to zero: small forecast quirks (e.g. negative
            // values in the Ember "Other Renewables" series) break stacked-area
            // rendering. Generation should not be negative for chart purposes.
            var v = applyUnitConversion(row.base[i], unitKey, view.days[i]);
            data.push(v < 0 ? 0 : v);
        }
        var color = getSourceColor(label);
        datasets.push({
            label: label,
            data: data,
            backgroundColor: color + 'CC',
            borderColor: color,
            borderWidth: 1,
            // For stacked area (line) charts, fill back to the previous
            // dataset's line so each layer renders ONLY as its incremental
            // band, not from 0 up to its cumulative top. Without this, fading
            // others would still leave the hovered dataset's fill spanning
            // 0 -> cumulative top.
            fill: s === 0 ? 'origin' : '-1',
            stack: 'gen',
            tension: 0.2,
            pointRadius: 0,
        });
    }

    var modeLabel = (sources.length === getIndividualSourceLabels().length && rawSelected.length === 1 && rawSelected[0] === DATA.total_row)
        ? 'Total (stacked breakdown)'
        : sources.join(' + ');
    document.getElementById('buildupTitle').textContent =
        'Generation — ' + modeLabel + ' (' + getHeaderUnitLabel(unitKey) + efficiencySuffix() + ')';

    var chartType = (agg === 'Monthly') ? 'line' : 'bar';

    if (buildupChart) { buildupChart.destroy(); buildupChart = null; }
    var canvas = document.getElementById('buildupCanvas');
    if (!canvas) return;
    buildupChart = new Chart(canvas, {
        type: chartType,
        data: { labels: labels, datasets: datasets },
        options: chartOptionsStacked(unitKey)
    });
    initBuildupHover(canvas);
    renderHtmlLegend(buildupChart, 'buildupLegend', {});
}

function chartOptionsStacked(unitKey) {
    var fam = "'Gotham Book','Segoe UI',Calibri,sans-serif";
    return {
        responsive: true, maintainAspectRatio: false,
        interaction: { mode: 'index', intersect: false },
        plugins: {
            // Custom HTML legend - see above for rationale.
            legend: { display: false },
            tooltip: {
                mode: 'index', intersect: false,
                titleFont: scaledFont(11, fam),
                bodyFont: scaledFont(10.5, fam),
                footerFont: scaledFont(10.5, fam),
                callbacks: {
                    label: function(ctx) { return ctx.dataset.label + ': ' + (ctx.parsed.y === null || ctx.parsed.y === undefined ? '-' : ctx.parsed.y.toFixed(1)); },
                    footer: function(items) {
                        var total = 0;
                        for (var i = 0; i < items.length; i++) {
                            var v = items[i].parsed.y;
                            if (typeof v === 'number' && !isNaN(v)) total += v;
                        }
                        return 'Total: ' + total.toFixed(1);
                    }
                }
            }
        },
        scales: {
            x: { stacked: true, grid: { display: false }, ticks: { font: scaledFont(10, fam), color: '#272962', maxRotation: 0, autoSkip: true, autoSkipPadding: 12 } },
            y: {
                stacked: true,
                beginAtZero: true,  // critical: keep all unit views on the same baseline
                title: { display: true, text: getHeaderUnitLabel(unitKey), font: scaledFont(10.5, fam), color: '#272962' },
                grid: { color: '#E0E0E8' },
                ticks: { font: scaledFont(10, fam), color: '#272962' }
            }
        }
    };
}

/* Custom HTML legend renderer.
   Chart.js's native legend has a hardcoded strike-through on hidden items
   that can't be suppressed via standard config. We render our own legend
   below the canvas so we can fade hidden items via CSS opacity. */
function renderHtmlLegend(chart, containerId, opts) {
    opts = opts || {};
    var container = document.getElementById(containerId);
    if (!container) return;
    container.innerHTML = '';

    var datasets = chart.data.datasets;
    for (var i = 0; i < datasets.length; i++) {
        var ds = datasets[i];
        var label = ds.label || '';
        if (opts.filter && !opts.filter(label)) continue;

        // Determine the marker color. For filled areas (range, stacked sources)
        // use the dataset's backgroundColor; for lines use borderColor.
        var bg = ds.backgroundColor;
        var border = ds.borderColor;
        var markerFill, markerBorder;
        // If a transparent background, use border color (line series).
        var bgIsTransparent = (bg === 'rgba(0,0,0,0)') || (typeof bg === 'string' && bg.indexOf('rgba(0,0,0,0)') === 0);
        if (bgIsTransparent || !bg) {
            markerFill = border || '#888';
            markerBorder = border || '#888';
        } else {
            markerFill = bg;
            markerBorder = border || bg;
        }

        var hidden = chart.getDatasetMeta(i).hidden === true;

        var item = document.createElement('div');
        item.className = 'cl-item' + (hidden ? ' cl-hidden' : '');
        item.setAttribute('data-idx', i);

        // Marker style depends on the dataset shape: filled box for areas/bands,
        // line-stripe for line series, dashed line for forecasts.
        var marker = document.createElement('span');
        marker.className = 'cl-marker';
        if (bgIsTransparent) {
            marker.classList.add('cl-marker-line');
            marker.style.setProperty('--ml', markerFill);
            if (ds.borderDash && ds.borderDash.length) marker.classList.add('cl-marker-dashed');
        } else {
            marker.classList.add('cl-marker-box');
            marker.style.backgroundColor = markerFill;
            marker.style.borderColor = markerBorder;
        }
        item.appendChild(marker);

        var text = document.createElement('span');
        text.className = 'cl-text';
        text.textContent = label;
        item.appendChild(text);

        item.addEventListener('click', function() {
            var idx = parseInt(this.getAttribute('data-idx'));
            var meta = chart.getDatasetMeta(idx);
            meta.hidden = meta.hidden === true ? false : true;
            chart.update();
            renderHtmlLegend(chart, containerId, opts);
        });

        container.appendChild(item);
    }
}

/* Hover-highlight on the buildup chart canvas.
   Chart.js's built-in hit detection on stacked AREA charts only fires on
   line points - not in the filled regions. So we compute the hovered
   dataset manually by mapping the cursor's Y position into the cumulative
   stack at the cursor's X column. */
var buildupHoverInited = false;
function initBuildupHover(canvas) {
    if (buildupHoverInited) return;
    buildupHoverInited = true;
    canvas.addEventListener('mousemove', onBuildupMousemove);
    canvas.addEventListener('mouseleave', onBuildupMouseleave);
}

function onBuildupMousemove(e) {
    if (!buildupChart) return;
    var canvas = buildupChart.canvas;
    var rect = canvas.getBoundingClientRect();
    var cx = e.clientX - rect.left;
    var cy = e.clientY - rect.top;

    var chartArea = buildupChart.chartArea;
    if (!chartArea || cx < chartArea.left || cx > chartArea.right || cy < chartArea.top || cy > chartArea.bottom) {
        applyBuildupHover(-1);
        return;
    }

    var xScale = buildupChart.scales.x;
    var yScale = buildupChart.scales.y;
    var labels = buildupChart.data.labels;
    var datasets = buildupChart.data.datasets;

    // Find nearest x category
    var dataIdx = 0, minDist = Infinity;
    for (var i = 0; i < labels.length; i++) {
        var px = xScale.getPixelForValue(i);
        var d = Math.abs(px - cx);
        if (d < minDist) { minDist = d; dataIdx = i; }
    }

    // Walk the cumulative stack at this x to find which dataset the cursor
    // is sitting within (in data-value space, not pixels).
    var cursorYVal = yScale.getValueForPixel(cy);
    var cum = 0;
    var hoveredIdx = -1;
    for (var i = 0; i < datasets.length; i++) {
        if (buildupChart.getDatasetMeta(i).hidden) continue;
        var v = datasets[i].data[dataIdx] || 0;
        var top = cum + v;
        if (cursorYVal >= cum && cursorYVal <= top) {
            hoveredIdx = i;
            break;
        }
        cum = top;
    }
    applyBuildupHover(hoveredIdx);
}

function onBuildupMouseleave() { applyBuildupHover(-1); }

function applyBuildupHover(hoveredIdx) {
    if (!buildupChart) return;
    var changed = false;
    var datasets = buildupChart.data.datasets;
    for (var i = 0; i < datasets.length; i++) {
        var ds = datasets[i];
        if (ds._origBg === undefined) ds._origBg = ds.backgroundColor;
        if (ds._origBorder === undefined) ds._origBorder = ds.borderColor;
        var targetBg, targetBorder;
        if (hoveredIdx === -1 || hoveredIdx === i) {
            targetBg = ds._origBg;
            targetBorder = ds._origBorder;
        } else {
            targetBg = fadeColor(ds._origBg, 0.10);
            targetBorder = fadeColor(ds._origBorder, 0.20);
        }
        if (ds.backgroundColor !== targetBg) { ds.backgroundColor = targetBg; changed = true; }
        if (ds.borderColor !== targetBorder) { ds.borderColor = targetBorder; changed = true; }
    }
    if (changed) buildupChart.update('none');
}

/* ---- Chart-control rebuilders ---- */
function rebuildChartControls() {
    if (!DATA.charts_enabled) return;
    var individuals = getIndividualSourceLabels();
    var totalLabel = DATA.total_row;
    buildMultiSelect('msSeasonality', individuals, totalLabel, [totalLabel], function() { updateSeasonalityChart(); });
    buildMultiSelect('msBuildup',     individuals, totalLabel, [totalLabel], function() { updateBuildupChart(); });
    populateLookbackOptions();
    populateBuildupRangeSelectors();
}

function populateLookbackOptions() {
    var sel = document.getElementById('seasonalityLookback');
    if (!sel) return;
    var cgy = currentGY();
    var firstYear = DATA.views.Monthly.col_meta[0].year;
    var maxLookback = Math.max(3, cgy - firstYear);
    sel.innerHTML = '';
    for (var n = 3; n <= maxLookback; n++) {
        var s = (n === 5) ? ' selected' : '';
        sel.innerHTML += '<option value="' + n + '"' + s + '>' + n + ' years</option>';
    }
}

function populateBuildupRangeSelectors() {
    var aggEl = document.getElementById('buildupAgg');
    if (!aggEl) return;
    var agg = aggEl.value;
    var view = DATA.views[agg];
    if (!view) return;
    var fs = document.getElementById('buildupFrom');
    var ts = document.getElementById('buildupTo');
    fs.innerHTML = ''; ts.innerHTML = '';
    for (var i = 0; i < view.col_meta.length; i++) {
        var lbl = view.col_meta[i].label;
        fs.innerHTML += '<option value="' + i + '">' + lbl + '</option>';
        ts.innerHTML += '<option value="' + i + '">' + lbl + '</option>';
    }
    // Default: clip to admin display range
    var defaultFrom = 0, defaultTo = view.col_meta.length - 1;
    for (var i = 0; i < view.col_meta.length; i++) {
        if (view.col_meta[i].year >= DATA.selectable_start && view.col_meta[i].year <= DATA.selectable_end) { defaultFrom = i; break; }
    }
    for (var i = view.col_meta.length - 1; i >= 0; i--) {
        if (view.col_meta[i].year >= DATA.selectable_start && view.col_meta[i].year <= DATA.selectable_end) { defaultTo = i; break; }
    }
    fs.value = defaultFrom;
    ts.value = defaultTo;
}

function onBuildupAggChange() {
    populateBuildupRangeSelectors();
    updateBuildupChart();
}

function updateCharts() {
    if (!DATA.charts_enabled) return;
    updateSeasonalityChart();
    updateBuildupChart();
}

/* === EVENTS === */
function onPeriodChange() {
    updateRangeSelectors();
    updateGrowthTypeSelector();
    updatePeriodNote();
    updateAll();
}
function onRangeChange() { updateAll(); }
function onGrowthTypeChange() { growthType=document.getElementById('growthTypeSelector').value; updateGrowthTable(); }
function updateAll() {
    updateTableValueModeLabels();
    updateEfficiencyNote();
    updateTable();
    updateGrowthTable();
    if (DATA.charts_enabled) updateCharts();
}

document.addEventListener('click', function(e) {
    var td = e.target.closest('td[data-toggle]');
    if (td) {
        var idx=parseInt(td.getAttribute('data-toggle'));
        var label=DATA.hierarchy[idx].label;
        expandedMain[label]=!expandedMain[label];
        updateTable();
        return;
    }
    var td2 = e.target.closest('td[data-toggle-g]');
    if (td2) {
        var idx=parseInt(td2.getAttribute('data-toggle-g'));
        var label=DATA.hierarchy[idx].label;
        expandedGrowth[label]=!expandedGrowth[label];
        updateGrowthTable();
    }
});

function setupHover(tbodyId) {
    var container = document.getElementById(tbodyId);
    if (!container) return;
    container.addEventListener('mouseover', function(e) {
        var td = e.target.closest('td');
        if (!td) return;
        var tr = td.closest('tr');
        if (!tr) return;
        var colIdx = Array.from(tr.children).indexOf(td);
        var rowIdx = Array.from(container.children).indexOf(tr);
        if (colIdx===highlightedCol && rowIdx===highlightedRow) return;
        clearHighlight();
        highlightedCol=colIdx; highlightedRow=rowIdx;
        var allRows = container.querySelectorAll('tr');
        for (var r=0;r<allRows.length;r++) {
            var cells=allRows[r].children;
            if (colIdx<cells.length) {
                cells[colIdx].classList.add(r===rowIdx?'cell-highlight':'col-highlight');
            }
        }
        for (var c=0;c<tr.children.length;c++) {
            if (c!==colIdx) tr.children[c].classList.add('row-highlight');
        }
    });
    container.addEventListener('mouseleave', function() { clearHighlight(); });
}
function clearHighlight() {
    var els = document.querySelectorAll('.col-highlight,.row-highlight,.cell-highlight');
    for (var i=0;i<els.length;i++) els[i].classList.remove('col-highlight','row-highlight','cell-highlight');
    highlightedCol=-1; highlightedRow=-1;
}

document.addEventListener('DOMContentLoaded', function() {
    updateDatasetUI();
    rebuildUnitSelector();
    updateRangeSelectors();
    updateGrowthTypeSelector();
    updatePeriodNote();
    if (DATA.charts_enabled) rebuildChartControls();
    updateAll();
    setupHover('tableBody');
    setupHover('growthBody');
});
"""

    js = js.replace("__DATA_PLACEHOLDER__", data_json)

    # Build dataset toggle buttons
    toggle_buttons = []
    for k in ordered_keys:
        ds = datasets_by_key[k]
        toggle_buttons.append(
            f'<button class="dataset-toggle-btn{" active" if k == ordered_keys[0] else ""}" '
            f'id="btnDataset-{k}" onclick="switchDataset(\'{k}\')">{ds["tab_label"]}</button>'
        )

    html = '<!DOCTYPE html>\n<html lang="en">\n<head>\n'
    html += '<meta charset="UTF-8">\n'
    html += '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
    html += '<title>Palissy Advisors</title>\n'
    html += '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>\n'
    html += '<style>\n' + css + '\n</style>\n'
    html += '</head>\n<body>\n\n'

    html += '<div class="header">\n'
    html += '    <img src="data:image/png;base64,' + logo_b64 + '" alt="Palissy Advisors">\n'
    html += '    <h1 id="pageTitle"></h1>\n'
    html += '</div>\n\n'

    html += '<div class="dataset-toggle-bar">\n    '
    html += '\n    '.join(toggle_buttons)
    html += '\n</div>\n\n'

    html += '<div class="controls">\n'
    html += '    <div class="control-group">\n'
    html += '        <label>Period</label>\n'
    html += '        <select id="periodSelector" onchange="onPeriodChange()">\n'
    html += '            <option value="Monthly" selected>Monthly</option>\n'
    html += '            <option value="Quarterly">Quarterly</option>\n'
    html += '            <option value="Annual CY">Annual (Calendar Year)</option>\n'
    html += '            <option value="Gas Year">Annual (Gas Year)</option>\n'
    html += '            <option value="Winter">Winter (Oct-Mar)</option>\n'
    html += '            <option value="Summer">Summer (Apr-Sep)</option>\n'
    html += '        </select>\n'
    html += '    </div>\n'
    html += '    <div class="control-group">\n'
    html += '        <label>Unit</label>\n'
    html += '        <select id="unitSelector" onchange="updateAll()"></select>\n'
    html += '        <span class="unit-efficiency-note" id="unitEfficiencyNote"></span>\n'
    html += '    </div>\n'
    html += '    <div class="control-group">\n'
    html += '        <label>From</label>\n'
    html += '        <select id="rangeFrom" onchange="onRangeChange()"></select>\n'
    html += '    </div>\n'
    html += '    <div class="control-group">\n'
    html += '        <label>To</label>\n'
    html += '        <select id="rangeTo" onchange="onRangeChange()"></select>\n'
    html += '    </div>\n'
    html += '    <button class="reset-btn" onclick="resetRange()">Reset</button>\n'
    html += '</div>\n'
    html += '<div class="period-note" id="periodNote"></div>\n\n'

    html += '<div class="main-grid">\n'

    # === Top-left: table ===
    html += '<div class="grid-quad quad-table">\n'
    html += '    <div class="table-value-toggle-bar">\n'
    html += '        <div class="value-toggle">\n'
    html += '            <button id="btnTblGen" class="active" onclick="setTableValueMode(\'gen\')">Generation</button>\n'
    html += '            <button id="btnTblPct" onclick="setTableValueMode(\'pct\')">% of Total Fuel</button>\n'
    html += '        </div>\n'
    html += '    </div>\n'
    html += '    <div class="table-container" id="tableContainer">\n'
    html += '        <table id="dataTable">\n'
    html += '            <thead id="tableHead"></thead>\n'
    html += '            <tbody id="tableBody"></tbody>\n'
    html += '        </table>\n'
    html += '    </div>\n'
    html += '</div>\n\n'

    # === Bottom-left: growth table ===
    html += '<div class="grid-quad quad-growth">\n'
    html += '    <div class="growth-section">\n'
    html += '        <div class="growth-controls">\n'
    html += '            <div class="control-group">\n'
    html += '                <label>Change</label>\n'
    html += '                <select id="growthTypeSelector" onchange="onGrowthTypeChange()"></select>\n'
    html += '            </div>\n'
    html += '            <div class="growth-toggle">\n'
    html += '                <button id="btnPct" class="active" onclick="setGrowthMode(\'pct\')">% Change</button>\n'
    html += '                <button id="btnAbs" onclick="setGrowthMode(\'abs\')">Absolute</button>\n'
    html += '            </div>\n'
    html += '        </div>\n'
    html += '        <div class="growth-table-container">\n'
    html += '            <table>\n'
    html += '                <thead id="growthHead"></thead>\n'
    html += '                <tbody id="growthBody"></tbody>\n'
    html += '            </table>\n'
    html += '        </div>\n'
    html += '    </div>\n'
    html += '</div>\n\n'

    # === Top-right: seasonality chart (only shown for datasets with charts_enabled) ===
    html += '<div class="grid-quad quad-seasonality chart-quadrant">\n'
    html += '    <div class="chart-controls">\n'
    html += '        <div class="control-group">\n'
    html += '            <label>Sources</label>\n'
    html += '            <div class="multi-select" id="msSeasonality"></div>\n'
    html += '        </div>\n'
    html += '        <div class="control-group">\n'
    html += '            <label>Lookback</label>\n'
    html += '            <select id="seasonalityLookback" onchange="updateSeasonalityChart()"></select>\n'
    html += '        </div>\n'
    html += '    </div>\n'
    html += '    <div class="chart-title" id="seasonalityTitle">Seasonality</div>\n'
    html += '    <div class="chart-canvas-wrap">\n'
    html += '        <canvas id="seasonalityCanvas"></canvas>\n'
    html += '    </div>\n'
    html += '    <div class="custom-legend" id="seasonalityLegend"></div>\n'
    html += '</div>\n\n'

    # === Bottom-right: generation buildup chart ===
    html += '<div class="grid-quad quad-buildup chart-quadrant">\n'
    html += '    <div class="chart-controls">\n'
    html += '        <div class="control-group">\n'
    html += '            <label>Sources</label>\n'
    html += '            <div class="multi-select" id="msBuildup"></div>\n'
    html += '        </div>\n'
    html += '        <div class="control-group">\n'
    html += '            <label>Period</label>\n'
    html += '            <select id="buildupAgg" onchange="onBuildupAggChange()">\n'
    html += '                <option value="Monthly">Monthly</option>\n'
    html += '                <option value="Annual CY">Annual (Calendar Year)</option>\n'
    html += '                <option value="Gas Year">Annual (Gas Year)</option>\n'
    html += '            </select>\n'
    html += '        </div>\n'
    html += '        <div class="control-group">\n'
    html += '            <label>From</label>\n'
    html += '            <select id="buildupFrom" onchange="updateBuildupChart()"></select>\n'
    html += '        </div>\n'
    html += '        <div class="control-group">\n'
    html += '            <label>To</label>\n'
    html += '            <select id="buildupTo" onchange="updateBuildupChart()"></select>\n'
    html += '        </div>\n'
    html += '    </div>\n'
    html += '    <div class="chart-title" id="buildupTitle">Generation</div>\n'
    html += '    <div class="chart-canvas-wrap">\n'
    html += '        <canvas id="buildupCanvas"></canvas>\n'
    html += '    </div>\n'
    html += '    <div class="custom-legend" id="buildupLegend"></div>\n'
    html += '</div>\n\n'

    html += '</div>\n\n'  # end main-grid

    html += '<div class="footer">\n'
    html += '    <span>Source: Palissy Advisors</span>\n'
    html += '    &nbsp;&bull;&nbsp;\n'
    html += '    <span>Last updated: ' + generated + '</span>\n'
    html += '</div>\n\n'

    html += '<script>\n' + js + '\n</script>\n\n'
    html += '</body>\n</html>'
    return html


def main():
    print("=" * 60)
    print("Palissy Multi-Dataset Dashboard Generator")
    print("=" * 60)

    datasets_by_key = {}
    ordered_keys = []
    for config in DATASETS:
        print(f"\n--- Building dataset: {config['key']} ({config['title']}) ---")
        blob = build_dataset_blob(config)
        datasets_by_key[config["key"]] = blob
        ordered_keys.append(config["key"])

    print("\nLoading assets...")
    assets = load_assets()

    print("\nGenerating HTML...")
    html = generate_html(datasets_by_key, ordered_keys, assets)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\nDashboard saved to: {OUTPUT_FILE}")
    print(f"  File size: {os.path.getsize(OUTPUT_FILE)/1024:.0f} KB")
    print(f"  Datasets: {', '.join(ordered_keys)}")
    print(f"  Display range: {DISPLAY_START_YEAR} - {DISPLAY_END_YEAR}")
    print("=" * 60)
    print("Done! Open output/index.html in a browser.")


if __name__ == "__main__":
    main()
