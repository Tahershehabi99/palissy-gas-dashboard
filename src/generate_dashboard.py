"""
Palissy European Gas Balance Dashboard Generator

Reads monthly bcf data from INPUT/gas_model_input.xlsx and generates
a self-contained HTML dashboard with:
- Expandable/collapsible rows
- Unit conversion (bcf, bcf/d, bcm, mmcm/d, TWh, mmt)
- Time period aggregation (Monthly, Quarterly, Annual CY, Annual GY, Summer, Winter)
- Palissy brand styling
- Admin-configurable display range
"""

import openpyxl
import json
import os
import base64
from datetime import datetime, date
from calendar import monthrange

# ============================================================
# ADMIN CONFIGURATION - Edit these to change display range
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

# Conversion factors (from bcf)
CONVERSIONS = {
    "bcf": 1.0,
    "bcm": 1.0 / 35.3,       # 1 bcm = 35.3 bcf
    "TWh": 1.0 / 3.41,        # 1 TWh = 3.41 bcf
    "mmt": 1.0 / 48.0,        # 1 mmt = 48 bcf
}

# Rate conversion labels
RATE_UNITS = {
    "bcf": "bcf/d",
    "bcm": "mmcm/d",
    "TWh": "GWh/d",
    "mmt": "kt/d",
}

# Rate conversion multipliers (from bcf/d)
RATE_CONVERSIONS = {
    "bcf": 1.0,                       # bcf/d
    "bcm": 1000.0 / 35.3,             # mmcm/d
    "TWh": 1000.0 / 3.41,             # GWh/d
    "mmt": 1000.0 / 48.0,             # kt/d
}

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


def read_input_data():
    """Read the INPUT Excel file and return structured data."""
    print("Reading input data...")
    wb = openpyxl.load_workbook(INPUT_FILE, read_only=True, data_only=True)
    ws = wb['Monthly Data']

    # Find last column
    last_col = 1
    for row in ws.iter_rows(min_row=1, max_row=1, max_col=500, values_only=False):
        for cell in row:
            if cell.value is not None:
                last_col = cell.column

    # Row 1: Dates
    dates = []
    for row in ws.iter_rows(min_row=1, max_row=1, min_col=2, max_col=last_col, values_only=True):
        for val in row:
            if val is not None:
                if isinstance(val, datetime):
                    dates.append(val)
                elif isinstance(val, str):
                    dates.append(datetime.strptime(val, "%Y-%m-%d"))
            else:
                dates.append(None)

    # Row 2: Days per month
    days_per_month = []
    for row in ws.iter_rows(min_row=2, max_row=2, min_col=2, max_col=last_col, values_only=True):
        for val in row:
            days_per_month.append(int(val) if val is not None else 30)

    # Rows 5+: Data rows (after header row 4 "bcf")
    rows = []
    row_idx = 0
    for row in ws.iter_rows(min_row=5, max_row=100, min_col=1, max_col=last_col, values_only=True):
        label = row[0]
        if label is None:
            break
        values = []
        for val in row[1:]:
            values.append(float(val) if val is not None else 0.0)
        rows.append({
            "label": str(label).strip(),
            "values": values
        })
        row_idx += 1

    wb.close()

    n_months = len(dates)
    print(f"  Read {len(rows)} data rows x {n_months} months")
    print(f"  Date range: {dates[0].strftime('%b %Y')} to {dates[-1].strftime('%b %Y')}")

    return dates, days_per_month, rows


def detect_hierarchy(rows):
    """
    Detect parent-child hierarchy from row labels.

    In the gas model, children appear BEFORE their parent total:
        Russia              <- child
        Norway              <- child
        + Imports           <- parent/total (children are above)

    Rules:
    - Rows with '+' or '-' prefix = parent/total/standalone
    - Un-prefixed rows that appear between two parent rows = children of
      the NEXT parent row (the total row that follows them)
    - Opening Storage, Closing Storage, Storage percentage = standalone
    """
    standalone = {'Opening Storage', 'Closing Storage', 'Storage percentage'}

    # First pass: classify each row
    classified = []
    for i, row in enumerate(rows):
        label = row["label"]
        is_parent = label.startswith('+') or label.startswith('-')
        is_standalone = label in standalone
        classified.append({
            "index": i,
            "label": label,
            "is_parent": is_parent,
            "is_standalone": is_standalone
        })

    # Second pass: group children with the parent total that follows them
    hierarchy = []
    pending_children = []

    for item in classified:
        if item["is_standalone"]:
            # Flush any pending children as standalone rows first
            for child in pending_children:
                hierarchy.append({
                    "label": child["label"],
                    "row_index": child["index"],
                    "children": [],
                    "type": "standalone"
                })
            pending_children = []
            hierarchy.append({
                "label": item["label"],
                "row_index": item["index"],
                "children": [],
                "type": "standalone"
            })
        elif item["is_parent"]:
            if pending_children:
                # These children belong to THIS parent total
                children = [{"label": c["label"], "row_index": c["index"]}
                           for c in pending_children]
                hierarchy.append({
                    "label": item["label"],
                    "row_index": item["index"],
                    "children": children,
                    "type": "parent"
                })
                pending_children = []
            else:
                hierarchy.append({
                    "label": item["label"],
                    "row_index": item["index"],
                    "children": [],
                    "type": "standalone"
                })
        else:
            # Un-prefixed row — collect as pending child
            pending_children.append(item)

    # Flush any remaining pending children
    for child in pending_children:
        hierarchy.append({
            "label": child["label"],
            "row_index": child["index"],
            "children": [],
            "type": "standalone"
        })

    return hierarchy


def is_stock_row(label):
    """Check if a row is a stock (level) vs flow row."""
    stock_labels = {'Opening Storage', 'Closing Storage', 'Storage percentage'}
    return label in stock_labels


def is_percentage_row(label):
    """Check if a row shows percentages (not converted to units)."""
    return label == 'Storage percentage'


def aggregate_monthly_to_periods(dates, days_per_month, rows):
    """
    Compute all time period aggregations from monthly data.
    Returns dict of {period_name: {columns, data}} where data
    contains aggregated values for each row.
    """
    n_months = len(dates)
    results = {}

    # Helper to get year/month from date index
    def ym(idx):
        return dates[idx].year, dates[idx].month

    # ========== MONTHLY ==========
    monthly_cols = []
    for i, d in enumerate(dates):
        monthly_cols.append({
            "label": d.strftime("%b %Y"),  # "Jan 2020"
            "short": d.strftime("%b-%y"),   # "Jan-20"
            "year": d.year,
            "month": d.month,
            "indices": [i],
            "days": days_per_month[i]
        })
    results["Monthly"] = monthly_cols

    # ========== QUARTERLY ==========
    quarterly_cols = []
    # Group months by year and quarter
    quarters = {}
    for i, d in enumerate(dates):
        y = d.year
        q = (d.month - 1) // 3 + 1
        key = (y, q)
        if key not in quarters:
            quarters[key] = {"indices": [], "days": 0}
        quarters[key]["indices"].append(i)
        quarters[key]["days"] += days_per_month[i]

    for (y, q), info in sorted(quarters.items()):
        if len(info["indices"]) == 3:  # Complete quarter only
            quarterly_cols.append({
                "label": f"Q{q} {y}",
                "short": f"Q{q}-{str(y)[2:]}",
                "year": y,
                "quarter": q,
                "indices": info["indices"],
                "days": info["days"]
            })
    results["Quarterly"] = quarterly_cols

    # ========== ANNUAL CALENDAR YEAR ==========
    annual_cy_cols = []
    years = {}
    for i, d in enumerate(dates):
        y = d.year
        if y not in years:
            years[y] = {"indices": [], "days": 0}
        years[y]["indices"].append(i)
        years[y]["days"] += days_per_month[i]

    for y, info in sorted(years.items()):
        if len(info["indices"]) == 12:  # Complete year only
            annual_cy_cols.append({
                "label": str(y),
                "short": str(y),
                "year": y,
                "indices": info["indices"],
                "days": info["days"]
            })
    results["Annual CY"] = annual_cy_cols

    # ========== ANNUAL GAS YEAR ==========
    # Gas Year N = Oct Year N through Sep Year N+1
    annual_gy_cols = []
    gas_years = {}
    for i, d in enumerate(dates):
        gy = d.year if d.month >= 10 else d.year - 1
        if gy not in gas_years:
            gas_years[gy] = {"indices": [], "days": 0}
        gas_years[gy]["indices"].append(i)
        gas_years[gy]["days"] += days_per_month[i]

    for gy, info in sorted(gas_years.items()):
        if len(info["indices"]) == 12:  # Complete gas year only
            label = f"{str(gy)[2:]}/{str(gy+1)[2:]}"
            annual_gy_cols.append({
                "label": f"GY {label}",
                "short": label,
                "year": gy,
                "indices": info["indices"],
                "days": info["days"]
            })
    results["Gas Year"] = annual_gy_cols

    # ========== WINTERS ==========
    # Winter of GY N = Oct Year N through Mar Year N+1
    winter_cols = []
    winters = {}
    for i, d in enumerate(dates):
        if d.month >= 10:
            gy = d.year
        elif d.month <= 3:
            gy = d.year - 1
        else:
            continue
        if gy not in winters:
            winters[gy] = {"indices": [], "days": 0}
        winters[gy]["indices"].append(i)
        winters[gy]["days"] += days_per_month[i]

    for gy, info in sorted(winters.items()):
        if len(info["indices"]) == 6:  # Complete winter only
            label = f"Win {str(gy)[2:]}/{str(gy+1)[2:]}"
            winter_cols.append({
                "label": label,
                "short": label,
                "year": gy,
                "indices": info["indices"],
                "days": info["days"]
            })
    results["Winter"] = winter_cols

    # ========== SUMMERS ==========
    # Summer = Apr Year N through Sep Year N
    summer_cols = []
    summers = {}
    for i, d in enumerate(dates):
        if 4 <= d.month <= 9:
            y = d.year
            if y not in summers:
                summers[y] = {"indices": [], "days": 0}
            summers[y]["indices"].append(i)
            summers[y]["days"] += days_per_month[i]

    for y, info in sorted(summers.items()):
        if len(info["indices"]) == 6:  # Complete summer only
            label = f"Sum {y}"
            summer_cols.append({
                "label": label,
                "short": label,
                "year": y,
                "indices": info["indices"],
                "days": info["days"]
            })
    results["Summer"] = summer_cols

    return results


def compute_period_values(rows, period_cols):
    """
    For each time period column, compute the aggregated bcf value for each row.
    - Stock rows (Opening/Closing Storage): take first/last month value
    - Storage percentage: take last month value
    - Flow rows: sum of monthly values
    """
    aggregated = []
    for row in rows:
        label = row["label"]
        values = row["values"]
        period_values = []

        for col in period_cols:
            indices = col["indices"]
            if not indices:
                period_values.append(0)
                continue

            if label == 'Opening Storage':
                # First month's value
                period_values.append(values[indices[0]])
            elif label == 'Closing Storage':
                # Last month's value
                period_values.append(values[indices[-1]])
            elif label == 'Storage percentage':
                # Last month's value
                period_values.append(values[indices[-1]])
            else:
                # Sum of months
                total = sum(values[idx] for idx in indices)
                period_values.append(total)

        aggregated.append({
            "label": label,
            "bcf_values": period_values
        })

    return aggregated


def build_dashboard_data(dates, days_per_month, rows, hierarchy, period_results):
    """Build the complete data structure for the dashboard JSON."""

    # Build period data for each view
    views = {}
    for view_name, period_cols in period_results.items():
        # Filter to display range
        filtered_cols = []
        for col in period_cols:
            year = col.get("year", 0)
            if view_name == "Monthly":
                if DISPLAY_START_YEAR <= year <= DISPLAY_END_YEAR:
                    filtered_cols.append(col)
            elif view_name == "Quarterly":
                if DISPLAY_START_YEAR <= year <= DISPLAY_END_YEAR:
                    filtered_cols.append(col)
            elif view_name in ("Annual CY",):
                if DISPLAY_START_YEAR <= year <= DISPLAY_END_YEAR:
                    filtered_cols.append(col)
            elif view_name in ("Gas Year",):
                # GY year is the start year (Oct). Show if GY starts in range
                # GY 19/20 starts Oct 2019, ends Sep 2020
                # Show GY if its start year >= DISPLAY_START_YEAR - 1 and <= DISPLAY_END_YEAR - 1
                if (DISPLAY_START_YEAR - 1) <= year <= (DISPLAY_END_YEAR - 1):
                    filtered_cols.append(col)
            elif view_name == "Winter":
                if (DISPLAY_START_YEAR - 1) <= year <= (DISPLAY_END_YEAR - 1):
                    filtered_cols.append(col)
            elif view_name == "Summer":
                if DISPLAY_START_YEAR <= year <= DISPLAY_END_YEAR:
                    filtered_cols.append(col)

        # Compute aggregated values
        aggregated = compute_period_values(rows, filtered_cols)

        # Build days array for rate conversion
        days_array = [col["days"] for col in filtered_cols]

        views[view_name] = {
            "columns": [col["label"] for col in filtered_cols],
            "short_columns": [col["short"] for col in filtered_cols],
            "days": days_array,
            "rows": []
        }

        for agg_row in aggregated:
            views[view_name]["rows"].append({
                "label": agg_row["label"],
                "bcf": agg_row["bcf_values"]
            })

    # Build hierarchy for the UI
    ui_hierarchy = []
    for item in hierarchy:
        entry = {
            "label": item["label"],
            "index": item["row_index"],
            "type": item["type"],
            "is_stock": is_stock_row(item["label"]),
            "is_pct": is_percentage_row(item["label"]),
            "children": []
        }
        for child in item["children"]:
            entry["children"].append({
                "label": child["label"],
                "index": child["row_index"],
                "is_stock": False,
                "is_pct": False
            })
        ui_hierarchy.append(entry)

    return {
        "views": views,
        "hierarchy": ui_hierarchy,
        "conversions": CONVERSIONS,
        "rate_conversions": RATE_CONVERSIONS,
        "rate_units": RATE_UNITS,
        "display_start": DISPLAY_START_YEAR,
        "display_end": DISPLAY_END_YEAR,
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }


def load_assets():
    """Load logo and font as base64 for embedding."""
    assets = {}

    if os.path.exists(LOGO_FILE):
        with open(LOGO_FILE, "rb") as f:
            assets["logo_b64"] = base64.b64encode(f.read()).decode("utf-8")
        print(f"  Logo loaded: {LOGO_FILE}")
    else:
        assets["logo_b64"] = ""
        print(f"  WARNING: Logo not found at {LOGO_FILE}")

    if os.path.exists(FONT_FILE):
        with open(FONT_FILE, "rb") as f:
            assets["font_b64"] = base64.b64encode(f.read()).decode("utf-8")
        print(f"  Font loaded: {FONT_FILE}")
    else:
        assets["font_b64"] = ""
        print(f"  WARNING: Font not found at {FONT_FILE}")

    return assets


def generate_html(dashboard_data, assets):
    """Generate the self-contained HTML dashboard."""

    data_json = json.dumps(dashboard_data, separators=(',', ':'))
    logo_b64 = assets.get("logo_b64", "")
    font_b64 = assets.get("font_b64", "")
    generated = dashboard_data["generated"]
    db = COLORS["dark_blue"]
    grey = COLORS["grey"]
    red = COLORS["red"]
    card = COLORS["card_bg"]
    border = COLORS["border"]
    grid = COLORS["grid"]

    # Build HTML using string concatenation to avoid f-string issues with JS
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
    border-bottom: 2px solid """ + db + """; margin-bottom: 15px;
}
.header img { height: 70px; object-fit: contain; margin-bottom: 8px; }
.header h1 {
    font-size: 18px; font-weight: bold; color: """ + db + """;
    letter-spacing: 1px; text-transform: uppercase;
}

.controls {
    display: flex; justify-content: center; align-items: center; gap: 20px;
    padding: 14px 24px; background: """ + card + """;
    border: 1px solid """ + border + """; border-radius: 12px;
    margin: 0 20px 15px; flex-wrap: wrap;
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
.control-group select:hover {
    border-color: """ + db + """;
}
.control-group select:focus {
    outline: none; border-color: """ + db + """;
    box-shadow: 0 0 0 3px rgba(39, 41, 98, 0.12);
}

.table-container {
    margin: 0 20px 20px; overflow-x: auto;
    border: 1px solid """ + border + """; border-radius: 10px;
    max-height: calc(100vh - 230px); overflow-y: auto;
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
.table-container::-webkit-scrollbar { height: 8px; width: 8px; }
.table-container::-webkit-scrollbar-track { background: #f0f0f4; border-radius: 4px; }
.table-container::-webkit-scrollbar-thumb { background: """ + grey + """; border-radius: 4px; }
.footer { text-align: center; padding: 12px; font-size: 10px; color: """ + grey + """; }
.unit-label { font-size: 10px; color: rgba(255,255,255,0.6); font-style: italic; margin-left: 6px; }
@media (max-width: 1024px) {
    .header { padding: 16px 16px 8px; }
    .header img { height: 55px; }
    .header h1 { font-size: 16px; }
    .controls { margin: 0 12px 12px; padding: 12px 18px; gap: 14px; }
    .table-container { margin: 0 12px 15px; max-height: calc(100vh - 210px); }
    thead th { font-size: 10.5px; padding: 8px 10px; }
    tbody td { font-size: 11.5px; padding: 5px 10px; }
    thead th:first-child, tbody td:first-child { min-width: 200px; }
    tr.child-row td:first-child { padding-left: 26px; }
}
@media (max-width: 768px) {
    .header { padding: 14px 12px 8px; margin-bottom: 10px; }
    .header img { height: 45px; margin-bottom: 6px; }
    .header h1 { font-size: 14px; letter-spacing: 0.5px; }
    .controls {
        flex-direction: column; gap: 10px; padding: 10px 14px;
        margin: 0 8px 10px; border-radius: 10px;
    }
    .control-group { width: 100%; justify-content: space-between; }
    .control-group select { flex: 1; font-size: 13px; padding: 10px 32px 10px 12px; }
    .control-group label { font-size: 10px; min-width: 70px; }
    .table-container {
        margin: 0 8px 12px; border-radius: 8px;
        max-height: calc(100vh - 200px);
        -webkit-overflow-scrolling: touch;
    }
    table { font-size: 11px; }
    thead th { font-size: 10px; padding: 8px 8px; }
    tbody td { font-size: 11px; padding: 5px 8px; }
    thead th:first-child, tbody td:first-child { min-width: 150px; font-size: 10.5px; }
    tr.child-row td:first-child { padding-left: 22px; }
    .footer { font-size: 9px; padding: 10px 8px; }
    .unit-label { display: none; }
}
@media (max-width: 480px) {
    .header { padding: 10px 8px 6px; margin-bottom: 8px; }
    .header img { height: 36px; margin-bottom: 4px; }
    .header h1 { font-size: 12px; }
    .controls { margin: 0 6px 8px; padding: 8px 10px; gap: 8px; border-radius: 8px; }
    .control-group select { font-size: 12px; padding: 8px 28px 8px 10px; }
    .table-container { margin: 0 6px 10px; border-radius: 6px; }
    thead th { font-size: 9px; padding: 6px 6px; }
    tbody td { font-size: 10px; padding: 4px 6px; }
    thead th:first-child, tbody td:first-child { min-width: 120px; font-size: 9.5px; }
    tr.child-row td:first-child { padding-left: 18px; }
    .toggle-arrow { width: 10px; font-size: 8px; }
}
"""

    # JavaScript as a plain string (no f-string escaping issues)
    js = r"""
const DATA = __DATA_PLACEHOLDER__;

let expandedGroups = new Set();
let highlightedCol = -1;
let highlightedRow = -1;

const UNIT_CONFIG = {
    'bcf':    { volLabel: 'bcf',   rateLabel: 'bcf/d',  volFactor: 1.0,           rateFactor: 1.0 },
    'bcf/d':  { volLabel: 'bcf',   rateLabel: 'bcf/d',  volFactor: 1.0,           rateFactor: 1.0,          isRate: true },
    'bcm':    { volLabel: 'bcm',   rateLabel: 'mmcm/d', volFactor: 1.0/35.3,      rateFactor: 1000.0/35.3 },
    'mmcm/d': { volLabel: 'bcm',   rateLabel: 'mmcm/d', volFactor: 1.0/35.3,      rateFactor: 1000.0/35.3,  isRate: true },
    'TWh':    { volLabel: 'TWh',   rateLabel: 'GWh/d',  volFactor: 1.0/3.41,      rateFactor: 1000.0/3.41 },
    'GWh/d':  { volLabel: 'TWh',   rateLabel: 'GWh/d',  volFactor: 1.0/3.41,      rateFactor: 1000.0/3.41,  isRate: true },
    'mmt':    { volLabel: 'mmt',   rateLabel: 'kt/d',   volFactor: 1.0/48.0,      rateFactor: 1000.0/48.0 },
    'kt/d':   { volLabel: 'mmt',   rateLabel: 'kt/d',   volFactor: 1.0/48.0,      rateFactor: 1000.0/48.0,  isRate: true }
};

function formatNumber(val, isPct) {
    if (isPct) return (val * 100).toFixed(1) + '%';
    var abs = Math.abs(val);
    if (abs < 0.005) return '0';
    if (abs >= 100) {
        return Math.round(val).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    }
    if (abs >= 10) return val.toFixed(1);
    return val.toFixed(2);
}

function computeDisplayValue(bcfVal, unitKey, isStock, isPct, days) {
    if (isPct) return bcfVal;
    var cfg = UNIT_CONFIG[unitKey];
    if (isStock) {
        return bcfVal * cfg.volFactor;
    }
    if (cfg.isRate) {
        return (bcfVal / days) * cfg.rateFactor;
    }
    return bcfVal * cfg.volFactor;
}

function getStockLabel(label, unitKey) {
    var cfg = UNIT_CONFIG[unitKey];
    if (cfg.isRate) {
        return label + ' (' + cfg.volLabel + ')';
    }
    return label;
}

function getHeaderUnitLabel(unitKey) {
    var cfg = UNIT_CONFIG[unitKey];
    if (cfg.isRate) return cfg.rateLabel;
    return cfg.volLabel;
}

function updateTable() {
    var period = document.getElementById('periodSelector').value;
    var unitKey = document.getElementById('unitSelector').value;
    var view = DATA.views[period];
    if (!view) return;

    var columns = view.columns;
    var shortCols = view.short_columns;
    var days = view.days;
    var rows = view.rows;
    var hierarchy = DATA.hierarchy;
    var cfg = UNIT_CONFIG[unitKey];
    var unitLabel = getHeaderUnitLabel(unitKey);

    // Build header
    var thead = document.getElementById('tableHead');
    var hHtml = '<tr><th>' + unitLabel + '<span class="unit-label">(' + period + ')</span></th>';
    for (var i = 0; i < columns.length; i++) {
        hHtml += '<th>' + (shortCols ? shortCols[i] : columns[i]) + '</th>';
    }
    hHtml += '</tr>';
    thead.innerHTML = hHtml;

    // Build body
    var tbody = document.getElementById('tableBody');
    var bHtml = '';

    for (var h = 0; h < hierarchy.length; h++) {
        var item = hierarchy[h];
        var rowData = rows[item.index];
        var isExpanded = expandedGroups.has(item.label);
        var hasChildren = item.children && item.children.length > 0;
        var isPct = item.is_pct;
        var isStock = item.is_stock;

        var rowClass = item.type === 'standalone' ? 'standalone-row' : 'parent-row';
        if (isPct) rowClass += ' pct-row';

        bHtml += '<tr class="' + rowClass + '">';

        // Label cell - only this cell is clickable for expand/collapse
        var lbl = '';
        if (hasChildren) {
            lbl += '<span class="toggle-arrow' + (isExpanded ? ' expanded' : '') + '">&#9654;</span> ';
        }
        var displayLabel = isStock ? getStockLabel(item.label, unitKey) : item.label;
        lbl += displayLabel.replace(/&/g,'&amp;').replace(/</g,'&lt;');
        if (hasChildren) {
            bHtml += '<td data-toggle="' + h + '">' + lbl + '</td>';
        } else {
            bHtml += '<td>' + lbl + '</td>';
        }

        // Values
        for (var i = 0; i < rowData.bcf.length; i++) {
            var bcfVal = rowData.bcf[i];
            var dv;
            if (isPct) {
                dv = formatNumber(bcfVal, true);
            } else {
                dv = formatNumber(computeDisplayValue(bcfVal, unitKey, isStock, false, days[i]), false);
            }
            bHtml += '<td>' + dv + '</td>';
        }
        bHtml += '</tr>';

        // Children
        if (hasChildren) {
            for (var c = 0; c < item.children.length; c++) {
                var child = item.children[c];
                var childData = rows[child.index];
                var hidden = !isExpanded ? ' hidden' : '';
                bHtml += '<tr class="child-row' + hidden + '">';
                bHtml += '<td>' + child.label.replace(/&/g,'&amp;').replace(/</g,'&lt;') + '</td>';
                for (var i = 0; i < childData.bcf.length; i++) {
                    var bcfVal = childData.bcf[i];
                    var dv = formatNumber(computeDisplayValue(bcfVal, unitKey, false, false, days[i]), false);
                    bHtml += '<td>' + dv + '</td>';
                }
                bHtml += '</tr>';
            }
        }
    }
    tbody.innerHTML = bHtml;
}

// Toggle expand/collapse only when clicking on the first column (label cell)
document.addEventListener('click', function(e) {
    var td = e.target.closest('td[data-toggle]');
    if (!td) return;
    var idx = parseInt(td.getAttribute('data-toggle'));
    var label = DATA.hierarchy[idx].label;
    if (expandedGroups.has(label)) {
        expandedGroups.delete(label);
    } else {
        expandedGroups.add(label);
    }
    updateTable();
});

// Crosshair highlight on hover
document.addEventListener('mouseover', function(e) {
    var td = e.target.closest('td');
    if (!td) return;
    var tr = td.closest('tr');
    var tbody = document.getElementById('tableBody');
    if (!tr || !tbody.contains(tr)) return;

    var colIdx = Array.from(tr.children).indexOf(td);
    var rowIdx = Array.from(tbody.children).indexOf(tr);
    if (colIdx === highlightedCol && rowIdx === highlightedRow) return;
    clearHighlight();
    highlightedCol = colIdx;
    highlightedRow = rowIdx;

    // Highlight all cells in same column
    var allRows = tbody.querySelectorAll('tr');
    for (var r = 0; r < allRows.length; r++) {
        var cells = allRows[r].children;
        if (colIdx < cells.length) {
            if (r === rowIdx) {
                cells[colIdx].classList.add('cell-highlight');
            } else {
                cells[colIdx].classList.add('col-highlight');
            }
        }
    }
    // Highlight all cells in same row
    var rowCells = tr.children;
    for (var c = 0; c < rowCells.length; c++) {
        if (c !== colIdx) {
            rowCells[c].classList.add('row-highlight');
        }
    }
});

document.addEventListener('mouseout', function(e) {
    var td = e.target.closest('td');
    if (td) {
        var related = e.relatedTarget;
        if (related && related.closest && related.closest('td')) return;
    }
    clearHighlight();
});

function clearHighlight() {
    var els = document.querySelectorAll('.col-highlight, .row-highlight, .cell-highlight');
    for (var i = 0; i < els.length; i++) {
        els[i].classList.remove('col-highlight', 'row-highlight', 'cell-highlight');
    }
    highlightedCol = -1;
    highlightedRow = -1;
}

document.addEventListener('DOMContentLoaded', function() {
    updateTable();
});
"""

    # Replace the data placeholder
    js = js.replace('__DATA_PLACEHOLDER__', data_json)

    html = '<!DOCTYPE html>\n<html lang="en">\n<head>\n'
    html += '<meta charset="UTF-8">\n'
    html += '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
    html += '<title>Palissy Advisors - European Gas Balance</title>\n'
    html += '<style>\n' + css + '\n</style>\n'
    html += '</head>\n<body>\n\n'

    html += '<div class="header">\n'
    html += '    <img src="data:image/png;base64,' + logo_b64 + '" alt="Palissy Advisors">\n'
    html += '    <h1>European Gas Balance</h1>\n'
    html += '</div>\n\n'

    html += '<div class="controls">\n'
    html += '    <div class="control-group">\n'
    html += '        <label>Time Period</label>\n'
    html += '        <select id="periodSelector" onchange="updateTable()">\n'
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
    html += '        <select id="unitSelector" onchange="updateTable()">\n'
    html += '            <option value="bcf" selected>bcf</option>\n'
    html += '            <option value="bcf/d">bcf/d</option>\n'
    html += '            <option value="bcm">bcm</option>\n'
    html += '            <option value="mmcm/d">mmcm/d</option>\n'
    html += '            <option value="TWh">TWh</option>\n'
    html += '            <option value="GWh/d">GWh/d</option>\n'
    html += '            <option value="mmt">mmt</option>\n'
    html += '            <option value="kt/d">kt/d</option>\n'
    html += '        </select>\n'
    html += '    </div>\n'
    html += '</div>\n\n'

    html += '<div class="table-container" id="tableContainer">\n'
    html += '    <table id="dataTable">\n'
    html += '        <thead id="tableHead"></thead>\n'
    html += '        <tbody id="tableBody"></tbody>\n'
    html += '    </table>\n'
    html += '</div>\n\n'

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
    print("Palissy European Gas Balance Dashboard Generator")
    print("=" * 60)

    # Read input
    dates, days_per_month, rows = read_input_data()

    # Detect hierarchy
    hierarchy = detect_hierarchy(rows)
    print(f"\nHierarchy detected:")
    for item in hierarchy:
        children_str = f" -> {len(item['children'])} children" if item['children'] else ""
        print(f"  {item['label']} ({item['type']}){children_str}")

    # Compute all period aggregations
    print("\nComputing aggregations...")
    period_results = aggregate_monthly_to_periods(dates, days_per_month, rows)
    for name, cols in period_results.items():
        print(f"  {name}: {len(cols)} periods")

    # Build dashboard data
    print("\nBuilding dashboard data...")
    dashboard_data = build_dashboard_data(dates, days_per_month, rows, hierarchy, period_results)

    # Load assets
    print("\nLoading assets...")
    assets = load_assets()

    # Generate HTML
    print("\nGenerating HTML dashboard...")
    html = generate_html(dashboard_data, assets)

    # Write output
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    file_size = os.path.getsize(OUTPUT_FILE)
    print(f"\nDashboard saved to: {OUTPUT_FILE}")
    print(f"  File size: {file_size / 1024:.0f} KB")
    print(f"  Display range: {DISPLAY_START_YEAR} - {DISPLAY_END_YEAR}")
    print("=" * 60)
    print("Done! Open output/index.html in a browser to preview.")


if __name__ == "__main__":
    main()
