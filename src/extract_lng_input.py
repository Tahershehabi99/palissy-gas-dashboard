"""
Extract LNG monthly data from the master model into a clean INPUT workbook.

Source : Context/AKAP Global LNG Model.xlsx  (READ-ONLY — never modified)
Output : INPUT/lng_model_input.xlsx

Copied faithfully (values only), preserving original row/column positions so the
master's row numbers still line up:
  - 'Monthly Imports' : source rows 4-21, cols A..col 290 (Jan 2017 - Dec 2040)
        row 4  = month date header
        rows 7-21 = EU+UK (potential), EU+UK, China, Japan, South Korea, Taiwan,
                    India, Other Asia, LatAm, Middle East, Egypt, Turkey, RoW,
                    Unaccounted demand, Total
  - 'Monthly Exports' : source rows 4-44, cols A..col 293 (Jan 2017 - Dec 2040)
        row 4  = month date header (data begins at col F; B-E are 'Working' cols)
        rows 8-44 = Production (mmt) block by country, with regional subtotals
                    (Asia, LatAm, MENA and Europe, North America,
                     Sub-Saharan Africa) and Grand total (row 44)

NOTE: This is a one-shot hardcode of current data. When the user finalises the
paste/refresh methodology, this is the script to adjust.
"""
import openpyxl
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "Context" / "AKAP Global LNG Model.xlsx"
OUT = ROOT / "INPUT" / "lng_model_input.xlsx"

# (source sheet, first row, last row, first col, last col)
BLOCKS = [
    ("Monthly Imports", 4, 21, 1, 290),
    ("Monthly Exports", 4, 44, 1, 293),
]


def main():
    # NOTE: load WITHOUT read_only. Random .cell() access on a read_only sheet is
    # pathologically slow (re-scans the stream each call), which hung earlier runs.
    src = openpyxl.load_workbook(SRC, read_only=False, data_only=True)
    out = openpyxl.Workbook()
    out.remove(out.active)

    for sheet, r0, r1, c0, c1 in BLOCKS:
        ws_in = src[sheet]
        ws_out = out.create_sheet(sheet)
        copied = 0
        for r in range(r0, r1 + 1):
            for c in range(c0, c1 + 1):
                v = ws_in.cell(row=r, column=c).value
                if v is not None:
                    ws_out.cell(row=r, column=c, value=v)
                    copied += 1
        print(f"{sheet}: rows {r0}-{r1}, cols {c0}-{c1} -> {copied} non-empty cells")

    build_projects(src, out)

    OUT.parent.mkdir(exist_ok=True)
    out.save(OUT)
    print(f"Saved {OUT}")


# ============================================================
# LNG PROJECTS (for the LNG Projects tab)
# ------------------------------------------------------------
# Build a clean projects dataset from the master and write three sheets:
#   'LNG Projects'        - one row per MAIN terminal (trains excluded), joined
#                           to the Assumptions tab metadata + computed Risked Cap.
#   'LNG Proj Production' - mains x months (Reported mmt profile, reported+forecast)
#   'LNG Proj Utilisation'- mains x months (Utilisation %)
# Trains (col C='T' in Monthly Exports 'Reported' section) are dropped entirely;
# only mains (col C='M') are kept. Country comes from that section's aggregate
# rows; region from REGION_OF; metadata is joined from Assumptions by name.
# ============================================================
REGION_OF = {
    "Brunei": "Asia", "Indonesia": "Asia", "Malaysia": "Asia",
    "Papua New Guinea": "Asia", "Timor Leste": "Asia",
    "Australia": "Australia",
    "Argentina": "LatAm", "Mexico": "LatAm", "Peru": "LatAm",
    "Suriname": "LatAm", "Trinidad": "LatAm", "Venezuela": "LatAm",
    "Algeria": "MENA and Europe", "Egypt": "MENA and Europe", "Israel": "MENA and Europe",
    "Mauritania": "MENA and Europe", "Norway": "MENA and Europe", "Oman": "MENA and Europe",
    "Qatar": "MENA and Europe", "UAE": "MENA and Europe",
    "Canada": "North America", "United States": "North America",
    "Russia": "Russia",
    "Angola": "Sub-Saharan Africa", "Cameroon": "Sub-Saharan Africa",
    "Congo (Rep.)": "Sub-Saharan Africa", "Equatorial Guinea": "Sub-Saharan Africa",
    "Mozambique": "Sub-Saharan Africa", "Nigeria": "Sub-Saharan Africa",
    "Tanzania": "Sub-Saharan Africa",
}
COUNTRIES = set(REGION_OF)

# Monthly Exports project-section layout (verified against the master)
EXP_DATE_ROW = 4
EXP_COL0, EXP_COL1 = 6, 293            # data cols F..(Dec 2040), 288 months
REP_R0, REP_R1 = 165, 396              # 'Reported (mmt)' project rows (165) .. before Grand total
UTIL_R0, UTIL_R1 = 874, 1110           # 'Utilisation (%)' project rows


def _clean(v):
    return v.strip() if isinstance(v, str) else v


def build_projects(src, out):
    me = src["Monthly Exports"]
    asm = src["Assumptions"]

    # --- Assumptions metadata, keyed by project name ---
    # cols: A name, B country, C status, D unrisked mmt, E unrisked bcf/d, F start,
    #       G CoS, H util forecast, I util decline, J decline start,
    #       W(23) operator, X(24) partners, Y(25) primary stake, Z-AF(26-32) S1-S7
    meta = {}
    for r in range(3, asm.max_row + 1):
        name = _clean(asm.cell(r, 1).value)
        if not name or name == "Total":
            continue
        start = asm.cell(r, 6).value
        meta[name] = {
            "status": _clean(asm.cell(r, 3).value) or "",
            "unrisked_mmt": asm.cell(r, 4).value,
            "unrisked_bcfd": asm.cell(r, 5).value,
            "start": start.strftime("%Y-%m-%d") if isinstance(start, datetime) else "",
            "cos": asm.cell(r, 7).value,
            "util_forecast": asm.cell(r, 8).value,
            "util_decline": asm.cell(r, 9).value,
            "operator": _clean(asm.cell(r, 23).value) or "",
            "partners": _clean(asm.cell(r, 24).value) or "",
            "primary_stake": asm.cell(r, 25).value,
            "stakes": [asm.cell(r, c).value for c in range(26, 33)],  # S1..S7
        }

    # --- Utilisation series, keyed by name ---
    util_by_name = {}
    for r in range(UTIL_R0, UTIL_R1 + 1):
        name = _clean(me.cell(r, 1).value)
        if not name or name in COUNTRIES or name in ("Grand total",):
            continue
        util_by_name[name] = [me.cell(r, c).value for c in range(EXP_COL0, EXP_COL1 + 1)]

    # --- Walk the Reported section: accumulate project rows, assign country on the
    #     country aggregate row; keep mains (M) only. ---
    dates = [me.cell(EXP_DATE_ROW, c).value for c in range(EXP_COL0, EXP_COL1 + 1)]
    projects = []           # ordered list of main dicts
    pending = []            # rows since last country marker
    for r in range(REP_R0, REP_R1 + 1):
        name = _clean(me.cell(r, 1).value)
        if not name:
            continue
        marker = _clean(me.cell(r, 3).value)
        if marker in ("M", "T"):
            pending.append((name, marker, [me.cell(r, c).value for c in range(EXP_COL0, EXP_COL1 + 1)]))
        elif name in COUNTRIES:
            country = name
            for pname, pmark, series in pending:
                if pmark != "M":
                    continue                      # drop trains
                m = meta.get(pname, {})
                cos = m.get("cos")
                unr = m.get("unrisked_mmt")
                risked = (unr * cos) if (isinstance(unr, (int, float)) and isinstance(cos, (int, float))) \
                    else (unr if isinstance(unr, (int, float)) else None)  # producing: blank CoS -> 1.0
                projects.append({
                    "name": pname, "country": country, "region": REGION_OF.get(country, country),
                    "prod": series, "util": util_by_name.get(pname), **m, "risked_mmt": risked,
                })
            pending = []
        # else: regional/grand-total rows -> ignore (and reset so they don't leak)
        elif name in ("Grand total",):
            pending = []

    print(f"LNG Projects: {len(projects)} mains across "
          f"{len(set(p['country'] for p in projects))} countries")
    missing = [p['name'] for p in projects if p['name'] not in meta]
    if missing:
        print(f"  (no Assumptions match for {len(missing)}: {missing})")

    # --- Write 'LNG Projects' metadata sheet ---
    ws = out.create_sheet("LNG Projects")
    headers = ["Project", "Country", "Region", "Status", "Unrisked (mmt)", "Unrisked (bcf/d)",
               "CoS", "Risked (mmt)", "Start date", "Util forecast", "Util decline",
               "Operator", "Partners", "Primary stake", "S1", "S2", "S3", "S4", "S5", "S6", "S7"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=1, column=c, value=h)
    for i, p in enumerate(projects, start=2):
        row = [p["name"], p["country"], p["region"], p.get("status", ""),
               p.get("unrisked_mmt"), p.get("unrisked_bcfd"), p.get("cos"), p.get("risked_mmt"),
               p.get("start", ""), p.get("util_forecast"), p.get("util_decline"),
               p.get("operator", ""), p.get("partners", ""), p.get("primary_stake")]
        row += (p.get("stakes") or [None] * 7)[:7]
        for c, v in enumerate(row, start=1):
            ws.cell(row=i, column=c, value=v)

    # --- Write series sheets (row 1 = dates from col B; col A = project name) ---
    for sheet_name, key in [("LNG Proj Production", "prod"), ("LNG Proj Utilisation", "util")]:
        sh = out.create_sheet(sheet_name)
        for j, d in enumerate(dates, start=2):
            sh.cell(row=1, column=j, value=d)
        for i, p in enumerate(projects, start=2):
            sh.cell(row=i, column=1, value=p["name"])
            series = p.get(key)
            if series:
                for j, v in enumerate(series, start=2):
                    if v is not None:
                        sh.cell(row=i, column=j, value=v)
    print("LNG Projects sheets written: 'LNG Projects', 'LNG Proj Production', 'LNG Proj Utilisation'")


if __name__ == "__main__":
    main()
