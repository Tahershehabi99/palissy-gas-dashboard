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

    OUT.parent.mkdir(exist_ok=True)
    out.save(OUT)
    print(f"Saved {OUT}")


if __name__ == "__main__":
    main()
