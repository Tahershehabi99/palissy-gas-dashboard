"""
Ask the user for the last ACTUAL month of the gas balance and power data and
store it in WORKING/gas_settings.json, which generate_dashboard.py reads to
draw the actual/forecast boundary on the Gas Balance and Power tabs.

The master carries the boundary as CELL COLOUR: black = reported/actual, blue
(or red-on-yellow) = forecast. We read that per row in fast read-only mode and
propose a boundary per tab; the user confirms it at the prompt (Enter = accept,
or type another month). Called by INPUT/update_gas.bat after the extract.

Per-row detection lives in extract_gas_input.row_boundaries (and is written to
the extract's 'Boundaries' sheet, so each table row is shaded from its OWN
month). What is set here is the TAB-level boundary: the header shading, the
"Actuals through ..." note, the range chart's current-line cap, and the fallback
for rows with no colour information (derived rows like '- Consumption').
Suggested = the month most coloured rows agree on; rows that differ are printed.

  py -3 src/set_gas_actual.py            # detect + interactive prompts
  py -3 src/set_gas_actual.py --show     # just print the current values
  py -3 src/set_gas_actual.py --detect   # print the per-row detection only
"""
import json
import re
import sys
from datetime import datetime
from pathlib import Path

from collections import Counter

import openpyxl

from extract_gas_input import SRC, LayoutError, row_boundaries

ROOT = Path(__file__).resolve().parent.parent
SETTINGS = ROOT / "WORKING" / "gas_settings.json"
# Must match the "actual_end" defaults in generate_dashboard.py DATASETS.
DEFAULTS = {"gas": "2026-05", "power": "2026-05"}
LABELS = {"gas": "Gas Balance", "power": "Power"}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def load():
    cur = dict(DEFAULTS)
    if SETTINGS.exists():
        try:
            cur.update({k: v for k, v in json.loads(SETTINGS.read_text(encoding="utf-8")).items()
                        if k in DEFAULTS})
        except Exception:
            pass
    return cur


def pretty(ym):
    y, m = ym.split("-")
    return f"{MONTHS[int(m) - 1]} {y}"


def parse(text, fallback):
    """Accept '2026-08', '2026/08', '08-2026', '08/2026', 'Aug 2026', 'aug26', '202608'."""
    t = text.strip()
    if not t:
        return fallback
    t = t.replace("/", "-").replace(".", "-")
    m = re.fullmatch(r"(\d{4})-?(\d{1,2})", t)
    if m:
        y, mo = int(m.group(1)), int(m.group(2))
    else:
        m = re.fullmatch(r"(\d{1,2})-(\d{4})", t)
        if m:
            mo, y = int(m.group(1)), int(m.group(2))
        else:
            m = re.fullmatch(r"([A-Za-z]{3})[A-Za-z]*[\s-]*(\d{2}|\d{4})", t)
            if not m:
                return None
            try:
                mo = MONTHS.index(m.group(1).title()) + 1
            except ValueError:
                return None
            y = int(m.group(2))
            y = y + 2000 if y < 100 else y
    if not (1 <= mo <= 12 and 2015 <= y <= 2100):
        return None
    return f"{y:04d}-{mo:02d}"


def detect():
    """Read the master's cell colours. Returns {key: (suggested 'YYYY-MM' | None, [lines])}."""
    out = {}
    if not SRC.exists():
        return out
    try:
        wb = openpyxl.load_workbook(SRC, read_only=True, data_only=True)
    except Exception as e:
        print(f"  (could not open the master for colour detection: {e})")
        return out
    try:
        bounds = row_boundaries(wb)
    except (LayoutError, KeyError) as e:
        print(f"  (colour detection skipped - layout problem: {e})")
        return out
    finally:
        wb.close()
    for key, found in bounds.items():
        if not found:
            out[key] = (None, ["  no colour boundary found (all rows black)"])
            continue
        common, _ = Counter(found.values()).most_common(1)[0]
        lines = [f"  {LABELS[key]}: {sum(1 for v in found.values() if v == common)} of "
                 f"{len(found)} coloured rows end actuals at {pretty(common)}"]
        for label, v in found.items():
            if v != common:
                lines.append(f"      {label}: {pretty(v)}  (shaded from its own month)")
        out[key] = (common, lines)
    return out


def main():
    cur = load()
    if "--show" in sys.argv:
        for k in DEFAULTS:
            print(f"  {LABELS[k]}: last actual month = {pretty(cur[k])}")
        return 0

    detected = detect()
    if "--detect" in sys.argv:
        for k in DEFAULTS:
            sug, lines = detected.get(k, (None, ["  (not detected)"]))
            print("\n".join(lines))
            print(f"  -> suggested {LABELS[k]} boundary: {pretty(sug) if sug else '(none)'}")
        return 0

    print("-" * 60)
    print("LAST ACTUAL MONTH (everything after it is shown as forecast)")
    print("-" * 60)
    if detected:
        print("Read from the master's cell colours (black = reported, coloured = forecast):")
        for k in DEFAULTS:
            for line in detected.get(k, (None, []))[1]:
                print(line)
        print()
    print("Press Enter to accept the suggested month, or type another as YYYY-MM")
    print("(e.g. 2026-08) or 'Aug 2026'.")
    print()
    new = {}
    prev_answer = None
    for k in DEFAULTS:
        sug = detected.get(k, (None, None))[0]
        # suggested from colours > (power) the gas answer > the current value
        default = sug or prev_answer or cur[k]
        while True:
            try:
                ans = input(f"  {LABELS[k]} - last actual month [{pretty(default)}]: ")
            except EOFError:                   # non-interactive: keep current
                ans = ""
            val = parse(ans, default)
            if val is None:
                print("    Sorry, I didn't understand that - try e.g. 2026-08")
                continue
            if datetime.strptime(val, "%Y-%m") > datetime.now():
                print("    That's in the future - the last actual month can't be later than today.")
                continue
            break
        new[k] = val
        prev_answer = val

    SETTINGS.parent.mkdir(parents=True, exist_ok=True)
    SETTINGS.write_text(json.dumps(new, indent=2), encoding="utf-8")
    print()
    for k in DEFAULTS:
        tag = "" if new[k] == cur[k] else f"  (was {pretty(cur[k])})"
        print(f"  {LABELS[k]}: actuals through {pretty(new[k])}{tag}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
