"""
LNG update PREFLIGHT - validate the master's STRUCTURE before extracting/publishing.

Run before every refresh. Locates the Monthly-Exports blocks by anchor (catching
layout changes) and compares the project roster against the last known-good
manifest (INPUT/lng_manifest.json):

  exit 0  -> SAFE to proceed. Either no structural change, or new project(s) in an
             EXISTING region/country (their names are printed so colour/region
             mapping can be checked after the build).
  exit 2  -> STOP, needs review: a missing anchor / layout change, a removed or
             renamed project, or a NEW region/country. Call Claude before updating.

Roster changes are found by NAME (a set diff), NOT by row position, so a project
inserted anywhere in the list is reported precisely.

  py -3 src/preflight_lng.py            # check against the manifest
  py -3 src/preflight_lng.py --write    # (re)initialise the manifest from the master
                                        #   (done automatically after an approved update)
"""
import json
import sys
from pathlib import Path

import openpyxl

from extract_lng_input import SRC, REGION_OF, COUNTRIES, _clean, locate_exports

ROOT = Path(__file__).resolve().parent.parent
MANIFEST = ROOT / "WORKING" / "lng_manifest.json"   # pipeline state (not in INPUT)


def snapshot():
    """Structural fingerprint of the master: project roster (name -> country/region/
    status) + the country/region sets + the anchor block positions."""
    wb = openpyxl.load_workbook(SRC, data_only=True)
    try:
        me = wb["Monthly Exports"]
        asm = wb["Assumptions"]
        loc = locate_exports(me)                       # raises if an anchor is missing
        status_of = {}
        for r in range(3, asm.max_row + 1):
            nm = _clean(asm.cell(r, 1).value)
            if nm and nm != "Total":
                status_of[nm] = _clean(asm.cell(r, 3).value) or ""
        # Walk the Reported section exactly as the extractor does: mains (M) only,
        # country assigned on the country aggregate row.
        projects, pending = {}, []
        for r in range(loc["rep"][0], loc["rep"][1] + 1):
            nm = _clean(me.cell(r, 1).value)
            if not nm:
                continue
            mk = _clean(me.cell(r, 3).value)
            if mk in ("M", "T"):
                pending.append((nm, mk))
            elif nm in COUNTRIES:
                for pn, pm in pending:
                    if pm == "M":
                        projects[pn] = {"country": nm, "region": REGION_OF.get(nm, nm),
                                        "status": status_of.get(pn, "")}
                pending = []
            elif nm == "Grand total":
                pending = []
        return {
            "projects": projects,
            "countries": sorted({p["country"] for p in projects.values()}),
            "regions": sorted({p["region"] for p in projects.values()}),
            "anchors": {k: (list(v) if isinstance(v, tuple) else v) for k, v in loc.items()},
        }
    finally:
        wb.close()


def main():
    write = "--write" in sys.argv
    try:
        snap = snapshot()
    except Exception as e:
        print(f"[PREFLIGHT] LAYOUT ERROR: {e}")
        print("STOP - the master's layout could not be read by the anchors. Call Claude.")
        return 2

    if write or not MANIFEST.exists():
        first = not MANIFEST.exists()
        MANIFEST.parent.mkdir(exist_ok=True)
        MANIFEST.write_text(json.dumps(snap, indent=1))
        print(f"[PREFLIGHT] manifest {'initialised' if first else 'rewritten'}: "
              f"{len(snap['projects'])} projects, {len(snap['countries'])} countries, "
              f"{len(snap['regions'])} regions")
        return 0

    old = json.loads(MANIFEST.read_text())
    on, nn = set(old["projects"]), set(snap["projects"])
    added = sorted(nn - on)
    removed = sorted(on - nn)
    statchg = [(p, old["projects"][p]["status"], snap["projects"][p]["status"])
               for p in sorted(on & nn)
               if old["projects"][p]["status"] != snap["projects"][p]["status"]]
    new_regions = sorted(set(snap["regions"]) - set(old["regions"]))
    new_countries = sorted(set(snap["countries"]) - set(old["countries"]))

    print(f"[PREFLIGHT] master: {len(snap['projects'])} projects | "
          f"added {len(added)}, removed {len(removed)}, status-changes {len(statchg)}")
    for p in added:
        print(f"   + NEW PROJECT: {p}  ({snap['projects'][p]['country']}, {snap['projects'][p]['status']})")
    for p in removed:
        print(f"   - REMOVED/RENAMED: {p}")
    for p, o, n in statchg:
        print(f"   ~ STATUS CHANGE: {p}: {o or '(blank)'} -> {n or '(blank)'}")
    if new_regions:
        print(f"   ! NEW REGION(S): {new_regions}")
    if new_countries:
        print(f"   ! NEW COUNTRY(IES): {new_countries}")

    if removed or new_regions or new_countries:
        print("STOP - structural change needs review (removed/renamed project or new "
              "region/country). Call Claude before updating.")
        return 2
    if added:
        print("PROCEED (with notice) - new project(s) in existing regions; check their "
              "colour/region mapping after the build.")
        return 0
    print("OK - no structural change. Safe to auto-update (value/status updates only).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
