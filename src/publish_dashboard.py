"""
Palissy Dashboard Publisher (main dashboard)

Publishes output/index.html as the live main dashboard:
    https://tahershehabi99.github.io/palissy-gas-dashboard/

All publishing goes through src/site_publish.py, which keeps a private working copy of the GitHub
repository outside OneDrive and only ever touches this page's own files (docs/index.html and
output/index.html in the repository). The other pages in the same repository - Pipeline Imports,
Demand, Global Power, Pricing - have their own folders and are never touched from here, and they
never touch the main dashboard.

After a successful publish, docs/index.html in this folder is updated too: it is the "last
published" copy that the Demand and Global Power pages read the gas model from.

Run AFTER update_dashboard.bat has regenerated output/index.html (and you've previewed it).

Usage:
    py -3 src/publish_dashboard.py
or double-click publish.bat
"""
import os
import shutil
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
OUTPUT_FILE = os.path.join(PROJECT_DIR, "output", "index.html")
DOCS_FILE = os.path.join(PROJECT_DIR, "docs", "index.html")

sys.path.insert(0, SCRIPT_DIR)
from site_publish import publish_page, page_url  # noqa: E402


def main():
    print("=" * 60)
    print("Palissy Dashboard Publisher")
    print("=" * 60)

    if not os.path.exists(OUTPUT_FILE):
        print(f"\nERROR: {OUTPUT_FILE} not found.")
        print("  Run update_dashboard.bat first to generate it.")
        sys.exit(1)

    size_kb = os.path.getsize(OUTPUT_FILE) // 1024
    print(f"\nPublishing output/index.html ({size_kb} KB)...")
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    ok, msg = publish_page("main", OUTPUT_FILE, f"Update dashboard data ({stamp})")
    if not ok:
        print("\nERROR: not published.")
        print("  " + msg)
        print("\n  Common causes:")
        print("  - No internet connection")
        print("  - GitHub login expired on this computer (git will ask you to sign in again)")
        print("  Nothing was changed on the live site. Fix and run publish.bat again.")
        sys.exit(1)

    print("  " + msg)
    os.makedirs(os.path.dirname(DOCS_FILE), exist_ok=True)
    shutil.copyfile(OUTPUT_FILE, DOCS_FILE)
    print("  docs/index.html (the 'last published' copy) updated.")

    print("\n" + "=" * 60)
    print("DONE. Live site updates in 30-90 seconds:")
    print(f"  {page_url('main')}")
    print("(hard-refresh with Ctrl+Shift+R if you see old data)")
    print("=" * 60)


if __name__ == "__main__":
    main()
