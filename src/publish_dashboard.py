"""
Palissy Dashboard Publisher

Syncs output/index.html -> docs/index.html, commits the change,
and pushes to GitHub so the live site updates.

Run AFTER update_dashboard.bat has regenerated output/index.html
(and you've previewed it locally).

Usage:
    py -3 src/publish_dashboard.py
or double-click publish.bat
"""

import os
import shutil
import subprocess
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
OUTPUT_FILE = os.path.join(PROJECT_DIR, "output", "index.html")
DOCS_FILE = os.path.join(PROJECT_DIR, "docs", "index.html")
LIVE_URL = "https://tahershehabi99.github.io/palissy-gas-dashboard/"


def run_git(args, check=True):
    result = subprocess.run(
        ["git"] + args,
        cwd=PROJECT_DIR,
        capture_output=True,
        text=True,
    )
    if check and result.returncode != 0:
        print(f"\nERROR: git {' '.join(args)} failed")
        print(result.stderr.strip())
        sys.exit(1)
    return result


def main():
    print("=" * 60)
    print("Palissy Dashboard Publisher")
    print("=" * 60)

    if not os.path.exists(OUTPUT_FILE):
        print(f"\nERROR: {OUTPUT_FILE} not found.")
        print("Run update_dashboard.bat first to generate it.")
        sys.exit(1)

    print(f"\nSyncing docs/index.html with output/index.html...")
    shutil.copyfile(OUTPUT_FILE, DOCS_FILE)
    size_kb = os.path.getsize(DOCS_FILE) // 1024
    print(f"  Copied ({size_kb} KB)")

    print("\nChecking for changes...")
    status = run_git(["status", "--porcelain", "output/index.html", "docs/index.html"])
    if not status.stdout.strip():
        print("  No changes to publish - live site already matches local build.")
        print("=" * 60)
        return

    print("  Changes detected:")
    for line in status.stdout.strip().splitlines():
        print(f"    {line}")

    print("\nStaging files...")
    run_git(["add", "output/index.html", "docs/index.html"])

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    commit_msg = f"Update dashboard data ({timestamp})"
    print(f"\nCommitting: \"{commit_msg}\"")
    run_git(["commit", "-m", commit_msg])

    print("\nPushing to GitHub...")
    push_result = run_git(["push", "origin", "master"], check=False)
    if push_result.returncode != 0:
        print("  Push failed:")
        print(push_result.stderr.strip())
        print("\n  Common causes:")
        print("  - No internet connection")
        print("  - GitHub credentials expired (try: gh auth login)")
        print("  - Someone else pushed first (try: git pull --rebase, then re-run)")
        sys.exit(1)

    print("  Push successful.")
    print("\n" + "=" * 60)
    print("DONE. Live site updates in 30-90 seconds:")
    print(f"  {LIVE_URL}")
    print("(hard-refresh with Ctrl+Shift+R if you see old data)")
    print("=" * 60)


if __name__ == "__main__":
    main()
