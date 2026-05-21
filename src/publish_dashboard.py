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

# Hard ceiling on any single git call. If git hangs on a credential prompt
# (despite stdin=DEVNULL it can still wait on a GUI helper), we surface a
# clear error instead of leaving the user staring at a frozen window.
GIT_TIMEOUT_SEC = 90


def run_git(args, check=True, timeout=GIT_TIMEOUT_SEC):
    try:
        result = subprocess.run(
            ["git"] + args,
            cwd=PROJECT_DIR,
            capture_output=True,
            text=True,
            stdin=subprocess.DEVNULL,   # never wait for terminal input
            timeout=timeout,
        )
    except subprocess.TimeoutExpired:
        print(f"\nERROR: git {' '.join(args)} timed out after {timeout}s.")
        print("  This usually means git is waiting on a credential prompt.")
        print("  Open Git Bash or PowerShell and run:")
        print("    git -C \"" + PROJECT_DIR + "\" push origin master")
        print("  Follow any login prompts that appear, then retry publish.bat.")
        sys.exit(1)
    if check and result.returncode != 0:
        print(f"\nERROR: git {' '.join(args)} failed (exit {result.returncode})")
        if result.stdout.strip():
            print(result.stdout.strip())
        if result.stderr.strip():
            print(result.stderr.strip())
        sys.exit(1)
    return result


def preflight():
    if shutil.which("git") is None:
        print("\nERROR: git is not on PATH.")
        print("  Install Git for Windows from https://git-scm.com/ and reopen the terminal.")
        sys.exit(1)
    if not os.path.exists(os.path.join(PROJECT_DIR, ".git")):
        print("\nERROR: this folder is not a git repository.")
        print(f"  Expected .git/ in: {PROJECT_DIR}")
        sys.exit(1)
    if not os.path.exists(OUTPUT_FILE):
        print(f"\nERROR: {OUTPUT_FILE} not found.")
        print("  Run update_dashboard.bat first to generate it.")
        sys.exit(1)


def main():
    print("=" * 60)
    print("Palissy Dashboard Publisher")
    print("=" * 60)

    preflight()

    print("\nSyncing docs/index.html with output/index.html...")
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
    print(f'\nCommitting: "{commit_msg}"')
    run_git(["commit", "-m", commit_msg])

    print("\nPushing to GitHub...")
    push_result = run_git(["push", "origin", "master"], check=False)
    if push_result.returncode != 0:
        print("  Push failed:")
        if push_result.stdout.strip():
            print(push_result.stdout.strip())
        if push_result.stderr.strip():
            print(push_result.stderr.strip())
        print("\n  Common causes:")
        print("  - No internet connection")
        print("  - GitHub credentials expired -> run: gh auth login")
        print("  - Someone else pushed first -> run: git pull --rebase, then re-run")
        print("\n  Note: the local commit is already saved. You can re-run publish.bat after fixing.")
        sys.exit(1)

    print("  Push successful.")
    print("\n" + "=" * 60)
    print("DONE. Live site updates in 30-90 seconds:")
    print(f"  {LIVE_URL}")
    print("(hard-refresh with Ctrl+Shift+R if you see old data)")
    print("=" * 60)


if __name__ == "__main__":
    main()
