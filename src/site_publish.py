r"""
ONE place that publishes every page of the Palissy gas dashboard to GitHub.

Everything lives in ONE GitHub repository, Tahershehabi99/palissy-gas-dashboard (branch master,
GitHub Pages serves its docs\ folder). Each page has its own folder, and a publish only ever
touches that page's own file:

    main dashboard    docs\index.html            https://tahershehabi99.github.io/palissy-gas-dashboard/
    Pipeline Imports  docs\imports\index.html    .../palissy-gas-dashboard/imports/
    Demand            docs\demand\index.html     .../palissy-gas-dashboard/demand/
    Global Power      docs\power\index.html      .../palissy-gas-dashboard/power/
    Pricing           docs\pricing\index.html    .../palissy-gas-dashboard/pricing/

Used by: src\publish_dashboard.py (main dashboard, publish.bat) and the daily page builders in the
sibling projects (Imports, ENTSOG, Ember Electricity, Pricing).

How a publish works (so pages can never overwrite each other):
  1. A private working copy of the repository is kept OUTSIDE OneDrive, in
     %USERPROFILE%\palissy-dashboard-site. It is created automatically the first time.
  2. It is first reset to exactly what GitHub has now.
  3. Only this page's file is copied in and committed, then pushed.
  4. If another page was published in the same minute, steps 2-3 are simply repeated.
Nothing in a project folder (OneDrive) is ever changed by this module.

    from site_publish import publish_page, page_url
    ok, msg = publish_page("imports", r"...\site\index.html")

To add a new page: add one line to PAGES below.
"""
import os
import shutil
import subprocess
import time
from datetime import datetime

REPO_URL = os.environ.get("PALISSY_SITE_REPO", "https://github.com/Tahershehabi99/palissy-gas-dashboard.git")
BRANCH = "master"
SITE_ROOT = "https://tahershehabi99.github.io/palissy-gas-dashboard/"
CLONE_DIR = os.environ.get("PALISSY_SITE_CLONE", os.path.join(os.path.expanduser("~"), "palissy-dashboard-site"))

# page name -> (label, files inside the repository that receive the page, address under SITE_ROOT)
PAGES = {
    "main":    ("Main dashboard",   ["docs/index.html", "output/index.html"], ""),
    "imports": ("Pipeline Imports", ["docs/imports/index.html"], "imports/"),
    "demand":  ("Demand",           ["docs/demand/index.html"], "demand/"),
    "power":   ("Global Power",     ["docs/power/index.html"], "power/"),
    "pricing": ("Pricing",          ["docs/pricing/index.html"], "pricing/"),
}
ATTEMPTS = 4
LOCK_WAIT_SEC = 300
LOCK_STALE_SEC = 900


def page_url(name):
    """Public address of a page."""
    return SITE_ROOT + PAGES[name][2]


def _git(args, cwd, timeout=180):
    try:
        r = subprocess.run(["git"] + args, cwd=cwd, capture_output=True, text=True,
                           stdin=subprocess.DEVNULL, timeout=timeout)
    except subprocess.TimeoutExpired:
        return 1, f"git {' '.join(args[:2])} timed out (is git waiting for a GitHub login?)"
    return r.returncode, (r.stdout + r.stderr).strip()


class _Lock:
    """Only one publish at a time on this computer (the pages share one working copy)."""

    def __init__(self):
        self.path = CLONE_DIR.rstrip("\\/") + ".lock"
        self.held = False

    def __enter__(self):
        t0 = time.time()
        while True:
            try:
                fd = os.open(self.path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                os.write(fd, f"{os.getpid()} {datetime.now():%Y-%m-%d %H:%M:%S}".encode())
                os.close(fd)
                self.held = True
                return self
            except FileExistsError:
                try:
                    if time.time() - os.path.getmtime(self.path) > LOCK_STALE_SEC:
                        os.remove(self.path)      # left behind by a run that was killed
                        continue
                except OSError:
                    pass
                if time.time() - t0 > LOCK_WAIT_SEC:
                    raise TimeoutError("another publish has been running for more than 5 minutes")
                time.sleep(3)

    def __exit__(self, *exc):
        if self.held:
            try:
                os.remove(self.path)
            except OSError:
                pass


def _sync_clone():
    """Make the working copy identical to GitHub. Returns (ok, message)."""
    if not os.path.isdir(os.path.join(CLONE_DIR, ".git")):
        if os.path.isdir(CLONE_DIR) and os.listdir(CLONE_DIR):
            return False, f"{CLONE_DIR} exists but is not a working copy - rename or remove that folder"
        rc, out = _git(["clone", "--depth", "1", "--branch", BRANCH, REPO_URL, CLONE_DIR],
                       cwd=os.path.expanduser("~"), timeout=600)
        if rc:
            return False, f"could not download the repository ({REPO_URL}): {out[-300:]}"
        return True, "working copy created"
    rc, out = _git(["fetch", "--depth", "1", "origin", BRANCH], CLONE_DIR, timeout=300)
    if rc:
        return False, f"could not reach GitHub: {out[-300:]}"
    rc, out = _git(["reset", "--hard", "FETCH_HEAD"], CLONE_DIR)
    if rc:
        return False, f"could not refresh the working copy: {out[-300:]}"
    return True, "working copy refreshed"


def _identity_args():
    rc, out = _git(["config", "user.email"], CLONE_DIR)
    if rc == 0 and out.strip():
        return []
    return ["-c", "user.name=Taher Shehabi", "-c", "user.email=taher@akapenergy.com"]


def publish_page(name, src_file, message=None):
    """Publish one page. Returns (ok, message). Never raises for ordinary failures."""
    if name not in PAGES:
        return False, f"unknown page '{name}' (known: {', '.join(PAGES)})"
    label, targets, _ = PAGES[name]
    if not os.path.isfile(src_file):
        return False, f"page file not found: {src_file}"
    if os.path.getsize(src_file) < 1000:
        return False, f"page file looks empty ({os.path.getsize(src_file)} bytes): {src_file}"
    if shutil.which("git") is None:
        return False, "git is not installed on this computer"
    message = message or f"{label} update {datetime.now():%Y-%m-%d %H:%M}"
    try:
        with _Lock():
            last = ""
            for attempt in range(1, ATTEMPTS + 1):
                ok, msg = _sync_clone()
                if not ok:
                    last = msg
                    time.sleep(5 * attempt)
                    continue
                for rel in targets:
                    dst = os.path.join(CLONE_DIR, *rel.split("/"))
                    os.makedirs(os.path.dirname(dst), exist_ok=True)
                    shutil.copyfile(src_file, dst)
                rc, out = _git(["add", "--"] + targets, CLONE_DIR)
                if rc:
                    return False, f"git add failed: {out[-300:]}"
                if _git(["diff", "--cached", "--quiet"], CLONE_DIR)[0] == 0:
                    return True, f"live page already matches ({page_url(name)})"
                rc, out = _git(_identity_args() + ["commit", "-m", message, "--"] + targets, CLONE_DIR)
                if rc:
                    return False, f"git commit failed: {out[-300:]}"
                rc, out = _git(["push", "origin", f"HEAD:{BRANCH}"], CLONE_DIR, timeout=300)
                if rc == 0:
                    return True, f"published to {page_url(name)}"
                last = f"git push failed: {out[-300:]}"     # usually: another page was published just now
                time.sleep(5 * attempt)
            return False, f"{last} (after {ATTEMPTS} attempts)"
    except TimeoutError as e:
        return False, str(e)
    except OSError as e:
        return False, f"file error while publishing: {e}"


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("usage: python site_publish.py <page name> <page file>\npages: " + ", ".join(PAGES))
        sys.exit(2)
    ok_, msg_ = publish_page(sys.argv[1], sys.argv[2])
    print(msg_)
    sys.exit(0 if ok_ else 3)
