# Palissy gas dashboard

Everything the dashboard shows is published from this one repository. GitHub Pages serves the `docs/` folder.

| Page | File here | Address | Published by |
|---|---|---|---|
| Main dashboard (Gas Balance, Global LNG, Europe power, Map) | `docs/index.html` | https://tahershehabi99.github.io/palissy-gas-dashboard/ | `publish.bat`, by hand, when the gas or LNG model changes |
| Pipeline Imports | `docs/imports/index.html` | `/imports/` | dev machine, every morning after the imports pipeline |
| Demand | `docs/demand/index.html` | `/demand/` | dev machine, every morning after the demand pipeline |
| Global Power (Asia + Combined) | `docs/power/index.html` | `/power/` | dev machine, weekdays, when Ember data changes |
| Pricing | `docs/pricing/index.html` | `/pricing/` | dev machine, daily pricing job |

Each page is one self-contained file in its own folder. Every publish goes through `src/site_publish.py`,
which only ever commits that page's own file, so the pages cannot overwrite each other.
Storage and LNG Sendout are still served from the older `eu-gas-dashboard` site.
