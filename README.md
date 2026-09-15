# Sales Report Tool

Streamlit-based sales performance analysis tool with automatic data classification and forecast integration.

**Version:** 3.6 | **Build Date:** May 2026

---

## Quick Start

**End Users (Windows):** Double-click `SalesReportTool.exe`

**End Users (macOS, Apple Silicon):**

1. Download `SalesReportTool-macOS-arm64.zip` from the latest Release and unzip it.
2. Move `SalesReportTool.app` into `/Applications`.
3. **First launch only:** right-click the app → **Open** → **Open**. The app is unsigned, so
   double-clicking shows a Gatekeeper warning the first time.
4. Put your data files in `~/Library/Application Support/SalesReportTool/data/`
   (`Over the Years/`, `Current Year/`, `FCST/` — created automatically on first launch).

**Developers (macOS / Linux):**
```bash
pip install -r requirements.txt
chmod +x run.sh build-mac.sh   # first time only
./run.sh                       # dev server (streamlit run app/app.py)
```

**Developers (Windows):**
```powershell
.\venv311\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app/app.py       # dev server
```

### Where things live per platform

| | Windows (`.exe`) | macOS (`.app`) |
|---|---|---|
| Data folder | `data\` next to the `.exe` | `~/Library/Application Support/SalesReportTool/data/` |
| Log file | `salesreport.log` next to the `.exe` | `~/Library/Logs/SalesReportTool/salesreport.log` |
| Category overrides | `app\overrides.json` | `~/Library/Application Support/SalesReportTool/app/overrides.json` |
| Local build | `build.bat` | `./build-mac.sh` |
| CI workflow | `build-windows.yml` | `build-macos.yml` |

> A macOS `.app` bundle is read-only, so on each launch the launcher mirrors the bundled `app/`
> folder into Application Support and runs Streamlit from there (`overrides.json` is preserved).
> The code under `app/` is identical on both platforms; all platform differences live in
> `launcher.py`.

---

## Build & Release

Release builds are produced in **CI**, not locally. PyInstaller cannot cross-compile, so each
platform is packaged on its own runner:

| Platform | Workflow | Runner | Artifact |
|----------|----------|--------|----------|
| Windows | `.github/workflows/build-windows.yml` | `windows-latest` | `SalesReportTool-windows.zip` |
| macOS (Apple Silicon) | `.github/workflows/build-macos.yml` | `macos-14` (arm64) | `SalesReportTool-macOS-arm64.zip` |

### Releasing a new version (recommended)

1. Commit and push your changes.
2. Tag the release and push the tag:
   ```bash
   git tag v3.7.0
   git push origin v3.7.0
   ```
3. **Both** workflows run and attach their zip to the **same** GitHub Release for that tag.
4. End users download the zip for their platform from the Release.

> Before tagging, bump `FileVersion` / `ProductVersion` in `version_info.txt` (Windows exe version
> resource) to match the new tag — it is not derived from the tag automatically.

You can also run either workflow manually from the **Actions** tab (`workflow_dispatch`); the
`.zip` is then available as a downloadable workflow artifact.

### Local builds (smoke-test only)

| Platform | Command | Output |
|----------|---------|--------|
| Windows | `build.bat` | `dist\SalesReportTool\` (Windows / VM fallback) |
| macOS | `./build-mac.sh` | `dist/SalesReportTool.app` (arm64, unsigned) |

> `build-mac.sh` is the single source of truth for the macOS PyInstaller flags — CI calls it with
> `--skip-deps`, so local and CI builds cannot drift apart. It generates `assets/app.icns` from
> `assets/app.ico` on first run (git-ignored). When changing shared PyInstaller flags, update
> `build.bat`, `build-windows.yml` and `build-mac.sh` together.

### Code signing / notarization

Out of scope for now: the `.app` is **unsigned**, so first launch requires right-click → **Open**
(or `xattr -dr com.apple.quarantine /Applications/SalesReportTool.app`). `build-macos.yml`
contains a marked hook where `codesign` + `xcrun notarytool` should be added once a Developer ID
certificate is available.

---

## Directory Structure

```
SalesReportTool/
├── .github/
│   └── workflows/
│       ├── build-windows.yml  # CI: build Windows .exe + attach to Release
│       └── build-macos.yml    # CI: build arm64 .app + attach to same Release
├── app/                       # Platform-agnostic Streamlit app (shared, never forked)
│   ├── app.py              # Streamlit UI
│   ├── charts.py           # Altair chart functions
│   ├── fcst_loader.py      # FCST parser, blending, budget
│   ├── utils.py            # Data loading, classification, KPIs
│   ├── aliases.json        # Name alias mappings
│   └── overrides.json      # Category overrides (auto-generated)
├── data/
│   ├── Over the Years/
│   │   └── historical.csv  # All past years merged (run scripts/merge_historical.py)
│   ├── Current Year/
│   │   └── *.xlsx          # Current-year Shipping Record (Actual sheet)
│   └── FCST/               # Latest FCST xlsx (auto-selected by mtime)
├── scripts/
│   └── merge_historical.py # One-time migration: year folders → historical.csv
├── launcher.py             # Entry point; all Windows/macOS divergence lives here
├── build.bat               # Windows build (fallback)
├── build-mac.sh            # macOS arm64 .app build (also used by CI)
├── run.sh                  # macOS/Linux dev server
└── requirements.txt
```

> **Migrating from year-based folders:** run `python scripts/merge_historical.py` once to merge
> `data/2024/`, `data/2025/`, … into `data/Over the Years/historical.csv`, then move
> the current-year xlsx into `data/Current Year/`.

---

## Data Requirements

### Shipping Record — Required Columns

| Column | Notes |
|--------|-------|
| `Customer Name` | Normalized at load time |
| `Ship Date` | Fault-tolerant parse; NaT rows skipped |
| `QTY` | Used for CDR/Tablet categories |
| `SALES Total AMT` | Revenue in TWD |
| `final GP(NTD,data from Financial Report)` | Gross Profit |
| `Part Number` | — |
| `Category` | Direct category or fallback destination |

### Shipping Record — Optional Columns

| Column | Purpose |
|--------|---------|
| `DES` | Keyword-based category classification |
| `SALE_Person` | Sales rep filter |
| `Currency`, `UP`, `TP(USD)` | Shipping Record Search tab |

### FCST File

- **Location:** `data/FCST/*.xlsx` (latest by mtime)
- **Sheets:** `Div.1&2_All`, `VT`, `Signify`
- **Units:** AMT/GP stored in thousands (千元); auto-scaled ×1,000 at parse time

---

## Category Classification

Priority order:
1. **Customer Name** — `CUSTOMER_CATEGORY_MAP` in `utils.py` (e.g. SIGNIFY → Signify)
2. **Category column** — direct match (case-insensitive)
3. **DES keywords** — substring match via `DES_RULES` in `utils.py`
4. **Fallback** → Others

Valid categories: `Tablet` / `CDR` / `Tablet ACC` / `CDR ACC` / `AI_SW` / `Signify` / `Others`

---

## Configuration

### Name Aliases — `app/aliases.json`

```json
{
  "customer":      { "AZUGA INC": "AZUGA Inc." },
  "sales_person":  { "KILLIAN": "Killian Chen" },
  "fcst_customer": { "Zonar-CDR": "Zonar System Inc.", "Zonar-Tablet": "Zonar System Inc." }
}
```

- Keys must be in normalized form (uppercase for customer, Title Case for sales person)
- `fcst_customer`: maps FCST Excel names → Shipping Record canonical names; unmatched → `{sheet}_Others`

### Category Overrides — `app/overrides.json`

Auto-generated via UI. Manual format:
```json
{ "[\"Customer A\", \"PN-001\", \"2026-01\", \"desc\"]": "Tablet ACC" }
```

On macOS the packaged app writes this file to
`~/Library/Application Support/SalesReportTool/app/overrides.json`; it is preserved when you
install a newer `.app`.

### Excluded Customers — `utils.py`

```python
EXCLUDED_CUSTOMERS = {"MITAC COMPUTERKUNSHAN COLTD"}  # normalized form, no punctuation
```

---

## Tabs & Features

### Performance Report
Monthly sales trends, category breakdowns, GP%, YoY comparison, Excel export.

### Shipping Record Search
Part number keyword search, UP/TP(USD) trend, GP% analysis.

### Company Dashboard
- KPI cards: Revenue, GP, GP%, QTY, Customers, Categories (with YoY deltas)
- **Forecast row:** Full-Year Forecast (Revenue, GP, GP%, QTY)
- **Budget row:** Budget Achievement% (YTD Actual / FY Budget) + FY Budget Revenue
- Monthly trend: Actual (solid blue) / Forecast (dashed green) / Budget (dashed gray)
- Category breakdown: donut + stacked bar + AI_SW trend + FCST category chart
- Top N customers with FY Forecast and Achievement%
- **Customer Drill-Down:** per-customer blended revenue chart + FY Forecast KPIs + category/QTY/PN detail

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| "No data found" on startup | Place `historical.csv` in `data/Over the Years/` and/or xlsx in `data/Current Year/` |
| Missing columns error | Check required column names match exactly |
| Historical years missing from selector | Re-run `scripts/merge_historical.py` to regenerate `historical.csv` |
| FCST not appearing | Ensure current year selected + `.xlsx` exists in `data/FCST/` |
| FCST customer warnings | Add mapping to `aliases.json` → `fcst_customer` section |
| Name not normalizing | Check alias key is in normalized form; restart app after editing |
| Build fails | Run `pip install -r requirements.txt` first |
| CI build fails on tag push | Open the **Actions** tab → the failed **Build Windows EXE** / **Build macOS App** run and check the failed step's log |
| Antivirus (e.g. Trend Micro) quarantines `SalesReportTool.exe` as a Trojan/ML heuristic hit | Unzip to a fixed folder such as `C:\Tools\SalesReportTool\` and run it from there rather than directly from the zip or a temp/Downloads folder; submit the exe's SHA256 (printed at the end of the build) to Trend Micro / IT for allowlisting or a false-positive report |

### macOS specific

| Issue | Fix |
|-------|-----|
| "App is damaged and can't be opened" / unidentified developer | Right-click → **Open**, or `xattr -dr com.apple.quarantine /Applications/SalesReportTool.app` |
| Nothing happens after launch | Check `~/Library/Logs/SalesReportTool/salesreport.log` (also reachable via menu bar icon → **Open Log**) |
| App opens but shows no data | Data must be in `~/Library/Application Support/SalesReportTool/data/` (not inside the `.app`) |
| Want a clean slate | Delete `~/Library/Application Support/SalesReportTool/` — it is recreated on next launch |
| `build-mac.sh` fails at `iconutil` | Run it on macOS with Xcode command line tools installed (`sips` / `iconutil` are macOS-only) |
| Second launch does nothing visible | Single-instance protection: the running instance's browser tab is reopened instead |

---

## Change Log

### Unreleased
- macOS (Apple Silicon) packaging: `build-mac.sh` now produces a real arm64 `SalesReportTool.app`
  (unsigned), with `assets/app.icns` generated from `assets/app.ico`
- New `build-macos.yml` workflow: `macos-14` runner, attaches `SalesReportTool-macOS-arm64.zip`
  to the same Release as the Windows zip on `vX.Y.Z` tags
- `launcher.py` now resolves macOS-specific log, app and data locations under `~/Library`;
  `app/` code remains platform-agnostic

### v3.9 (September 2026)
- Hardened Windows build to reduce antivirus false positives (Trend Micro ML engine flagged the
  packaged exe as `Troj.Win32.TRX.XXPE50FFF109`): disabled UPX compression (`--noupx`) and added a
  Windows version resource (`version_info.txt` via `--version-file`) so the exe carries proper
  CompanyName/ProductName/FileVersion metadata
- Build scripts (`build.bat`, `build-windows.yml`) now print the SHA256 of the built exe, for
  submitting to Trend Micro / IT allowlisting
- Troubleshooting: added guidance to unzip to a fixed folder (e.g. `C:\Tools\SalesReportTool\`)
  rather than running from the zip or a temp folder

### v3.6 (May 2026)
- Data folder restructure: year-based `data/{year}/` replaced with `data/Over the Years/historical.csv` (all past years) + `data/Current Year/*.xlsx` (current year)
- `scripts/merge_historical.py`: one-time migration helper to merge year folders into `historical.csv` (UTF-8-BOM)
- Year selector now derived from Ship Date values in loaded data; historical years from CSV are automatically available
- YoY comparison simplified: both years come from a single combined DataFrame

### v3.5 (April 2026)
- Budget integration: `agg_budget_monthly()`, Budget Achievement% KPI cards, dashed-gray Budget line in charts
- Customer Drill-Down: per-customer FCST blend with FY Forecast KPIs and blended revenue chart
- Unmatched FCST customer warnings surfaced in Dashboard body

### v3.4 (April 2026)
- Signify as independent product category (DES keyword + customer name override + purple chart color)
- `EXCLUDED_CUSTOMERS` uses normalized (no-punctuation) customer name

### v3.3 (April 2026)
- `fcst_loader.py`: FCST blend engine, customer name mapping, AMT/GP ×1,000 auto-scaling
- Company Dashboard: FY Forecast KPI row, Actual/Forecast trend charts, FCST category chart

### v3.2 (April 2026)
- Customer/sales person name normalization with `aliases.json` alias maps

### v3.1
- DES keyword classification, Shipping Record Search, Company Dashboard KPIs, override system

---

*For internal use. Last Updated: 2026-09-14*
