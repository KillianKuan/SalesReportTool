# Sales Report Tool

Streamlit-based sales performance analysis tool with automatic data classification and forecast integration.

**Version:** 4.1 | **Build Date:** September 2026

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
| User settings (ignore list / account match) | `app\settings.json` | `~/Library/Application Support/SalesReportTool/app/settings.json` |
| Local build | `build.bat` | `./build-mac.sh` |
| CI workflow | `build-windows.yml` | `build-macos.yml` |

> A macOS `.app` bundle is read-only, so on each launch the launcher mirrors the bundled `app/`
> folder into Application Support and runs Streamlit from there (`overrides.json` and
> `settings.json` are preserved). The code under `app/` is identical on both platforms; all
> platform differences live in `launcher.py`.

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
│   ├── app.py              # Streamlit UI, sidebar page navigation
│   ├── charts.py           # Altair chart functions
│   ├── components.py       # Neutral kpi_card() / card_title() UI helpers
│   ├── palette.py          # Fixed (non-theme) chart mark + KPI delta colors
│   ├── fcst_loader.py      # FCST parser, blending, budget
│   ├── utils.py            # Data loading, classification, KPIs, layout CSS
│   ├── aliases.json        # Name alias mappings (shipped defaults, git-tracked)
│   ├── overrides.json      # Category overrides (auto-generated)
│   └── settings.json       # User settings: ignore list, account match (auto-generated)
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

Most day-to-day configuration — ignored customers and FCST↔Performance Report account
mapping — is meant to be done from the **⚙️ Settings** page in the app (see
[Pages & Features](#pages--features)), which writes to `app/settings.json`. The files below are the
underlying storage and the shipped defaults that `settings.json` layers on top of.

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
- **This file is the shipped default and the app never writes to it.** It's git-tracked and meant
  for developers to edit; end users add/override mappings from **⚙️ Settings ▸ Account Match**
  instead, which is layered on top at load time (settings override, defaults fill the rest).

### Category Overrides — `app/overrides.json`

Auto-generated via UI. Manual format:
```json
{ "[\"Customer A\", \"PN-001\", \"2026-01\", \"desc\"]": "Tablet ACC" }
```

On macOS the packaged app writes this file to
`~/Library/Application Support/SalesReportTool/app/overrides.json`; it is preserved when you
install a newer `.app`.

### User Settings — `app/settings.json`

Auto-generated the first time a user changes anything in the **⚙️ Settings** page. Same
persistence model as `overrides.json` — user-writable, preserved across restarts and `.app`
upgrades (see `launcher.py`'s `USER_STATE_FILES`), never committed with real data.

```json
{
  "ignored_customers": ["MITAC COMPUTERKUNSHAN COLTD"],
  "customer_aliases": {},
  "fcst_customer_aliases": {},
  "custom_groups": []
}
```

> A `settings.json` written before v4.1 may still have a `"theme"` key (the app no longer has a
> theme system — colors come entirely from Streamlit's native light/dark theme). It's harmless:
> `load_settings()` only reads keys present in `DEFAULT_SETTINGS`, so a legacy `theme` key is
> silently ignored and dropped the next time settings are saved.

- `ignored_customers`: customer names (any casing/punctuation) to drop entirely from both the
  Performance Report and FCST pipelines; defaults to the legacy `EXCLUDED_CUSTOMERS` set
- `customer_aliases` / `fcst_customer_aliases`: user overrides layered on top of `aliases.json`'s
  `customer` / `fcst_customer` sections
- `custom_groups`: extra "Others"-style buckets (e.g. `"Others - EMEA Distributors"`) that unmatched
  FCST customers can be assigned to instead of the default `{sheet}_Others`
- Any change here invalidates the relevant `@st.cache_data` caches immediately (`utils._rules_key()`
  and `fcst_loader.load_fcst()` both fold in a hash of the current settings)

### Excluded Customers — default source

```python
# utils.py — DEFAULT_SETTINGS["ignored_customers"] seeds from this constant
EXCLUDED_CUSTOMERS = {"MITAC COMPUTERKUNSHAN COLTD"}  # normalized form, no punctuation
```

This is only the shipped **default** for `settings.json`'s `ignored_customers` list — the actual
filtering happens once, inside `load_single_file()` / `load_historical_csv()` /
`fcst_loader._parse_sheet()`, using whatever is currently in `settings.json`. End users manage the
list from **⚙️ Settings ▸ Customer Ignore List**, which also shows how many rows are being excluded.

---

## Pages & Features

Navigation is a sidebar page list (`Company Dashboard` / `Performance Report` / `Shipping Record
Search`, with `Settings` pinned at the bottom, below the collapsed `FCST` and `System Info`
sections) rather than horizontal tabs — only the active page's content renders in the main area,
in a bordered-card layout. The Sales Person filter at the top of the sidebar, and the FCST sheet
picker, apply across all pages; switching pages preserves each page's own filter state (year
selection, search text, etc.) via Streamlit session state.

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

### Settings
UI-based configuration, persisted to `app/settings.json` and effective immediately (no restart,
no editing JSON by hand). Two sections, switched with a radio control rather than nested tabs
(`st.tabs()` resets to its first item on the programmatic rerun each save triggers):

- **🚫 Customer Ignore List** — add customers (from current data or free text) or remove them from
  the ignore list; shows how many rows are currently excluded, split by Performance Report vs. FCST.
- **🔗 Account Match** — a "Needs Mapping" table lists every unmatched FCST customer (name, sheet,
  forecast amount) with an inline selector to assign it to an existing Performance Report customer
  or a custom group; also supports creating/renaming/deleting custom groups and editing the plain
  Shipping Record customer-name aliases, with validation warnings for self-referencing or dangling
  mappings.

Each section has its own "Reset to Default" button, plus a "Reset ALL Settings" button at the
bottom (with confirmation). There is no app-level theme setting — the app follows whichever
light/dark theme is configured in Streamlit itself.

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| "No data found" on startup | Place `historical.csv` in `data/Over the Years/` and/or xlsx in `data/Current Year/` |
| Missing columns error | Check required column names match exactly |
| Historical years missing from selector | Re-run `scripts/merge_historical.py` to regenerate `historical.csv` |
| FCST not appearing | Ensure current year selected + `.xlsx` exists in `data/FCST/` |
| FCST customer warnings | Assign the customer from **⚙️ Settings ▸ Account Match ▸ Needs Mapping** (no restart needed); developers can also add a shipped default to `aliases.json` → `fcst_customer` |
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

### v4.1 (September 2026)
- **Sidebar page navigation**: the horizontal main tabs (Performance Report / Shipping Record
  Search / Company Dashboard / Settings) were replaced with a sidebar page list — Company
  Dashboard, Performance Report, Shipping Record Search, then the collapsed FCST and System Info
  sections, then Settings pinned at the bottom. Only the active page renders in the main area;
  each page's own filters keep their state when you switch pages.
- **Removed the app-level Theme system entirely.** There is no more Light/Dark/System setting —
  the app now relies solely on Streamlit's own native theme (configured via `.streamlit/config.toml`
  or the viewer's own Settings menu), which already adapts colors, borders, and chart rendering to
  light/dark automatically.
  - Deleted `app/theme.py` (the Light/Dark design-token module), `utils.inject_theme_css()`,
    `utils.resolve_theme_mode()`, and `charts.apply_altair_theme()`. Charts no longer take a `mode`
    argument or register a competing Altair theme — `st.altair_chart()`'s default
    `theme="streamlit"` now themes axes/legends/text automatically, matching the app's actual theme
    instead of a value resolved once at startup.
  - `kpi_card()` / `card_title()` moved to the new neutral `app/components.py`; their markup is
    unchanged but they now use Streamlit's own `var(--text-color)` for text instead of custom
    tokens.
  - Fixed (non-Light/Dark) mark colors for charts and the KPI delta pills moved to the new
    `app/palette.py` — a single color set, not a light/dark pair, since chart marks need a real
    color and can't inherit CSS variables the way UI chrome can.
  - `utils.inject_layout_css()` replaces `inject_theme_css()`: spacing, radius, control max-width,
    and the sidebar's own fixed brand-green background are layout-only CSS now; card/text colors
    are left entirely to Streamlit.
  - Removed the `theme` key from `DEFAULT_SETTINGS` / the Settings ▸ Theme section. A
    `settings.json` written by an older build that still has a `"theme"` key keeps loading
    normally — the key is just ignored and dropped on next save.

### v4.0 (September 2026)
- **UI redesign**: the visual layer was rebuilt around a single design-token module
  (`app/theme.py`) — deep-green chrome, mint canvas, white/dark rounded cards, and Material
  Symbols icons in place of emoji throughout. No data loading, blending, or KPI math changed.
  - `app/theme.py` is now the only place colors, radii, spacing and category/source palettes are
    defined; `get_tokens(mode)`, `CATEGORY_COLORS(mode)` and `SOURCE_COLORS(mode)` resolve light
    vs. dark values. `utils.inject_theme_css()` emits them as CSS custom properties, and
    `utils.resolve_theme_mode()` picks an explicit light/dark for chart rendering (the CSS itself
    still follows `prefers-color-scheme` live when the Settings theme is "System").
  - `charts.py` registers one shared Altair theme (`apply_altair_theme(mode)`) — token colors,
    horizontal-only gridlines, no chart border, top legends, and abbreviated numeric axis labels
    (1.2M / 340K) — and every chart now takes an explicit `mode` instead of hardcoded hex colors.
  - New `kpi_card()` / `card_title()` helpers (in `theme.py`) replace `st.metric()` and the old
    double-title pattern; every dashboard section is now a bordered card with consistent spacing.
  - Company Dashboard's first screen is a `[2, 1]` layout: the monthly trend chart on the left,
    a stacked KPI summary + compact Top Customers list on the right.
  - Sidebar: the per-customer Sales Person checkbox list became a single multiselect; the FCST
    sheet picker and System Info moved into expanders so the sidebar fits one screen.
- New **⚙️ Settings** tab: Theme (Light/Dark/System), Customer Ignore List, and Account Match
  (FCST ↔ Performance Report customer mapping with custom groups), all persisted to the new
  `app/settings.json` and effective immediately — see [Settings](#settings) and
  [User Settings — `app/settings.json`](#user-settings--appsettingsjson)
- `EXCLUDED_CUSTOMERS` migrated from a hardcoded constant to a `settings.json`-backed, user-editable
  ignore list; filtering now happens once, inside the loaders themselves, for both the Performance
  Report and FCST pipelines
- `aliases.json`'s `customer` / `fcst_customer` sections can now be extended/overridden per-user via
  `settings.json` without editing the shipped file
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

*For internal use. Last Updated: 2026-09-15*
