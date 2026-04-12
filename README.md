# Sales Report Tool

A Streamlit-based interactive data analysis platform for sales performance reporting with intelligent data normalization, classification, and forecast integration.

**Version:** 3.3 | **Status:** Production-ready  
**Build Date:** April 2026

---

## 📋 Table of Contents

1. [Quick Start](#quick-start)
2. [Features](#features)
3. [System Architecture](#system-architecture)
4. [Data Requirements](#data-requirements)
5. [Configuration](#configuration)
6. [Installation & Build](#installation--build)
7. [Usage](#usage)
8. [Troubleshooting](#troubleshooting)

---

## ⚡ Quick Start

### For End Users (Packaged Version)

1. **Download & Extract:** Unzip the distributed `SalesReportTool` folder
2. **Run:** Double-click `SalesReportTool.exe`
3. **Select Year(s):** Choose data year(s) from the sidebar
4. **Analyze:** Explore Performance Report, Shipping Record Search, or Company Dashboard

### For Developers (Source Code)

```bash
# Clone/navigate to workspace
cd SalesReportTool

# Activate venv
./venv311/Scripts/Activate.ps1

# Install dependencies
pip install -r requirements.txt

# Run locally
streamlit run app/app.py

# Build executable
build.bat
```

---

## ✨ Features

### Core Functionality

- **Performance Report** — Multi-year sales trends, category breakdowns, margin analysis
- **Shipping Record Search** — Currency conversion, unit price tracking, cost-per-unit analysis
- **Company Dashboard** — KPI tracking, YoY comparison, monthly trends, customer rankings, **Forecast integration**
- **Sales Person Filter** — Sidebar filter to segment data by sales representative (current year)

### Data Intelligence

#### 1. Automatic Category Classification
- **Direct Matching:** Reads `Category` column with normalization
- **Keyword-based (DES):** Falls back to `DES` field pattern matching:
  - CDR ACC: cdr, gemini, evo, sprint, sd card, panic button, iosix, uvc camera, k220, k245, k265, smart link dongle, safetycam
  - Tablet ACC: tablet, prometheus, chiron, hera, phaeton, surfing pro, cradle, f840, ulmo, fleet cable
  - AI_SW: visionmax
  - Others: catch-all
- **Valid Categories:** `Tablet`, `CDR`, `Tablet ACC`, `CDR ACC`, `AI_SW`, `Others`

#### 2. Name Normalization (v3.2+)
Automatically normalizes customer names and sales person names to handle variations:

**Customer Names:**
- Removes punctuation (commas, periods)
- Compresses whitespace
- Converts to UPPERCASE
- Optionally applies alias maps (see [Configuration](#configuration))

**Sales Person:**
- Removes punctuation
- Compresses whitespace
- Converts to Title Case
- Optionally applies alias maps

**Example:**
```
"AZUGA, INC."  →  "AZUGA INC"
"azuga  inc"   →  "AZUGA INC"
"Killian Chen" →  "Killian Chen"
"killian"      →  (optional alias) "Killian Chen"
```

#### 3. Override System
Store custom category overrides in `app/overrides.json` using composite keys:
```json
{
  "[\"Customer\", \"PN\", \"2026-01\", \"DES Description\"]": "Tablet ACC"
}
```
- Persists across sessions
- Keyed by: Customer Name, Part Number, Month, DES (if present)
- Applied after initial classification

#### 4. Forecast Integration (v3.3+)
Company Dashboard blends Shipping Record actuals with FCST file data:

- **Data Source:** Latest `.xlsx` in `data/FCST/`; supports sheets `Div.1&2_All`, `VT`, `Signify`
- **Blend Logic:** Past months use Actual; current month and future months use Forecast
- **Sidebar Control:** FCST Sheet selector (`All Sheets` / `Div.1&2_All` / `VT` / `Signify`)
- **KPI Row:** Full-Year Forecast (Revenue, GP, GP%, QTY) shown below existing KPI cards
- **Monthly Trend Charts:** Actual = solid blue line, Forecast = dashed green line, with a boundary marker
- **Category Chart:** 4th column shows FCST category revenue breakdown
- **Customer Mapping:** `FCST_TO_CANONICAL` dict maps FCST file names to Performance Report canonical names; unmatched customers are bucketed as `{sheet}_Others`
- **Unit Handling:** FCST AMT/GP values are in thousands (千元) and are automatically scaled ×1,000 to match Shipping Record units

---

## 🏗️ System Architecture

### Directory Structure

```
SalesReportTool/
├── app/
│   ├── app.py              # Main Streamlit application
│   ├── charts.py           # Chart generation functions
│   ├── fcst_loader.py      # FCST Excel parser, blending, and customer mapping
│   ├── utils.py            # Data loading, classification, report helpers
│   ├── aliases.json        # Name alias mappings (customer, sales_person)
│   └── overrides.json      # Category overrides (auto-generated)
├── data/
│   ├── 2024/
│   │   └── sales_2024.xlsx
│   ├── 2025/
│   ├── 2026/
│   └── FCST/               # Latest FCST Excel file (auto-selected by mtime)
│       └── FCST_2026_wNN.xlsx
├── assets/
│   └── app.ico
├── launcher.py             # PyInstaller entry point
├── build.bat               # Build script (Windows)
├── SalesReportTool.spec    # PyInstaller configuration
├── requirements.txt        # Python dependencies
└── README.md              # This file
```

### Data Flow

```
[Shipping Record Excel Files]
    ↓
load_single_file()
    ├─ Parse & validate columns
    ├─ Normalize Customer Name + SALE_Person
    ├─ Apply alias maps
    ├─ Classify Category (direct → DES keyword → Others)
    └─ Return DataFrame
    ↓
[Merged DataFrame]
    ↓
apply_overrides()
    └─ Replace Category for matching rows
    ↓
[Streamlit Tabs]
    ├─ Performance Report
    ├─ Shipping Record Search
    └─ Company Dashboard
            ↓
        fcst_loader.get_fcst_for_dashboard()
            ├─ find_latest_fcst_file()
            ├─ _parse_sheet() × N sheets
            │   ├─ normalize_fcst_customer()  ← FCST_TO_CANONICAL + aliases.json
            │   └─ AMT/GP ×1000 scaling
            └─ pivot_table (long → wide)
            ↓
        blend_actual_fcst()
            ├─ month < current  → Actual
            └─ month ≥ current  → Forecast
            ↓
        agg_blended_monthly()
            └─ KPI cards + blended charts
```

### Key Modules

#### `fcst_loader.py`

| Function | Purpose |
|----------|---------|
| `find_latest_fcst_file(data_dir)` | Returns most-recently modified `.xlsx` in `data/FCST/` |
| `load_fcst(data_dir, sheet_name)` | Parses one or all FCST sheets into long-format DataFrame |
| `get_fcst_for_dashboard(data_dir, customer, sheet_name)` | Pivots to wide format for blending |
| `blend_actual_fcst(actual_df, fcst_df, current_month)` | Merges actual + forecast by month boundary |
| `agg_blended_monthly(blended_df)` | Company-level monthly totals with Source column |
| `agg_fcst_category_monthly(fcst_df)` | Monthly × Cat aggregation for category chart |
| `normalize_fcst_customer(fcst_name, sheet_name)` | Maps FCST names → canonical; fallback to `{sheet}_Others` |

#### `utils.py` Normalization Functions

```python
_normalize_name(name, upper=True)
  └─ Remove punctuation, compress whitespace, unify case

normalize_customer_name(name)
  └─ _normalize_name() + alias mapping

normalize_sales_person(name)
  └─ _normalize_name(upper=False) + alias mapping

_load_aliases(kind: "customer" | "sales_person")
  └─ Load from app/aliases.json
```

---

## 📊 Data Requirements

### Shipping Record — Required Columns (exact names)

| Column | Type | Notes |
|--------|------|-------|
| `Customer Name` | string | Will be normalized |
| `Ship Date` | date | Parsed with fault tolerance; NaT rows skipped |
| `QTY` | numeric | Quantity (used for CDR/Tablet categories) |
| `SALES Total AMT` | numeric | Revenue in TWD |
| `final GP(NTD,data from Financial Report)` | numeric | Gross Profit in TWD |
| `Part Number` | string | Product identifier |
| `Category` | string | Direct category or fallback destination |

### Shipping Record — Optional Columns

| Column | Type | Purpose |
|--------|------|---------|
| `DES` | string | Product description for keyword-based classification |
| `SALE_Person` | string | Sales rep name (will be normalized) |
| `Currency` | string | Used in Shipping Record Search |
| `UP` | numeric | Unit Price |
| `TP(USD)` | numeric | Total Price in USD |

### FCST File

- **Location:** `data/FCST/*.xlsx` (latest by modified time is auto-selected)
- **Sheets:** `Div.1&2_All`, `VT`, `Signify`
- **Structure:**
  - Row 1: (blank)
  - Row 2: Exchange rate + Month labels (merged cells, forward-filled)
  - Row 3: Column headers (`Budget` / `Forecast` / `PO` / `Shipped` / `Deviation`)
  - Row 4+: Data rows — each Customer × Model = 3 rows (QTY / AMT / GP in Detail column)
- **Units:** AMT and GP values are in thousands (千元); auto-scaled ×1,000 at parse time

### File Format (Shipping Record)

- **Format:** `.xlsx` (Excel 2007+)
- **Sheet:** Must contain `Actual` sheet
- **Location:** `data/{YEAR}/` (e.g., `data/2024/sales_2024.xlsx`)
- **Loading:** Latest modified `.xlsx` per year folder is auto-selected

### Data Exclusions

Customers in `EXCLUDED_CUSTOMERS` set are automatically filtered:
```python
EXCLUDED_CUSTOMERS = {"MITAC Computer(Kunshan) Co.,Ltd"}
```
Modify in `utils.py` to add/remove excluded customers.

---

## ⚙️ Configuration

### 1. Category Classification Rules

Edit `DES_RULES` in [app/utils.py](app/utils.py):

```python
DES_RULES = {
    "CDR ACC": ["cdr", "gemini", "evo", ...],
    "Tablet ACC": ["tablet", "prometheus", ...],
    "AI_SW": ["visionmax"],
}
```

When `Category` is empty/invalid, keywords from `DES` field are matched (substring, case-insensitive).

### 2. Name Aliases

Create or edit [app/aliases.json](app/aliases.json):

```json
{
  "customer": {
    "AZUGA INC": "AZUGA, Inc.",
    "KILLIAN": "Killian Chen"
  },
  "sales_person": {
    "KILLIAN": "Killian Chen",
    "JASON TAN": "Jason Tan"
  }
}
```

**Format:**
- After normalization (uppercase for customer, Title Case for sales), apply these mappings
- Use normalized form as the key
- Value is the canonical name to use

### 3. FCST Customer Mapping

Edit `FCST_TO_CANONICAL` in [app/fcst_loader.py](app/fcst_loader.py):

```python
FCST_TO_CANONICAL: dict[str, str] = {
    "AKAM":                   "AKAM Netherlands BV",
    "TTT/WFS/BMS(WFS)":       "Bridgestone Mobility Solutions B.V.",
    "Zonar-CDR":              "Zonar System Inc.",
    "Zonar-Tablet":           "Zonar System Inc.",
    # ... add more as needed
}
```

**Rules:**
- Keys are exact strings as they appear in the FCST Excel file
- Values are the canonical names used in the Shipping Record
- Multiple FCST keys can map to one canonical name (e.g., Zonar-CDR + Zonar-Tablet → Zonar System Inc.)
- Unmatched FCST customers are automatically bucketed as `"{sheet_name}_Others"` (e.g., `Div.1&2_All_Others`)
- `aliases.json` is checked as a secondary fallback before the Others bucket

### 4. Category Overrides

[app/overrides.json](app/overrides.json) is auto-generated when you adjust categories in the UI. Manual entries:

```json
{
  "[\"Company A\", \"PN-001\", \"2026-01\", \"Some Product Desc\"]": "Tablet"
}
```

Keys are JSON arrays: `[Customer, Part Number, Month (YYYY-MM), DES]`

---

## 🔧 Installation & Build

### Prerequisites

- **Windows 10+** (build script targets Windows)
- **Python 3.11+** (venv at `venv311/`)
- **pip** (for dependency installation)

### Dependencies

| Package | Version | Purpose |
|---------|---------|---------|
| streamlit | ≥1.32.0 | Web framework |
| pandas | ≥2.0.0 | Data manipulation |
| openpyxl | ≥3.1.0 | Excel reading (fallback) |
| pyinstaller | ≥6.0.0 | Executable packaging |
| python-calamine | ≥0.3.0 | Fast Excel reading (primary) |

See [requirements.txt](requirements.txt) for complete list.

### Build Process

Run `build.bat`:

```batch
@echo off
REM [1/5] Install dependencies
REM [2/5] Clean old output
REM [3/5] PyInstaller: launcher.py → SalesReportTool.exe
REM [4/5] Copy app\ and data\ folders
REM [5/5] Verify output
```

**Output:** `dist/SalesReportTool/`

**Distribution:** Zip the entire `dist/SalesReportTool/` folder for end users.

### Development Workflow

```bash
# After modifying app.py or fcst_loader.py
python -m streamlit run app/app.py

# After modifying utils.py or adding dependencies
pip install -r requirements.txt
python -m streamlit run app/app.py

# After modifying launcher.py
build.bat

# Small fixes (app.py / fcst_loader.py only)
# → Replace dist/SalesReportTool/app/<file> directly, no rebuild needed
```

---

## 🎯 Usage

### Command Line

```bash
# Development
streamlit run app/app.py

# With port override
streamlit run app/app.py --server.port 8502

# Frozen (from built executable)
SalesReportTool.exe
```

### UI Workflow

1. **Sidebar — Year Selection:**
   - Multi-select years to analyze
   - File auto-discovery from `data/{YEAR}/` folders

2. **Sidebar — Sales Person Filter (optional):**
   - Filter by sales rep (current year only)
   - Shows related customers

3. **Sidebar — FCST Sheet:**
   - `All Sheets` (default) — loads and merges `Div.1&2_All` + `VT` + `Signify`
   - Individual sheet options for isolated view

4. **Main Tabs:**

   **📄 Performance Report**
   - Summary metrics by month
   - Category breakdowns (stacked trends, donut charts)
   - Margin % (GP / Sales)
   - YoY comparison (if available)

   **🚚 Shipping Record Search**
   - Currency conversion
   - Unit price & quantity analysis
   - Requires: `Currency`, `UP`, `TP(USD)` columns

   **📊 Company Dashboard**
   - KPI cards (Total Sales, GP, GP%, QTY, Customers, Categories)
   - Full-Year Forecast row (Revenue, GP, GP%, QTY) — shown when current year is selected and FCST file exists
   - Monthly trend with Actual (solid) / Forecast (dashed) overlay
   - Category breakdown (donut + stacked bar + AI_SW trend + FCST category chart)
   - Top N customers (bar chart + table)
   - Customer detail drill-down with monthly revenue, category donut, QTY by category, part number detail

### Export Data

- Click **"📋 Download as CSV"** on report tables
- Charts: Right-click → Save as image (browser dependent)

---

## 🆘 Troubleshooting

### Issue: "Data folder not found"

**Cause:** Missing or empty `data/` directory  
**Fix:** Create `data/2024/`, `data/2025/`, etc., and place `.xlsx` files inside

### Issue: "No .xlsx files found in year folders"

**Cause:** Excel files in wrong location or wrong extension  
**Fix:** Ensure files are in `data/{YEAR}/*.xlsx` format

### Issue: "'DES' column not found; DES classification disabled"

**Cause:** Your Excel file doesn't have a `DES` column  
**Fix:** Either add the column or accept category fallback to "Others"

### Issue: "SALE_Person' column not found; Sales Person filter disabled"

**Cause:** Your Excel file doesn't have `SALE_Person` column  
**Fix:** Optional; add the column if you want sales rep filtering

### Issue: "X row(s) with invalid Ship Date skipped"

**Cause:** Unparseable dates in `Ship Date` column  
**Fix:** Ensure all dates are in recognizable format (e.g., YYYY-MM-DD, MM/DD/YYYY)

### Issue: Customer/Sales Person names not normalizing

**Cause:** Aliases not configured or names not found  
**Check:**
1. [app/aliases.json](app/aliases.json) has the mapping
2. Key is in normalized form: uppercase for customer, Title Case for sales
3. Reload browser/restart app after editing `aliases.json`

### Issue: FCST data not appearing in Company Dashboard

**Cause:** No FCST file found, or selected year is not the current calendar year  
**Check:**
1. A `.xlsx` file exists in `data/FCST/`
2. The current year (e.g., 2026) is selected in the Year Selection sidebar
3. Console output for `[fcst_loader]` warnings about unmapped customer names

### Issue: FCST Forecast numbers look 1,000× too small

**Cause:** Old FCST file before automatic unit scaling was added  
**Note:** v3.3+ automatically multiplies FCST AMT/GP by 1,000 at parse time — no manual action needed

### Issue: Build fails with "PyInstaller not found"

**Cause:** `pyinstaller` not installed  
**Fix:** Run `pip install -r requirements.txt` before `build.bat`

### Issue: "calamine" or "python_calamine" import errors

**Cause:** Dependency not installed  
**Fix:** Auto-handled; `openpyxl` fallback used if `calamine` unavailable

---

## 📝 Change Log

### v3.3 (April 2026)
- ✨ **FCST Integration in Company Dashboard**
  - New `fcst_loader.py` module: FCST Excel parser, blending engine, customer mapping
  - Blends Shipping Record actuals (past months) with FCST data (current + future months)
  - Full-Year Forecast KPI row (Revenue, GP, GP%, QTY)
  - Monthly trend charts show Actual (solid) / Forecast (dashed) with boundary marker
  - FCST category revenue chart added as 4th column in Category Analysis section
  - FCST Sheet sidebar selector: `All Sheets` / `Div.1&2_All` / `VT` / `Signify`
- 🗺️ **FCST Customer Name Mapping**
  - `FCST_TO_CANONICAL` dict (15 entries) maps FCST names → Performance Report canonical names
  - Three-stage lookup: `FCST_TO_CANONICAL` → `aliases.json` → `{sheet}_Others` bucket
  - AMT/GP auto-scaled ×1,000 (FCST stores values in thousands)

### v3.2 (April 2026)
- ✨ **Name Normalization System**
  - Auto-normalize Customer Name & Sales Person
  - Configurable alias maps via `aliases.json`
  - Handles punctuation, whitespace, case variations
- 🔧 **Build Process Improvements**
  - Added `python-calamine` to requirements (faster Excel reading)
  - Added `pyinstaller` to explicit dependencies
  - Archived deprecated `啟動程式.spec`

### v3.1 (Earlier)
- Category classification with DES keyword matching
- Shipping Record Search tab
- Company Dashboard with KPIs
- Override system for manual corrections
- Multi-year data support

---

## 👥 Developer Notes

### Key Decision Rationale

**FCST Blend Logic:**
- Past months (`< current_month`) always use Actual data — shipping records are authoritative
- Current month and future months (`>= current_month`) always use Forecast — actuals are incomplete or zero
- This ensures month-end numbers are never polluted by partial actuals

**FCST Customer Mapping:**
- FCST file uses short/internal names that differ from Shipping Record customer names
- `FCST_TO_CANONICAL` is maintained in `fcst_loader.py` as the single source of truth
- Unmatched customers land in `{sheet}_Others` buckets rather than being silently dropped
- Two FCST names can map to one canonical name (e.g., Zonar-CDR + Zonar-Tablet → Zonar System Inc.); the pivot's `aggfunc="sum"` handles the merge automatically

**`{sheet}_Others` Bucket:**
- Preserves all FCST revenue even for unmapped customers
- Bucketed per sheet (e.g., `Div.1&2_All_Others`, `VT_Others`, `Signify_Others`) so the source is traceable
- Appears as Forecast-only lines in blended charts (no Actual counterpart)

**Name Normalization:**
- Customer names often have punctuation/case variations across different sheets
- Sales person names may be entered inconsistently
- Normalizing at load-time ensures consistent filtering and aggregation

**Composite Keys (overrides.json):**
- Avoids index shift issues when Excel files are reorganized
- Part Number + Month + DES + Customer → unique product transaction row

**Lazy Alias Loading:**
- `_load_aliases()` reads JSON on every `normalize_*()` call
- `_load_fcst_customer_aliases()` in `fcst_loader.py` uses a module-level cache (`_ALIASES_CACHE`) — read once per process
- No Streamlit caching needed; file updates reflected after app restart

**Calamine + openpyxl Fallback:**
- `calamine` (Rust-based) is ~5–10x faster for large sheets
- `openpyxl` (pure Python) is more widely available, used as fallback

### Session State Management

- **`others_overrides`:** Loaded at startup; any UI edits saved to `overrides.json`
- **`sp_cust__*`:** Sales Person related customer checkboxes (UI-only)
- **`fcst_sheet`:** FCST Sheet radio selection (sidebar)
- **File Cache:** `@st.cache_data` on `load_single_file()` using `_rules_key()` for DES_RULES version

---

## 📄 License & Contact

*For internal use. Contact project maintainer for distribution.*

---

**Last Updated:** 2026-04-12
