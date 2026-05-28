# Sales Report Tool

Streamlit-based sales performance analysis tool with automatic data classification and forecast integration.

**Version:** 3.6 | **Build Date:** May 2026

---

## Quick Start

**End Users:** Double-click `SalesReportTool.exe`

**Developers:**
```bash
./venv311/Scripts/Activate.ps1
pip install -r requirements.txt
streamlit run app/app.py   # dev server
build.bat                  # build exe
```

---

## Directory Structure

```
SalesReportTool/
├── app/
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
├── launcher.py
├── build.bat
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

---

## Change Log

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

*For internal use. Last Updated: 2026-05-28*
