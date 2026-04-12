"""
fcst_loader.py — FCST / PO pivot-style Excel parser.

Reads the weekly-updated FCST file from data/FCST/ and returns
a normalized long-format DataFrame for integration with Company Dashboard.

Expected file structure (Excel):
  Row 1: (blank)
  Row 2: Exchange rate (left) + Month labels (merged across 5 cols each)
  Row 3: BU, Region, Customer, ... Detail (left) + Budget/Forecast/PO/Shipped/Deviation
  Row 4+: Data rows (each Customer × Model = 3 rows: QTY, AMT, GP in Detail column)

Sheets: "Div.1&2_All", "VT"
"""

import json
import os
import re
import string
import pandas as pd
import numpy as np
from pathlib import Path
from typing import Optional

# ── Customer name mapping: FCST file → Performance Report canonical ──
# Keys are the exact strings as they appear in the FCST Excel file.
# Values are the normalized canonical names used in the Shipping Record.

FCST_TO_CANONICAL: dict[str, str] = {
    "AKAM":                      "AKAM Netherlands BV",
    "All-Connects":              "All-Connects NV",
    "Astrata SG":                "Astrata Group Pte Ltd",
    "TTT/WFS/BMS(AZUGA)":        "AZUGA INC.",
    "TTT/WFS/BMS(WFS)":          "Bridgestone Mobility Solutions B.V.",
    "CalAmp":                    "CalAmp Wireless Networks Corporation",
    "TeletracNavman-AU":         "Navman Wireless Australia Pty Ltd",
    "TeletracNavman-NZ":         "Navman Wireless New Zealand",
    "TeletracNavman-US  & MX":   "Teletrac Navman US Ltd.",
    "Geotab (SmarterAI)":        "Geotab Inc.",
    "Zonar-CDR":                 "Zonar System Inc.",
    "Zonar-Tablet":              "Zonar System Inc.",
    "Pedigree":                  "Pedigree Technologies LLC",
    "Texim":                     "Texim Europe B.V.",
    "Signify-EMS.21789.Signify Netherlands BV.Patek": "SIGNIFY NETHERLANDS BV",
}

# ── Configuration ─────────────────────────────────────────

FCST_FOLDER = "FCST"
FCST_SHEETS = ["Div.1&2_All", "VT", "Signify"]

COL_CUSTOMER = 2   # Column C
COL_CAT = 5        # Column F
COL_SALES = 6      # Column G
COL_DETAIL = 8     # Column I — contains MetricGroup: QTY / AMT / GP
LEFT_COLS_COUNT = 9 # A~I (data columns start from Column J)

# Header row indices (0-indexed)
ROW_MONTH = 1       # Excel Row 2: Exchange rate + Month labels
ROW_SUBCOL = 2      # Excel Row 3: Column headers + Budget/Forecast/PO/Shipped/Deviation
DATA_START_ROW = 3  # Excel Row 4: First data row

MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
QUARTERS = ["Q1", "Q2", "Q3", "Q4"]
ALL_PERIODS = MONTHS + QUARTERS + ["Annual"]
METRIC_GROUPS = ["QTY", "AMT", "GP"]
SUB_COLUMNS = ["Budget", "Forecast", "PO", "Shipped", "Deviation"]
MONTH_INDEX = {m: i + 1 for i, m in enumerate(MONTHS)}

# Module-level cache so aliases.json is only read once per process.
_ALIASES_CACHE: Optional[dict] = None


def _load_fcst_customer_aliases() -> dict:
    """Return the 'customer' section of app/aliases.json (cached)."""
    global _ALIASES_CACHE
    if _ALIASES_CACHE is None:
        try:
            path = Path(__file__).resolve().parent / "aliases.json"
            with open(path, encoding="utf-8") as f:
                _ALIASES_CACHE = json.load(f).get("customer", {})
        except Exception:
            _ALIASES_CACHE = {}
    return _ALIASES_CACHE


def _normalize_fcst_name(name: str) -> str:
    """Strip punctuation, compress whitespace, uppercase — mirrors utils._normalize_name."""
    name = name.translate(str.maketrans("", "", string.punctuation))
    name = re.sub(r"\s+", " ", name.strip())
    return name.upper()


def normalize_fcst_customer(fcst_name: str, sheet_name: str) -> str:
    """Map a raw FCST customer name to the Performance Report canonical name.

    Lookup order:
      1. FCST_TO_CANONICAL  — exact match (strip only), then case-insensitive
      2. aliases.json       — same normalization as utils.normalize_customer_name
      3. Fallback           — ``"{sheet_name}_Others"``
    """
    name = str(fcst_name).strip()

    # 1. FCST_TO_CANONICAL — exact
    if name in FCST_TO_CANONICAL:
        return FCST_TO_CANONICAL[name]
    # 1b. case-insensitive fallback
    name_lower = name.lower()
    for k, v in FCST_TO_CANONICAL.items():
        if k.lower() == name_lower:
            return v

    # 2. aliases.json
    aliases = _load_fcst_customer_aliases()
    norm = _normalize_fcst_name(name)
    if norm in aliases:
        return aliases[norm]

    # 3. Unknown → Others bucket keyed by sheet
    print(f"[fcst_loader] No mapping for '{name}' (sheet={sheet_name}) "
          f"→ {sheet_name}_Others")
    return f"{sheet_name}_Others"


# ── Public API ────────────────────────────────────────────

def find_latest_fcst_file(data_dir: str) -> Optional[str]:
    fcst_dir = os.path.join(data_dir, FCST_FOLDER)
    if not os.path.isdir(fcst_dir):
        return None
    xlsx_files = [
        os.path.join(fcst_dir, f)
        for f in os.listdir(fcst_dir)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]
    if not xlsx_files:
        return None
    return max(xlsx_files, key=os.path.getmtime)


def load_fcst(data_dir: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    filepath = find_latest_fcst_file(data_dir)
    if filepath is None:
        return pd.DataFrame()
    sheets = [sheet_name] if sheet_name else FCST_SHEETS
    frames = []
    for sn in sheets:
        try:
            df = _parse_sheet(filepath, sn)
            if not df.empty:
                df["Sheet"] = sn
                frames.append(df)
        except Exception as e:
            print(f"[fcst_loader] Warning: Failed to parse sheet '{sn}': {e}")
            continue
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)


def get_fcst_for_dashboard(
    data_dir: str,
    customer: Optional[str] = None,
    sheet_name: str = "Div.1&2_All",
) -> pd.DataFrame:
    df = load_fcst(data_dir, sheet_name=sheet_name)
    if df.empty:
        return df
    df = df[df["Period"].isin(MONTHS)].copy()
    if customer:
        df = df[df["Customer"].str.upper() == customer.upper()]
    pivot = df.pivot_table(
        index=["Customer", "Cat", "Sales", "Period", "MonthIndex"],
        columns="MetricGroup",
        values=["Budget", "Forecast", "PO", "Shipped"],
        aggfunc="sum",
        fill_value=0,
    ).reset_index()
    pivot.columns = [
        f"{col[1]}_{col[0]}" if col[1] else col[0]
        for col in pivot.columns
    ]
    return pivot


def blend_actual_fcst(
    actual_df: pd.DataFrame,
    fcst_df: pd.DataFrame,
    current_month: int,
    qty_col: str = "QTY",
    amt_col: str = "SALES Total AMT",
    gp_col: str = "final GP(NTD,data from Financial Report)",
) -> pd.DataFrame:
    records = []
    all_customers = set()
    if not actual_df.empty:
        all_customers.update(actual_df["Customer"].unique())
    if not fcst_df.empty:
        all_customers.update(fcst_df["Customer"].unique())
    for customer in all_customers:
        for month_name, month_idx in MONTH_INDEX.items():
            act = _get_actual_values(actual_df, customer, month_idx,
                                     qty_col, amt_col, gp_col)
            fct = _get_fcst_values(fcst_df, customer, month_idx)
            if month_idx < current_month:
                source = "Actual"
                qty, amt, gp = act["qty"], act["amt"], act["gp"]
            else:  # >= current_month: always Forecast
                source = "Forecast"
                qty, amt, gp = fct["qty"], fct["amt"], fct["gp"]
            records.append({
                "Customer": customer, "Period": month_name,
                "MonthIndex": month_idx, "QTY": qty,
                "AMT": amt, "GP": gp, "Source": source,
            })
    return pd.DataFrame(records)


def agg_blended_monthly(blended_df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate per-customer blended data to company-level monthly totals.

    Input: output of blend_actual_fcst().
    Output columns: Period, MonthIndex, Source, Revenue, GP, QTY, GP%.
    """
    if blended_df.empty:
        return pd.DataFrame()
    agg = (
        blended_df
        .groupby(["Period", "MonthIndex", "Source"], sort=False)
        .agg(Revenue=("AMT", "sum"), GP=("GP", "sum"), QTY=("QTY", "sum"))
        .reset_index()
    )
    agg["GP%"] = agg.apply(
        lambda r: r["GP"] / r["Revenue"] * 100 if r["Revenue"] else 0.0, axis=1
    )
    return agg.sort_values("MonthIndex").reset_index(drop=True)


def agg_fcst_category_monthly(fcst_df: pd.DataFrame) -> pd.DataFrame:
    """Aggregate FCST data by month and Cat for category breakdown chart.

    Input: output of get_fcst_for_dashboard().
    Output columns: Period, MonthIndex, Cat, Revenue.
    """
    if fcst_df.empty or "AMT_Forecast" not in fcst_df.columns:
        return pd.DataFrame()
    agg = (
        fcst_df
        .groupby(["Period", "MonthIndex", "Cat"], sort=False)
        .agg(Revenue=("AMT_Forecast", "sum"))
        .reset_index()
    )
    return agg.sort_values("MonthIndex").reset_index(drop=True)


# ── Internal helpers ──────────────────────────────────────

def _parse_sheet(filepath: str, sheet_name: str) -> pd.DataFrame:
    # Read header rows (Row 1~3, 0-indexed 0~2)
    header_df = pd.read_excel(
        filepath, sheet_name=sheet_name,
        header=None, nrows=DATA_START_ROW, dtype=str,
    )
    # Read data rows (from Excel Row 4 onward)
    data_df = pd.read_excel(
        filepath, sheet_name=sheet_name,
        header=None, skiprows=DATA_START_ROW,
    )
    if data_df.empty:
        return pd.DataFrame()
    exchange_rate = _extract_exchange_rate(header_df)
    num_data_cols = data_df.shape[1] - LEFT_COLS_COUNT
    if num_data_cols <= 0:
        return pd.DataFrame()

    # Build column mapping from header rows
    # ROW_MONTH (row 1): month labels — forward-fill merged cells
    month_row = header_df.iloc[ROW_MONTH, LEFT_COLS_COUNT:].ffill().values
    # ROW_SUBCOL (row 2): Budget/Forecast/PO/Shipped/Deviation
    sub_col_row = header_df.iloc[ROW_SUBCOL, LEFT_COLS_COUNT:].values

    # col_map: col_offset → (period, sub_col)
    col_map = []
    for i in range(min(num_data_cols, len(month_row))):
        period = _normalize_period(str(month_row[i]).strip()) if pd.notna(month_row[i]) else ""
        sub = _normalize_sub_col(str(sub_col_row[i]).strip()) if pd.notna(sub_col_row[i]) else ""
        col_map.append((period, sub))

    # Parse data rows — MetricGroup comes from Detail column (H)
    records = []
    for _, row in data_df.iterrows():
        customer = row.iloc[COL_CUSTOMER]
        if pd.isna(customer) or str(customer).strip() == "":
            continue
        # Normalize to Performance Report canonical name (or sheet_Others bucket)
        customer_name = normalize_fcst_customer(str(customer).strip(), sheet_name)
        cat = row.iloc[COL_CAT] if pd.notna(row.iloc[COL_CAT]) else ""
        sales = row.iloc[COL_SALES] if pd.notna(row.iloc[COL_SALES]) else ""

        # Read MetricGroup from Detail column (QTY / AMT / GP)
        detail_val = row.iloc[COL_DETAIL] if pd.notna(row.iloc[COL_DETAIL]) else ""
        metric = _normalize_metric(str(detail_val).strip())
        if not metric:
            continue  # skip rows without a valid metric type

        for col_offset, (period, sub) in enumerate(col_map):
            if not period or not sub:
                continue
            abs_col = LEFT_COLS_COUNT + col_offset
            if abs_col >= len(row):
                break
            value = row.iloc[abs_col]
            value = 0 if pd.isna(value) else value
            # FCST file stores AMT/GP in thousands (千元) — scale to match Shipping Record
            if metric in ("AMT", "GP"):
                value = value * 1000
            records.append({
                "Customer": customer_name,
                "Cat": str(cat).strip(),
                "Sales": str(sales).strip(),
                "Period": period,
                "MonthIndex": MONTH_INDEX.get(period, 0),
                "MetricGroup": metric,
                "SubColumn": sub,
                "Value": value,
            })
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    df = df.pivot_table(
        index=["Customer", "Cat", "Sales", "Period", "MonthIndex", "MetricGroup"],
        columns="SubColumn", values="Value", aggfunc="sum", fill_value=0,
    ).reset_index()
    df.columns.name = None
    for col in SUB_COLUMNS:
        if col not in df.columns:
            df[col] = 0
    df["ExchangeRate"] = exchange_rate
    return df


def _extract_exchange_rate(header_df: pd.DataFrame) -> float:
    # Exchange rate is in Excel Row 2 (0-indexed row 1)
    for col_idx in range(min(5, header_df.shape[1])):
        val = str(header_df.iloc[ROW_MONTH, col_idx])  # Row 2 = index 1
        if "exchange" in val.lower() or "匯率" in val:
            match = re.search(r"[\d.]+", val.split(":")[-1] if ":" in val else val)
            if match:
                return float(match.group())
    return 29.3


def _normalize_period(label: str) -> str:
    mapping = {
        "jan": "Jan", "jan.": "Jan", "january": "Jan",
        "feb": "Feb", "feb.": "Feb", "february": "Feb",
        "mar": "Mar", "mar.": "Mar", "march": "Mar",
        "apr": "Apr", "apr.": "Apr", "april": "Apr",
        "may": "May", "may.": "May",
        "jun": "Jun", "jun.": "Jun", "june": "Jun",
        "jul": "Jul", "jul.": "Jul", "july": "Jul",
        "aug": "Aug", "aug.": "Aug", "august": "Aug",
        "sep": "Sep", "sep.": "Sep", "september": "Sep",
        "oct": "Oct", "oct.": "Oct", "october": "Oct",
        "nov": "Nov", "nov.": "Nov", "november": "Nov",
        "dec": "Dec", "dec.": "Dec", "december": "Dec",
        "q1": "Q1", "q2": "Q2", "q3": "Q3", "q4": "Q4",
        "annual": "Annual", "total": "Annual", "fy": "Annual",
    }
    return mapping.get(label.lower(), label if label in ALL_PERIODS else "")


def _normalize_sub_col(label: str) -> str:
    mapping = {
        "budget": "Budget", "forecast": "Forecast", "fcst": "Forecast",
        "po": "PO", "shipped": "Shipped", "ship": "Shipped",
        "deviation": "Deviation", "dev": "Deviation", "var": "Deviation",
    }
    return mapping.get(label.lower(), "")


def _normalize_metric(label: str) -> str:
    mapping = {
        "qty": "QTY", "quantity": "QTY",
        "amt": "AMT", "amount": "AMT", "revenue": "AMT",
        "gp": "GP", "gross profit": "GP",
    }
    return mapping.get(label.lower(), "")


def _get_actual_values(df, customer, month_idx, qty_col, amt_col, gp_col):
    if df.empty:
        return {"qty": 0, "amt": 0, "gp": 0}
    mask = (df["Customer"] == customer) & (df["Month"] == month_idx)
    subset = df[mask]
    if subset.empty:
        return {"qty": 0, "amt": 0, "gp": 0}
    return {
        "qty": subset[qty_col].sum() if qty_col in subset.columns else 0,
        "amt": subset[amt_col].sum() if amt_col in subset.columns else 0,
        "gp": subset[gp_col].sum() if gp_col in subset.columns else 0,
    }


def _get_fcst_values(df, customer, month_idx):
    if df.empty:
        return {"qty": 0, "amt": 0, "gp": 0}
    mask = (df["Customer"] == customer) & (df["MonthIndex"] == month_idx)
    subset = df[mask]
    if subset.empty:
        return {"qty": 0, "amt": 0, "gp": 0}
    return {
        "qty": subset.get("QTY_Forecast", pd.Series([0])).sum(),
        "amt": subset.get("AMT_Forecast", pd.Series([0])).sum(),
        "gp": subset.get("GP_Forecast", pd.Series([0])).sum(),
    }