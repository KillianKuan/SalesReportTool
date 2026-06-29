"""utils.py — Data loading, cleaning, classification, and report helpers."""

import functools
import json
import os
import re
import string
from pathlib import Path

import pandas as pd
import streamlit as st

# ── Constants ─────────────────────────────────────────────────────
REQUIRED_COLS = [
    "Customer Name", "Ship Date", "QTY",
    "SALES Total AMT", "final GP(NTD,data from Financial Report)",
    "Part Number", "Category",
]
SHIPPING_COLS = ["Currency", "UP", "TP(USD)"]
VALID_CATEGORIES = {"Tablet", "CDR", "Tablet ACC", "CDR ACC", "AI_SW", "Signify"}
_VALID_CAT_MAP = {" ".join(c.upper().split()): c for c in VALID_CATEGORIES}
GP_COL = "final GP(NTD,data from Financial Report)"
CAT_ORDER = ["CDR", "CDR ACC", "Tablet", "Tablet ACC", "AI_SW", "Signify", "Others"]
EXCLUDED_CUSTOMERS = {"MITAC COMPUTERKUNSHAN COLTD"}
QTY_CATEGORIES = {"CDR", "Tablet"}
DES_RULES = {
    "CDR ACC":    ["cdr", "gemini", "evo", "sprint", "sd card", "panic button",
                   "iosix", "uvc camera", "k220", "k245", "k265",
                   "smart link dongle", "safetycam"],
    "Tablet ACC": ["tablet", "prometheus", "chiron", "hera", "phaeton", "surfing pro",
                   "cradle", "f840", "ulmo", "fleet cable"],
    "AI_SW":      ["visionmax"],
    "Signify":    ["signify"],
}

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR.parent / "data"
HISTORICAL_DIR = DATA_DIR / "Over the Years"
CURRENT_YEAR_DIR = DATA_DIR / "Current Year"
HISTORICAL_CSV = HISTORICAL_DIR / "historical.csv"
HISTORICAL_PARQUET = HISTORICAL_DIR / "historical.parquet"
OVERRIDES_FILE = str(APP_DIR / "overrides.json")

# Translation table used to strip punctuation during name normalization.
_PUNCT_TABLE = str.maketrans("", "", string.punctuation)


# ── Data folder scanning ─────────────────────────────────────────
def scan_current_year_folder() -> Path | None:
    """Return the most-recently-modified .xlsx in Current Year/, or None."""
    if not CURRENT_YEAR_DIR.exists():
        return None
    xlsx_files = list(CURRENT_YEAR_DIR.glob("*.xlsx"))
    if not xlsx_files:
        return None
    return max(xlsx_files, key=lambda f: f.stat().st_mtime)


def get_latest_xlsx(year_dir: Path) -> Path | None:
    """Return the most-recently-modified .xlsx file in *year_dir*, or None."""
    xlsx_files = list(year_dir.glob("*.xlsx"))
    if not xlsx_files:
        return None
    return max(xlsx_files, key=lambda f: f.stat().st_mtime)


# ── Overrides ─────────────────────────────────────────────────────
def save_overrides(ov):
    try:
        with open(OVERRIDES_FILE, "w", encoding="utf-8") as f:
            json.dump([[list(k), v] for k, v in ov.items()], f,
                      ensure_ascii=False, indent=2)
    except Exception:
        pass


def load_overrides():
    try:
        if os.path.exists(OVERRIDES_FILE):
            with open(OVERRIDES_FILE, encoding="utf-8") as f:
                return {tuple(row[0]): row[1] for row in json.load(f)}
    except Exception:
        pass
    return {}


# ── Name normalization ───────────────────────────────────────────
def _normalize_name(name, upper=True):
    """Remove punctuation, compress whitespace, unify case."""
    if not isinstance(name, str):
        return ""
    # Remove punctuation
    name = name.translate(_PUNCT_TABLE)
    # Compress whitespace
    name = re.sub(r'\s+', ' ', name.strip())
    # Unify case
    return name.upper() if upper else name.lower()


@functools.lru_cache(maxsize=None)
def _load_aliases(kind):
    """Load aliases from app/aliases.json.

    Cached per process: the file is read and parsed only once per ``kind``
    instead of on every call (previously re-opened for every row).
    """
    try:
        with open(APP_DIR / "aliases.json", encoding="utf-8") as f:
            data = json.load(f)
        return data.get(kind, {})
    except Exception:
        return {}


def normalize_customer_name(name):
    """Normalize customer name with alias mapping."""
    normalized = _normalize_name(name, upper=True)
    aliases = _load_aliases("customer")
    return aliases.get(normalized, normalized)


def normalize_sales_person(name):
    """Normalize sales person name with alias mapping."""
    normalized = _normalize_name(name, upper=False)
    aliases = _load_aliases("sales_person")
    return aliases.get(normalized, normalized)


def _normalize_series(names, upper=True):
    """Vectorized equivalent of _normalize_name() over a pandas Series."""
    out = (
        names.astype(str)
        .str.translate(_PUNCT_TABLE)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    return out.str.upper() if upper else out.str.lower()


def normalize_customer_series(names):
    """Vectorized customer-name normalization with alias mapping.

    Equivalent to ``names.apply(normalize_customer_name)`` but avoids per-row
    work: aliases are loaded once and applied with a dict-based ``.map()``.
    Unmatched names keep their normalized value (same fallback as before).
    """
    normalized = _normalize_series(names, upper=True)
    aliases = _load_aliases("customer")
    return normalized.map(aliases).fillna(normalized)


def normalize_sales_person_series(names):
    """Vectorized sales-person normalization with alias mapping.

    Equivalent to ``names.apply(normalize_sales_person)``.
    """
    normalized = _normalize_series(names, upper=False)
    aliases = _load_aliases("sales_person")
    return normalized.map(aliases).fillna(normalized)


# ── Data loading (cached) ────────────────────────────────────────
def _rules_key():
    """Convert DES_RULES to a hashable tuple for cache busting."""
    return tuple((k, tuple(v)) for k, v in DES_RULES.items())


@st.cache_data
def load_single_file(file_path: str, rules_key):
    """Load and clean a single .xlsx file.
    Returns (df, nat_count, err, ambiguous, has_des, has_shipping).
    """
    try:
        xl = pd.ExcelFile(file_path, engine="calamine")
    except ImportError:
        xl = pd.ExcelFile(file_path)
    except Exception as e:
        return None, 0, f"Cannot read {file_path}: {e}", [], False, False
    if "Actual" not in xl.sheet_names:
        return (None, 0,
                f"'{Path(file_path).name}': 'Actual' sheet not found. "
                f"Available: {xl.sheet_names}", [], False, False)
    raw = xl.parse("Actual")
    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        return (None, 0,
                f"'{Path(file_path).name}': Missing columns: {missing}",
                [], False, False)

    has_des = "DES" in raw.columns
    has_sp = "SALE_Person" in raw.columns
    has_shipping = all(c in raw.columns for c in SHIPPING_COLS)

    use_cols = (REQUIRED_COLS
                + (["DES"] if has_des else [])
                + (["SALE_Person"] if has_sp else [])
                + (SHIPPING_COLS if has_shipping else []))
    df = raw[use_cols].copy()
    df["Ship Date"] = pd.to_datetime(
        df["Ship Date"].astype(str).str.strip(), errors="coerce"
    )
    nat_count = int(df["Ship Date"].isna().sum())
    df = df.dropna(subset=["Ship Date"])
    df["Month"] = df["Ship Date"].dt.strftime("%Y-%m")
    df["Category"] = df["Category"].astype(str).str.strip()
    if has_des:
        df["DES"] = df["DES"].astype(str).str.strip()
    if has_sp:
        df["SALE_Person"] = df["SALE_Person"].astype(str).str.strip()

    # ── Vectorized category classification ──
    ambiguous = []
    orig_cat = df["Category"].copy()

    # Customer-name-based override (runs before Category/DES logic)
    CUSTOMER_CATEGORY_MAP = {
        "SIGNIFY": "Signify",
    }
    cust_upper = df["Customer Name"].str.strip().str.upper()
    customer_aliases = _load_aliases("customer")
    cust_upper = cust_upper.map(customer_aliases).fillna(cust_upper)
    customer_cat = cust_upper.map(CUSTOMER_CATEGORY_MAP)

    cat_upper = df["Category"].str.upper().str.split().str.join(" ")
    df["Category"] = cat_upper.map(_VALID_CAT_MAP)

    needs_des = df["Category"].isna()
    if has_des and needs_des.any():
        des_lower = df.loc[needs_des, "DES"].str.lower()
        match_cats = {}
        for cat_name, keywords in DES_RULES.items():
            pattern = "|".join(re.escape(k) for k in keywords)
            match_cats[cat_name] = des_lower.str.contains(pattern, na=False)

        match_count = sum(m.astype(int) for m in match_cats.values())
        ambiguous_mask = match_count > 1
        if ambiguous_mask.any():
            for idx in ambiguous_mask[ambiguous_mask].index:
                matched_names = [c for c, m in match_cats.items() if m[idx]]
                ambiguous.append({
                    "Part Number": df.at[idx, "Part Number"],
                    "DES": df.at[idx, "DES"],
                    "Original Category": orig_cat[idx],
                    "Matched": " / ".join(matched_names),
                    "Assigned": matched_names[0],
                })

        for cat_name, matched in match_cats.items():
            still_na = df.loc[needs_des, "Category"].isna()
            to_fill = still_na & matched
            df.loc[to_fill[to_fill].index, "Category"] = cat_name

    # Apply customer-name override (takes priority over Category/DES results)
    df.loc[customer_cat.notna(), "Category"] = customer_cat[customer_cat.notna()]

    df["Category"] = df["Category"].fillna("Others")
    for col in ["QTY", "SALES Total AMT", GP_COL]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if has_shipping:
        for col in ["UP", "TP(USD)"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        df["Currency"] = df["Currency"].astype(str).str.strip()
    df["Customer Name"] = df["Customer Name"].astype(str).str.strip()
    df["Customer Name"] = normalize_customer_series(df["Customer Name"])
    if has_sp:
        df["SALE_Person"] = df["SALE_Person"].astype(str).str.strip()
        df["SALE_Person"] = normalize_sales_person_series(df["SALE_Person"])
    df = df[~df["Customer Name"].isin(["nan", "NaN", ""])]
    df["Part Number"] = (
        df["Part Number"].astype(str).str.strip()
        .replace({"None": "", "nan": "", "NaN": ""})
    )
    return df, nat_count, None, ambiguous, has_des, has_shipping


def _try_read_csv_with_encodings(file_path: str, encodings):
    last_exc = None
    for encoding in encodings:
        try:
            return pd.read_csv(file_path, encoding=encoding, low_memory=False)
        except UnicodeDecodeError as e:
            last_exc = e
            continue
        except Exception as e:
            last_exc = e
            break
    raise last_exc if last_exc is not None else ValueError(
        f"Unable to read CSV file: {file_path}"
    )


@st.cache_data
def load_historical_csv(file_path: str, rules_key):
    """Load data/Over the Years/historical.csv (or its Parquet sibling).
    Applies the same cleaning pipeline as load_single_file().
    Returns (df, nat_count, err, ambiguous, has_des, has_shipping).

    For faster cold-start loading, a sibling ``historical.parquet`` is used
    when present; otherwise the multi-encoding CSV reader is used.
    """
    parquet_path = Path(file_path).with_suffix(".parquet")
    try:
        if parquet_path.exists():
            raw = pd.read_parquet(parquet_path)
        else:
            raw = _try_read_csv_with_encodings(
                file_path,
                ["utf-8-sig", "utf-8", "cp950", "cp936", "latin1"],
            )
    except Exception as e:
        return None, 0, f"Cannot read historical data: {e}", [], False, False

    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        return (None, 0,
                f"historical.csv: Missing columns: {missing}",
                [], False, False)

    has_des = "DES" in raw.columns
    has_sp = "SALE_Person" in raw.columns
    has_shipping = all(c in raw.columns for c in SHIPPING_COLS)

    use_cols = (REQUIRED_COLS
                + (["DES"] if has_des else [])
                + (["SALE_Person"] if has_sp else [])
                + (SHIPPING_COLS if has_shipping else []))
    df = raw[use_cols].copy()
    df["Ship Date"] = pd.to_datetime(
        df["Ship Date"].astype(str).str.strip(), errors="coerce"
    )
    nat_count = int(df["Ship Date"].isna().sum())
    df = df.dropna(subset=["Ship Date"])
    df["Month"] = df["Ship Date"].dt.strftime("%Y-%m")
    df["Category"] = df["Category"].astype(str).str.strip()
    if has_des:
        df["DES"] = df["DES"].astype(str).str.strip()
    if has_sp:
        df["SALE_Person"] = df["SALE_Person"].astype(str).str.strip()

    ambiguous = []
    orig_cat = df["Category"].copy()

    CUSTOMER_CATEGORY_MAP = {
        "SIGNIFY": "Signify",
    }
    cust_upper = df["Customer Name"].str.strip().str.upper()
    customer_aliases = _load_aliases("customer")
    cust_upper = cust_upper.map(customer_aliases).fillna(cust_upper)
    customer_cat = cust_upper.map(CUSTOMER_CATEGORY_MAP)

    cat_upper = df["Category"].str.upper().str.split().str.join(" ")
    df["Category"] = cat_upper.map(_VALID_CAT_MAP)

    needs_des = df["Category"].isna()
    if has_des and needs_des.any():
        des_lower = df.loc[needs_des, "DES"].str.lower()
        match_cats = {}
        for cat_name, keywords in DES_RULES.items():
            pattern = "|".join(re.escape(k) for k in keywords)
            match_cats[cat_name] = des_lower.str.contains(pattern, na=False)

        match_count = sum(m.astype(int) for m in match_cats.values())
        ambiguous_mask = match_count > 1
        if ambiguous_mask.any():
            for idx in ambiguous_mask[ambiguous_mask].index:
                matched_names = [c for c, m in match_cats.items() if m[idx]]
                ambiguous.append({
                    "Part Number": df.at[idx, "Part Number"],
                    "DES": df.at[idx, "DES"],
                    "Original Category": orig_cat[idx],
                    "Matched": " / ".join(matched_names),
                    "Assigned": matched_names[0],
                })

        for cat_name, matched in match_cats.items():
            still_na = df.loc[needs_des, "Category"].isna()
            to_fill = still_na & matched
            df.loc[to_fill[to_fill].index, "Category"] = cat_name

    df.loc[customer_cat.notna(), "Category"] = customer_cat[customer_cat.notna()]

    df["Category"] = df["Category"].fillna("Others")
    for col in ["QTY", "SALES Total AMT", GP_COL]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    if has_shipping:
        for col in ["UP", "TP(USD)"]:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
        df["Currency"] = df["Currency"].astype(str).str.strip()
    df["Customer Name"] = df["Customer Name"].astype(str).str.strip()
    df["Customer Name"] = normalize_customer_series(df["Customer Name"])
    if has_sp:
        df["SALE_Person"] = df["SALE_Person"].astype(str).str.strip()
        df["SALE_Person"] = normalize_sales_person_series(df["SALE_Person"])
    df = df[~df["Customer Name"].isin(["nan", "NaN", ""])]
    df["Part Number"] = (
        df["Part Number"].astype(str).str.strip()
        .replace({"None": "", "nan": "", "NaN": ""})
    )
    return df, nat_count, None, ambiguous, has_des, has_shipping


# ── Report helpers ────────────────────────────────────────────────
def build_summary(base, qty_only):
    src = base[base["Category"].isin({"Tablet", "CDR"})] if qty_only else base
    agg = base.groupby("Month", sort=True).agg(
        **{"SALES Total AMT": ("SALES Total AMT", "sum"),
           "final GP(NTD)": (GP_COL, "sum")}
    ).reset_index()
    qty = (src.groupby("Month")["QTY"].sum()
           .reset_index().rename(columns={"QTY": "QTY (All)"}))
    return agg.merge(qty, on="Month", how="left").fillna({"QTY (All)": 0})


def build_bycat(base, qty_only, merge_cdr, merge_tab):
    cat_df = base.copy()
    orig = cat_df["Category"].copy()
    if merge_cdr:
        cat_df["Category"] = cat_df["Category"].replace("CDR ACC", "CDR")
    if merge_tab:
        cat_df["Category"] = cat_df["Category"].replace("Tablet ACC", "Tablet")
    agg = cat_df.groupby(["Month", "Category"], sort=True).agg(
        **{"SALES Total AMT": ("SALES Total AMT", "sum"),
           "final GP(NTD)": (GP_COL, "sum")}
    ).reset_index()
    mask = (orig.isin({"Tablet", "CDR"}) if qty_only
            else pd.Series(True, index=cat_df.index))
    qty = (cat_df[mask].groupby(["Month", "Category"])["QTY"].sum()
           .reset_index().rename(columns={"QTY": "QTY (All)"}))
    long = agg.merge(qty, on=["Month", "Category"], how="left")
    long["QTY (All)"] = long["QTY (All)"].fillna(0)
    return long


def to_wide_summary(long_df):
    metrics = ["QTY (All)", "SALES Total AMT", "final GP(NTD)"]
    m = long_df.melt(id_vars=["Month"], value_vars=metrics,
                     var_name="Metric", value_name="Value")
    p = (m.pivot_table(index="Metric", columns="Month",
                       values="Value", aggfunc="sum").reindex(metrics))
    p.columns.name = None
    month_cols = list(p.columns)
    years = sorted(set(c[:4] for c in month_cols))
    if len(years) > 1:
        for yr in years:
            yr_cols = [c for c in month_cols if c.startswith(yr)]
            p[f"{yr} Total"] = p[yr_cols].sum(axis=1)
    p["Total"] = p[month_cols].sum(axis=1)
    result = p.reset_index()
    val_cols = [c for c in result.columns if c != "Metric"]
    s_vals = result.loc[result["Metric"] == "SALES Total AMT", val_cols].values[0]
    g_vals = result.loc[result["Metric"] == "final GP(NTD)", val_cols].values[0]
    gp_row = pd.DataFrame(
        [["GP%"] + [f"{g/s*100:.1f}%" if pd.notna(s) and s != 0
                     else "-" for g, s in zip(g_vals, s_vals)]],
        columns=["Metric"] + val_cols,
    )
    return pd.concat([result, gp_row], ignore_index=True)


def to_wide_one_cat(long_df, cat, all_months):
    metrics = ["QTY (All)", "SALES Total AMT", "final GP(NTD)"]
    sub = long_df[long_df["Category"] == cat]
    m = sub.melt(id_vars=["Month"], value_vars=metrics,
                 var_name="Metric", value_name="Value")
    p = (m.pivot_table(index="Metric", columns="Month",
                       values="Value", aggfunc="sum").reindex(metrics))
    p = p.reindex(columns=all_months, fill_value=0).fillna(0)
    p.columns.name = None
    result = p.reset_index()
    val_cols = [c for c in result.columns if c != "Metric"]
    s_vals = result.loc[result["Metric"] == "SALES Total AMT", val_cols].values[0]
    g_vals = result.loc[result["Metric"] == "final GP(NTD)", val_cols].values[0]
    gp_row = pd.DataFrame(
        [["GP%"] + [f"{g/s*100:.1f}%" if pd.notna(s) and s != 0
                     else "-" for g, s in zip(g_vals, s_vals)]],
        columns=["Metric"] + val_cols,
    )
    return pd.concat([result, gp_row], ignore_index=True)


def to_fcst_wide_summary(fcst_df, selected_customers, month_cols):
    """Convert FCST DataFrame to wide summary format for appending to report.

    fcst_df: output of get_fcst_for_dashboard()
    selected_customers: list of customer names to filter
    month_cols: list of month column names from actual summary (e.g., ['2024-01', '2024-02', ...])
    """
    if fcst_df.empty:
        return pd.DataFrame()

    # Filter for selected customers (case-insensitive)
    fcst_filtered = fcst_df[fcst_df["Customer"].str.upper().isin([c.upper() for c in selected_customers])].copy()
    if fcst_filtered.empty:
        return pd.DataFrame()

    # Aggregate by Period (month name like 'Jan', 'Feb')
    agg = fcst_filtered.groupby("Period").agg({
        "QTY_Forecast": "sum",
        "AMT_Forecast": "sum",
        "GP_Forecast": "sum"
    }).reset_index()

    # Metric mapping
    metric_to_col = {
        "FCST QTY": "QTY_Forecast",
        "FCST AMT(TWD)": "AMT_Forecast",
        "FCST GP(TWD)": "GP_Forecast"
    }

    # Create the wide DataFrame with same columns as wide_summary
    data = []
    for metric in ["FCST QTY", "FCST AMT(TWD)", "FCST GP(TWD)"]:
        row = {"Metric": metric}
        col_name = metric_to_col[metric]
        for col in month_cols:
            if col == "Total":
                # Calculate total across all months
                total = 0
                for m_col in month_cols:
                    if m_col != "Total" and not m_col.endswith(" Total"):
                        try:
                            dt = pd.to_datetime(m_col)
                            period = dt.strftime("%b")
                            if period in agg["Period"].values:
                                val = agg.loc[agg["Period"] == period, col_name].values[0]
                                total += val
                        except:
                            pass
                row[col] = total
            elif col.endswith(" Total"):
                # Year total
                year = col.split()[0]
                year_total = 0
                for m_col in month_cols:
                    if m_col.startswith(year + "-"):
                        try:
                            dt = pd.to_datetime(m_col)
                            period = dt.strftime("%b")
                            if period in agg["Period"].values:
                                val = agg.loc[agg["Period"] == period, col_name].values[0]
                                year_total += val
                        except:
                            pass
                row[col] = year_total
            else:
                # Month column
                try:
                    dt = pd.to_datetime(col)
                    period = dt.strftime("%b")
                    if period in agg["Period"].values:
                        val = agg.loc[agg["Period"] == period, col_name].values[0]
                    else:
                        val = 0
                except:
                    val = 0
                row[col] = val
        data.append(row)

    return pd.DataFrame(data)


def sorted_cats(long_bycat):
    present = long_bycat["Category"].unique().tolist()
    ordered = [c for c in CAT_ORDER if c in present]
    ordered += sorted(c for c in present if c not in CAT_ORDER)
    return ordered


def fmt(df_display):
    nc = [c for c in df_display.columns if c != df_display.columns[0]]
    num_idx = df_display.index[
        (df_display.iloc[:, 0] != "GP%") &
        (df_display.iloc[:, 0] != "---") &
        (~df_display.iloc[:, 0].str.startswith("FCST", na=False))
    ].tolist()
    styled = df_display.style.format(
        "{:,.0f}", subset=pd.IndexSlice[num_idx, nc], na_rep="0"
    )
    # For FCST rows, format as numbers
    fcst_idx = df_display.index[df_display.iloc[:, 0].str.startswith("FCST", na=False)].tolist()
    if fcst_idx:
        styled = styled.format(
            "{:,.0f}", subset=pd.IndexSlice[fcst_idx, nc], na_rep="0"
        )
    return styled


def fmt_num(n) -> str:
    """Format a number with M/B/K suffix for KPI display. Returns 'N/A' for None/NaN."""
    import math
    if n is None:
        return "N/A"
    try:
        if math.isnan(n):
            return "N/A"
    except TypeError:
        return "N/A"
    sign = "-" if n < 0 else ""
    abs_n = abs(n)
    if abs_n >= 1e9:
        return f"{sign}{abs_n / 1e9:.2f}B"
    if abs_n >= 1e6:
        return f"{sign}{abs_n / 1e6:.1f}M"
    if abs_n >= 1e3:
        return f"{sign}{abs_n / 1e3:.1f}K"
    return f"{'-' if n < 0 else ''}{abs_n:,.0f}"


def show_bycat(long_bycat):
    all_months = sorted(long_bycat["Month"].unique().tolist())
    for cat in sorted_cats(long_bycat):
        st.markdown(f"**{cat}**")
        st.dataframe(
            fmt(to_wide_one_cat(long_bycat, cat, all_months)),
            use_container_width=True,
        )


# ── Cached shipping search ────────────────────────────────────────
@st.cache_data
def cached_search_indices(part_numbers: tuple, keywords: tuple) -> list:
    """Return matching row indices. Cached across Streamlit reruns."""
    matched = set()
    for i, pn in enumerate(part_numbers):
        pn_lower = str(pn).lower()
        for kw in keywords:
            if kw.lower() in pn_lower:
                matched.add(i)
                break
    return sorted(matched)


# ── Dashboard helpers ─────────────────────────────────────────────
def calc_dashboard_kpis(df, prev_df=None):
    """Calculate top-level KPI metrics with YoY deltas."""
    revenue = df["SALES Total AMT"].sum()
    gp = df[GP_COL].sum()
    gp_pct = gp / revenue * 100 if revenue else 0.0
    qty = df[df["Category"].isin(QTY_CATEGORIES)]["QTY"].sum()

    result = {
        "revenue": revenue, "gp": gp, "gp_pct": gp_pct, "qty": qty,
    }

    if prev_df is not None and not prev_df.empty:
        p_rev = prev_df["SALES Total AMT"].sum()
        p_gp = prev_df[GP_COL].sum()
        p_gp_pct = p_gp / p_rev * 100 if p_rev else 0.0
        p_qty = prev_df[prev_df["Category"].isin(QTY_CATEGORIES)]["QTY"].sum()

        result["revenue_yoy"] = (revenue - p_rev) / p_rev * 100 if p_rev else None
        result["gp_yoy"] = (gp - p_gp) / p_gp * 100 if p_gp else None
        result["gp_pct_yoy"] = gp_pct - p_gp_pct  # ppt change
        result["qty_yoy"] = (qty - p_qty) / p_qty * 100 if p_qty else None
    else:
        for k in ("revenue_yoy", "gp_yoy", "gp_pct_yoy", "qty_yoy"):
            result[k] = None

    return result


def build_monthly_trend(df):
    """Aggregate monthly: Revenue, GP, GP%, with Year column for multi-year overlay."""
    m = df.copy()
    m["Year"] = m["Ship Date"].dt.year.astype(str)
    m["MonthNum"] = m["Ship Date"].dt.month
    m["Month"] = m["Ship Date"].dt.strftime("%Y-%m")

    agg = m.groupby(["Year", "MonthNum", "Month"], sort=True).agg(
        Revenue=("SALES Total AMT", "sum"),
        GP=(GP_COL, "sum"),
    ).reset_index()
    agg["GP%"] = agg.apply(
        lambda r: r["GP"] / r["Revenue"] * 100 if r["Revenue"] else 0.0, axis=1
    )
    return agg


def build_category_breakdown(df):
    """Category-level aggregation: Revenue, GP, QTY, percentage share."""
    agg = df.groupby("Category", sort=False).agg(
        Revenue=("SALES Total AMT", "sum"),
        GP=(GP_COL, "sum"),
        QTY=("QTY", "sum"),
    ).reset_index()
    total_rev = agg["Revenue"].sum()
    agg["Pct"] = agg["Revenue"] / total_rev * 100 if total_rev else 0.0
    agg["GP%"] = agg.apply(
        lambda r: r["GP"] / r["Revenue"] * 100 if r["Revenue"] else 0.0, axis=1
    )
    cat_rank = {c: i for i, c in enumerate(CAT_ORDER)}
    agg["_rank"] = agg["Category"].map(cat_rank).fillna(len(CAT_ORDER))
    agg = agg.sort_values("_rank").drop(columns=["_rank"]).reset_index(drop=True)
    return agg


def build_monthly_category(df):
    """Monthly x Category aggregation for stacked chart."""
    m = df.copy()
    m["Month"] = m["Ship Date"].dt.strftime("%Y-%m")
    agg = m.groupby(["Month", "Category"], sort=True).agg(
        Revenue=("SALES Total AMT", "sum"),
    ).reset_index()
    return agg


def build_customer_monthly_qty_by_cat(df: pd.DataFrame) -> pd.DataFrame:
    """Monthly QTY grouped by Category for a customer subset.
    Returns DataFrame with columns: Month, Category, QTY.
    """
    m = df.copy()
    m["Month"] = m["Ship Date"].dt.strftime("%Y-%m")
    agg = (
        m.groupby(["Month", "Category"], sort=True)["QTY"]
        .sum()
        .reset_index()
    )
    cat_rank = {c: i for i, c in enumerate(CAT_ORDER)}
    agg["_rank"] = agg["Category"].map(cat_rank).fillna(len(CAT_ORDER))
    agg = agg.sort_values(["Month", "_rank"]).drop(columns=["_rank"]).reset_index(drop=True)
    return agg


def build_top_customers(df, n=10, prev_df=None, fcst_df=None):
    """Top N customers by revenue with GP, GP%, QTY, YoY, and optional FY Forecast."""
    agg = df.groupby("Customer Name", sort=False).agg(
        Revenue=("SALES Total AMT", "sum"),
        GP=(GP_COL, "sum"),
    ).reset_index()
    _qty_agg = (
        df[df["Category"].isin(QTY_CATEGORIES)]
        .groupby("Customer Name", sort=False)["QTY"].sum()
        .reset_index()
    )
    agg = agg.merge(_qty_agg, on="Customer Name", how="left").fillna({"QTY": 0})
    agg["GP%"] = agg.apply(
        lambda r: r["GP"] / r["Revenue"] * 100 if r["Revenue"] else 0.0, axis=1
    )

    if prev_df is not None and not prev_df.empty:
        prev_agg = prev_df.groupby("Customer Name", sort=False).agg(
            Prev_Revenue=("SALES Total AMT", "sum"),
        ).reset_index()
        agg = agg.merge(prev_agg, on="Customer Name", how="left")
        agg["YoY%"] = agg.apply(
            lambda r: (r["Revenue"] - r["Prev_Revenue"]) / r["Prev_Revenue"] * 100
            if pd.notna(r.get("Prev_Revenue")) and r["Prev_Revenue"] != 0
            else None,
            axis=1,
        )
        agg = agg.drop(columns=["Prev_Revenue"])
    else:
        agg["YoY%"] = None

    agg = agg.sort_values("Revenue", ascending=False).head(n).reset_index(drop=True)

    if fcst_df is not None and not fcst_df.empty:
        fy_agg = fcst_df.groupby("Customer", sort=False)["AMT"].sum().reset_index()
        fy_agg.columns = ["Customer Name", "FY Forecast"]
        ytd_agg = (
            fcst_df[fcst_df["Source"] == "Actual"]
            .groupby("Customer", sort=False)["AMT"].sum()
            .reset_index()
        )
        ytd_agg.columns = ["Customer Name", "_YTD"]
        agg = agg.merge(fy_agg, on="Customer Name", how="left").fillna({"FY Forecast": 0})
        agg = agg.merge(ytd_agg, on="Customer Name", how="left").fillna({"_YTD": 0})
        agg["Achievement%"] = agg.apply(
            lambda r: r["_YTD"] / r["FY Forecast"] * 100 if r["FY Forecast"] else None,
            axis=1,
        )
        agg = agg.drop(columns=["_YTD"])

    agg.index = agg.index + 1
    agg.index.name = "Rank"
    return agg


def build_customer_detail(df, customers):
    """Customer(s) monthly breakdown for drill-down.
    Returns (kpis_dict, monthly_df, category_df).
    *customers* can be a single name (str) or a list of names.
    """
    if isinstance(customers, str):
        customers = [customers]
    cust_df = df[df["Customer Name"].isin(customers)].copy()
    if cust_df.empty:
        return {}, pd.DataFrame(), pd.DataFrame()

    kpis = {
        "revenue": cust_df["SALES Total AMT"].sum(),
        "gp": cust_df[GP_COL].sum(),
        "qty": cust_df[cust_df["Category"].isin(QTY_CATEGORIES)]["QTY"].sum(),
    }
    total_rev = kpis["revenue"]
    kpis["gp_pct"] = kpis["gp"] / total_rev * 100 if total_rev else 0.0

    cust_df["Month"] = cust_df["Ship Date"].dt.strftime("%Y-%m")
    monthly = cust_df.groupby("Month", sort=True).agg(
        Revenue=("SALES Total AMT", "sum"),
        GP=(GP_COL, "sum"),
    ).reset_index()
    _qty_mo = (
        cust_df[cust_df["Category"].isin(QTY_CATEGORIES)]
        .groupby("Month", sort=True)["QTY"].sum()
        .reset_index()
    )
    monthly = monthly.merge(_qty_mo, on="Month", how="left").fillna({"QTY": 0})
    monthly["GP%"] = monthly.apply(
        lambda r: r["GP"] / r["Revenue"] * 100 if r["Revenue"] else 0.0, axis=1
    )

    cat_agg = cust_df.groupby("Category", sort=False).agg(
        Revenue=("SALES Total AMT", "sum"),
    ).reset_index()
    cat_total = cat_agg["Revenue"].sum()
    cat_agg["Pct"] = cat_agg["Revenue"] / cat_total * 100 if cat_total else 0.0
    cat_rank = {c: i for i, c in enumerate(CAT_ORDER)}
    cat_agg["_rank"] = cat_agg["Category"].map(cat_rank).fillna(len(CAT_ORDER))
    cat_agg = cat_agg.sort_values("_rank").drop(columns=["_rank"]).reset_index(drop=True)

    return kpis, monthly, cat_agg


def build_pn_detail(df, has_shipping=False):
    """Part Number breakdown for CDR/Tablet: QTY sum + latest UP."""
    sub = df[df["Category"].isin({"CDR", "Tablet"})].copy()
    if sub.empty:
        return pd.DataFrame()
    has_des = "DES" in sub.columns
    grp_cols = ["Category", "Part Number"] + (["DES"] if has_des else [])
    agg = sub.groupby(grp_cols, sort=False).agg(
        QTY=("QTY", "sum"),
    ).reset_index()
    if has_shipping and "UP" in sub.columns:
        latest = (
            sub.sort_values("Ship Date")
            .groupby("Part Number", sort=False)["UP"]
            .last()
            .reset_index()
            .rename(columns={"UP": "Latest UP"})
        )
        agg = agg.merge(latest, on="Part Number", how="left")
    agg = agg.sort_values(["Category", "QTY"], ascending=[True, False]).reset_index(drop=True)
    return agg
