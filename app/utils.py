"""utils.py — Data loading, cleaning, classification, and report helpers."""

import functools
import hashlib
import html
import json
import os
import re
import string
from pathlib import Path

import pandas as pd
import streamlit as st

import palette

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
SETTINGS_FILE = str(APP_DIR / "settings.json")

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
_MISSING_KEY_TOKENS = {"", "nan", "NaN", "None"}


def override_key(customer, part_number, month, des) -> tuple[str, str, str, str]:
    """Normalize an override key: every field -> str, stripped, '' if missing."""
    def _norm(v) -> str:
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except (TypeError, ValueError):
            pass
        s = str(v).strip()
        return "" if s in _MISSING_KEY_TOKENS else s
    return (_norm(customer), _norm(part_number), _norm(month), _norm(des))


def override_key_series(df: pd.DataFrame) -> pd.Series:
    """Vectorized equivalent of override_key() for an entire DataFrame."""
    def _clean(col: pd.Series) -> pd.Series:
        s = col.astype(str).str.strip()
        bad = col.isna() | s.isin(_MISSING_KEY_TOKENS)
        return s.mask(bad, "")

    cust = _clean(df["Customer Name"])
    pn = _clean(df["Part Number"])
    month = _clean(df["Month"])
    des = _clean(df["DES"]) if "DES" in df.columns else pd.Series("", index=df.index)
    return pd.Series(list(zip(cust, pn, month, des)), index=df.index)


def save_overrides(ov):
    """Keys must already be normalized via override_key(). Raises on failure
    (e.g. a non-serializable value) so bugs surface during development instead
    of silently writing non-standard NaN into the file."""
    with open(OVERRIDES_FILE, "w", encoding="utf-8") as f:
        json.dump([[list(k), v] for k, v in ov.items()], f,
                  ensure_ascii=False, indent=2, allow_nan=False)


def load_overrides():
    try:
        if not os.path.exists(OVERRIDES_FILE):
            return {}
        with open(OVERRIDES_FILE, encoding="utf-8") as f:
            raw = json.load(f)
        migrated = {}
        changed = False
        for row in raw:
            key = tuple(row[0])
            new_key = override_key(*key)
            if new_key != key:
                changed = True
            migrated[new_key] = row[1]
    except Exception:
        return {}

    if changed:
        try:
            save_overrides(migrated)
        except Exception:
            pass
    return migrated


# ── User-facing Settings (ignore list / account match) ─────────────
# Default source for the ignore list; settings.json overrides this default.
DEFAULT_SETTINGS = {
    "ignored_customers": sorted(EXCLUDED_CUSTOMERS),
    "customer_aliases": {},
    "fcst_customer_aliases": {},
    "custom_groups": [],
}


def load_settings() -> dict:
    """Load app/settings.json layered over schema defaults.

    Tolerant of a missing or corrupt file (falls back to defaults, never
    raises). Unknown keys (including a legacy "theme" key from settings.json
    files written before the theme system was removed) or keys with the
    wrong type are ignored.
    """
    settings = {
        k: (list(v) if isinstance(v, list) else dict(v) if isinstance(v, dict) else v)
        for k, v in DEFAULT_SETTINGS.items()
    }
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                for key, default_val in DEFAULT_SETTINGS.items():
                    val = data.get(key)
                    if val is not None and isinstance(val, type(default_val)):
                        settings[key] = val
    except Exception:
        pass
    return settings


def save_settings(settings: dict) -> None:
    """Persist *settings* to app/settings.json (merged onto the current file)."""
    try:
        merged = load_settings()
        merged.update({k: v for k, v in settings.items() if k in DEFAULT_SETTINGS})
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(merged, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def get_shipped_aliases(kind: str) -> dict:
    """Public accessor for the read-only aliases shipped in app/aliases.json."""
    return dict(_load_aliases(kind))


def validate_alias_mappings(mapping: dict, valid_targets: set) -> list:
    """Return human-readable warnings for an alias/mapping dict.

    Checks for: self-referencing mappings (key normalizes to the same value
    as its target), duplicate keys that collide after normalization with
    conflicting targets, and targets that don't exist in *valid_targets*.
    """
    warnings = []
    seen_norm_keys = {}
    norm_valid_targets = {_normalize_name(t, upper=True) for t in valid_targets}
    for key, value in mapping.items():
        norm_key = _normalize_name(key, upper=True)
        norm_val = _normalize_name(value, upper=True)
        if norm_key == norm_val:
            warnings.append(f"Self-referencing mapping: '{key}' → '{value}'")
        if norm_key in seen_norm_keys and seen_norm_keys[norm_key] != value:
            warnings.append(
                f"Duplicate key after normalization: '{key}' collides with "
                f"another entry mapping to a different target"
            )
        seen_norm_keys[norm_key] = value
        if valid_targets and norm_val not in norm_valid_targets:
            warnings.append(f"Target '{value}' (for '{key}') not found in current data or groups")
    return warnings


def _settings_hash() -> str:
    """Hash of the settings fields that affect loaded data, for cache busting."""
    s = load_settings()
    relevant = {
        "ignored_customers": sorted(s.get("ignored_customers", [])),
        "customer_aliases": s.get("customer_aliases", {}),
        "fcst_customer_aliases": s.get("fcst_customer_aliases", {}),
    }
    blob = json.dumps(relevant, sort_keys=True, ensure_ascii=False)
    return hashlib.md5(blob.encode("utf-8")).hexdigest()


def inject_layout_css() -> None:
    """Inject layout-only CSS: card spacing/radius, control widths, sidebar
    nav structure. Colors are either inherited from Streamlit (``var(--text-
    color)`` etc. — no hardcoded Light/Dark palette) or, for the sidebar
    brand chrome and the KPI delta pills, a single fixed value from
    palette.py that does not change between Streamlit's light and dark
    themes (it never did — see the sidebar rule below).
    """
    st.markdown(
        f"""
<style>
[data-testid="stSidebar"] {{
    background-color: {palette.PRIMARY};
}}
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p, [data-testid="stSidebar"] label,
[data-testid="stSidebar"] .stMarkdown, [data-testid="stSidebar"] .stCaption,
[data-testid="stSidebar"] .stRadio label span, [data-testid="stSidebar"] .stCheckbox label span,
[data-testid="stSidebar"] small {{
    color: #FFFFFF !important;
}}
[data-testid="stSidebar"] input,
[data-testid="stSidebar"] [data-baseweb="select"] *,
[data-testid="stSidebar"] [data-baseweb="tag"] * {{
    color: var(--text-color) !important;
}}
[data-testid="stSidebar"] .stButton > button {{
    width: 100%;
    justify-content: flex-start;
    text-align: left;
    border-radius: 8px;
}}
/* Nav buttons sit on their own (near-white) button background, not the
   sidebar's dark green — so their label must use the ordinary Streamlit
   text color, not the white forced onto the rest of the sidebar chrome. */
[data-testid="stSidebar"] .stButton button[kind="secondary"] p,
[data-testid="stSidebar"] .stButton button[kind="secondary"] div {{
    color: var(--text-color) !important;
}}
/* FCST / System Info expander headers (sidebar only — main-content
   expanders are untouched). The header's own background is transparent,
   so it always shows the sidebar's dark green through it; the broad "p,
   label, ..." rule above already whitens the label text, but the toggle
   chevron and any other icon glyph under the header are plain <span>s
   that rule doesn't reach, so they're left at Streamlit's default dark
   text color — nearly invisible on dark green. Whiten every element in
   the header instead of guessing at each icon's selector, and keep the
   header surface itself transparent/white-tinted (not Streamlit's native
   light hover surface) so label, icon and chevron read as one coherent,
   sidebar-colored control in every state. */
[data-testid="stSidebar"] [data-testid="stExpander"] summary,
[data-testid="stSidebar"] [data-testid="stExpander"] summary * {{
    color: #FFFFFF !important;
}}
[data-testid="stSidebar"] [data-testid="stExpander"] details {{
    background-color: transparent;
    border-color: rgba(255, 255, 255, 0.35);
}}
[data-testid="stSidebar"] [data-testid="stExpander"] summary {{
    background-color: transparent;
}}
[data-testid="stSidebar"] [data-testid="stExpander"] summary:hover,
[data-testid="stSidebar"] [data-testid="stExpander"] summary:hover * {{
    background-color: transparent;
    color: rgba(255, 255, 255, 0.85) !important;
}}
[data-testid="stSidebar"] [data-testid="stExpander"] summary:focus-visible {{
    outline: 2px solid #FFFFFF;
    outline-offset: -2px;
    background-color: rgba(255, 255, 255, 0.12);
    /* Suppress Streamlit's default reddish primary-color focus ring so the
       header shows a single coherent white outline instead of two rings. */
    box-shadow: none !important;
}}

div[data-testid="stVerticalBlockBorderWrapper"] {{
    border-radius: 12px;
    padding: 4px 4px 12px 4px;
    margin-bottom: 24px;
}}

.main .stTextInput, [data-testid="stMain"] .stTextInput,
.main .stMultiSelect, [data-testid="stMain"] .stMultiSelect,
.main .stSelectbox, [data-testid="stMain"] .stSelectbox,
.main .stNumberInput, [data-testid="stMain"] .stNumberInput {{
    max-width: 480px;
}}

.sr-card-title {{
    font-size: 13px;
    font-weight: 600;
    color: var(--text-color);
    opacity: 0.65;
    margin-bottom: 8px;
}}

.sr-kpi-card {{ margin-bottom: 4px; }}
.sr-kpi-label {{
    font-size: 12px;
    font-weight: 600;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    color: var(--text-color);
    opacity: 0.65;
    margin-bottom: 4px;
}}
.sr-kpi-value {{
    font-size: 28px;
    font-weight: 600;
    color: var(--text-color);
    display: flex;
    align-items: baseline;
    gap: 8px;
    flex-wrap: wrap;
}}
.sr-kpi-delta {{
    font-size: 12px;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 999px;
}}
.sr-kpi-delta-pos {{ background-color: {palette.POSITIVE_BG}; color: {palette.POSITIVE}; }}
.sr-kpi-delta-neg {{ background-color: {palette.NEGATIVE_BG}; color: {palette.NEGATIVE}; }}
.sr-kpi-delta-na {{ background-color: {palette.MUTED_BG}; color: {palette.MUTED}; }}
.sr-kpi-caption {{ font-size: 12px; color: var(--text-color); opacity: 0.65; margin-top: 2px; }}

/* Report tables and PR headings key every color off `currentColor` (the
   real, correctly-inherited text color) rather than `var(--text-color)` /
   `var(--secondary-background-color)` / `var(--primary-color)` — this
   Streamlit version does not actually define those as CSS custom
   properties (confirmed: `var(--text-color, red)` resolves to red at the
   document root), so anything other than the `color` property itself
   would silently fall back to `transparent`/initial for non-inherited
   properties like `background-color`. `currentColor` needs no such
   variable: it just reads the element's own (correctly inherited) `color`,
   so tinting backgrounds/borders off it stays theme-correct for free. */
.sr-report-table-wrap {{
    width: 100%;
    overflow-x: auto;
    margin-bottom: 8px;
}}
.sr-report-table {{
    border-collapse: collapse;
    width: 100%;
    min-width: max-content;
    font-size: 14px;
}}
.sr-report-table th, .sr-report-table td {{
    padding: 6px 14px;
    white-space: nowrap;
    border: 1px solid color-mix(in srgb, currentColor 18%, transparent);
}}
.sr-report-table thead th {{
    background-color: color-mix(in srgb, currentColor 10%, transparent);
    font-weight: 600;
    text-align: left;
}}
.sr-report-table tbody td:first-child {{
    background-color: color-mix(in srgb, currentColor 6%, transparent);
    font-weight: 600;
    text-align: left;
}}
.sr-report-table tbody td:not(:first-child) {{
    text-align: left;
}}
.sr-report-table tbody tr:nth-child(even) td:not(:first-child) {{
    background-color: color-mix(in srgb, currentColor 5%, transparent);
}}
.sr-report-table tbody tr:hover td {{
    background-color: color-mix(in srgb, currentColor 16%, transparent) !important;
}}

.sr-pr-heading {{
    display: flex;
    align-items: center;
    font-size: 15px;
    font-weight: 700;
    padding: 6px 12px;
    margin-bottom: 10px;
    border-radius: 6px;
    border-left: 4px solid color-mix(in srgb, currentColor 55%, transparent);
    background-color: color-mix(in srgb, currentColor 8%, transparent);
}}
</style>
""",
        unsafe_allow_html=True,
    )


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


def _merged_customer_aliases() -> dict:
    """aliases.json 'customer' section with settings.json overrides layered on top."""
    aliases = dict(_load_aliases("customer"))
    aliases.update(load_settings().get("customer_aliases", {}))
    return aliases


def normalize_customer_name(name):
    """Normalize customer name with alias mapping."""
    normalized = _normalize_name(name, upper=True)
    aliases = _merged_customer_aliases()
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
    aliases = _merged_customer_aliases()
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
    """Convert DES_RULES + Settings into a hashable tuple for cache busting.

    Any change to the ignore list or the alias mappings edited in the
    Settings tab must invalidate @st.cache_data results, so the settings
    hash is folded in alongside the existing DES_RULES key.
    """
    return (
        tuple((k, tuple(v)) for k, v in DES_RULES.items()),
        _settings_hash(),
    )


def _apply_ignore_list(df: pd.DataFrame) -> tuple:
    """Drop rows whose (already-normalized) Customer Name is on the ignore list.

    Returns (filtered_df, ignored_row_count). Matching is on the normalized
    (no-punctuation, uppercase) form, same as the legacy EXCLUDED_CUSTOMERS.
    """
    ignored = {
        _normalize_name(c, upper=True)
        for c in load_settings().get("ignored_customers", [])
    }
    if not ignored or df.empty:
        return df, 0
    mask = df["Customer Name"].isin(ignored)
    ignored_count = int(mask.sum())
    return df[~mask].copy(), ignored_count


@st.cache_data
def load_single_file(file_path: str, rules_key):
    """Load and clean a single .xlsx file.
    Returns (df, nat_count, err, ambiguous, has_des, has_shipping, ignored_count).
    """
    try:
        xl = pd.ExcelFile(file_path, engine="calamine")
    except ImportError:
        xl = pd.ExcelFile(file_path)
    except Exception as e:
        return None, 0, f"Cannot read {file_path}: {e}", [], False, False, 0
    if "Actual" not in xl.sheet_names:
        return (None, 0,
                f"'{Path(file_path).name}': 'Actual' sheet not found. "
                f"Available: {xl.sheet_names}", [], False, False, 0)
    raw = xl.parse("Actual")
    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        return (None, 0,
                f"'{Path(file_path).name}': Missing columns: {missing}",
                [], False, False, 0)

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
    customer_aliases = _merged_customer_aliases()
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
    df, ignored_count = _apply_ignore_list(df)
    return df, nat_count, None, ambiguous, has_des, has_shipping, ignored_count


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
    Returns (df, nat_count, err, ambiguous, has_des, has_shipping, ignored_count).

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
        return None, 0, f"Cannot read historical data: {e}", [], False, False, 0

    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        return (None, 0,
                f"historical.csv: Missing columns: {missing}",
                [], False, False, 0)

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
    customer_aliases = _merged_customer_aliases()
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
    df, ignored_count = _apply_ignore_list(df)
    return df, nat_count, None, ambiguous, has_des, has_shipping, ignored_count


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


def build_bycat(base, qty_only, merge_acc):
    cat_df = base.copy()
    orig = cat_df["Category"].copy()
    if merge_acc:
        cat_df["Category"] = cat_df["Category"].replace(
            {"CDR ACC": "CDR", "Tablet ACC": "Tablet"}
        )
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


def _report_cell_text(label, value) -> str:
    """Format one Performance Report table cell to display text.

    Every row is a plain number (comma-grouped) except the "GP%" row —
    and any legacy "---"/"FCST..." label rows a shared caller might still
    pass in — which already carry pre-formatted display strings.
    """
    is_plain_row = label != "GP%" and label != "---" and not str(label).startswith("FCST")
    if pd.isna(value):
        return "0" if is_plain_row else str(value)
    return f"{value:,.0f}" if is_plain_row else str(value)


def pr_section_heading(text: str, icon: str | None = None) -> None:
    """Consistent, theme-adaptive heading for Performance Report sections:
    Filters, Results, Summary, By Category, and each category name (CDR,
    Tablet, ...). Uses the ``.sr-pr-heading`` CSS from
    ``inject_layout_css()``, which tints its background/border off
    ``currentColor`` — see the comment there for why (this Streamlit
    version doesn't actually expose ``--secondary-background-color`` /
    ``--primary-color`` as usable CSS variables).

    Mirrors ``components.card_title()``'s icon/text split (the
    ``:material/...:`` shorthand only expands in plain markdown, not in
    raw HTML), but is a separate class so it doesn't change the look of
    ``card_title()`` on other pages.
    """
    if icon:
        _icon_col, _text_col = st.columns([0.05, 0.95], gap="small")
        with _icon_col:
            st.markdown(f":material/{icon}:")
        with _text_col:
            st.markdown(f'<div class="sr-pr-heading">{text}</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="sr-pr-heading">{text}</div>', unsafe_allow_html=True)


def render_report_table(df_display: pd.DataFrame) -> None:
    """Render a Performance Report table (Summary or a single category) as
    a full-width, horizontally scrollable HTML table.

    Uses the shared ``.sr-report-table`` CSS from ``inject_layout_css()``:
    left-aligned numeric/percentage cells, a theme-colored header and
    label column, zebra striping and row hover — all derived from
    ``currentColor`` (see the comment in ``inject_layout_css()``), never a
    hard-coded palette. Rendered as plain HTML (not
    ``st.dataframe``/``st.table``) so columns keep a readable width and
    overflow via horizontal scroll instead of being compressed or wrapped.
    """
    label_col = df_display.columns[0]
    value_cols = [c for c in df_display.columns if c != label_col]

    header_html = "".join(f"<th>{html.escape(str(c))}</th>" for c in df_display.columns)
    body_rows = []
    for _, row in df_display.iterrows():
        label = row[label_col]
        cells = [f"<td>{html.escape(str(label))}</td>"]
        cells += [
            f"<td>{html.escape(_report_cell_text(label, row[c]))}</td>"
            for c in value_cols
        ]
        body_rows.append("<tr>" + "".join(cells) + "</tr>")

    st.markdown(
        '<div class="sr-report-table-wrap">'
        '<table class="sr-report-table">'
        f"<thead><tr>{header_html}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table></div>",
        unsafe_allow_html=True,
    )


def show_bycat(long_bycat):
    all_months = sorted(long_bycat["Month"].unique().tolist())
    for cat in sorted_cats(long_bycat):
        pr_section_heading(cat)
        render_report_table(to_wide_one_cat(long_bycat, cat, all_months))


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
