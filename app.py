import io
import json
import os
from datetime import datetime
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Performance Report Analysis Tool", layout="wide")
st.title("📊 Performance Report Data Analysis Tool")

REQUIRED_COLS = [
    "Customer Name", "Ship Date", "QTY",
    "SALES Total AMT", "final GP(NTD,data from Financial Report)",
    "Part Number", "Category",
]
VALID_CATEGORIES = {"Tablet", "CDR", "Tablet ACC", "CDR ACC"}
_VALID_CAT_MAP   = {" ".join(c.upper().split()): c for c in VALID_CATEGORIES}
GP_COL           = "final GP(NTD,data from Financial Report)"
CAT_ORDER        = ["CDR", "CDR ACC", "Tablet", "Tablet ACC", "AI_SW", "Others"]
DES_RULES = {
    "CDR ACC":    ["cdr", "gemini", "evo", "sprint", "sd card", "panic button",
                   "iosix", "uvc camera", "k220", "k245", "k265",
                   "smart link dongle", "safetycam"],
    "Tablet ACC": ["tablet", "prometheus", "chiron", "hera", "phaeton", "surfing pro",
                   "cradle", "f840", "ulmo", "fleet cable"],
    "AI_SW":      ["visionmax"],
}

# ── 0. Data folder scanning ──────────────────────────────────────
# Resolve DATA_DIR relative to app.py's location (works both in dev and PyInstaller)
APP_DIR  = Path(__file__).resolve().parent
DATA_DIR = APP_DIR.parent / "data"


def scan_data_folders() -> dict[int, Path]:
    """Return {year: folder_path} for all valid year-named subdirs in DATA_DIR."""
    folders = {}
    if not DATA_DIR.exists():
        return folders
    for entry in sorted(DATA_DIR.iterdir()):
        if entry.is_dir() and entry.name.isdigit():
            year = int(entry.name)
            if 2019 <= year <= 2099:
                folders[year] = entry
    return folders


def get_latest_xlsx(year_dir: Path) -> Path | None:
    """Return the most-recently-modified .xlsx file in *year_dir*, or None."""
    xlsx_files = list(year_dir.glob("*.xlsx"))
    if not xlsx_files:
        return None
    return max(xlsx_files, key=lambda f: f.stat().st_mtime)


# ── 1. Year selection (sidebar) ──────────────────────────────────
year_folders = scan_data_folders()

if not year_folders:
    st.error(
        f"找不到資料夾。請在以下路徑建立年份資料夾並放入 .xlsx 檔案：\n\n"
        f"`{DATA_DIR}`\n\n"
        f"範例：`data/2024/sales_2024.xlsx`"
    )
    st.stop()

available_years = sorted(year_folders.keys())
current_year = datetime.now().year
default_years = [current_year] if current_year in available_years else [available_years[-1]]

st.sidebar.header("📅 年份選擇")
selected_years = st.sidebar.multiselect(
    "選擇要分析的年份",
    options=available_years,
    default=default_years,
    format_func=lambda y: f"{y}{'  ⬅ 當年度' if y == current_year else ''}",
)

if not selected_years:
    st.info("請在左側選擇至少一個年份。")
    st.stop()

# Show which file will be used per year
file_map: dict[int, Path] = {}
for yr in selected_years:
    f = get_latest_xlsx(year_folders[yr])
    if f:
        file_map[yr] = f

missing_years = [yr for yr in selected_years if yr not in file_map]
if missing_years:
    st.warning(f"⚠️ 以下年份資料夾內沒有 .xlsx 檔案：{missing_years}")
if not file_map:
    st.error("所有選取的年份都沒有可用的 .xlsx 檔案。")
    st.stop()

with st.sidebar.expander("📄 使用中的檔案", expanded=True):
    for yr in sorted(file_map):
        f = file_map[yr]
        mtime = datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        st.markdown(f"**{yr}**: `{f.name}`  \n<small>修改時間: {mtime}</small>", unsafe_allow_html=True)

# ── 2. Load & clean (cached) ─────────────────────────────────────
def _rules_key():
    """Convert DES_RULES to a hashable tuple for cache busting."""
    return tuple((k, tuple(v)) for k, v in DES_RULES.items())


@st.cache_data
def load_single_file(file_path: str, _rules_key):
    """Load and clean a single .xlsx file. Returns (df, nat_count, err, ambiguous, has_des)."""
    try:
        xl = pd.ExcelFile(file_path)
    except Exception as e:
        return None, 0, f"Cannot read {file_path}: {e}", [], False
    if "Actual" not in xl.sheet_names:
        return None, 0, f"'{Path(file_path).name}': 'Actual' sheet not found. Available: {xl.sheet_names}", [], False
    raw = xl.parse("Actual")
    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        return None, 0, f"'{Path(file_path).name}': Missing columns: {missing}", [], False

    has_des = "DES" in raw.columns
    df = raw[REQUIRED_COLS + (["DES"] if has_des else [])].copy()
    df["Ship Date"] = pd.to_datetime(df["Ship Date"].astype(str).str.strip(), errors="coerce")
    nat_count = int(df["Ship Date"].isna().sum())
    df = df.dropna(subset=["Ship Date"])
    df["Month"]    = df["Ship Date"].dt.strftime("%Y-%m")
    df["Category"] = df["Category"].astype(str).str.strip()
    if has_des:
        df["DES"] = df["DES"].astype(str).str.strip()

    def by_des(des):
        d = des.lower()
        return [c for c, kws in DES_RULES.items() if any(k in d for k in kws)]

    ambiguous = []

    def norm_cat(row):
        cat = row["Category"]
        key = " ".join(cat.upper().split())
        if key in _VALID_CAT_MAP:
            return _VALID_CAT_MAP[key]
        if not has_des:
            return "Others"
        hits = by_des(str(row.get("DES", "")))
        if not hits:
            return "Others"
        if len(hits) == 1:
            return hits[0]
        ambiguous.append({"Part Number": row.get("Part Number", ""),
                          "DES": row.get("DES", ""),
                          "Original Category": cat,
                          "Matched": " / ".join(hits),
                          "Assigned": hits[0]})
        return hits[0]

    df["Category"] = df.apply(norm_cat, axis=1)
    for col in ["QTY", "SALES Total AMT", GP_COL]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    df["Customer Name"] = df["Customer Name"].astype(str).str.strip()
    df["Part Number"]   = (df["Part Number"].astype(str).str.strip()
                           .replace({"None": "", "nan": "", "NaN": ""}))
    return df, nat_count, None, ambiguous, has_des


# Load and merge all selected years
all_dfs = []
total_nat = 0
all_ambiguous = []
global_has_des = True

for yr in sorted(file_map):
    fp = file_map[yr]
    df_yr, nat_yr, err_yr, amb_yr, hd_yr = load_single_file(str(fp), _rules_key())
    if err_yr:
        st.error(f"❌ {err_yr}")
        continue
    all_dfs.append(df_yr)
    total_nat += nat_yr
    all_ambiguous.extend(amb_yr)
    if not hd_yr:
        global_has_des = False

if not all_dfs:
    st.error("沒有成功載入任何檔案。"); st.stop()

df = pd.concat(all_dfs, ignore_index=True)
has_des = global_has_des

# ── Overrides: persistent composite-key store ──────────────────
OVERRIDES_FILE = str(APP_DIR / "overrides.json")

def _save_overrides(ov):
    try:
        with open(OVERRIDES_FILE, "w", encoding="utf-8") as f:
            json.dump([[list(k), v] for k, v in ov.items()], f,
                      ensure_ascii=False, indent=2)
    except Exception:
        pass

def _load_overrides():
    try:
        if os.path.exists(OVERRIDES_FILE):
            with open(OVERRIDES_FILE, encoding="utf-8") as f:
                return {tuple(row[0]): row[1] for row in json.load(f)}
    except Exception:
        pass
    return {}

if "others_overrides" not in st.session_state:
    st.session_state["others_overrides"] = _load_overrides()

if st.session_state["others_overrides"]:
    df = df.copy()
    for (cust, pn, month, des), new_cat in st.session_state["others_overrides"].items():
        if has_des:
            mask = (
                (df["Customer Name"] == cust) &
                (df["Part Number"]   == pn)   &
                (df["Month"]         == month) &
                (df["DES"]           == des)
            )
        else:
            mask = (
                (df["Customer Name"] == cust) &
                (df["Part Number"]   == pn)   &
                (df["Month"]         == month)
            )
        df.loc[mask, "Category"] = new_cat

if not has_des:
    st.warning("⚠️ 'DES' column not found in one or more files — DES classification disabled.")
if total_nat:
    st.warning(f"⚠️ {total_nat} row(s) with invalid Ship Date skipped.")
if all_ambiguous:
    with st.expander(f"⚠️ {len(all_ambiguous)} row(s) matched multiple DES categories. Assigned to first match:", expanded=True):
        st.dataframe(pd.DataFrame(all_ambiguous), use_container_width=True)

st.sidebar.markdown(f"---\n**已載入 {len(df):,} 筆資料**（{len(file_map)} 個年份）")

# ── 3. Customer search ───────────────────────────────────────────
st.subheader("🔍 Customer Name")
cust_query = st.text_input("Enter keyword (substring, case-insensitive)")
all_customers = sorted(df["Customer Name"].unique())
if cust_query.strip():
    matched = [c for c in all_customers if cust_query.strip().lower() in c.lower()]
    if not matched:
        st.warning("No matching customers found.")
    else:
        st.markdown(f"**Found {len(matched)} customer(s):**")
        for c in matched:
            st.session_state.setdefault(f"cust__{c}", False)
            st.checkbox(c, key=f"cust__{c}")
else:
    st.info("Enter a keyword to search for customers.")

selected = [c for c in all_customers if st.session_state.get(f"cust__{c}", False)]
if selected:
    st.markdown("**✅ Selected ({}):** {}".format(
        len(selected), "\u3000".join(f"`{c}`" for c in selected)
    ))
    if st.button("🗑 Clear all selections"):
        for c in all_customers:
            st.session_state.pop(f"cust__{c}", None)
        st.rerun()

st.divider()

# ── 4. Options ───────────────────────────────────────────────────
qty_only  = st.checkbox("QTY: sum only Tablet & CDR (exclude ACC)", value=True)
by_cat    = st.checkbox("Split report by Category", value=True)
merge_cdr = merge_tab = False
if by_cat:
    merge_cdr = st.checkbox("  ↳ Merge CDR ACC into CDR",       value=True)
    merge_tab = st.checkbox("  ↳ Merge Tablet ACC into Tablet",  value=True)

_opts = (qty_only, by_cat, merge_cdr, merge_tab, tuple(sorted(selected)))

# ── 5. Helpers ───────────────────────────────────────────────────
def build_summary(base, qty_only):
    src = base[base["Category"].isin({"Tablet", "CDR"})] if qty_only else base
    agg = base.groupby("Month", sort=True).agg(
        **{"SALES Total AMT": ("SALES Total AMT", "sum"), "final GP(NTD)": (GP_COL, "sum")}
    ).reset_index()
    qty = src.groupby("Month")["QTY"].sum().reset_index().rename(columns={"QTY": "QTY (All)"})
    return agg.merge(qty, on="Month", how="left").fillna({"QTY (All)": 0})

def build_bycat(base, qty_only, merge_cdr, merge_tab):
    cat_df = base.copy()
    orig   = cat_df["Category"].copy()
    if merge_cdr: cat_df["Category"] = cat_df["Category"].replace("CDR ACC",    "CDR")
    if merge_tab: cat_df["Category"] = cat_df["Category"].replace("Tablet ACC", "Tablet")
    agg = cat_df.groupby(["Month", "Category"], sort=True).agg(
        **{"SALES Total AMT": ("SALES Total AMT", "sum"), "final GP(NTD)": (GP_COL, "sum")}
    ).reset_index()
    mask = orig.isin({"Tablet", "CDR"}) if qty_only else pd.Series(True, index=cat_df.index)
    qty  = (cat_df[mask].groupby(["Month", "Category"])["QTY"].sum()
            .reset_index().rename(columns={"QTY": "QTY (All)"}))
    long = agg.merge(qty, on=["Month", "Category"], how="left")
    long["QTY (All)"] = long["QTY (All)"].fillna(0)
    return long

def to_wide_summary(long_df):
    metrics = ["QTY (All)", "SALES Total AMT", "final GP(NTD)"]
    m = long_df.melt(id_vars=["Month"], value_vars=metrics, var_name="Metric", value_name="Value")
    p = m.pivot_table(index="Metric", columns="Month", values="Value", aggfunc="sum").reindex(metrics)
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
    g_vals = result.loc[result["Metric"] == "final GP(NTD)",   val_cols].values[0]
    gp_row = pd.DataFrame(
        [["GP%"] + [f"{g/s*100:.1f}%" if s != 0 else "-" for g, s in zip(g_vals, s_vals)]],
        columns=["Metric"] + val_cols,
    )
    return pd.concat([result, gp_row], ignore_index=True)

def to_wide_one_cat(long_df, cat, all_months):
    metrics = ["QTY (All)", "SALES Total AMT", "final GP(NTD)"]
    sub = long_df[long_df["Category"] == cat]
    m   = sub.melt(id_vars=["Month"], value_vars=metrics, var_name="Metric", value_name="Value")
    p   = m.pivot_table(index="Metric", columns="Month", values="Value", aggfunc="sum").reindex(metrics)
    p   = p.reindex(columns=all_months, fill_value=0).fillna(0)
    p.columns.name = None
    result = p.reset_index()
    val_cols = [c for c in result.columns if c != "Metric"]
    s_vals = result.loc[result["Metric"] == "SALES Total AMT", val_cols].values[0]
    g_vals = result.loc[result["Metric"] == "final GP(NTD)",   val_cols].values[0]
    gp_row = pd.DataFrame(
        [["GP%"] + [f"{g/s*100:.1f}%" if s != 0 else "-" for g, s in zip(g_vals, s_vals)]],
        columns=["Metric"] + val_cols,
    )
    return pd.concat([result, gp_row], ignore_index=True)

def sorted_cats(long_bycat):
    present = long_bycat["Category"].unique().tolist()
    ordered = [c for c in CAT_ORDER if c in present]
    ordered += sorted(c for c in present if c not in CAT_ORDER)
    return ordered

def fmt(df):
    nc      = [c for c in df.columns if c != df.columns[0]]
    num_idx = df.index[df.iloc[:, 0] != "GP%"].tolist()
    return df.style.format("{:,.0f}", subset=pd.IndexSlice[num_idx, nc], na_rep="0")

def show_bycat(long_bycat):
    all_months = sorted(long_bycat["Month"].unique().tolist())
    for cat in sorted_cats(long_bycat):
        st.markdown(f"**{cat}**")
        st.dataframe(fmt(to_wide_one_cat(long_bycat, cat, all_months)), use_container_width=True)

# ── 6. Run ───────────────────────────────────────────────────────
if st.button("▶ Run"):
    if not selected:
        st.warning("Please select at least one customer."); st.stop()
    base = df[df["Customer Name"].isin(selected)].copy()
    if base.empty:
        st.warning("No data for selected customer(s)."); st.stop()

    with st.spinner("Generating report..."):
        wide_summary = to_wide_summary(build_summary(base, qty_only))
        long_bycat   = build_bycat(base, qty_only, merge_cdr, merge_tab) if by_cat else pd.DataFrame()
        others_df    = base[base["Category"] == "Others"].copy()

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            wide_summary.to_excel(w, sheet_name="Summary", index=False)
            if not long_bycat.empty:
                all_months = sorted(long_bycat["Month"].unique().tolist())
                frames = []
                for cat in sorted_cats(long_bycat):
                    wc = to_wide_one_cat(long_bycat, cat, all_months)
                    wc.insert(0, "Category", cat)
                    frames.append(wc)
                pd.concat(frames, ignore_index=True).to_excel(w, sheet_name="ByCategory", index=False)
        buf.seek(0)

    st.session_state["rpt_summary"]    = wide_summary
    st.session_state["rpt_long_bycat"] = long_bycat
    st.session_state["rpt_others"]     = others_df
    st.session_state["rpt_buf"]        = buf.getvalue()
    st.session_state["rpt_has_des"]    = has_des
    st.session_state["rpt_opts"]       = _opts

# ── 7. Display ───────────────────────────────────────────────────
if "rpt_summary" not in st.session_state:
    st.stop()

if st.session_state.get("rpt_opts") != _opts:
    st.info("ℹ️ Options have changed — press **▶ Run** to refresh the report.")
_report_customers = list(st.session_state["rpt_opts"][4])
st.markdown("**Customer(s):** " + "\u3000".join(f"`{c}`" for c in _report_customers))

_summary    = st.session_state["rpt_summary"]
_long_bycat = st.session_state["rpt_long_bycat"]
_others     = st.session_state["rpt_others"]
_buf        = st.session_state["rpt_buf"]
_has_des    = st.session_state["rpt_has_des"]

tab_labels = ["📋 Summary"]
if not _long_bycat.empty:
    tab_labels.append("📊 ByCategory")
tabs = st.tabs(tab_labels)

with tabs[0]:
    st.dataframe(fmt(_summary), use_container_width=True)

if not _long_bycat.empty:
    with tabs[1]:
        show_bycat(_long_bycat)

if not _others.empty:
    _override_opts = ["Others (keep)"] + [c for c in CAT_ORDER if c != "Others"]
    with st.expander(f"⚠️ Others ({len(_others)} row(s)) — review & reassign category"):
        for _i, _row in _others.iterrows():
            _c1, _c2 = st.columns([4, 1])
            with _c1:
                _des_str = f" | DES: {_row['DES']}" if _has_des else ""
                st.markdown(
                    f"`{_row['Part Number']}`{_des_str}&nbsp;&nbsp;"
                    f"Month: **{_row['Month']}** | AMT: {int(_row['SALES Total AMT']):,}"
                )
            with _c2:
                _ok = (
                    _row["Customer Name"],
                    _row["Part Number"],
                    _row["Month"],
                    _row["DES"] if _has_des else "",
                )
                _cur = st.session_state["others_overrides"].get(_ok, "Others (keep)")
                if _cur not in _override_opts:
                    _cur = "Others (keep)"
                _choice = st.selectbox(
                    "Reassign",
                    _override_opts,
                    index=_override_opts.index(_cur),
                    key=f"override_{_i}_{'__'.join(str(x) for x in _ok)}",
                    label_visibility="collapsed",
                )
                if _choice != "Others (keep)":
                    st.session_state["others_overrides"][_ok] = _choice
                    _save_overrides(st.session_state["others_overrides"])
                elif _ok in st.session_state["others_overrides"]:
                    del st.session_state["others_overrides"][_ok]
                    _save_overrides(st.session_state["others_overrides"])
        if st.session_state["others_overrides"]:
            st.info("ℹ️ Overrides set — press **▶ Run** to apply to the report.")

st.download_button(
    "⬇️ Download Excel Report",
    data=_buf,
    file_name=datetime.now().strftime("sales_report_%Y%m%d_%H%M.xlsx"),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
