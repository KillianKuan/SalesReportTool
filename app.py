import io
from datetime import datetime
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Performance Report Analysis Tool", layout="wide")
st.title("📊 Performance Report Data Analysis Tool")

REQUIRED_COLS = [
    "Customer Name", "Ship Date", "QTY",
    "SALES Total AMT", "final GP(NTD,data from Financial Report)",
    "Part Number", "Category"
]
VALID_CATEGORIES = {"Tablet", "CDR", "Tablet ACC", "CDR ACC"}
GP_COL = "final GP(NTD,data from Financial Report)"
# ── DES keyword classification rules (edit here; sync with page table) ──
DES_RULES = {
    "CDR ACC":    ["cdr", "gemini", "evo", "sprint", "sd card", "panic button", "iosix", "uvc camera", "k220", "k245", "k265"],
    "Tablet ACC": ["tablet", "chiron", "hera", "phaeton", "surfing pro", "cradle", "f840"],
}

# ── 1. Upload file ───────────────────────────────────────────────
uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])
if not uploaded:
    st.info("Please upload a .xlsx file containing the 'Actual' sheet.")
    st.stop()

# ── 2. Load, validate, clean (cached) ───────────────────────────
@st.cache_data
def load_and_clean(file_bytes):
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception as e:
        return None, 0, f"Cannot read file: {e}", [], False

    if "Actual" not in xl.sheet_names:
        return None, 0, f"'Actual' sheet not found. Available sheets: {xl.sheet_names}", [], False

    raw = xl.parse("Actual")
    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        return None, 0, f"Missing required columns: {missing}", [], False

    has_des = "DES" in raw.columns
    read_cols = REQUIRED_COLS + (["DES"] if has_des else [])
    df = raw[read_cols].copy()

    df["Ship Date"] = df["Ship Date"].astype(str).str.strip()
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    nat_count = int(df["Ship Date"].isna().sum())
    df = df.dropna(subset=["Ship Date"])
    df["Month"] = df["Ship Date"].dt.strftime("%Y-%m")

    df["Category"] = df["Category"].astype(str).str.strip()
    if has_des:
        df["DES"] = df["DES"].astype(str).str.strip()

    def classify_by_des(des_val):
        des_lower = des_val.lower()
        return [cat for cat, kws in DES_RULES.items() if any(kw in des_lower for kw in kws)]

    ambiguous_rows = []

    def normalize_category(row):
        cat = row["Category"]
        if cat in VALID_CATEGORIES:
            return cat
        if not has_des:
            return "Others"
        matches = classify_by_des(str(row.get("DES", "")))
        if len(matches) == 0:
            return "Others"
        elif len(matches) == 1:
            return matches[0]
        else:
            ambiguous_rows.append({
                "Part Number": row.get("Part Number", ""),
                "DES": row.get("DES", ""),
                "Original Category": cat,
                "Matched Categories": " / ".join(matches),
                "Assigned Category": matches[0],
            })
            return matches[0]

    df["Category"] = df.apply(normalize_category, axis=1)

    df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0)
    df["SALES Total AMT"] = pd.to_numeric(df["SALES Total AMT"], errors="coerce").fillna(0)
    df[GP_COL] = pd.to_numeric(df[GP_COL], errors="coerce").fillna(0)
    df["Customer Name"] = df["Customer Name"].astype(str).str.strip()
    df["Part Number"] = df["Part Number"].astype(str).str.strip()

    return df, nat_count, None, ambiguous_rows, has_des

df, nat_count, error_msg, ambiguous_rows, has_des = load_and_clean(uploaded.read())

if error_msg:
    st.error(error_msg)
    st.stop()
if not has_des:
    st.warning("⚠️ 'DES' column not found. DES-based classification disabled; unknown categories will fall back to Others.")
if nat_count > 0:
    st.warning(f"⚠️ {nat_count} row(s) with invalid or blank Ship Date skipped.")
if ambiguous_rows:
    st.warning(
        f"⚠️ {len(ambiguous_rows)} row(s) matched multiple DES categories. "
        f"Temporarily assigned to '{list(DES_RULES.keys())[0]}'. Please review:"
    )
    st.dataframe(pd.DataFrame(ambiguous_rows), use_container_width=True)

# ── 3. Customer search and selection ────────────────────────────
st.subheader("🔍 Customer Name")
cust_query = st.text_input("Enter Customer Name keyword (substring, case-insensitive)")

all_customers = sorted(df["Customer Name"].unique())

if not cust_query.strip():
    st.info("Enter a keyword to search for customers.")
    st.stop()

if st.session_state.get("_last_query") != cust_query:
    for k in list(st.session_state.keys()):
        if k.startswith("cust__"):
            del st.session_state[k]
    st.session_state["_last_query"] = cust_query

matched = [c for c in all_customers if cust_query.strip().lower() in c.lower()]

if not matched:
    st.warning("No matching customers found. Showing 0 rows.")
    st.stop()

st.markdown(f"**Found {len(matched)} customer(s). Select below:**")

selected_customers = []
for cust in matched:
    key = f"cust__{cust}"
    st.session_state.setdefault(key, True)
    if st.checkbox(cust, key=key):
        selected_customers.append(cust)

st.divider()
# ── 4. QTY: Tablet & CDR only ───────────────────────────────────
use_tablet_cdr_only = st.checkbox("QTY: sum only Tablet & CDR categories (exclude ACC)", value=True)

# ── 5. Category split ────────────────────────────────────────────
use_cat_split = st.checkbox("Split report by Category", value=True)
merge_cdr_acc = False
merge_tablet_acc = False
if use_cat_split:
    merge_cdr_acc = st.checkbox("  ↳ Merge CDR ACC into CDR", value=True)
    merge_tablet_acc = st.checkbox("  ↳ Merge Tablet ACC into Tablet", value=True)

# Fingerprint of current options — used to detect stale report
_current_opts = (
    use_tablet_cdr_only, use_cat_split, merge_cdr_acc, merge_tablet_acc,
    tuple(sorted(selected_customers))
)

# ── 6. Aggregation helpers ───────────────────────────────────────
def build_summary(base, use_tablet_cdr_only):
    """Build month-level summary (no category split)."""
    qty_base = base[base["Category"].isin({"Tablet", "CDR"})] if use_tablet_cdr_only else base
    agg = base.groupby("Month", sort=True).agg(
        **{"SALES Total AMT": ("SALES Total AMT", "sum"), "final GP(NTD)": (GP_COL, "sum")}
    ).reset_index()
    qty_all = qty_base.groupby("Month")["QTY"].sum().reset_index().rename(columns={"QTY": "QTY (All)"})
    return agg.merge(qty_all, on="Month", how="left").fillna({"QTY (All)": 0})


def build_bycat(base, use_tablet_cdr_only, merge_cdr_acc, merge_tablet_acc):
    """Build month × category long table.

    - CDR ACC / Tablet ACC appear as separate rows when merge is OFF.
    - QTY counts only original Tablet & CDR rows when use_tablet_cdr_only is True.
    """
    cat_df = base.copy()
    # Keep original category labels before any merge for QTY masking
    original_cat = cat_df["Category"].copy()

    if merge_cdr_acc:
        cat_df["Category"] = cat_df["Category"].replace("CDR ACC", "CDR")
    if merge_tablet_acc:
        cat_df["Category"] = cat_df["Category"].replace("Tablet ACC", "Tablet")

    # SALES + GP: aggregate all rows per (Month, Category)
    agg = cat_df.groupby(["Month", "Category"], sort=True).agg(
        **{"SALES Total AMT": ("SALES Total AMT", "sum"), "final GP(NTD)": (GP_COL, "sum")}
    ).reset_index()

    # QTY: count only rows whose ORIGINAL category is Tablet or CDR
    if use_tablet_cdr_only:
        qty_mask = original_cat.isin({"Tablet", "CDR"})
    else:
        qty_mask = pd.Series(True, index=cat_df.index)

    qty_src = cat_df[qty_mask]
    qty_agg = (
        qty_src.groupby(["Month", "Category"])["QTY"].sum()
        .reset_index().rename(columns={"QTY": "QTY (All)"})
    )
    long = agg.merge(qty_agg, on=["Month", "Category"], how="left")
    long["QTY (All)"] = long["QTY (All)"].fillna(0)
    return long


def to_wide_summary(long_df):
    metrics = ["QTY (All)", "SALES Total AMT", "final GP(NTD)"]
    melted = long_df.melt(id_vars=["Month"], value_vars=metrics, var_name="Metric", value_name="Value")
    pivot = melted.pivot_table(index="Metric", columns="Month", values="Value", aggfunc="sum")
    pivot = pivot.reindex(metrics)
    pivot.columns.name = None
    month_cols = list(pivot.columns)
    pivot["Total"] = pivot[month_cols].sum(axis=1)
    return pivot.reset_index()


def to_wide_bycat(long_df):
    metrics = ["QTY (All)", "SALES Total AMT", "final GP(NTD)"]
    long_df = long_df.copy()
    long_df["Row"] = long_df["Category"] + " | " + long_df.get("_dummy", "")
    melted = long_df.melt(id_vars=["Month", "Category"], value_vars=metrics,
                          var_name="Metric", value_name="Value")
    melted["Row"] = melted["Category"] + " | " + melted["Metric"]
    pivot = melted.pivot_table(index="Row", columns="Month", values="Value", aggfunc="sum")
    pivot.columns.name = None
    return pivot.reset_index()


def format_wide(df):
    label_col = df.columns[0]
    num_cols = [c for c in df.columns if c != label_col]
    return df.style.format(formatter="{:,.0f}", subset=num_cols, na_rep="-")


def display_bycat_subtables(wide_bycat):
    """Render ByCategory as one small table per Category."""
    row_col = wide_bycat.columns[0]
    # Extract unique categories preserving order
    categories = list(dict.fromkeys(
        wide_bycat[row_col].str.split(" | ").str[0].tolist()
    ))
    for cat in categories:
        st.markdown(f"**{cat}**")
        cat_rows = wide_bycat[wide_bycat[row_col].str.startswith(cat + " | ")].copy()
        cat_rows[row_col] = cat_rows[row_col].str.replace(cat + " | ", "", regex=False)
        st.dataframe(format_wide(cat_rows.reset_index(drop=True)), use_container_width=True)

# ── 7. Run computation ───────────────────────────────────────────
if st.button("▶ Run"):
    if not selected_customers:
        st.warning("Please select at least one customer.")
        st.stop()

    base = df[df["Customer Name"].isin(selected_customers)].copy()
    if base.empty:
        st.warning("No data found for selected customer(s). Showing 0 rows.")
        st.stop()

    with st.spinner("Generating report..."):
        # Summary
        long_summary = build_summary(base, use_tablet_cdr_only)
        wide_summary = to_wide_summary(long_summary)

        # ByCategory
        wide_bycat = pd.DataFrame()
        if use_cat_split:
            long_bycat = build_bycat(base, use_tablet_cdr_only, merge_cdr_acc, merge_tablet_acc)
            wide_bycat = to_wide_bycat(long_bycat)

        # Others
        others_df = base[base["Category"] == "Others"].copy()

        # Excel export
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            wide_summary.to_excel(writer, sheet_name="Summary", index=False)
            if not wide_bycat.empty:
                wide_bycat.to_excel(writer, sheet_name="ByCategory", index=False)
        buf.seek(0)

        # Persist results + option snapshot
        st.session_state["rpt_summary"] = wide_summary
        st.session_state["rpt_bycat"] = wide_bycat
        st.session_state["rpt_others"] = others_df
        st.session_state["rpt_buf"] = buf.getvalue()
        st.session_state["rpt_has_des"] = has_des
        st.session_state["rpt_opts"] = _current_opts

# ── 8. Display (persists across reruns until next Run) ───────────
if "rpt_summary" in st.session_state:
    # Warn if options have changed since last Run
    if st.session_state.get("rpt_opts") != _current_opts:
        st.info("ℹ️ Options have changed — press **▶ Run** to refresh the report.")

    _summary = st.session_state["rpt_summary"]
    _bycat = st.session_state["rpt_bycat"]
    _others = st.session_state["rpt_others"]
    _buf = st.session_state["rpt_buf"]
    _has_des = st.session_state["rpt_has_des"]

    tab_labels = ["📋 Summary"]
    if not _bycat.empty:
        tab_labels.append("📊 ByCategory")
    tabs = st.tabs(tab_labels)

    with tabs[0]:
        st.dataframe(format_wide(_summary), use_container_width=True)

    if not _bycat.empty:
        with tabs[1]:
            display_bycat_subtables(_bycat)

    if not _others.empty:
        show_cols = ["Customer Name", "Month", "Part Number", "Category", "QTY", "SALES Total AMT"]
        if _has_des:
            show_cols.insert(3, "DES")
        with st.expander(f"⚠️ Others ({len(_others)} row(s)) — unclassified data, excluded from report"):
            st.dataframe(_others[show_cols].reset_index(drop=True), use_container_width=True)

    filename = datetime.now().strftime("sales_report_%Y%m%d_%H%M.xlsx")
    st.download_button(
        label="⬇️ Download Excel Report",
        data=_buf,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
