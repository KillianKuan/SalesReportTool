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

    # Date parsing
    df["Ship Date"] = df["Ship Date"].astype(str).str.strip()
    df["Ship Date"] = pd.to_datetime(df["Ship Date"], errors="coerce")
    nat_count = int(df["Ship Date"].isna().sum())
    df = df.dropna(subset=["Ship Date"])
    df["Month"] = df["Ship Date"].dt.strftime("%Y-%m")

    # Category normalization
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

    # Numeric conversion
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

# ── 6. Aggregation functions ─────────────────────────────────────
def build_long(grp_df, qty_df, group_cols):
    agg = grp_df.groupby(group_cols, sort=True).agg(
        **{
            "SALES Total AMT": ("SALES Total AMT", "sum"),
            "final GP(NTD)": (GP_COL, "sum"),
        }
    ).reset_index()
    qty_all = (
        qty_df.groupby(group_cols)["QTY"].sum()
        .reset_index().rename(columns={"QTY": "QTY (All)"})
    )
    agg = agg.merge(qty_all, on=group_cols, how="left")
    agg["QTY (All)"] = agg["QTY (All)"].fillna(0)
    return agg


def to_wide(long_df, group_cols, add_total=False):
    month_col = "Month"
    extra_cols = [c for c in group_cols if c != month_col]
    metrics = ["QTY (All)", "SALES Total AMT", "final GP(NTD)"]
    id_vars = extra_cols + [month_col]
    melted = long_df.melt(id_vars=id_vars, value_vars=metrics,
                          var_name="Metric", value_name="Value")
    if extra_cols:
        melted["Row"] = melted[extra_cols[0]] + " | " + melted["Metric"]
        pivot = melted.pivot_table(index="Row", columns=month_col,
                                   values="Value", aggfunc="sum")
    else:
        pivot = melted.pivot_table(index="Metric", columns=month_col,
                                   values="Value", aggfunc="sum")
        pivot = pivot.reindex(metrics)
    pivot.columns.name = None
    if add_total:
        month_cols = list(pivot.columns)
        pivot["Total"] = pivot[month_cols].sum(axis=1)
    pivot = pivot.reset_index()
    return pivot


def format_wide(df):
    label_col = df.columns[0]
    num_cols = [c for c in df.columns if c != label_col]
    return df.style.format(formatter="{:,.0f}", subset=num_cols, na_rep="-")


def display_bycat_subtables(wide_bycat):
    """Render ByCategory as one small table per Category (outer grouping)."""
    row_col = wide_bycat.columns[0]  # "Row" column: "CDR | QTY (All)", etc.
    categories = wide_bycat[row_col].str.split(" | ").str[0].unique()
    for cat in categories:
        st.markdown(f"**{cat}**")
        cat_rows = wide_bycat[wide_bycat[row_col].str.startswith(cat + " | ")].copy()
        # Strip "Category | " prefix → show only the Metric name
        cat_rows[row_col] = cat_rows[row_col].str.replace(f"{cat} | ", "", regex=False)
        st.dataframe(format_wide(cat_rows.reset_index(drop=True)), use_container_width=True)

# ── 7. Generate report ───────────────────────────────────────────
if st.button("▶ Run"):
    if not selected_customers:
        st.warning("Please select at least one customer.")
        st.stop()

    base = df[df["Customer Name"].isin(selected_customers)].copy()
    if base.empty:
        st.warning("No data found for selected customer(s). Showing 0 rows.")
        st.stop()

    with st.spinner("Generating report..."):
        qty_base = base[base["Category"].isin({"Tablet", "CDR"})] if use_tablet_cdr_only else base

        long_summary = build_long(base, qty_base, ["Month"])
        wide_summary = to_wide(long_summary, ["Month"], add_total=True)

        wide_bycat = pd.DataFrame()
        if use_cat_split:
            cat_df = base.copy()
            if merge_cdr_acc:
                cat_df["Category"] = cat_df["Category"].replace("CDR ACC", "CDR")
            if merge_tablet_acc:
                cat_df["Category"] = cat_df["Category"].replace("Tablet ACC", "Tablet")
            qty_base_cat = cat_df[base["Category"].isin({"Tablet", "CDR"})] if use_tablet_cdr_only else cat_df
            long_bycat = build_long(cat_df, qty_base_cat, ["Month", "Category"])
            wide_bycat = to_wide(long_bycat, ["Month", "Category"], add_total=False)

        # Collect Others rows for expander
        others_df = base[base["Category"] == "Others"].copy()

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            wide_summary.to_excel(writer, sheet_name="Summary", index=False)
            if use_cat_split and not wide_bycat.empty:
                wide_bycat.to_excel(writer, sheet_name="ByCategory", index=False)
        buf.seek(0)

    # ── Display results in tabs ──────────────────────────────────
    tab_labels = ["📋 Summary"]
    if use_cat_split and not wide_bycat.empty:
        tab_labels.append("📊 ByCategory")
    tabs = st.tabs(tab_labels)

    with tabs[0]:
        st.dataframe(format_wide(wide_summary), use_container_width=True)

    if use_cat_split and not wide_bycat.empty:
        with tabs[1]:
            display_bycat_subtables(wide_bycat)

    # ── Others expander ──────────────────────────────────────────
    if not others_df.empty:
        show_cols = ["Customer Name", "Month", "Part Number", "Category", "QTY", "SALES Total AMT"]
        if has_des:
            show_cols.insert(3, "DES")
        with st.expander(f"⚠️ Others ({len(others_df)} row(s)) — unclassified data, excluded from report"):
            st.dataframe(others_df[show_cols].reset_index(drop=True), use_container_width=True)

    # ── Download ─────────────────────────────────────────────────
    filename = datetime.now().strftime("sales_report_%Y%m%d_%H%M.xlsx")
    st.download_button(
        label="⬇️ Download Excel Report",
        data=buf,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
