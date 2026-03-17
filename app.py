import io
from datetime import datetime
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
                   "iosix", "uvc camera", "k220", "k245", "k265"],
    "Tablet ACC": ["tablet", "prometheus", "chiron", "hera", "phaeton", "surfing pro", "cradle", "f840"],
    "AI_SW":      ["visionmax"],
}

# ── 1. Upload ────────────────────────────────────────────────────
uploaded = st.file_uploader("Upload .xlsx file", type=["xlsx"])
if not uploaded:
    st.info("Please upload a .xlsx file containing the 'Actual' sheet.")
    st.stop()

# ── 2. Load & clean (cached) ─────────────────────────────────────
@st.cache_data
def load_and_clean(file_bytes):
    try:
        xl = pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception as e:
        return None, 0, f"Cannot read file: {e}", [], False
    if "Actual" not in xl.sheet_names:
        return None, 0, f"'Actual' sheet not found. Available: {xl.sheet_names}", [], False
    raw = xl.parse("Actual")
    missing = [c for c in REQUIRED_COLS if c not in raw.columns]
    if missing:
        return None, 0, f"Missing columns: {missing}", [], False

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
    df["Part Number"]   = df["Part Number"].astype(str).str.strip()
    return df, nat_count, None, ambiguous, has_des

df, nat_count, err, ambiguous, has_des = load_and_clean(uploaded.read())
if err:
    st.error(err); st.stop()
if not has_des:
    st.warning("⚠️ 'DES' column not found — DES classification disabled.")
if nat_count:
    st.warning(f"⚠️ {nat_count} row(s) with invalid Ship Date skipped.")
if ambiguous:
    st.warning(f"⚠️ {len(ambiguous)} row(s) matched multiple DES categories. Assigned to first match:")
    st.dataframe(pd.DataFrame(ambiguous), use_container_width=True)

# ── 3. Customer search ───────────────────────────────────────────
st.subheader("🔍 Customer Name")
cust_query = st.text_input("Enter keyword (substring, case-insensitive)")
all_customers = sorted(df["Customer Name"].unique())
if not cust_query.strip():
    st.info("Enter a keyword to search for customers.")
    st.stop()
if st.session_state.get("_last_query") != cust_query:
    for k in [k for k in st.session_state if k.startswith("cust__")]:
        del st.session_state[k]
    st.session_state["_last_query"] = cust_query
matched = [c for c in all_customers if cust_query.strip().lower() in c.lower()]
if not matched:
    st.warning("No matching customers found."); st.stop()
st.markdown(f"**Found {len(matched)} customer(s):**")
selected = []
for c in matched:
    st.session_state.setdefault(f"cust__{c}", True)
    if st.checkbox(c, key=f"cust__{c}"):
        selected.append(c)

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
    p["Total"] = p.sum(axis=1)
    return p.reset_index()

def to_wide_one_cat(long_df, cat, all_months):
    """Pivot one category → wide; pad missing months with 0."""
    metrics = ["QTY (All)", "SALES Total AMT", "final GP(NTD)"]
    sub = long_df[long_df["Category"] == cat]
    m   = sub.melt(id_vars=["Month"], value_vars=metrics, var_name="Metric", value_name="Value")
    p   = m.pivot_table(index="Metric", columns="Month", values="Value", aggfunc="sum").reindex(metrics)
    p   = p.reindex(columns=all_months, fill_value=0).fillna(0)
    p.columns.name = None
    return p.reset_index()

def sorted_cats(long_bycat):
    """Return categories in CAT_ORDER; unknowns appended alphabetically."""
    present = long_bycat["Category"].unique().tolist()
    ordered = [c for c in CAT_ORDER if c in present]
    ordered += sorted(c for c in present if c not in CAT_ORDER)
    return ordered

def fmt(df):
    nc = [c for c in df.columns if c != df.columns[0]]
    return df.style.format("{:,.0f}", subset=nc, na_rep="0")

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
    cols = ["Customer Name", "Month", "Part Number", "Category", "QTY", "SALES Total AMT"]
    if _has_des:
        cols.insert(3, "DES")
    with st.expander(f"⚠️ Others ({len(_others)} row(s)) — unclassified, excluded from report"):
        st.dataframe(_others[cols].reset_index(drop=True), use_container_width=True)

st.download_button(
    "⬇️ Download Excel Report",
    data=_buf,
    file_name=datetime.now().strftime("sales_report_%Y%m%d_%H%M.xlsx"),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
