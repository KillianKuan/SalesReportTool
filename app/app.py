"""app.py - Performance Report Analysis Tool (v7.2)."""

import io
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

# Ensure local modules are importable
sys.path.insert(0, str(Path(__file__).resolve().parent))

import fcst_loader

from utils import (
    DATA_DIR, CAT_ORDER, _rules_key,
    scan_data_folders, get_latest_xlsx,
    load_single_file, load_overrides, save_overrides,
    build_summary, build_bycat,
    to_wide_summary, to_wide_one_cat,
    sorted_cats, fmt, show_bycat,
    cached_search_indices,
    EXCLUDED_CUSTOMERS,
    calc_dashboard_kpis, build_monthly_trend,
    build_category_breakdown, build_monthly_category,
    build_top_customers, build_customer_detail,
    build_customer_monthly_qty_by_cat,
    build_pn_detail,
)
from charts import (
    chart_up_tp_trend, chart_qty_by_year, chart_qty_by_month, chart_gp_pct_trend,
    chart_revenue_trend, chart_gp_dual_axis,
    chart_category_donut, chart_category_stacked, chart_ai_sw_revenue_trend,
    chart_top_customers_bar, chart_customer_monthly, chart_customer_cat_donut,
    chart_customer_qty_by_cat,
    chart_revenue_trend_blended, chart_gp_trend_blended,
)

st.set_page_config(
    page_title="Performance Report Analysis Tool",
    page_icon="📊",
    layout="wide",
)
st.title("📊 Performance Report Data Analysis Tool")

def inject_heartbeat() -> None:
    """Ping the launcher so it can close when the browser tab goes away."""
    port = os.environ.get("APP_HEARTBEAT_PORT")
    if not port:
        return

    components.html(
        f"""
        <script>
        (() => {{
          const url = "http://127.0.0.1:{port}/heartbeat";
          const ping = () => fetch(
            url,
            {{ method: "GET", mode: "cors", cache: "no-store" }}
          ).catch(() => null);
          ping();
          const id = window.setInterval(ping, 3000);
          window.addEventListener("beforeunload", ping);
          document.addEventListener("visibilitychange", () => {{
            if (document.visibilityState === "visible") ping();
          }});
          window.addEventListener("pagehide", () => {{
            window.clearInterval(id);
            if (navigator.sendBeacon) {{
              navigator.sendBeacon(url);
            }} else {{
              ping();
            }}
          }});
        }})();
        </script>
        """,
        height=0,
        width=0,
    )


inject_heartbeat()

# ?? 0. Year selection (sidebar) ??????????????????????????????????
year_folders = scan_data_folders()

if not year_folders:
    st.error(
        f"Data folder not found. Please create year folders and place .xlsx files in:\n\n"
        f"`{DATA_DIR}`\n\n"
        f"Example: `data/2024/sales_2024.xlsx`"
    )
    st.stop()

available_years = sorted(year_folders.keys(), reverse=True)
current_year = datetime.now().year
default_years = [current_year] if current_year in available_years else [available_years[-1]]

st.sidebar.header("📅 Year Selection")
selected_years = st.sidebar.multiselect(
    "Select years to analyze",
    options=available_years,
    default=default_years,
    format_func=str,
)

if not selected_years:
    st.info("Please select at least one year from the sidebar.")
    st.stop()

file_map: dict[int, Path] = {}
for yr in selected_years:
    f = get_latest_xlsx(year_folders[yr])
    if f:
        file_map[yr] = f

missing_years = [yr for yr in selected_years if yr not in file_map]
if missing_years:
    st.warning(f"⚠️ No .xlsx files found in year folders: {missing_years}")
if not file_map:
    st.error("No usable .xlsx files found for any of the selected years.")
    st.stop()

# ?? 1. Load & merge ???????????????????????????????????????????
all_dfs = []
total_nat = 0
all_ambiguous = []
global_has_des = True
global_has_shipping = True

for yr in sorted(file_map):
    fp = file_map[yr]
    df_yr, nat_yr, err_yr, amb_yr, hd_yr, hs_yr = load_single_file(
        str(fp), _rules_key()
    )
    if err_yr:
        st.error(f"❌ {err_yr}")
        continue
    all_dfs.append(df_yr)
    total_nat += nat_yr
    all_ambiguous.extend(amb_yr)
    if not hd_yr:
        global_has_des = False
    if not hs_yr:
        global_has_shipping = False

if not all_dfs:
    st.error("No files were loaded successfully.")
    st.stop()

df = pd.concat(all_dfs, ignore_index=True)
df = df[~df["Customer Name"].isin(EXCLUDED_CUSTOMERS)]
has_des = global_has_des
has_shipping = global_has_shipping

# ?? Overrides ????????????????????????????????????????????????????
if "others_overrides" not in st.session_state:
    st.session_state["others_overrides"] = load_overrides()

if st.session_state["others_overrides"]:
    df = df.copy()
    for (cust, pn, month, des), new_cat in st.session_state["others_overrides"].items():
        if "DES" in df.columns:
            mask = (
                (df["Customer Name"] == cust)
                & (df["Part Number"] == pn)
                & (df["Month"] == month)
                & (df["DES"] == des)
            )
        else:
            mask = (
                (df["Customer Name"] == cust)
                & (df["Part Number"] == pn)
                & (df["Month"] == month)
            )
        df.loc[mask, "Category"] = new_cat

# ?? Sales Person filter (sidebar) ????????????????????????????????
_sp_visible_custs: set[str] = set()
if "SALE_Person" in df.columns:
    _sp_year_rows = df[df["Ship Date"].dt.year == current_year]
    _sp_names = sorted(
        sp for sp in _sp_year_rows["SALE_Person"].dropna().unique()
        if sp not in ("nan", "NaN", "")
    )
    if _sp_names:
        st.sidebar.header("👤 Sales Person")
        _sel_persons = st.sidebar.multiselect(
            "Filter by Sales Person (current year)",
            options=_sp_names,
        )
        if _sel_persons:
            _sp_custs = sorted(
                c for c in _sp_year_rows[
                    _sp_year_rows["SALE_Person"].isin(_sel_persons)
                ]["Customer Name"].dropna().unique()
                if c not in ("nan", "NaN", "")
            )
            if _sp_custs:
                _sp_visible_custs = set(_sp_custs)
                st.sidebar.markdown(f"**🔗 Related Customers ({len(_sp_custs)}):**")
                for c in _sp_custs:
                    st.session_state.setdefault(f"sp_cust__{c}", False)
                    st.sidebar.checkbox(c, key=f"sp_cust__{c}")

# ?? Sidebar: System Info ???????????????????????????????????????
st.sidebar.header("📈 FCST")
_fcst_sheet = st.sidebar.radio(
    "FCST Sheet",
    options=["All Sheets", "Div.1&2_All", "VT", "Signify"],
    index=0,
    key="fcst_sheet",
)

with st.sidebar.expander("ℹ️ System Info", expanded=False):
    st.markdown(f"**Loaded {len(df):,} rows** ({len(file_map)} year(s))")
    for yr in sorted(file_map):
        f = file_map[yr]
        mtime = datetime.fromtimestamp(f.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        st.markdown(
            f"**{yr}**: `{f.name}`  \n<small>Modified: {mtime}</small>",
            unsafe_allow_html=True,
        )
    if not has_des:
        st.warning("'DES' column not found; DES classification disabled.")
    if not has_shipping:
        st.warning(
            "'Currency'/'UP'/'TP(USD)' columns not found; "
            "Shipping Record Search disabled."
        )
    if "SALE_Person" not in df.columns:
        st.warning("'SALE_Person' column not found; Sales Person filter disabled.")
    if total_nat:
        st.warning(f"{total_nat} row(s) with invalid Ship Date skipped.")
    if all_ambiguous:
        st.warning(
            f"{len(all_ambiguous)} row(s) matched multiple DES categories. "
            "Assigned to first match."
        )
        st.dataframe(pd.DataFrame(all_ambiguous), use_container_width=True)


# ?? YoY comparison data (for Dashboard) ??????????????????????????
yoy_df = None
_max_sel_year = max(selected_years)
_yoy_year = _max_sel_year - 1
if _yoy_year in year_folders:
    if _yoy_year in selected_years:
        yoy_df = df[df["Ship Date"].dt.year == _yoy_year].copy()
    else:
        _yoy_file = get_latest_xlsx(year_folders[_yoy_year])
        if _yoy_file:
            _yoy_raw, _, _yoy_err, _, _, _ = load_single_file(
                str(_yoy_file), _rules_key()
            )
            if _yoy_raw is not None:
                yoy_df = _yoy_raw[
                    ~_yoy_raw["Customer Name"].isin(EXCLUDED_CUSTOMERS)
                ]

# ??????????????????????????????????????????????????????????????
# MAIN TABS
# ??????????????????????????????????????????????????????????????
main_tab1, main_tab2, main_tab3 = st.tabs(
    ["📄 Performance Report", "🚚 Shipping Record Search", "📊 Company Dashboard"]
)

# ?? TAB 1: Performance Report ????????????????????????????????????
with main_tab1:
    st.subheader("🔎 Customer Name")
    cust_query = st.text_input("Enter keyword (substring, case-insensitive)")
    all_customers = sorted(df["Customer Name"].dropna().unique())
    if cust_query.strip():
        matched = [c for c in all_customers if cust_query.strip().lower() in c.lower()]
        if not matched:
            st.warning("No matching customers found.")
        else:
            st.markdown(f"**Found {len(matched)} customer(s):**")
            for c in matched:
                st.session_state.setdefault(f"cust__{c}", False)
                st.checkbox(c, key=f"cust__{c}")

    selected_keyword = [
        c for c in all_customers if st.session_state.get(f"cust__{c}", False)
    ]
    selected_sp = [
        c for c in all_customers
        if c in _sp_visible_custs and st.session_state.get(f"sp_cust__{c}", False)
    ]
    selected = sorted(set(selected_keyword + selected_sp))
    if selected:
        st.markdown(
            "**Selected ({}):** {}".format(
                len(selected), "\u3000".join(f"`{c}`" for c in selected)
            )
        )
        if st.button("🧹 Clear all selections"):
            for c in all_customers:
                st.session_state.pop(f"cust__{c}", None)
                st.session_state.pop(f"sp_cust__{c}", None)
            st.rerun()

    st.divider()

    qty_only = st.checkbox("QTY: sum only Tablet & CDR (exclude ACC)", value=True)
    by_cat = st.checkbox("Split report by Category", value=True)
    merge_cdr = merge_tab = False
    if by_cat:
        merge_cdr = st.checkbox("  ↪ Merge CDR ACC into CDR", value=True)
        merge_tab = st.checkbox("  ↪ Merge Tablet ACC into Tablet", value=True)

    _opts = (qty_only, by_cat, merge_cdr, merge_tab, tuple(sorted(selected)))

    if st.button("▶ Run"):
        if not selected:
            st.warning("Please select at least one customer.")
        else:
            base = df[df["Customer Name"].isin(selected)].copy()
            if base.empty:
                st.warning("No data for selected customer(s).")
            else:
                with st.spinner("Generating report..."):
                    wide_summary = to_wide_summary(build_summary(base, qty_only))
                    long_bycat = (
                        build_bycat(base, qty_only, merge_cdr, merge_tab)
                        if by_cat else pd.DataFrame()
                    )
                    others_df = base[base["Category"] == "Others"].copy()

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
                            pd.concat(frames, ignore_index=True).to_excel(
                                w, sheet_name="ByCategory", index=False
                            )
                    buf.seek(0)

                st.session_state["rpt_summary"] = wide_summary
                st.session_state["rpt_long_bycat"] = long_bycat
                st.session_state["rpt_others"] = others_df
                st.session_state["rpt_buf"] = buf.getvalue()
                st.session_state["rpt_has_des"] = has_des
                st.session_state["rpt_opts"] = _opts

    if "rpt_summary" in st.session_state:
        if st.session_state.get("rpt_opts") != _opts:
            st.info("Options have changed; press **▶ Run** to refresh the report.")
        _report_customers = list(st.session_state["rpt_opts"][4])
        st.markdown(
            "**Customer(s):** "
            + "\u3000".join(f"`{c}`" for c in _report_customers)
        )

        _summary = st.session_state["rpt_summary"]
        _long_bycat = st.session_state["rpt_long_bycat"]
        _others = st.session_state["rpt_others"]
        _buf = st.session_state["rpt_buf"]
        _has_des = st.session_state["rpt_has_des"]

        tab_labels = ["📋 Summary"]
        if not _long_bycat.empty:
            tab_labels.append("🗂️ By Category")
        tabs = st.tabs(tab_labels)

        with tabs[0]:
            st.dataframe(fmt(_summary), use_container_width=True)

        if not _long_bycat.empty:
            with tabs[1]:
                show_bycat(_long_bycat)

        if not _others.empty:
            _override_opts = ["Others (keep)"] + [
                c for c in CAT_ORDER if c != "Others"
            ]
            with st.expander(
                f"⚠️ Others ({len(_others)} row(s)) - review & reassign category"
            ):
                for _i, _row in _others.iterrows():
                    _c1, _c2 = st.columns([4, 1])
                    with _c1:
                        _des_str = f" | DES: {_row['DES']}" if _has_des else ""
                        st.markdown(
                            f"`{_row['Part Number']}`{_des_str}&nbsp;&nbsp;"
                            f"Month: **{_row['Month']}** | "
                            f"AMT: {int(_row['SALES Total AMT']):,}"
                        )
                    with _c2:
                        _ok = (
                            _row["Customer Name"],
                            _row["Part Number"],
                            _row["Month"],
                            _row["DES"] if _has_des else "",
                        )
                        _cur = st.session_state["others_overrides"].get(
                            _ok, "Others (keep)"
                        )
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
                            save_overrides(st.session_state["others_overrides"])
                        elif _ok in st.session_state["others_overrides"]:
                            del st.session_state["others_overrides"][_ok]
                            save_overrides(st.session_state["others_overrides"])
                if st.session_state["others_overrides"]:
                    st.info(
                        "Overrides updated; press **▶ Run** to apply them to the report."
                    )

        st.download_button(
            "📥 Download Excel Report",
            data=_buf,
            file_name=datetime.now().strftime("sales_report_%Y%m%d_%H%M.xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# ?? TAB 2: Shipping Record Search ????????????????????????????????
with main_tab2:
    if not has_shipping:
        st.warning(
            "Shipping Record Search requires **Currency**, **UP**, and **TP(USD)** "
            "columns in the data files. These columns were not found."
        )
    else:
        st.subheader("🔎 Search by Part Number")

        # ?? Search history (quick-select) ??
        if "search_history" not in st.session_state:
            st.session_state["search_history"] = []
        if st.session_state["search_history"]:
            st.caption("🕘 Recent searches")
            _hist_cols = st.columns(
                min(len(st.session_state["search_history"]), 5)
            )
            for _hi, _hq in enumerate(
                st.session_state["search_history"][:5]
            ):
                if _hist_cols[_hi].button(_hq, key=f"hist_{_hi}"):
                    st.session_state["shipping_pn_query"] = _hq
                    st.rerun()

        pn_query = st.text_input(
            "Enter Part Number keyword(s)",
            placeholder="e.g. K245, F840  (comma or space separated)",
            key="shipping_pn_query",
        )

        if pn_query.strip():
            keywords = [
                k.strip()
                for k in re.split(r"[,\s]+", pn_query.strip())
                if k.strip()
            ]

            # ?? Cached search ??
            _indices = cached_search_indices(
                tuple(df["Part Number"].tolist()),
                tuple(keywords),
            )

            _result_cols = (
                ["Ship Date", "Customer Name", "Part Number"]
                + (["DES"] if has_des else [])
                + ["QTY", "Currency", "UP", "TP(USD)"]
            )
            results = df.iloc[_indices][_result_cols].copy()

            if results.empty:
                st.info("No matching shipping records found.")
            else:
                # ?? Save to search history ??
                _q = pn_query.strip()
                _hist = st.session_state["search_history"]
                if _q in _hist:
                    _hist.remove(_q)
                _hist.insert(0, _q)
                st.session_state["search_history"] = _hist[:10]

                # ?? Part Number selection (narrow down) ??
                _matched_pns = sorted(results["Part Number"].unique())
                if len(_matched_pns) > 1:
                    _pn_selection = st.multiselect(
                        "Matching Part Numbers (select to narrow down)",
                        options=_matched_pns,
                        key="shipping_pn_selection",
                    )
                    if _pn_selection:
                        results = results[
                            results["Part Number"].isin(_pn_selection)
                        ]

                results["GP"] = results["UP"] - results["TP(USD)"]

                # ?? Customer filter (optional) ??
                _matched_custs = sorted(results["Customer Name"].unique())
                _cust_filter_active = False
                if len(_matched_custs) > 1:
                    _cust_filter = st.multiselect(
                        "Filter by Customer (optional)",
                        options=_matched_custs,
                        key="shipping_cust_filter",
                    )
                    if _cust_filter:
                        results = results[
                            results["Customer Name"].isin(_cust_filter)
                        ]
                        _cust_filter_active = True

                st.markdown(
                    f"**Found {len(results):,} record(s)** matching: "
                    + ", ".join(f"`{k}`" for k in keywords)
                )

                # ?? QTY TTL & weighted averages ??
                _total_qty = results["QTY"].sum()
                if _total_qty > 0:
                    _wavg_up = (
                        (results["UP"] * results["QTY"]).sum() / _total_qty
                    )
                    _wavg_tp = (
                        (results["TP(USD)"] * results["QTY"]).sum() / _total_qty
                    )
                else:
                    _wavg_up = _wavg_tp = 0.0
                _wavg_gp = _wavg_up - _wavg_tp
                _wavg_gp_pct = (
                    f"{_wavg_gp / _wavg_up * 100:.1f}%"
                    if _wavg_up != 0 else "-"
                )

                # ?? Metric cards ??
                _mc1, _mc2, _mc3, _mc4 = st.columns(4)
                _mc1.metric("Total QTY", f"{_total_qty:,}")
                _mc2.metric("Avg UP (wavg)", f"{_wavg_up:,.2f}")
                _mc3.metric("Avg TP(USD) (wavg)", f"{_wavg_tp:,.2f}")
                _mc4.metric("GP%", _wavg_gp_pct)

                # ?? Charts (3 columns, all Altair) ??
                _cc1, _cc2, _cc3 = st.columns(3)
                with _cc1:
                    st.markdown("**📈 UP & TP(USD) Monthly Trend**")
                    st.altair_chart(
                        chart_up_tp_trend(results),
                        use_container_width=True,
                    )
                with _cc2:
                    _qty_mode = st.radio(
                        "QTY grouping",
                        ["By Year", "By Month"],
                        horizontal=True,
                        key="qty_chart_mode",
                        label_visibility="collapsed",
                    )
                    if _qty_mode == "By Year":
                        st.markdown("**📊 QTY by Year**")
                        st.altair_chart(
                            chart_qty_by_year(results),
                            use_container_width=True,
                        )
                    else:
                        st.markdown("**📊 QTY by Month**")
                        st.altair_chart(
                            chart_qty_by_month(results),
                            use_container_width=True,
                        )
                with _cc3:
                    st.markdown("**📉 GP% Monthly Trend**")
                    st.altair_chart(
                        chart_gp_pct_trend(results),
                        use_container_width=True,
                    )

                # ?? GP% conditional formatting helper ??
                def _color_gp_pct(val):
                    if isinstance(val, str) and "%" in val:
                        try:
                            n = float(val.replace("%", ""))
                            if n >= 30:
                                return "color: #2e7d32; font-weight: bold"
                            elif n < 15:
                                return "color: #c62828; font-weight: bold"
                        except ValueError:
                            pass
                    return ""

                # ?? Format display ??
                results["GP%"] = results.apply(
                    lambda r: f"{r['GP'] / r['UP'] * 100:.1f}%"
                    if r["UP"] != 0 else "-",
                    axis=1,
                )
                results = results.sort_values(
                    "Ship Date", ascending=False
                ).reset_index(drop=True)
                results["Ship Date"] = results[
                    "Ship Date"
                ].dt.strftime("%Y-%m-%d")

                styled = (
                    results.style
                    .format({
                        "QTY": "{:,.0f}",
                        "UP": "{:,.2f}",
                        "TP(USD)": "{:,.2f}",
                        "GP": "{:,.2f}",
                    })
                    .map(_color_gp_pct, subset=["GP%"])
                )
                st.dataframe(styled, use_container_width=True)

                # ?? Part Number summary table (when customer filter active) ??
                if _cust_filter_active:
                    st.divider()
                    st.markdown("**📋 Part Number Summary**")
                    if has_des and "DES" in results.columns:
                        _grp_df = results.copy()
                        _grp_df["DES"] = _grp_df["DES"].fillna("")
                        _grp_cols = ["Part Number", "DES"]
                    else:
                        _grp_df = results
                        _grp_cols = ["Part Number"]
                    _summary_tbl = (
                        _grp_df.groupby(_grp_cols, as_index=False)["QTY"]
                        .sum()
                        .sort_values("QTY", ascending=False)
                        .reset_index(drop=True)
                    )
                    _summary_tbl = _summary_tbl.rename(
                        columns={"QTY": "SUM(QTY)"}
                    )
                    st.dataframe(
                        _summary_tbl.style.format({"SUM(QTY)": "{:,.0f}"}),
                        use_container_width=True,
                        hide_index=True,
                    )

# ?? TAB 3: Company Dashboard ????????????????????????????????????
with main_tab3:
    # ?? Data prep (apply SALE_Person filter) ??
    dash_df = df.copy()
    dash_yoy = yoy_df.copy() if yoy_df is not None else None
    if _sp_visible_custs:
        dash_df = dash_df[dash_df["Customer Name"].isin(_sp_visible_custs)]
        if dash_yoy is not None:
            dash_yoy = dash_yoy[
                dash_yoy["Customer Name"].isin(_sp_visible_custs)
            ]

    # YoY delta: compare max selected year vs previous year
    _dash_max_yr = max(selected_years)
    _dash_curr = dash_df[dash_df["Ship Date"].dt.year == _dash_max_yr]
    _kpis_yoy = calc_dashboard_kpis(_dash_curr, dash_yoy)
    _kpis_all = calc_dashboard_kpis(dash_df)

    # ── FCST Integration ──────────────────────────────────────────
    _now = datetime.now()
    _current_month = _now.month
    _current_yr = _now.year
    # Only blend when the current calendar year is in the selection
    _do_fcst = _dash_max_yr == _current_yr

    _fcst_raw = pd.DataFrame()
    _blended_monthly = pd.DataFrame()
    _budget_monthly = pd.DataFrame()
    _fcst_cat_monthly = pd.DataFrame()

    if _do_fcst:
        try:
            _sheet_arg = None if _fcst_sheet == "All Sheets" else _fcst_sheet
            _fcst_raw = fcst_loader.get_fcst_for_dashboard(
                str(DATA_DIR), customer=None, sheet_name=_sheet_arg
            )
            if not _fcst_raw.empty:
                # Customer names are already normalized inside fcst_loader._parse_sheet()
                # via normalize_fcst_customer() — no post-processing needed here.
                # Prepare actual_for_blend: current-year rows, renamed for blend function
                _act_yr = dash_df[dash_df["Ship Date"].dt.year == _current_yr].copy()
                _act_yr = _act_yr.rename(columns={"Customer Name": "Customer"})
                _act_yr["Month"] = _act_yr["Ship Date"].dt.month

                _blended_raw = fcst_loader.blend_actual_fcst(
                    _act_yr, _fcst_raw, _current_month
                )
                _blended_monthly = fcst_loader.agg_blended_monthly(_blended_raw)
                _budget_monthly = fcst_loader.agg_budget_monthly(_fcst_raw)
                _fcst_cat_monthly = fcst_loader.agg_fcst_category_monthly(_fcst_raw)
        except Exception as _fcst_err:
            print(f"[FCST] Warning: failed to load/blend FCST data: {_fcst_err}")

    # Collect unmatched FCST customers for warning
    _unmatched_fcst = fcst_loader.get_unmatched_customers() if _do_fcst else set()
    if _unmatched_fcst:
        customer_list = [f"'{name}' (sheet: {sheet})" for name, sheet in sorted(_unmatched_fcst)]
        st.warning(
            f"⚠️ FCST 未匹配客戶 ({len(_unmatched_fcst)} 個): {', '.join(customer_list)}。請更新 aliases.json 的 fcst_customer section。"
        )

    # ?? Section 1: KPI Metric Cards ??
    st.subheader("📌 Overview")
    if dash_yoy is not None:
        st.caption(f"YoY delta: {_dash_max_yr} vs {_dash_max_yr - 1}")

    def _fmt_delta(val, suffix="%"):
        if val is None:
            return "N/A"
        return f"{val:+,.1f}{suffix}"

    _k1, _k2, _k3, _k4, _k5, _k6 = st.columns(6)
    _k1.metric("💰 Revenue", f"{_kpis_all['revenue']:,.0f}",
               delta=_fmt_delta(_kpis_yoy["revenue_yoy"]))
    _k2.metric("💵 GP", f"{_kpis_all['gp']:,.0f}",
               delta=_fmt_delta(_kpis_yoy["gp_yoy"]))
    _k3.metric("📈 GP%", f"{_kpis_all['gp_pct']:.1f}%",
               delta=_fmt_delta(_kpis_yoy["gp_pct_yoy"], " ppt"))
    _k4.metric("📦 QTY (CDR+Tablet)", f"{_kpis_all['qty']:,.0f}",
               delta=_fmt_delta(_kpis_yoy["qty_yoy"]))
    _k5.metric(
        "👥 Customers", f"{_kpis_all['customers']:,}",
        delta=(
            f"{_kpis_yoy['customers_yoy']:+,}"
            if _kpis_yoy["customers_yoy"] is not None else "N/A"
        ),
    )
    _k6.metric("🗂️ Categories", f"{_kpis_all['active_cats']}")

    # Full-Year Forecast row (only shown when current year is selected and FCST loaded)
    if not _blended_monthly.empty:
        _ytd_rev = _blended_monthly[_blended_monthly["Source"] == "Actual"]["Revenue"].sum()
        _ytd_gp = _blended_monthly[_blended_monthly["Source"] == "Actual"]["GP"].sum()
        _fy_rev = _blended_monthly["Revenue"].sum()
        _fy_gp = _blended_monthly["GP"].sum()
        _fy_gp_pct = _fy_gp / _fy_rev * 100 if _fy_rev else 0.0
        _fy_qty = _blended_monthly["QTY"].sum()

        st.caption(
            f"📊 Full-Year Forecast (YTD Actual + Remaining FCST) — sheet: **{_fcst_sheet}**"
        )
        _fk1, _fk2, _fk3, _fk4 = st.columns(4)
        _fk1.metric(
            "💰 FY Revenue Forecast", f"{_fy_rev:,.0f}",
            delta=f"YTD: {_ytd_rev:,.0f}", delta_color="off",
        )
        _fk2.metric(
            "💵 FY GP Forecast", f"{_fy_gp:,.0f}",
            delta=f"YTD: {_ytd_gp:,.0f}", delta_color="off",
        )
        _fk3.metric("📈 FY GP% Forecast", f"{_fy_gp_pct:.1f}%")
        _fk4.metric("📦 FY QTY Forecast", f"{_fy_qty:,.0f}")

        # Budget Achievement metrics (2nd row)
        try:
            _budget_monthly = fcst_loader.agg_budget_monthly(_fcst_raw)
            if not _budget_monthly.empty:
                _fy_budget_rev = _budget_monthly["Revenue"].sum()
                _budget_achievement_pct = (
                    _ytd_rev / _fy_budget_rev * 100 if _fy_budget_rev else 0.0
                )
                _bk1, _bk2 = st.columns(2)
                _bk1.metric(
                    "🎯 Budget Achievement%",
                    f"{_budget_achievement_pct:.1f}%",
                    delta=f"YTD vs FY Budget",
                    delta_color="off",
                )
                _bk2.metric(
                    "📊 FY Budget Revenue",
                    f"{_fy_budget_rev:,.0f}",
                    delta=f"Current: {_ytd_rev:,.0f}",
                    delta_color="off",
                )
        except Exception as _budget_err:
            print(f"[Dashboard] Warning: Failed to calculate Budget Achievement%: {_budget_err}")


    st.divider()

    # ?? Section 2: Monthly Trends ??
    st.subheader("📈 Monthly Trends")
    _trend = build_monthly_trend(dash_df)
    _multi_yr = len(selected_years) > 1
    _tr1, _tr2 = st.columns(2)
    
    # Prepare blended data with budget (if available)
    _chart_data_blended = _blended_monthly.copy() if not _blended_monthly.empty else pd.DataFrame()
    if not _budget_monthly.empty and not _chart_data_blended.empty:
        _chart_data_blended = pd.concat(
            [_chart_data_blended, _budget_monthly],
            ignore_index=True
        )
    
    with _tr1:
        st.markdown("**📈 Monthly Revenue Trend**")
        if not _chart_data_blended.empty:
            st.altair_chart(
                chart_revenue_trend_blended(_chart_data_blended),
                use_container_width=True,
            )
        else:
            st.altair_chart(
                chart_revenue_trend(_trend, multi_year=_multi_yr),
                use_container_width=True,
            )
    with _tr2:
        st.markdown("**📉 Monthly GP & GP% Trend**")
        if not _chart_data_blended.empty:
            st.altair_chart(
                chart_gp_trend_blended(_chart_data_blended),
                use_container_width=True,
            )
        else:
            st.altair_chart(
                chart_gp_dual_axis(_trend),
                use_container_width=True,
            )

    st.divider()

    # ?? Section 3: Category Analysis ??
    st.subheader("🧩 Category Analysis")
    _cat_br = build_category_breakdown(dash_df)
    _cat_mo = build_monthly_category(dash_df)
    _cat_cols_count = 4 if not _fcst_cat_monthly.empty else 3
    _cat_cols = st.columns(_cat_cols_count)
    with _cat_cols[0]:
        st.markdown("**🍩 Revenue by Category**")
        st.altair_chart(
            chart_category_donut(_cat_br), use_container_width=True,
        )
    with _cat_cols[1]:
        st.markdown("**📊 Category Revenue Trend**")
        st.altair_chart(
            chart_category_stacked(_cat_mo), use_container_width=True,
        )
    with _cat_cols[2]:
        st.markdown("**🤖 AI_SW Monthly Revenue Trend**")
        st.altair_chart(
            chart_ai_sw_revenue_trend(_cat_mo), use_container_width=True,
        )
    if not _fcst_cat_monthly.empty:
        with _cat_cols[3]:
            st.markdown("**📊 FCST Category Revenue**")
            _fcst_cat_display = _fcst_cat_monthly.rename(
                columns={"Cat": "Category", "Period": "Month"}
            )
            st.altair_chart(
                chart_category_stacked(_fcst_cat_display), use_container_width=True,
            )

    st.divider()

    # ?? Section 4: Top N Customers ??
    st.subheader("🏆 Top Customers")
    _top_n = st.slider(
        "Number of customers", 5, 30, 10, key="dash_top_n",
    )
    _top = build_top_customers(dash_df, _top_n, dash_yoy)
    _tn1, _tn2 = st.columns(2)
    with _tn1:
        st.markdown(f"**🏆 Top {_top_n} Customers by Revenue**")
        st.altair_chart(
            chart_top_customers_bar(_top), use_container_width=True,
        )
    with _tn2:
        st.markdown(f"**📋 Top {_top_n} Customers**")

        def _gp_color(val):
            try:
                n = float(val)
                if n >= 30:
                    return "color: #2e7d32; font-weight: bold"
                if n < 15:
                    return "color: #c62828; font-weight: bold"
            except (ValueError, TypeError):
                pass
            return ""

        st.dataframe(
            _top.style.format(
                {"Revenue": "{:,.0f}", "GP": "{:,.0f}",
                 "GP%": "{:.1f}", "QTY": "{:,.0f}", "YoY%": "{:+.1f}"},
                na_rep="-",
            ).map(_gp_color, subset=["GP%"]),
            use_container_width=True,
        )

    st.divider()

    # ?? Section 5: Customer Drill-Down ??
    st.subheader("🔍 Customer Drill-Down")
    _all_dash_custs = sorted(dash_df["Customer Name"].dropna().unique())
    _top_names = _top["Customer Name"].tolist()

    _dd1, _dd2 = st.columns([2, 1])
    with _dd1:
        _dd_cust = st.selectbox(
            "Select from Top customers",
            options=[""] + _top_names,
            format_func=lambda x: "Select a customer..." if x == "" else x,
            key="dash_dd_cust",
        )
    with _dd2:
        _dd_search = st.text_input(
            "Or search by name", key="dash_dd_search",
        )

    _targets = []
    if _dd_search.strip():
        _matches = [
            c for c in _all_dash_custs
            if _dd_search.strip().lower() in c.lower()
        ]
        if _matches:
            _targets = st.multiselect(
                "Matching customers (multi-select)",
                _matches, key="dash_dd_match",
            )
        else:
            st.info("No matching customers.")
    elif _dd_cust:
        _targets = [_dd_cust]

    if _targets:
        _dk, _dm, _dcat = build_customer_detail(dash_df, _targets)
        if _dk:
            # Prepare FCST data for selected customers
            _blended_chart_df = pd.DataFrame()
            _fy_forecast_revenue = 0
            _fy_forecast_gp = 0
            _fy_budget_revenue = 0
            _budget_achievement_pct = 0
            _fcst_filtered = pd.DataFrame()
            if _do_fcst and not _fcst_raw.empty:
                _fcst_filtered = _fcst_raw[_fcst_raw["Customer"].isin(_targets)].copy()
                if not _fcst_filtered.empty:
                    _actual_df = dash_df[dash_df["Customer Name"].isin(_targets)].rename(
                        columns={"Customer Name": "Customer"}
                    ).copy()
                    _actual_df["Month"] = _actual_df["Ship Date"].dt.month
                    _blended_df = fcst_loader.blend_actual_fcst(
                        _actual_df, _fcst_filtered, _current_month
                    )
                    _blended_monthly = fcst_loader.agg_blended_monthly(_blended_df)
                    _budget_monthly = fcst_loader.agg_budget_monthly(_fcst_filtered)
                    _blended_chart_df = pd.concat([_blended_monthly, _budget_monthly], ignore_index=True)
                    # FY Forecast KPIs
                    _fy_forecast_revenue = _blended_df[_blended_df["Source"] == "Forecast"]["AMT"].sum()
                    _fy_forecast_gp = _blended_df[_blended_df["Source"] == "Forecast"]["GP"].sum()
                    _fy_budget_revenue = _budget_monthly["Revenue"].sum() if not _budget_monthly.empty else 0
                    # YTD Actual Revenue (up to current month)
                    _ytd_actual_revenue = _actual_df[
                        _actual_df["Ship Date"].dt.month <= _current_month
                    ]["SALES Total AMT"].sum()
                    _budget_achievement_pct = (
                        _ytd_actual_revenue / _fy_budget_revenue * 100
                    ) if _fy_budget_revenue else 0

            _label = (
                ", ".join(_targets)
                if len(_targets) <= 3
                else f"{_targets[0]} + {len(_targets)-1} others"
            )
            st.markdown(f"### {_label}")
            _dkc1, _dkc2, _dkc3, _dkc4 = st.columns(4)
            _dkc1.metric("Revenue", f"{_dk['revenue']:,.0f}")
            _dkc2.metric("GP", f"{_dk['gp']:,.0f}")
            _dkc3.metric("GP%", f"{_dk['gp_pct']:.1f}%")
            _dkc4.metric("QTY (CDR+Tablet)", f"{_dk['qty']:,.0f}")

            # FY Forecast KPIs row
            if _do_fcst and not _fcst_raw.empty and not _fcst_filtered.empty:
                st.markdown("**FY Forecast KPIs**")
                _fkc1, _fkc2, _fkc3, _fkc4 = st.columns(4)
                _fkc1.metric("FY Forecast Revenue", f"{_fy_forecast_revenue:,.0f}")
                _fkc2.metric("FY Forecast GP", f"{_fy_forecast_gp:,.0f}")
                _fkc3.metric("Budget Achievement%", f"{_budget_achievement_pct:.1f}%")
                _fkc4.metric("FY Budget Revenue", f"{_fy_budget_revenue:,.0f}")

            if not _dm.empty and not _dcat.empty:
                _ddc1, _ddc2 = st.columns(2)
                with _ddc1:
                    if not _blended_chart_df.empty:
                        st.markdown("**📈 Monthly Revenue (Actual + Forecast + Budget)**")
                        st.altair_chart(
                            chart_revenue_trend_blended(_blended_chart_df),
                            use_container_width=True,
                        )
                    else:
                        st.markdown("**📈 Monthly Revenue**")
                        st.altair_chart(
                            chart_customer_monthly(_dm),
                            use_container_width=True,
                        )
                with _ddc2:
                    st.markdown("**🍩 Category Breakdown**")
                    st.altair_chart(
                        chart_customer_cat_donut(_dcat),
                        use_container_width=True,
                    )

            # QTY by Category grouped bar
            _qty_cat = build_customer_monthly_qty_by_cat(
                dash_df[dash_df["Customer Name"].isin(_targets)]
            )
            if not _qty_cat.empty:
                st.markdown("**📦 Monthly QTY by Category**")
                st.altair_chart(
                    chart_customer_qty_by_cat(_qty_cat),
                    use_container_width=True,
                )

            if not _dm.empty:
                st.markdown("**📋 Monthly Detail**")
                st.dataframe(
                    _dm.style.format({
                        "Revenue": "{:,.0f}", "GP": "{:,.0f}",
                        "GP%": "{:.1f}", "QTY": "{:,.0f}",
                    }),
                    use_container_width=True, hide_index=True,
                )

            # Part Number detail (CDR + Tablet)
            _pn = build_pn_detail(
                dash_df[dash_df["Customer Name"].isin(_targets)],
                has_shipping,
            )
            if not _pn.empty:
                st.markdown("**📦 Part Number Detail (CDR + Tablet)**")
                _pn_fmt = {"QTY": "{:,.0f}"}
                if "Latest UP" in _pn.columns:
                    _pn_fmt["Latest UP"] = "{:,.2f}"
                st.dataframe(
                    _pn.style.format(_pn_fmt),
                    use_container_width=True, hide_index=True,
                )
        else:
            st.info("No data found for selected customer(s).")


