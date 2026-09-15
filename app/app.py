"""app.py - Performance Report Analysis Tool (v7.2)."""

import io
import os
import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st
# Ensure local modules are importable
sys.path.insert(0, str(Path(__file__).resolve().parent))

import fcst_loader

from utils import (
    DATA_DIR, CAT_ORDER, _rules_key, _normalize_name,
    HISTORICAL_CSV, scan_current_year_folder,
    load_single_file, load_historical_csv, load_overrides, save_overrides,
    override_key, override_key_series,
    build_summary, build_bycat,
    to_wide_summary, to_wide_one_cat,
    sorted_cats, fmt, show_bycat,
    cached_search_indices,
    calc_dashboard_kpis, build_monthly_trend,
    build_category_breakdown, build_monthly_category,
    build_top_customers, build_customer_detail,
    build_customer_monthly_qty_by_cat,
    build_pn_detail,
    fmt_num,
    load_settings, save_settings, DEFAULT_SETTINGS,
    inject_theme_css, resolve_theme_mode,
    get_shipped_aliases, validate_alias_mappings,
)
from charts import (
    apply_altair_theme,
    chart_up_tp_trend, chart_qty_by_year, chart_qty_by_month, chart_gp_pct_trend,
    chart_revenue_trend, chart_gp_dual_axis,
    chart_category_donut, chart_category_stacked, chart_ai_sw_revenue_trend,
    chart_top_customers_bar, chart_customer_monthly, chart_customer_cat_donut,
    chart_customer_qty_by_cat,
    chart_revenue_trend_blended, chart_gp_trend_blended, chart_qty_trend_blended,
)
from theme import kpi_card, card_title, get_tokens

st.set_page_config(
    page_title="Performance Report Analysis Tool",
    page_icon="📊",
    layout="wide",
)

if "app_settings" not in st.session_state:
    st.session_state["app_settings"] = load_settings()
_theme_setting = st.session_state["app_settings"].get("theme", "system")
inject_theme_css(_theme_setting)
_theme_mode = resolve_theme_mode(_theme_setting)
apply_altair_theme(_theme_mode)

st.title(":material/insert_chart: Performance Report Data Analysis Tool")

# -- 0. Year selection (sidebar) ---------------------------------
current_year = datetime.now().year
historical_exists = HISTORICAL_CSV.exists()
current_xlsx = scan_current_year_folder()

historical_df = None
historical_nat = 0
historical_amb = []
historical_has_des = False
historical_has_shipping = False
historical_ignored = 0
current_df = None
current_nat = 0
current_amb = []
current_has_des = False
current_has_shipping = False
current_ignored = 0

if historical_exists:
    (historical_df, historical_nat, historical_err, historical_amb,
     historical_has_des, historical_has_shipping, historical_ignored) = (
        load_historical_csv(str(HISTORICAL_CSV), _rules_key())
    )
    if historical_err:
        st.error(f"❌ {historical_err}")
        st.stop()
    historical_df = historical_df[historical_df["Ship Date"].dt.year != current_year].copy()

if current_xlsx is not None:
    (current_df, current_nat, current_err, current_amb,
     current_has_des, current_has_shipping, current_ignored) = (
        load_single_file(str(current_xlsx), _rules_key())
    )
    if current_err:
        st.error(f"❌ {current_err}")
        st.stop()
    current_df = current_df[current_df["Ship Date"].dt.year == current_year].copy()
    if current_df.empty:
        current_df = None

available_years = sorted(
    set(
        ([] if historical_df is None else historical_df["Ship Date"].dt.year.astype(int).unique().tolist())
        + ([current_year] if current_df is not None else [])
    ),
    reverse=True,
)

if not available_years:
    st.error(
        f"No data found. Please create {DATA_DIR / 'Over the Years'} and/or {DATA_DIR / 'Current Year'} "
        f"with the expected files."
    )
    st.stop()

default_year = current_year if current_year in available_years else available_years[0]

# -- 1. Load & merge ---------------------------------------------
all_dfs = []
total_nat = 0
all_ambiguous = []
global_has_des = False
global_has_shipping = False

if historical_df is not None:
    all_dfs.append(historical_df)
    total_nat += historical_nat
    all_ambiguous.extend(historical_amb)
    global_has_des = global_has_des or historical_has_des
    global_has_shipping = global_has_shipping or historical_has_shipping

if current_df is not None:
    all_dfs.append(current_df)
    total_nat += current_nat
    all_ambiguous.extend(current_amb)
    global_has_des = global_has_des or current_has_des
    global_has_shipping = global_has_shipping or current_has_shipping

if not all_dfs:
    st.error("No usable data found.")
    st.stop()

all_df = pd.concat(all_dfs, ignore_index=True)
has_des = global_has_des
has_shipping = global_has_shipping
total_ignored_rows = historical_ignored + current_ignored

# -- Overrides ---------------------------------------------------
if "others_overrides" not in st.session_state:
    st.session_state["others_overrides"] = load_overrides()

st.session_state["unmatched_overrides"] = []
if st.session_state["others_overrides"]:
    all_df = all_df.copy()
    _override_key_col = override_key_series(all_df)
    _override_mapped = _override_key_col.map(st.session_state["others_overrides"])
    _override_mask = _override_mapped.notna()
    all_df.loc[_override_mask, "Category"] = _override_mapped[_override_mask]
    _override_present_keys = set(_override_key_col)
    st.session_state["unmatched_overrides"] = [
        k for k in st.session_state["others_overrides"]
        if k not in _override_present_keys
    ]

# -- Sales Person filter (sidebar) -------------------------------
_sp_visible_custs: set[str] = set()
if "SALE_Person" in all_df.columns:
    _sp_year_rows = all_df[all_df["Ship Date"].dt.year == current_year]
    _sp_names = sorted(
        sp for sp in _sp_year_rows["SALE_Person"].dropna().unique()
        if sp not in ("nan", "NaN", "")
    )
    if _sp_names:
        st.sidebar.header(":material/person: Sales Person")
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
                _sp_selected = st.sidebar.multiselect(
                    f"Related Customers ({len(_sp_custs)})",
                    options=_sp_custs,
                    default=_sp_custs,
                    key="sp_custs_multiselect",
                )
                _sp_visible_custs = set(_sp_custs)
                for c in _sp_custs:
                    st.session_state[f"sp_cust__{c}"] = c in _sp_selected

# -- Sidebar: FCST + System Info ----------------------------------
with st.sidebar.expander(":material/trending_up: FCST", expanded=False):
    _fcst_sheet = st.radio(
        "FCST Sheet",
        options=["All Sheets", "Div.1&2_All", "VT", "Signify"],
        index=0,
        key="fcst_sheet",
    )

with st.sidebar.expander(":material/info: System Info", expanded=False):
    st.markdown(f"**Loaded {len(all_df):,} rows** ({len(available_years)} year(s) available)")
    if historical_df is not None:
        hist_mtime = datetime.fromtimestamp(HISTORICAL_CSV.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        st.markdown(
            f"**Historical**: `historical.csv`  \n<small>Modified: {hist_mtime}</small>",
            unsafe_allow_html=True,
        )
    if current_xlsx is not None:
        mtime = datetime.fromtimestamp(current_xlsx.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        st.markdown(
            f"**{current_year}**: `{current_xlsx.name}`  \n<small>Modified: {mtime}</small>",
            unsafe_allow_html=True,
        )
    if not has_des:
        st.warning("'DES' column not found; DES classification disabled.")
    if not has_shipping:
        st.warning(
            "'Currency'/'UP'/'TP(USD)' columns not found; "
            "Shipping Record Search disabled."
        )
    if "SALE_Person" not in all_df.columns:
        st.warning("'SALE_Person' column not found; Sales Person filter disabled.")
    if total_nat:
        st.warning(f"{total_nat} row(s) with invalid Ship Date skipped.")
    if all_ambiguous:
        st.warning(
            f"{len(all_ambiguous)} row(s) matched multiple DES categories. "
            "Assigned to first match."
        )
        st.dataframe(pd.DataFrame(all_ambiguous), use_container_width=True)
    if st.session_state.get("unmatched_overrides"):
        st.warning(
            f"{len(st.session_state['unmatched_overrides'])} saved category "
            "override(s) matched 0 rows in the current data (the underlying "
            "row may no longer exist)."
        )


# -- YoY comparison data (for Dashboard) -------------------------

# ------------------------------------------------------------------
# MAIN TABS
# ------------------------------------------------------------------
main_tab1, main_tab2, main_tab3, main_tab4 = st.tabs(
    [
        ":material/insert_chart: Performance Report",
        ":material/local_shipping: Shipping Record Search",
        ":material/dashboard: Company Dashboard",
        ":material/settings: Settings",
    ]
)

# -- TAB 1: Performance Report -----------------------------------
with main_tab1:
    with st.container(border=True):
        card_title("Filters", icon="filter_alt")
        _perf_col, _ = st.columns([2, 3])
        with _perf_col:
            _perf_years = st.multiselect(
                "Select years",
                options=available_years,
                default=[default_year],
                key="year_perf",
            )
        df = all_df[all_df["Ship Date"].dt.year.isin(_perf_years)].copy() if _perf_years else pd.DataFrame(columns=all_df.columns)

        st.markdown("**:material/search: Customer Name**")
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
            if st.button("Clear all selections", icon=":material/backspace:"):
                for c in all_customers:
                    st.session_state.pop(f"cust__{c}", None)
                    st.session_state.pop(f"sp_cust__{c}", None)
                st.rerun()

        qty_only = st.checkbox("QTY: sum only Tablet & CDR (exclude ACC)", value=True)
        by_cat = st.checkbox("Split report by Category", value=True)
        merge_cdr = merge_tab = False
        if by_cat:
            merge_cdr = st.checkbox("  ↪ Merge CDR ACC into CDR", value=True)
            merge_tab = st.checkbox("  ↪ Merge Tablet ACC into Tablet", value=True)

        _opts = (qty_only, by_cat, merge_cdr, merge_tab, tuple(sorted(selected)))

        if st.button("Run", icon=":material/play_arrow:", type="primary"):
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
        with st.container(border=True):
            card_title("Results", icon="insert_chart")
            if st.session_state.get("rpt_opts") != _opts:
                st.info("Options have changed; press **Run** to refresh the report.")
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

            tab_labels = [":material/summarize: Summary"]
            if not _long_bycat.empty:
                tab_labels.append(":material/donut_small: By Category")
            tabs = st.tabs(tab_labels)

            with tabs[0]:
                st.dataframe(fmt(_summary), use_container_width=True)

            if not _long_bycat.empty:
                with tabs[1]:
                    show_bycat(_long_bycat)

            st.download_button(
                "Download Excel Report",
                icon=":material/download:",
                data=_buf,
                file_name=datetime.now().strftime("sales_report_%Y%m%d_%H%M.xlsx"),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        if not _others.empty:
            with st.container(border=True):
                card_title(f"Others ({len(_others)} row(s)) — review & reassign category", icon="warning")
                _override_opts = ["Others (keep)"] + [
                    c for c in CAT_ORDER if c != "Others"
                ]
                with st.expander("Show rows"):
                    for _i, _row in _others.iterrows():
                        _ok = override_key(
                            _row["Customer Name"],
                            _row["Part Number"],
                            _row["Month"],
                            _row["DES"] if _has_des else "",
                        )
                        _c1, _c2 = st.columns([4, 1])
                        with _c1:
                            _des_str = f" | DES: {_row['DES']}" if _has_des else ""
                            _pn_display = _ok[1] if _ok[1] else "(no P/N)"
                            st.markdown(
                                f"`{_pn_display}`{_des_str}&nbsp;&nbsp;"
                                f"Month: **{_row['Month']}** | "
                                f"AMT: {int(_row['SALES Total AMT']):,}"
                            )
                        with _c2:
                            _cur = st.session_state["others_overrides"].get(
                                _ok, "Others (keep)"
                            )
                            if _cur not in _override_opts:
                                _cur = "Others (keep)"
                            _choice = st.selectbox(
                                "Reassign",
                                _override_opts,
                                index=_override_opts.index(_cur),
                                key=f"override_{_i}_{'__'.join(_ok)}",
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
                            "Overrides updated; press **Run** to apply them to the report."
                        )
                    if st.session_state.get("unmatched_overrides"):
                        _unmatched_str = "; ".join(
                            "(" + ", ".join(x if x else "∅" for x in k) + ")"
                            for k in st.session_state["unmatched_overrides"]
                        )
                        st.warning(
                            f"{len(st.session_state['unmatched_overrides'])} saved "
                            "override(s) matched **0 rows** in the current data (the "
                            f"underlying row may no longer exist): {_unmatched_str}"
                        )

# -- TAB 2: Shipping Record Search -------------------------------
with main_tab2:
    with st.container(border=True):
        card_title("Search", icon="local_shipping")
        _ship_col, _ = st.columns([2, 3])
        with _ship_col:
            _ship_years = st.multiselect(
                "Select years",
                options=available_years,
                default=[default_year],
                key="year_shipping",
            )
        df = all_df[all_df["Ship Date"].dt.year.isin(_ship_years)].copy() if _ship_years else pd.DataFrame(columns=all_df.columns)

        if not has_shipping:
            st.warning(
                "Shipping Record Search requires **Currency**, **UP**, and **TP(USD)** "
                "columns in the data files. These columns were not found."
            )
            pn_query = ""
        else:
            if "search_history" not in st.session_state:
                st.session_state["search_history"] = []
            if st.session_state["search_history"]:
                st.caption(":material/history: Recent searches")
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

                # -- Cached search -----------------------------------------------
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
                    # -- Save to search history --------------------------------------
                    _q = pn_query.strip()
                    _hist = st.session_state["search_history"]
                    if _q in _hist:
                        _hist.remove(_q)
                    _hist.insert(0, _q)
                    st.session_state["search_history"] = _hist[:10]

                    # -- Part Number selection (narrow down) -------------------------
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

                    # -- Customer filter (optional) ----------------------------------
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

    if has_shipping and pn_query.strip() and not results.empty:
        # -- QTY TTL & weighted averages ---------------------------------
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

        # -- KPI row -------------------------------------------------------
        with st.container(border=True):
            card_title("Summary", icon="summarize")
            _mc1, _mc2, _mc3, _mc4 = st.columns(4)
            with _mc1:
                kpi_card("Total QTY", f"{_total_qty:,}")
            with _mc2:
                kpi_card("Avg UP (wavg)", f"{_wavg_up:,.2f}")
            with _mc3:
                kpi_card("Avg TP(USD) (wavg)", f"{_wavg_tp:,.2f}")
            with _mc4:
                kpi_card("GP%", _wavg_gp_pct)

        # -- Charts (3 columns, all Altair) ------------------------------
        with st.container(border=True):
            card_title("Trends", icon="trending_up")
            _cc1, _cc2, _cc3 = st.columns(3)
            with _cc1:
                card_title("UP & TP(USD) Monthly Trend")
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
                    card_title("QTY by Year")
                    st.altair_chart(
                        chart_qty_by_year(results),
                        use_container_width=True,
                    )
                else:
                    card_title("QTY by Month")
                    st.altair_chart(
                        chart_qty_by_month(results),
                        use_container_width=True,
                    )
            with _cc3:
                card_title("GP% Monthly Trend")
                st.altair_chart(
                    chart_gp_pct_trend(results, mode=_theme_mode),
                    use_container_width=True,
                )

        # -- GP% conditional formatting helper ---------------------------
        _gp_pos_color = get_tokens(_theme_mode)["positive"]
        _gp_neg_color = get_tokens(_theme_mode)["negative"]

        def _color_gp_pct(val):
            if isinstance(val, str) and "%" in val:
                try:
                    n = float(val.replace("%", ""))
                    if n >= 30:
                        return f"color: {_gp_pos_color}; font-weight: bold"
                    elif n < 15:
                        return f"color: {_gp_neg_color}; font-weight: bold"
                except ValueError:
                    pass
            return ""

        # -- Format display ----------------------------------------------
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

        with st.container(border=True):
            card_title("Results", icon="insert_chart")
            st.dataframe(styled, use_container_width=True)

            # -- Part Number summary table (when customer filter active) -----
            if _cust_filter_active:
                card_title("Part Number Summary")
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

# -- TAB 3: Company Dashboard ------------------------------------
with main_tab3:
    # -- Data prep (apply SALE_Person filter) ------------------------
    _dash_col, _ = st.columns([2, 3])
    with _dash_col:
        _dash_years = st.multiselect(
            "Select years",
            options=available_years,
            default=[default_year],
            key="year_dashboard",
        )

    # YoY data: year before the max selected year
    _dash_max_yr = max(_dash_years) if _dash_years else current_year
    _yoy_year = _dash_max_yr - 1
    yoy_df = (
        all_df[all_df["Ship Date"].dt.year == _yoy_year].copy()
        if _yoy_year in all_df["Ship Date"].dt.year.unique()
        else None
    )

    # Data prep: filter all_df to selected years, apply SALE_Person filter
    dash_df = all_df[all_df["Ship Date"].dt.year.isin(_dash_years)].copy() if _dash_years else pd.DataFrame(columns=all_df.columns)
    dash_yoy = yoy_df.copy() if yoy_df is not None else None
    if _sp_visible_custs:
        dash_df = dash_df[dash_df["Customer Name"].isin(_sp_visible_custs)]
        if dash_yoy is not None:
            dash_yoy = dash_yoy[
                dash_yoy["Customer Name"].isin(_sp_visible_custs)
            ]

    _dash_curr = dash_df[dash_df["Ship Date"].dt.year == _dash_max_yr]
    _kpis_yoy = calc_dashboard_kpis(_dash_curr, dash_yoy)
    _kpis_all = calc_dashboard_kpis(dash_df)

    # -- FCST Integration ----------------------------------------------
    _now = datetime.now()
    _current_month = _now.month
    _current_yr = _now.year
    # Only blend when the current calendar year is in the selection
    _do_fcst = _dash_max_yr == _current_yr

    _fcst_raw = pd.DataFrame()
    _blended_raw = pd.DataFrame()
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
        st.warning(
            f":material/warning: FCST 未匹配客戶 ({len(_unmatched_fcst)} 個) — "
            f"請至 **:material/settings: Settings ▸ Account Match** 指定對應的 Performance Report 客戶或自訂群組。"
        )
        with st.expander(f":material/search: 點擊查看未匹配客戶詳情 ({len(_unmatched_fcst)} 個)", expanded=False):
            _unmatched_df = pd.DataFrame(
                sorted(_unmatched_fcst),
                columns=["Customer Name", "Sheet"],
            )
            st.dataframe(_unmatched_df, hide_index=True, use_container_width=True)

    def _fmt_delta(val, suffix="%"):
        if val is None:
            return "N/A"
        return f"{val:+,.1f}{suffix}"

    # -- Prep for the first-screen [2, 1] row -------------------------
    _trend = build_monthly_trend(dash_df)
    _multi_yr = len(_dash_years) > 1
    _chart_data_blended = _blended_monthly.copy() if not _blended_monthly.empty else pd.DataFrame()
    if not _budget_monthly.empty and not _chart_data_blended.empty:
        _chart_data_blended = pd.concat(
            [_chart_data_blended, _budget_monthly],
            ignore_index=True
        )
    _top_preview = build_top_customers(dash_df, 5, dash_yoy)

    # -- First screen: [2, 1] trend + KPI/Top Customers summary -------
    _fs_left, _fs_right = st.columns([2, 1])
    with _fs_left:
        with st.container(border=True):
            card_title("Monthly Trend", icon="trending_up")
            _dash_trend_metric = st.radio(
                "Metric",
                ["Revenue", "QTY"],
                horizontal=True,
                key="dash_trend_metric",
                label_visibility="collapsed",
            )
            if _dash_trend_metric == "Revenue":
                if not _chart_data_blended.empty:
                    st.altair_chart(
                        chart_revenue_trend_blended(_chart_data_blended, mode=_theme_mode),
                        use_container_width=True,
                    )
                else:
                    st.altair_chart(
                        chart_revenue_trend(_trend, multi_year=_multi_yr),
                        use_container_width=True,
                    )
            else:
                if not _blended_monthly.empty:
                    st.altair_chart(
                        chart_qty_trend_blended(_blended_monthly, mode=_theme_mode),
                        use_container_width=True,
                    )
                else:
                    st.altair_chart(
                        chart_qty_by_month(dash_df),
                        use_container_width=True,
                    )
    with _fs_right:
        with st.container(border=True):
            card_title("Overview", icon="summarize")
            if dash_yoy is not None:
                st.caption(f"YoY delta: {_dash_max_yr} vs {_dash_max_yr - 1}")
            kpi_card("Revenue (TWD)", fmt_num(_kpis_all['revenue']),
                     delta=_fmt_delta(_kpis_yoy["revenue_yoy"]))
            kpi_card("Gross Profit (TWD)", fmt_num(_kpis_all['gp']),
                     delta=_fmt_delta(_kpis_yoy["gp_yoy"]))
            kpi_card("GP Margin", f"{_kpis_all['gp_pct']:.1f}%",
                     delta=_fmt_delta(_kpis_yoy["gp_pct_yoy"], " ppt"))
            kpi_card("Units Sold", f"{_kpis_all['qty']:,.0f}",
                     delta=_fmt_delta(_kpis_yoy["qty_yoy"]),
                     caption="CDR + Tablet only")
        with st.container(border=True):
            card_title("Top Customers", icon="leaderboard")
            st.dataframe(
                _top_preview[["Customer Name", "Revenue"]].style.format({"Revenue": "{:,.0f}"}),
                use_container_width=True, hide_index=True,
            )

    # -- Full-Year Forecast (only shown when current year is selected and FCST loaded)
    if not _blended_monthly.empty:
        _ytd_rev = _blended_monthly[_blended_monthly["Source"] == "Actual"]["Revenue"].sum()
        _ytd_gp = _blended_monthly[_blended_monthly["Source"] == "Actual"]["GP"].sum()
        _fy_rev = _blended_monthly["Revenue"].sum()
        _fy_gp = _blended_monthly["GP"].sum()
        _fy_gp_pct = _fy_gp / _fy_rev * 100 if _fy_rev else 0.0
        _fy_qty = _blended_monthly["QTY"].sum()

        with st.container(border=True):
            card_title("Full-Year Forecast", icon="trending_up")
            st.caption(f"YTD Actual + Remaining FCST — sheet: **{_fcst_sheet}**")
            _fk1, _fk2, _fk3, _fk4 = st.columns(4)
            with _fk1:
                kpi_card("FY Revenue Forecast (TWD)", fmt_num(_fy_rev), caption=f"YTD: {fmt_num(_ytd_rev)}")
            with _fk2:
                kpi_card("FY Gross Profit Forecast (TWD)", fmt_num(_fy_gp), caption=f"YTD: {fmt_num(_ytd_gp)}")
            with _fk3:
                kpi_card("FY GP Margin Forecast", f"{_fy_gp_pct:.1f}%")
            with _fk4:
                kpi_card("FY Units Forecast", f"{_fy_qty:,.0f}")

        # Budget Achievement & PO Coverage
        try:
            _budget_monthly = fcst_loader.agg_budget_monthly(_fcst_raw)
            _po_cov = fcst_loader.agg_po_coverage(_fcst_raw)
            _has_budget = not _budget_monthly.empty
            if _has_budget:
                _fy_budget_rev = _budget_monthly["Revenue"].sum()
                _budget_achievement_pct = (
                    _ytd_rev / _fy_budget_rev * 100 if _fy_budget_rev else 0.0
                )
            _po_pct = _po_cov["po_coverage_pct"]
            if _has_budget or _po_pct is not None:
                with st.container(border=True):
                    card_title("Budget Achievement & PO Coverage", icon="flag")
                    _bk1, _bk2, _bk3 = st.columns(3)
                    if _has_budget:
                        with _bk1:
                            kpi_card("Budget Achievement", f"{_budget_achievement_pct:.1f}%", caption="YTD vs FY Budget")
                        with _bk2:
                            kpi_card("FY Budget Revenue (TWD)", fmt_num(_fy_budget_rev), caption=f"Current: {fmt_num(_ytd_rev)}")
                    with _bk3:
                        kpi_card("PO Coverage%", f"{_po_pct:.1f}%" if _po_pct is not None else "-",
                                 caption="PO / Forecast (FY, AMT)")
        except Exception as _budget_err:
            print(f"[Dashboard] Warning: Failed to calculate Budget Achievement / PO Coverage: {_budget_err}")

    # -- Monthly Trends (GP & GP% — Revenue/QTY trend is in the first screen above)
    with st.container(border=True):
        card_title("Monthly Trends", icon="trending_up")
        if not _chart_data_blended.empty:
            st.altair_chart(
                chart_gp_trend_blended(_chart_data_blended, mode=_theme_mode),
                use_container_width=True,
            )
        else:
            st.altair_chart(
                chart_gp_dual_axis(_trend, mode=_theme_mode),
                use_container_width=True,
            )

    # -- Category Analysis ---------------------------------------------
    with st.container(border=True):
        card_title("Category Analysis", icon="donut_small")
        _cat_br = build_category_breakdown(dash_df)
        _cat_mo = build_monthly_category(dash_df)
        _cat_row1_c1, _cat_row1_c2 = st.columns(2)
        with _cat_row1_c1:
            card_title("Revenue by Category")
            st.altair_chart(
                chart_category_donut(_cat_br, mode=_theme_mode), use_container_width=True,
            )
        with _cat_row1_c2:
            card_title("Category Revenue Trend")
            st.altair_chart(
                chart_category_stacked(_cat_mo, mode=_theme_mode), use_container_width=True,
            )
        if not _fcst_cat_monthly.empty:
            _cat_row2_c1, _cat_row2_c2 = st.columns(2)
        else:
            _cat_row2_c1 = st.columns(1)[0]
        with _cat_row2_c1:
            card_title("AI_SW Monthly Revenue Trend")
            st.altair_chart(
                chart_ai_sw_revenue_trend(_cat_mo, mode=_theme_mode), use_container_width=True,
            )
        if not _fcst_cat_monthly.empty:
            with _cat_row2_c2:
                card_title("FCST Category Revenue")
                _fcst_cat_display = _fcst_cat_monthly.rename(
                    columns={"Cat": "Category", "Period": "Month"}
                )
                st.altair_chart(
                    chart_category_stacked(_fcst_cat_display, mode=_theme_mode), use_container_width=True,
                )

    # -- Top Customers ---------------------------------------------------
    with st.container(border=True):
        card_title("Top Customers", icon="leaderboard")
        _top_n = st.slider(
            "Number of customers", 5, 30, 10, key="dash_top_n",
        )
        _fcst_blended_for_top = _blended_raw if (_do_fcst and not _blended_raw.empty) else None
        _top = build_top_customers(dash_df, _top_n, dash_yoy, fcst_df=_fcst_blended_for_top)
        _tn1, _tn2 = st.columns(2)
        with _tn1:
            card_title(f"Top {_top_n} Customers by Revenue")
            st.altair_chart(
                chart_top_customers_bar(_top, mode=_theme_mode), use_container_width=True,
            )
        with _tn2:
            card_title(f"Top {_top_n} Customers")

            _gp_pos_color = get_tokens(_theme_mode)["positive"]
            _gp_neg_color = get_tokens(_theme_mode)["negative"]

            def _gp_color(val):
                try:
                    n = float(val)
                    if n >= 30:
                        return f"color: {_gp_pos_color}; font-weight: bold"
                    if n < 15:
                        return f"color: {_gp_neg_color}; font-weight: bold"
                except (ValueError, TypeError):
                    pass
                return ""

            def _achievement_color(val):
                try:
                    n = float(val)
                    if n >= 80:
                        return f"color: {_gp_pos_color}; font-weight: bold"
                    if n < 50:
                        return f"color: {_gp_neg_color}; font-weight: bold"
                except (ValueError, TypeError):
                    pass
                return ""

            _fmt = {"Revenue": "{:,.0f}", "GP": "{:,.0f}",
                    "GP%": "{:.1f}", "QTY": "{:,.0f}", "YoY%": "{:+.1f}"}
            _style_subsets = [("GP%", _gp_color)]
            if "FY Forecast" in _top.columns:
                _fmt["FY Forecast"] = "{:,.0f}"
                _fmt["Achievement%"] = "{:.1f}"
                _style_subsets.append(("Achievement%", _achievement_color))

            _styled = _top.style.format(_fmt, na_rep="-")
            for _col, _fn in _style_subsets:
                _styled = _styled.map(_fn, subset=[_col])

            st.dataframe(_styled, use_container_width=True)

    # -- Customer Drill-Down ---------------------------------------------
    with st.container(border=True):
        card_title("Customer Drill-Down", icon="search")
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
                _cust_po_cov = {"po_total": 0, "forecast_total": 0, "po_coverage_pct": None}
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
                        _fy_forecast_revenue = _blended_df["AMT"].sum()
                        _fy_forecast_gp = _blended_df["GP"].sum()
                        _fy_budget_revenue = _budget_monthly["Revenue"].sum() if not _budget_monthly.empty else 0
                        # YTD Actual Revenue (up to current month)
                        _ytd_actual_revenue = _actual_df[
                            _actual_df["Ship Date"].dt.month <= _current_month
                        ]["SALES Total AMT"].sum()
                        _budget_achievement_pct = (
                            _ytd_actual_revenue / _fy_budget_revenue * 100
                        ) if _fy_budget_revenue else 0
                        _cust_po_cov = fcst_loader.agg_po_coverage(_fcst_filtered)

                _label = (
                    ", ".join(_targets)
                    if len(_targets) <= 3
                    else f"{_targets[0]} + {len(_targets)-1} others"
                )
                card_title(_label)

                # FY Forecast KPIs row
                if _do_fcst and not _fcst_raw.empty and not _fcst_filtered.empty:
                    card_title("FY Forecast KPIs")
                    _fkc1, _fkc2, _fkc3, _fkc4, _fkc5 = st.columns(5)
                    _rev_diff = _fy_forecast_revenue - _fy_budget_revenue
                    with _fkc1:
                        kpi_card("FY Forecast Revenue (TWD)", fmt_num(_fy_forecast_revenue),
                                 delta=('+' if _rev_diff >= 0 else '') + fmt_num(_rev_diff))
                    with _fkc2:
                        kpi_card("FY Forecast Gross Profit (TWD)", fmt_num(_fy_forecast_gp))
                    with _fkc3:
                        kpi_card("Budget Achievement", f"{_budget_achievement_pct:.1f}%",
                                 delta=f"{_budget_achievement_pct - 100:+.1f}%")
                    _ytd_diff = _ytd_actual_revenue - _fy_budget_revenue
                    with _fkc4:
                        kpi_card("FY Budget Revenue (TWD)", fmt_num(_fy_budget_revenue),
                                 delta=('+' if _ytd_diff >= 0 else '') + fmt_num(_ytd_diff))
                    _cust_po_pct = _cust_po_cov["po_coverage_pct"]
                    with _fkc5:
                        kpi_card("PO Coverage%",
                                 f"{_cust_po_pct:.1f}%" if _cust_po_pct is not None else "-",
                                 caption="PO / Forecast (FY, AMT)")

                _dkc1, _dkc2, _dkc3, _dkc4 = st.columns(4)
                with _dkc1:
                    kpi_card("Revenue (TWD)", fmt_num(_dk['revenue']))
                with _dkc2:
                    kpi_card("Gross Profit (TWD)", fmt_num(_dk['gp']))
                with _dkc3:
                    kpi_card("GP Margin", f"{_dk['gp_pct']:.1f}%")
                with _dkc4:
                    kpi_card("Units Sold", f"{_dk['qty']:,.0f}", caption="CDR + Tablet only")

                if not _dm.empty and not _dcat.empty:
                    _ddc1, _ddc2 = st.columns(2)
                    with _ddc1:
                        if not _blended_chart_df.empty:
                            card_title("Monthly Revenue (Actual + Forecast + Budget)")
                            st.altair_chart(
                                chart_revenue_trend_blended(_blended_chart_df, mode=_theme_mode),
                                use_container_width=True,
                            )
                        else:
                            card_title("Monthly Revenue")
                            st.altair_chart(
                                chart_customer_monthly(_dm, mode=_theme_mode),
                                use_container_width=True,
                            )
                    with _ddc2:
                        card_title("Category Breakdown")
                        st.altair_chart(
                            chart_customer_cat_donut(_dcat, mode=_theme_mode),
                            use_container_width=True,
                        )

                # QTY by Category grouped bar
                _qty_cat = build_customer_monthly_qty_by_cat(
                    dash_df[dash_df["Customer Name"].isin(_targets)]
                )
                if not _qty_cat.empty:
                    card_title("Monthly QTY by Category")
                    st.altair_chart(
                        chart_customer_qty_by_cat(_qty_cat, mode=_theme_mode),
                        use_container_width=True,
                    )

                if not _dm.empty:
                    card_title("Monthly Detail")
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
                    card_title("Part Number Detail (CDR + Tablet)")
                    _pn_fmt = {"QTY": "{:,.0f}"}
                    if "Latest UP" in _pn.columns:
                        _pn_fmt["Latest UP"] = "{:,.2f}"
                    st.dataframe(
                        _pn.style.format(_pn_fmt),
                        use_container_width=True, hide_index=True,
                    )
            else:
                st.info("No data found for selected customer(s).")

# -- TAB 4: Settings -----------------------------------------------
with main_tab4:
    st.header(":material/settings: Settings")
    st.caption(
        "Changes here are saved to `app/settings.json` immediately and "
        "survive restarts and version upgrades."
    )

    if "app_settings" not in st.session_state:
        st.session_state["app_settings"] = load_settings()
    _set_settings = st.session_state["app_settings"]

    def _set_persist(new_settings):
        save_settings(new_settings)
        st.session_state["app_settings"] = load_settings()
        st.rerun()

    # A st.radio (not st.tabs) is used for this sub-navigation: every edit in
    # Section B/C calls st.rerun() to apply immediately, and st.tabs() resets
    # to its first item on a programmatic rerun — a plain widget key survives it.
    _set_section = st.radio(
        "Settings section",
        [
            ":material/palette: Theme",
            ":material/block: Customer Ignore List",
            ":material/link: Account Match",
        ],
        horizontal=True, label_visibility="collapsed", key="settings_section",
    )

    # -- Section A: Theme --------------------------------------------
    if _set_section == ":material/palette: Theme":
        with st.container(border=True):
            card_title("Theme", icon="palette")
            _set_theme_labels = ["Light", "Dark", "System"]
            _set_theme_to_label = {"light": "Light", "dark": "Dark", "system": "System"}
            _set_label_to_theme = {v: k for k, v in _set_theme_to_label.items()}
            _set_cur_label = _set_theme_to_label.get(_set_settings.get("theme", "system"), "System")
            _set_new_label = st.radio(
                "App theme", _set_theme_labels,
                index=_set_theme_labels.index(_set_cur_label),
                horizontal=True, key="settings_theme_radio",
            )
            _set_new_theme = _set_label_to_theme[_set_new_label]
            if _set_new_theme != _set_settings.get("theme", "system"):
                _set_settings["theme"] = _set_new_theme
                _set_persist(_set_settings)
            if _set_new_theme in ("light", "dark"):
                st.caption(
                    ":material/info: Applied immediately for this session. Restart the app for "
                    "full native theming fidelity (some Streamlit chrome updates only after restart)."
                )
            if st.button("Reset Theme to Default", icon=":material/restart_alt:", key="settings_reset_theme"):
                _set_settings["theme"] = DEFAULT_SETTINGS["theme"]
                _set_persist(_set_settings)

    # -- Section B: Customer Ignore List -------------------------------
    if _set_section == ":material/block: Customer Ignore List":
        with st.container(border=True):
            card_title("Customer Ignore List", icon="block")
            st.caption(
                "Rows whose Customer Name (normalized) matches an entry below are "
                "dropped entirely — from every KPI, chart, drill-down, and export, "
                "in both the Performance Report and FCST pipelines."
            )
            _set_ignored = list(_set_settings.get("ignored_customers", []))
            _set_known_custs = (
                sorted(all_df["Customer Name"].dropna().unique().tolist())
                if not all_df.empty else []
            )

            _set_add_col1, _set_add_col2 = st.columns([3, 1])
            with _set_add_col1:
                _set_add_ms = st.multiselect(
                    "Add customer(s) from current data",
                    options=[c for c in _set_known_custs if c not in _set_ignored],
                    key="settings_ignore_add_ms",
                )
                _set_add_free = st.text_input(
                    "Or type a customer name not present in current data",
                    key="settings_ignore_add_free",
                )
            with _set_add_col2:
                st.write("")
                st.write("")
                if st.button("Add to ignore list", icon=":material/add:", key="settings_ignore_add_btn"):
                    _set_to_add = list(_set_add_ms)
                    if _set_add_free.strip():
                        _set_to_add.append(_set_add_free.strip())
                    _set_existing_norms = {_normalize_name(x, upper=True) for x in _set_ignored}
                    _set_added_any = False
                    for _set_name in _set_to_add:
                        _set_norm = _normalize_name(_set_name, upper=True)
                        if _set_norm and _set_norm not in _set_existing_norms:
                            _set_ignored.append(_set_name.strip())
                            _set_existing_norms.add(_set_norm)
                            _set_added_any = True
                    if _set_added_any:
                        _set_settings["ignored_customers"] = _set_ignored
                        _set_persist(_set_settings)
                    else:
                        st.warning("Nothing to add — select or type a customer name first.")

            if _set_ignored:
                st.markdown(f"**Currently ignored ({len(_set_ignored)}):**")
                for _set_name in list(_set_ignored):
                    _set_r1, _set_r2 = st.columns([5, 1])
                    _set_r1.write(f"`{_set_name}`")
                    if _set_r2.button("Remove", icon=":material/close:", key=f"settings_ignore_rm_{_set_name}"):
                        _set_ignored.remove(_set_name)
                        _set_settings["ignored_customers"] = _set_ignored
                        _set_persist(_set_settings)
            else:
                st.caption("No customers are currently ignored.")

            _set_fcst_ignored = fcst_loader.get_ignored_row_count() if _do_fcst else 0
            kpi_card(
                "Rows excluded by ignore list (current load)",
                f"{total_ignored_rows + _set_fcst_ignored:,}",
                caption=f"Performance Report: {total_ignored_rows:,} row(s)  ·  "
                        f"FCST: {_set_fcst_ignored:,} row(s)",
            )

            if st.button("Reset Ignore List to Default", icon=":material/restart_alt:", key="settings_reset_ignore"):
                _set_settings["ignored_customers"] = list(DEFAULT_SETTINGS["ignored_customers"])
                _set_persist(_set_settings)

    # -- Section C: Account Match (FCST <-> Performance Report) -------
    if _set_section == ":material/link: Account Match":
        with st.container(border=True):
            card_title("Account Match (FCST ↔ Performance Report)", icon="link")
            st.caption(
                "The Performance Report customer name is the source of truth. "
                "Map each FCST-only customer name to an existing Performance Report "
                "customer, or to a custom group bucket."
            )
            _set_known_pr = (
                sorted(all_df["Customer Name"].dropna().unique().tolist())
                if not all_df.empty else []
            )
            _set_groups = list(_set_settings.get("custom_groups", []))
            _set_fcst_aliases = dict(_set_settings.get("fcst_customer_aliases", {}))
            _set_cust_aliases = dict(_set_settings.get("customer_aliases", {}))

            card_title("Needs Mapping", icon="warning")
            if not _do_fcst:
                st.info(
                    "Select the current calendar year on the **:material/dashboard: Company Dashboard** "
                    "tab to load FCST data and see unmatched customers here."
                )
            else:
                _set_unmatched = sorted(fcst_loader.get_unmatched_customers())
                _set_unmatched_amts = fcst_loader.get_unmatched_customer_amounts()
                if not _set_unmatched:
                    st.caption("No unmatched FCST customers in the current load.")
                else:
                    _set_targets = (
                        ["Others (default)"] + _set_known_pr
                        + [f"[Group] {g}" for g in _set_groups]
                    )
                    _hc1, _hc2, _hc3, _hc4 = st.columns([3, 2, 2, 3])
                    _hc1.markdown("**FCST Name**")
                    _hc2.markdown("**Sheet**")
                    _hc3.markdown("**Amount**")
                    _hc4.markdown("**Assign To**")
                    for _set_fname, _set_sheet in _set_unmatched:
                        _set_amt = _set_unmatched_amts.get((_set_fname, _set_sheet), 0.0)
                        _c1, _c2, _c3, _c4 = st.columns([3, 2, 2, 3])
                        _c1.write(f"`{_set_fname}`")
                        _c2.write(_set_sheet)
                        _c3.write(fmt_num(_set_amt))
                        _set_existing = _set_fcst_aliases.get(_set_fname, "")
                        if not _set_existing:
                            _set_existing_label = "Others (default)"
                        elif _set_existing in _set_groups:
                            _set_existing_label = f"[Group] {_set_existing}"
                        else:
                            _set_existing_label = _set_existing
                        _set_targets_local = (
                            _set_targets if _set_existing_label in _set_targets
                            else _set_targets + [_set_existing_label]
                        )
                        _set_choice = _c4.selectbox(
                            "Assign", _set_targets_local,
                            index=_set_targets_local.index(_set_existing_label),
                            key=f"settings_fcst_map_{_set_fname}_{_set_sheet}",
                            label_visibility="collapsed",
                        )
                        if _set_choice != _set_existing_label:
                            if _set_choice == "Others (default)":
                                _set_fcst_aliases.pop(_set_fname, None)
                            elif _set_choice.startswith("[Group] "):
                                _set_fcst_aliases[_set_fname] = _set_choice[len("[Group] "):]
                            else:
                                _set_fcst_aliases[_set_fname] = _set_choice
                            _set_settings["fcst_customer_aliases"] = _set_fcst_aliases
                            _set_persist(_set_settings)

            card_title("Custom Groups")
            st.caption(
                "Additional 'Others'-style buckets (e.g. 'Others - EMEA Distributors') "
                "selectable anywhere a mapping target is chosen."
            )
            _set_gcol1, _set_gcol2 = st.columns([3, 1])
            with _set_gcol1:
                _set_new_group = st.text_input(
                    "New group name", key="settings_new_group",
                    placeholder="Others - EMEA Distributors",
                )
            with _set_gcol2:
                st.write("")
                st.write("")
                if st.button("Create Group", icon=":material/add:", key="settings_add_group"):
                    _set_gname = _set_new_group.strip()
                    if not _set_gname:
                        st.warning("Group name cannot be empty.")
                    elif _set_gname in _set_groups:
                        st.warning("Group already exists.")
                    else:
                        _set_groups.append(_set_gname)
                        _set_settings["custom_groups"] = _set_groups
                        _set_persist(_set_settings)

            for _set_g in list(_set_groups):
                _gc1, _gc2, _gc3 = st.columns([3, 1, 1])
                _set_rename_val = _gc1.text_input(
                    f"Rename '{_set_g}'", value=_set_g,
                    key=f"settings_group_rename_{_set_g}",
                    label_visibility="collapsed",
                )
                if _gc2.button("Save", icon=":material/save:", key=f"settings_group_save_{_set_g}"):
                    _set_new_name = _set_rename_val.strip()
                    if _set_new_name and _set_new_name != _set_g:
                        _set_groups[_set_groups.index(_set_g)] = _set_new_name
                        for _set_d in (_set_fcst_aliases, _set_cust_aliases):
                            for _set_k, _set_v in list(_set_d.items()):
                                if _set_v == _set_g:
                                    _set_d[_set_k] = _set_new_name
                        _set_settings["custom_groups"] = _set_groups
                        _set_settings["fcst_customer_aliases"] = _set_fcst_aliases
                        _set_settings["customer_aliases"] = _set_cust_aliases
                        _set_persist(_set_settings)
                if _gc3.button("Delete", icon=":material/delete:", key=f"settings_group_del_{_set_g}"):
                    _set_groups.remove(_set_g)
                    for _set_d in (_set_fcst_aliases, _set_cust_aliases):
                        for _set_k, _set_v in list(_set_d.items()):
                            if _set_v == _set_g:
                                del _set_d[_set_k]
                    _set_settings["custom_groups"] = _set_groups
                    _set_settings["fcst_customer_aliases"] = _set_fcst_aliases
                    _set_settings["customer_aliases"] = _set_cust_aliases
                    _set_persist(_set_settings)
            if not _set_groups:
                st.caption("No custom groups yet.")

            card_title("Shipping Record Customer Name Aliases")
            st.caption(
                "Overrides/additions to aliases.json's 'customer' section "
                "(used to normalize Customer Name when loading Performance Report data)."
            )
            with st.expander("View shipped defaults (aliases.json, read-only)"):
                st.json(get_shipped_aliases("customer"))
            _set_akcol1, _set_akcol2, _set_akcol3 = st.columns([2, 2, 1])
            with _set_akcol1:
                _set_new_alias_key = st.text_input(
                    "Customer name (as it appears in source file)",
                    key="settings_new_alias_key",
                )
            with _set_akcol2:
                _set_new_alias_val = st.text_input(
                    "Normalize to", key="settings_new_alias_val",
                )
            with _set_akcol3:
                st.write("")
                st.write("")
                if st.button("Add / Update", icon=":material/add:", key="settings_add_alias"):
                    if _set_new_alias_key.strip() and _set_new_alias_val.strip():
                        _set_norm_key = _normalize_name(_set_new_alias_key, upper=True)
                        _set_cust_aliases[_set_norm_key] = _set_new_alias_val.strip()
                        _set_settings["customer_aliases"] = _set_cust_aliases
                        _set_persist(_set_settings)
                    else:
                        st.warning("Both fields are required.")

            if _set_cust_aliases:
                for _set_k, _set_v in list(_set_cust_aliases.items()):
                    _ac1, _ac2, _ac3 = st.columns([3, 3, 1])
                    _ac1.write(f"`{_set_k}`")
                    _ac2.write(f"→ `{_set_v}`")
                    if _ac3.button("Remove", icon=":material/close:", key=f"settings_alias_rm_{_set_k}"):
                        del _set_cust_aliases[_set_k]
                        _set_settings["customer_aliases"] = _set_cust_aliases
                        _set_persist(_set_settings)
            else:
                st.caption("No custom Shipping Record aliases yet.")

            _set_valid_targets = set(_set_known_pr) | set(_set_groups) | {"Signify"}
            _set_warnings = validate_alias_mappings(_set_cust_aliases, _set_valid_targets)
            _set_warnings += validate_alias_mappings(_set_fcst_aliases, _set_valid_targets)
            if _set_warnings:
                st.warning(":material/warning: Mapping validation:\n" + "\n".join(f"- {w}" for w in _set_warnings))

            if st.button("Reset Account Match to Default", icon=":material/restart_alt:", key="settings_reset_match"):
                _set_settings["customer_aliases"] = {}
                _set_settings["fcst_customer_aliases"] = {}
                _set_settings["custom_groups"] = []
                _set_persist(_set_settings)

    if st.button("Reset ALL Settings to Default", icon=":material/restart_alt:", key="settings_reset_all", type="primary"):
        st.session_state["settings_confirm_reset_all"] = True
    if st.session_state.get("settings_confirm_reset_all"):
        st.warning("This resets Theme, Ignore List, and Account Match to their shipped defaults.")
        _rc1, _rc2 = st.columns(2)
        if _rc1.button("Confirm Reset All", icon=":material/check:", key="settings_confirm_reset_all_btn"):
            save_settings(dict(DEFAULT_SETTINGS))
            st.session_state["app_settings"] = load_settings()
            st.session_state["settings_confirm_reset_all"] = False
            st.rerun()
        if _rc2.button("Cancel", key="settings_cancel_reset_all"):
            st.session_state["settings_confirm_reset_all"] = False
            st.rerun()
