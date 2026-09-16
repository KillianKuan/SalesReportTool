"""charts.py — Altair chart builders for Shipping Record Search & Dashboard tabs.

Mark colors come from palette.py (CATEGORY_COLORS / SOURCE_COLORS / POSITIVE /
NEGATIVE / PRIMARY / MUTED) — a single fixed set, not a Light/Dark palette.
Everything else (axis/legend/title text color, chart background) is left to
Streamlit's own automatic Altair theme (``st.altair_chart``'s default
``theme="streamlit"``), which already adapts to the active light/dark theme —
so charts here never register or enable a competing Altair theme.
"""

import altair as alt
import pandas as pd

from palette import CATEGORY_COLORS, SOURCE_COLORS, MUTED, NEGATIVE, PRIMARY

_MONTHS_ORDER = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]

# Vega expression: abbreviate numeric axis labels (1.2M / 340K); pass
# through non-numeric (nominal/ordinal) labels unchanged.
_NUM_LABEL_EXPR = (
    "isNumber(datum.value) ? "
    "(abs(datum.value) >= 1e9 ? format(datum.value / 1e9, '.2f') + 'B' : "
    "abs(datum.value) >= 1e6 ? format(datum.value / 1e6, '.1f') + 'M' : "
    "abs(datum.value) >= 1e3 ? format(datum.value / 1e3, '.0f') + 'K' : "
    "format(datum.value, ',')) : datum.value"
)


def chart_up_tp_trend(results: pd.DataFrame) -> alt.LayerChart:
    """UP & TP(USD) monthly average trend — line chart with tooltips & point markers."""
    m = results.copy()
    m["Month"] = m["Ship Date"].dt.to_period("M").astype(str)
    avg = m.groupby("Month", sort=True)[["UP", "TP(USD)"]].mean().reset_index()
    long = avg.melt("Month", var_name="Metric", value_name="Price")

    return (
        alt.Chart(long)
        .mark_line(point=alt.OverlayMarkDef(size=30))
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Price:Q", title="Price", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
            color=alt.Color("Metric:N"),
            tooltip=["Month:N", "Metric:N", alt.Tooltip("Price:Q", format=",.2f")],
        )
    )


def chart_qty_by_year(results: pd.DataFrame) -> alt.LayerChart:
    """QTY sum by year — bar chart with value labels on top."""
    y = results.copy()
    y["Year"] = y["Ship Date"].dt.year.astype(str)
    qty = y.groupby("Year", sort=True)["QTY"].sum().reset_index()

    bars = alt.Chart(qty).mark_bar().encode(
        x=alt.X("Year:N", title="Year", sort=None),
        y=alt.Y("QTY:Q", title="QTY", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
        tooltip=["Year:N", alt.Tooltip("QTY:Q", format=",")],
    )
    text = bars.mark_text(dy=-10, fontSize=12).encode(
        text=alt.Text("QTY:Q", format=","),
    )
    return bars + text


def chart_qty_by_month(results: pd.DataFrame) -> alt.LayerChart:
    """QTY sum by month — bar chart with value labels on top."""
    m = results.copy()
    m["Month"] = m["Ship Date"].dt.to_period("M").astype(str)
    qty = m.groupby("Month", sort=True)["QTY"].sum().reset_index()

    bars = alt.Chart(qty).mark_bar().encode(
        x=alt.X("Month:N", title="Month", sort=None),
        y=alt.Y("QTY:Q", title="QTY", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
        tooltip=["Month:N", alt.Tooltip("QTY:Q", format=",")],
    )
    text = bars.mark_text(dy=-10, fontSize=11).encode(
        text=alt.Text("QTY:Q", format=","),
    )
    return bars + text


def chart_gp_pct_trend(results: pd.DataFrame) -> alt.Chart:
    """GP% monthly weighted-average trend — line chart."""
    m = results.copy()
    m["Month"] = m["Ship Date"].dt.to_period("M").astype(str)
    m["_gp_val"] = (m["UP"] - m["TP(USD)"]) * m["QTY"]
    m["_up_val"] = m["UP"] * m["QTY"]

    agg = m.groupby("Month", sort=True)[["_gp_val", "_up_val"]].sum().reset_index()
    agg["GP%"] = agg.apply(
        lambda r: r["_gp_val"] / r["_up_val"] * 100 if r["_up_val"] != 0 else 0,
        axis=1,
    )

    return (
        alt.Chart(agg)
        .mark_line(point=alt.OverlayMarkDef(size=30), color=NEGATIVE)
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("GP%:Q", title="GP%",
                     axis=alt.Axis(format=".1f", labelExpr="format(datum.value, '.1f')")),
            tooltip=["Month:N", alt.Tooltip("GP%:Q", format=".1f")],
        )
    )


def _cat_color_scale() -> alt.Scale:
    """Altair color scale for consistent category colors."""
    return alt.Scale(domain=list(CATEGORY_COLORS.keys()), range=list(CATEGORY_COLORS.values()))


# ── Dashboard charts ──────────────────────────────────────────────
def chart_revenue_trend(monthly_df, multi_year=False):
    """Monthly revenue line chart. Multi-year: color by year."""
    if multi_year:
        return (
            alt.Chart(monthly_df)
            .mark_line(point=alt.OverlayMarkDef(size=30))
            .encode(
                x=alt.X("MonthNum:O", title="Month",
                         axis=alt.Axis(labelExpr="datum.value")),
                y=alt.Y("Revenue:Q", title="Revenue", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
                color=alt.Color("Year:N"),
                tooltip=[
                    "Year:N", "MonthNum:O",
                    alt.Tooltip("Revenue:Q", format=",.0f"),
                ],
            )
        )
    return (
        alt.Chart(monthly_df)
        .mark_line(point=alt.OverlayMarkDef(size=30))
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Revenue:Q", title="Revenue", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
            tooltip=["Month:N", alt.Tooltip("Revenue:Q", format=",.0f")],
        )
    )


def chart_gp_dual_axis(monthly_df):
    """GP bar + GP% line dual-axis chart."""
    base = alt.Chart(monthly_df).encode(
        x=alt.X("Month:N", title="Month", sort=None),
    )
    bars = base.mark_bar(opacity=0.75, color=PRIMARY).encode(
        y=alt.Y("GP:Q", title="GP", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
        tooltip=["Month:N", alt.Tooltip("GP:Q", format=",.0f")],
    )
    line = base.mark_line(
        color=NEGATIVE,
        point=alt.OverlayMarkDef(size=30, color=NEGATIVE),
    ).encode(
        y=alt.Y("GP%:Q", title="GP%",
                 axis=alt.Axis(format=".1f", labelExpr="format(datum.value, '.1f')")),
        tooltip=["Month:N", alt.Tooltip("GP%:Q", format=".1f")],
    )
    return alt.layer(bars, line).resolve_scale(y="independent")


def chart_category_donut(cat_df):
    """Category revenue share donut chart."""
    return (
        alt.Chart(cat_df)
        .mark_arc(innerRadius=60)
        .encode(
            theta=alt.Theta("Revenue:Q"),
            color=alt.Color("Category:N", scale=_cat_color_scale(),
                           legend=alt.Legend(title="Category")),
            tooltip=[
                "Category:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
                alt.Tooltip("Pct:Q", format=".1f", title="Share %"),
            ],
        )
    )


def chart_category_stacked(monthly_cat_df):
    """Stacked bar: monthly revenue by category."""
    return (
        alt.Chart(monthly_cat_df)
        .mark_bar()
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Revenue:Q", title="Revenue", stack="zero", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
            color=alt.Color("Category:N", scale=_cat_color_scale()),
            tooltip=[
                "Month:N", "Category:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
            ],
        )
    )


def chart_customer_qty_by_cat(monthly_qty_cat_df: pd.DataFrame) -> alt.Chart:
    """Grouped bar chart: monthly QTY by Category for drill-down."""
    return (
        alt.Chart(monthly_qty_cat_df)
        .mark_bar()
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("QTY:Q", title="QTY", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
            color=alt.Color("Category:N", scale=_cat_color_scale(),
                           legend=alt.Legend(title="Category")),
            xOffset=alt.XOffset("Category:N"),
            tooltip=[
                "Month:N", "Category:N",
                alt.Tooltip("QTY:Q", format=",")
            ],
        )
    )


def chart_ai_sw_revenue_trend(monthly_cat_df: pd.DataFrame) -> alt.Chart:
    """AI_SW monthly revenue line chart."""
    ai_sw = monthly_cat_df[monthly_cat_df["Category"] == "AI_SW"].copy()
    if ai_sw.empty:
        return alt.Chart(pd.DataFrame({"Month": [], "Revenue": []})).mark_text(
            text="No AI_SW data", fontSize=14
        ).encode()

    return (
        alt.Chart(ai_sw)
        .mark_line(
            point=alt.OverlayMarkDef(size=30),
            color=CATEGORY_COLORS.get("AI_SW"),
        )
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Revenue:Q", title="Revenue", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
            tooltip=[
                "Month:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
            ],
        )
    )


def chart_top_customers_bar(top_df):
    """Horizontal bar chart for top N customers."""
    data = top_df.reset_index().copy()
    bars = alt.Chart(data).mark_bar(color=PRIMARY).encode(
        y=alt.Y("Customer Name:N", sort="-x", title=None),
        x=alt.X("Revenue:Q", title="Revenue", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
        tooltip=[
            "Customer Name:N",
            alt.Tooltip("Revenue:Q", format=",.0f"),
            alt.Tooltip("GP%:Q", format=".1f"),
        ],
    )
    text = bars.mark_text(align="left", dx=3, fontSize=11).encode(
        text=alt.Text("Revenue:Q", format=",.0f"),
    )
    return bars + text


def chart_customer_monthly(detail_monthly_df):
    """Single customer monthly revenue trend."""
    return (
        alt.Chart(detail_monthly_df)
        .mark_line(point=alt.OverlayMarkDef(size=30), color=PRIMARY)
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Revenue:Q", title="Revenue", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
            tooltip=[
                "Month:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
                alt.Tooltip("GP%:Q", format=".1f"),
            ],
        )
    )


def chart_customer_cat_donut(cat_df):
    """Single customer category breakdown donut."""
    return (
        alt.Chart(cat_df)
        .mark_arc(innerRadius=50)
        .encode(
            theta=alt.Theta("Revenue:Q"),
            color=alt.Color("Category:N", scale=_cat_color_scale()),
            tooltip=[
                "Category:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
                alt.Tooltip("Pct:Q", format=".1f", title="Share %"),
            ],
        )
    )


# ── Blended Actual + Forecast + Budget charts ────────────────
def _source_scales():
    """(color scale, dash scale) for the Actual/Forecast/Budget Source field."""
    domain = list(SOURCE_COLORS.keys())
    return (
        alt.Scale(domain=domain, range=[SOURCE_COLORS[k]["color"] for k in domain]),
        alt.Scale(domain=domain, range=[SOURCE_COLORS[k]["dash"] for k in domain]),
    )


def _actual_forecast_boundary_rule(df: pd.DataFrame):
    """Vertical rule at the start of the first Forecast month, or None."""
    actual_max = (
        df[df["Source"] == "Actual"]["MonthIndex"].max()
        if "Actual" in df["Source"].values else None
    )
    if actual_max is None or int(actual_max) >= 12:
        return None
    boundary_period = _MONTHS_ORDER[int(actual_max)]
    return (
        alt.Chart({"values": [{"Period": boundary_period}]})
        .mark_rule(color=MUTED, strokeDash=[4, 2], opacity=0.6, size=1)
        .encode(x=alt.X("Period:N", sort=_MONTHS_ORDER))
    )


def chart_revenue_trend_blended(blended_monthly_df: pd.DataFrame) -> alt.LayerChart:
    """Monthly revenue line: Actual = solid primary, Forecast = dashed accent, Budget = dashed muted.

    Input: output of fcst_loader.agg_blended_monthly() or concat with agg_budget_monthly().
    Columns required: Period, MonthIndex, Source, Revenue.
    """
    df = blended_monthly_df.copy()
    color_scale, dash_scale = _source_scales()
    line = (
        alt.Chart(df)
        .mark_line(point=alt.OverlayMarkDef(size=30))
        .encode(
            x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
            y=alt.Y("Revenue:Q", title="Revenue", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
            color=alt.Color("Source:N", scale=color_scale, legend=alt.Legend(title="")),
            strokeDash=alt.StrokeDash("Source:N", scale=dash_scale, legend=None),
            tooltip=[
                "Period:N", "Source:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
            ],
        )
    )
    rule = _actual_forecast_boundary_rule(df)
    return alt.layer(line, rule) if rule is not None else alt.layer(line)


def chart_qty_trend_blended(blended_monthly_df: pd.DataFrame) -> alt.LayerChart:
    """Monthly QTY line chart: Actual = solid primary, Forecast = dashed accent. Budget excluded.

    Input: output of fcst_loader.agg_blended_monthly().
    Columns required: Period, MonthIndex, Source, QTY.
    """
    df = blended_monthly_df[blended_monthly_df["Source"] != "Budget"].copy()
    color_scale, dash_scale = _source_scales()
    line = (
        alt.Chart(df)
        .mark_line(point=alt.OverlayMarkDef(size=30))
        .encode(
            x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
            y=alt.Y("QTY:Q", title="QTY", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
            color=alt.Color("Source:N", scale=color_scale, legend=alt.Legend(title="")),
            strokeDash=alt.StrokeDash("Source:N", scale=dash_scale, legend=None),
            tooltip=[
                "Period:N", "Source:N",
                alt.Tooltip("QTY:Q", format=","),
            ],
        )
    )
    rule = _actual_forecast_boundary_rule(df)
    return alt.layer(line, rule) if rule is not None else alt.layer(line)


def chart_gp_trend_blended(blended_monthly_df: pd.DataFrame) -> alt.LayerChart:
    """GP bar + GP% line dual-axis chart with Actual/Forecast/Budget color coding.

    Budget bars are rendered as a separate low-opacity layer (reference only) so
    they overlap rather than stack with Actual/Forecast bars.
    GP% line uses only Actual/Forecast rows to avoid jumps from Budget GP%.

    Input: output of fcst_loader.agg_blended_monthly() or concat with agg_budget_monthly().
    Columns required: Period, MonthIndex, Source, GP, GP%.
    """
    color_scale, _ = _source_scales()
    df = blended_monthly_df.copy()
    df_main = df[df["Source"] != "Budget"]
    df_budget = df[df["Source"] == "Budget"]

    bars_main = (
        alt.Chart(df_main)
        .mark_bar(opacity=0.75)
        .encode(
            x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
            y=alt.Y("GP:Q", title="GP", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
            color=alt.Color("Source:N", scale=color_scale, legend=alt.Legend(title="")),
            tooltip=[
                "Period:N", "Source:N",
                alt.Tooltip("GP:Q", format=",.0f"),
            ],
        )
    )
    bars_budget = (
        alt.Chart(df_budget)
        .mark_bar(opacity=0.3)
        .encode(
            x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
            y=alt.Y("GP:Q", title="GP", axis=alt.Axis(labelExpr=_NUM_LABEL_EXPR)),
            color=alt.Color("Source:N", scale=color_scale, legend=alt.Legend(title="")),
            tooltip=[
                "Period:N", "Source:N",
                alt.Tooltip("GP:Q", format=",.0f"),
            ],
        )
    )
    line = (
        alt.Chart(df_main)
        .mark_line(
            color=NEGATIVE,
            point=alt.OverlayMarkDef(size=30, color=NEGATIVE),
        )
        .encode(
            x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
            y=alt.Y("GP%:Q", title="GP%",
                     axis=alt.Axis(format=".1f", labelExpr="format(datum.value, '.1f')")),
            tooltip=["Period:N", alt.Tooltip("GP%:Q", format=".1f")],
        )
    )
    return alt.layer(bars_main, bars_budget, line).resolve_scale(y="independent")
