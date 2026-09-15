"""charts.py — Altair chart builders for Shipping Record Search & Dashboard tabs.

All colors come from theme.py (CATEGORY_COLORS / SOURCE_COLORS / get_tokens) —
no hex literals here. Call ``apply_altair_theme(mode)`` once per script run
(after resolving the theme setting) before rendering any chart.
"""

import altair as alt
import pandas as pd

from theme import CATEGORY_COLORS, SOURCE_COLORS, get_tokens

_MONTHS_ORDER = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]

_ALTAIR_THEME_NAME = "sales_report_tool"

# Vega expression: abbreviate numeric axis labels (1.2M / 340K); pass
# through non-numeric (nominal/ordinal) labels unchanged.
_NUM_LABEL_EXPR = (
    "isNumber(datum.value) ? "
    "(abs(datum.value) >= 1e9 ? format(datum.value / 1e9, '.2f') + 'B' : "
    "abs(datum.value) >= 1e6 ? format(datum.value / 1e6, '.1f') + 'M' : "
    "abs(datum.value) >= 1e3 ? format(datum.value / 1e3, '.0f') + 'K' : "
    "format(datum.value, ',')) : datum.value"
)


def apply_altair_theme(mode: str) -> None:
    """Register + enable the shared Altair theme for *mode* ('light'/'dark').

    Re-registering under the same theme name is idempotent, so calling this
    once near the top of every Streamlit script run (with the currently
    resolved mode) keeps every chart in sync with the Settings theme.
    """
    t = get_tokens(mode)

    def _theme_config():
        return alt.theme.ThemeConfig({
            "config": {
                "background": "transparent",
                "font": "-apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
                "title": {"color": t["text"], "fontSize": 14, "fontWeight": 600},
                "text": {"color": t["text"]},
                "view": {"stroke": None},
                "point": {"size": 30},
                "axis": {
                    "domain": False,
                    "tickColor": t["border"],
                    "labelColor": t["text"],
                    "labelFontSize": 11,
                    "titleColor": t["muted"],
                    "titleFontSize": 11,
                    "titleFontWeight": 600,
                    "labelExpr": _NUM_LABEL_EXPR,
                    "gridColor": t["border"],
                },
                "axisX": {"grid": False},
                "axisY": {"grid": True},
                "legend": {
                    "orient": "top",
                    "title": None,
                    "labelColor": t["text"],
                    "labelFontSize": 11,
                },
            }
        })

    alt.theme.register(_ALTAIR_THEME_NAME, enable=True)(_theme_config)


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
            y=alt.Y("Price:Q", title="Price"),
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
        y=alt.Y("QTY:Q", title="QTY"),
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
        y=alt.Y("QTY:Q", title="QTY"),
        tooltip=["Month:N", alt.Tooltip("QTY:Q", format=",")],
    )
    text = bars.mark_text(dy=-10, fontSize=11).encode(
        text=alt.Text("QTY:Q", format=","),
    )
    return bars + text


def chart_gp_pct_trend(results: pd.DataFrame, mode: str = "light") -> alt.Chart:
    """GP% monthly weighted-average trend — line chart."""
    t = get_tokens(mode)
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
        .mark_line(point=alt.OverlayMarkDef(size=30), color=t["negative"])
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("GP%:Q", title="GP%",
                     axis=alt.Axis(format=".1f", labelExpr="format(datum.value, '.1f')")),
            tooltip=["Month:N", alt.Tooltip("GP%:Q", format=".1f")],
        )
    )


def _cat_color_scale(mode: str) -> alt.Scale:
    """Altair color scale for consistent category colors."""
    cats = CATEGORY_COLORS(mode)
    return alt.Scale(domain=list(cats.keys()), range=list(cats.values()))


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
                y=alt.Y("Revenue:Q", title="Revenue"),
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
            y=alt.Y("Revenue:Q", title="Revenue"),
            tooltip=["Month:N", alt.Tooltip("Revenue:Q", format=",.0f")],
        )
    )


def chart_gp_dual_axis(monthly_df, mode: str = "light"):
    """GP bar + GP% line dual-axis chart."""
    t = get_tokens(mode)
    base = alt.Chart(monthly_df).encode(
        x=alt.X("Month:N", title="Month", sort=None),
    )
    bars = base.mark_bar(opacity=0.75, color=t["primary"]).encode(
        y=alt.Y("GP:Q", title="GP"),
        tooltip=["Month:N", alt.Tooltip("GP:Q", format=",.0f")],
    )
    line = base.mark_line(
        color=t["negative"],
        point=alt.OverlayMarkDef(size=30, color=t["negative"]),
    ).encode(
        y=alt.Y("GP%:Q", title="GP%",
                 axis=alt.Axis(format=".1f", labelExpr="format(datum.value, '.1f')")),
        tooltip=["Month:N", alt.Tooltip("GP%:Q", format=".1f")],
    )
    return alt.layer(bars, line).resolve_scale(y="independent")


def chart_category_donut(cat_df, mode: str = "light"):
    """Category revenue share donut chart."""
    return (
        alt.Chart(cat_df)
        .mark_arc(innerRadius=60)
        .encode(
            theta=alt.Theta("Revenue:Q"),
            color=alt.Color("Category:N", scale=_cat_color_scale(mode),
                           legend=alt.Legend(title="Category")),
            tooltip=[
                "Category:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
                alt.Tooltip("Pct:Q", format=".1f", title="Share %"),
            ],
        )
    )


def chart_category_stacked(monthly_cat_df, mode: str = "light"):
    """Stacked bar: monthly revenue by category."""
    return (
        alt.Chart(monthly_cat_df)
        .mark_bar()
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Revenue:Q", title="Revenue", stack="zero"),
            color=alt.Color("Category:N", scale=_cat_color_scale(mode)),
            tooltip=[
                "Month:N", "Category:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
            ],
        )
    )


def chart_customer_qty_by_cat(monthly_qty_cat_df: pd.DataFrame, mode: str = "light") -> alt.Chart:
    """Grouped bar chart: monthly QTY by Category for drill-down."""
    return (
        alt.Chart(monthly_qty_cat_df)
        .mark_bar()
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("QTY:Q", title="QTY"),
            color=alt.Color("Category:N", scale=_cat_color_scale(mode),
                           legend=alt.Legend(title="Category")),
            xOffset=alt.XOffset("Category:N"),
            tooltip=[
                "Month:N", "Category:N",
                alt.Tooltip("QTY:Q", format=",")
            ],
        )
    )


def chart_ai_sw_revenue_trend(monthly_cat_df: pd.DataFrame, mode: str = "light") -> alt.Chart:
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
            color=CATEGORY_COLORS(mode).get("AI_SW"),
        )
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Revenue:Q", title="Revenue"),
            tooltip=[
                "Month:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
            ],
        )
    )


def chart_top_customers_bar(top_df, mode: str = "light"):
    """Horizontal bar chart for top N customers."""
    t = get_tokens(mode)
    data = top_df.reset_index().copy()
    bars = alt.Chart(data).mark_bar(color=t["primary"]).encode(
        y=alt.Y("Customer Name:N", sort="-x", title=None),
        x=alt.X("Revenue:Q", title="Revenue"),
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


def chart_customer_monthly(detail_monthly_df, mode: str = "light"):
    """Single customer monthly revenue trend."""
    t = get_tokens(mode)
    return (
        alt.Chart(detail_monthly_df)
        .mark_line(point=alt.OverlayMarkDef(size=30), color=t["primary"])
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Revenue:Q", title="Revenue"),
            tooltip=[
                "Month:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
                alt.Tooltip("GP%:Q", format=".1f"),
            ],
        )
    )


def chart_customer_cat_donut(cat_df, mode: str = "light"):
    """Single customer category breakdown donut."""
    return (
        alt.Chart(cat_df)
        .mark_arc(innerRadius=50)
        .encode(
            theta=alt.Theta("Revenue:Q"),
            color=alt.Color("Category:N", scale=_cat_color_scale(mode)),
            tooltip=[
                "Category:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
                alt.Tooltip("Pct:Q", format=".1f", title="Share %"),
            ],
        )
    )


# ── Blended Actual + Forecast + Budget charts ────────────────
def _source_scales(mode: str):
    """(color scale, dash scale) for the Actual/Forecast/Budget Source field."""
    sc = SOURCE_COLORS(mode)
    domain = list(sc.keys())
    return (
        alt.Scale(domain=domain, range=[sc[k]["color"] for k in domain]),
        alt.Scale(domain=domain, range=[sc[k]["dash"] for k in domain]),
    )


def _actual_forecast_boundary_rule(df: pd.DataFrame, mode: str):
    """Vertical rule at the start of the first Forecast month, or None."""
    t = get_tokens(mode)
    actual_max = (
        df[df["Source"] == "Actual"]["MonthIndex"].max()
        if "Actual" in df["Source"].values else None
    )
    if actual_max is None or int(actual_max) >= 12:
        return None
    boundary_period = _MONTHS_ORDER[int(actual_max)]
    return (
        alt.Chart({"values": [{"Period": boundary_period}]})
        .mark_rule(color=t["muted"], strokeDash=[4, 2], opacity=0.6, size=1)
        .encode(x=alt.X("Period:N", sort=_MONTHS_ORDER))
    )


def chart_revenue_trend_blended(blended_monthly_df: pd.DataFrame, mode: str = "light") -> alt.LayerChart:
    """Monthly revenue line: Actual = solid primary, Forecast = dashed accent, Budget = dashed muted.

    Input: output of fcst_loader.agg_blended_monthly() or concat with agg_budget_monthly().
    Columns required: Period, MonthIndex, Source, Revenue.
    """
    df = blended_monthly_df.copy()
    color_scale, dash_scale = _source_scales(mode)
    line = (
        alt.Chart(df)
        .mark_line(point=alt.OverlayMarkDef(size=30))
        .encode(
            x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
            y=alt.Y("Revenue:Q", title="Revenue"),
            color=alt.Color("Source:N", scale=color_scale, legend=alt.Legend(title="")),
            strokeDash=alt.StrokeDash("Source:N", scale=dash_scale, legend=None),
            tooltip=[
                "Period:N", "Source:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
            ],
        )
    )
    rule = _actual_forecast_boundary_rule(df, mode)
    return alt.layer(line, rule) if rule is not None else alt.layer(line)


def chart_qty_trend_blended(blended_monthly_df: pd.DataFrame, mode: str = "light") -> alt.LayerChart:
    """Monthly QTY line chart: Actual = solid primary, Forecast = dashed accent. Budget excluded.

    Input: output of fcst_loader.agg_blended_monthly().
    Columns required: Period, MonthIndex, Source, QTY.
    """
    df = blended_monthly_df[blended_monthly_df["Source"] != "Budget"].copy()
    color_scale, dash_scale = _source_scales(mode)
    line = (
        alt.Chart(df)
        .mark_line(point=alt.OverlayMarkDef(size=30))
        .encode(
            x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
            y=alt.Y("QTY:Q", title="QTY"),
            color=alt.Color("Source:N", scale=color_scale, legend=alt.Legend(title="")),
            strokeDash=alt.StrokeDash("Source:N", scale=dash_scale, legend=None),
            tooltip=[
                "Period:N", "Source:N",
                alt.Tooltip("QTY:Q", format=","),
            ],
        )
    )
    rule = _actual_forecast_boundary_rule(df, mode)
    return alt.layer(line, rule) if rule is not None else alt.layer(line)


def chart_gp_trend_blended(blended_monthly_df: pd.DataFrame, mode: str = "light") -> alt.LayerChart:
    """GP bar + GP% line dual-axis chart with Actual/Forecast/Budget color coding.

    Budget bars are rendered as a separate low-opacity layer (reference only) so
    they overlap rather than stack with Actual/Forecast bars.
    GP% line uses only Actual/Forecast rows to avoid jumps from Budget GP%.

    Input: output of fcst_loader.agg_blended_monthly() or concat with agg_budget_monthly().
    Columns required: Period, MonthIndex, Source, GP, GP%.
    """
    t = get_tokens(mode)
    color_scale, _ = _source_scales(mode)
    df = blended_monthly_df.copy()
    df_main = df[df["Source"] != "Budget"]
    df_budget = df[df["Source"] == "Budget"]

    bars_main = (
        alt.Chart(df_main)
        .mark_bar(opacity=0.75)
        .encode(
            x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
            y=alt.Y("GP:Q", title="GP"),
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
            y=alt.Y("GP:Q", title="GP"),
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
            color=t["negative"],
            point=alt.OverlayMarkDef(size=30, color=t["negative"]),
        )
        .encode(
            x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
            y=alt.Y("GP%:Q", title="GP%",
                     axis=alt.Axis(format=".1f", labelExpr="format(datum.value, '.1f')")),
            tooltip=["Period:N", alt.Tooltip("GP%:Q", format=".1f")],
        )
    )
    return alt.layer(bars_main, bars_budget, line).resolve_scale(y="independent")
