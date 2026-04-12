"""charts.py — Altair chart builders for Shipping Record Search tab."""

import altair as alt
import pandas as pd

_MONTHS_ORDER = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
]


def chart_up_tp_trend(results: pd.DataFrame) -> alt.LayerChart:
    """UP & TP(USD) monthly average trend — line chart with tooltips & point markers."""
    m = results.copy()
    m["Month"] = m["Ship Date"].dt.to_period("M").astype(str)
    avg = m.groupby("Month", sort=True)[["UP", "TP(USD)"]].mean().reset_index()
    long = avg.melt("Month", var_name="Metric", value_name="Price")

    return (
        alt.Chart(long)
        .mark_line(point=alt.OverlayMarkDef(size=40))
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
        .mark_line(point=alt.OverlayMarkDef(size=40), color="#ff7f0e")
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("GP%:Q", title="GP%"),
            tooltip=["Month:N", alt.Tooltip("GP%:Q", format=".1f")],
        )
    )


# ── Dashboard color scheme ────────────────────────────────────────
CAT_COLORS = {
    "CDR":        "#1f77b4",
    "CDR ACC":    "#aec7e8",
    "Tablet":     "#ff7f0e",
    "Tablet ACC": "#ffbb78",
    "AI_SW":      "#2ca02c",
    "Others":     "#d62728",
}


def _cat_color_scale():
    """Altair color scale for consistent category colors."""
    return alt.Scale(
        domain=list(CAT_COLORS.keys()),
        range=list(CAT_COLORS.values()),
    )


# ── Dashboard charts ──────────────────────────────────────────────
def chart_revenue_trend(monthly_df, multi_year=False):
    """Monthly revenue line chart. Multi-year: color by year."""
    if multi_year:
        return (
            alt.Chart(monthly_df)
            .mark_line(point=alt.OverlayMarkDef(size=40))
            .encode(
                x=alt.X("MonthNum:O", title="Month",
                         axis=alt.Axis(labelExpr="datum.value")),
                y=alt.Y("Revenue:Q", title="Revenue",
                         axis=alt.Axis(format=",.0f")),
                color=alt.Color("Year:N"),
                tooltip=[
                    "Year:N", "MonthNum:O",
                    alt.Tooltip("Revenue:Q", format=",.0f"),
                ],
            )
        )
    return (
        alt.Chart(monthly_df)
        .mark_line(point=alt.OverlayMarkDef(size=40))
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Revenue:Q", title="Revenue",
                     axis=alt.Axis(format=",.0f")),
            tooltip=["Month:N", alt.Tooltip("Revenue:Q", format=",.0f")],
        )
    )


def chart_gp_dual_axis(monthly_df):
    """GP bar + GP% line dual-axis chart."""
    base = alt.Chart(monthly_df).encode(
        x=alt.X("Month:N", title="Month", sort=None),
    )
    bars = base.mark_bar(opacity=0.6, color="#5470c6").encode(
        y=alt.Y("GP:Q", title="GP", axis=alt.Axis(format=",.0f")),
        tooltip=["Month:N", alt.Tooltip("GP:Q", format=",.0f")],
    )
    line = base.mark_line(
        color="#ee6666",
        point=alt.OverlayMarkDef(size=40, color="#ee6666"),
    ).encode(
        y=alt.Y("GP%:Q", title="GP%", axis=alt.Axis(format=".1f")),
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
            y=alt.Y("Revenue:Q", title="Revenue", stack="zero",
                     axis=alt.Axis(format=",.0f")),
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
            y=alt.Y("QTY:Q", title="QTY",
                     axis=alt.Axis(format=",.0f")),
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
            point=alt.OverlayMarkDef(size=50),
            color=CAT_COLORS.get("AI_SW", "#2ca02c"),
        )
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Revenue:Q", title="Revenue", axis=alt.Axis(format=",.0f")),
            tooltip=[
                "Month:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
            ],
        )
    )


def chart_top_customers_bar(top_df):
    """Horizontal bar chart for top N customers."""
    data = top_df.reset_index().copy()
    bars = alt.Chart(data).mark_bar().encode(
        y=alt.Y("Customer Name:N", sort="-x", title=None),
        x=alt.X("Revenue:Q", title="Revenue",
                 axis=alt.Axis(format=",.0f")),
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
        .mark_line(point=alt.OverlayMarkDef(size=40))
        .encode(
            x=alt.X("Month:N", title="Month", sort=None),
            y=alt.Y("Revenue:Q", title="Revenue",
                     axis=alt.Axis(format=",.0f")),
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


# ── Blended Actual + Forecast charts ─────────────────────────────

_SOURCE_COLOR = alt.Scale(
    domain=["Actual", "Forecast"],
    range=["#5470c6", "#91cc75"],
)
_SOURCE_DASH = alt.Scale(
    domain=["Actual", "Forecast"],
    range=[[1, 0], [6, 3]],
)


def chart_revenue_trend_blended(blended_monthly_df: pd.DataFrame) -> alt.LayerChart:
    """Monthly revenue line: Actual = solid blue, Forecast = dashed green.

    Input: output of fcst_loader.agg_blended_monthly().
    Columns required: Period, MonthIndex, Source, Revenue.
    """
    df = blended_monthly_df.copy()
    line = (
        alt.Chart(df)
        .mark_line(point=alt.OverlayMarkDef(size=40))
        .encode(
            x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
            y=alt.Y("Revenue:Q", title="Revenue", axis=alt.Axis(format=",.0f")),
            color=alt.Color("Source:N", scale=_SOURCE_COLOR,
                            legend=alt.Legend(title="")),
            strokeDash=alt.StrokeDash("Source:N", scale=_SOURCE_DASH,
                                      legend=None),
            tooltip=[
                "Period:N", "Source:N",
                alt.Tooltip("Revenue:Q", format=",.0f"),
            ],
        )
    )
    # Vertical rule at the start of the first Forecast month
    actual_max = (
        df[df["Source"] == "Actual"]["MonthIndex"].max()
        if "Actual" in df["Source"].values else None
    )
    if actual_max is not None and int(actual_max) < 12:
        boundary_period = _MONTHS_ORDER[int(actual_max)]  # first Forecast month
        rule = (
            alt.Chart({"values": [{"Period": boundary_period}]})
            .mark_rule(color="#aaaaaa", strokeDash=[4, 2], opacity=0.6, size=1)
            .encode(x=alt.X("Period:N", sort=_MONTHS_ORDER))
        )
        return alt.layer(line, rule)
    return alt.layer(line)


def chart_gp_trend_blended(blended_monthly_df: pd.DataFrame) -> alt.LayerChart:
    """GP bar + GP% line dual-axis chart with Actual/Forecast color coding.

    Input: output of fcst_loader.agg_blended_monthly().
    Columns required: Period, MonthIndex, Source, GP, GP%.
    """
    df = blended_monthly_df.copy()
    base = alt.Chart(df).encode(
        x=alt.X("Period:N", title="Month", sort=_MONTHS_ORDER),
    )
    bars = base.mark_bar(opacity=0.65).encode(
        y=alt.Y("GP:Q", title="GP", axis=alt.Axis(format=",.0f")),
        color=alt.Color("Source:N", scale=_SOURCE_COLOR,
                        legend=alt.Legend(title="")),
        tooltip=[
            "Period:N", "Source:N",
            alt.Tooltip("GP:Q", format=",.0f"),
        ],
    )
    line = base.mark_line(
        color="#ee6666",
        point=alt.OverlayMarkDef(size=40, color="#ee6666"),
    ).encode(
        y=alt.Y("GP%:Q", title="GP%", axis=alt.Axis(format=".1f")),
        tooltip=["Period:N", alt.Tooltip("GP%:Q", format=".1f")],
    )
    return alt.layer(bars, line).resolve_scale(y="independent")
