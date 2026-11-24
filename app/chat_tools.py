"""chat_tools.py - read-only data tools for the in-app chatbot.

Pure pandas (no Streamlit state). Every function takes the already-filtered
shipping DataFrame and returns a small JSON-serialisable dict, so the LLM only
ever sees aggregates - never raw rows - and does no arithmetic of its own.
"""
from __future__ import annotations

import math
from typing import Any, Callable, Optional

import pandas as pd

from utils import GP_COL, QTY_CATEGORIES, calc_dashboard_kpis, _normalize_name

MAX_ROWS = 50
DEFAULT_ROWS = 10

_GROUP_COLS = {
    "customer": "Customer Name",
    "category": "Category",
    "month": "Month",
    "year": "_Year",
    "part_number": "Part Number",
    "none": None,
}
_METRICS = ("revenue", "gp", "gp_pct", "qty")


def _clean(v: Any) -> Any:
    """Make numpy / NaN values JSON friendly."""
    if v is None:
        return None
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    if hasattr(v, "item"):
        v = v.item()
    return round(v, 2) if isinstance(v, float) else v


def _agg(df: pd.DataFrame) -> dict:
    revenue = float(df["SALES Total AMT"].sum())
    gp = float(df[GP_COL].sum())
    qty = float(df[df["Category"].isin(QTY_CATEGORIES)]["QTY"].sum())
    return {
        "revenue": revenue,
        "gp": gp,
        "gp_pct": gp / revenue * 100 if revenue else 0.0,
        "qty": qty,
    }


def match_customers(df: pd.DataFrame, text: str) -> list[str]:
    """Customer Names whose normalized form contains the normalized *text*."""
    needle = _normalize_name(text)
    if not needle:
        return []
    names = [n for n in df["Customer Name"].dropna().unique() if isinstance(n, str)]
    exact = [n for n in names if _normalize_name(n) == needle]
    if exact:
        return sorted(exact)
    return sorted(n for n in names if needle in _normalize_name(n))


def query_sales(
    df: pd.DataFrame,
    group_by: str = "none",
    years: Optional[list] = None,
    category: Optional[str] = None,
    customer: Optional[str] = None,
    months: Optional[list] = None,
    sort_by: str = "revenue",
    top_n: int = DEFAULT_ROWS,
) -> dict:
    """Aggregate actual shipments. Amounts are TWD."""
    group_by = (group_by or "none").lower()
    if group_by not in _GROUP_COLS:
        return {"error": f"group_by must be one of {sorted(_GROUP_COLS)}"}
    sort_by = (sort_by or "revenue").lower()
    if sort_by not in _METRICS:
        return {"error": f"sort_by must be one of {list(_METRICS)}"}

    work = df.copy()
    work["_Year"] = work["Ship Date"].dt.year.astype(int)
    filters: dict[str, Any] = {}

    if years:
        years = [int(y) for y in years]
        work = work[work["_Year"].isin(years)]
        filters["years"] = years
    if months:
        months = [int(m) for m in months]
        work = work[work["Ship Date"].dt.month.isin(months)]
        filters["months"] = months
    if category:
        cat = category.strip().lower()
        work = work[work["Category"].str.lower() == cat]
        filters["category"] = category
    matched: list[str] = []
    if customer:
        matched = match_customers(df, customer)
        if not matched:
            return {"error": f"No customer matches '{customer}'.", "filters": filters}
        work = work[work["Customer Name"].isin(matched)]
        filters["customers_matched"] = matched[:10]

    if work.empty:
        return {"filters": filters, "rows": [], "note": "No rows match these filters."}

    top_n = max(1, min(int(top_n or DEFAULT_ROWS), MAX_ROWS))
    gcol = _GROUP_COLS[group_by]
    if gcol is None:
        rows = [{k: _clean(v) for k, v in _agg(work).items()}]
        return {"filters": filters, "rows": rows, "currency": "TWD"}

    records = []
    for key, part in work.groupby(gcol, dropna=True):
        rec = {group_by: _clean(key)}
        rec.update({k: _clean(v) for k, v in _agg(part).items()})
        records.append(rec)
    if group_by in ("month", "year"):
        records.sort(key=lambda r: r[group_by])
    else:
        records.sort(key=lambda r: r[sort_by] or 0, reverse=True)
    total_groups = len(records)
    out = {
        "filters": filters,
        "currency": "TWD",
        "rows": records[:top_n],
        "total_groups": total_groups,
    }
    if total_groups > top_n:
        out["note"] = f"Showing {top_n} of {total_groups} groups."
    return out


CHART_TYPES = ("bar", "line", "pie")


def show_chart(
    df: pd.DataFrame,
    chart_type: str,
    group_by: str,
    metric: str = "revenue",
    title: Optional[str] = None,
    **query_args: Any,
) -> dict:
    """Chart request: runs query_sales and returns a render-ready spec.

    The UI draws the chart from ``result["chart"]``; the model only learns that
    the chart was shown, so it still never computes or sees raw rows.
    """
    chart_type = (chart_type or "bar").lower()
    if chart_type not in CHART_TYPES:
        return {"error": f"chart_type must be one of {list(CHART_TYPES)}"}
    metric = (metric or "revenue").lower()
    if metric not in _METRICS:
        return {"error": f"metric must be one of {list(_METRICS)}"}
    if (group_by or "none").lower() == "none":
        return {"error": "A chart needs group_by (customer, category, month, year or part_number)."}
    query_args.setdefault("sort_by", metric)
    res = query_sales(df, group_by=group_by, **query_args)
    if "error" in res or not res.get("rows"):
        return res
    group_by = group_by.lower()
    points = [{"label": str(r[group_by]), "value": r[metric]} for r in res["rows"]]
    return {
        "filters": res.get("filters", {}),
        "currency": "TWD",
        "rows": res["rows"],
        "chart": {
            "type": chart_type,
            "x": group_by,
            "metric": metric,
            "title": title or f"{metric} by {group_by}",
            "points": points,
            "chronological": group_by in ("month", "year"),
        },
        "note": "Chart displayed to the user. Do not repeat all the numbers; summarise the insight.",
    }


def get_kpis(df: pd.DataFrame, year: int) -> dict:
    """Headline KPIs for *year* with YoY change vs the previous year."""
    years = df["Ship Date"].dt.year
    cur = df[years == int(year)]
    if cur.empty:
        return {"error": f"No shipments found for {year}."}
    prev = df[years == int(year) - 1]
    k = calc_dashboard_kpis(cur, prev if not prev.empty else None)
    return {"year": int(year), "currency": "TWD", **{n: _clean(v) for n, v in k.items()}}


def get_forecast(
    df: pd.DataFrame,
    fcst_df: Optional[pd.DataFrame],
    current_year: int,
    current_month: int,
    customer: Optional[str] = None,
) -> dict:
    """FY forecast = Actual for months < current month + Forecast from the
    current month on (same rule as the dashboard), plus Budget."""
    if fcst_df is None or fcst_df.empty:
        return {"error": "No FCST file is loaded, so there is no forecast data."}
    import fcst_loader  # local import: keeps module import light for tests

    actual = df[df["Ship Date"].dt.year == current_year].copy()
    fc = fcst_df
    filters: dict[str, Any] = {"year": current_year}
    if customer:
        matched = match_customers(
            pd.DataFrame({"Customer Name": list(set(fc["Customer"]) | set(actual["Customer Name"]))}),
            customer,
        )
        if not matched:
            return {"error": f"No customer matches '{customer}' in FCST or actuals."}
        actual = actual[actual["Customer Name"].isin(matched)]
        fc = fc[fc["Customer"].isin(matched)]
        filters["customers_matched"] = matched[:10]

    actual = actual.rename(columns={"Customer Name": "Customer"})
    actual["Month"] = actual["Ship Date"].dt.month
    blended = fcst_loader.blend_actual_fcst(actual, fc, current_month)
    monthly = fcst_loader.agg_blended_monthly(blended)
    budget = fcst_loader.agg_budget_monthly(fc)
    if monthly.empty:
        return {"error": "Forecast data is empty for this selection.", "filters": filters}

    fy_rev = float(monthly["Revenue"].sum())
    fy_gp = float(monthly["GP"].sum())
    ytd_actual = float(monthly[monthly["Source"] == "Actual"]["Revenue"].sum())
    out = {
        "filters": filters,
        "currency": "TWD",
        "fy_forecast_revenue": _clean(fy_rev),
        "fy_forecast_gp": _clean(fy_gp),
        "fy_forecast_gp_pct": _clean(fy_gp / fy_rev * 100 if fy_rev else 0.0),
        "ytd_actual_revenue": _clean(ytd_actual),
        "basis": f"Actual for months before {current_month}, Forecast from month {current_month}.",
    }
    if not budget.empty:
        fy_budget = float(budget["Revenue"].sum())
        out["fy_budget_revenue"] = _clean(fy_budget)
        out["budget_achievement_pct_ytd"] = _clean(
            ytd_actual / fy_budget * 100 if fy_budget else None
        )
    return out


# -- Tool schemas (Anthropic tool-use format) ---------------------------------
TOOL_SCHEMAS = [
    {
        "name": "query_sales",
        "description": (
            "Aggregate ACTUAL shipment records (amounts in TWD). Returns revenue, gp, "
            "gp_pct and qty (qty counts only CDR and Tablet categories). Use for top-N "
            "customers, category breakdowns, monthly trends, part-number rankings, and "
            "totals for a customer/category/period."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "group_by": {"type": "string", "enum": sorted(_GROUP_COLS),
                             "description": "Dimension to group by; 'none' for one grand total."},
                "years": {"type": "array", "items": {"type": "integer"},
                          "description": "Calendar years to include; omit for all years."},
                "months": {"type": "array", "items": {"type": "integer"},
                           "description": "Month numbers 1-12 to include; omit for all."},
                "category": {"type": "string",
                             "description": "One of Tablet, CDR, Tablet ACC, CDR ACC, AI_SW, Signify, Others."},
                "customer": {"type": "string",
                             "description": "Customer name or part of it (fuzzy match)."},
                "sort_by": {"type": "string", "enum": list(_METRICS)},
                "top_n": {"type": "integer", "description": f"Max rows (default {DEFAULT_ROWS}, max {MAX_ROWS})."},
            },
            "required": ["group_by"],
        },
    },
    {
        "name": "show_chart",
        "description": (
            "Display a chart to the user (bar, line or pie) of one metric grouped by a "
            "dimension. Use when the user asks for a chart/graph/plot/visualisation or a "
            "trend. Takes the same filters as query_sales; line suits month/year trends, "
            "bar suits rankings, pie suits shares of a total."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "chart_type": {"type": "string", "enum": list(CHART_TYPES)},
                "group_by": {"type": "string", "enum": [g for g in sorted(_GROUP_COLS) if g != "none"]},
                "metric": {"type": "string", "enum": list(_METRICS),
                           "description": "Value to plot (default revenue)."},
                "title": {"type": "string"},
                "years": {"type": "array", "items": {"type": "integer"}},
                "months": {"type": "array", "items": {"type": "integer"}},
                "category": {"type": "string"},
                "customer": {"type": "string"},
                "top_n": {"type": "integer", "description": f"Max points (default {DEFAULT_ROWS}, max {MAX_ROWS})."},
            },
            "required": ["chart_type", "group_by"],
        },
    },
    {
        "name": "get_kpis",
        "description": "Headline KPIs (revenue, gp, gp_pct, qty) for one year with YoY change vs the prior year.",
        "input_schema": {
            "type": "object",
            "properties": {"year": {"type": "integer"}},
            "required": ["year"],
        },
    },
    {
        "name": "get_forecast",
        "description": (
            "Current-year full-year forecast (Actual for past months + FCST for the rest), "
            "plus budget and YTD budget achievement. Optional customer filter."
        ),
        "input_schema": {
            "type": "object",
            "properties": {"customer": {"type": "string"}},
        },
    },
]


def make_executor(
    df: pd.DataFrame,
    fcst_df: Any,
    current_year: int,
    current_month: int,
) -> Callable[[str, dict], dict]:
    """Bind the data to a (tool_name, args) -> result callable.

    *fcst_df* may be a DataFrame, None, or a zero-arg callable returning one
    (loaded lazily, only if the model actually asks for the forecast).
    """

    def run(name: str, args: dict) -> dict:
        args = args or {}
        try:
            if name == "query_sales":
                return query_sales(df, **args)
            if name == "show_chart":
                return show_chart(df, **args)
            if name == "get_kpis":
                return get_kpis(df, **args)
            if name == "get_forecast":
                fc = fcst_df() if callable(fcst_df) else fcst_df
                return get_forecast(df, fc, current_year, current_month, **args)
        except TypeError as exc:  # unexpected / missing argument from the model
            return {"error": f"Bad arguments for {name}: {exc}"}
        except Exception as exc:  # never let a tool crash the chat
            return {"error": f"{name} failed: {exc}"}
        return {"error": f"Unknown tool '{name}'."}

    return run
