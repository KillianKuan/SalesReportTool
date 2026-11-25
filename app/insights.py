"""insights.py - anomaly detection + AI executive summaries for the dashboard.

Two layers:
  1. detect_findings(): deterministic pandas rules (revenue drops, GP% erosion,
     category/customer declines, budget gap). Always available, no API needed.
  2. AI narration: ONE API call turns the findings + a small fact sheet into an
     executive headline, top-3 alert wording and a one/two-line caption per
     dashboard section. The model is told to reuse only the supplied numbers.

Results are cached in insights_cache.json (keyed by a data fingerprint) so the
API is only hit when the data changes or the refresh schedule (weekly/monthly)
says the stored summary is stale. No Streamlit imports: unit-testable.
"""
from __future__ import annotations

import hashlib
import json
import logging
import os
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

import pandas as pd

import chatbot

_logger = logging.getLogger(__name__)

REV = "SALES Total AMT"
GP = "final GP(NTD,data from Financial Report)"
CAT = "Category"
CUST = "Customer Name"

SECTIONS = ("trend", "margin", "category", "customers", "forecast")
REFRESH_DAYS = {"manual": None, "weekly": 7, "monthly": 30}
CACHE_FILE = Path(__file__).resolve().parent / "insights_cache.json"
MAX_CACHE_ENTRIES = 30
FAIL_COOLDOWN_SEC = 600  # don't re-hit a failing API on every Streamlit rerun

# Thresholds
DROP_WARN, DROP_CRIT = -0.20, -0.35
GP_PPT_WARN, GP_PPT_CRIT = -3.0, -5.0
CAT_DROP, CAT_GROWTH, CAT_MIN_SHARE = -0.20, 0.30, 0.05
CUST_DROP_WARN, CUST_DROP_CRIT, CUST_TOP_N = -0.30, -0.60, 10
CONCENTRATION = 0.30
BUDGET_PACE_GAP_PPT = 10.0
FORECAST_SHORTFALL = 0.05

_SEV_RANK = {"critical": 0, "warning": 1, "info": 2}
_failed: dict[str, float] = {}


# ── formatting / templates ──────────────────────────────────────────
def fmt_money(x: float) -> str:
    sign = "-" if x < 0 else ""
    x = abs(x)
    if x >= 1e9:
        return f"{sign}NT${x / 1e9:.2f}B"
    if x >= 1e6:
        return f"{sign}NT${x / 1e6:.1f}M"
    if x >= 1e3:
        return f"{sign}NT${x / 1e3:.0f}K"
    return f"{sign}NT${x:.0f}"


def _pct(x: float) -> str:
    return f"{x * 100:+.0f}%"


_EN_MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _months(ms: list[int], lang: str) -> str:
    """'7–8月' / 'Jul–Aug' for a consecutive run."""
    if lang == "en":
        lab = [_EN_MONTHS[m - 1] for m in ms]
        return lab[0] if len(lab) == 1 else f"{lab[0]}–{lab[-1]}"
    return f"{ms[0]}月" if len(ms) == 1 else f"{ms[0]}–{ms[-1]}月"


def _month_list(ms: list[int], lang: str) -> str:
    """'7月、8月、10月' / 'Jul, Aug, Oct' for scattered months."""
    if lang == "en":
        return ", ".join(_EN_MONTHS[m - 1] for m in ms)
    return "、".join(f"{m}月" for m in ms)


_T = {
    "zh-TW": {
        "drop_title": "{m}營收下滑 {chg}（對比{basis}）",
        "drop_detail": "較{basis}少 {gap}，期間營收 {rev}。",
        "basis_yoy": "去年同期", "basis_mom": "前一個月",
        "gp_title": "{m}毛利率偏低（最低 {gp:.1f}%）",
        "gp_detail": "低於全年平均 {avg:.1f}% 約 {gap:.1f} 個百分點。",
        "cat_down_title": "{cat} 營收年減 {chg}",
        "cat_down_detail": "累計 {rev}，去年同期 {prev}，占去年營收 {share:.0f}%。",
        "cat_up_title": "{cat} 營收年增 {chg}",
        "cat_up_detail": "累計 {rev}，去年同期 {prev}。",
        "cust_title": "{cust} 營收年減 {chg}",
        "cust_detail": "累計 {rev}，去年同期 {prev}，減少 {gap}。",
        "conc_title": "客戶集中度偏高：{cust} 占 {share:.0f}%",
        "conc_detail": "單一客戶占累計營收超過 {thr:.0f}%，需留意依賴風險。",
        "pace_title": "預算達成率落後進度",
        "pace_detail": "累計達成 {ach:.0f}%，時間進度已過 {pace:.0f}%。",
        "fc_title": "全年預測低於預算 {gap:.0f}%",
        "fc_detail": "全年預測 {fc}，全年預算 {bud}。",
        "ask_drop": "請說明 {year} 年 {m} 的營收下滑主要來自哪些客戶與產品類別。",
        "ask_gp": "請分析 {year} 年 {m} 毛利率偏低的原因（哪些類別或客戶）。",
        "ask_cat": "{year} 年 {cat} 類別各月份營收趨勢，與去年相比如何？",
        "ask_cust": "{cust} 在 {year} 年的每月營收與去年相比如何？按類別拆解。",
        "ask_conc": "{year} 年前五大客戶的營收占比各是多少？",
        "ask_budget": "{year} 年預算達成與全年預測的差距有多少？",
        "head_fallback": "{year} 年累計營收 {rev}，毛利率 {gp:.1f}%{yoy}。",
        "yoy": "，較去年 {chg}",
        "calm": "目前未偵測到明顯異常。",
        "lang_name": "Traditional Chinese (繁體中文)",
    },
    "en": {
        "drop_title": "{m} revenue down {chg} vs {basis}",
        "drop_detail": "{gap} below {basis}; revenue in the period was {rev}.",
        "basis_yoy": "last year", "basis_mom": "the prior month",
        "gp_title": "{m} gross margin low (as low as {gp:.1f}%)",
        "gp_detail": "About {gap:.1f} ppt below the year's average of {avg:.1f}%.",
        "cat_down_title": "{cat} revenue down {chg} YoY",
        "cat_down_detail": "YTD {rev} vs {prev} last year ({share:.0f}% of last year's revenue).",
        "cat_up_title": "{cat} revenue up {chg} YoY",
        "cat_up_detail": "YTD {rev} vs {prev} last year.",
        "cust_title": "{cust} revenue down {chg} YoY",
        "cust_detail": "YTD {rev} vs {prev} last year, {gap} lower.",
        "conc_title": "High customer concentration: {cust} at {share:.0f}%",
        "conc_detail": "One customer exceeds {thr:.0f}% of YTD revenue; watch dependency risk.",
        "pace_title": "Budget achievement behind pace",
        "pace_detail": "{ach:.0f}% achieved YTD with {pace:.0f}% of the year elapsed.",
        "fc_title": "Full-year forecast {gap:.0f}% below budget",
        "fc_detail": "Forecast {fc} vs budget {bud}.",
        "ask_drop": "Explain the {m} {year} revenue drop: which customers and categories drove it?",
        "ask_gp": "Why was gross margin low in {m} {year}? Which categories or customers?",
        "ask_cat": "Show monthly {cat} revenue for {year} and compare with last year.",
        "ask_cust": "Show {cust}'s monthly revenue in {year} vs last year, by category.",
        "ask_conc": "What share of {year} revenue do the top 5 customers represent?",
        "ask_budget": "How far is {year} budget achievement and the full-year forecast from budget?",
        "head_fallback": "{year} YTD revenue {rev}, GP margin {gp:.1f}%{yoy}.",
        "yoy": ", {chg} vs last year",
        "calm": "No significant anomalies detected.",
        "lang_name": "English",
    },
}
LANGS = {"zh-TW": "繁體中文", "en": "English"}


def _tr(lang: str) -> dict:
    return _T.get(lang) or _T["zh-TW"]


# ── data helpers ────────────────────────────────────────────────────
PARTIAL_MONTH_DAY = 25  # data ending before this day of a month = month still partial


def _completed_months(df: pd.DataFrame, year: int, today: datetime) -> int:
    """Completed months of *year*: bounded by the calendar AND by the loaded data.

    A month counts only once it is over and the data reaches (nearly) its end, so a
    shipping file last updated mid-May doesn't make May look like a collapse.
    """
    if year > today.year:
        return 0
    cal = 12 if year < today.year else max(today.month - 1, 0)
    dates = df.loc[df["Ship Date"].dt.year == year, "Ship Date"]
    if dates.empty:
        return cal
    last = dates.max()
    data_n = last.month if last.day >= PARTIAL_MONTH_DAY else last.month - 1
    return min(cal, data_n)


def _by_month(df: pd.DataFrame, year: int) -> pd.DataFrame:
    d = df[df["Ship Date"].dt.year == year]
    return d.groupby(d["Ship Date"].dt.month).agg(rev=(REV, "sum"), gp=(GP, "sum"))


def _runs(months: list[int]) -> list[list[int]]:
    out: list[list[int]] = []
    for m in sorted(months):
        if out and m == out[-1][-1] + 1:
            out[-1].append(m)
        else:
            out.append([m])
    return out


def _ytd(df: pd.DataFrame, year: int, upto: int) -> pd.DataFrame:
    return df[(df["Ship Date"].dt.year == year) & (df["Ship Date"].dt.month <= upto)]


# ── rule-based detection ────────────────────────────────────────────
def detect_findings(
    df: pd.DataFrame,
    year: int,
    today: Optional[datetime] = None,
    lang: str = "zh-TW",
    budget: Optional[dict] = None,
) -> list[dict]:
    """Deterministic anomaly list for *year*, most important first.

    Each finding: id, severity (critical|warning|info), section, title, detail,
    impact (TWD, for ranking) and question (a ready-made chatbot drill-down).
    *budget* (optional): {"ytd_rev", "fy_budget", "fy_forecast"} from FCST/Budget.
    """
    today = today or datetime.now()
    t = _tr(lang)
    if df is None or df.empty:
        return []
    n = _completed_months(df, year, today)
    if n == 0:
        return []
    out: list[dict] = []

    cur = _by_month(df, year)
    prev = _by_month(df, year - 1)

    # 1. Revenue drops (YoY same month when available, else MoM), clustered
    flagged: dict[int, tuple[float, float, str]] = {}
    for m in range(1, n + 1):
        if m not in cur.index:
            continue
        rev = float(cur.at[m, "rev"])
        if m in prev.index and prev.at[m, "rev"] > 0:
            base, basis = float(prev.at[m, "rev"]), "yoy"
        elif m - 1 in cur.index and cur.at[m - 1, "rev"] > 0:
            base, basis = float(cur.at[m - 1, "rev"]), "mom"
        else:
            continue
        if (rev - base) / base <= DROP_WARN:
            flagged[m] = (rev, base, basis)
    for run in _runs(list(flagged)):
        rev = sum(flagged[m][0] for m in run)
        base = sum(flagged[m][1] for m in run)
        chg = (rev - base) / base
        basis = "yoy" if all(flagged[m][2] == "yoy" for m in run) else "mom"
        label = _months(run, lang)
        out.append({
            "id": f"drop_{year}_{run[0]:02d}_{run[-1]:02d}",
            "severity": "critical" if chg <= DROP_CRIT else "warning",
            "section": "trend",
            "title": t["drop_title"].format(m=label, chg=_pct(chg), basis=t[f"basis_{basis}"]),
            "detail": t["drop_detail"].format(
                gap=fmt_money(base - rev), basis=t[f"basis_{basis}"], rev=fmt_money(rev)),
            "impact": base - rev,
            "question": t["ask_drop"].format(m=label, year=year),
        })

    # 2. Gross-margin erosion vs the year's average
    done = cur[(cur.index <= n) & (cur["rev"] > 0)]
    if len(done) >= 3:
        avg = float(done["gp"].sum() / done["rev"].sum() * 100)
        low = []
        for m, r in done.iterrows():
            pct = float(r["gp"] / r["rev"] * 100)
            if pct - avg <= GP_PPT_WARN:
                low.append((int(m), pct, avg - pct, float(r["rev"])))
        if low:
            worst = max(low, key=lambda x: x[2])
            months = [m for m, *_ in low][:3]
            label = _month_list(months, lang)
            out.append({
                "id": f"gp_{year}_" + "_".join(f"{m:02d}" for m in months),
                "severity": "critical" if worst[2] >= -GP_PPT_CRIT else "warning",
                "section": "margin",
                "title": t["gp_title"].format(m=label, gp=worst[1]),
                "detail": t["gp_detail"].format(gap=worst[2], avg=avg),
                "impact": worst[2] / 100 * worst[3],
                "question": t["ask_gp"].format(m=label, year=year),
            })

    # 3. Categories (YTD over completed months vs same months last year)
    c_cur = _ytd(df, year, n).groupby(CAT)[REV].sum()
    c_prev = _ytd(df, year - 1, n).groupby(CAT)[REV].sum()
    prev_total = float(c_prev.sum())
    ups = []
    for cat, p in c_prev.items():
        if p <= 0 or prev_total <= 0 or p / prev_total < CAT_MIN_SHARE:
            continue
        c = float(c_cur.get(cat, 0.0))
        chg = (c - p) / p
        if chg <= CAT_DROP:
            out.append({
                "id": f"cat_down_{year}_{cat}",
                "severity": "critical" if chg <= DROP_CRIT else "warning",
                "section": "category",
                "title": t["cat_down_title"].format(cat=cat, chg=_pct(chg)),
                "detail": t["cat_down_detail"].format(
                    rev=fmt_money(c), prev=fmt_money(p), share=p / prev_total * 100),
                "impact": p - c,
                "question": t["ask_cat"].format(cat=cat, year=year),
            })
        elif chg >= CAT_GROWTH:
            ups.append((cat, c, float(p), chg))
    if ups:
        cat, c, p, chg = max(ups, key=lambda x: x[1] - x[2])
        out.append({
            "id": f"cat_up_{year}_{cat}", "severity": "info", "section": "category",
            "title": t["cat_up_title"].format(cat=cat, chg=_pct(chg)),
            "detail": t["cat_up_detail"].format(rev=fmt_money(c), prev=fmt_money(p)),
            "impact": c - p,
            "question": t["ask_cat"].format(cat=cat, year=year),
        })

    # 4. Customers: top accounts last year that are shrinking, plus concentration
    k_cur = _ytd(df, year, n).groupby(CUST)[REV].sum()
    k_prev = _ytd(df, year - 1, n).groupby(CUST)[REV].sum().sort_values(ascending=False)
    hits = []
    for cust, p in k_prev.head(CUST_TOP_N).items():
        if p <= 0:
            continue
        c = float(k_cur.get(cust, 0.0))
        chg = (c - p) / p
        if chg <= CUST_DROP_WARN:
            hits.append((cust, c, float(p), chg))
    for cust, c, p, chg in sorted(hits, key=lambda x: x[1] - x[2])[:3]:
        out.append({
            "id": f"cust_{year}_{cust}",
            "severity": "critical" if chg <= CUST_DROP_CRIT else "warning",
            "section": "customers",
            "title": t["cust_title"].format(cust=cust, chg=_pct(chg)),
            "detail": t["cust_detail"].format(
                rev=fmt_money(c), prev=fmt_money(p), gap=fmt_money(p - c)),
            "impact": p - c,
            "question": t["ask_cust"].format(cust=cust, year=year),
        })
    if len(k_cur) and k_cur.sum() > 0:
        top = k_cur.idxmax()
        share = float(k_cur.max() / k_cur.sum())
        if share >= CONCENTRATION:
            out.append({
                "id": f"conc_{year}_{top}", "severity": "warning", "section": "customers",
                "title": t["conc_title"].format(cust=top, share=share * 100),
                "detail": t["conc_detail"].format(thr=CONCENTRATION * 100),
                "impact": float(k_cur.max()) * 0.01,
                "question": t["ask_conc"].format(year=year),
            })

    # 5. Budget pace and full-year forecast vs budget
    b = budget or {}
    fy_budget = float(b.get("fy_budget") or 0)
    if fy_budget > 0:
        ach = float(b.get("ytd_rev") or 0) / fy_budget * 100
        pace = n / 12 * 100
        if ach < pace - BUDGET_PACE_GAP_PPT:
            out.append({
                "id": f"pace_{year}", "severity": "warning", "section": "forecast",
                "title": t["pace_title"],
                "detail": t["pace_detail"].format(ach=ach, pace=pace),
                "impact": (pace - ach) / 100 * fy_budget,
                "question": t["ask_budget"].format(year=year),
            })
        fc = float(b.get("fy_forecast") or 0)
        if fc > 0 and fc < fy_budget * (1 - FORECAST_SHORTFALL):
            gap = (fy_budget - fc) / fy_budget * 100
            out.append({
                "id": f"fc_{year}", "severity": "critical" if gap >= 15 else "warning",
                "section": "forecast",
                "title": t["fc_title"].format(gap=gap),
                "detail": t["fc_detail"].format(fc=fmt_money(fc), bud=fmt_money(fy_budget)),
                "impact": fy_budget - fc,
                "question": t["ask_budget"].format(year=year),
            })

    out.sort(key=lambda f: (_SEV_RANK[f["severity"]], -abs(f["impact"])))
    return out


def build_facts(df: pd.DataFrame, year: int, today: Optional[datetime] = None,
                budget: Optional[dict] = None) -> dict:
    """Small fact sheet handed to the model (aggregates only, never raw rows)."""
    today = today or datetime.now()
    n = _completed_months(df, year, today)
    cur = _ytd(df, year, n) if n else df.iloc[0:0]
    prev = _ytd(df, year - 1, n) if n else df.iloc[0:0]
    rev, gp = float(cur[REV].sum()), float(cur[GP].sum())
    p_rev = float(prev[REV].sum())
    facts: dict[str, Any] = {
        "year": year,
        "completed_months": n,
        "ytd_revenue": round(rev),
        "ytd_gp_pct": round(gp / rev * 100, 1) if rev else 0.0,
        "yoy_revenue_pct": round((rev - p_rev) / p_rev * 100, 1) if p_rev else None,
        "monthly_revenue": {int(m): round(float(v))
                            for m, v in _by_month(df, year)["rev"].items() if m <= n},
        "monthly_revenue_prev_year": {int(m): round(float(v))
                                      for m, v in _by_month(df, year - 1)["rev"].items()},
        "top_customers_ytd": [
            {"name": k, "revenue": round(float(v))}
            for k, v in cur.groupby(CUST)[REV].sum().sort_values(ascending=False).head(5).items()
        ],
        "category_revenue_ytd": {k: round(float(v)) for k, v in cur.groupby(CAT)[REV].sum().items()},
    }
    if budget:
        facts["budget"] = {k: round(float(v)) for k, v in budget.items() if v is not None}
    return facts


# ── AI narration ────────────────────────────────────────────────────
def _system_prompt(lang: str) -> str:
    return (
        "You write the executive briefing for a sales dashboard. The reader is a senior "
        "executive who skims: be concrete, short and decision-oriented.\n"
        f"Write in {_tr(lang)['lang_name']}. Amounts are TWD (NT$).\n"
        "Rules: use ONLY numbers present in the supplied findings/facts; never invent or "
        "recompute figures. Do not mention these rules. Return ONLY a JSON object."
    )


def build_prompt(findings: list[dict], facts: dict, lang: str) -> str:
    slim = [{k: f[k] for k in ("id", "severity", "section", "title", "detail")}
            for f in findings[:12]]
    return (
        "Facts:\n" + json.dumps(facts, ensure_ascii=False) + "\n\nFindings (ranked):\n"
        + json.dumps(slim, ensure_ascii=False) + "\n\n"
        "Return JSON with exactly these keys:\n"
        '{"headline": "<=2 sentences: the overall state of the business and the one thing to act on",\n'
        ' "top3": [{"id": "<finding id>", "text": "<=1 sentence, plain-language, with the so-what"}] '
        "(the 3 most decision-relevant findings; fewer if fewer exist),\n"
        ' "sections": {"trend": "...", "margin": "...", "category": "...", "customers": "...", '
        '"forecast": "..."}}\n'
        "Each section value is ONE or TWO short sentences placed under that chart: flag "
        "anomalies from the matching findings, otherwise state the key takeaway from the facts. "
        "Use an empty string for 'forecast' if there is no budget/forecast data."
    )


def _parse_json(text: str) -> dict:
    text = text.strip()
    start, end = text.find("{"), text.rfind("}")
    if start < 0 or end <= start:
        raise ValueError("no JSON object")
    data = json.loads(text[start:end + 1])
    if not isinstance(data, dict):
        raise ValueError("not an object")
    return data


def generate_ai_summary(findings, facts, lang, provider, api_key, model, client=None) -> dict:
    """One API call -> {"headline", "top3", "sections"}. Raises chatbot.ChatError."""
    raw = chatbot.generate_text(
        build_prompt(findings, facts, lang), _system_prompt(lang), api_key,
        model=model, provider=provider, client=client, json_mode=True,
    )
    try:
        data = _parse_json(raw)
    except ValueError as exc:  # JSONDecodeError is a ValueError
        raise chatbot.ChatError("The AI summary could not be parsed. Try refreshing.") from exc
    ids = {f["id"] for f in findings}
    top3 = [
        {"id": x["id"], "text": str(x.get("text") or "").strip()}
        for x in (data.get("top3") or [])
        if isinstance(x, dict) and x.get("id") in ids
    ][:3]
    secs = data.get("sections") if isinstance(data.get("sections"), dict) else {}
    return {
        "headline": str(data.get("headline") or "").strip(),
        "top3": top3,
        "sections": {s: str(secs.get(s) or "").strip() for s in SECTIONS},
    }


def rule_summary(findings: list[dict], facts: dict, lang: str) -> dict:
    """No-API fallback built from the findings themselves."""
    t = _tr(lang)
    yoy = facts.get("yoy_revenue_pct")
    headline = t["head_fallback"].format(
        year=facts["year"], rev=fmt_money(facts["ytd_revenue"]), gp=facts["ytd_gp_pct"],
        yoy=t["yoy"].format(chg=f"{yoy:+.1f}%") if yoy is not None else "",
    )
    if not findings:
        headline += " " + t["calm"]
    stop = "." if lang == "en" else "。"
    sections = {
        s: " ".join(f["title"] + stop for f in [x for x in findings if x["section"] == s][:2])
        for s in SECTIONS
    }
    return {
        "headline": headline,
        "top3": [{"id": f["id"], "text": f["detail"]} for f in findings[:3]],
        "sections": sections,
    }


def compose(findings, facts, lang, ai: Optional[dict]) -> dict:
    """AI text where present, rule-based text for anything the model left out."""
    base = rule_summary(findings, facts, lang)
    if not ai:
        return base
    top3 = [x for x in ai.get("top3", []) if x.get("text")] or base["top3"]
    ai_secs = ai.get("sections") or {}
    return {
        "headline": ai.get("headline") or base["headline"],
        "top3": top3,
        "sections": {s: ai_secs.get(s) or base["sections"][s] for s in SECTIONS},
    }


# ── cache ───────────────────────────────────────────────────────────
def fingerprint(df: pd.DataFrame, year: int, lang: str, today: datetime,
                budget: Optional[dict]) -> str:
    d = df[df["Ship Date"].dt.year.isin([year, year - 1])]
    parts = [
        year, lang, _completed_months(df, year, today), len(d),
        round(float(d[REV].sum())), round(float(d[GP].sum())), str(d["Ship Date"].max()),
        json.dumps({k: round(float(v or 0)) for k, v in (budget or {}).items()}, sort_keys=True),
    ]
    return hashlib.sha1("|".join(map(str, parts)).encode()).hexdigest()[:16]


def load_cache(path: Optional[Path] = None) -> dict:
    try:
        with open(path or CACHE_FILE, encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def save_cache(cache: dict, path: Optional[Path] = None) -> None:
    path = Path(path or CACHE_FILE)
    if len(cache) > MAX_CACHE_ENTRIES:
        newest = sorted(cache.items(), key=lambda kv: kv[1].get("generated_at", ""), reverse=True)
        cache = dict(newest[:MAX_CACHE_ENTRIES])
    try:
        tmp = path.with_suffix(".tmp")
        tmp.write_text(json.dumps(cache, ensure_ascii=False, indent=1), encoding="utf-8")
        os.replace(tmp, path)
    except OSError:
        _logger.warning("Could not write insights cache", exc_info=True)


def is_fresh(entry: dict, refresh: str, now: Optional[datetime] = None) -> bool:
    days = REFRESH_DAYS.get(refresh, REFRESH_DAYS["weekly"])
    if days is None:  # manual: only a data change (new fingerprint) invalidates
        return True
    try:
        age = (now or datetime.now()) - datetime.fromisoformat(entry["generated_at"])
    except (KeyError, ValueError):
        return False
    return age.total_seconds() < days * 86400


# ── public entry point ──────────────────────────────────────────────
def get_insights(
    df: pd.DataFrame,
    year: int,
    *,
    settings: Optional[dict] = None,
    today: Optional[datetime] = None,
    budget: Optional[dict] = None,
    force: bool = False,
    allow_generate: bool = True,
    client: Any = None,
    cache_path: Optional[Path] = None,
) -> dict:
    """Findings + executive summary for the dashboard.

    Returns {"findings", "summary", "source" ("ai"|"rules"), "generated_at", "error"}.
    Uses the cache when fresh; otherwise (and with an API key) regenerates once.
    """
    settings = settings or {}
    today = today or datetime.now()
    lang = settings.get("insights_lang") or "zh-TW"
    refresh = settings.get("insights_refresh") or "weekly"
    findings = detect_findings(df, year, today, lang, budget)
    facts = build_facts(df, year, today, budget)
    result: dict[str, Any] = {
        "findings": findings, "source": "rules", "generated_at": None, "error": None,
    }

    provider, key, model = chatbot.provider_config(settings)
    fp = fingerprint(df, year, lang, today, budget)
    cache = load_cache(cache_path)
    entry = cache.get(fp)
    ai = None

    if entry and not force and is_fresh(entry, refresh, today):
        ai, result["generated_at"] = entry.get("ai"), entry.get("generated_at")
    elif allow_generate and (key or client is not None):
        if not force and time.time() - _failed.get(fp, 0) < FAIL_COOLDOWN_SEC:
            result["error"] = "cooldown"
            ai = entry.get("ai") if entry else None  # stale beats nothing
        else:
            try:
                ai = generate_ai_summary(findings, facts, lang, provider, key, model, client)
                stamp = today.isoformat(timespec="seconds")
                cache[fp] = {"generated_at": stamp, "ai": ai, "provider": provider, "model": model}
                save_cache(cache, cache_path)
                result["generated_at"] = stamp
                _failed.pop(fp, None)
            except chatbot.ChatError as exc:
                _failed[fp] = time.time()
                result["error"] = str(exc)
                ai = entry.get("ai") if entry else None
                if ai:
                    result["generated_at"] = entry.get("generated_at")
    elif entry:  # no key available but an older summary exists
        ai, result["generated_at"] = entry.get("ai"), entry.get("generated_at")

    result["summary"] = compose(findings, facts, lang, ai)
    result["source"] = "ai" if ai else "rules"
    return result
