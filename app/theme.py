"""theme.py — Single source of design tokens (colors, radii, spacing, font sizes).

charts.py, app.py, and utils.py must import from here rather than hardcoding
colors. See CLAUDE.md / README for the v4.0 UI redesign notes.
"""

import colorsys

import streamlit as st

# ── Design tokens ────────────────────────────────────────────────
_LIGHT = {
    "primary": "#1F5E46",
    "primary_hover": "#2D7A5C",
    "accent": "#7FBF9E",
    "canvas": "#E8F3ED",
    "surface": "#FFFFFF",
    "border": "#D6E5DC",
    "text": "#1C2B24",
    "muted": "#6B7C74",
    "positive": "#2E9E6B",
    "negative": "#D9534F",
}
_DARK = {
    "primary": "#7FBF9E",
    "primary_hover": "#9BD1B5",
    "accent": "#A8DCC0",
    "canvas": "#111A16",
    "surface": "#1A2621",
    "border": "#2C3A34",
    "text": "#EAF2ED",
    "muted": "#9AAAA2",
    "positive": "#5FC694",
    "negative": "#E8756F",
}
_COMMON = {
    "radius": "12px",
    "card_padding": "16px",
    "section_gap": "24px",
}

# Sidebar chrome always uses the (dark-green) light-mode primary as its
# background, with white text — it does not flip with light/dark mode.
SIDEBAR_BG = _LIGHT["primary"]


def _hex_to_rgb(hex_color: str) -> tuple:
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


def _hex_to_rgba(hex_color: str, alpha: float) -> str:
    r, g, b = _hex_to_rgb(hex_color)
    return f"rgba({r}, {g}, {b}, {alpha})"


def get_tokens(mode: str) -> dict:
    """Return the full token dict for *mode* ('light' or 'dark').

    Includes derived low-alpha tints (``*_bg``) for delta pills etc., so
    callers never need to compute rgba() blends themselves.
    """
    base = _DARK if mode == "dark" else _LIGHT
    tokens = {**base, **_COMMON}
    tokens["positive_bg"] = _hex_to_rgba(tokens["positive"], 0.16)
    tokens["negative_bg"] = _hex_to_rgba(tokens["negative"], 0.16)
    tokens["muted_bg"] = _hex_to_rgba(tokens["muted"], 0.16)
    return tokens


# ── Category palette ──────────────────────────────────────────────
_CATEGORY_COLORS_LIGHT = {
    "CDR": "#1F5E46",
    "CDR ACC": "#7FBF9E",
    "Tablet": "#C97B3C",
    "Tablet ACC": "#E8B583",
    "AI_SW": "#3D7EA6",
    "Signify": "#8A6BA8",
    "Others": "#9AA5A0",
}


def _lighten(hex_color: str, amount: float = 0.13) -> str:
    """Lighten a hex color by *amount* of HSL lightness, keeping hue/sat."""
    r, g, b = (v / 255 for v in _hex_to_rgb(hex_color))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = min(1.0, l + amount)
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return "#{:02X}{:02X}{:02X}".format(round(r * 255), round(g * 255), round(b * 255))


def CATEGORY_COLORS(mode: str) -> dict:
    """Category -> hex color. Same keys/order for every mode."""
    if mode == "dark":
        return {k: _lighten(v) for k, v in _CATEGORY_COLORS_LIGHT.items()}
    return dict(_CATEGORY_COLORS_LIGHT)


def SOURCE_COLORS(mode: str) -> dict:
    """Actual/Forecast/Budget -> {'color': hex, 'dash': [a, b]}."""
    t = get_tokens(mode)
    return {
        "Actual":   {"color": t["primary"], "dash": [1, 0]},
        "Forecast": {"color": t["accent"],  "dash": [6, 3]},
        "Budget":   {"color": t["muted"],   "dash": [4, 4]},
    }


# ── Components ─────────────────────────────────────────────────────
# Streamlit's ":material/<name>:" shorthand is only expanded in plain
# markdown text — embedding it inside a raw HTML string passed with
# unsafe_allow_html=True renders as garbled literal text instead of the
# icon glyph. So the icon (native, plain st.markdown) and the custom-styled
# label (raw HTML div, text only) are rendered as two adjacent elements
# rather than concatenated into one HTML string.
def card_title(text: str, icon: str | None = None) -> None:
    """Small (13px/600) per-card heading, muted-to-text color.

    Replaces the old ``st.markdown("**...**")`` double-title pattern.
    """
    if icon:
        _icon_col, _text_col = st.columns([0.04, 0.96], gap="small")
        with _icon_col:
            st.markdown(f":material/{icon}:")
        with _text_col:
            st.markdown(f'<div class="sr-card-title">{text}</div>', unsafe_allow_html=True)
    else:
        st.markdown(f'<div class="sr-card-title">{text}</div>', unsafe_allow_html=True)


def kpi_card(label, value, delta=None, delta_positive=None, caption=None, icon=None) -> None:
    """Custom KPI tile replacing ``st.metric``.

    Uppercase muted label, 28px/600 value, an optional delta pill
    (positive/negative/neutral "N/A"), and an optional caption line —
    same semantics as the ``st.metric(..., delta=...)`` calls it replaces.
    """
    if icon:
        st.markdown(f":material/{icon}:")
    delta_html = ""
    if delta is not None:
        is_na = str(delta).strip().upper() == "N/A"
        if is_na:
            pill_cls = "sr-kpi-delta-na"
        elif delta_positive is None:
            pill_cls = "sr-kpi-delta-neg" if str(delta).strip().startswith("-") else "sr-kpi-delta-pos"
        else:
            pill_cls = "sr-kpi-delta-pos" if delta_positive else "sr-kpi-delta-neg"
        delta_html = f'<span class="sr-kpi-delta {pill_cls}">{delta}</span>'
    caption_html = f'<div class="sr-kpi-caption">{caption}</div>' if caption else ""
    st.markdown(
        f"""<div class="sr-kpi-card">
    <div class="sr-kpi-label">{label}</div>
    <div class="sr-kpi-value">{value}{delta_html}</div>
    {caption_html}
</div>""",
        unsafe_allow_html=True,
    )
