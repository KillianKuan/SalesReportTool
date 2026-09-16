"""palette.py — Fixed, theme-independent colors for chart marks and semantic
(positive/negative/neutral) UI accents.

These are NOT a Light/Dark palette: there is a single value per key, used
unchanged under Streamlit's native light or dark theme. Chart marks need an
explicit color (Vega-Lite can't inherit CSS variables), so they stay here as
plain hex/rgba. Any color that *can* inherit from Streamlit (card/text/
background chrome) must not be defined here — see components.py.
"""


def _hex_to_rgb(hex_color: str) -> tuple:
    hex_color = hex_color.lstrip("#")
    return tuple(int(hex_color[i:i + 2], 16) for i in (0, 2, 4))


def _hex_to_rgba(hex_color: str, alpha: float) -> str:
    r, g, b = _hex_to_rgb(hex_color)
    return f"rgba({r}, {g}, {b}, {alpha})"


# ── Brand / chart accent colors ─────────────────────────────────────
PRIMARY = "#1F5E46"
ACCENT = "#7FBF9E"
MUTED = "#6B7C74"
POSITIVE = "#2E9E6B"
NEGATIVE = "#D9534F"

POSITIVE_BG = _hex_to_rgba(POSITIVE, 0.16)
NEGATIVE_BG = _hex_to_rgba(NEGATIVE, 0.16)
MUTED_BG = _hex_to_rgba(MUTED, 0.16)

# ── Category palette ──────────────────────────────────────────────
CATEGORY_COLORS = {
    "CDR": "#1F5E46",
    "CDR ACC": "#7FBF9E",
    "Tablet": "#C97B3C",
    "Tablet ACC": "#E8B583",
    "AI_SW": "#3D7EA6",
    "Signify": "#8A6BA8",
    "Others": "#9AA5A0",
}

# ── Actual / Forecast / Budget source palette ───────────────────────
SOURCE_COLORS = {
    "Actual":   {"color": PRIMARY, "dash": [1, 0]},
    "Forecast": {"color": ACCENT,  "dash": [6, 3]},
    "Budget":   {"color": MUTED,   "dash": [4, 4]},
}
