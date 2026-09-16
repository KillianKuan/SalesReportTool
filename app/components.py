"""components.py — Neutral KPI/card-title UI helpers.

Markup and layout only. Colors are either inherited from Streamlit (text,
backgrounds) or, for the positive/negative delta pills, a fixed semantic
accent from palette.py — never a Light/Dark palette. See inject_layout_css()
in utils.py for the CSS classes these helpers render into.
"""

import streamlit as st

from palette import NEGATIVE, NEGATIVE_BG, POSITIVE, POSITIVE_BG, MUTED_BG


# Streamlit's ":material/<name>:" shorthand is only expanded in plain
# markdown text — embedding it inside a raw HTML string passed with
# unsafe_allow_html=True renders as garbled literal text instead of the
# icon glyph. So the icon (native, plain st.markdown) and the custom-styled
# label (raw HTML div, text only) are rendered as two adjacent elements
# rather than concatenated into one HTML string.
def card_title(text: str, icon: str | None = None) -> None:
    """Small (13px/600) per-card heading, muted-to-text color."""
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
