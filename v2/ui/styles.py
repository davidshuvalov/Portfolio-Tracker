"""
Professional branding assets and CSS for Portfolio Tracker.

Usage:
    from ui.styles import inject_styles, render_logo, render_sidebar_logo
    inject_styles()          # call once per page, after set_page_config
    render_logo()            # home-page hero logo
    render_sidebar_logo()    # compact sidebar logo
"""
from __future__ import annotations

import streamlit as st

# ── Logo SVG — home-page hero (320 × 80) ─────────────────────────────────────
# Four ascending bars (portfolio strategies) with a gold trend line.
LOGO_FULL = """\
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 340 80">
  <rect x="2"  y="52" width="12" height="24" rx="2" fill="#2563eb" opacity="0.35"/>
  <rect x="18" y="38" width="12" height="38" rx="2" fill="#2563eb" opacity="0.55"/>
  <rect x="34" y="22" width="12" height="54" rx="2" fill="#2563eb" opacity="0.78"/>
  <rect x="50" y="8"  width="12" height="68" rx="2" fill="#3b82f6"/>
  <polyline points="8,52 24,38 40,22 56,8"
    stroke="#d4a843" stroke-width="2.5" fill="none"
    stroke-linecap="round" stroke-linejoin="round"/>
  <circle cx="56" cy="8" r="3.5" fill="#d4a843"/>
  <text x="82" y="37"
    font-family="system-ui,sans-serif"
    font-size="22" font-weight="700" letter-spacing="2"
    fill="#dde3ee">PORTFOLIO TRACKER</text>
  <text x="83" y="56"
    font-family="system-ui,sans-serif"
    font-size="11" font-weight="400" letter-spacing="3.5"
    fill="#94a3b8">A TOOL FOR MULTIWALK</text>
  <line x1="82" y1="62" x2="334" y2="62" stroke="#1e3a6e" stroke-width="0.75"/>
</svg>"""

# ── Logo SVG — sidebar compact (240 × 52) ────────────────────────────────────
LOGO_SIDEBAR = """\
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 234 52">
  <rect x="2"  y="36" width="9" height="13" rx="1.5" fill="#2563eb" opacity="0.35"/>
  <rect x="14" y="27" width="9" height="22" rx="1.5" fill="#2563eb" opacity="0.55"/>
  <rect x="26" y="16" width="9" height="33" rx="1.5" fill="#2563eb" opacity="0.78"/>
  <rect x="38" y="6"  width="9" height="43" rx="1.5" fill="#3b82f6"/>
  <polyline points="6.5,36 18.5,27 30.5,16 42.5,6"
    stroke="#d4a843" stroke-width="2" fill="none"
    stroke-linecap="round" stroke-linejoin="round"/>
  <circle cx="42.5" cy="6" r="2.5" fill="#d4a843"/>
  <text x="58" y="24"
    font-family="system-ui,sans-serif"
    font-size="13" font-weight="700" letter-spacing="1.5"
    fill="#dde3ee">PORTFOLIO TRACKER</text>
  <text x="59" y="38"
    font-family="system-ui,sans-serif"
    font-size="8.5" font-weight="400" letter-spacing="2.5"
    fill="#94a3b8">A TOOL FOR MULTIWALK</text>
</svg>"""

# ── Professional CSS ──────────────────────────────────────────────────────────
_CSS = """<style>
/* Layout */
.main .block-container {
    padding-top: 1.5rem;
    padding-bottom: 3rem;
    max-width: 1440px;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #060c1a !important;
    border-right: 1px solid #101e35 !important;
}

/* Metric cards */
[data-testid="stMetric"] {
    background: #0d1626;
    border: 1px solid #192840;
    border-radius: 10px;
    padding: 1rem 1.25rem;
}
[data-testid="stMetricValue"] {
    font-size: 1.55rem !important;
    font-weight: 700 !important;
    letter-spacing: -0.02em !important;
}
[data-testid="stMetricLabel"] {
    font-size: 0.72rem !important;
    text-transform: uppercase !important;
    letter-spacing: 0.09em !important;
    color: #64748b !important;
}

/* Bordered containers (step cards, analytics cards) */
[data-testid="stVerticalBlockBorderWrapper"] {
    border: 1px solid #192840 !important;
    border-radius: 12px !important;
    background: #08101f !important;
    transition: border-color 0.15s ease, background 0.15s ease;
}
[data-testid="stVerticalBlockBorderWrapper"]:hover {
    border-color: #2563eb !important;
    background: #0a1428 !important;
}

/* Dividers */
hr {
    border-color: #131f33 !important;
    margin: 1.75rem 0 !important;
}

/* Expanders */
[data-testid="stExpander"] {
    border: 1px solid #192840 !important;
    border-radius: 10px !important;
    background: #080e1c !important;
}

/* DataFrames */
[data-testid="stDataFrame"] {
    border: 1px solid #192840;
    border-radius: 8px;
    overflow: hidden;
}

/* Primary buttons */
.stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #1d4ed8 0%, #2563eb 100%) !important;
    border: none !important;
    font-weight: 600 !important;
    letter-spacing: 0.04em !important;
    box-shadow: 0 2px 8px rgba(37,99,235,0.25) !important;
    border-radius: 8px !important;
}
.stButton > button[kind="primary"]:hover {
    background: linear-gradient(135deg, #1e40af 0%, #1d4ed8 100%) !important;
    box-shadow: 0 4px 16px rgba(37,99,235,0.40) !important;
}

/* Secondary buttons */
.stButton > button[kind="secondary"] {
    border: 1px solid #192840 !important;
    border-radius: 8px !important;
    font-weight: 500 !important;
}
.stButton > button[kind="secondary"]:hover {
    border-color: #2563eb !important;
    color: #60a5fa !important;
}

/* Alert banners */
[data-testid="stAlert"] {
    border-radius: 8px !important;
}

/* Page links */
[data-testid="stPageLink"] a {
    font-weight: 500 !important;
    font-size: 0.88rem !important;
}
</style>"""


def inject_styles() -> None:
    """Inject professional CSS into the current Streamlit page."""
    st.markdown(_CSS, unsafe_allow_html=True)


def render_logo() -> None:
    """Render the full-size logo SVG (for the home page hero)."""
    st.markdown(
        f'<div style="margin: 0.5rem 0 0.25rem 0">{LOGO_FULL}</div>',
        unsafe_allow_html=True,
    )


def render_sidebar_logo() -> None:
    """Render the compact logo SVG in the sidebar."""
    st.markdown(
        f'<div style="padding: 0.75rem 0 0.25rem 0">{LOGO_SIDEBAR}</div>',
        unsafe_allow_html=True,
    )
