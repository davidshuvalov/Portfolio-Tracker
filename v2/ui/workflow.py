"""
Shared workflow status helpers.

Used by the home page and sidebar of core workflow pages to show
order-of-operations progress and provide clickable navigation.
"""
from __future__ import annotations

import streamlit as st

# Step definitions: (session_state_key_or_check, label, page_path, description)
_STEPS = [
    (
        "folders",
        "Add Folders",
        "ui/pages/01_Import.py",
        "Point the app at your MultiWalk strategy folders",
    ),
    (
        "data",
        "Import Data",
        "ui/pages/01_Import.py",
        "Scan and load all strategy CSV files",
    ),
    (
        "strategies",
        "Review Strategies",
        "ui/pages/02_Strategies.py",
        "Set status, contracts, symbol and sector for each strategy",
    ),
    (
        "portfolio",
        "Build Portfolio",
        "ui/pages/03_Portfolio.py",
        "Aggregate Live strategies and view portfolio metrics",
    ),
]


def step_status() -> dict[str, bool]:
    """Return completion status for each workflow step."""
    config = st.session_state.get("config")

    has_folders = bool(config and getattr(config, "folders", None))
    has_data = st.session_state.get("imported_data") is not None

    has_live = False
    if has_data:
        try:
            from core.portfolio.strategies import load_strategies
            strategies = load_strategies()
            has_live = any(s.get("status") == "Live" for s in strategies) if strategies else False
        except Exception:
            pass

    has_portfolio = st.session_state.get("portfolio_data") is not None

    return {
        "folders": has_folders,
        "data": has_data,
        "strategies": has_live,
        "portfolio": has_portfolio,
    }


def render_workflow_sidebar() -> None:
    """Render compact workflow progress in the sidebar."""
    status = step_status()
    steps = [
        ("folders",    "Add Folders",        "ui/pages/01_Import.py"),
        ("data",       "Import Data",        "ui/pages/01_Import.py"),
        ("strategies", "Review Strategies",  "ui/pages/02_Strategies.py"),
        ("portfolio",  "Build Portfolio",    "ui/pages/03_Portfolio.py"),
    ]

    st.markdown(
        '<p style="color:#64748b;font-size:0.72rem;text-transform:uppercase;'
        'letter-spacing:0.12em;margin:0 0 0.4rem 0">Setup</p>',
        unsafe_allow_html=True,
    )

    for key, label, page in steps:
        done = status[key]
        st.page_link(page, label=f"{'✅' if done else '⭕'} {label}")

    st.divider()

    # Analytics links — shown once portfolio is built
    if status["portfolio"]:
        st.markdown(
            '<p style="color:#64748b;font-size:0.72rem;text-transform:uppercase;'
            'letter-spacing:0.12em;margin:0 0 0.4rem 0">Analytics</p>',
            unsafe_allow_html=True,
        )
        _analytics = [
            ("Monte Carlo",          "ui/pages/_04_Monte_Carlo.py"),
            ("Correlations",         "ui/pages/_05_Correlations.py"),
            ("Diversification",      "ui/pages/_06_Diversification.py"),
            ("Leave One Out",        "ui/pages/_07_Leave_One_Out.py"),
            ("Backtest",             "ui/pages/_08_Backtest.py"),
            ("Eligibility Backtest", "ui/pages/_09_Eligibility_Backtest.py"),
            ("Margin Tracking",      "ui/pages/_10_Margin_Tracking.py"),
            ("Position Check",       "ui/pages/_11_Position_Check.py"),
        ]
        for _label, _page in _analytics:
            st.page_link(_page, label=_label)
        st.divider()
