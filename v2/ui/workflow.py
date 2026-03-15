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
    labels = {
        "folders": "Add folders",
        "data": "Import data",
        "strategies": "Review strategies",
        "portfolio": "Build portfolio",
    }
    pages = {
        "folders": "ui/pages/01_Import.py",
        "data": "ui/pages/01_Import.py",
        "strategies": "ui/pages/02_Strategies.py",
        "portfolio": "ui/pages/03_Portfolio.py",
    }

    st.markdown("**Setup**")
    all_done = all(status.values())
    for key, label in labels.items():
        done = status[key]
        icon = "✅" if done else "⬜"
        if not done:
            st.page_link(pages[key], label=f"{icon} {label}")
        else:
            st.markdown(f"{icon} {label}")
    if all_done:
        st.caption("Ready to explore analytics")
    st.divider()
