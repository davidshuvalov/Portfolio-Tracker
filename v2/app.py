"""
Portfolio Tracker v2 — Streamlit entrypoint.

Handles:
- First-run setup (license, folder config, optional v1.24 migration)
- Session state initialisation
- License gate (DEV_MODE bypasses during development)
"""

import streamlit as st
from pathlib import Path

# ── Dev mode bypass — remove before release ───────────────────────────────────
DEV_MODE = True

# ── Page config (must be first Streamlit call) ────────────────────────────────
st.set_page_config(
    page_title="Portfolio Tracker v2",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

from core.config import AppConfig

# ── Session state defaults ────────────────────────────────────────────────────
if "config" not in st.session_state:
    st.session_state.config = AppConfig.load()

if "imported_data" not in st.session_state:
    st.session_state.imported_data = None

if "portfolio_data" not in st.session_state:
    st.session_state.portfolio_data = None


# ── License check ─────────────────────────────────────────────────────────────
def _check_license() -> bool:
    if DEV_MODE:
        return True
    try:
        from core.licensing.license_manager import validate_full
        customer_id = st.session_state.config.customer_id
        if not customer_id:
            return False
        valid, message = validate_full(customer_id)
        if not valid:
            st.error(f"License error: {message}")
        return valid
    except Exception as e:
        st.error(f"License check failed: {e}")
        return False


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if not _check_license():
        st.stop()

    config: AppConfig = st.session_state.config

    # Sidebar: folder status + quick stats
    with st.sidebar:
        st.title("Portfolio Tracker")
        st.caption("v2.0.0")

        if DEV_MODE:
            st.warning("DEV MODE — license check disabled")

        st.divider()

        if st.session_state.imported_data is not None:
            data = st.session_state.imported_data
            n_strats = len(data.strategy_names)
            start, end = data.date_range
            st.metric("Strategies loaded", n_strats)
            st.caption(f"{start} → {end}")
        else:
            st.info("No data loaded. Go to **Import** to get started.")

        st.divider()
        st.caption(f"{len(config.folders)} base folder(s) configured")

    # Home page content
    st.title("Portfolio Tracker v2")
    st.markdown("""
    Welcome. Use the pages in the sidebar to:

    | Page | Description |
    |------|-------------|
    | **Import** | Scan MultiWalk folders and load strategy data |
    | **Strategies** | Configure strategy status, sector, symbol, contracts |
    | **Portfolio** | View aggregated portfolio metrics |
    | **Monte Carlo** | Run MC simulations with risk-of-ruin targeting |
    | **Correlations** | Analyse strategy correlations (3 modes) |
    | **Diversification** | Find optimal strategy combinations |
    | **Leave One Out** | Portfolio sensitivity analysis |
    | **Backtest** | Historical period performance |
    | **Eligibility Backtest** | Walk-forward portfolio construction rules |
    | **Margin Tracking** | Contract margin by symbol and sector |
    | **Position Check** | Current live positions |
    """)

    if not config.folders:
        st.info(
            "**Getting started:** Go to the **Import** page to add your "
            "MultiWalk base folders."
        )


if __name__ == "__main__":
    main()
