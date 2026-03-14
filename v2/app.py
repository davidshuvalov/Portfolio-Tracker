"""
Portfolio Tracker v2 — Streamlit entrypoint.

Handles:
- Session state initialisation
- License gate (checks MultiWalk DLL; shows customer-ID entry if not configured)
- Sidebar with data status and nav hints
"""

import streamlit as st

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
    config: AppConfig = st.session_state.config
    customer_id = config.customer_id

    if not customer_id:
        _show_license_entry("Enter your TradeStation Customer ID to activate Portfolio Tracker.")
        return False

    try:
        from core.licensing.license_manager import validate_full
        valid, message = validate_full(customer_id)
    except Exception as e:
        valid, message = False, str(e)

    if not valid:
        _show_license_entry(f"License check failed: {message}")
        return False

    return True


def _show_license_entry(prompt: str) -> None:
    st.title("Portfolio Tracker v2 — License Required")
    st.warning(prompt)
    st.markdown(
        "Enter your **TradeStation Customer ID** (the number you use to log in to "
        "TradeStation). This is verified against the MultiWalk license DLL."
    )
    with st.form("license_form"):
        cid = st.number_input(
            "TradeStation Customer ID",
            min_value=1,
            max_value=9_999_999,
            value=0,
            step=1,
        )
        submitted = st.form_submit_button("Activate", type="primary")
        if submitted and cid:
            st.session_state.config.customer_id = int(cid)
            st.session_state.config.save()
            st.rerun()

    st.info(
        "If you need help or don't have a license, contact "
        "david@portfoliotracker.com"
    )
    st.stop()


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if not _check_license():
        return

    config: AppConfig = st.session_state.config

    # Sidebar: data status
    with st.sidebar:
        st.title("Portfolio Tracker")
        st.caption("v2.0.0")
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
    | **Settings** | License, export/import settings, Excel export |
    """)

    if not config.folders:
        st.info(
            "**Getting started:** Go to the **Import** page to add your "
            "MultiWalk base folders."
        )


if __name__ == "__main__":
    main()
