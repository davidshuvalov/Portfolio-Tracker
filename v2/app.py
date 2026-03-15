"""
Portfolio Tracker v2 — Streamlit entrypoint.

Handles:
- Session state initialisation
- License gate (checks MultiWalk DLL; shows customer-ID entry if not configured)
- Home page with order-of-operations workflow guide
- Sidebar with workflow progress status
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
            value=1,
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

    from ui.workflow import step_status, render_workflow_sidebar

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        st.title("Portfolio Tracker")
        st.caption("v2.0.0")
        st.divider()
        render_workflow_sidebar()

        if st.session_state.imported_data is not None:
            data = st.session_state.imported_data
            n_strats = len(data.strategy_names)
            start, end = data.date_range
            st.metric("Strategies loaded", n_strats)
            st.caption(f"{start} → {end}")

    # ── Home page ─────────────────────────────────────────────────────────────
    st.title("Portfolio Tracker v2")
    st.markdown("Follow these four steps to get started. Complete them in order.")

    status = step_status()

    # Determine the current active step (first incomplete)
    _keys = ["folders", "data", "strategies", "portfolio"]
    active_step = next((i for i, k in enumerate(_keys) if not status[k]), len(_keys))

    steps = [
        {
            "key": "folders",
            "num": 1,
            "title": "Add Folders",
            "desc": "Tell the app where your MultiWalk strategy folders live on disk.",
            "action": "Open Import",
            "page": "ui/pages/01_Import.py",
        },
        {
            "key": "data",
            "num": 2,
            "title": "Import Data",
            "desc": "Scan your folders and load all strategy EquityData CSVs into memory.",
            "action": "Go to Import",
            "page": "ui/pages/01_Import.py",
        },
        {
            "key": "strategies",
            "num": 3,
            "title": "Review Strategies",
            "desc": "Set each strategy's status (Live/Paper/Retired), contracts, symbol, and sector.",
            "action": "Open Strategies",
            "page": "ui/pages/02_Strategies.py",
        },
        {
            "key": "portfolio",
            "num": 4,
            "title": "Build Portfolio",
            "desc": "Aggregate all Live strategies and view combined portfolio metrics.",
            "action": "Open Portfolio",
            "page": "ui/pages/03_Portfolio.py",
        },
    ]

    cols = st.columns(4)
    for i, (col, step) in enumerate(zip(cols, steps)):
        done = status[step["key"]]
        is_active = i == active_step
        with col:
            if done:
                st.success(f"**Step {step['num']} — {step['title']}** ✅")
            elif is_active:
                st.info(f"**Step {step['num']} — {step['title']}**")
            else:
                st.markdown(
                    f"<div style='padding:1em;border:1px solid #444;border-radius:6px;opacity:0.5'>"
                    f"<b>Step {step['num']} — {step['title']}</b></div>",
                    unsafe_allow_html=True,
                )
                st.write("")  # spacing
                continue

            st.caption(step["desc"])
            st.page_link(step["page"], label=step["action"])

    st.divider()

    # Analytics section — only show once all 4 steps are done
    if active_step == len(_keys):
        st.markdown("### Analytics")
        st.markdown("All setup steps complete. Explore your portfolio:")

        analytics = [
            ("Monte Carlo", "ui/pages/04_Monte_Carlo.py", "Risk-of-ruin simulation"),
            ("Correlations", "ui/pages/05_Correlations.py", "Strategy correlation analysis"),
            ("Diversification", "ui/pages/06_Diversification.py", "Sector & symbol breakdown"),
            ("Leave One Out", "ui/pages/07_Leave_One_Out.py", "Portfolio sensitivity"),
            ("Backtest", "ui/pages/08_Backtest.py", "Historical period performance"),
            ("Eligibility Backtest", "ui/pages/09_Eligibility_Backtest.py", "Walk-forward rules"),
            ("Margin Tracking", "ui/pages/10_Margin_Tracking.py", "Daily margin by symbol"),
            ("Position Check", "ui/pages/11_Position_Check.py", "Current live positions"),
        ]

        a_cols = st.columns(4)
        for j, (label, page, desc) in enumerate(analytics):
            with a_cols[j % 4]:
                st.page_link(page, label=f"**{label}**")
                st.caption(desc)
    else:
        st.markdown(
            f"_Complete step {active_step + 1} above to continue._"
        )


if __name__ == "__main__":
    main()
