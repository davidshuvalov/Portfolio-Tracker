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
    page_title="Portfolio Tracker — A Tool for MultiWalk",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

from core.config import AppConfig
from ui.styles import inject_styles, render_logo, render_sidebar_logo

inject_styles()

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
        valid, message = validate_full(customer_id, config.multiwalk_folder)
    except Exception as e:
        valid, message = False, str(e)

    if not valid:
        _show_license_entry(f"License check failed: {message}")
        return False

    return True


def _show_license_entry(prompt: str) -> None:
    render_logo()
    st.divider()
    st.warning(prompt)

    config: AppConfig = st.session_state.config

    with st.form("license_form"):
        st.markdown("**TradeStation Customer ID**")
        st.caption("The number you use to log in to TradeStation — verified against the MultiWalk license DLL.")
        cid = st.number_input(
            "Customer ID",
            min_value=1,
            max_value=9_999_999,
            value=int(config.customer_id) if config.customer_id else 1,
            step=1,
            label_visibility="collapsed",
        )

        st.markdown("**MultiWalk Program Folder**")
        st.caption("Folder containing `MultiWalkLicense64.dll`. Leave blank to auto-detect from registry.")
        folder = st.text_input(
            "MultiWalk Program Folder",
            value=config.multiwalk_folder or "",
            placeholder=r"e.g. C:\Users\you\Documents\MultiWalk\Program",
            label_visibility="collapsed",
        )

        submitted = st.form_submit_button("Activate", type="primary")
        if submitted and cid:
            st.session_state.config.customer_id = int(cid)
            st.session_state.config.multiwalk_folder = folder.strip()
            st.session_state.config.save()
            st.rerun()

    st.info("If you need help or don't have a license, contact david@portfoliotracker.com")
    st.stop()


# ── Step card renderer ────────────────────────────────────────────────────────
def _step_card(
    col,
    num: int,
    title: str,
    desc: str,
    done: bool,
    active: bool,
    page: str,
    action: str,
) -> None:
    with col:
        with st.container(border=True):
            if done:
                badge = (
                    '<span style="background:#071f12;color:#10b981;border-radius:4px;'
                    'padding:2px 9px;font-size:0.68rem;font-weight:700;letter-spacing:0.1em">'
                    "✓ COMPLETE</span>"
                )
                num_color = "#0c2a1a"
            elif active:
                badge = (
                    '<span style="background:#071428;color:#3b82f6;border-radius:4px;'
                    'padding:2px 9px;font-size:0.68rem;font-weight:700;letter-spacing:0.1em">'
                    "● CURRENT</span>"
                )
                num_color = "#0d1f3d"
            else:
                badge = (
                    '<span style="color:#1e2d47;font-size:0.68rem;font-weight:700;'
                    'letter-spacing:0.1em">PENDING</span>'
                )
                num_color = "#0d1626"

            st.markdown(
                f'<div style="display:flex;justify-content:space-between;'
                f'align-items:flex-start;margin-bottom:0.6rem">'
                f'<span style="font-size:2.8rem;font-weight:800;color:{num_color};'
                f'line-height:1;letter-spacing:-0.04em">{num:02d}</span>'
                f'{badge}</div>',
                unsafe_allow_html=True,
            )
            st.markdown(f"**{title}**")
            st.caption(desc)
            if done or active:
                st.markdown(
                    f'<span style="color:#3b82f6;font-size:0.85rem">{action}</span>',
                    unsafe_allow_html=True,
                )


# ── Analytics card renderer ───────────────────────────────────────────────────
def _analytics_card(col, title: str, desc: str, page: str) -> None:
    with col:
        with st.container(border=True):
            st.markdown(f"**{title}**")
            st.caption(desc)
            st.page_link(page, label="Open →")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if not _check_license():
        return

    from ui.workflow import step_status, render_workflow_sidebar

    # ── Sidebar ───────────────────────────────────────────────────────────────
    with st.sidebar:
        render_sidebar_logo()
        st.divider()
        render_workflow_sidebar()

        if st.session_state.imported_data is not None:
            data = st.session_state.imported_data
            n_strats = len(data.strategy_names)
            start, end = data.date_range
            st.metric("Strategies loaded", n_strats)
            st.caption(f"{start} → {end}")

    # ── Hero ──────────────────────────────────────────────────────────────────
    render_logo()
    st.markdown(
        '<p style="color:#64748b;font-size:0.95rem;margin-top:0.25rem;margin-bottom:0">'
        "Systematic portfolio analytics for futures traders using MultiWalk Pro."
        "</p>",
        unsafe_allow_html=True,
    )
    st.divider()

    # ── Workflow steps ────────────────────────────────────────────────────────
    st.markdown(
        '<p style="color:#94a3b8;font-size:0.8rem;text-transform:uppercase;'
        'letter-spacing:0.12em;margin-bottom:0.75rem">Setup Workflow</p>',
        unsafe_allow_html=True,
    )

    status = step_status()
    _keys = ["folders", "data", "strategies", "portfolio"]
    active_step = next((i for i, k in enumerate(_keys) if not status[k]), len(_keys))

    steps = [
        (
            "folders", 1, "Add Folders",
            "Point the app at your MultiWalk strategy folders on disk.",
            "Open Import →", "ui/pages/01_Import.py",
        ),
        (
            "data", 2, "Import Data",
            "Scan your folders and load all strategy equity, trade, and walkforward CSVs.",
            "Go to Import →", "ui/pages/01_Import.py",
        ),
        (
            "strategies", 3, "Review Strategies",
            "Set each strategy's status (Live / Paper / Retired), contracts, symbol, and sector.",
            "Open Strategies →", "ui/pages/02_Strategies.py",
        ),
        (
            "portfolio", 4, "Build Portfolio",
            "Aggregate all Live strategies and explore the combined portfolio metrics.",
            "Open Portfolio →", "ui/pages/03_Portfolio.py",
        ),
    ]

    cols = st.columns(4, gap="medium")
    for i, (col, (key, num, title, desc, action, page)) in enumerate(zip(cols, steps)):
        _step_card(
            col, num, title, desc,
            done=status[key],
            active=(i == active_step),
            page=page,
            action=action,
        )

    st.divider()

    # ── Analytics ─────────────────────────────────────────────────────────────
    if active_step == len(_keys):
        st.markdown(
            '<p style="color:#94a3b8;font-size:0.8rem;text-transform:uppercase;'
            'letter-spacing:0.12em;margin-bottom:0.75rem">Analytics</p>',
            unsafe_allow_html=True,
        )

        analytics = [
            ("Monte Carlo",          "ui/pages/_04_Monte_Carlo.py",          "Risk-of-ruin simulation via trade resampling"),
            ("Correlations",         "ui/pages/_05_Correlations.py",         "Pairwise strategy correlation — IS, OOS & IS+OOS"),
            ("Diversification",      "ui/pages/_06_Diversification.py",      "Portfolio composition by sector, symbol and type"),
            ("Leave One Out",        "ui/pages/_07_Leave_One_Out.py",        "Impact on portfolio metrics of removing each strategy"),
            ("Backtest",             "ui/pages/_08_Backtest.py",             "Recreate the performance of your actual traded portfolio"),
            ("Eligibility Backtest", "ui/pages/_09_Eligibility_Backtest.py", "Walk-forward rule validation across OOS windows"),
            ("Margin Tracking",      "ui/pages/_10_Margin_Tracking.py",      "Historical daily margin utilisation by symbol and sector"),
            ("Position Check",       "ui/pages/_11_Position_Check.py",       "Compare current MultiWalk positions to your live account"),
        ]

        a_cols = st.columns(4, gap="medium")
        for j, (label, page, desc) in enumerate(analytics):
            _analytics_card(a_cols[j % 4], label, desc, page)

        st.divider()
    else:
        st.markdown(
            f'<p style="color:#64748b;font-size:0.9rem">'
            f"Complete step {active_step + 1} above to continue.</p>",
            unsafe_allow_html=True,
        )

    # ── How to Use ────────────────────────────────────────────────────────────
    with st.expander("How to Use — Portfolio Tracker & MultiWalk Workflow"):
        st.markdown("""
### What is Portfolio Tracker?

Portfolio Tracker is built for systematic futures traders using **MultiWalk Pro**
(MultiCharts). MultiWalk runs walk-forward optimisation across all your strategies
simultaneously — Portfolio Tracker then aggregates, analyses, and visualises that
data so you can manage a diversified algorithmic portfolio in minutes, not hours.

---

### Before You Start

**Run MultiWalk Trader Pro** to rerun and reoptimise all your strategy folders.
This refreshes the underlying CSV files (`EquityData.csv`, `TradeData.csv`,
`Walkforward Details.csv`) that Portfolio Tracker reads.

Each strategy should be its own MultiWalk workspace (folder). Organise folders
by category — for example:

| Folder | Contents |
|---|---|
| `Live/` | Strategies currently being traded |
| `Incubation/` | Strategies in out-of-sample testing |
| `Past/` | Retired strategies |
| `BuyHold/` | Simple ATR-tracking workspace for position sizing |

---

### Step-by-Step Workflow

**1 — Add Folders**
Go to the **Import** page and paste the root folder paths where your MultiWalk
strategy subfolders live. You can add multiple top-level folders. A Buy & Hold
folder provides ATR reference data used for position sizing.

**2 — Import Data**
Click **Scan** to discover all strategy subfolders automatically, then click
**Import**. Hundreds of strategies load in seconds. The app reads equity curves,
trade-level data, and walk-forward in-sample / out-of-sample date ranges.

**3 — Review Strategies**
Every discovered strategy appears in an editable table. For each strategy, set:
- **Status** — Live, Paper, Incubating, Retired, Past, etc. Only *Live* strategies
  appear in the portfolio.
- **Contracts** — position size multiplier (use decimals for micro-contract fractions).
- **Symbol / Sector / Timeframe / Type / Horizon** — metadata used by analytics pages.

Use the bulk actions to set all new strategies at once or reset contracts.

**4 — Build Portfolio**
Click **Build Portfolio** to aggregate all Live strategies. Review the combined
equity curve, monthly P&L heatmap, and per-strategy KPI table. Adjust contracts
and rebuild to see how sizing changes affect portfolio-level metrics.

---

### Analytics Modules

Once all four steps are complete, eight analytics modules unlock:

| Module | Purpose |
|---|---|
| **Monte Carlo** | Thousands of resampled trade sequences to estimate risk of ruin, drawdown percentiles, and profit distributions |
| **Correlations** | Pairwise strategy correlations across IS, OOS, and IS+OOS periods — spot dangerous clusters |
| **Diversification** | Portfolio composition by sector, symbol, type, and horizon — measure diversification benefit |
| **Leave One Out** | Run the portfolio with each strategy removed to find which add or detract value |
| **Backtest** | Recreate the exact portfolio you have actually been trading for any date range |
| **Eligibility Backtest** | Apply walk-forward criteria (min Sharpe, max drawdown) to see which strategies would have qualified in each OOS window |
| **Margin Tracking** | Estimate daily margin requirements from historical positions and current broker margins |
| **Position Check** | Compare today's MultiWalk open positions to your live broker account — highlights discrepancies |

---

### Tips

- Run MultiWalk Trader Pro **weekly** (or after any reoptimisation) then
  re-import in Portfolio Tracker to keep your data fresh.
- The **Backtest** module is ideal for tracking how your live trading has
  performed versus the MultiWalk hypothetical — import only the strategies
  you were actually trading for each sub-period.
- Use **Eligibility Backtest** to apply objective rules (e.g. OOS Sharpe > 0.5)
  that filter strategies before they enter the live portfolio.
- The **Buy & Hold** folder provides current ATR values across all markets,
  useful for volatility-adjusted position sizing.
""")


if __name__ == "__main__":
    main()
