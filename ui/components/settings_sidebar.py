"""
Compact settings sidebar panels.

Each `render_*_sidebar` function renders a collapsible settings block in the
current Streamlit context (call inside `with st.sidebar:` or at page level).
Returns True when the user saved changes so callers can trigger a rerun.

Usage:
    with st.sidebar:
        saved = render_portfolio_sidebar(config)
    if saved:
        st.rerun()
"""

from __future__ import annotations

import streamlit as st

from core.config import AppConfig


# ── Portfolio settings ────────────────────────────────────────────────────────

def render_portfolio_sidebar(config: AppConfig) -> bool:
    """
    Compact portfolio / data-scope settings.
    Shows: lookback period, cutoff toggle, active statuses.
    """
    with st.expander("Portfolio Settings", expanded=False):
        with st.form("sidebar_portfolio_form"):
            period_years = st.number_input(
                "Lookback (years)",
                min_value=0.5, max_value=20.0, step=0.5,
                value=float(config.portfolio.period_years), format="%.1f",
                help="How many years of history to use for summary metrics.",
            )
            use_cutoff = st.checkbox(
                "Use cutoff date",
                value=config.portfolio.use_cutoff,
            )
            cutoff_str = st.text_input(
                "Cutoff (YYYY-MM-DD)",
                value=config.portfolio.cutoff_date or "",
                disabled=not use_cutoff,
                help="Treat data after this date as the OOS end for all strategies.",
            )
            status_raw = st.text_input(
                "Active statuses",
                value=", ".join(config.eligibility.status_include),
                help="Comma-separated statuses eligible for the portfolio.",
            )

            if st.form_submit_button("Save", type="primary", use_container_width=True):
                config.portfolio.period_years = float(period_years)
                config.portfolio.use_cutoff   = use_cutoff
                if use_cutoff and cutoff_str.strip():
                    try:
                        from datetime import date as _date
                        _date.fromisoformat(cutoff_str.strip())
                        config.portfolio.cutoff_date = cutoff_str.strip()
                    except ValueError:
                        st.error("Invalid date — use YYYY-MM-DD.")
                        return False
                elif not use_cutoff:
                    config.portfolio.cutoff_date = None
                config.eligibility.status_include = (
                    [s.strip() for s in status_raw.split(",") if s.strip()] or ["Live"]
                )
                config.save()
                st.session_state.config = config
                return True
    return False


# ── Monte Carlo settings ──────────────────────────────────────────────────────

def render_mc_sidebar(config: AppConfig) -> bool:
    """
    Compact Monte Carlo extended settings (the page already shows the main ones).
    Shows: starting equity, output samples, remove-best %, solve-for-ROR.
    """
    with st.expander("MC Advanced Settings", expanded=False):
        with st.form("sidebar_mc_form"):
            cs = config.contract_sizing
            mc = config.monte_carlo

            starting_equity = st.number_input(
                "Starting equity ($)",
                min_value=10_000.0, max_value=100_000_000.0, step=5_000.0,
                value=float(cs.starting_equity), format="%.0f",
                help="Used when 'Solve for ROR' is off.",
            )
            solve_for_ror = st.checkbox(
                "Solve for ROR target",
                value=mc.solve_for_ror,
                help="Iterate starting equity until portfolio ROR matches the target.",
            )
            output_samples = st.number_input(
                "Output samples",
                min_value=1, max_value=500, step=5,
                value=int(mc.output_samples),
                help="Number of scenario paths to emit in the results table.",
            )
            remove_best = st.slider(
                "Remove best % days",
                0.0, 0.10, float(mc.remove_best_pct), 0.005,
                format="%.1%%",
                help="Trim top-N% days before sampling to stress-test the distribution.",
            )

            if st.form_submit_button("Save", type="primary", use_container_width=True):
                config.contract_sizing.starting_equity = float(starting_equity)
                config.monte_carlo.solve_for_ror       = solve_for_ror
                config.monte_carlo.output_samples      = int(output_samples)
                config.monte_carlo.remove_best_pct     = float(remove_best)
                config.save()
                st.session_state.config = config
                return True
    return False


# ── Contract sizing settings ──────────────────────────────────────────────────

def render_contract_sizing_sidebar(config: AppConfig) -> bool:
    """
    Compact contract sizing settings.
    Shows: ATR window, blend ratio, reweight scope + gain.
    """
    with st.expander("Contract Sizing", expanded=False):
        with st.form("sidebar_contract_form"):
            cs = config.contract_sizing

            atr_window = st.selectbox(
                "ATR window",
                ["ATR Last 3 Months", "ATR Last 6 Months", "ATR Last 12 Months"],
                index=["ATR Last 3 Months", "ATR Last 6 Months", "ATR Last 12 Months"]
                      .index(cs.atr_window),
                help="Rolling window for dollar ATR from trade MFE+MAE.",
            )
            blend = st.slider(
                "ATR / Margin blend",
                0.0, 1.0, float(cs.contract_ratio_margin_atr), 0.05,
                help="0 = pure margin, 1 = pure ATR.",
            )
            rw_scope = st.selectbox(
                "Reweight scope",
                ["None", "All", "Index Only"],
                index=["None", "All", "Index Only"].index(cs.reweight_scope),
                help="Which strategies get ATR-based reweighting.",
            )
            rw_gain = st.slider(
                "Reweight gain",
                0.5, 3.0, float(cs.reweight_gain), 0.05,
                disabled=(rw_scope == "None"),
                help="Multiplier on top of ATR reweighting.",
            )

            if st.form_submit_button("Save", type="primary", use_container_width=True):
                config.contract_sizing.atr_window                = atr_window
                config.contract_sizing.contract_ratio_margin_atr = float(blend)
                config.contract_sizing.reweight_scope            = rw_scope
                config.contract_sizing.reweight_gain             = float(rw_gain)
                config.save()
                st.session_state.config = config
                return True
    return False


# ── Strategy ranking settings ─────────────────────────────────────────────────

_RANK_LABELS = {
    "rtd_oos":                "RTD (OOS)",
    "rtd_12_months":          "RTD (12M)",
    "sharpe_isoos":           "Sharpe IS+OOS",
    "profit_since_oos_start": "OOS Total ($)",
    "profit_last_12_months":  "Last 12M ($)",
    "k_factor":               "K-Factor",
    "ulcer_index":            "Ulcer Index",
    "contracts":              "Contracts",
}


def render_ranking_sidebar(config: AppConfig) -> bool:
    """
    Compact ranking settings.
    Shows: metric, ascending, eligible-only, group-by-sector.
    """
    with st.expander("Ranking Options", expanded=False):
        with st.form("sidebar_ranking_form"):
            rk = config.ranking

            rank_metric = st.selectbox(
                "Rank by",
                list(_RANK_LABELS.keys()),
                index=list(_RANK_LABELS.keys()).index(rk.metric),
                format_func=lambda k: _RANK_LABELS[k],
            )
            col1, col2 = st.columns(2)
            with col1:
                ascending    = st.checkbox("Ascending",        value=rk.ascending)
                elig_only    = st.checkbox("Eligible only",    value=rk.eligible_only)
            with col2:
                group_sector    = st.checkbox("Group by sector",   value=rk.group_by_sector)
                group_contracts = st.checkbox("Sub-sort by contracts", value=rk.group_by_contracts)

            if st.form_submit_button("Save", type="primary", use_container_width=True):
                config.ranking.metric             = rank_metric
                config.ranking.ascending          = ascending
                config.ranking.eligible_only      = elig_only
                config.ranking.group_by_sector    = group_sector
                config.ranking.group_by_contracts = group_contracts
                config.save()
                st.session_state.config = config
                return True
    return False
