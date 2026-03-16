"""
Inputs page — central configuration hub, mirrors the VBA 'Inputs' tab.

Sections (matching spreadsheet layout):
  1. Data Folders & Date Settings
  2. Portfolio Settings
  3. Incubation Settings         ← Incubation_Period / Min_Incubation_Profit
  4. Quit Point Settings         ← Quitting_Method / Quit_Dollar / Quit_percent / SD_Multiple
  5. Strategy Eligibility Settings ← all Yes/No toggles from the spreadsheet
  6. Monte Carlo Settings
  7. Correlation Thresholds
  8. Margin Settings
"""

from __future__ import annotations

import datetime
from pathlib import Path

import streamlit as st

from core.config import AppConfig, StrategyRankingConfig

st.set_page_config(page_title="Inputs", layout="wide")

st.title("Inputs")
st.caption(
    "All configurable parameters that drive the analysis — mirrors the VBA Inputs tab. "
    "Changes are saved and take effect on the next portfolio rebuild."
)

config: AppConfig = st.session_state.get("config", AppConfig.load())
_any_saved = False

tab_data, tab_elig, tab_analytics, tab_portfolio = st.tabs([
    "📁 Data & Dates",
    "🎯 Eligibility",
    "📊 Analytics",
    "💼 Portfolio",
])


def _save(new_config: AppConfig) -> None:
    global _any_saved
    new_config.save()
    st.session_state.config = new_config
    st.session_state.pop("portfolio_data", None)   # force rebuild
    _any_saved = True


tab_data.__enter__()

# ══════════════════════════════════════════════════════════════════════════════
# 1. DATA FOLDERS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("Data Folders")
st.caption(
    "MultiWalk base folders — each should contain one sub-folder per strategy "
    "(with EquityData, TradeData and Walkforward Details CSVs)."
)

if config.folders:
    for i, folder in enumerate(config.folders):
        c_path, c_ok, c_rm = st.columns([5, 1, 1])
        c_path.code(str(folder))
        c_ok.markdown("✅" if folder.exists() else "❌")
        if c_rm.button("Remove", key=f"rm_folder_{i}"):
            nc = config.model_copy(deep=True)
            nc.folders = [f for f in nc.folders if f != folder]
            _save(nc)
            st.rerun()
else:
    st.info("No folders added yet.")

with st.form("add_folder_form", clear_on_submit=True):
    new_folder = st.text_input("Add folder path", placeholder="/path/to/MultiWalk/Strategies")
    if st.form_submit_button("Add Folder"):
        p = Path(new_folder.strip())
        if str(p) and p not in config.folders:
            nc = config.model_copy(deep=True)
            nc.folders.append(p)
            _save(nc)
            st.rerun()
        elif p in config.folders:
            st.warning("Folder already added.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# 2. DATE & PORTFOLIO SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("Date & Portfolio Settings")

with st.form("date_portfolio_form"):
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        date_format = st.selectbox(
            "CSV Date Format",
            ["DMY", "MDY"],
            index=0 if config.date_format == "DMY" else 1,
            help="DMY = DD/MM/YYYY (EU/UK/AU). MDY = MM/DD/YYYY (US).",
        )
    with col2:
        period_years = st.number_input(
            "Lookback period (years)",
            min_value=0.5, max_value=20.0, step=0.5,
            value=float(config.portfolio.period_years),
        )
    with col3:
        live_status = st.text_input("Live status name", value=config.portfolio.live_status)
    with col4:
        pass_status = st.text_input("Pass status name", value=config.portfolio.pass_status)
    with col5:
        buy_hold_status = st.text_input("Buy & Hold status", value=config.portfolio.buy_and_hold_status)

    col6, col7 = st.columns([1, 2])
    with col6:
        use_cutoff = st.checkbox("Apply cutoff date", value=config.portfolio.use_cutoff)
    with col7:
        default_cutoff = None
        if config.portfolio.cutoff_date:
            try:
                default_cutoff = datetime.date.fromisoformat(config.portfolio.cutoff_date)
            except ValueError:
                pass
        cutoff_date = st.date_input("Cutoff date", value=default_cutoff, disabled=not use_cutoff)

    if st.form_submit_button("Save Date & Portfolio Settings", type="primary"):
        nc = config.model_copy(deep=True)
        nc.date_format = date_format
        nc.portfolio.period_years = period_years
        nc.portfolio.live_status = live_status.strip() or "Live"
        nc.portfolio.pass_status = pass_status.strip() or "Pass"
        nc.portfolio.buy_and_hold_status = buy_hold_status.strip() or "Buy&Hold"
        nc.portfolio.use_cutoff = use_cutoff
        nc.portfolio.cutoff_date = cutoff_date.isoformat() if (use_cutoff and cutoff_date) else None
        _save(nc)
        st.success("Saved.")

tab_data.__exit__(None, None, None)
tab_elig.__enter__()

# ══════════════════════════════════════════════════════════════════════════════
# 3. INCUBATION SETTINGS   (mirrors VBA rows 3–7 of Inputs tab)
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("Incubation Settings")
st.caption(
    "A strategy passes incubation when its cumulative OOS profit ≥ "
    "expected_daily × elapsed_days × min_ratio, **after** the minimum period has elapsed.  \n"
    "**Passed** — target hit.  **Not Passed** — enough history, target never reached.  "
    "**Incubating** — not enough OOS history yet."
)

with st.form("incubation_form"):
    col_i1, col_i2, _ = st.columns([1, 1, 2])

    with col_i1:
        inc_months = st.number_input(
            "Incubation Time in Months",
            min_value=1, max_value=60, step=1,
            value=int(config.incubation.months),
            help="Minimum OOS months before incubation is evaluated. (VBA: Incubation_Period)",
        )
    with col_i2:
        inc_ratio_pct = st.number_input(
            "Minimum Incubation Profit (%)",
            min_value=1, max_value=500, step=5,
            value=int(round(config.incubation.min_profit_ratio * 100)),
            help=(
                "Cumulative OOS profit must reach this % of the expected daily rate × elapsed days. "
                "25% means 25%% of expected pace. (VBA: Min_Incubation_Profit)"
            ),
        )

    if st.form_submit_button("Save Incubation Settings", type="primary"):
        nc = config.model_copy(deep=True)
        nc.incubation.months = int(inc_months)
        nc.incubation.min_profit_ratio = float(inc_ratio_pct) / 100.0
        _save(nc)
        st.success("Saved.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# 4. QUIT POINT SETTINGS   (mirrors VBA rows 9–13 of Inputs tab)
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("Quit Point Settings")
st.caption(
    "Defines when a strategy should be removed from live trading based on OOS performance.  \n"
    "**Drawdown**: quit when equity drops below peak − MIN(Max Dollars, Max % × IS_MaxDD).  \n"
    "**Standard Deviation**: quit when equity falls below the statistical lower bound."
)

with st.form("quitting_form"):
    col_q1, col_q2, col_q3, col_q4 = st.columns(4)

    with col_q1:
        quit_method = st.selectbox(
            "Quitting Point Method",
            ["Drawdown", "Standard Deviation", "None"],
            index=["Drawdown", "Standard Deviation", "None"].index(config.quitting.method),
            help="Quitting_Method named range.",
        )
    with col_q2:
        quit_max_dollars = st.number_input(
            "Max Dollars ($)",
            min_value=0.0, max_value=10_000_000.0, step=1_000.0, format="%.0f",
            value=float(config.quitting.max_dollars),
            help="Quit_Dollar — hard cap on the quitting drawdown amount.",
        )
    with col_q3:
        quit_pct = st.number_input(
            "Max Percent of Max Drawdown (%)",
            min_value=0.0, max_value=1000.0, step=10.0, format="%.0f",
            value=float(config.quitting.max_percent_drawdown * 100),
            help="Quit_percent — as % of IS max drawdown (e.g. 150 = 150%).",
        )
    with col_q4:
        quit_sd = st.number_input(
            "Multiple of Standard Deviation",
            min_value=0.0, max_value=10.0, step=0.01, format="%.3f",
            value=float(config.quitting.sd_multiple),
            help="Quitting_SD_Multiple — used when method = Standard Deviation.",
        )

    if st.form_submit_button("Save Quit Point Settings", type="primary"):
        nc = config.model_copy(deep=True)
        nc.quitting.method = quit_method
        nc.quitting.max_dollars = float(quit_max_dollars)
        nc.quitting.max_percent_drawdown = float(quit_pct) / 100.0
        nc.quitting.sd_multiple = float(quit_sd)
        _save(nc)
        st.success("Saved.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# 5. STRATEGY ELIGIBILITY SETTINGS   (mirrors VBA rows 15–28 of Inputs tab)
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("Strategy Eligibility Settings")
st.caption(
    "Controls which strategies are eligible at each rebalance. "
    "Enable a check to require that condition for inclusion."
)

with st.form("eligibility_form"):
    e = config.eligibility   # shorthand

    # ── Row 1 header labels ───────────────────────────────────────────────────
    st.markdown("**Profit Checks (> $0)** &nbsp;&nbsp;&nbsp; **Efficiency Checks (> Ratio)** &nbsp;&nbsp;&nbsp; **Status / Threshold**")
    st.markdown("---")

    # ── Grid: Profit | Efficiency | Status columns ────────────────────────────
    # Each row: left = profit toggle, mid = efficiency toggle, right = other
    col_p, col_e, col_s, col_d = st.columns([2, 2, 2, 2])

    with col_p:
        st.markdown("**Profit Checks**")
        p1m    = st.checkbox("Profit Last 1 Month > $0",             value=e.profit_1m,    key="p_1m")
        p3m    = st.checkbox("Profit Last 3 Months > $0",            value=e.profit_3m,    key="p_3m")
        p6m    = st.checkbox("Profit Last 6 Months > $0",            value=e.profit_6m,    key="p_6m")
        p3or6m = st.checkbox("Profit Last 3 OR 6 Months > $0",       value=e.profit_3or6m, key="p_3or6m",
                              help="Pass if either 3M or 6M profit is > $0.")
        p9m    = st.checkbox("Profit Last 9 Months > $0",            value=e.profit_9m,    key="p_9m")
        p12m   = st.checkbox("Profit Last 12 Months > $0",           value=e.profit_12m,   key="p_12m")
        poos   = st.checkbox("Profit Since OOS Start > $0",          value=e.profit_oos,   key="p_oos")

        st.markdown("---")
        st.markdown("**Loss Disqualifiers (< $0)**")
        l1m = st.checkbox("Profit Last 1 Month < $0 → exclude",  value=e.loss_1m, key="l_1m")
        l3m = st.checkbox("Profit Last 3 Months < $0 → exclude", value=e.loss_3m, key="l_3m")
        l6m = st.checkbox("Profit Last 6 Months < $0 → exclude", value=e.loss_6m, key="l_6m")

    with col_e:
        st.markdown("**Efficiency Checks**")
        eff1m  = st.checkbox("Efficiency Last 1 Month > Ratio",      value=e.efficiency_1m,  key="eff_1m")
        eff3m  = st.checkbox("Efficiency Last 3 Months > Ratio",     value=e.efficiency_3m,  key="eff_3m")
        eff6m  = st.checkbox("Efficiency Last 6 Months > Ratio",     value=e.efficiency_6m,  key="eff_6m")
        eff9m  = st.checkbox("Efficiency Last 9 Months > Ratio",     value=e.efficiency_9m,  key="eff_9m")
        eff12m = st.checkbox("Efficiency Last 12 Months > Ratio",    value=e.efficiency_12m, key="eff_12m")
        effoos = st.checkbox("Efficiency Since OOS Start > Ratio",   value=e.efficiency_oos, key="eff_oos")

        st.markdown("---")
        st.markdown("**Efficiency Disqualifiers (< Ratio)**")
        el1m = st.checkbox("Efficiency Last 1 Month < Ratio → exclude",  value=e.efficiency_loss_1m, key="el_1m")
        el3m = st.checkbox("Efficiency Last 3 Months < Ratio → exclude", value=e.efficiency_loss_3m, key="el_3m")
        el6m = st.checkbox("Efficiency Last 6 Months < Ratio → exclude", value=e.efficiency_loss_6m, key="el_6m")

    with col_s:
        st.markdown("**Status Gates**")
        use_inc  = st.checkbox("Incubation Status must be Passed",
                               value=e.use_incubation,  key="use_inc",
                               help="EligibilityIncubation — strategy must have passed incubation.")
        use_quit = st.checkbox("Quitting Status: exclude Quit strategies",
                               value=e.use_quitting,   key="use_quit",
                               help="EligibilityQuitting — strategies currently 'Quit' are excluded.")

        st.markdown("---")
        st.markdown("**Count Profitable Months**")
        use_cpm = st.checkbox("Enable monthly profit count check",
                              value=e.use_count_monthly_profits, key="use_cpm",
                              help="EligibilityCountMonthlyProfits")
        min_pos_months = st.number_input(
            "Min Months > 0",
            min_value=1, max_value=36, step=1,
            value=int(e.min_positive_months),
            disabled=not use_cpm,
            help="EligibilityMinimumMonths",
        )
        monthly_op = st.selectbox(
            '">0" or "≥0"',
            [">0", ">=0"],
            index=0 if e.monthly_profit_operator == ">0" else 1,
            disabled=not use_cpm,
            help="EligibilityGreaterThan",
        )

        st.markdown("---")
        st.markdown("**Additional User Filter**")
        add_filter = st.checkbox("Enable additional filter",
                                 value=e.additional_user_filter, key="add_filter",
                                 help="AdditionalUserFilter")
        add_col = st.text_input(
            "Filter column",
            value=e.additional_user_filter_column,
            disabled=not add_filter,
        )
        add_min = st.number_input(
            "Min value",
            value=float(e.additional_user_filter_min_value),
            format="%.2f",
            disabled=not add_filter,
        )

    with col_d:
        st.markdown("**Thresholds**")
        eff_ratio_pct = st.number_input(
            "Efficiency Ratio (%)",
            min_value=0, max_value=500, step=5,
            value=int(round(e.efficiency_ratio * 100)),
            help="EfficiencyRatio — used by all efficiency checks.",
        )
        days_thresh = st.number_input(
            "Days Threshold",
            min_value=0, max_value=31, step=1,
            value=int(e.days_threshold_oos),
            help=(
                "EligibilityDaysThreshold — if > 0, profit windows snap to "
                "month-end boundaries unless current month has ≥ this many days. "
                "0 = rolling windows."
            ),
        )
        elig_months_total = st.number_input(
            "Total Months (profit count window)",
            min_value=1, max_value=36, step=1,
            value=int(e.eligibility_months),
            help="EligibilityTotalMonths — lookback for counting profitable months.",
        )
        oos_dd_cap = st.number_input(
            "Max OOS DD / IS DD ratio",
            min_value=0.0, max_value=10.0, step=0.1, format="%.1f",
            value=float(e.oos_dd_vs_is_cap),
            help="Max ratio of OOS to IS max drawdown. 0 = disabled.",
        )
        date_type = st.selectbox(
            "Eligibility Date Type",
            ["OOS Start Date", "Incubation Pass Date"],
            index=0 if e.date_type == "OOS Start Date" else 1,
        )
        max_horizon = st.number_input(
            "Max Horizon (months)",
            min_value=1, max_value=60, step=1,
            value=int(e.max_horizon),
        )
        status_raw = st.text_input(
            "Status values to include",
            value=", ".join(e.status_include),
            help="Comma-separated status names that are candidates for eligibility.",
        )

    if st.form_submit_button("Save Eligibility Settings", type="primary"):
        nc = config.model_copy(deep=True)
        eg = nc.eligibility

        # Profit checks
        eg.profit_1m    = p1m
        eg.profit_3m    = p3m
        eg.profit_6m    = p6m
        eg.profit_3or6m = p3or6m
        eg.profit_9m    = p9m
        eg.profit_12m   = p12m
        eg.profit_oos   = poos

        # Efficiency checks
        eg.efficiency_1m  = eff1m
        eg.efficiency_3m  = eff3m
        eg.efficiency_6m  = eff6m
        eg.efficiency_9m  = eff9m
        eg.efficiency_12m = eff12m
        eg.efficiency_oos = effoos

        # Loss disqualifiers
        eg.loss_1m = l1m
        eg.loss_3m = l3m
        eg.loss_6m = l6m

        # Efficiency disqualifiers
        eg.efficiency_loss_1m = el1m
        eg.efficiency_loss_3m = el3m
        eg.efficiency_loss_6m = el6m

        # Status gates
        eg.use_incubation = use_inc
        eg.use_quitting   = use_quit

        # Count monthly profits
        eg.use_count_monthly_profits = use_cpm
        eg.min_positive_months       = int(min_pos_months)
        eg.eligibility_months        = int(elig_months_total)
        eg.monthly_profit_operator   = monthly_op

        # Additional filter
        eg.additional_user_filter              = add_filter
        eg.additional_user_filter_column       = add_col.strip()
        eg.additional_user_filter_min_value    = float(add_min)

        # Thresholds
        eg.efficiency_ratio    = float(eff_ratio_pct) / 100.0
        eg.days_threshold_oos  = int(days_thresh)
        eg.oos_dd_vs_is_cap    = float(oos_dd_cap)
        eg.date_type           = date_type
        eg.max_horizon         = int(max_horizon)
        eg.status_include      = [s.strip() for s in status_raw.split(",") if s.strip()] or ["Live"]

        _save(nc)
        st.success("Saved.")

tab_elig.__exit__(None, None, None)
tab_analytics.__enter__()

# ══════════════════════════════════════════════════════════════════════════════
# 6. MONTE CARLO SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("Monte Carlo Settings")

with st.form("mc_form"):
    col_m1, col_m2, col_m3, col_m4, col_m5 = st.columns(5)
    with col_m1:
        mc_sims = st.number_input("Simulations", min_value=1_000, max_value=100_000, step=1_000,
                                  value=int(config.monte_carlo.simulations))
    with col_m2:
        mc_period = st.selectbox("Period", ["IS", "OOS", "IS+OOS"],
                                 index=["IS", "OOS", "IS+OOS"].index(config.monte_carlo.period))
    with col_m3:
        mc_trade = st.selectbox("Trade data", ["Closed", "M2M"],
                                index=0 if config.monte_carlo.trade_option == "Closed" else 1)
    with col_m4:
        mc_ror = st.number_input("Risk-of-ruin target", min_value=0.01, max_value=1.0,
                                 step=0.01, format="%.2f", value=float(config.monte_carlo.risk_ruin_target))
    with col_m5:
        mc_adj = st.number_input("Trade adjustment ($)", min_value=-10_000.0, max_value=10_000.0,
                                 step=100.0, format="%.0f", value=float(config.monte_carlo.trade_adjustment))

    if st.form_submit_button("Save MC Settings", type="primary"):
        nc = config.model_copy(deep=True)
        nc.monte_carlo.simulations   = int(mc_sims)
        nc.monte_carlo.period        = mc_period
        nc.monte_carlo.trade_option  = mc_trade
        nc.monte_carlo.risk_ruin_target   = float(mc_ror)
        nc.monte_carlo.trade_adjustment   = float(mc_adj)
        _save(nc)
        st.success("Saved.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# 7. CORRELATION THRESHOLDS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("Correlation Thresholds")
st.caption("Pairs above these thresholds are flagged as high-correlation.")

with st.form("corr_form"):
    col_c1, col_c2, col_c3, _ = st.columns(4)
    with col_c1:
        corr_n = st.slider("Normal mode",   0.0, 1.0, float(config.corr_normal_threshold),   0.05)
    with col_c2:
        corr_neg = st.slider("Negative mode", 0.0, 1.0, float(config.corr_negative_threshold), 0.05)
    with col_c3:
        corr_d = st.slider("Drawdown mode",  0.0, 1.0, float(config.corr_drawdown_threshold),  0.05)

    if st.form_submit_button("Save Correlation Settings", type="primary"):
        nc = config.model_copy(deep=True)
        nc.corr_normal_threshold   = float(corr_n)
        nc.corr_negative_threshold = float(corr_neg)
        nc.corr_drawdown_threshold = float(corr_d)
        _save(nc)
        st.success("Saved.")

tab_analytics.__exit__(None, None, None)
tab_portfolio.__enter__()

# ══════════════════════════════════════════════════════════════════════════════
# 8. MARGIN SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("Margin Settings")
st.caption(
    "Controls which margin data source is used in the Margin Tracking page.  \n"
    "**MultiWalk** — uses the Overnight Maintenance margin from the Walkforward Details CSV.  \n"
    "**TradeStation / Interactive Brokers** — looks up from the reference tables imported during migration.  \n"
    "**Manual** — per-symbol margins entered directly in the Margin Tracking page."
)

with st.form("margin_form"):
    col_mg1, col_mg2, col_mg3, _ = st.columns(4)

    _m_sources = ["MultiWalk", "TradeStation", "InteractiveBrokers", "Manual"]
    _m_types   = ["Maintenance", "Initial", "Average"]
    _safe_ms = config.margin_source if config.margin_source in _m_sources else "MultiWalk"
    _safe_mt = config.margin_type   if config.margin_type   in _m_types   else "Maintenance"

    with col_mg1:
        m_source = st.selectbox(
            "Margin source",
            _m_sources,
            index=_m_sources.index(_safe_ms),
            help="Mirrors VBA Margin_Source named range.",
        )
    with col_mg2:
        m_type = st.selectbox(
            "Margin type",
            _m_types,
            index=_m_types.index(_safe_mt),
            help=(
                "Maintenance = Overnight Maintenance (lower).  "
                "Initial = Overnight Initial (higher).  "
                "Average = (Initial + Maintenance) / 2.  "
                "Mirrors VBA Margin_Choice named range."
            ),
        )
    with col_mg3:
        m_default = st.number_input(
            "Default margin per contract ($)",
            min_value=0.0, max_value=500_000.0, step=500.0, format="%.0f",
            value=float(config.default_margin),
            help="Fallback when a symbol is not found in the selected margin source.",
        )

    if st.form_submit_button("Save Margin Settings", type="primary"):
        nc = config.model_copy(deep=True)
        nc.margin_source  = m_source
        nc.margin_type    = m_type
        nc.default_margin = float(m_default)
        _save(nc)
        st.success("Saved.")

# ══════════════════════════════════════════════════════════════════════════════
# 9. CONTRACT SIZING SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("Contract Sizing Settings")
st.caption(
    "Controls how contract counts are sized for portfolio backtests using the volatility "
    "(ATR) of each strategy.  \n"
    "**Starting equity** is used as the base for percentage-of-equity sizing.  \n"
    "**ATR blend** mixes dollar-ATR risk with margin requirement.  "
    "1.0 = pure ATR, 0.0 = pure margin."
)

with st.form("contract_sizing_form"):
    cs = config.contract_sizing
    col_cs1, col_cs2, col_cs3, col_cs4 = st.columns(4)

    with col_cs1:
        cs_equity = st.number_input(
            "Starting equity ($)",
            min_value=10_000.0, max_value=100_000_000.0, step=5_000.0,
            value=float(cs.starting_equity), format="%.0f",
            help="Base equity used for percentage-of-equity contract sizing.",
        )
        cs_pct = st.number_input(
            "Contract size % of equity",
            min_value=0.001, max_value=0.20, step=0.005, format="%.3f",
            value=float(cs.contract_size_pct_equity),
            help="1% of starting equity is allocated per contract (e.g. 0.01).",
        )

    with col_cs2:
        cs_cease_type = st.selectbox(
            "Cease trading type",
            ["Percentage", "Dollar"],
            index=0 if cs.cease_type == "Percentage" else 1,
            help="Stop adding new positions when portfolio drawdown hits this threshold.",
        )
        cs_cease_thresh = st.number_input(
            "Cease trading threshold",
            min_value=0.0, step=0.01 if cs.cease_type == "Percentage" else 1_000.0,
            value=float(cs.cease_trading_threshold), format="%.2f",
            help="0.25 = 25% drawdown from equity peak triggers cease.",
        )

    with col_cs3:
        cs_atr_window = st.selectbox(
            "ATR window",
            ["ATR Last 3 Months", "ATR Last 6 Months", "ATR Last 12 Months"],
            index=["ATR Last 3 Months", "ATR Last 6 Months", "ATR Last 12 Months"].index(cs.atr_window),
            help="Rolling window used to compute dollar ATR from trade MFE+MAE.",
        )
        cs_margin_mult = st.slider(
            "Margin multiple",
            0.0, 2.0, float(cs.contract_margin_multiple), 0.05,
            help="Fraction of margin requirement used in sizing (0.50 = 50%).",
        )

    with col_cs4:
        cs_blend = st.slider(
            "ATR vs Margin blend",
            0.0, 1.0, float(cs.contract_ratio_margin_atr), 0.05,
            help="0 = pure margin sizing, 1 = pure ATR sizing.",
        )
        cs_rw_scope = st.selectbox(
            "Reweighting scope",
            ["None", "All", "Index Only"],
            index=["None", "All", "Index Only"].index(cs.reweight_scope),
            help="Which strategies' historical contracts are rescaled by ATR ratio.",
        )
        cs_rw_gain = st.slider(
            "Reweight gain",
            0.5, 3.0, float(cs.reweight_gain), 0.05,
            disabled=(cs_rw_scope == "None"),
            help="Multiplier applied on top of ATR reweighting (1.0 = no gain).",
        )

    if st.form_submit_button("Save Contract Sizing Settings", type="primary"):
        nc = config.model_copy(deep=True)
        nc.contract_sizing.starting_equity           = float(cs_equity)
        nc.contract_sizing.contract_size_pct_equity  = float(cs_pct)
        nc.contract_sizing.cease_type                = cs_cease_type
        nc.contract_sizing.cease_trading_threshold   = float(cs_cease_thresh)
        nc.contract_sizing.atr_window                = cs_atr_window
        nc.contract_sizing.contract_margin_multiple  = float(cs_margin_mult)
        nc.contract_sizing.contract_ratio_margin_atr = float(cs_blend)
        nc.contract_sizing.reweight_scope            = cs_rw_scope
        nc.contract_sizing.reweight_gain             = float(cs_rw_gain)
        _save(nc)
        st.success("Saved.")

st.divider()


# ══════════════════════════════════════════════════════════════════════════════
# 10. STRATEGY RANKING SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("Strategy Ranking Settings")
st.caption(
    "Default ranking behaviour for the Eligibility Backtest → Strategy Screener tab.  \n"
    "Strategies are ranked by the chosen metric and can be grouped by sector."
)

with st.form("ranking_form"):
    _RANK_METRIC_LABELS = {
        "rtd_oos":                "Return-to-Drawdown (OOS)",
        "rtd_12_months":          "Return-to-Drawdown (12M)",
        "sharpe_isoos":           "Sharpe IS+OOS",
        "profit_since_oos_start": "Total OOS Profit ($)",
        "profit_last_12_months":  "Last 12M Profit ($)",
        "k_factor":               "K-Factor",
        "ulcer_index":            "Ulcer Index (lower = better)",
        "contracts":              "Contracts",
    }
    rk = config.ranking
    col_rk1, col_rk2, col_rk3 = st.columns([2, 1, 1])

    with col_rk1:
        rk_metric = st.selectbox(
            "Default ranking metric",
            list(_RANK_METRIC_LABELS.keys()),
            index=list(_RANK_METRIC_LABELS.keys()).index(rk.metric),
            format_func=lambda k: _RANK_METRIC_LABELS[k],
            help="Primary sort metric for the Strategy Screener.",
        )
    with col_rk2:
        rk_ascending = st.checkbox(
            "Ascending (lower = better)",
            value=rk.ascending,
            help="Enable for Ulcer Index and similar lower-is-better metrics.",
        )
        rk_elig_only = st.checkbox(
            "Eligible strategies only",
            value=rk.eligible_only,
            help="Hide ineligible strategies from the ranking view.",
        )
    with col_rk3:
        rk_sector = st.checkbox(
            "Group by sector",
            value=rk.group_by_sector,
            help="Sort and group the ranking table by sector.",
        )
        rk_contracts = st.checkbox(
            "Sub-sort by contracts",
            value=rk.group_by_contracts,
            help="Break ties within each sector by contract count.",
        )

    if st.form_submit_button("Save Ranking Settings", type="primary"):
        nc = config.model_copy(deep=True)
        nc.ranking.metric             = rk_metric
        nc.ranking.ascending          = rk_ascending
        nc.ranking.eligible_only      = rk_elig_only
        nc.ranking.group_by_sector    = rk_sector
        nc.ranking.group_by_contracts = rk_contracts
        _save(nc)
        st.success("Saved.")

tab_portfolio.__exit__(None, None, None)

# ── Saved banner ──────────────────────────────────────────────────────────────
if _any_saved:
    st.info(
        "Settings saved — go to **Portfolio** and click **Rebuild Portfolio** to apply.",
        icon="✅",
    )

st.divider()
_nav = st.columns(4)
with _nav[0]:
    st.page_link("ui/pages/01_Import.py", label="→ Import")
with _nav[1]:
    st.page_link("ui/pages/02_Strategies.py", label="→ Strategies")
with _nav[2]:
    st.page_link("ui/pages/03_Portfolio.py", label="→ Portfolio")
