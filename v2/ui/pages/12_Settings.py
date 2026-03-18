"""
Settings page — license, config backup/restore, and all export options.

Sections:
  1. License        — status + update TradeStation Customer ID
  2. Export / Import — ZIP config backup/restore
  3. Excel Export   — Raw Data | Computed Output | Correlations
  4. PDF Export     — Portfolio summary report
  5. App Preferences — date format, correlation thresholds
"""

from __future__ import annotations

import streamlit as st

from core.config import AppConfig, StrategyRankingConfig
from core.reporting.settings_io import (
    default_export_filename,
    export_settings,
    import_settings,
)

st.set_page_config(page_title="Settings", layout="wide")
st.title("Settings")

config: AppConfig          = st.session_state.get("config", AppConfig.load())
imported                   = st.session_state.get("imported_data")
portfolio                  = st.session_state.get("portfolio_data")
mc_result                  = st.session_state.get("mc_result")
mc_label                   = st.session_state.get("mc_target_label", "Portfolio")
loo_result                 = st.session_state.get("loo_result")
loo_base_profit            = st.session_state.get("loo_base_profit", 0.0)
loo_base_sharpe            = st.session_state.get("loo_base_sharpe", 0.0)
corr_data                  = st.session_state.get("corr_matrices") or st.session_state.get("correlation_result")


# ── 1. License ────────────────────────────────────────────────────────────────
st.header("License")

from core.licensing.license_manager import is_known_customer, validate_full

current_id = config.customer_id
if current_id:
    col_a, col_b = st.columns([1, 3])
    with col_a:
        st.metric("Current Customer ID", current_id)
    with col_b:
        if is_known_customer(current_id):
            st.success("Customer ID is in the licensed-customer list.")
        else:
            st.error("Customer ID is NOT in the licensed-customer list.")
else:
    st.info("No TradeStation Customer ID configured yet.")

with st.expander("Change / verify Customer ID"):
    with st.form("update_license"):
        new_id = st.number_input(
            "TradeStation Customer ID",
            min_value=1, max_value=9_999_999,
            value=current_id or 1, step=1,
        )
        c1, c2 = st.columns(2)
        save_btn  = c1.form_submit_button("Save", type="primary", use_container_width=True)
        check_btn = c2.form_submit_button("Save & Check DLL", use_container_width=True)
        if save_btn or check_btn:
            config.customer_id = int(new_id)
            config.save()
            st.session_state.config = config
            if check_btn:
                with st.spinner("Checking license via MultiWalk DLL…"):
                    valid, msg = validate_full(int(new_id))
                if valid:
                    st.success("License validated successfully.")
                else:
                    st.error(f"License check failed: {msg}")
            else:
                st.success("Customer ID saved.")
            st.rerun()

st.divider()


# ── 2. Export / Import Config ─────────────────────────────────────────────────
st.header("Export / Import Settings")

col_exp, col_imp = st.columns(2)

with col_exp:
    st.subheader("Export")
    st.caption("Download a ZIP of your strategies, margin tables, and app settings.")
    if st.button("Generate Config ZIP", use_container_width=True):
        zip_bytes = export_settings()
        st.download_button(
            label="⬇ Download config ZIP",
            data=zip_bytes,
            file_name=default_export_filename(),
            mime="application/zip",
            use_container_width=True,
        )

with col_imp:
    st.subheader("Import")
    st.caption("Upload a previously exported ZIP to restore your configuration. **Overwrites current settings.**")
    uploaded = st.file_uploader("Select config ZIP", type=["zip"], label_visibility="collapsed")
    if uploaded is not None:
        if st.button("Restore from ZIP", type="primary", use_container_width=True):
            ok, err, restored = import_settings(uploaded.read())
            if ok:
                st.success(f"Restored: {', '.join(restored)}")
                st.session_state.config = AppConfig.load()
                st.session_state.pop("margin_tables", None)
                st.rerun()
            else:
                st.error(f"Import failed: {err}")

st.divider()


# ── 3. Excel Export ───────────────────────────────────────────────────────────
st.header("Excel Export")

tab_raw, tab_output, tab_mc, tab_loo, tab_corr = st.tabs([
    "Raw Data", "Computed Output", "Monte Carlo", "Leave-One-Out", "Correlations"
])

# ── 3a. Raw Data ──────────────────────────────────────────────────────────────
with tab_raw:
    st.caption(
        "Export the raw imported DataFrames: Daily M2M, Closed Trades, "
        "In-Market Long/Short, and individual trade records."
    )
    if imported is None:
        st.info("No data loaded. Go to **Import** first.")
    else:
        start, end = imported.date_range
        n_strats = len(imported.strategy_names)
        st.metric("Strategies", n_strats)
        st.caption(f"Date range: {start} → {end} | Trades: {len(imported.trades):,}")
        if st.button("Export Raw Data to Excel", use_container_width=True, key="raw_xlsx"):
            try:
                from core.reporting.excel_export import export_raw_data, raw_data_export_filename
                xlsx = export_raw_data(imported)
                st.download_button(
                    "⬇ Download raw_data.xlsx", data=xlsx,
                    file_name=raw_data_export_filename(),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

# ── 3b. Computed Output ───────────────────────────────────────────────────────
with tab_output:
    st.caption(
        "Export computed strategy metrics (80+ columns) plus the "
        "portfolio cumulative equity curve."
    )
    if portfolio is None:
        st.info("No portfolio built. Go to **Portfolio** page first.")
    else:
        st.metric("Active strategies", len(portfolio.strategies))
        if st.button("Export Portfolio Output to Excel", use_container_width=True, key="output_xlsx"):
            try:
                from core.reporting.excel_export import export_portfolio, portfolio_export_filename
                xlsx = export_portfolio(portfolio, config)
                st.download_button(
                    "⬇ Download portfolio_output.xlsx", data=xlsx,
                    file_name=portfolio_export_filename(),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

# ── 3c. Monte Carlo ───────────────────────────────────────────────────────────
with tab_mc:
    st.caption("Export MC summary metrics and scenario distribution.")
    if mc_result is None:
        st.info("No MC results cached. Run **Monte Carlo** page first.")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Starting Equity",    f"${mc_result.starting_equity:,.0f}")
        c2.metric("Expected Profit",    f"${mc_result.expected_profit:,.0f}")
        c3.metric("Risk of Ruin",       f"{mc_result.risk_of_ruin:.1%}")
        if st.button("Export MC Results to Excel", use_container_width=True, key="mc_xlsx"):
            try:
                from core.reporting.excel_export import export_mc_result, mc_export_filename
                xlsx = export_mc_result(mc_result, label=mc_label)
                st.download_button(
                    "⬇ Download mc_results.xlsx", data=xlsx,
                    file_name=mc_export_filename(mc_label),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

# ── 3d. Leave-One-Out ─────────────────────────────────────────────────────────
with tab_loo:
    st.caption("Export the Leave-One-Out analysis table.")
    if loo_result is None:
        st.info("No LOO results cached. Run **Leave One Out** page first.")
    else:
        st.metric("Strategies analysed", len(loo_result))
        if st.button("Export LOO Results to Excel", use_container_width=True, key="loo_xlsx"):
            try:
                from core.reporting.excel_export import export_loo_result, loo_export_filename
                xlsx = export_loo_result(loo_result, loo_base_profit, loo_base_sharpe)
                st.download_button(
                    "⬇ Download leave_one_out.xlsx", data=xlsx,
                    file_name=loo_export_filename(),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

# ── 3e. Correlations ─────────────────────────────────────────────────────────
with tab_corr:
    st.caption("Export a correlation matrix to Excel.")
    if corr_data is None:
        st.info("No correlation results cached. Run **Correlations** page first.")
    else:
        if isinstance(corr_data, dict):
            mode = st.selectbox("Mode", list(corr_data.keys()), key="corr_mode_sel")
            corr_df = corr_data[mode]
        else:
            mode = "Normal"
            corr_df = corr_data
        st.caption(f"{len(corr_df)} × {len(corr_df.columns)} matrix")
        if st.button("Export Correlations to Excel", use_container_width=True, key="corr_xlsx"):
            try:
                from core.reporting.excel_export import export_correlations, correlations_export_filename
                xlsx = export_correlations(corr_df, mode)
                st.download_button(
                    "⬇ Download correlations.xlsx", data=xlsx,
                    file_name=correlations_export_filename(mode),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

st.divider()


# ── 4. PDF Export ─────────────────────────────────────────────────────────────
st.header("PDF Export")

st.caption(
    "Export a portfolio summary report as a PDF — cover page, overview metrics, "
    "MC summary (if available), and strategy list."
)

if portfolio is None:
    st.info("No portfolio built. Go to **Portfolio** page first.")
else:
    n_active = len(portfolio.strategies)
    col1, col2 = st.columns(2)
    col1.metric("Active strategies", n_active)
    if mc_result is not None:
        col2.metric("MC results included", "Yes")

    if st.button("Export Portfolio Report to PDF", type="primary", use_container_width=True):
        try:
            from core.reporting.pdf_export import export_portfolio_pdf, pdf_export_filename
            with st.spinner("Generating PDF…"):
                pdf_bytes = export_portfolio_pdf(portfolio, config, mc_result=mc_result)
            st.download_button(
                "⬇ Download portfolio_report.pdf", data=pdf_bytes,
                file_name=pdf_export_filename(),
                mime="application/pdf",
                use_container_width=True,
            )
        except ImportError as e:
            st.error(str(e))

st.divider()


# ── 5. App Preferences ────────────────────────────────────────────────────────
st.header("App Preferences")

with st.form("preferences_form"):
    st.subheader("Portfolio")
    period_years = st.number_input(
        "Lookback period (years)",
        min_value=0.5, max_value=20.0, step=0.5,
        value=float(config.portfolio.period_years), format="%.1f",
        help="How many years of history to use when computing summary metrics.",
    )
    use_cutoff = st.checkbox(
        "Use cutoff date",
        value=config.portfolio.use_cutoff,
        help="Treat data after this date as OOS-end for all strategies.",
    )
    cutoff_str = st.text_input(
        "Cutoff date (YYYY-MM-DD)",
        value=config.portfolio.cutoff_date or "",
        disabled=not use_cutoff,
        help="Only used when 'Use cutoff date' is checked.",
    )

    st.subheader("Monte Carlo defaults")
    mc_sims = st.number_input(
        "Default simulations",
        min_value=1_000, max_value=100_000, step=1_000,
        value=int(config.monte_carlo.simulations),
    )
    mc_period = st.selectbox(
        "Default period",
        ["OOS", "IS", "IS+OOS"],
        index=["OOS", "IS", "IS+OOS"].index(config.monte_carlo.period),
    )
    mc_ror = st.slider(
        "Default risk-of-ruin target %",
        min_value=1, max_value=30,
        value=int(config.monte_carlo.risk_ruin_target * 100),
    )
    mc_trade_opt = st.radio(
        "Default trade data",
        ["M2M", "Closed"],
        index=["M2M", "Closed"].index(config.monte_carlo.trade_option),
        horizontal=True,
    )

    st.subheader("Portfolio Contract Sizing")
    cs = config.contract_sizing
    starting_equity = st.number_input(
        "Starting equity ($)",
        min_value=10_000.0, max_value=100_000_000.0, step=5_000.0,
        value=float(cs.starting_equity), format="%.0f",
        help="Used as fixed starting equity when 'Solve for ROR' is off.",
    )
    solve_for_ror = st.checkbox(
        "Solve for Risk-of-Ruin target (overrides starting equity)",
        value=config.monte_carlo.solve_for_ror,
        help="When on, the MC engine iterates equity until portfolio ROR matches the target.",
    )
    col_cease1, col_cease2 = st.columns(2)
    with col_cease1:
        cease_type = st.selectbox(
            "Cease trading type",
            ["Percentage", "Dollar"],
            index=0 if cs.cease_type == "Percentage" else 1,
            help="Stop adding new positions when portfolio drawdown hits threshold.",
        )
    with col_cease2:
        cease_threshold = st.number_input(
            "Cease trading threshold",
            min_value=0.0, max_value=1.0 if cease_type == "Percentage" else 10_000_000.0,
            step=0.01 if cease_type == "Percentage" else 1_000.0,
            value=float(cs.cease_trading_threshold),
            format="%.2f" if cease_type == "Percentage" else "%.0f",
            help="0.25 = 25% drawdown from equity peak triggers cease",
        )

    st.subheader("Contract Sizing — Estimated Vol (ATR + Margin blend)")
    col_atr1, col_atr2 = st.columns(2)
    with col_atr1:
        atr_window = st.selectbox(
            "ATR window",
            ["ATR Last 3 Months", "ATR Last 6 Months", "ATR Last 12 Months"],
            index=["ATR Last 3 Months", "ATR Last 6 Months", "ATR Last 12 Months"].index(cs.atr_window),
            help="Rolling window used to compute dollar ATR from trade MFE+MAE",
        )
        contract_margin_multiple = st.slider(
            "Margin multiple",
            0.0, 2.0, float(cs.contract_margin_multiple), 0.05,
            help="Fraction of margin requirement used in sizing (0.50 = 50%)",
        )
    with col_atr2:
        contract_ratio = st.slider(
            "ATR vs Margin ratio",
            0.0, 1.0, float(cs.contract_ratio_margin_atr), 0.05,
            help="0 = pure margin sizing, 1 = pure ATR sizing, 0.5 = equal blend",
        )
        contract_pct_equity = st.number_input(
            "Contract size % of equity",
            min_value=0.001, max_value=0.20, step=0.005,
            value=float(cs.contract_size_pct_equity), format="%.3f",
            help="1% of starting equity per contract position (e.g. 0.01)",
        )

    st.subheader("Portfolio Backtest Historical Sizing")
    col_rw1, col_rw2 = st.columns(2)
    with col_rw1:
        reweight_scope = st.selectbox(
            "ATR reweighting scope",
            ["None", "All", "Index Only"],
            index=["None", "All", "Index Only"].index(cs.reweight_scope),
            help=(
                "None = no ATR scaling. "
                "All = scale every strategy's historical contracts by current_ATR / historical_ATR. "
                "Index Only = restrict scaling to index/benchmark contracts."
            ),
        )
    with col_rw2:
        reweight_gain = st.slider(
            "Reweight gain (multiplier)",
            min_value=0.5, max_value=3.0,
            value=float(cs.reweight_gain), step=0.05,
            disabled=(reweight_scope == "None"),
            help="Scale factor applied on top of ATR reweighting. 1.0 = no additional gain.",
        )

    # ── Buy & Hold strategies ─────────────────────────────────────────────────
    _bh_strats = []
    if portfolio is not None:
        _bh_strats = [
            s for s in portfolio.strategies
            if "buy" in s.status.lower() and "hold" in s.status.lower()
        ]
    elif imported is not None:
        _bh_strats = [
            s for s in imported.strategies
            if "buy" in s.status.lower() and "hold" in s.status.lower()
        ]

    with st.expander(
        f"Buy & Hold Strategies ({len(_bh_strats)} loaded)",
        expanded=bool(_bh_strats),
    ):
        if not _bh_strats:
            st.caption("No Buy & Hold strategies found in the current portfolio/imported data.")
        else:
            import pandas as pd
            _bh_rows = [
                {
                    "Name":     s.name,
                    "Symbol":   s.symbol,
                    "Sector":   s.sector,
                    "Contracts": s.contracts,
                    "OOS Start": str(s.oos_start or "—"),
                }
                for s in _bh_strats
            ]
            st.dataframe(pd.DataFrame(_bh_rows), hide_index=True, use_container_width=True)
            st.caption(
                "B&H strategies are benchmarks only — excluded from eligibility scoring "
                "when 'Exclude Buy & Hold' is enabled in Eligibility Settings."
            )

    st.subheader("Monte Carlo — Additional Settings")
    mc_output_samples = st.number_input(
        "Output samples",
        min_value=1, max_value=500, step=5,
        value=int(config.monte_carlo.output_samples),
        help="Number of scenario paths to include in output",
    )
    mc_remove_best = st.slider(
        "Remove % best days/weeks before MC",
        0.0, 0.10, float(config.monte_carlo.remove_best_pct), 0.005,
        format="%.1%%",
        help="Trim top-N% best trading days to stress-test the distribution",
    )

    st.subheader("Strategy Ranking defaults")
    _RANK_METRICS = {
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
    rank_metric = st.selectbox(
        "Default ranking metric",
        list(_RANK_METRICS.keys()),
        index=list(_RANK_METRICS.keys()).index(rk.metric),
        format_func=lambda k: _RANK_METRICS[k],
        help="Metric used to rank strategies in the Strategy Screener.",
    )
    rank_ascending = st.checkbox(
        "Rank ascending (lower = better)",
        value=rk.ascending,
        help="Enable for Ulcer Index or other 'lower is better' metrics.",
    )
    rank_eligible_only = st.checkbox(
        "Ranking: eligible strategies only",
        value=rk.eligible_only,
        help="When on, only base-eligible strategies appear in the ranking view.",
    )
    col_rk1, col_rk2 = st.columns(2)
    with col_rk1:
        rank_group_sector = st.checkbox(
            "Group ranking by sector",
            value=rk.group_by_sector,
            help="Display strategies grouped under their sector header.",
        )
    with col_rk2:
        rank_group_contracts = st.checkbox(
            "Sub-sort by contracts within sector",
            value=rk.group_by_contracts,
            help="Within each sector group, break ties by contract count (descending).",
        )

    st.subheader("Eligibility defaults")
    elig_status = st.multiselect(
        "Default eligible statuses",
        ["Live", "Paper", "Pass", "Retired"],
        default=list(config.eligibility.status_include),
    )

    with st.expander("ℹ️ How the Month-End Days Threshold works", expanded=False):
        st.markdown(
            """
**Month-End Days Threshold** controls how profit windows (1M, 3M, 6M…) are anchored:

| Value | Mode | Behaviour |
|-------|------|-----------|
| **0** | Rolling | Windows count exact calendar months back from today (e.g. 15 Mar → 15 Feb for 1M). |
| **1–31** | Calendar snap | If the current month has **fewer days** than this threshold, treat last month-end as the effective end date and snap all windows to the 1st of each month. Otherwise use today. |

**Examples with threshold = 5:**
- Today = **Mar 3** (3 days into month, < 5) → effective end = **Feb 28**; 1M = Feb 1–28, 3M = Dec 1–Feb 28
- Today = **Mar 15** (15 days into month, ≥ 5) → effective end = **Mar 15**; 1M = Mar 1–15, 3M = Jan 1–Mar 15

Use a higher threshold (e.g. 10–15) for month-end reporting so results stay stable during the first days of a new month.
Set to **0** for a rolling window that always ends today.
            """
        )

    elig_days = st.number_input(
        "Month-End Days Threshold (0 = rolling, 1–31 = calendar snap)",
        min_value=0, max_value=31, step=1,
        value=min(int(config.eligibility.days_threshold_oos), 31),
        help=(
            "EligibilityDaysThreshold — 0 = rolling windows ending today. "
            "1–31 = snap profit windows to calendar month boundaries: if fewer than "
            "this many days have elapsed in the current month, use the previous "
            "month-end as the effective end date."
        ),
    )
    elig_eff = st.slider(
        "Default efficiency ratio",
        0.0, 2.0, float(config.eligibility.efficiency_ratio), 0.05,
    )

    st.subheader("Display")
    date_fmt = st.selectbox(
        "Date format in CSV files",
        ["DMY", "MDY"],
        index=["DMY", "MDY"].index(config.date_format),
        help="DMY = DD/MM/YYYY (TradeStation default). MDY = MM/DD/YYYY.",
    )
    corr_normal = st.number_input(
        "Correlation threshold — Normal mode",
        min_value=0.0, max_value=1.0, step=0.05,
        value=config.corr_normal_threshold, format="%.2f",
    )
    corr_drawdown = st.number_input(
        "Correlation threshold — Drawdown mode",
        min_value=0.0, max_value=1.0, step=0.05,
        value=config.corr_drawdown_threshold, format="%.2f",
    )
    corr_negative = st.number_input(
        "Correlation threshold — Negative mode",
        min_value=0.0, max_value=1.0, step=0.05,
        value=config.corr_negative_threshold, format="%.2f",
        help="Pairs BELOW this are considered negatively correlated (diversifying).",
    )

    if st.form_submit_button("Save All Preferences", type="primary"):
        # Portfolio
        config.portfolio.period_years = float(period_years)
        config.portfolio.use_cutoff   = use_cutoff
        if use_cutoff and cutoff_str.strip():
            try:
                from datetime import date as _date
                _date.fromisoformat(cutoff_str.strip())   # validate
                config.portfolio.cutoff_date = cutoff_str.strip()
            except ValueError:
                st.error("Invalid cutoff date — use YYYY-MM-DD format.")
                st.stop()
        elif not use_cutoff:
            config.portfolio.cutoff_date = None

        # Contract sizing
        config.contract_sizing.starting_equity           = float(starting_equity)
        config.contract_sizing.cease_type                = cease_type
        config.contract_sizing.cease_trading_threshold   = float(cease_threshold)
        config.contract_sizing.atr_window                = atr_window
        config.contract_sizing.contract_margin_multiple  = float(contract_margin_multiple)
        config.contract_sizing.contract_ratio_margin_atr = float(contract_ratio)
        config.contract_sizing.contract_size_pct_equity  = float(contract_pct_equity)
        config.contract_sizing.reweight_scope             = reweight_scope
        config.contract_sizing.reweight_gain              = float(reweight_gain)

        # Monte Carlo
        config.monte_carlo.simulations      = int(mc_sims)
        config.monte_carlo.period           = mc_period
        config.monte_carlo.risk_ruin_target = mc_ror / 100.0
        config.monte_carlo.trade_option     = mc_trade_opt
        config.monte_carlo.solve_for_ror    = solve_for_ror
        config.monte_carlo.output_samples   = int(mc_output_samples)
        config.monte_carlo.remove_best_pct  = float(mc_remove_best)

        # Strategy Ranking
        config.ranking.metric            = rank_metric
        config.ranking.ascending         = rank_ascending
        config.ranking.eligible_only     = rank_eligible_only
        config.ranking.group_by_sector   = rank_group_sector
        config.ranking.group_by_contracts = rank_group_contracts

        # Eligibility
        config.eligibility.status_include      = elig_status if elig_status else ["Live"]
        config.eligibility.days_threshold_oos  = int(elig_days)
        config.eligibility.efficiency_ratio    = float(elig_eff)

        # Display
        config.date_format             = date_fmt
        config.corr_normal_threshold   = corr_normal
        config.corr_drawdown_threshold = corr_drawdown
        config.corr_negative_threshold = corr_negative

        config.save()
        st.session_state.config = config
        st.success("All preferences saved.")
