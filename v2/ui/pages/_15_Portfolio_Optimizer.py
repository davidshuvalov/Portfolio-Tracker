"""
Portfolio Optimizer — suggest an optimal, diversified portfolio.

Configurable workflow:  filter → size → rank → select → adjust.

Users can enable/disable steps, reorder them, and tune every parameter.
The optimizer builds on existing analytics (ATR, correlations, eligibility)
already computed elsewhere in the app.
"""

from __future__ import annotations

import streamlit as st
import pandas as pd

from core.config import AppConfig
from core.portfolio.strategies import load_strategies, save_strategies

st.set_page_config(page_title="Portfolio Optimizer", layout="wide")

st.title("Portfolio Optimizer")
st.caption(
    "Build a suggested portfolio from eligible strategies using a configurable "
    "workflow of filters, sizing, ranking, selection, and adjustment steps."
)

# ── Session state + config ────────────────────────────────────────────────────
config: AppConfig = st.session_state.get("config", AppConfig.load())
opt_cfg = config.optimizer
cs_cfg  = config.contract_sizing

# Initialise workflow state from config (once per session)
if "opt_workflow" not in st.session_state:
    st.session_state.opt_workflow = list(opt_cfg.workflow_steps)
if "opt_enabled" not in st.session_state:
    st.session_state.opt_enabled = set(opt_cfg.enabled_steps)

_STEP_LABELS: dict[str, str] = {
    "filter_eligibility":      "Filter: Eligibility Gates",
    "filter_excluded_symbols": "Filter: Excluded Symbols",
    "filter_contract_size":    "Filter: Contract Too Large",
    "rank":                    "Rank: By Metric",
    "size_contracts":          "Size: ATR/Margin Blend",
    "select_strategies":       "Select: By Diversity + Margin",
    "adjust_correlations":     "Adjust: Correlation Limits",
    "adjust_gross_margins":    "Adjust: Gross Margin Limits",
    "adjust_drawdowns":        "Adjust: Drawdown Controls",
    "adjust_mc":               "Adjust: Monte Carlo Target",
}

_STEP_ICONS: dict[str, str] = {
    "filter_eligibility":      "🔵",
    "filter_excluded_symbols": "🔵",
    "filter_contract_size":    "🔵",
    "rank":                    "🟡",
    "size_contracts":          "🟢",
    "select_strategies":       "🟢",
    "adjust_correlations":     "🟠",
    "adjust_gross_margins":    "🟠",
    "adjust_drawdowns":        "🟠",
    "adjust_mc":               "🎲",
}

_ALL_STEPS = list(_STEP_LABELS.keys())

_RANK_METRIC_LABELS = {
    "rtd_oos":                "RTD (OOS)",
    "rtd_12_months":          "RTD (12M)",
    "sharpe_isoos":           "Sharpe IS+OOS",
    "profit_since_oos_start": "OOS Total ($)",
    "profit_last_12_months":  "Last 12M ($)",
    "profit_last_9_months":   "Last 9M ($)",
    "profit_last_6_months":   "Last 6M ($)",
    "profit_last_3_months":   "Last 3M ($)",
    "k_factor":               "K-Factor",
    "ulcer_index":            "Ulcer Index",
    "efficiency_oos":         "Efficiency (OOS)",
    "efficiency_12m":         "Efficiency (12M)",
    "efficiency_6m":          "Efficiency (6M)",
    "oos_monthly_win_rate":   "Monthly Win Rate (OOS)",
    "max_oos_drawdown":       "OOS Max Drawdown ($)",
    "sharpe_oos":             "Sharpe (OOS)",
}


# ── Sidebar: workflow builder ─────────────────────────────────────────────────
with st.sidebar:
    st.header("Workflow")
    st.caption("Enable/disable and reorder steps. Changes take effect on next run.")

    workflow: list[str] = st.session_state.opt_workflow
    enabled: set[str] = st.session_state.opt_enabled

    for i, step in enumerate(workflow):
        label = f"{_STEP_ICONS.get(step, '⚪')} {_STEP_LABELS.get(step, step)}"
        c_en, c_lbl, c_up, c_dn = st.columns([0.08, 0.62, 0.15, 0.15])
        with c_en:
            checked = st.checkbox(
                "en", value=(step in enabled),
                key=f"opt_en_{i}_{step}", label_visibility="collapsed",
            )
            if checked:
                enabled.add(step)
            else:
                enabled.discard(step)
        with c_lbl:
            st.caption(label)
        with c_up:
            if i > 0 and st.button("↑", key=f"opt_up_{i}", use_container_width=True):
                workflow[i - 1], workflow[i] = workflow[i], workflow[i - 1]
                st.rerun()
        with c_dn:
            if i < len(workflow) - 1 and st.button("↓", key=f"opt_dn_{i}", use_container_width=True):
                workflow[i + 1], workflow[i] = workflow[i], workflow[i + 1]
                st.rerun()

    st.divider()

    # Add / remove steps
    with st.expander("Add / Remove Steps", expanded=False):
        in_workflow = set(workflow)
        addable = [s for s in _ALL_STEPS if s not in in_workflow]
        if addable:
            new_step = st.selectbox(
                "Add step", options=addable,
                format_func=lambda s: _STEP_LABELS.get(s, s),
                key="opt_add_step_sel",
            )
            if st.button("Add →", key="opt_add_step_btn"):
                workflow.append(new_step)
                enabled.add(new_step)
                st.rerun()
        remove_step = st.selectbox(
            "Remove step", options=[""] + workflow,
            format_func=lambda s: _STEP_LABELS.get(s, s) if s else "— choose —",
            key="opt_rm_step_sel",
        )
        if remove_step and st.button("Remove ✕", key="opt_rm_step_btn"):
            workflow.remove(remove_step)
            enabled.discard(remove_step)
            st.rerun()

    if st.button("Reset to Defaults", key="opt_reset_workflow"):
        st.session_state.opt_workflow = list(opt_cfg.workflow_steps)
        st.session_state.opt_enabled  = set(opt_cfg.enabled_steps)
        st.rerun()

    st.divider()
    st.header("Parameters")

    # ── Equity ────────────────────────────────────────────────────────────
    equity_input = st.number_input(
        "Account Equity ($)",
        min_value=10_000.0, max_value=100_000_000.0, step=10_000.0,
        value=float(cs_cfg.starting_equity), format="%.0f",
        key="opt_equity",
    )

    # ── Contract sizing ───────────────────────────────────────────────────
    with st.expander("Contract Sizing", expanded=False):
        opt_ratio = st.slider(
            "ATR / Margin blend",
            0.0, 1.0, float(cs_cfg.contract_ratio_margin_atr), 0.05,
            help="0 = pure margin, 1 = pure ATR.",
            key="opt_blend",
        )
        opt_margin_multiple = st.slider(
            "Margin multiple",
            0.1, 1.0, float(cs_cfg.contract_margin_multiple), 0.05,
            help="Fraction of full maintenance margin to use in sizing.",
            key="opt_margin_mult",
        )
        opt_pct_equity = st.number_input(
            "% equity per contract",
            min_value=0.001, max_value=0.10, step=0.001,
            value=float(cs_cfg.contract_size_pct_equity), format="%.3f",
            key="opt_pct_eq",
        )
        opt_atr_window = st.selectbox(
            "ATR window",
            ["ATR Last 3 Months", "ATR Last 6 Months", "ATR Last 12 Months"],
            index=["ATR Last 3 Months", "ATR Last 6 Months", "ATR Last 12 Months"]
                  .index(cs_cfg.atr_window),
            key="opt_atr_window",
        )
        opt_min_fraction = st.selectbox(
            "Minimum contract fraction",
            [0.1, 0.25, 0.5, 1.0],
            index=0,
            help="Contracts are rounded down to this increment.",
            key="opt_min_frac",
        )
        opt_min_threshold = st.number_input(
            "Contract size threshold",
            min_value=0.1, max_value=2.0, step=0.05,
            value=float(opt_cfg.min_contract_size_threshold), format="%.2f",
            help="Strategies with computed contracts < this are excluded (contract too large).",
            key="opt_min_thresh",
        )

    # ── Eligibility ───────────────────────────────────────────────────────
    with st.expander("Eligibility Settings", expanded=False):
        _e = config.eligibility
        st.caption("Override eligibility rules for this optimiser run.")

        ec1, ec2 = st.columns(2)
        with ec1:
            st.markdown("**Profit Gates**")
            st.checkbox("Last 12M > $0",         value=_e.profit_12m,    key="opt_elig_p12m")
            st.checkbox("Last 3M OR 6M > $0",    value=_e.profit_3or6m, key="opt_elig_p3or6m")
            st.checkbox("Since OOS start > $0",  value=_e.profit_oos,   key="opt_elig_poos")
            st.checkbox("Last 9M > $0",          value=_e.profit_9m,    key="opt_elig_p9m")
            st.checkbox("Last 6M > $0",          value=_e.profit_6m,    key="opt_elig_p6m")
            st.checkbox("Last 3M > $0",          value=_e.profit_3m,    key="opt_elig_p3m")
            st.checkbox("Last 1M > $0",          value=_e.profit_1m,    key="opt_elig_p1m")
            st.markdown("**Loss Disqualifiers**")
            st.checkbox("Last 1M < $0 → exclude", value=_e.loss_1m, key="opt_elig_l1m")
            st.checkbox("Last 3M < $0 → exclude", value=_e.loss_3m, key="opt_elig_l3m")
            st.checkbox("Last 6M < $0 → exclude", value=_e.loss_6m, key="opt_elig_l6m")

        with ec2:
            st.markdown("**Efficiency Gates**")
            st.checkbox("OOS Efficiency > ratio",  value=_e.efficiency_oos,  key="opt_elig_effoos")
            st.checkbox("12M Efficiency > ratio",  value=_e.efficiency_12m,  key="opt_elig_eff12m")
            st.checkbox("6M Efficiency > ratio",   value=_e.efficiency_6m,   key="opt_elig_eff6m")
            st.checkbox("3M Efficiency > ratio",   value=_e.efficiency_3m,   key="opt_elig_eff3m")
            st.number_input(
                "Efficiency ratio (%)",
                min_value=0, max_value=500, step=5,
                value=int(round(_e.efficiency_ratio * 100)),
                key="opt_elig_effratio",
            )
            st.markdown("**Status Gates**")
            st.checkbox("Incubation must be Passed",   value=_e.use_incubation, key="opt_elig_inc")
            st.checkbox("Exclude Quit strategies",     value=_e.use_quitting,  key="opt_elig_quit")
            st.markdown("**Thresholds**")
            st.number_input(
                "Max OOS DD / IS DD ratio",
                min_value=0.0, max_value=10.0, step=0.1, format="%.1f",
                value=float(_e.oos_dd_vs_is_cap),
                key="opt_elig_oos_dd",
            )

    # ── Ranking ───────────────────────────────────────────────────────────
    with st.expander("Ranking", expanded=False):
        opt_rank_metric = st.selectbox(
            "Rank by",
            list(_RANK_METRIC_LABELS.keys()),
            index=list(_RANK_METRIC_LABELS.keys()).index(
                opt_cfg.rank_metric if opt_cfg.rank_metric in _RANK_METRIC_LABELS
                else "rtd_oos"
            ),
            format_func=lambda k: _RANK_METRIC_LABELS[k],
            key="opt_rank_metric",
        )
        opt_rank_asc = st.checkbox(
            "Ascending", value=opt_cfg.rank_ascending, key="opt_rank_asc",
            help="Ascending = lower values are better (e.g. Ulcer Index).",
        )

    # ── Selection ─────────────────────────────────────────────────────────
    with st.expander("Strategy Selection", expanded=False):
        opt_max_strats = st.number_input(
            "Max strategies",
            min_value=5, max_value=200, step=5,
            value=int(opt_cfg.max_strategies), format="%d",
            key="opt_max_strats",
        )
        opt_max_margin = st.slider(
            "Max total margin (% of equity)",
            0.10, 1.50, float(opt_cfg.max_margin_pct), 0.05,
            format="%.0f%%",
            key="opt_max_margin",
        )
        opt_per_symbol_first = st.checkbox(
            "Add best-per-symbol first",
            value=opt_cfg.per_symbol_first, key="opt_sym_first",
            help="Pass 1: one strategy per unique symbol, then fill by rank.",
        )
        opt_excluded_symbols = st.text_input(
            "Excluded symbols (comma-separated)",
            value=", ".join(opt_cfg.excluded_symbols),
            key="opt_excl_syms",
            help="e.g. 'S, YM' — strategies on these symbols are removed.",
        )

    # ── Correlation ───────────────────────────────────────────────────────
    with st.expander("Correlation Limits", expanded=False):
        opt_max_corr = st.slider(
            "Max positive correlation",
            0.0, 1.0, float(opt_cfg.max_correlation), 0.05,
            key="opt_max_corr",
        )
        opt_max_neg = st.slider(
            "Max negative correlation (abs)",
            0.0, 1.0, float(opt_cfg.max_negative_correlation), 0.05,
            key="opt_max_neg",
        )

    # ── Gross margins ─────────────────────────────────────────────────────
    with st.expander("Gross Margin Limits", expanded=False):
        st.caption(
            "Expressed as a share of total portfolio margin. "
            "1% = very tight; 100% = no limit."
        )
        opt_max_sym_pct = st.slider(
            "Max single symbol margin %",
            0.01, 1.0, float(opt_cfg.max_single_contract_margin_pct), 0.01,
            format="%.0f%%",
            key="opt_max_sym_pct",
        )
        opt_max_sec_pct = st.slider(
            "Max single sector margin %",
            0.01, 1.0, float(opt_cfg.max_sector_margin_pct), 0.01,
            format="%.0f%%",
            key="opt_max_sec_pct",
        )

    # ── Drawdowns ─────────────────────────────────────────────────────────
    with st.expander("Drawdown Controls", expanded=False):
        opt_max_avg_dd = st.slider(
            "Max avg strategy drawdown (% equity)",
            0.01, 1.0, float(opt_cfg.max_avg_drawdown_pct), 0.01,
            format="%.0f%%",
            key="opt_max_avg_dd",
        )
        opt_max_single_dd = st.slider(
            "Max single strategy drawdown (% equity)",
            0.01, 1.0, float(opt_cfg.max_single_drawdown_pct), 0.01,
            format="%.0f%%",
            key="opt_max_single_dd",
        )

    # ── Monte Carlo Targeting ─────────────────────────────────────────────
    with st.expander("Monte Carlo Target", expanded=False):
        _MC_MODE_LABELS = {
            "drawdown": "Max Drawdown",
            "margin":   "Margin Utilisation",
            "off":      "Off",
        }
        opt_mc_mode = st.radio(
            "Target mode",
            options=list(_MC_MODE_LABELS.keys()),
            index=list(_MC_MODE_LABELS.keys()).index(
                opt_cfg.mc_target_mode
                if opt_cfg.mc_target_mode in _MC_MODE_LABELS else "drawdown"
            ),
            format_func=lambda m: _MC_MODE_LABELS[m],
            key="opt_mc_mode",
            help=(
                "**Max Drawdown**: scale all contracts via MC simulation until "
                "the portfolio's median max drawdown hits the target.\n\n"
                "**Margin Utilisation**: scale contracts so total margin equals "
                "the target fraction of equity (no MC needed)."
            ),
        )
        if opt_mc_mode == "drawdown":
            st.slider(
                "Target max drawdown (% equity)",
                0.05, 0.50, float(opt_cfg.mc_target_drawdown_pct), 0.01,
                format="%.0f%%",
                key="opt_mc_dd_target",
                help="Contracts are scaled until MC median max dd ≈ this level.",
            )
        elif opt_mc_mode == "margin":
            st.slider(
                "Target margin utilisation (% equity)",
                0.10, 1.00, float(opt_cfg.mc_target_margin_pct), 0.05,
                format="%.0f%%",
                key="opt_mc_margin_target",
                help="Contracts are scaled so total margin usage = target × equity.",
            )
        st.number_input(
            "MC simulations",
            min_value=500, max_value=20_000, step=500,
            value=int(opt_cfg.mc_simulations), format="%d",
            key="opt_mc_sims",
            help="Number of scenarios per iteration. 2,000 balances speed and precision.",
        )
        st.slider(
            "Max scale factor",
            1.0, 5.0, float(opt_cfg.mc_max_scale), 0.5,
            key="opt_mc_max_scale",
            help="Cap on upward contract scaling. Prevents over-leveraging.",
        )

    st.divider()
    if st.button("Save as Defaults", key="opt_save_defaults"):
        config.optimizer.workflow_steps      = list(st.session_state.opt_workflow)
        config.optimizer.enabled_steps       = list(st.session_state.opt_enabled)
        config.optimizer.excluded_symbols    = [
            s.strip() for s in opt_excluded_symbols.split(",") if s.strip()
        ]
        config.optimizer.min_contract_size_threshold = float(opt_min_threshold)
        config.optimizer.rank_metric         = opt_rank_metric
        config.optimizer.rank_ascending      = opt_rank_asc
        config.optimizer.max_strategies      = int(opt_max_strats)
        config.optimizer.max_margin_pct      = float(opt_max_margin)
        config.optimizer.per_symbol_first    = opt_per_symbol_first
        config.optimizer.max_correlation     = float(opt_max_corr)
        config.optimizer.max_negative_correlation = float(opt_max_neg)
        config.optimizer.max_single_contract_margin_pct = float(opt_max_sym_pct)
        config.optimizer.max_sector_margin_pct          = float(opt_max_sec_pct)
        config.optimizer.max_avg_drawdown_pct    = float(opt_max_avg_dd)
        config.optimizer.max_single_drawdown_pct = float(opt_max_single_dd)
        config.optimizer.mc_target_mode          = st.session_state.get("opt_mc_mode", opt_cfg.mc_target_mode)
        config.optimizer.mc_target_drawdown_pct  = float(st.session_state.get("opt_mc_dd_target", opt_cfg.mc_target_drawdown_pct))
        config.optimizer.mc_target_margin_pct    = float(st.session_state.get("opt_mc_margin_target", opt_cfg.mc_target_margin_pct))
        config.optimizer.mc_simulations          = int(st.session_state.get("opt_mc_sims", opt_cfg.mc_simulations))
        config.optimizer.mc_max_scale            = float(st.session_state.get("opt_mc_max_scale", opt_cfg.mc_max_scale))
        config.save()
        st.session_state.config = config
        st.success("Defaults saved.")


# ── Prerequisite checks ────────────────────────────────────────────────────────
_imported = st.session_state.get("imported_data")
_summary_df: pd.DataFrame | None = st.session_state.get("all_strategies_summary_cache")
_strategies = load_strategies()

_prereqs_ok = True

if _imported is None:
    st.warning("No imported data found. Please import strategy data first.")
    st.page_link("ui/pages/01_Import.py", label="→ Go to Import")
    _prereqs_ok = False

if not _strategies:
    st.warning("No strategies configured. Please set up strategies first.")
    st.page_link("ui/pages/02_Strategies.py", label="→ Go to Strategies")
    _prereqs_ok = False

if _prereqs_ok and _summary_df is None:
    st.info(
        "Strategy summary not yet computed. Run **Compute / Refresh** on the "
        "Strategies page for the best results. The optimizer will still run "
        "but ranking metrics may be missing."
    )


# ── Run ────────────────────────────────────────────────────────────────────────
col_run, col_save_wf, _ = st.columns([1, 1, 5])

with col_run:
    run_btn = st.button(
        "▶ Run Optimizer",
        type="primary",
        disabled=not _prereqs_ok,
        use_container_width=True,
    )

with col_save_wf:
    if st.button("💾 Save Workflow", disabled=not _prereqs_ok, use_container_width=True):
        config.optimizer.workflow_steps = list(st.session_state.opt_workflow)
        config.optimizer.enabled_steps  = list(st.session_state.opt_enabled)
        config.save()
        st.session_state.config = config
        st.success("Workflow order saved.")


if run_btn and _prereqs_ok:
    with st.spinner("Running optimizer workflow…"):
        from core.portfolio.optimizer import (
            build_candidates,
            run_workflow,
            step_filter_eligibility,
            step_filter_excluded_symbols,
            step_filter_contract_size,
            step_rank,
            step_size_contracts,
            step_select_strategies,
            step_adjust_correlations,
            step_adjust_gross_margins,
            step_adjust_drawdowns,
            step_adjust_mc,
            portfolio_summary,
        )
        from core.analytics.atr import compute_atr
        from core.portfolio.summary import apply_eligibility_rules

        equity       = float(st.session_state.get("opt_equity", cs_cfg.starting_equity))
        margins      = config.symbol_margins
        default_margin = config.default_margin

        # ── 1. Compute ATR ────────────────────────────────────────────────
        trades_df = getattr(_imported, "trades", None)
        _atr_ser: pd.Series | None = None
        if trades_df is not None and not trades_df.empty:
            _atr_window = st.session_state.get("opt_atr_window", cs_cfg.atr_window)
            _atr_ser = compute_atr(trades_df, _atr_window)

        # ── 2. Build candidate list ───────────────────────────────────────
        candidates = build_candidates(
            strategies=_strategies,
            summary_df=_summary_df,
            atr_series=_atr_ser,
            margins=margins,
            default_margin=default_margin,
        )

        # ── 3. Eligibility mask (with per-run overrides) ─────────────────
        _eligible_mask: dict[str, bool] = {}
        if _summary_df is not None and not _summary_df.empty:
            try:
                _elig = config.eligibility.model_copy(deep=True)
                ss = st.session_state
                _elig.profit_1m       = bool(ss.get("opt_elig_p1m",    _elig.profit_1m))
                _elig.profit_3m       = bool(ss.get("opt_elig_p3m",    _elig.profit_3m))
                _elig.profit_6m       = bool(ss.get("opt_elig_p6m",    _elig.profit_6m))
                _elig.profit_9m       = bool(ss.get("opt_elig_p9m",    _elig.profit_9m))
                _elig.profit_12m      = bool(ss.get("opt_elig_p12m",   _elig.profit_12m))
                _elig.profit_3or6m    = bool(ss.get("opt_elig_p3or6m", _elig.profit_3or6m))
                _elig.profit_oos      = bool(ss.get("opt_elig_poos",   _elig.profit_oos))
                _elig.loss_1m         = bool(ss.get("opt_elig_l1m",    _elig.loss_1m))
                _elig.loss_3m         = bool(ss.get("opt_elig_l3m",    _elig.loss_3m))
                _elig.loss_6m         = bool(ss.get("opt_elig_l6m",    _elig.loss_6m))
                _elig.efficiency_oos  = bool(ss.get("opt_elig_effoos", _elig.efficiency_oos))
                _elig.efficiency_12m  = bool(ss.get("opt_elig_eff12m", _elig.efficiency_12m))
                _elig.efficiency_6m   = bool(ss.get("opt_elig_eff6m",  _elig.efficiency_6m))
                _elig.efficiency_3m   = bool(ss.get("opt_elig_eff3m",  _elig.efficiency_3m))
                _elig.efficiency_ratio = float(ss.get("opt_elig_effratio", int(round(_elig.efficiency_ratio * 100)))) / 100.0
                _elig.use_incubation  = bool(ss.get("opt_elig_inc",    _elig.use_incubation))
                _elig.use_quitting    = bool(ss.get("opt_elig_quit",   _elig.use_quitting))
                _elig.oos_dd_vs_is_cap = float(ss.get("opt_elig_oos_dd", _elig.oos_dd_vs_is_cap))
                _eligible_mask = apply_eligibility_rules(_summary_df, _elig)
            except Exception:
                pass

        # ── 4. Correlation matrix (and daily_m2m for MC step) ────────────
        daily_m2m = getattr(_imported, "daily_m2m", None)
        _corr: pd.DataFrame | None = st.session_state.get("correlation_matrix")
        if _corr is None and daily_m2m is not None and not daily_m2m.empty:
            try:
                from core.analytics.correlations import compute_correlation_matrix
                _corr = compute_correlation_matrix(daily_m2m)
            except Exception:
                pass

        # ── 5. Build workflow steps ───────────────────────────────────────
        _active_workflow = [
            s for s in st.session_state.opt_workflow
            if s in st.session_state.opt_enabled
        ]

        _excluded_syms = [
            s.strip() for s in st.session_state.get("opt_excl_syms", "").split(",")
            if s.strip()
        ]

        _step_map = {
            "filter_eligibility": (
                step_filter_eligibility,
                {"eligible_mask": _eligible_mask},
            ),
            "filter_excluded_symbols": (
                step_filter_excluded_symbols,
                {"excluded_symbols": _excluded_syms},
            ),
            "filter_contract_size": (
                step_filter_contract_size,
                {"min_threshold": float(st.session_state.get("opt_min_thresh", opt_cfg.min_contract_size_threshold))},
            ),
            "rank": (
                step_rank,
                {
                    "metric": st.session_state.get("opt_rank_metric", opt_cfg.rank_metric),
                    "ascending": bool(st.session_state.get("opt_rank_asc", opt_cfg.rank_ascending)),
                },
            ),
            "size_contracts": (
                step_size_contracts,
                {
                    "equity": equity,
                    "contract_size_pct": float(st.session_state.get("opt_pct_eq", cs_cfg.contract_size_pct_equity)),
                    "atr": dict(_atr_ser) if _atr_ser is not None else {},
                    "margins": margins,
                    "ratio": float(st.session_state.get("opt_blend", cs_cfg.contract_ratio_margin_atr)),
                    "contract_margin_multiple": float(st.session_state.get("opt_margin_mult", cs_cfg.contract_margin_multiple)),
                    "min_fraction": float(st.session_state.get("opt_min_frac", opt_cfg.min_contract_fraction)),
                },
            ),
            "select_strategies": (
                step_select_strategies,
                {
                    "margins": margins,
                    "contract_margin_multiple": float(st.session_state.get("opt_margin_mult", cs_cfg.contract_margin_multiple)),
                    "max_margin_pct": float(st.session_state.get("opt_max_margin", opt_cfg.max_margin_pct)),
                    "max_strategies": int(st.session_state.get("opt_max_strats", opt_cfg.max_strategies)),
                    "per_symbol_first": bool(st.session_state.get("opt_sym_first", opt_cfg.per_symbol_first)),
                },
            ),
            "adjust_correlations": (
                step_adjust_correlations,
                {
                    "corr_matrix": _corr,
                    "max_corr": float(st.session_state.get("opt_max_corr", opt_cfg.max_correlation)),
                    "max_neg_corr": float(st.session_state.get("opt_max_neg", opt_cfg.max_negative_correlation)),
                },
            ),
            "adjust_gross_margins": (
                step_adjust_gross_margins,
                {
                    "margins": margins,
                    "contract_margin_multiple": float(st.session_state.get("opt_margin_mult", cs_cfg.contract_margin_multiple)),
                    "equity": equity,
                    "max_single_pct": float(st.session_state.get("opt_max_sym_pct", opt_cfg.max_single_contract_margin_pct)),
                    "max_sector_pct": float(st.session_state.get("opt_max_sec_pct", opt_cfg.max_sector_margin_pct)),
                },
            ),
            "adjust_drawdowns": (
                step_adjust_drawdowns,
                {
                    "equity": equity,
                    "max_avg_pct": float(st.session_state.get("opt_max_avg_dd", opt_cfg.max_avg_drawdown_pct)),
                    "max_single_pct": float(st.session_state.get("opt_max_single_dd", opt_cfg.max_single_drawdown_pct)),
                    "max_single_trade_pct": float(opt_cfg.max_single_trade_loss_pct),
                    "min_fraction": float(st.session_state.get("opt_min_frac", opt_cfg.min_contract_fraction)),
                },
            ),
            "adjust_mc": (
                step_adjust_mc,
                {
                    "margins": margins,
                    "contract_margin_multiple": float(st.session_state.get("opt_margin_mult", cs_cfg.contract_margin_multiple)),
                    "daily_m2m": daily_m2m,
                    "target_drawdown_pct": (
                        float(st.session_state.get("opt_mc_dd_target", opt_cfg.mc_target_drawdown_pct))
                        if st.session_state.get("opt_mc_mode", opt_cfg.mc_target_mode) == "drawdown"
                        else None
                    ),
                    "target_margin_pct": (
                        float(st.session_state.get("opt_mc_margin_target", opt_cfg.mc_target_margin_pct))
                        if st.session_state.get("opt_mc_mode", opt_cfg.mc_target_mode) == "margin"
                        else None
                    ),
                    "n_simulations": int(st.session_state.get("opt_mc_sims", opt_cfg.mc_simulations)),
                    "max_scale": float(st.session_state.get("opt_mc_max_scale", opt_cfg.mc_max_scale)),
                    "tolerance": float(opt_cfg.mc_tolerance),
                    "min_fraction": float(st.session_state.get("opt_min_frac", opt_cfg.min_contract_fraction)),
                },
            ),
        }

        steps = [_step_map[s] for s in _active_workflow if s in _step_map]
        state = run_workflow(steps, candidates, equity)
        st.session_state["opt_result"] = state

# ── Results ────────────────────────────────────────────────────────────────────
_result = st.session_state.get("opt_result")

if _result is not None:
    from core.portfolio.optimizer import portfolio_summary

    _margins     = config.symbol_margins
    _marg_mult   = float(st.session_state.get("opt_margin_mult", cs_cfg.contract_margin_multiple))
    _stats       = portfolio_summary(_result, _margins, _marg_mult)
    _equity_used = float(st.session_state.get("opt_equity", cs_cfg.starting_equity))

    # ── KPIs ──────────────────────────────────────────────────────────────
    st.divider()
    kc1, kc2, kc3, kc4, kc5 = st.columns(5)
    kc1.metric("Strategies selected", _stats["n_strategies"])
    kc2.metric("Strategies removed", _stats["n_excluded"])
    kc3.metric(
        "Total margin used",
        f"${_stats['total_margin']:,.0f}",
        f"{_stats['margin_pct_equity']:.1%} of equity",
    )
    kc4.metric("Top symbol margin", f"{_stats['top_symbol_pct']:.1%}")
    kc5.metric("Top sector margin", f"{_stats['top_sector_pct']:.1%}")

    tab_port, tab_log, tab_breakdown, tab_apply = st.tabs(
        ["📋 Portfolio", "📜 Step Log", "📊 Margin Breakdown", "✅ Apply"]
    )

    # ── Portfolio table ────────────────────────────────────────────────────
    with tab_port:
        if _result.candidates:
            _rows = []
            for c in _result.candidates:
                name   = c["name"]
                sym    = c.get("symbol", "")
                sec    = c.get("sector", "")
                n      = _result.contracts.get(name, 0.0)
                m      = _margins.get(sym, config.default_margin) * _marg_mult
                _rows.append({
                    "Strategy":         name,
                    "Symbol":           sym,
                    "Sector":           sec,
                    "Contracts":        n,
                    "Margin/contract":  _margins.get(sym, config.default_margin),
                    "Total margin":     n * m,
                    "Margin %":         (n * m) / _stats["total_margin"] if _stats["total_margin"] else 0.0,
                    "RTD OOS":          c.get("rtd_oos"),
                    "Last 12M ($)":     c.get("profit_last_12_months"),
                    "Last 3M ($)":      c.get("profit_last_3_months"),
                    "OOS Max DD ($)":   c.get("max_oos_drawdown"),
                    "ATR ($)":          c.get("atr"),
                })
            _port_df = pd.DataFrame(_rows)
            st.dataframe(
                _port_df,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Contracts":        st.column_config.NumberColumn(format="%.1f"),
                    "Margin/contract":  st.column_config.NumberColumn(format="$%.0f"),
                    "Total margin":     st.column_config.NumberColumn(format="$%.0f"),
                    "Margin %":         st.column_config.NumberColumn(format="%.1f%%"),
                    "RTD OOS":          st.column_config.NumberColumn(format="%.2f"),
                    "Last 12M ($)":     st.column_config.NumberColumn(format="$%.0f"),
                    "Last 3M ($)":      st.column_config.NumberColumn(format="$%.0f"),
                    "OOS Max DD ($)":   st.column_config.NumberColumn(format="$%.0f"),
                    "ATR ($)":          st.column_config.NumberColumn(format="$%.0f"),
                },
            )
        else:
            st.warning("No strategies selected — try relaxing constraints.")

    # ── Step log ───────────────────────────────────────────────────────────
    with tab_log:
        st.subheader("Workflow Log")
        for line in _result.log:
            st.text(line)

        if _result.excluded:
            st.subheader(f"Excluded Strategies ({len(_result.excluded)})")
            _excl_df = pd.DataFrame([
                {"Strategy": e.name, "Step": e.step, "Reason": e.reason}
                for e in _result.excluded
            ])
            st.dataframe(_excl_df, use_container_width=True, hide_index=True)

    # ── Margin breakdown ───────────────────────────────────────────────────
    with tab_breakdown:
        bc1, bc2 = st.columns(2)

        with bc1:
            st.subheader("By Symbol")
            _sym_rows = [
                {
                    "Symbol": sym,
                    "Margin ($)": m,
                    "Share": m / _stats["total_margin"] if _stats["total_margin"] else 0.0,
                }
                for sym, m in sorted(
                    _stats["symbol_margin"].items(),
                    key=lambda x: -x[1],
                )
            ]
            st.dataframe(
                pd.DataFrame(_sym_rows),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Margin ($)": st.column_config.NumberColumn(format="$%.0f"),
                    "Share":      st.column_config.NumberColumn(format="%.1f%%"),
                },
            )

        with bc2:
            st.subheader("By Sector")
            _sec_rows = [
                {
                    "Sector": sec,
                    "Margin ($)": m,
                    "Share": m / _stats["total_margin"] if _stats["total_margin"] else 0.0,
                }
                for sec, m in sorted(
                    _stats["sector_margin"].items(),
                    key=lambda x: -x[1],
                )
            ]
            st.dataframe(
                pd.DataFrame(_sec_rows),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Margin ($)": st.column_config.NumberColumn(format="$%.0f"),
                    "Share":      st.column_config.NumberColumn(format="%.1f%%"),
                },
            )

    # ── Apply ──────────────────────────────────────────────────────────────
    with tab_apply:
        st.caption(
            "Apply the suggested portfolio to your strategy configuration. "
            "Strategies in the suggestion are set to **Live** with the suggested "
            "contract count. All other Live strategies are set to **Pass**. "
            "Review the diff below before applying."
        )

        _current  = load_strategies()
        _live_status = config.portfolio.live_status
        _suggested_names = {c["name"]: _result.contracts.get(c["name"], 1) for c in _result.candidates}
        _current_live    = {s["name"] for s in _current if s.get("status") == _live_status}

        _will_add    = [n for n in _suggested_names if n not in _current_live]
        _will_remove = [n for n in _current_live if n not in _suggested_names]
        _will_change = []
        for s in _current:
            n = s.get("name", "")
            if n in _suggested_names and n in _current_live:
                old_c = int(s.get("contracts") or 1)
                new_c = _suggested_names[n]
                if abs(old_c - new_c) >= 0.05:
                    _will_change.append({"name": n, "old": old_c, "new": new_c})

        dc1, dc2, dc3 = st.columns(3)
        dc1.metric("Strategies to add", len(_will_add))
        dc2.metric("Strategies to remove", len(_will_remove))
        dc3.metric("Contract changes", len(_will_change))

        if _will_add:
            with st.expander(f"✅ Add to Live ({len(_will_add)})", expanded=len(_will_add) <= 15):
                for n in _will_add:
                    st.write(f"- **{n}** — {_suggested_names[n]:.1f} contracts")

        if _will_remove:
            with st.expander(f"❌ Remove from Live ({len(_will_remove)})", expanded=len(_will_remove) <= 15):
                for n in _will_remove:
                    st.write(f"- **{n}**")

        if _will_change:
            with st.expander(f"🔄 Change contracts ({len(_will_change)})", expanded=len(_will_change) <= 15):
                for ch in _will_change:
                    arrow = "▲" if ch["new"] > ch["old"] else "▼"
                    st.write(f"- **{ch['name']}** {ch['old']} → **{ch['new']:.1f}** {arrow}")

        st.divider()

        if not _result.candidates:
            st.warning("No strategies in the suggestion — nothing to apply.")
        else:
            if st.button(
                f"✅ Apply Suggested Portfolio ({len(_result.candidates)} strategies)",
                type="primary",
                key="opt_apply_btn",
            ):
                updated = []
                for s in _current:
                    nm = s.get("name", "")
                    s = dict(s)
                    if nm in _suggested_names:
                        s["status"]    = _live_status
                        s["contracts"] = _suggested_names[nm]
                    elif s.get("status") == _live_status:
                        s["status"] = "Pass"
                    updated.append(s)
                save_strategies(updated)
                st.session_state.portfolio_data = None
                live_n = sum(1 for s in updated if s.get("status") == _live_status)
                st.success(
                    f"Applied! {live_n} Live strategies saved. "
                    "Rebuild the portfolio to see updated analytics."
                )
                st.page_link("ui/pages/03_Portfolio.py", label="→ Rebuild Portfolio")
