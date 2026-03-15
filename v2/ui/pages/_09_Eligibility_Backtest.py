"""
Eligibility Backtest page — 2 tabs mirroring VBA U_BackTest_Eligibility.

Tab 1 — Rule Statistics
  Walk-forward: for each month, evaluate all 70 rules and record
  N / Win% / $/Month / vs-Baseline across horizons 1–12.
  Displays as a heatmap-coloured table.

Tab 2 — Portfolio Construction Backtest
  Select rule(s) + settings → walk-forward equity curve vs baseline.
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.analytics.eligibility.portfolio_backtest import (
    PortfolioBacktestConfig,
    PortfolioBacktestResult,
    run_portfolio_backtest,
)
from core.analytics.eligibility.rule_backtest import run_rule_backtest
from core.analytics.eligibility.rules import build_rule_catalogue, evaluate_rule
from core.config import AppConfig, EligibilityConfig
from core.data_types import PortfolioData
from ui.components.settings_sidebar import render_ranking_sidebar, render_contract_sizing_sidebar

st.set_page_config(page_title="Eligibility Backtest", layout="wide")
st.title("Eligibility Backtest")

config: AppConfig = st.session_state.get("config", AppConfig.load())
portfolio: PortfolioData | None = st.session_state.get("portfolio_data")
imported = st.session_state.get("imported_data")

if imported is None:
    st.info("No data loaded yet.")
    st.page_link("ui/pages/01_Import.py", label="Go to Import →")
    st.stop()

if portfolio is None:
    st.info("Portfolio not built yet.")
    st.page_link("ui/pages/03_Portfolio.py", label="Go to Portfolio →")
    st.stop()

# ── Shared: build summary with status column ──────────────────────────────────
@st.cache_data(show_spinner=False)
def _build_summary_with_status(_portfolio_hash, _imported_hash):
    """Merge portfolio status, contracts and sector into summary_metrics."""
    summary = portfolio.summary_metrics.copy() if not portfolio.summary_metrics.empty else pd.DataFrame()
    if summary.empty:
        return summary

    status_map    = {s.name: s.status    for s in portfolio.strategies}
    contracts_map = {s.name: s.contracts for s in portfolio.strategies}
    sector_map    = {s.name: s.sector    for s in portfolio.strategies}

    for name in summary.index:
        if name not in status_map:
            status_map[name] = ""
    summary["status"]    = [status_map.get(n, "")    for n in summary.index]
    summary["contracts"] = [contracts_map.get(n, 1)  for n in summary.index]
    summary["sector"]    = [sector_map.get(n, "")    for n in summary.index]
    return summary


# Use id() as a cheap hash for the cache key
summary = _build_summary_with_status(id(portfolio), id(imported))

# Fall back if summary is empty or missing required columns
if summary.empty or "status" not in summary.columns:
    st.warning(
        "Summary metrics not available. Run the **Portfolio** page first to build strategy metrics."
    )
    st.stop()

rules = build_rule_catalogue()
rule_map = {r.id: r for r in rules}
rule_labels = [r.label for r in rules]

# ── Eligibility config sidebar ────────────────────────────────────────────────
with st.sidebar:
    st.header("Eligibility Settings")

    status_include = st.multiselect(
        "Eligible statuses",
        ["Live", "Paper", "Pass", "Retired"],
        default=list(config.eligibility.status_include),
    )
    days_threshold = st.number_input(
        "Min OOS days",
        min_value=0, max_value=730, value=int(config.eligibility.days_threshold_oos),
        help="Strategy must have ≥ N days of OOS data to be eligible",
    )
    dd_cap = st.number_input(
        "OOS DD / IS DD cap (0 = disabled)",
        min_value=0.0, max_value=10.0,
        value=float(config.eligibility.oos_dd_vs_is_cap), step=0.1,
        help="Exclude strategy if OOS max drawdown > cap × IS max drawdown",
    )
    eff_ratio = st.slider(
        "Efficiency ratio",
        0.0, 2.0, float(config.eligibility.efficiency_ratio), 0.05,
        help="Used by threshold rules: ann return ≥ eff_ratio × expected",
    )
    date_type = st.radio(
        "Eligibility date type",
        ["OOS Start Date", "Incubation Pass Date"],
        index=0 if config.eligibility.date_type == "OOS Start Date" else 1,
    )

    st.divider()
    st.subheader("Backtest Scope")

    data_scope = st.radio(
        "Data scope",
        ["OOS", "IS+OOS"],
        index=0 if config.eligibility.backtest_data_scope == "OOS" else 1,
        horizontal=True,
        help="OOS: use only out-of-sample data for P&L windows.\nIS+OOS: include in-sample history in rolling profit calculations.",
    )
    exclude_bh = st.checkbox(
        "Exclude Buy & Hold",
        value=config.eligibility.exclude_buy_and_hold,
        help="Buy & Hold strategies are excluded from eligibility scoring and backtest",
    )
    exclude_quit = st.checkbox(
        "Exclude previously quit",
        value=config.eligibility.exclude_previously_quit,
        help="Exclude strategies that have ever hit a quitting threshold (quitting_date is set)",
    )

    st.divider()
    if st.button("Save as defaults", use_container_width=True, help="Persist these settings so they load next session"):
        config.eligibility.days_threshold_oos    = int(days_threshold)
        config.eligibility.oos_dd_vs_is_cap      = float(dd_cap)
        config.eligibility.status_include        = status_include if status_include else ["Live"]
        config.eligibility.efficiency_ratio      = float(eff_ratio)
        config.eligibility.date_type             = date_type
        config.eligibility.backtest_data_scope   = data_scope
        config.eligibility.exclude_buy_and_hold  = exclude_bh
        config.eligibility.exclude_previously_quit = exclude_quit
        config.save()
        st.session_state.config = config
        st.success("Saved.")

    st.divider()
    if render_ranking_sidebar(config):
        st.rerun()
    if render_contract_sizing_sidebar(config):
        st.rerun()

elig_config = EligibilityConfig(
    days_threshold_oos=int(days_threshold),
    oos_dd_vs_is_cap=float(dd_cap),
    status_include=status_include if status_include else ["Live"],
    efficiency_ratio=float(eff_ratio),
    date_type=date_type,
    backtest_data_scope=data_scope,
    exclude_buy_and_hold=exclude_bh,
    exclude_previously_quit=exclude_quit,
)

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab1, tab2, tab3 = st.tabs(["📊 Rule Statistics", "📈 Portfolio Construction", "🔍 Strategy Screener"])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1: Rule Statistics
# ══════════════════════════════════════════════════════════════════════════════

with tab1:
    st.subheader("Walk-Forward Rule Statistics")
    st.caption(
        "For each calendar month, each rule is evaluated on eligible strategies. "
        "Forward returns are tracked for horizons 1–12 months. "
        "**vs Base** = % difference vs 'Baseline (All Eligible)'."
    )

    col_h, col_metric, col_run = st.columns([1, 1, 1])
    with col_h:
        horizon = st.selectbox("Horizon (months)", list(range(1, 13)), index=2)
    with col_metric:
        metric = st.selectbox(
            "Display metric",
            ["vs_base", "win_pct", "avg_pnl", "N"],
            format_func=lambda x: {
                "vs_base": "vs Baseline (%)",
                "win_pct": "Win % (profitable next period)",
                "avg_pnl": "Avg $/Strategy",
                "N":       "N (sample count)",
            }[x],
        )
    with col_run:
        st.write("")
        run_btn1 = st.button("Run Rule Backtest", type="primary", use_container_width=True)

    rule_bt_result = st.session_state.get("rule_bt_result")
    rule_bt_config_key = st.session_state.get("rule_bt_config_key")
    current_key = (
        tuple(sorted(status_include)), int(days_threshold),
        float(dd_cap), float(eff_ratio), date_type,
    )

    if run_btn1 or (rule_bt_result is None):
        with st.status(f"Evaluating {len(rules)} rules × 12 horizons…", expanded=True) as _rbt_status:
            st.write(f"Processing {len(portfolio.strategies)} strategies…")
            rule_bt_result = run_rule_backtest(
                daily_pnl=imported.daily_m2m,
                summary=summary,
                config=elig_config,
                rules=rules,
                max_horizon=12,
            )
            _rbt_status.update(
                label=f"Done — {len(rules)} rules × 12 horizons",
                state="complete",
                expanded=False,
            )
        st.session_state.rule_bt_result = rule_bt_result
        st.session_state.rule_bt_config_key = current_key

    if rule_bt_result is not None and not rule_bt_result.empty:
        col_name = f"{metric}_{horizon}"
        if col_name not in rule_bt_result.columns:
            st.error(f"Column '{col_name}' not found in results.")
        else:
            display = rule_bt_result[["label", f"N_{horizon}", f"win_pct_{horizon}",
                                      f"avg_pnl_{horizon}", f"vs_base_{horizon}"]].copy()
            display.columns = ["Rule", f"N", "Win %", "Avg $/Strat", "vs Base %"]
            display = display.sort_values("vs Base %", ascending=False).reset_index(drop=True)

            # Colour-code vs_base column
            def _style_vs_base(val):
                try:
                    v = float(val)
                    if v >= 20:    return "background-color: #1b5e20; color: white"
                    if v >= 10:    return "background-color: #4caf50; color: white"
                    if v >= 0:     return "background-color: #c8e6c9"
                    if v >= -10:   return "background-color: #ffccbc"
                    return "background-color: #b71c1c; color: white"
                except Exception:
                    return ""

            styled = display.style.applymap(_style_vs_base, subset=["vs Base %"])
            st.dataframe(styled, hide_index=True, use_container_width=True, height=600)

            # Top-10 bar chart
            top10 = display.head(10)
            fig = px.bar(
                top10, x="vs Base %", y="Rule", orientation="h",
                title=f"Top 10 Rules by vs Baseline% — Horizon {horizon}M",
                color="vs Base %",
                color_continuous_scale=["#F44336", "#FFEB3B", "#4CAF50"],
                color_continuous_midpoint=0,
            )
            fig.add_vline(x=0, line_color="black", line_width=1)
            fig.update_layout(
                height=350,
                coloraxis_showscale=False,
                yaxis={"categoryorder": "total ascending"},
            )
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Click **Run Rule Backtest** to compute statistics.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2: Portfolio Construction Backtest
# ══════════════════════════════════════════════════════════════════════════════

with tab2:
    st.subheader("Portfolio Construction Backtest")
    st.caption(
        "At each month, select strategies that pass the chosen rule(s). "
        "Track the resulting portfolio's equity curve vs the all-eligible baseline."
    )

    col_rules, col_opts = st.columns([2, 1])

    with col_rules:
        selected_rule_labels = st.multiselect(
            "Rules to test (multi-select)",
            rule_labels,
            default=[rule_labels[1]] if len(rule_labels) > 1 else rule_labels[:1],
            help="Each selected rule generates a separate equity curve",
        )
        selected_rule_ids = [r.id for r in rules if r.label in selected_rule_labels]

    with col_opts:
        max_strats = st.number_input(
            "Max strategies (0 = all passing)",
            min_value=0, max_value=100, value=0,
        )
        ranking_metric = st.selectbox(
            "Ranking metric (for top-N cap)",
            ["oos_pnl", "momentum_3m", "momentum_6m", "expected_return"],
            format_func=lambda x: {
                "oos_pnl":          "OOS Total PnL",
                "momentum_3m":      "Momentum 3M",
                "momentum_6m":      "Momentum 6M",
                "expected_return":  "Expected Annual Return",
            }[x],
        )
        weighting = st.radio("Weighting", ["equal", "by_contracts"], horizontal=True)

    run_btn2 = st.button("Run Portfolio Backtest", type="primary")

    pb_results: dict[str, PortfolioBacktestResult] | None = st.session_state.get("pb_results")

    if run_btn2:
        if not selected_rule_ids:
            st.error("Select at least one rule.")
        else:
            bt_config = PortfolioBacktestConfig(
                rule_ids=selected_rule_ids,
                max_strategies=int(max_strats) if max_strats > 0 else None,
                ranking_metric=ranking_metric,
                weighting=weighting,
                include_baseline=True,
            )
            with st.spinner("Running portfolio construction backtest…"):
                pb_results = run_portfolio_backtest(
                    daily_pnl=imported.daily_m2m,
                    summary=summary,
                    config=elig_config,
                    backtest_config=bt_config,
                    rules=rules,
                )
            st.session_state.pb_results = pb_results

    if pb_results:
        # ── Equity curve comparison ───────────────────────────────────────────
        st.subheader("Equity Curves")
        palette = px.colors.qualitative.Plotly

        fig_eq = go.Figure()
        for i, (lbl, res) in enumerate(pb_results.items()):
            is_baseline = lbl.startswith("Baseline")
            fig_eq.add_trace(go.Scatter(
                x=res.equity_curve.index,
                y=res.equity_curve.values,
                name=lbl,
                line=dict(
                    width=3 if is_baseline else 1.5,
                    dash="dash" if is_baseline else "solid",
                    color="#546E7A" if is_baseline else palette[i % len(palette)],
                ),
            ))
        fig_eq.update_layout(
            height=420,
            xaxis_title="Date",
            yaxis_title="Cumulative PnL ($)",
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig_eq, use_container_width=True)

        # ── Summary metrics table ─────────────────────────────────────────────
        st.subheader("Performance Summary")
        rows = []
        for lbl, res in pb_results.items():
            rows.append({
                "Rule":            lbl,
                "Avg Monthly $":   f"${res.avg_monthly_pnl:,.0f}",
                "Win Rate":        f"{res.win_rate:.1%}",
                "Max Drawdown $":  f"${res.max_drawdown:,.0f}",
                "Sharpe":          f"{res.sharpe_ratio:.2f}",
                "vs Baseline %":   f"{res.vs_baseline_pct:+.1f}%",
            })
        summary_df = pd.DataFrame(rows)
        st.dataframe(summary_df, hide_index=True, use_container_width=True)

        # ── Monthly strategy count ────────────────────────────────────────────
        with st.expander("Monthly Strategy Count", expanded=False):
            fig_cnt = go.Figure()
            for i, (lbl, res) in enumerate(pb_results.items()):
                if lbl.startswith("Baseline"):
                    continue
                fig_cnt.add_trace(go.Scatter(
                    x=res.monthly_strategy_count.index,
                    y=res.monthly_strategy_count.values,
                    name=lbl,
                    mode="lines+markers",
                    line=dict(color=palette[i % len(palette)]),
                ))
            fig_cnt.update_layout(
                height=280,
                xaxis_title="Date",
                yaxis_title="# Strategies Selected",
                hovermode="x unified",
            )
            st.plotly_chart(fig_cnt, use_container_width=True)

    elif not run_btn2:
        st.info("Select rule(s) and click **Run Portfolio Backtest**.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3: Strategy Screener — "as of today"
# ══════════════════════════════════════════════════════════════════════════════

# ── Ranking-metric label map (shared between Tab 3 controls and display) ──────
_RANK_METRIC_LABELS: dict[str, str] = {
    "rtd_oos":                "RTD (OOS)",
    "rtd_12_months":          "RTD (12M)",
    "sharpe_isoos":           "Sharpe IS+OOS",
    "profit_since_oos_start": "OOS Total ($)",
    "profit_last_12_months":  "Last 12M ($)",
    "k_factor":               "K-Factor",
    "ulcer_index":            "Ulcer Index",
    "contracts":              "Contracts",
}
_rk_default = config.ranking

with tab3:
    st.subheader("Strategy Screener")
    st.caption(
        "Evaluate any combination of rules against every strategy using all data available today. "
        "See which strategies currently pass your chosen criteria — independently of the walk-forward backtest."
    )

    # ── Ranking & filter controls ─────────────────────────────────────────────
    with st.expander("Ranking & Display Options", expanded=True):
        col_rk1, col_rk2, col_rk3 = st.columns([2, 1, 1])
        with col_rk1:
            rank_metric = st.selectbox(
                "Rank by",
                list(_RANK_METRIC_LABELS.keys()),
                index=list(_RANK_METRIC_LABELS.keys()).index(_rk_default.metric),
                format_func=lambda k: _RANK_METRIC_LABELS[k],
                help="Metric used to rank strategies. Highest value ranked first unless 'ascending' is checked.",
            )
        with col_rk2:
            rank_ascending = st.checkbox(
                "Ascending (lower = better)",
                value=_rk_default.ascending,
                help="Enable for Ulcer Index and similar 'lower is better' metrics.",
            )
            rank_eligible_only = st.checkbox(
                "Eligible strategies only",
                value=_rk_default.eligible_only,
                help="Hide strategies that fail the base eligibility checks.",
            )
        with col_rk3:
            rank_group_sector = st.checkbox(
                "Group by sector",
                value=_rk_default.group_by_sector,
                help="Sort and group the ranking table by sector.",
            )
            rank_group_contracts = st.checkbox(
                "Sub-sort by contracts",
                value=_rk_default.group_by_contracts,
                help="Within each sector, break ties by contract count (descending).",
            )

    # ── Rule selector ──────────────────────────────────────────────────────────
    col_sc1, col_sc2 = st.columns([3, 1])
    with col_sc1:
        default_screener_labels = [r.label for r in rules if r.id in (1, 6, 10)]
        screener_rule_labels = st.multiselect(
            "Rules to apply",
            rule_labels,
            default=default_screener_labels,
            help="Strategies must pass ALL selected rules to appear in the 'passing' section.",
        )
    with col_sc2:
        st.write("")
        show_all = st.checkbox(
            "Show non-eligible strategies too",
            value=False,
            disabled=rank_eligible_only,
        )
        only_passing = st.checkbox("Show only strategies passing all rules", value=False)

    screener_rules = [r for r in rules if r.label in screener_rule_labels]

    # ── Helper: last-N monthly sum ─────────────────────────────────────────────
    def _ln(arr: np.ndarray, n: int) -> float:
        if n <= 0 or len(arr) == 0:
            return 0.0
        return float(arr[-n:].sum())

    # ── Build monthly series (full history) ────────────────────────────────────
    monthly_all = imported.daily_m2m.resample("ME").sum()
    months_index = monthly_all.index
    now_ts = pd.Timestamp.now()

    # ── Evaluate each strategy ─────────────────────────────────────────────────
    rows = []
    bh_rows = []      # separate list for B&H strategies

    for name in summary.index:
        row = summary.loc[name]
        status  = str(row.get("status", ""))
        sector  = str(row.get("sector", "") or "")

        # --- Buy & Hold detection ---
        _is_bh = "buy" in status.lower() and "hold" in status.lower()

        # --- Base eligibility ---
        status_ok = status in (elig_config.status_include or ["Live"])

        oos_begin_raw = row.get("oos_begin")
        incub_raw     = row.get("incubation_date")
        date_raw = (
            incub_raw if elig_config.date_type == "Incubation Pass Date" else oos_begin_raw
        )
        if pd.notna(date_raw) if not isinstance(date_raw, type(None)) else False:
            try:
                oos_days = int((now_ts - pd.Timestamp(date_raw)).days)
            except Exception:
                oos_days = 0
        else:
            oos_days = 0

        days_ok = oos_days >= elig_config.days_threshold_oos

        dd_cap_ok = True
        if elig_config.oos_dd_vs_is_cap > 0:
            is_dd  = abs(float(row.get("max_drawdown_is", 0) or 0))
            oos_dd = abs(float(row.get("max_oos_drawdown", 0) or 0))
            dd_cap_ok = (is_dd == 0) or (oos_dd <= elig_config.oos_dd_vs_is_cap * is_dd)

        base_eligible = status_ok and days_ok and dd_cap_ok

        _exclude = not base_eligible and (rank_eligible_only or not show_all)
        if _exclude and not _is_bh:
            continue

        # --- Get monthly PnL array for this strategy ---
        if name not in monthly_all.columns:
            continue
        m_pnl = monthly_all[name].values.astype(float)

        # --- OOS start index ---
        oos_begin_ts = None
        try:
            oos_begin_ts = pd.Timestamp(oos_begin_raw) if pd.notna(oos_begin_raw) else None
        except Exception:
            pass
        if oos_begin_ts is not None:
            oos_idx = int(min(
                months_index.searchsorted(oos_begin_ts, side="left"),
                len(months_index) - 1,
            ))
        else:
            oos_idx = 0

        # --- Expected annual & contracts ---
        exp_annual = float(row.get("expected_annual_profit", 0) or 0)
        contracts  = int(row.get("contracts", 1) or 1)

        # --- Compute PnL display metrics ---
        last_1m  = _ln(m_pnl, 1)  * contracts
        last_3m  = _ln(m_pnl, 3)  * contracts
        last_6m  = _ln(m_pnl, 6)  * contracts
        last_12m = _ln(m_pnl, 12) * contracts
        oos_total = float(m_pnl[oos_idx:].sum()) * contracts if oos_idx < len(m_pnl) else 0.0

        # --- Ranking metric value (from precomputed summary) ---
        _rank_val = None
        if rank_metric in ("profit_since_oos_start", "profit_last_12_months"):
            # Use contract-scaled live values
            _rank_val = oos_total if rank_metric == "profit_since_oos_start" else last_12m
        else:
            raw_val = row.get(rank_metric)
            _rank_val = float(raw_val) if raw_val is not None and pd.notna(raw_val) else None

        # --- B&H: capture separately, skip from main ranking ---
        if _is_bh:
            bh_rows.append({
                "Strategy":     name,
                "Symbol":       str(row.get("symbol", "") or ""),
                "Sector":       sector,
                "Contracts":    contracts,
                "OOS Days":     oos_days,
                "Last 1M ($)":  last_1m,
                "Last 3M ($)":  last_3m,
                "Last 12M ($)": last_12m,
                "OOS Total ($)": oos_total,
            })
            continue

        # --- Evaluate selected rules ---
        rule_results: dict[str, bool] = {}
        for sr in screener_rules:
            rule_results[sr.label] = evaluate_rule(
                sr, m_pnl, oos_idx, exp_annual, elig_config.efficiency_ratio
            )

        n_passed = sum(rule_results.values())
        passes_all = (n_passed == len(screener_rules)) if screener_rules else True

        if only_passing and not passes_all:
            continue

        entry: dict = {
            "Strategy":      name,
            "Sector":        sector,
            "Status":        status,
            "Eligible":      "✓" if base_eligible else "✗",
            "OOS Days":      oos_days,
            "Contracts":     contracts,
            _RANK_METRIC_LABELS[rank_metric]: _rank_val,
            "Last 1M ($)":   last_1m,
            "Last 3M ($)":   last_3m,
            "Last 6M ($)":   last_6m,
            "Last 12M ($)":  last_12m,
            "OOS Total ($)": oos_total,
        }
        for sr in screener_rules:
            entry[sr.label] = rule_results[sr.label]
        if len(screener_rules) > 1:
            entry["Rules Passed"] = f"{n_passed}/{len(screener_rules)}"
        entry["_passes_all"]  = passes_all
        entry["_rank_val"]    = _rank_val if _rank_val is not None else (float("inf") if rank_ascending else float("-inf"))
        rows.append(entry)

    if not rows:
        st.info("No strategies to display with the current eligibility settings.")
    else:
        df_screen = pd.DataFrame(rows)
        passes_all_col = df_screen.pop("_passes_all")
        rank_sort_col  = df_screen.pop("_rank_val")

        # ── Sort: passing-all first, then by ranking metric ───────────────────
        df_screen["_pass_sort"] = passes_all_col.astype(int)
        sort_cols = ["_pass_sort"]
        sort_asc  = [False]
        if rank_group_sector:
            sort_cols.insert(0, "Sector")
            sort_asc.insert(0, True)
        sort_cols.append("_rank_col")
        sort_asc.append(rank_ascending)
        if rank_group_contracts:
            sort_cols.append("Contracts")
            sort_asc.append(False)

        df_screen["_rank_col"] = rank_sort_col.values
        df_screen = df_screen.sort_values(sort_cols, ascending=sort_asc)
        df_screen = df_screen.drop(columns=["_pass_sort", "_rank_col"]).reset_index(drop=True)

        # ── Summary banner ─────────────────────────────────────────────────────
        n_elig = (df_screen["Eligible"] == "✓").sum()
        n_pass = int(passes_all_col.reindex(df_screen.index, fill_value=False).sum()) if screener_rules else n_elig
        rule_summary = (
            f"**{n_pass}** of **{n_elig}** eligible strategies pass all selected rules."
            if screener_rules
            else f"**{n_elig}** eligible strategies (no rules selected)."
        )
        st.info(rule_summary)

        # ── Colour rule columns ────────────────────────────────────────────────
        rule_col_names = [sr.label for sr in screener_rules]
        pnl_col_names  = ["Last 1M ($)", "Last 3M ($)", "Last 6M ($)", "Last 12M ($)", "OOS Total ($)"]

        def _style_screener(df: pd.DataFrame):
            styles = pd.DataFrame("", index=df.index, columns=df.columns)
            for col in rule_col_names:
                if col not in df.columns:
                    continue
                for idx, val in df[col].items():
                    styles.at[idx, col] = (
                        "background-color: #c8e6c9" if val is True else "background-color: #ffcdd2"
                    )
            for col in pnl_col_names:
                if col not in df.columns:
                    continue
                for idx, val in df[col].items():
                    try:
                        styles.at[idx, col] = (
                            "color: #2e7d32" if float(val) >= 0 else "color: #c62828"
                        )
                    except Exception:
                        pass
            return styles

        display_df = df_screen.copy()
        for col in rule_col_names:
            display_df[col] = display_df[col].map({True: "✓", False: "✗"})
        for col in pnl_col_names:
            display_df[col] = display_df[col].apply(
                lambda v: f"${v:,.0f}" if isinstance(v, (int, float)) and not np.isnan(v) else v
            )

        styled_screen = display_df.style.apply(_style_screener, axis=None)
        st.dataframe(styled_screen, hide_index=True, use_container_width=True, height=550)

        # ── OOS Total bar chart (coloured by rank metric, grouped by sector) ──
        if not df_screen.empty:
            chart_df = df_screen[["Strategy", "Sector", "OOS Total ($)"]].copy()
            chart_df["pass"] = passes_all_col.reindex(df_screen.index, fill_value=False).values
            chart_df["color"] = chart_df["pass"].map({True: "Passes All Rules", False: "Fails a Rule"})

            _color_col = "Sector" if rank_group_sector else "color"
            fig_sc = px.bar(
                chart_df.sort_values("OOS Total ($)", ascending=True),
                x="OOS Total ($)",
                y="Strategy",
                orientation="h",
                color=_color_col,
                title=f"OOS Total PnL per Strategy — ranked by {_RANK_METRIC_LABELS[rank_metric]}",
                **({"color_discrete_map": {"Passes All Rules": "#4CAF50", "Fails a Rule": "#EF9A9A"}}
                   if not rank_group_sector else {}),
            )
            fig_sc.add_vline(x=0, line_color="black", line_width=1)
            fig_sc.update_layout(
                height=max(300, len(chart_df) * 28 + 100),
                coloraxis_showscale=False,
                legend_title_text="Sector" if rank_group_sector else "",
            )
            st.plotly_chart(fig_sc, use_container_width=True)

    # ── Buy & Hold panel ──────────────────────────────────────────────────────
    with st.expander(f"Buy & Hold Strategies ({len(bh_rows)} found)", expanded=bool(bh_rows)):
        if not bh_rows:
            st.caption(
                "No Buy & Hold strategies are currently loaded. "
                "Add B&H strategies in the Portfolio page and set status to 'Buy & Hold'."
            )
        else:
            st.caption(
                "B&H strategies are used as benchmarks only — excluded from eligibility "
                "scoring and rule evaluation."
            )
            df_bh = pd.DataFrame(bh_rows)
            pnl_bh_cols = ["Last 1M ($)", "Last 3M ($)", "Last 12M ($)", "OOS Total ($)"]

            def _style_bh(df: pd.DataFrame):
                styles = pd.DataFrame("", index=df.index, columns=df.columns)
                for col in pnl_bh_cols:
                    if col not in df.columns:
                        continue
                    for idx, val in df[col].items():
                        try:
                            styles.at[idx, col] = (
                                "color: #2e7d32" if float(val) >= 0 else "color: #c62828"
                            )
                        except Exception:
                            pass
                return styles

            display_bh = df_bh.copy()
            for col in pnl_bh_cols:
                display_bh[col] = display_bh[col].apply(
                    lambda v: f"${v:,.0f}" if isinstance(v, (int, float)) and not np.isnan(v) else v
                )
            st.dataframe(
                display_bh.style.apply(_style_bh, axis=None),
                hide_index=True, use_container_width=True,
            )
