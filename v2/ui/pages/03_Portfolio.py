"""
Portfolio page — aggregated metrics for all Live strategies.
Mirrors the VBA Portfolio tab: filtered to Live status, scaled by contracts.

Sections:
1. Portfolio header stats (equity, drawdown, win rate)
2. Equity curve chart
3. Per-strategy portfolio metrics table
4. Monthly PnL heatmap
"""

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import date

from core.config import AppConfig
from core.portfolio.strategies import load_strategies, save_strategies
from core.portfolio.aggregator import (
    build_portfolio,
    portfolio_total_pnl,
    portfolio_equity_curve,
    monthly_portfolio_pnl,
    portfolio_summary_stats,
)
from core.portfolio.summary import apply_eligibility_rules, compute_summary
from ui.strategy_labels import render_strategy_picker, build_label_map, render_legend

st.set_page_config(page_title="Portfolio", layout="wide")

# ── Sidebar: workflow status + portfolio / contract sizing settings ────────────
try:
    from ui.workflow import render_workflow_sidebar
    with st.sidebar:
        render_workflow_sidebar()
except Exception:
    pass

try:
    from ui.components.settings_sidebar import render_portfolio_sidebar, render_contract_sizing_sidebar
    _sidebar_cfg = st.session_state.get("config", AppConfig.load())
    with st.sidebar:
        st.divider()
        if render_portfolio_sidebar(_sidebar_cfg):
            st.rerun()
        if render_contract_sizing_sidebar(_sidebar_cfg):
            st.rerun()
except Exception:
    pass

st.title("Portfolio")

_pnav_l, _ = st.columns([1, 7])
with _pnav_l:
    st.page_link("ui/pages/02_Strategies.py", label="← Strategies")

st.caption("Step 4 of 4 — build the portfolio to aggregate all Live strategies.")

config: AppConfig = st.session_state.get("config", AppConfig.load())
imported = st.session_state.get("imported_data")
strategies_config = load_strategies()

# ── Guard: require imported data ──────────────────────────────────────────────

if imported is None:
    st.info("No data loaded yet.")
    st.page_link("ui/pages/01_Import.py", label="Go to Import →")
    st.stop()

if not strategies_config:
    st.info("No strategies configured yet.")
    st.page_link("ui/pages/02_Strategies.py", label="Go to Strategies →")
    st.stop()

# ── Check for cached portfolio ────────────────────────────────────────────────

portfolio = st.session_state.get("portfolio_data")
_needs_rebuild = portfolio is None

if _needs_rebuild:
    st.warning(
        "Portfolio has not been built yet, or strategy changes require a rebuild. "
        "Click **Build Portfolio** to continue."
    )

col_rebuild, col_snap, col_status = st.columns([1, 1, 4])
with col_rebuild:
    if st.button("Build Portfolio" if _needs_rebuild else "Rebuild Portfolio",
                 type="primary" if _needs_rebuild else "secondary"):
        _needs_rebuild = True
with col_snap:
    _port_snap_btn = st.button(
        "📸 Set Live Portfolio",
        help="Save the current Live portfolio as a snapshot and show trading instructions vs the previous baseline.",
        key="port_set_live_btn",
    )

if _needs_rebuild:
    with st.spinner("Computing portfolio metrics..."):
        # Build strategy folders lookup for WF CSV paths
        from core.ingestion.folder_scanner import scan_folders
        scan_result = scan_folders(config.folders) if config.folders else None
        sf_list = scan_result.strategies if scan_result else []

        # Compute per-strategy summary metrics
        cutoff = None
        if config.portfolio.use_cutoff and config.portfolio.cutoff_date:
            try:
                cutoff = date.fromisoformat(config.portfolio.cutoff_date)
            except ValueError:
                cutoff = None

        summary_df = compute_summary(
            imported=imported,
            strategy_folders=sf_list,
            date_format=config.date_format,
            use_cutoff=config.portfolio.use_cutoff,
            cutoff_date=cutoff,
            incubation_months=config.incubation.months,
            min_incubation_ratio=config.incubation.min_profit_ratio,
            eligibility_months=config.eligibility.eligibility_months,
            quitting_method=config.quitting.method,
            quitting_max_dollars=config.quitting.max_dollars,
            quitting_max_percent=config.quitting.max_percent_drawdown,
            quitting_sd_multiple=config.quitting.sd_multiple,
        )

        # Apply eligibility rules and add as a column
        if not summary_df.empty:
            elig = apply_eligibility_rules(summary_df, config.eligibility)
            summary_df["eligibility_status"] = elig.map({True: "Yes", False: "No"})

        portfolio = build_portfolio(
            imported=imported,
            strategies_config=strategies_config,
            summary_metrics=summary_df,
            live_status=config.portfolio.live_status,
        )
        st.session_state.portfolio_data = portfolio

with col_status:
    if portfolio and portfolio.strategies:
        st.caption(
            f"Portfolio: **{len(portfolio.strategies)}** live strategies active"
        )

# ── Set Live Portfolio handler ─────────────────────────────────────────────────
if _port_snap_btn:
    from core.portfolio.snapshot import (
        compare_portfolios as _port_cmp,
        list_snapshots as _port_list_snaps,
        load_snapshot as _port_load_snap,
        save_snapshot as _port_save_snap,
    )
    from datetime import datetime as _port_dt
    _port_all = load_strategies()
    _port_snaps = _port_list_snaps()
    _port_prev = _port_load_snap(_port_snaps[0]["filename"]) if _port_snaps else []
    _port_result = _port_cmp(_port_all, _port_prev, live_status=config.portfolio.live_status)
    _port_label = _port_dt.now().strftime("%Y-%m-%d %H:%M")
    _port_save_snap(_port_all, _port_label)
    _port_live_n = sum(1 for s in _port_all if s.get("status") == config.portfolio.live_status)
    st.success(f"📸 Live portfolio saved — {_port_live_n} strategies ({_port_label})")

    if _port_result.has_changes:
        st.subheader("Trading Instructions vs previous baseline")
        if _port_result.new_strategies:
            st.markdown("#### ✅ Enable in trading system")
            for _ns in _port_result.new_strategies:
                _c = _ns.get("contracts", 1)
                st.markdown(f"- **{_ns['name']}**  ({_ns.get('symbol','')}, {_c} contract{'s' if _c != 1 else ''})")
        if _port_result.removed_strategies:
            st.markdown("#### ❌ Disable in trading system")
            for _rs in _port_result.removed_strategies:
                st.markdown(f"- **{_rs['name']}**  ({_rs.get('symbol','')})")
        if _port_result.contract_changes:
            st.markdown("#### 🔄 Adjust contracts")
            for _chg in _port_result.contract_changes:
                _arrow = "▲" if _chg["delta"] > 0 else "▼"
                st.markdown(
                    f"- **{_chg['name']}** ({_chg['symbol']})  "
                    f"{_chg['old_contracts']} → **{_chg['new_contracts']}** {_arrow}"
                )
    elif _port_snaps:
        st.info("No changes vs previous baseline.")
    else:
        st.info("First baseline saved.")

# ── Quick portfolio add/remove ─────────────────────────────────────────────────
_live_status = config.portfolio.live_status
_all_strats_port = load_strategies()
_live_names_port = [s["name"] for s in _all_strats_port if s.get("status") == _live_status]
_non_live_names_port = [s["name"] for s in _all_strats_port if s.get("status") != _live_status]

with st.expander("➕ / ➖ Modify Portfolio Composition", expanded=False):
    _pc1, _pc2 = st.columns(2)
    with _pc1:
        st.markdown("**Add to portfolio**")
        _port_to_add = st.multiselect(
            "Add strategies as Live",
            options=_non_live_names_port,
            key="port_quick_add",
            placeholder="Search…",
        )
        if st.button("➕ Add selected", key="port_add_btn", disabled=not _port_to_add):
            _updated = [
                dict(s, status=_live_status) if s.get("name") in _port_to_add else s
                for s in _all_strats_port
            ]
            save_strategies(_updated)
            st.session_state.portfolio_data = None
            st.success(f"Added {len(_port_to_add)} strategies. Rebuild portfolio to apply.")
            st.rerun()
    with _pc2:
        st.markdown("**Remove from portfolio**")
        _port_to_remove = st.multiselect(
            "Remove Live strategies",
            options=_live_names_port,
            key="port_quick_remove",
            placeholder="Search…",
        )
        if st.button("➖ Remove selected", key="port_remove_btn", disabled=not _port_to_remove):
            _updated = [
                dict(s, status="Pass") if s.get("name") in _port_to_remove else s
                for s in _all_strats_port
            ]
            save_strategies(_updated)
            st.session_state.portfolio_data = None
            st.success(f"Removed {len(_port_to_remove)} strategies (set to Pass). Rebuild portfolio to apply.")
            st.rerun()

if portfolio and portfolio.strategies:
    with st.sidebar:
        st.divider()
        render_strategy_picker(portfolio.strategies, key="port_strat_picker")

if portfolio is None or not portfolio.strategies:
    st.warning("No live strategies in portfolio.")
    st.page_link("ui/pages/02_Strategies.py", label="Go to Strategies to mark strategies as Live →")
    st.stop()


# ── 1. Header metrics ──────────────────────────────────────────────────────────

stats = portfolio_summary_stats(portfolio)

st.subheader("Portfolio Summary")

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Live Strategies", stats.get("n_strategies", 0))
c2.metric("Total P&L", f"${stats.get('total_profit', 0):,.0f}")
c3.metric("Annualised P&L", f"${stats.get('annualised_profit', 0):,.0f}")
c4.metric("Expected Annual", f"${stats.get('expected_annual_profit', 0):,.0f}")
c5.metric("Max Drawdown", f"${stats.get('max_drawdown', 0):,.0f}")
c6.metric(
    "Monthly Win Rate",
    f"{stats.get('monthly_win_rate', 0) * 100:.0f}%",
)

st.divider()


# ── 2. Equity curve ────────────────────────────────────────────────────────────

st.subheader("Portfolio Equity Curve")

equity = portfolio_equity_curve(portfolio)
total_pnl = portfolio_total_pnl(portfolio)

# Per-strategy equity curves
strat_equity = portfolio.daily_pnl.cumsum()

fig = go.Figure()

# Individual strategy lines (thin, semi-transparent)
for col in strat_equity.columns:
    fig.add_trace(go.Scatter(
        x=strat_equity.index,
        y=strat_equity[col],
        mode="lines",
        name=col,
        line=dict(width=1),
        opacity=0.4,
        showlegend=False,
    ))

# Total portfolio line (bold)
fig.add_trace(go.Scatter(
    x=equity.index,
    y=equity.values,
    mode="lines",
    name="Portfolio Total",
    line=dict(width=2.5, color="#1f77b4"),
))

fig.update_layout(
    xaxis_title="Date",
    yaxis_title="Cumulative P&L ($)",
    hovermode="x unified",
    height=400,
    margin=dict(l=0, r=0, t=10, b=0),
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
)
st.plotly_chart(fig, use_container_width=True)

_eq_exp_col, _ = st.columns([1, 5])
with _eq_exp_col:
    if st.button("Export Equity Curves to Excel", key="export_equity_btn"):
        from core.reporting.excel_export import export_equity_curves, equity_curves_export_filename
        _xlsx = export_equity_curves(portfolio.daily_pnl)
        st.download_button(
            "📥 Download Equity Curves",
            data=_xlsx,
            file_name=equity_curves_export_filename(),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_equity_xlsx",
        )

st.divider()


# ── 3. Per-strategy metrics table ─────────────────────────────────────────────

st.subheader("Strategy Metrics")

if not portfolio.summary_metrics.empty:
    sm = portfolio.summary_metrics.copy()

    # Add contracts from strategies config
    if "contracts" not in sm.columns:
        contracts_map = {s.name: s.contracts for s in portfolio.strategies}
        sm["contracts"] = sm.index.map(contracts_map)

    # All available metric columns with friendly labels
    _ALL_METRIC_COLS: dict[str, str] = {
        "eligibility_status": "Eligible",
        "contracts": "Contracts",
        "symbol": "Symbol",
        "sector": "Sector",
        "oos_begin": "OOS Start",
        "oos_end": "OOS End",
        "oos_period_years": "OOS Years",
        "expected_annual_profit": "Exp. Annual ($)",
        "actual_annual_profit": "Act. Annual ($)",
        "return_efficiency": "Efficiency",
        "trades_per_year": "Trades/Yr",
        "overall_win_rate": "Win Rate",
        "sharpe_isoos": "Sharpe IS+OOS",
        "sharpe_is": "Sharpe IS",
        "max_drawdown_isoos": "Max DD IS+OOS ($)",
        "max_drawdown_is": "Max DD IS ($)",
        "profit_last_1_month": "Last 1M ($)",
        "profit_last_3_months": "Last 3M ($)",
        "profit_last_6_months": "Last 6M ($)",
        "profit_last_9_months": "Last 9M ($)",
        "profit_last_12_months": "Last 12M ($)",
        "profit_since_oos_start": "OOS P&L ($)",
        "max_oos_drawdown": "OOS Max DD ($)",
        "avg_oos_drawdown": "Avg OOS DD ($)",
        "rtd_oos": "R:DD OOS",
        "rtd_12_months": "R:DD 12M",
        "count_profit_months": "Profit Months",
        "incubation_status": "Incubation",
        "incubation_date": "Incub. Date",
        "quitting_status": "Quit Status",
        "quitting_date": "Quit Date",
        "profit_since_quit": "P&L Since Quit ($)",
        "k_factor": "K-Factor",
        "ulcer_index": "Ulcer Index",
        "best_month": "Best Month ($)",
        "worst_month": "Worst Month ($)",
        "max_consecutive_loss_months": "Max Loss Streak",
    }
    _DEFAULT_METRIC_COLS = [
        "eligibility_status", "contracts", "symbol", "sector",
        "expected_annual_profit", "actual_annual_profit", "return_efficiency",
        "profit_last_12_months", "max_oos_drawdown", "rtd_oos",
    ]

    _avail_metric_cols = [c for c in _ALL_METRIC_COLS if c in sm.columns]
    _avail_metric_labels = {c: _ALL_METRIC_COLS[c] for c in _avail_metric_cols}
    _default_sel = [c for c in _DEFAULT_METRIC_COLS if c in _avail_metric_cols]

    # Initialise session state on first load
    _col_key = "port_metrics_col_picker"
    if _col_key not in st.session_state:
        st.session_state[_col_key] = _default_sel

    with st.expander("⚙ Columns", expanded=False):
        _sel_cols = st.multiselect(
            "Select columns to display",
            options=_avail_metric_cols,
            format_func=lambda c: _avail_metric_labels.get(c, c),
            key=_col_key,
        )
        if st.button("Reset to defaults", key="port_metrics_cols_reset"):
            st.session_state[_col_key] = _default_sel
            st.rerun()

    _raw_sel = st.session_state.get(_col_key) or _default_sel
    display_cols = [c for c in _raw_sel if c in sm.columns] or _default_sel

    if display_cols:
        display_df = sm[display_cols].reset_index()
        display_df.rename(columns={"strategy_name": "Strategy"}, inplace=True)

        _readonly = [c for c in display_df.columns if c != "contracts"]

        _edited_metrics = st.data_editor(
            display_df,
            use_container_width=True,
            hide_index=True,
            disabled=_readonly,
            key="portfolio_metrics_editor",
            column_config={
                "eligibility_status": st.column_config.TextColumn(
                    "Eligible",
                    help="Yes = meets all configured eligibility rules; No = one or more rules failed.",
                ),
                "contracts": st.column_config.NumberColumn(
                    "Contr.", format="%d", min_value=0, max_value=999, step=1
                ),
                "symbol": st.column_config.TextColumn("Symbol"),
                "sector": st.column_config.TextColumn("Sector"),
                "trades_per_year": st.column_config.NumberColumn("Trades/Yr", format="%.1f"),
                "incubation_status": st.column_config.TextColumn("Incubation"),
                "oos_begin": st.column_config.DateColumn("OOS Start"),
                "oos_end": st.column_config.DateColumn("OOS End"),
                "expected_annual_profit": st.column_config.NumberColumn(
                    "Exp. Annual ($)", format="$%.0f"
                ),
                "actual_annual_profit": st.column_config.NumberColumn(
                    "Act. Annual ($)", format="$%.0f"
                ),
                "return_efficiency": st.column_config.NumberColumn(
                    "Efficiency", format="%.1%%"
                ),
                "overall_win_rate": st.column_config.NumberColumn(
                    "Win Rate", format="%.1%%"
                ),
                "max_drawdown_isoos": st.column_config.NumberColumn(
                    "Max DD IS+OOS ($)", format="$%.0f"
                ),
                "sharpe_isoos": st.column_config.NumberColumn("Sharpe IS+OOS", format="%.2f"),
                "profit_last_1_month": st.column_config.NumberColumn(
                    "Last 1M ($)", format="$%.0f"
                ),
                "profit_last_3_months": st.column_config.NumberColumn(
                    "Last 3M ($)", format="$%.0f"
                ),
                "profit_last_6_months": st.column_config.NumberColumn(
                    "Last 6M ($)", format="$%.0f"
                ),
                "profit_last_12_months": st.column_config.NumberColumn(
                    "Last 12M ($)", format="$%.0f"
                ),
                "profit_since_oos_start": st.column_config.NumberColumn(
                    "OOS P&L ($)", format="$%.0f"
                ),
                "max_oos_drawdown": st.column_config.NumberColumn(
                    "OOS Max DD ($)", format="$%.0f"
                ),
                "rtd_oos": st.column_config.NumberColumn(
                    "R:DD OOS", format="%.2f"
                ),
                "profit_last_9_months": st.column_config.NumberColumn(
                    "Last 9M ($)", format="$%.0f"
                ),
                "avg_oos_drawdown": st.column_config.NumberColumn(
                    "Avg OOS DD ($)", format="$%.0f"
                ),
                "rtd_12_months": st.column_config.NumberColumn(
                    "R:DD 12M", format="%.2f"
                ),
                "count_profit_months": st.column_config.NumberColumn(
                    "Profit Months", format="%d",
                    help="# of profitable months in the eligibility lookback window."
                ),
                "oos_period_years": st.column_config.NumberColumn(
                    "OOS Years", format="%.1f"
                ),
                "sharpe_is": st.column_config.NumberColumn(
                    "Sharpe IS", format="%.2f"
                ),
                "max_drawdown_is": st.column_config.NumberColumn(
                    "Max DD IS ($)", format="$%.0f"
                ),
                "incubation_date": st.column_config.DateColumn("Incub. Date"),
                "quitting_status": st.column_config.TextColumn(
                    "Quit Status",
                    help="Continue / Quit / Coming Back / Recovered / N/A"
                ),
                "quitting_date": st.column_config.DateColumn("Quit Date"),
                "profit_since_quit": st.column_config.NumberColumn(
                    "P&L Since Quit ($)", format="$%.0f",
                    help="Cumulative OOS P&L from the date the strategy entered 'Quit' status to now.",
                ),
                "k_factor": st.column_config.NumberColumn(
                    "K-Factor",
                    format="%.2f",
                    help="(Win rate / Loss rate) × (Avg win / Avg loss) — monthly P&L.",
                ),
                "ulcer_index": st.column_config.NumberColumn(
                    "Ulcer Index",
                    format="%.2f",
                    help="RMS % drawdown over OOS period (Peter Martin, 1987). Lower = smoother equity curve.",
                ),
                "best_month": st.column_config.NumberColumn(
                    "Best Month ($)", format="$%.0f"
                ),
                "worst_month": st.column_config.NumberColumn(
                    "Worst Month ($)", format="$%.0f"
                ),
                "max_consecutive_loss_months": st.column_config.NumberColumn(
                    "Max Loss Streak", format="%d",
                    help="Maximum consecutive losing months in the OOS period."
                ),
            },
        )

        _sc_col, _em_col, _ = st.columns([1, 1, 4])
        with _sc_col:
            _save_contracts_clicked = st.button("Save Contracts", key="save_contracts_btn")
        with _em_col:
            if st.button("Export Metrics to Excel", key="export_port_metrics_btn"):
                from core.reporting.excel_export import export_portfolio, portfolio_export_filename
                _xlsx = export_portfolio(portfolio, config)
                st.download_button(
                    "📥 Download Portfolio Metrics",
                    data=_xlsx,
                    file_name=portfolio_export_filename(),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_port_metrics_xlsx",
                )

        if _save_contracts_clicked:
            _all_strats = load_strategies()
            _contracts_edit = {
                r["Strategy"]: r["contracts"]
                for r in _edited_metrics.to_dict(orient="records")
                if "contracts" in r
            }
            _updated = []
            for _s in _all_strats:
                _nm = _s.get("name", "")
                if _nm in _contracts_edit:
                    _s = dict(_s)
                    try:
                        _s["contracts"] = int(round(float(_contracts_edit[_nm] or 1)))
                    except (ValueError, TypeError):
                        pass
                _updated.append(_s)
            save_strategies(_updated)
            st.session_state.portfolio_data = None
            st.success("Contracts saved — click **Rebuild Portfolio** to recalculate.")
            st.rerun()

    else:
        st.info("Run an import with Walkforward Details CSVs to see strategy metrics.")
else:
    # Show basic table from trade data
    rows = []
    for s in portfolio.strategies:
        col_pnl = portfolio.daily_pnl[s.name]
        rows.append({
            "Strategy": s.name,
            "Contracts": s.contracts,
            "Symbol": s.symbol,
            "Sector": s.sector,
            "Total P&L ($)": col_pnl.sum(),
            "Last 12M ($)": col_pnl.iloc[-252:].sum() if len(col_pnl) >= 252 else col_pnl.sum(),
        })
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

st.divider()


# ── 3b. Metric charts by strategy / symbol / sector ───────────────────────────

st.subheader("Metric Analysis")

_CHARTABLE_METRICS: dict[str, str] = {
    "max_oos_drawdown":       "Max OOS Drawdown ($)",
    "avg_oos_drawdown":       "Avg OOS Drawdown ($)",
    "profit_last_12_months":  "Last 12M P&L ($)",
    "profit_last_6_months":   "Last 6M P&L ($)",
    "profit_last_3_months":   "Last 3M P&L ($)",
    "profit_since_oos_start": "OOS Total P&L ($)",
    "rtd_oos":                "RTD OOS",
    "rtd_12_months":          "RTD 12M",
    "expected_annual_profit": "Expected Annual ($)",
    "actual_annual_profit":   "Actual Annual ($)",
    "return_efficiency":      "Return Efficiency",
    "sharpe_isoos":           "Sharpe IS+OOS",
    "k_factor":               "K-Factor",
    "ulcer_index":            "Ulcer Index",
    "contracts":              "Contracts",
}

_DOLLAR_METRICS = {
    "max_oos_drawdown", "avg_oos_drawdown",
    "profit_last_12_months", "profit_last_6_months", "profit_last_3_months",
    "profit_since_oos_start", "expected_annual_profit", "actual_annual_profit",
}
_PCT_METRICS = {"return_efficiency"}

_SECTOR_PALETTE = [
    "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
    "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf",
    "#aec7e8", "#ffbb78", "#98df8a", "#ff9896", "#c5b0d5",
]

# Determine which chartable metrics are actually in the data
if not portfolio.summary_metrics.empty:
    _sm_chart = portfolio.summary_metrics.copy()

    # Ensure symbol/sector present (add from strategies_config if missing)
    _strats_lookup = {s["name"]: s for s in strategies_config}
    for _col in ("symbol", "sector"):
        if _col not in _sm_chart.columns:
            _sm_chart[_col] = _sm_chart.index.map(
                lambda n: _strats_lookup.get(n, {}).get(_col, "")
            )
    if "contracts" not in _sm_chart.columns:
        _sm_chart["contracts"] = _sm_chart.index.map(
            lambda n: float(_strats_lookup.get(n, {}).get("contracts", 1) or 1)
        )

    _avail = {k: v for k, v in _CHARTABLE_METRICS.items() if k in _sm_chart.columns}

    if _avail:
        _mc1, _mc2, _mc3 = st.columns([2, 1, 3])
        with _mc1:
            _chart_metric_key = "port_chart_metric"
            _default_metric = next(
                (k for k in ("max_oos_drawdown", "profit_last_12_months", "rtd_oos") if k in _avail),
                next(iter(_avail)),
            )
            _sel_metric = st.selectbox(
                "Metric to chart",
                list(_avail.keys()),
                index=list(_avail.keys()).index(_default_metric),
                format_func=lambda k: _avail[k],
                key=_chart_metric_key,
            )
        with _mc2:
            _sel_agg = st.selectbox(
                "Group aggregate",
                ["Sum", "Average", "Max", "Min"],
                key="port_chart_agg",
                help="Used when grouping by Symbol or Sector.",
            )

        _agg_fn_name = {"Sum": "sum", "Average": "mean", "Max": "max", "Min": "min"}[_sel_agg]

        # Build chart DataFrame
        _cdf = _sm_chart[[_sel_metric, "symbol", "sector"]].copy()
        _cdf.index.name = "Strategy"
        _cdf = _cdf.reset_index()
        _cdf["symbol"] = _cdf["symbol"].fillna("?")
        _cdf["sector"] = _cdf["sector"].fillna("Other").replace("", "Other")
        _cdf = _cdf[_cdf[_sel_metric].notna()].copy()

        if _cdf.empty:
            st.info(f"No data available for **{_avail[_sel_metric]}** — compute strategy summary to populate this metric.")
        else:
            _cdf[_sel_metric] = _cdf[_sel_metric].astype(float)

            # Value formatter
            def _fmt(v: float) -> str:
                if _sel_metric in _DOLLAR_METRICS:
                    return f"-${abs(v):,.0f}" if v < 0 else f"${v:,.0f}"
                if _sel_metric in _PCT_METRICS:
                    return f"{v:.1%}"
                return f"{v:.2f}"

            # Sector → colour mapping
            _unique_sectors = sorted(_cdf["sector"].unique())
            _sec_color = {s: _SECTOR_PALETTE[i % len(_SECTOR_PALETTE)]
                          for i, s in enumerate(_unique_sectors)}

            _tab_strat, _tab_sym, _tab_sec = st.tabs(["By Strategy", "By Symbol", "By Sector"])

            # ── By Strategy ───────────────────────────────────────────────────
            with _tab_strat:
                _label_map = build_label_map(portfolio.strategies)
                _strat_df = _cdf.copy()
                _strat_df["short_label"] = _strat_df["Strategy"].map(
                    lambda n: _label_map.get(n, n)
                )
                _strat_df = _strat_df.sort_values(_sel_metric, ascending=True)

                _strat_palette = _SECTOR_PALETTE  # reuse same 15-color palette
                _strat_colors = [
                    _strat_palette[i % len(_strat_palette)]
                    for i in range(len(_strat_df))
                ]

                _fig_strat = go.Figure()
                for i, row in enumerate(_strat_df.itertuples()):
                    _fig_strat.add_trace(go.Bar(
                        y=[row.short_label],
                        x=[getattr(row, _sel_metric)],
                        orientation="h",
                        name=row.short_label,
                        marker_color=_strat_colors[i],
                        text=[_fmt(getattr(row, _sel_metric))],
                        textposition="outside",
                        showlegend=False,
                        hovertemplate=f"{row.short_label} ({row.Strategy}): %{{text}}<extra></extra>",
                    ))

                _fig_strat.update_layout(
                    xaxis_title=_avail[_sel_metric],
                    height=max(350, len(_strat_df) * 28 + 80),
                    margin=dict(l=0, r=80, t=10, b=0),
                    showlegend=False,
                    barmode="stack",
                )
                st.plotly_chart(_fig_strat, use_container_width=True)
                render_legend(portfolio.strategies)

            # ── By Symbol ─────────────────────────────────────────────────────
            with _tab_sym:
                _sym_grp = (
                    _cdf.groupby("symbol")[_sel_metric]
                    .agg(_agg_fn_name)
                    .reset_index()
                    .sort_values(_sel_metric, ascending=False)
                )
                _fig_sym = go.Figure(go.Bar(
                    x=_sym_grp["symbol"],
                    y=_sym_grp[_sel_metric],
                    text=[_fmt(v) for v in _sym_grp[_sel_metric]],
                    textposition="outside",
                    marker_color="#1f77b4",
                    hovertemplate="%{x}: %{text}<extra></extra>",
                ))
                _fig_sym.update_layout(
                    xaxis_title="Symbol",
                    yaxis_title=f"{_sel_agg} — {_avail[_sel_metric]}",
                    height=400,
                    margin=dict(l=0, r=0, t=10, b=0),
                )
                st.plotly_chart(_fig_sym, use_container_width=True)

            # ── By Sector ─────────────────────────────────────────────────────
            with _tab_sec:
                _sec_grp = (
                    _cdf.groupby("sector")[_sel_metric]
                    .agg(_agg_fn_name)
                    .reset_index()
                    .sort_values(_sel_metric, ascending=False)
                )
                _sec_bar_colors = [_sec_color.get(s, "#888") for s in _sec_grp["sector"]]
                _fig_sec = go.Figure(go.Bar(
                    x=_sec_grp["sector"],
                    y=_sec_grp[_sel_metric],
                    text=[_fmt(v) for v in _sec_grp[_sel_metric]],
                    textposition="outside",
                    marker_color=_sec_bar_colors,
                    hovertemplate="%{x}: %{text}<extra></extra>",
                ))
                _fig_sec.update_layout(
                    xaxis_title="Sector",
                    yaxis_title=f"{_sel_agg} — {_avail[_sel_metric]}",
                    height=400,
                    margin=dict(l=0, r=0, t=10, b=0),
                )
                st.plotly_chart(_fig_sec, use_container_width=True)

    else:
        st.info("No chartable metrics available — compute strategy summary to populate metric charts.")

st.divider()


# ── 4. Monthly P&L heatmap ────────────────────────────────────────────────────

st.subheader("Monthly P&L — Total Portfolio")

monthly = monthly_portfolio_pnl(portfolio)
if not monthly.empty and "Total" in monthly.columns:
    total_monthly = monthly["Total"].to_frame()
    total_monthly["Year"] = total_monthly.index.year
    total_monthly["Month"] = total_monthly.index.strftime("%b")

    pivot = total_monthly.pivot_table(
        values="Total", index="Year", columns="Month", aggfunc="sum"
    )

    # Reorder months
    month_order = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
                   "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    pivot = pivot.reindex(columns=[m for m in month_order if m in pivot.columns])
    pivot = pivot.sort_index(ascending=False)

    _hm_vals = pivot.values.astype(float)
    # Format each cell as "$12,345" or "-$12,345"; blank for NaN
    def _fmt_pnl(v: float) -> str:
        return f"-${abs(v):,.0f}" if v < 0 else f"${v:,.0f}"
    _hm_text = [
        [_fmt_pnl(v) if not pd.isna(v) else "" for v in row]
        for row in _hm_vals
    ]
    fig_hm = go.Figure(go.Heatmap(
        z=_hm_vals,
        x=list(pivot.columns),
        y=[str(y) for y in pivot.index],
        colorscale="RdYlGn",
        zmid=0,
        text=_hm_text,
        texttemplate="%{text}",
        textfont={"size": 11},
        hovertemplate="%{y} %{x}: %{text}<extra></extra>",
        colorbar=dict(title="P&L ($)"),
    ))
    fig_hm.update_layout(
        height=max(200, len(pivot) * 40 + 100),
        margin=dict(l=0, r=0, t=10, b=0),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig_hm, use_container_width=True)

    # Annual totals
    annual = total_monthly.groupby("Year")["Total"].sum().sort_index(ascending=False)
    st.caption("Annual P&L totals")
    annual_df = annual.reset_index()
    annual_df.columns = ["Year", "Total P&L ($)"]
    st.dataframe(annual_df, use_container_width=True, hide_index=True)

# ── Analytics navigation ────────────────────────────────────────────────────────

st.divider()
st.markdown("**Portfolio built. Explore analytics:**")
_a_cols = st.columns(4)
_analytics = [
    ("Monte Carlo", "ui/pages/_04_Monte_Carlo.py"),
    ("Correlations", "ui/pages/_05_Correlations.py"),
    ("Diversification", "ui/pages/_06_Diversification.py"),
    ("Leave One Out", "ui/pages/_07_Leave_One_Out.py"),
    ("Backtest", "ui/pages/_08_Backtest.py"),
    ("Eligibility Backtest", "ui/pages/_09_Eligibility_Backtest.py"),
    ("Margin Tracking", "ui/pages/_10_Margin_Tracking.py"),
    ("Position Check", "ui/pages/_11_Position_Check.py"),
]
for _i, (_label, _page) in enumerate(_analytics):
    with _a_cols[_i % 4]:
        st.page_link(_page, label=_label)
