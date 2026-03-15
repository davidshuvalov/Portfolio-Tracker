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
import plotly.express as px
from datetime import date

from core.config import AppConfig
from core.portfolio.strategies import load_strategies
from core.portfolio.aggregator import (
    build_portfolio,
    portfolio_total_pnl,
    portfolio_equity_curve,
    monthly_portfolio_pnl,
    portfolio_summary_stats,
)
from core.portfolio.summary import compute_summary

st.set_page_config(page_title="Portfolio", layout="wide")

# ── Sidebar workflow status ────────────────────────────────────────────────────
try:
    from ui.workflow import render_workflow_sidebar
    with st.sidebar:
        render_workflow_sidebar()
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

col_rebuild, col_status = st.columns([1, 5])
with col_rebuild:
    if st.button("Build Portfolio" if _needs_rebuild else "Rebuild Portfolio",
                 type="primary" if _needs_rebuild else "secondary"):
        _needs_rebuild = True

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
        )

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

st.divider()


# ── 3. Per-strategy metrics table ─────────────────────────────────────────────

st.subheader("Strategy Metrics")

if not portfolio.summary_metrics.empty:
    sm = portfolio.summary_metrics.copy()

    # Select key display columns
    display_cols = [
        c for c in [
            "contracts", "symbol", "sector", "oos_begin", "oos_end",
            "expected_annual_profit", "actual_annual_profit", "return_efficiency",
            "trades_per_year", "overall_win_rate",
            "max_drawdown_isoos", "sharpe_isoos",
            "profit_last_1_month", "profit_last_3_months", "profit_last_6_months",
            "profit_last_12_months", "profit_since_oos_start",
            "max_oos_drawdown", "rtd_oos",
            "incubation_status",
        ]
        if c in sm.columns
    ]

    # Add contracts from strategies config
    if "contracts" not in sm.columns:
        contracts_map = {s.name: s.contracts for s in portfolio.strategies}
        sm["contracts"] = sm.index.map(contracts_map)

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
                "contracts": st.column_config.NumberColumn(
                    "Contr.", format="%d", min_value=0, max_value=999, step=1
                ),
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
                "sharpe_isoos": st.column_config.NumberColumn("Sharpe", format="%.2f"),
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
            },
        )

        if st.button("Save Contracts", key="save_contracts_btn"):
            from core.portfolio.strategies import load_strategies, save_strategies as _save
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
                        _s["contracts"] = int(_contracts_edit[_nm] or 1)
                    except (ValueError, TypeError):
                        pass
                _updated.append(_s)
            _save(_updated)
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

    fig_hm = px.imshow(
        pivot,
        color_continuous_scale="RdYlGn",
        color_continuous_midpoint=0,
        text_auto=".0f",
        aspect="auto",
        labels={"color": "P&L ($)"},
    )
    fig_hm.update_layout(
        height=max(200, len(pivot) * 40 + 100),
        margin=dict(l=0, r=0, t=10, b=0),
        coloraxis_showscale=False,
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
