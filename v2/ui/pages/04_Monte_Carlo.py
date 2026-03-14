"""
Monte Carlo page — simulation for portfolio total or a single strategy.
Mirrors the VBA Monte Carlo tab (K_MonteCarlo.bas).

Layout:
  Sidebar: all config controls + Run button
  Main:    6 metric cards → two distribution charts → scenario stats table
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.analytics.monte_carlo import run_monte_carlo
from core.config import AppConfig, MCConfig
from core.data_types import PortfolioData, Strategy
from core.portfolio.aggregator import portfolio_total_pnl

st.set_page_config(page_title="Monte Carlo", layout="wide")
st.title("Monte Carlo Simulation")

config: AppConfig = st.session_state.get("config", AppConfig.load())
imported = st.session_state.get("imported_data")
portfolio: PortfolioData | None = st.session_state.get("portfolio_data")

if imported is None:
    st.info("No data loaded. Go to **Import** first.")
    st.stop()

# ── Sidebar config ────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("MC Settings")

    # Target: portfolio or single strategy
    target_options = ["Portfolio"] + list(imported.strategy_names)
    mc_target = st.selectbox(
        "Target",
        target_options,
        help="Run MC on total portfolio PnL or a single strategy",
    )

    period = st.selectbox(
        "Period",
        ["OOS", "IS", "IS+OOS"],
        index=0,
        help="Data period to sample trade PnL from",
    )

    trade_option = st.radio(
        "Trade data",
        ["M2M", "Closed"],
        horizontal=True,
        help="M2M: daily mark-to-market  |  Closed: closed-trade PnL",
    )

    simulations = st.number_input(
        "Simulations",
        min_value=1_000,
        max_value=100_000,
        value=int(config.monte_carlo.simulations),
        step=1_000,
    )

    trade_adj_pct = st.slider(
        "Trade adjustment %",
        min_value=-50,
        max_value=50,
        value=int(config.monte_carlo.trade_adjustment * 100),
        step=1,
        help="Reduce each trade by this % (stress test slippage / commission)",
    )
    trade_adjustment = trade_adj_pct / 100.0

    risk_ruin_pct = st.slider(
        "Risk-of-ruin target %",
        min_value=1,
        max_value=30,
        value=int(config.monte_carlo.risk_ruin_target * 100),
        step=1,
        help="Solver finds starting equity that achieves this ruin probability",
    )
    risk_ruin_target = risk_ruin_pct / 100.0

    margin_threshold = st.number_input(
        "Margin threshold ($)",
        min_value=0,
        max_value=10_000_000,
        value=5_000,
        step=500,
        help="Account balance below which = ruin (typically initial margin)",
    )

    run_btn = st.button("Run Monte Carlo", type="primary", use_container_width=True)

    st.divider()
    if st.button("Save as defaults", use_container_width=True, help="Persist these settings so they load next session"):
        config.monte_carlo.simulations      = int(simulations)
        config.monte_carlo.period           = period
        config.monte_carlo.risk_ruin_target = risk_ruin_target
        config.monte_carlo.trade_adjustment = trade_adjustment
        config.monte_carlo.trade_option     = trade_option
        config.save()
        st.session_state.config = config
        st.success("Saved.")


# ── Build PnL series for selected target ─────────────────────────────────────

def _build_pnl_series(
    target_name: str,
) -> tuple[pd.Series | None, pd.Series | None, Strategy | None]:
    """
    Return (daily_m2m, closed_daily, strategy_or_None) for the chosen target.
    Portfolio MC: sums active strategies (scaled by contracts).
    Strategy MC:  returns that strategy's column directly.
    """
    if target_name == "Portfolio":
        if portfolio is None:
            st.error("Portfolio not built. Go to **Portfolio** page first.")
            return None, None, None

        m2m_total = portfolio_total_pnl(portfolio)

        # Build portfolio closed-trade daily PnL (scaled by contracts)
        closed_daily: pd.Series | None = None
        if imported is not None and not imported.closed_trade_pnl.empty:
            active_names = [s.name for s in portfolio.strategies]
            contracts_map = {s.name: s.contracts for s in portfolio.strategies}
            avail = [n for n in active_names if n in imported.closed_trade_pnl.columns]
            if avail:
                cdf = imported.closed_trade_pnl[avail].copy()
                for n in avail:
                    cdf[n] *= contracts_map.get(n, 1)
                closed_daily = cdf.sum(axis=1)

        return m2m_total, closed_daily, None  # No single strategy → use full date range

    else:
        # Single strategy
        if target_name not in imported.strategy_names:
            st.error(f"Strategy '{target_name}' not found in imported data.")
            return None, None, None

        m2m = imported.daily_m2m[target_name]

        closed: pd.Series | None = None
        if target_name in imported.closed_trade_pnl.columns:
            closed = imported.closed_trade_pnl[target_name]

        strategy_obj: Strategy | None = next(
            (s for s in imported.strategies if s.name == target_name), None
        )
        return m2m, closed, strategy_obj


# ── Run ───────────────────────────────────────────────────────────────────────

result = st.session_state.get("mc_result")
mc_target_label = st.session_state.get("mc_target_label", "")

if run_btn:
    daily_m2m, closed_daily, strategy_obj = _build_pnl_series(mc_target)

    if daily_m2m is None or daily_m2m.empty:
        st.error("No PnL data available for the selected target.")
        st.stop()

    mc_config = MCConfig(
        simulations=int(simulations),
        period=period,
        risk_ruin_target=risk_ruin_target,
        risk_ruin_tolerance=config.monte_carlo.risk_ruin_tolerance,
        trade_adjustment=trade_adjustment,
        trade_option=trade_option,
    )

    with st.status(f"Running Monte Carlo — {mc_target}", expanded=True) as _mc_status:
        st.write("Building trade sequence…")
        _mc_status.update(label="Running simulations…")
        result = run_monte_carlo(
            daily_m2m=daily_m2m,
            config=mc_config,
            margin_threshold=float(margin_threshold),
            closed_daily=closed_daily,
            strategy=strategy_obj,
            return_scenarios=True,
        )
        _mc_status.update(
            label=f"Done — {simulations:,} scenarios complete",
            state="complete",
            expanded=False,
        )

    st.session_state.mc_result = result
    st.session_state.mc_target_label = mc_target
    mc_target_label = mc_target


# ── Display ───────────────────────────────────────────────────────────────────

if result is None:
    st.info("Configure settings in the sidebar and click **Run Monte Carlo**.")
    st.stop()

st.subheader(f"Results — {mc_target_label}")

# Metric cards
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric(
    "Starting Equity",
    f"${result.starting_equity:,.0f}",
    help="Capital required to achieve the target risk-of-ruin probability",
)
c2.metric(
    "Expected Annual Profit",
    f"${result.expected_profit:,.0f}",
)
c3.metric(
    "Risk of Ruin",
    f"{result.risk_of_ruin:.1%}" if not np.isnan(result.risk_of_ruin) else "N/A",
    help="Probability equity fell below margin threshold",
)
c4.metric(
    "Max Drawdown",
    f"{result.max_drawdown_pct:.1%}",
    help="Median max peak-to-trough drawdown across scenarios",
)
c5.metric("Sharpe Ratio", f"{result.sharpe_ratio:.2f}")
c6.metric("Return / Drawdown", f"{result.return_to_drawdown:.2f}")

st.divider()

# Distribution charts
if result.scenarios_df is not None and not result.scenarios_df.empty:
    df = result.scenarios_df

    col_left, col_right = st.columns(2)

    with col_left:
        median_profit = float(df["profit"].median())
        fig = px.histogram(
            df,
            x="profit",
            nbins=60,
            title="Distribution of Annual Profits",
            labels={"profit": "Annual Profit ($)"},
            color_discrete_sequence=["#2196F3"],
        )
        fig.add_vline(
            x=median_profit,
            line_dash="dash",
            line_color="orange",
            annotation_text=f"Median ${median_profit:,.0f}",
            annotation_position="top right",
        )
        fig.add_vline(x=0, line_color="red", opacity=0.5, line_width=1)
        fig.update_layout(height=360, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with col_right:
        median_dd = float(df["max_drawdown_pct"].median())
        fig2 = px.histogram(
            df,
            x="max_drawdown_pct",
            nbins=60,
            title="Distribution of Max Drawdowns",
            labels={"max_drawdown_pct": "Max Drawdown"},
            color_discrete_sequence=["#F44336"],
        )
        fig2.update_xaxes(tickformat=".0%")
        fig2.add_vline(
            x=median_dd,
            line_dash="dash",
            line_color="orange",
            annotation_text=f"Median {median_dd:.1%}",
            annotation_position="top right",
        )
        fig2.update_layout(height=360, showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    # Scenario stats table
    with st.expander("Scenario Statistics", expanded=False):
        pct_profitable = float((df["profit"] > 0).mean())
        stats = pd.DataFrame(
            {
                "Metric": [
                    "Expected Profit (Mean)",
                    "Median Profit",
                    "Std Dev of Profit",
                    "% Profitable Scenarios",
                    "5th Pct Profit",
                    "95th Pct Profit",
                    "Median Max Drawdown",
                    "95th Pct Max Drawdown",
                    "Expected Annual Return",
                    "Starting Equity",
                ],
                "Value": [
                    f"${df['profit'].mean():,.0f}",
                    f"${df['profit'].median():,.0f}",
                    f"${df['profit'].std():,.0f}",
                    f"{pct_profitable:.1%}",
                    f"${df['profit'].quantile(0.05):,.0f}",
                    f"${df['profit'].quantile(0.95):,.0f}",
                    f"{df['max_drawdown_pct'].median():.1%}",
                    f"{df['max_drawdown_pct'].quantile(0.95):.1%}",
                    f"{df['profit'].mean() / result.starting_equity:.1%}" if result.starting_equity > 0 else "N/A",
                    f"${result.starting_equity:,.0f}",
                ],
            }
        )
        st.dataframe(stats, hide_index=True, use_container_width=True)
