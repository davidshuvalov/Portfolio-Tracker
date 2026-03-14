"""
Leave-One-Out (LOO) sensitivity analysis page.
Mirrors the VBA LOO tab (J_LOO.bas).

For each live strategy: remove it → re-run Monte Carlo → show delta vs base.

Sections:
  1. Config + Run button
  2. Results table (delta metrics, sortable)
  3. Bar charts for delta_profit and delta_sharpe
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st

from core.analytics.leave_one_out import run_leave_one_out
from core.config import AppConfig, MCConfig
from core.data_types import PortfolioData

st.set_page_config(page_title="Leave One Out", layout="wide")
st.title("Leave-One-Out Analysis")

st.caption(
    "Remove each strategy from the portfolio one at a time and re-run Monte Carlo. "
    "Negative delta = removing that strategy **hurts** the portfolio (it adds value). "
    "Positive delta = removing it **helps** (it's a drag)."
)

config: AppConfig = st.session_state.get("config", AppConfig.load())
portfolio: PortfolioData | None = st.session_state.get("portfolio_data")

if portfolio is None:
    st.info("Portfolio not built. Go to **Portfolio** page first.")
    st.stop()

if len(portfolio.strategies) < 2:
    st.warning("Need at least 2 live strategies to run LOO analysis.")
    st.stop()

# ── Sidebar config ────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("LOO Settings")

    simulations = st.number_input(
        "Simulations per LOO run",
        min_value=500,
        max_value=20_000,
        value=min(int(config.monte_carlo.simulations), 2_000),
        step=500,
        help="Fewer simulations = faster runs. 2,000 gives stable results.",
    )

    margin_threshold = st.number_input(
        "Margin threshold ($)",
        min_value=0,
        max_value=10_000_000,
        value=5_000,
        step=500,
        help="Same ruin threshold as Monte Carlo page.",
    )

    trade_option = st.radio(
        "Trade data",
        ["M2M", "Closed"],
        horizontal=True,
    )

    trade_adj_pct = st.slider("Trade adjustment %", -50, 50, 0, 1)
    trade_adjustment = trade_adj_pct / 100.0

    risk_ruin_pct = st.slider(
        "Risk-of-ruin target %",
        1, 30,
        int(config.monte_carlo.risk_ruin_target * 100),
        1,
    )

    run_btn = st.button(
        f"Run LOO ({len(portfolio.strategies)} strategies)",
        type="primary",
        use_container_width=True,
    )

    st.divider()
    if st.button("Save as defaults", use_container_width=True, help="Persist these settings so they load next session"):
        config.monte_carlo.simulations      = int(simulations)
        config.monte_carlo.risk_ruin_target = risk_ruin_pct / 100.0
        config.monte_carlo.trade_adjustment = trade_adjustment
        config.monte_carlo.trade_option     = trade_option
        config.save()
        st.session_state.config = config
        st.success("Saved.")

# ── Run ───────────────────────────────────────────────────────────────────────
loo_result: pd.DataFrame | None = st.session_state.get("loo_result")

if run_btn:
    mc_config = MCConfig(
        simulations=int(simulations),
        period="IS+OOS",
        risk_ruin_target=risk_ruin_pct / 100.0,
        risk_ruin_tolerance=config.monte_carlo.risk_ruin_tolerance,
        trade_adjustment=trade_adjustment,
        trade_option=trade_option,
    )

    n = len(portfolio.strategies)
    progress = st.progress(0, text=f"Running LOO for {n} strategies…")

    # Monkey-patch progress into run_leave_one_out by calling it with a custom loop
    # We import the internal helpers directly for progress tracking
    from core.analytics.leave_one_out import _analyse_portfolio, _remove_strategy
    from core.analytics.monte_carlo import run_monte_carlo
    from core.portfolio.aggregator import portfolio_total_pnl

    base = _analyse_portfolio(portfolio, mc_config, float(margin_threshold))
    rows = []
    for idx, strategy in enumerate(portfolio.strategies):
        progress.progress(
            (idx + 1) / (n + 1),
            text=f"Removing {strategy.name} ({idx + 1}/{n})…",
        )
        reduced = _remove_strategy(portfolio, strategy.name)
        result = _analyse_portfolio(reduced, mc_config, float(margin_threshold))
        rows.append({
            "strategy":       strategy.name,
            "delta_profit":   result.expected_profit - base.expected_profit,
            "delta_sharpe":   result.sharpe_ratio - base.sharpe_ratio,
            "delta_drawdown": result.max_drawdown_pct - base.max_drawdown_pct,
            "delta_rtd":      result.return_to_drawdown - base.return_to_drawdown,
            "delta_ror":      result.risk_of_ruin - base.risk_of_ruin,
            "base_profit":    base.expected_profit,
            "result_profit":  result.expected_profit,
        })

    progress.progress(1.0, text="Done.")
    loo_result = pd.DataFrame(rows).sort_values("delta_profit", ascending=True).reset_index(drop=True)
    st.session_state.loo_result = loo_result
    st.session_state.loo_base_profit = base.expected_profit
    st.session_state.loo_base_sharpe = base.sharpe_ratio

# ── Display ───────────────────────────────────────────────────────────────────
if loo_result is None:
    st.info("Configure settings in the sidebar and click **Run LOO**.")
    st.stop()

base_profit = st.session_state.get("loo_base_profit", 0.0)
base_sharpe = st.session_state.get("loo_base_sharpe", 0.0)

# Summary cards
c1, c2, c3 = st.columns(3)
c1.metric("Strategies Analysed", len(loo_result))
c2.metric("Base Portfolio Expected Profit", f"${base_profit:,.0f}")
c3.metric("Base Portfolio Sharpe", f"{base_sharpe:.2f}")

st.divider()

# ── Results table ─────────────────────────────────────────────────────────────
st.subheader("LOO Results Table")
st.caption("Sort by any column. Rows in **red** = strategy hurts portfolio (positive delta_profit).")

display = loo_result[[
    "strategy", "delta_profit", "delta_sharpe",
    "delta_drawdown", "delta_rtd", "delta_ror",
]].copy()
display["delta_profit"]   = display["delta_profit"].round(0).astype(int)
display["delta_sharpe"]   = display["delta_sharpe"].round(3)
display["delta_drawdown"] = display["delta_drawdown"].apply(lambda x: f"{x:+.1%}")
display["delta_rtd"]      = display["delta_rtd"].round(2)
display["delta_ror"]      = display["delta_ror"].apply(lambda x: f"{x:+.1%}")
display.columns = [
    "Strategy", "ΔProfit ($)", "ΔSharpe",
    "ΔDrawdown", "ΔRTD", "ΔRoR",
]

st.dataframe(display, hide_index=True, use_container_width=True)

st.divider()

# ── Bar charts ────────────────────────────────────────────────────────────────
col_left, col_right = st.columns(2)

with col_left:
    fig = px.bar(
        loo_result.sort_values("delta_profit"),
        x="delta_profit",
        y="strategy",
        orientation="h",
        title="ΔExpected Profit when Strategy Removed",
        labels={"delta_profit": "ΔProfit ($)", "strategy": ""},
        color="delta_profit",
        color_continuous_scale=["#4CAF50", "#FFEB3B", "#F44336"],
        color_continuous_midpoint=0,
    )
    fig.add_vline(x=0, line_color="black", line_width=1)
    fig.update_layout(
        height=max(350, len(loo_result) * 30 + 80),
        coloraxis_showscale=False,
        yaxis={"categoryorder": "total ascending"},
    )
    st.plotly_chart(fig, use_container_width=True)

with col_right:
    fig2 = px.bar(
        loo_result.sort_values("delta_sharpe"),
        x="delta_sharpe",
        y="strategy",
        orientation="h",
        title="ΔSharpe when Strategy Removed",
        labels={"delta_sharpe": "ΔSharpe", "strategy": ""},
        color="delta_sharpe",
        color_continuous_scale=["#4CAF50", "#FFEB3B", "#F44336"],
        color_continuous_midpoint=0,
    )
    fig2.add_vline(x=0, line_color="black", line_width=1)
    fig2.update_layout(
        height=max(350, len(loo_result) * 30 + 80),
        coloraxis_showscale=False,
        yaxis={"categoryorder": "total ascending"},
    )
    st.plotly_chart(fig2, use_container_width=True)

# ── Interpretation guide ──────────────────────────────────────────────────────
with st.expander("How to interpret these results", expanded=False):
    st.markdown("""
| Delta Sign | Meaning |
|------------|---------|
| **ΔProfit < 0** | Removing this strategy *reduces* expected profit → it adds value |
| **ΔProfit > 0** | Removing this strategy *increases* expected profit → it's a drag |
| **ΔSharpe < 0** | Strategy improves risk-adjusted return |
| **ΔDrawdown < 0** | Strategy reduces portfolio drawdown |
| **ΔRoR < 0** | Strategy reduces risk of ruin |

A strategy is **portfolio-positive** when ΔProfit and ΔSharpe are both negative.
A strategy is a **drag** when both are positive.
Mixed signals require judgement (e.g. adds profit but also adds drawdown).
""")
