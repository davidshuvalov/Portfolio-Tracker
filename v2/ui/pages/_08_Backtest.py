"""
Backtest page — historical period performance analysis.

Sections:
  1. Period selector (predefined windows or custom date range)
  2. Equity curve + per-strategy contribution
  3. Monthly PnL heatmap
  4. Summary metrics (return, drawdown, Sharpe, win rate)
"""

from __future__ import annotations

from datetime import date, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.config import AppConfig
from core.data_types import PortfolioData
from core.portfolio.aggregator import portfolio_equity_curve, portfolio_total_pnl
from ui.strategy_labels import render_strategy_picker

st.set_page_config(page_title="Backtest", layout="wide")
st.title("Backtest")

config: AppConfig = st.session_state.get("config", AppConfig.load())
portfolio: PortfolioData | None = st.session_state.get("portfolio_data")

if portfolio is None:
    st.info("Portfolio not built yet.")
    st.page_link("ui/pages/03_Portfolio.py", label="Go to Portfolio →")
    st.stop()

if not portfolio.strategies:
    st.warning("No live strategies in portfolio.")
    st.stop()

total_pnl = portfolio_total_pnl(portfolio)
if total_pnl.empty:
    st.warning("No PnL data available.")
    st.stop()

date_min = total_pnl.index.min().date()
date_max = total_pnl.index.max().date()

# ── Period controls ───────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Period")

    period_preset = st.selectbox(
        "Preset",
        ["All Data", "Last 1 Year", "Last 2 Years", "Last 3 Years",
         "Last 5 Years", "Custom"],
    )

    if period_preset == "Custom":
        start_date = st.date_input("Start", value=date_min, min_value=date_min, max_value=date_max)
        end_date   = st.date_input("End",   value=date_max, min_value=date_min, max_value=date_max)
    else:
        end_date = date_max
        years_map = {
            "All Data": None,
            "Last 1 Year": 1,
            "Last 2 Years": 2,
            "Last 3 Years": 3,
            "Last 5 Years": 5,
        }
        yrs = years_map[period_preset]
        start_date = (
            date_max - timedelta(days=int(yrs * 365.25))
            if yrs else date_min
        )

    show_individual = st.checkbox("Show individual strategies", value=True)
    normalize = st.checkbox("Normalise curves to 0 start", value=False)

    st.divider()
    render_strategy_picker(portfolio.strategies, key="bt_strat_picker")

# ── Filter to period ──────────────────────────────────────────────────────────
start_ts = pd.Timestamp(start_date)
end_ts   = pd.Timestamp(end_date)

pnl_window = total_pnl.loc[(total_pnl.index >= start_ts) & (total_pnl.index <= end_ts)]
port_pnl_window = portfolio.daily_pnl.loc[
    (portfolio.daily_pnl.index >= start_ts) & (portfolio.daily_pnl.index <= end_ts)
]

if pnl_window.empty:
    st.warning("No data in selected period.")
    st.stop()

# ── Summary metrics ───────────────────────────────────────────────────────────
total_return  = float(pnl_window.sum())
n_days        = len(pnl_window)
years         = max(n_days / 252.0, 1e-3)
ann_return    = total_return / years

equity        = pnl_window.cumsum()
peak          = equity.cummax()
dd            = peak - equity
max_dd        = float(dd.max())
avg_dd        = float(dd[dd > 0].mean()) if (dd > 0).any() else 0.0

monthly_w     = pnl_window.resample("ME").sum()
win_rate      = float((monthly_w > 0).mean()) if len(monthly_w) > 0 else 0.0
std_m         = float(monthly_w.std()) if len(monthly_w) > 1 else 0.0
sharpe        = (float(monthly_w.mean()) / std_m * np.sqrt(12)) if std_m > 1e-9 else 0.0

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Total Return",       f"${total_return:,.0f}")
c2.metric("Ann. Return",        f"${ann_return:,.0f}")
c3.metric("Max Drawdown",       f"${max_dd:,.0f}")
c4.metric("Avg Drawdown",       f"${avg_dd:,.0f}")
c5.metric("Monthly Win Rate",   f"{win_rate:.1%}")
c6.metric("Sharpe (Monthly)",   f"{sharpe:.2f}")

st.caption(f"Period: **{start_date}** → **{end_date}** · {n_days} trading days")
st.divider()

# ── Equity curve ──────────────────────────────────────────────────────────────
st.subheader("Equity Curve")
fig = go.Figure()

# Portfolio total
eq_curve = pnl_window.cumsum()
if normalize:
    eq_curve = eq_curve - eq_curve.iloc[0]

fig.add_trace(go.Scatter(
    x=eq_curve.index, y=eq_curve.values,
    name="Portfolio",
    line=dict(color="#1565C0", width=2.5),
))

# Individual strategies
if show_individual and not port_pnl_window.empty:
    palette = px.colors.qualitative.Plotly
    for i, strat in enumerate(port_pnl_window.columns):
        s_eq = port_pnl_window[strat].cumsum()
        if normalize:
            s_eq = s_eq - s_eq.iloc[0]
        fig.add_trace(go.Scatter(
            x=s_eq.index, y=s_eq.values,
            name=strat,
            line=dict(width=1, dash="dot", color=palette[i % len(palette)]),
            opacity=0.7,
        ))

fig.update_layout(
    height=420,
    xaxis_title="Date",
    yaxis_title="Cumulative PnL ($)",
    hovermode="x unified",
    legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
)
st.plotly_chart(fig, use_container_width=True)

# ── Drawdown chart ────────────────────────────────────────────────────────────
with st.expander("Drawdown", expanded=False):
    dd_series = -(peak - equity)
    fig_dd = go.Figure(go.Scatter(
        x=dd_series.index, y=dd_series.values,
        fill="tozeroy",
        line=dict(color="#F44336"),
        name="Drawdown ($)",
    ))
    fig_dd.update_layout(height=250, yaxis_title="Drawdown ($)", showlegend=False)
    st.plotly_chart(fig_dd, use_container_width=True)

# ── Monthly PnL heatmap ───────────────────────────────────────────────────────
st.subheader("Monthly PnL")

monthly_total = pnl_window.resample("ME").sum()

if len(monthly_total) > 0:
    # Build year × month grid
    monthly_df = pd.DataFrame({
        "year":  monthly_total.index.year,
        "month": monthly_total.index.month,
        "pnl":   monthly_total.values,
    })
    pivot = monthly_df.pivot(index="year", columns="month", values="pnl")
    pivot.columns = [
        ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][c - 1]
        for c in pivot.columns
    ]

    fig_heat = px.imshow(
        pivot,
        color_continuous_scale="RdYlGn",
        color_continuous_midpoint=0,
        text_auto=".0f",
        title="Monthly PnL ($) — Portfolio Total",
        aspect="auto",
    )
    fig_heat.update_layout(height=max(250, len(pivot) * 40 + 80))
    st.plotly_chart(fig_heat, use_container_width=True)

# ── Per-strategy contribution table ──────────────────────────────────────────
st.subheader("Per-Strategy Contribution")

if not port_pnl_window.empty:
    strat_totals = port_pnl_window.sum()
    strat_ann    = strat_totals / years
    strat_pct    = strat_totals / (abs(strat_totals).sum() + 1e-9) * 100

    contrib_df = pd.DataFrame({
        "Strategy":      strat_totals.index,
        "Total PnL ($)": strat_totals.values.round(0).astype(int),
        "Ann. PnL ($)":  strat_ann.values.round(0).astype(int),
        "Contribution %": strat_pct.values.round(1),
    })
    contrib_df = contrib_df.sort_values("Total PnL ($)", ascending=False)

    st.dataframe(contrib_df, hide_index=True, use_container_width=True)

    # Bar chart
    fig_bar = px.bar(
        contrib_df,
        x="Total PnL ($)",
        y="Strategy",
        orientation="h",
        title="Strategy Contribution",
        color="Total PnL ($)",
        color_continuous_scale=["#F44336", "#FFEB3B", "#4CAF50"],
        color_continuous_midpoint=0,
    )
    fig_bar.update_layout(height=max(300, len(contrib_df) * 32 + 80), coloraxis_showscale=False)
    st.plotly_chart(fig_bar, use_container_width=True)
