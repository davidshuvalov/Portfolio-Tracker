"""
Margin Tracking page — daily margin utilisation by symbol and sector.
Mirrors M_Margin_Tracking.bas (CreateContractMarginTracking).

Sections:
  1. Symbol margin configuration (editable table in sidebar)
  2. Summary header cards
  3. Total daily margin time series
  4. Margin by symbol (stacked area chart)
  5. Margin by sector
  6. Symbol activity calendar (days in-market per symbol)
"""

from __future__ import annotations

from datetime import date, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.analytics.margin import (
    compute_daily_margin,
    margin_by_sector,
    margin_by_symbol,
    margin_summary_stats,
)
from core.config import AppConfig
from core.data_types import PortfolioData

st.set_page_config(page_title="Margin Tracking", layout="wide")
st.title("Margin Tracking")

config: AppConfig = st.session_state.get("config", AppConfig.load())
portfolio: PortfolioData | None = st.session_state.get("portfolio_data")
imported = st.session_state.get("imported_data")

if imported is None:
    st.info("No data loaded. Go to **Import** first.")
    st.stop()

if portfolio is None:
    st.info("Portfolio not built. Go to **Portfolio** page first.")
    st.stop()

if not portfolio.strategies:
    st.warning("No live strategies in portfolio.")
    st.stop()

# ── Discover symbols from active strategies ───────────────────────────────────
all_symbols = sorted({s.symbol for s in portfolio.strategies if s.symbol})

# ── Sidebar: margin configuration ────────────────────────────────────────────
with st.sidebar:
    st.header("Margin Settings")

    default_margin = st.number_input(
        "Default margin / contract ($)",
        min_value=0, max_value=1_000_000,
        value=int(config.default_margin),
        step=500,
        help="Used for any symbol not listed below",
    )

    st.caption("Per-symbol margin requirements:")

    # Editable symbol → margin table
    init_rows = [
        {"Symbol": sym, "Margin ($)": int(config.symbol_margins.get(sym, default_margin))}
        for sym in all_symbols
    ]
    if not init_rows:
        init_rows = [{"Symbol": "", "Margin ($)": int(default_margin)}]

    margin_table = st.data_editor(
        pd.DataFrame(init_rows),
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "Symbol":    st.column_config.TextColumn("Symbol"),
            "Margin ($)": st.column_config.NumberColumn("Margin ($)", min_value=0, step=500),
        },
        key="margin_table_editor",
    )

    if st.button("Save Margins", use_container_width=True):
        new_margins = {
            row["Symbol"]: float(row["Margin ($)"])
            for _, row in margin_table.iterrows()
            if row["Symbol"] and not pd.isna(row["Symbol"])
        }
        config.symbol_margins = new_margins
        config.default_margin = float(default_margin)
        config.save()
        st.session_state.config = config
        st.success("Saved.")

    st.divider()

    period_label = st.selectbox(
        "Period",
        ["All Data", "Last 1 Year", "Last 2 Years", "Last 3 Years"],
    )
    period_map = {"All Data": None, "Last 1 Year": 1, "Last 2 Years": 2, "Last 3 Years": 3}
    period_years = period_map[period_label]

# ── Build symbol margins dict ─────────────────────────────────────────────────
symbol_margins: dict[str, float] = {
    row["Symbol"]: float(row["Margin ($)"])
    for _, row in margin_table.iterrows()
    if row["Symbol"] and not pd.isna(row["Symbol"])
}
dflt = float(default_margin)

# ── Compute margin series ─────────────────────────────────────────────────────
with st.spinner("Computing margin utilisation…"):
    daily_margin = compute_daily_margin(
        imported.in_market_long,
        imported.in_market_short,
        portfolio.strategies,
        symbol_margins,
        dflt,
    )
    sym_margin = margin_by_symbol(
        imported.in_market_long,
        imported.in_market_short,
        portfolio.strategies,
        symbol_margins,
        dflt,
    )
    sec_margin = margin_by_sector(sym_margin, portfolio.strategies, symbol_margins)

# ── Apply period filter ───────────────────────────────────────────────────────
if not daily_margin.empty and period_years is not None:
    cutoff = daily_margin.index[-1] - pd.DateOffset(years=period_years)
    daily_margin = daily_margin[daily_margin.index >= cutoff]
    if not sym_margin.empty:
        sym_margin = sym_margin[sym_margin.index >= cutoff]
    if not sec_margin.empty:
        sec_margin = sec_margin[sec_margin.index >= cutoff]

if daily_margin.empty:
    st.warning("No in-market data found. Check that in_market_long/short data was imported.")
    st.stop()

# ── Summary cards ─────────────────────────────────────────────────────────────
stats = margin_summary_stats(daily_margin, sym_margin)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Current Margin",   f"${stats.get('current_margin', 0):,.0f}")
c2.metric("Peak Margin",      f"${stats.get('peak_margin', 0):,.0f}")
c3.metric("Average Margin",   f"${stats.get('average_margin', 0):,.0f}")
c4.metric("Days at Peak",     str(stats.get("days_at_peak", 0)))
c5.metric(
    f"Top Symbol ({stats.get('top_symbol','N/A')})",
    f"${stats.get('top_symbol_avg', 0):,.0f} avg",
)

st.divider()

# ── Total margin time series ──────────────────────────────────────────────────
st.subheader("Total Portfolio Margin Over Time")

fig = go.Figure()
fig.add_trace(go.Scatter(
    x=daily_margin.index,
    y=daily_margin.values,
    fill="tozeroy",
    line=dict(color="#1565C0", width=1.5),
    fillcolor="rgba(21, 101, 192, 0.15)",
    name="Total Margin",
    hovertemplate="%{x|%Y-%m-%d}: $%{y:,.0f}<extra></extra>",
))
fig.update_layout(
    height=350,
    xaxis_title="Date",
    yaxis_title="Margin ($)",
    hovermode="x unified",
    showlegend=False,
)
st.plotly_chart(fig, use_container_width=True)

# ── Margin by symbol ──────────────────────────────────────────────────────────
if not sym_margin.empty:
    st.subheader("Margin by Symbol")
    palette = px.colors.qualitative.Plotly

    fig_sym = go.Figure()
    for i, sym in enumerate(sym_margin.columns):
        fig_sym.add_trace(go.Scatter(
            x=sym_margin.index,
            y=sym_margin[sym].values,
            name=sym,
            stackgroup="one",
            line=dict(width=0.5, color=palette[i % len(palette)]),
            hovertemplate=f"{sym}: $%{{y:,.0f}}<extra></extra>",
        ))
    fig_sym.update_layout(
        height=380,
        xaxis_title="Date",
        yaxis_title="Margin ($)",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    st.plotly_chart(fig_sym, use_container_width=True)

# ── Margin by sector ──────────────────────────────────────────────────────────
if not sec_margin.empty and len(sec_margin.columns) > 1:
    st.subheader("Margin by Sector")
    fig_sec = go.Figure()
    palette2 = px.colors.qualitative.Set2
    for i, sec in enumerate(sec_margin.columns):
        fig_sec.add_trace(go.Scatter(
            x=sec_margin.index,
            y=sec_margin[sec].values,
            name=sec,
            stackgroup="one",
            line=dict(width=0.5, color=palette2[i % len(palette2)]),
        ))
    fig_sec.update_layout(
        height=320,
        xaxis_title="Date",
        yaxis_title="Margin ($)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    st.plotly_chart(fig_sec, use_container_width=True)

# ── Symbol activity: days in-market ──────────────────────────────────────────
with st.expander("Symbol Activity (Days In-Market)", expanded=False):
    if not sym_margin.empty:
        in_market_days = (sym_margin > 0).sum().sort_values(ascending=False)
        total_days = len(sym_margin)
        pct_days = (in_market_days / total_days * 100).round(1)

        activity_df = pd.DataFrame({
            "Symbol":          in_market_days.index,
            "Days In-Market":  in_market_days.values,
            "% of Period":     pct_days.values,
        })
        st.dataframe(activity_df, hide_index=True, use_container_width=True)

        fig_act = px.bar(
            activity_df,
            x="Days In-Market",
            y="Symbol",
            orientation="h",
            title="Days In-Market per Symbol",
            color="% of Period",
            color_continuous_scale="Blues",
        )
        fig_act.update_layout(height=max(280, len(activity_df) * 30 + 80))
        st.plotly_chart(fig_act, use_container_width=True)

# ── Monthly peak margin heatmap ───────────────────────────────────────────────
with st.expander("Monthly Peak Margin Heatmap", expanded=False):
    monthly_peak = daily_margin.resample("ME").max()
    if len(monthly_peak) > 0:
        heat_df = pd.DataFrame({
            "year":  monthly_peak.index.year,
            "month": monthly_peak.index.month,
            "peak":  monthly_peak.values,
        })
        pivot = heat_df.pivot(index="year", columns="month", values="peak").fillna(0)
        pivot.columns = [
            ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][c - 1]
            for c in pivot.columns
        ]
        fig_heat = px.imshow(
            pivot,
            color_continuous_scale="Blues",
            text_auto=".0f",
            title="Monthly Peak Margin ($)",
            aspect="auto",
        )
        fig_heat.update_layout(height=max(200, len(pivot) * 40 + 80))
        st.plotly_chart(fig_heat, use_container_width=True)
