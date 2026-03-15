"""
Position Check page — current live position status for all active strategies.
Mirrors V_PositionCheck.bas (CreateLatestPositionsReport).

Shows:
  1. Current position for each Live strategy (Long / Short / Flat)
  2. Net position summary by symbol
  3. Stale data alerts (strategies that haven't updated recently)
"""

from __future__ import annotations

from datetime import date, timedelta

import pandas as pd
import streamlit as st

from core.analytics.margin import (
    PositionStatus,
    get_strategy_position_table,
    net_position_by_symbol,
)
from core.config import AppConfig
from core.data_types import PortfolioData

st.set_page_config(page_title="Position Check", layout="wide")
st.title("Position Check")

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

if not portfolio.strategies:
    st.warning("No live strategies in portfolio.")
    st.stop()

# ── Determine report date ─────────────────────────────────────────────────────
all_dates = []
if not imported.in_market_long.empty:
    all_dates.append(imported.in_market_long.index[-1])
if not imported.in_market_short.empty:
    all_dates.append(imported.in_market_short.index[-1])

report_ts = max(all_dates) if all_dates else None

# ── Build position table ──────────────────────────────────────────────────────
if report_ts is None:
    st.warning("No in-market position data available. Ensure EquityData CSVs were imported.")
    st.stop()

position_table = get_strategy_position_table(
    imported.in_market_long,
    imported.in_market_short,
    portfolio.strategies,
    as_of_date=report_ts,
)

report_date = report_ts.date() if hasattr(report_ts, "date") else report_ts

# ── Header ────────────────────────────────────────────────────────────────────
n_long  = int((position_table["position_status"] == PositionStatus.LONG.value).sum())
n_short = int((position_table["position_status"] == PositionStatus.SHORT.value).sum())
n_flat  = int((position_table["position_status"] == PositionStatus.FLAT.value).sum())
n_total = len(position_table)

st.caption(f"Report date: **{report_date}** · {n_total} strategies checked")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total Strategies", n_total)
c2.metric("Long",  n_long,  delta=None)
c3.metric("Short", n_short, delta=None)
c4.metric("Flat",  n_flat,  delta=None)

st.divider()

# ── Position table ────────────────────────────────────────────────────────────
st.subheader("Strategy Positions")

# Colour-code by position status
def _row_style(row):
    status = row.get("Position", "")
    if status == PositionStatus.LONG.value:
        return ["background-color: #c8e6c9"] * len(row)
    if status == PositionStatus.SHORT.value:
        return ["background-color: #ffcdd2"] * len(row)
    if status == PositionStatus.FLAT.value:
        return ["background-color: #fff9c4"] * len(row)
    return [""] * len(row)

display = position_table.rename(columns={
    "strategy":       "Strategy",
    "symbol":         "Symbol",
    "sector":         "Sector",
    "contracts":      "Contracts",
    "status":         "Live Status",
    "position_status": "Position",
    "last_date":      "Last Date",
})

# Sort: Long first, then Short, then Flat
sort_order = {PositionStatus.LONG.value: 0, PositionStatus.SHORT.value: 1, PositionStatus.FLAT.value: 2}
display["_sort"] = display["Position"].map(sort_order).fillna(3)
display = display.sort_values(["_sort", "Strategy"]).drop(columns=["_sort"])

st.dataframe(
    display.style.apply(_row_style, axis=1),
    hide_index=True,
    use_container_width=True,
    height=min(600, n_total * 38 + 50),
)

# ── Net position by symbol ────────────────────────────────────────────────────
st.divider()
st.subheader("Net Position by Symbol")

net_df = net_position_by_symbol(position_table)

if net_df.empty:
    st.info("No symbol data available. Configure symbols on the Strategies page.")
else:
    def _net_style(row):
        status = row.get("Net Status", "")
        if status == PositionStatus.LONG.value:
            return ["background-color: #c8e6c9"] * len(row)
        if status == PositionStatus.SHORT.value:
            return ["background-color: #ffcdd2"] * len(row)
        return ["background-color: #fff9c4"] * len(row)

    net_display = net_df.rename(columns={
        "symbol":      "Symbol",
        "long_count":  "Long Contracts",
        "short_count": "Short Contracts",
        "net":         "Net",
        "net_status":  "Net Status",
    })

    st.dataframe(
        net_display.style.apply(_net_style, axis=1),
        hide_index=True,
        use_container_width=True,
    )

    # Visual: net position bar chart
    import plotly.express as px
    fig = px.bar(
        net_df,
        x="symbol",
        y="net",
        title="Net Contract Position by Symbol",
        labels={"net": "Net Contracts", "symbol": "Symbol"},
        color="net",
        color_continuous_scale=["#F44336", "#FFFFFF", "#4CAF50"],
        color_continuous_midpoint=0,
    )
    fig.add_hline(y=0, line_color="black", line_width=1)
    fig.update_layout(height=320, coloraxis_showscale=False)
    st.plotly_chart(fig, use_container_width=True)

# ── Stale data alerts ─────────────────────────────────────────────────────────
st.divider()
st.subheader("Data Freshness")

# Find the most recent date per strategy in the daily_m2m data
freshness_rows = []
strat_names = set(position_table["strategy"])

for strat_name in sorted(strat_names):
    if strat_name in imported.daily_m2m.columns:
        strat_series = imported.daily_m2m[strat_name].dropna()
        if not strat_series.empty:
            last_data = strat_series.index[-1].date()
            days_stale = (date.today() - last_data).days
        else:
            last_data = None
            days_stale = 9999
    else:
        last_data = None
        days_stale = 9999

    freshness_rows.append({
        "Strategy":      strat_name,
        "Last Data Date": str(last_data) if last_data else "N/A",
        "Days Since Update": days_stale,
        "⚠ Stale":       days_stale > 5,
    })

freshness_df = pd.DataFrame(freshness_rows).sort_values("Days Since Update", ascending=False)

# Highlight stale rows
stale_count = int(freshness_df["⚠ Stale"].sum())
if stale_count > 0:
    st.warning(f"{stale_count} strategy / strategies may have stale data (>5 trading days old).")

def _stale_style(row):
    if row.get("⚠ Stale", False):
        return ["background-color: #ffccbc"] * len(row)
    return ["background-color: #e8f5e9"] * len(row)

st.dataframe(
    freshness_df.style.apply(_stale_style, axis=1),
    hide_index=True,
    use_container_width=True,
)
