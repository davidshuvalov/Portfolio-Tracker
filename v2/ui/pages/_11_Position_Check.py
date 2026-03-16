"""
Position Check page — current live position status for portfolio strategies only.

Shows:
  1. Summary counts (Long / Short / Flat)
  2. Strategies grouped into LONG / SHORT / FLAT visual sections
  3. Net position by symbol
  4. Data freshness alerts
"""

from __future__ import annotations

from datetime import date

import pandas as pd
import plotly.express as px
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
    st.warning("No active strategies in portfolio.")
    st.stop()

# ── Filter in-market data to portfolio strategies only ────────────────────────
portfolio_names = {s.name for s in portfolio.strategies}

def _filter_cols(df: pd.DataFrame) -> pd.DataFrame:
    keep = [c for c in df.columns if c in portfolio_names]
    return df[keep] if keep else df.iloc[:, :0]  # empty with index

long_df  = _filter_cols(imported.in_market_long)
short_df = _filter_cols(imported.in_market_short)

# ── Determine report date (last available across both frames) ─────────────────
all_dates = []
if not long_df.empty:
    all_dates.append(long_df.index[-1])
if not short_df.empty:
    all_dates.append(short_df.index[-1])

if not all_dates:
    st.warning("No in-market position data available for portfolio strategies.")
    st.stop()

report_ts   = max(all_dates)
report_date = report_ts.date() if hasattr(report_ts, "date") else report_ts

# ── Build position table (portfolio strategies only) ──────────────────────────
position_table = get_strategy_position_table(
    long_df,
    short_df,
    portfolio.strategies,
    as_of_date=report_ts,
)

# ── Summary counts ────────────────────────────────────────────────────────────
n_long  = int((position_table["position_status"] == PositionStatus.LONG.value).sum())
n_short = int((position_table["position_status"] == PositionStatus.SHORT.value).sum())
n_flat  = int((position_table["position_status"] == PositionStatus.FLAT.value).sum())
n_total = len(position_table)

st.caption(f"As of **{report_date}** · {n_total} portfolio strateg{'y' if n_total == 1 else 'ies'}")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Portfolio Strategies", n_total)
c2.metric("🟢 Long",  n_long)
c3.metric("🔴 Short", n_short)
c4.metric("⚪ Flat",  n_flat)

st.divider()

# ── Strategy card renderer ────────────────────────────────────────────────────
def _strategy_card(row: pd.Series, bg: str, border: str) -> str:
    """Return an HTML card string for one strategy row."""
    name      = row.get("strategy", "")
    symbol    = row.get("symbol", "") or "—"
    sector    = row.get("sector", "") or "—"
    contracts = int(row.get("contracts", 1) or 1)
    last_date = row.get("last_date", "")
    pos       = row.get("position_status", "")

    pos_colours = {
        PositionStatus.LONG.value:  ("#1b5e20", "▲ LONG"),
        PositionStatus.SHORT.value: ("#b71c1c", "▼ SHORT"),
        PositionStatus.FLAT.value:  ("#555555", "● FLAT"),
    }
    text_col, pos_label = pos_colours.get(pos, ("#333", pos))

    return f"""
<div style="
    background:{bg}; border:2px solid {border}; border-radius:10px;
    padding:14px 16px; margin-bottom:8px;
">
  <div style="font-size:1.0rem; font-weight:700; color:#1a1a1a; margin-bottom:4px;">
    {name}
  </div>
  <div style="font-size:1.5rem; font-weight:800; color:{text_col}; margin-bottom:8px; letter-spacing:0.04em;">
    {pos_label}
  </div>
  <div style="font-size:0.82rem; color:#444; line-height:1.7;">
    <b>Symbol:</b> {symbol}&nbsp;&nbsp;
    <b>Sector:</b> {sector}<br>
    <b>Contracts:</b> {contracts}&nbsp;&nbsp;
    <b>Last data:</b> {last_date}
  </div>
</div>
"""

# ── Render each position group ────────────────────────────────────────────────
CARDS_PER_ROW = 3

def _render_group(
    df: pd.DataFrame,
    section_title: str,
    header_colour: str,
    bg: str,
    border: str,
) -> None:
    if df.empty:
        return
    st.markdown(
        f"<h3 style='color:{header_colour}; margin-bottom:6px;'>{section_title}</h3>",
        unsafe_allow_html=True,
    )
    rows_list = [df.iloc[i] for i in range(len(df))]
    for chunk_start in range(0, len(rows_list), CARDS_PER_ROW):
        chunk = rows_list[chunk_start : chunk_start + CARDS_PER_ROW]
        cols = st.columns(CARDS_PER_ROW)
        for col, row in zip(cols, chunk):
            col.markdown(_strategy_card(row, bg, border), unsafe_allow_html=True)


long_rows  = position_table[position_table["position_status"] == PositionStatus.LONG.value]
short_rows = position_table[position_table["position_status"] == PositionStatus.SHORT.value]
flat_rows  = position_table[position_table["position_status"] == PositionStatus.FLAT.value]

_render_group(long_rows,  f"▲ Long ({n_long})",  "#1b5e20", "#e8f5e9", "#4CAF50")
if n_long > 0 and (n_short > 0 or n_flat > 0):
    st.write("")

_render_group(short_rows, f"▼ Short ({n_short})", "#b71c1c", "#ffebee", "#F44336")
if n_short > 0 and n_flat > 0:
    st.write("")

_render_group(flat_rows,  f"● Flat ({n_flat})",   "#424242", "#fafafa", "#9E9E9E")

# ── Net position by symbol ────────────────────────────────────────────────────
st.divider()
st.subheader("Net Position by Symbol")

net_df = net_position_by_symbol(position_table)

if net_df.empty or net_df["symbol"].eq("").all():
    st.info("No symbol data available. Configure symbols on the Strategies page.")
else:
    net_df = net_df[net_df["symbol"] != ""].copy()

    NET_COLOURS = {
        PositionStatus.LONG.value:  "#c8e6c9",
        PositionStatus.SHORT.value: "#ffcdd2",
        PositionStatus.FLAT.value:  "#f5f5f5",
    }

    def _net_style(row):
        bg = NET_COLOURS.get(row.get("Net Status", ""), "")
        return [f"background-color: {bg}"] * len(row)

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

    fig = px.bar(
        net_df,
        x="symbol",
        y="net",
        title="Net Contract Position by Symbol",
        labels={"net": "Net Contracts", "symbol": "Symbol"},
        color="net",
        color_continuous_scale=["#F44336", "#EEEEEE", "#4CAF50"],
        color_continuous_midpoint=0,
    )
    fig.add_hline(y=0, line_color="black", line_width=1)
    fig.update_layout(height=300, coloraxis_showscale=False, xaxis_title="")
    st.plotly_chart(fig, use_container_width=True)

# ── Data freshness ────────────────────────────────────────────────────────────
st.divider()
st.subheader("Data Freshness")

freshness_rows = []
for strat_name in sorted(portfolio_names):
    if strat_name in imported.daily_m2m.columns:
        series = imported.daily_m2m[strat_name].dropna()
        if not series.empty:
            last_data  = series.index[-1].date()
            days_stale = (date.today() - last_data).days
        else:
            last_data, days_stale = None, 9999
    else:
        last_data, days_stale = None, 9999

    freshness_rows.append({
        "Strategy":          strat_name,
        "Last Data Date":    str(last_data) if last_data else "N/A",
        "Days Since Update": days_stale,
        "Stale?":            "⚠ Yes" if days_stale > 5 else "✓ OK",
    })

freshness_df = pd.DataFrame(freshness_rows).sort_values("Days Since Update", ascending=False)

stale_count = int((freshness_df["Stale?"] == "⚠ Yes").sum())
if stale_count:
    st.warning(f"{stale_count} strateg{'y' if stale_count == 1 else 'ies'} may have stale data (>5 trading days old).")

def _fresh_style(row):
    if row.get("Stale?", "") == "⚠ Yes":
        return ["background-color: #ffccbc"] * len(row)
    return ["background-color: #e8f5e9"] * len(row)

st.dataframe(
    freshness_df.style.apply(_fresh_style, axis=1),
    hide_index=True,
    use_container_width=True,
)
