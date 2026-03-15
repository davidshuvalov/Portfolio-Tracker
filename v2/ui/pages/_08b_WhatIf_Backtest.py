"""
What-If Backtest — pick any strategies, set contracts, choose a date range.

Unlike the Portfolio Backtest page (which uses the live portfolio as-is),
this page lets you freely compose a hypothetical portfolio:
  1. Check/uncheck any imported strategy
  2. Override contract count per strategy
  3. Set a date range (presets or custom)
  4. See the combined backtest result
"""

from __future__ import annotations

from datetime import date, timedelta

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.config import AppConfig
from core.data_types import ImportedData

st.set_page_config(page_title="What-If Backtest", layout="wide")

# ── Sidebar workflow status ────────────────────────────────────────────────────
try:
    from ui.workflow import render_workflow_sidebar
    with st.sidebar:
        render_workflow_sidebar()
except Exception:
    pass

st.title("What-If Backtest")
st.caption("Compose a hypothetical portfolio — any strategies, any contract counts, any date range.")

config: AppConfig = st.session_state.get("config", AppConfig.load())
imported: ImportedData | None = st.session_state.get("imported_data")

if imported is None:
    st.info("No data imported yet.")
    st.page_link("ui/pages/01_Import.py", label="Go to Import →")
    st.stop()

all_names: list[str] = imported.strategy_names
if not all_names:
    st.warning("No strategies found in imported data.")
    st.stop()

# Build a lookup of current status + contracts from imported strategies list
_strat_meta: dict[str, dict] = {}
for s in imported.strategies:
    _strat_meta[s.name] = {"status": s.status, "contracts": s.contracts}

# ── Sidebar — strategy + contract controls ────────────────────────────────────
with st.sidebar:
    st.header("Strategies")

    # Bulk toggle buttons
    col_a, col_b = st.columns(2)
    if col_a.button("Select All", use_container_width=True):
        for n in all_names:
            st.session_state[f"wi_include_{n}"] = True
    if col_b.button("Deselect All", use_container_width=True):
        for n in all_names:
            st.session_state[f"wi_include_{n}"] = False

    st.divider()

    selected_names: list[str] = []
    contracts_override: dict[str, int] = {}

    for name in all_names:
        meta = _strat_meta.get(name, {})
        default_on = meta.get("status", "") == "Live"
        default_contracts = int(meta.get("contracts", 1) or 1)

        included = st.checkbox(
            name,
            value=st.session_state.get(f"wi_include_{name}", default_on),
            key=f"wi_include_{name}",
        )
        if included:
            c = st.number_input(
                f"Contracts — {name}",
                min_value=1,
                max_value=100,
                value=st.session_state.get(f"wi_contracts_{name}", default_contracts),
                step=1,
                key=f"wi_contracts_{name}",
                label_visibility="collapsed",
            )
            selected_names.append(name)
            contracts_override[name] = int(c)

    st.divider()
    st.header("Period")

    period_preset = st.selectbox(
        "Preset",
        ["All Data", "Last 1 Year", "Last 2 Years", "Last 3 Years", "Last 5 Years", "Custom"],
    )

    date_min = imported.daily_m2m.index.min().date()
    date_max = imported.daily_m2m.index.max().date()

    if period_preset == "Custom":
        start_date = st.date_input("Start", value=date_min, min_value=date_min, max_value=date_max)
        end_date   = st.date_input("End",   value=date_max, min_value=date_min, max_value=date_max)
    else:
        end_date = date_max
        years_map = {"All Data": None, "Last 1 Year": 1, "Last 2 Years": 2,
                     "Last 3 Years": 3, "Last 5 Years": 5}
        yrs = years_map[period_preset]
        start_date = (
            date_max - timedelta(days=int(yrs * 365.25)) if yrs else date_min
        )

    st.divider()
    show_individual = st.checkbox("Show individual strategies on chart", value=True)
    normalize       = st.checkbox("Normalise curves to 0 start", value=False)

# ── Guard: nothing selected ───────────────────────────────────────────────────
if not selected_names:
    st.warning("No strategies selected. Use the sidebar to select at least one strategy.")
    st.stop()

# ── Build what-if daily PnL matrix ────────────────────────────────────────────
raw_pnl = imported.daily_m2m[selected_names].copy()
for name in selected_names:
    raw_pnl[name] = raw_pnl[name] * contracts_override[name]

# ── Filter to date window ─────────────────────────────────────────────────────
start_ts = pd.Timestamp(start_date)
end_ts   = pd.Timestamp(end_date)
window   = raw_pnl.loc[(raw_pnl.index >= start_ts) & (raw_pnl.index <= end_ts)]

if window.empty:
    st.warning("No data in the selected period.")
    st.stop()

total_pnl = window.sum(axis=1)

# ── Summary metrics ───────────────────────────────────────────────────────────
total_return = float(total_pnl.sum())
n_days       = len(total_pnl)
years        = max(n_days / 252.0, 1e-3)
ann_return   = total_return / years

equity  = total_pnl.cumsum()
peak    = equity.cummax()
dd      = peak - equity
max_dd  = float(dd.max())
avg_dd  = float(dd[dd > 0].mean()) if (dd > 0).any() else 0.0

monthly_w = total_pnl.resample("ME").sum()
win_rate  = float((monthly_w > 0).mean()) if len(monthly_w) > 0 else 0.0
std_m     = float(monthly_w.std()) if len(monthly_w) > 1 else 0.0
sharpe    = (float(monthly_w.mean()) / std_m * np.sqrt(12)) if std_m > 1e-9 else 0.0

n_strats = len(selected_names)
total_contracts = sum(contracts_override.values())

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Total Return",     f"${total_return:,.0f}")
c2.metric("Ann. Return",      f"${ann_return:,.0f}")
c3.metric("Max Drawdown",     f"${max_dd:,.0f}")
c4.metric("Avg Drawdown",     f"${avg_dd:,.0f}")
c5.metric("Monthly Win Rate", f"{win_rate:.1%}")
c6.metric("Sharpe (Monthly)", f"{sharpe:.2f}")

st.caption(
    f"Period: **{start_date}** → **{end_date}** · {n_days} trading days · "
    f"{n_strats} strateg{'y' if n_strats == 1 else 'ies'} · {total_contracts} total contracts"
)
st.divider()

# ── Equity curve ──────────────────────────────────────────────────────────────
st.subheader("Equity Curve")
fig = go.Figure()

eq_curve = equity.copy()
if normalize:
    eq_curve = eq_curve - eq_curve.iloc[0]

fig.add_trace(go.Scatter(
    x=eq_curve.index, y=eq_curve.values,
    name="What-If Portfolio",
    line=dict(color="#1565C0", width=2.5),
))

if show_individual:
    palette = px.colors.qualitative.Plotly
    for i, name in enumerate(selected_names):
        s_eq = window[name].cumsum()
        if normalize:
            s_eq = s_eq - s_eq.iloc[0]
        fig.add_trace(go.Scatter(
            x=s_eq.index, y=s_eq.values,
            name=f"{name} (×{contracts_override[name]})",
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
monthly_total = total_pnl.resample("ME").sum()

if len(monthly_total) > 0:
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
        title="Monthly PnL ($) — What-If Portfolio",
        aspect="auto",
    )
    fig_heat.update_layout(height=max(250, len(pivot) * 40 + 80))
    st.plotly_chart(fig_heat, use_container_width=True)

# ── Per-strategy contribution ─────────────────────────────────────────────────
st.subheader("Per-Strategy Contribution")

strat_totals = window.sum()
strat_ann    = strat_totals / years
strat_pct    = strat_totals / (abs(strat_totals).sum() + 1e-9) * 100

contrib_df = pd.DataFrame({
    "Strategy":        strat_totals.index,
    "Contracts":       [contracts_override[n] for n in strat_totals.index],
    "Total PnL ($)":   strat_totals.values.round(0).astype(int),
    "Ann. PnL ($)":    strat_ann.values.round(0).astype(int),
    "Contribution %":  strat_pct.values.round(1),
})
contrib_df = contrib_df.sort_values("Total PnL ($)", ascending=False)

st.dataframe(contrib_df, hide_index=True, use_container_width=True)

fig_bar = px.bar(
    contrib_df,
    x="Total PnL ($)",
    y="Strategy",
    orientation="h",
    title="Strategy Contribution",
    color="Total PnL ($)",
    color_continuous_scale=["#F44336", "#FFEB3B", "#4CAF50"],
    color_continuous_midpoint=0,
    hover_data={"Contracts": True},
)
fig_bar.update_layout(
    height=max(300, len(contrib_df) * 32 + 80),
    coloraxis_showscale=False,
)
st.plotly_chart(fig_bar, use_container_width=True)
