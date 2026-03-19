"""
Correlations page — pairwise strategy correlation analysis (3 modes).
Mirrors the VBA Correlations tab (J_Correlations.bas).

Sections:
  1. Mode selector + correlation heatmap
  2. High-correlation alerts (threshold-based)
  3. All-pairs table sorted by correlation
  4. Average correlation summary by mode
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.analytics.correlations import (
    CorrelationMode,
    average_correlation,
    compute_all_modes,
    compute_correlation_matrix,
    flag_high_correlations,
    get_correlation_pairs,
)
from core.config import AppConfig
from core.data_types import PortfolioData
from ui.strategy_labels import build_label_map, relabel_matrix, render_legend, render_strategy_picker

st.set_page_config(page_title="Correlations", layout="wide")
st.title("Correlations")

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

if len(portfolio.strategies) < 2:
    st.warning("Need at least 2 live strategies to compute correlations.")
    st.stop()

# ── Sidebar controls ──────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Settings")

    mode_label = st.radio(
        "Correlation Mode",
        ["Normal", "Negative", "Drawdown"],
        help=(
            "**Normal** — standard Pearson on active trading days.\n\n"
            "**Negative** — exclude days both strategies are profitable "
            "(focuses on joint drawdown behaviour).\n\n"
            "**Drawdown** — correlate equity-curve drawdown series "
            "(measures timing synchronisation)."
        ),
    )
    mode_map = {
        "Normal":   CorrelationMode.NORMAL,
        "Negative": CorrelationMode.NEGATIVE,
        "Drawdown": CorrelationMode.DRAWDOWN,
    }
    mode = mode_map[mode_label]

    thresh_normal = st.slider(
        "High-corr threshold (Normal)",
        0.0, 1.0, float(config.corr_normal_threshold), 0.05,
    )
    thresh_negative = st.slider(
        "High-corr threshold (Negative)",
        0.0, 1.0, float(config.corr_negative_threshold), 0.05,
    )
    thresh_drawdown = st.slider(
        "High-corr threshold (Drawdown)",
        0.0, 1.0, float(config.corr_drawdown_threshold), 0.05,
    )

    st.divider()
    period_label = st.radio(
        "Lookback period",
        ["All Data", "1 Year", "3 Years", "5 Years", "10 Years", "Custom Range"],
        help=(
            "Restrict the correlation calculation to the most recent window. "
            "Mirrors VBA Correl_Short_Period / Correl_Long_Period named ranges."
        ),
    )
    _period_years_map = {
        "All Data": None, "1 Year": 1, "3 Years": 3, "5 Years": 5, "10 Years": 10,
    }

    # ── Custom date range slider ───────────────────────────────────────────────
    _data_start = portfolio.daily_pnl.index.min().date() if not portfolio.daily_pnl.empty else None
    _data_end   = portfolio.daily_pnl.index.max().date() if not portfolio.daily_pnl.empty else None

    if period_label == "Custom Range" and _data_start and _data_end:
        _custom_range = st.slider(
            "Date range",
            min_value=_data_start,
            max_value=_data_end,
            value=(_data_start, _data_end),
            key="corr_custom_range",
            help="Drag the handles to select the date window for correlation calculation.",
        )
        _custom_start, _custom_end = _custom_range
    else:
        _custom_start = _custom_end = None

    period_years = _period_years_map.get(period_label)

    compute_btn = st.button("Compute Correlations", type="primary", use_container_width=True)

    st.divider()
    render_strategy_picker(portfolio.strategies, key="corr_strat_picker")

threshold_map = {
    CorrelationMode.NORMAL:   thresh_normal,
    CorrelationMode.NEGATIVE: thresh_negative,
    CorrelationMode.DRAWDOWN: thresh_drawdown,
}

# ── Compute ───────────────────────────────────────────────────────────────────
corr_cache = st.session_state.get("corr_matrices")
corr_cache_period = st.session_state.get("corr_cache_period")

# Determine start_date from period window
_end = portfolio.daily_pnl.index.max() if not portfolio.daily_pnl.empty else None
if period_label == "Custom Range" and _custom_start and _custom_end:
    _start_date = pd.Timestamp(_custom_start)
    _end = pd.Timestamp(_custom_end)
elif period_years is not None and _end is not None:
    _start_date = _end - pd.DateOffset(years=period_years)
else:
    _start_date = None

# Invalidate cache when period changes
if corr_cache_period != period_label:
    corr_cache = None

if compute_btn or corr_cache is None:
    with st.spinner("Computing correlation matrices…"):
        _pnl_for_corr = portfolio.daily_pnl
        if _start_date is not None:
            _pnl_for_corr = _pnl_for_corr[_pnl_for_corr.index >= _start_date]
        if _end is not None and period_label == "Custom Range":
            _pnl_for_corr = _pnl_for_corr[_pnl_for_corr.index <= _end]
        corr_cache = compute_all_modes(_pnl_for_corr, start_date=None)
    st.session_state.corr_matrices = corr_cache
    st.session_state.corr_cache_period = period_label

matrix = corr_cache[mode.value]
threshold = threshold_map[mode]

# ── Short labels ───────────────────────────────────────────────────────────────
label_map = build_label_map(portfolio.strategies)
labeled_matrix = relabel_matrix(matrix, label_map)

# ── Heatmap ───────────────────────────────────────────────────────────────────
st.subheader(f"Correlation Matrix — {mode_label} Mode")

n = len(labeled_matrix)
# Set diagonal to NaN in z-values so cells appear white (not dark red)
# and blank out diagonal text overlay (self-correlation is trivially 1)
_vals = labeled_matrix.values.copy().astype(float)
np.fill_diagonal(_vals, float("nan"))
_text = np.where(np.eye(n, dtype=bool), "", np.round(labeled_matrix.values, 2).astype(str))
# Colour scale: green (negative) → white (zero) → red (positive)
fig = go.Figure(go.Heatmap(
    z=_vals,
    x=list(labeled_matrix.columns),
    y=list(labeled_matrix.index),
    colorscale="RdYlGn_r",
    zmin=-1.0,
    zmax=1.0,
    text=_text,
    texttemplate="%{text}",
    textfont={"size": 14 if n <= 15 else 11},
    hovertemplate="%{y} × %{x}: %{z:.3f}<extra></extra>",
    colorbar=dict(title="Correlation"),
))
fig.update_layout(
    height=max(400, n * 35 + 100),
    xaxis=dict(side="bottom", tickangle=-45 if n > 10 else 0),
    margin=dict(l=10, r=10, t=30, b=10),
)
st.plotly_chart(fig, use_container_width=True)
render_legend(portfolio.strategies)

# ── Summary row ───────────────────────────────────────────────────────────────
avg = average_correlation(matrix)
high = flag_high_correlations(matrix, threshold)

col_avg, col_high, col_pairs = st.columns(3)
col_avg.metric("Average Correlation", f"{avg:.3f}" if not np.isnan(avg) else "N/A")
col_high.metric(f"Pairs ≥ {threshold:.0%}", len(high))
col_pairs.metric("Total Pairs", n * (n - 1) // 2)

st.divider()

# ── High-correlation alerts ───────────────────────────────────────────────────
if high:
    st.subheader(f"⚠ High-Correlation Pairs (|r| ≥ {threshold:.0%})")
    alert_df = pd.DataFrame(high, columns=["Strategy A", "Strategy B", "Correlation"])
    alert_df["Strategy A"] = alert_df["Strategy A"].map(lambda n: label_map.get(n, n))
    alert_df["Strategy B"] = alert_df["Strategy B"].map(lambda n: label_map.get(n, n))
    alert_df["Correlation"] = alert_df["Correlation"].round(3)
    alert_df = alert_df.sort_values("Correlation", ascending=False)

    def _color_corr(val: float) -> str:
        if val >= 0.9:
            return "background-color: #ffcccc"
        if val >= threshold:
            return "background-color: #fff3cd"
        return ""

    st.dataframe(
        alert_df.style.map(_color_corr, subset=["Correlation"]),
        hide_index=True,
        use_container_width=True,
    )
else:
    st.success(f"No pairs exceed the {threshold:.0%} correlation threshold.")

# ── All pairs table ───────────────────────────────────────────────────────────
with st.expander("All Pairs", expanded=False):
    pairs = get_correlation_pairs(matrix)
    if not pairs.empty:
        pairs["strategy_a"] = pairs["strategy_a"].map(lambda n: label_map.get(n, n))
        pairs["strategy_b"] = pairs["strategy_b"].map(lambda n: label_map.get(n, n))
        pairs["correlation"] = pairs["correlation"].round(3)
        st.dataframe(pairs, hide_index=True, use_container_width=True)

# ── Cross-mode average summary ────────────────────────────────────────────────
with st.expander("Cross-mode Summary", expanded=False):
    summary_rows = []
    for m_label, m_enum in mode_map.items():
        m_matrix = corr_cache[m_enum.value]
        m_avg = average_correlation(m_matrix)
        m_thresh = threshold_map[m_enum]
        m_high = flag_high_correlations(m_matrix, m_thresh)
        summary_rows.append({
            "Mode": m_label,
            "Threshold": f"{m_thresh:.0%}",
            "Avg Correlation": f"{m_avg:.3f}" if not np.isnan(m_avg) else "N/A",
            f"Pairs ≥ Threshold": len(m_high),
        })
    st.dataframe(
        pd.DataFrame(summary_rows), hide_index=True, use_container_width=True
    )
