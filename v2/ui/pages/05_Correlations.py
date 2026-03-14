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

st.set_page_config(page_title="Correlations", layout="wide")
st.title("Correlations")

config: AppConfig = st.session_state.get("config", AppConfig.load())
portfolio: PortfolioData | None = st.session_state.get("portfolio_data")
imported = st.session_state.get("imported_data")

if imported is None:
    st.info("No data loaded. Go to **Import** first.")
    st.stop()

if portfolio is None:
    st.info("Portfolio not built. Go to **Portfolio** page first.")
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

    compute_btn = st.button("Compute Correlations", type="primary", use_container_width=True)

threshold_map = {
    CorrelationMode.NORMAL:   thresh_normal,
    CorrelationMode.NEGATIVE: thresh_negative,
    CorrelationMode.DRAWDOWN: thresh_drawdown,
}

# ── Compute ───────────────────────────────────────────────────────────────────
corr_cache = st.session_state.get("corr_matrices")

if compute_btn or corr_cache is None:
    with st.spinner("Computing correlation matrices…"):
        corr_cache = compute_all_modes(portfolio.daily_pnl)
    st.session_state.corr_matrices = corr_cache

matrix = corr_cache[mode.value]
threshold = threshold_map[mode]

# ── Heatmap ───────────────────────────────────────────────────────────────────
st.subheader(f"Correlation Matrix — {mode_label} Mode")

n = len(matrix)
# Colour scale: green (negative) → white (zero) → red (positive)
fig = go.Figure(go.Heatmap(
    z=matrix.values,
    x=list(matrix.columns),
    y=list(matrix.index),
    colorscale="RdYlGn_r",
    zmin=-1.0,
    zmax=1.0,
    text=np.round(matrix.values, 2),
    texttemplate="%{text}",
    textfont={"size": 10 if n <= 15 else 8},
    hovertemplate="%{y} × %{x}: %{z:.3f}<extra></extra>",
    colorbar=dict(title="Correlation"),
))
fig.update_layout(
    height=max(400, n * 35 + 100),
    xaxis=dict(side="bottom", tickangle=-45 if n > 10 else 0),
    margin=dict(l=10, r=10, t=30, b=10),
)
st.plotly_chart(fig, use_container_width=True)

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
