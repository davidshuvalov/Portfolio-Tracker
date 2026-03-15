"""
Diversification page — portfolio composition and correlation-based
diversification analysis.

Sections:
  1. Portfolio composition breakdown (sector, symbol, type, horizon)
  2. Correlation-based diversification metrics
  3. Highly-correlated clusters (risk concentrations)
  4. Strategy overlap matrix (shared trading days)
"""

from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.analytics.correlations import (
    CorrelationMode,
    compute_correlation_matrix,
    get_correlation_pairs,
)
from core.config import AppConfig
from core.data_types import PortfolioData, Strategy

st.set_page_config(page_title="Diversification", layout="wide")
st.title("Diversification")

config: AppConfig = st.session_state.get("config", AppConfig.load())
portfolio: PortfolioData | None = st.session_state.get("portfolio_data")

if portfolio is None:
    st.info("Portfolio not built yet.")
    st.page_link("ui/pages/03_Portfolio.py", label="Go to Portfolio →")
    st.stop()

if not portfolio.strategies:
    st.warning("No live strategies in portfolio.")
    st.stop()

strategies = portfolio.strategies


# ── Helper: build breakdown DataFrame ────────────────────────────────────────

def _build_breakdown(field: str) -> pd.DataFrame:
    values = [getattr(s, field, "") or "—" for s in strategies]
    df = pd.DataFrame({"strategy": [s.name for s in strategies], field: values})
    return df.groupby(field).size().reset_index(name="count").sort_values("count", ascending=False)


# ── Section 1: Composition ────────────────────────────────────────────────────
st.subheader("Portfolio Composition")

fields = {
    "sector":    "Sector",
    "symbol":    "Symbol",
    "type":      "Strategy Type",
    "horizon":   "Horizon",
    "timeframe": "Timeframe",
}

# Only show breakdowns that have variety
active_fields = {
    k: v for k, v in fields.items()
    if len({getattr(s, k, "") or "—" for s in strategies}) > 1
    or len(strategies) <= 3
}

if active_fields:
    cols = st.columns(min(len(active_fields), 3))
    for col, (field, label) in zip(cols * 10, active_fields.items()):
        df = _build_breakdown(field)
        if df.empty:
            continue
        fig = px.pie(
            df,
            names=field,
            values="count",
            title=label,
            hole=0.35,
        )
        fig.update_traces(textposition="inside", textinfo="percent+label")
        fig.update_layout(height=280, showlegend=False, margin=dict(t=40, b=10, l=10, r=10))
        col.plotly_chart(fig, use_container_width=True)
else:
    st.info("All strategies share the same sector/symbol/type — no composition breakdown available.")

st.divider()

# ── Section 2: Correlation diversification metrics ────────────────────────────
st.subheader("Correlation Diversification")

if len(strategies) >= 2:
    with st.spinner("Computing correlation matrices…"):
        # Use cached matrices if available, otherwise compute normal mode
        corr_cache = st.session_state.get("corr_matrices")
        if corr_cache is not None:
            normal_matrix = corr_cache["normal"]
            neg_matrix = corr_cache["negative"]
            dd_matrix = corr_cache["drawdown"]
        else:
            normal_matrix = compute_correlation_matrix(
                portfolio.daily_pnl, CorrelationMode.NORMAL
            )
            neg_matrix = compute_correlation_matrix(
                portfolio.daily_pnl, CorrelationMode.NEGATIVE
            )
            dd_matrix = compute_correlation_matrix(
                portfolio.daily_pnl, CorrelationMode.DRAWDOWN
            )

    def _off_diag(m: pd.DataFrame) -> np.ndarray:
        n = len(m)
        return np.array([m.iloc[i, j] for i in range(n) for j in range(i + 1, n) if not np.isnan(m.iloc[i, j])])

    norm_vals = _off_diag(normal_matrix)
    neg_vals = _off_diag(neg_matrix)
    dd_vals = _off_diag(dd_matrix)

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Avg Normal Corr",   f"{np.mean(norm_vals):.3f}" if len(norm_vals) > 0 else "N/A")
    c2.metric("Avg Negative Corr", f"{np.mean(neg_vals):.3f}"  if len(neg_vals)  > 0 else "N/A")
    c3.metric("Avg Drawdown Corr", f"{np.mean(dd_vals):.3f}"   if len(dd_vals)   > 0 else "N/A")

    n_high_normal = int(np.sum(norm_vals >= config.corr_normal_threshold))
    n_high_dd     = int(np.sum(dd_vals   >= config.corr_drawdown_threshold))
    c4.metric(f"Pairs ≥ {config.corr_normal_threshold:.0%} (Normal)", n_high_normal)
    c5.metric(f"Pairs ≥ {config.corr_drawdown_threshold:.0%} (Drawdown)", n_high_dd)

    # Distribution of pairwise correlations
    if len(norm_vals) > 0:
        st.divider()
        st.subheader("Pairwise Correlation Distribution")
        col_left, col_right = st.columns(2)

        with col_left:
            fig = px.histogram(
                x=norm_vals, nbins=30,
                title="Normal Correlation Distribution",
                labels={"x": "Correlation", "count": "Pairs"},
                color_discrete_sequence=["#2196F3"],
            )
            fig.add_vline(x=float(np.mean(norm_vals)), line_dash="dash",
                          line_color="orange", annotation_text="Mean")
            fig.add_vline(x=config.corr_normal_threshold, line_color="red",
                          line_dash="dot", annotation_text="Threshold")
            fig.update_layout(height=320, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        with col_right:
            fig2 = px.histogram(
                x=dd_vals, nbins=30,
                title="Drawdown Correlation Distribution",
                labels={"x": "Correlation", "count": "Pairs"},
                color_discrete_sequence=["#9C27B0"],
            )
            fig2.add_vline(x=float(np.mean(dd_vals)), line_dash="dash",
                           line_color="orange", annotation_text="Mean")
            fig2.add_vline(x=config.corr_drawdown_threshold, line_color="red",
                           line_dash="dot", annotation_text="Threshold")
            fig2.update_layout(height=320, showlegend=False)
            st.plotly_chart(fig2, use_container_width=True)

    # ── Highest-correlated pairs (risk concentrations) ────────────────────────
    st.divider()
    st.subheader("Highest-Correlation Pairs (Risk Concentrations)")

    all_pairs = get_correlation_pairs(normal_matrix)
    if not all_pairs.empty:
        top_n = min(20, len(all_pairs))
        top_pairs = all_pairs.head(top_n).copy()
        top_pairs["correlation"] = top_pairs["correlation"].round(3)
        top_pairs["⚠ High"] = top_pairs["correlation"] >= config.corr_normal_threshold

        def _row_style(row):
            if row["correlation"] >= 0.9:
                return ["background-color: #ffcccc"] * len(row)
            if row["correlation"] >= config.corr_normal_threshold:
                return ["background-color: #fff3cd"] * len(row)
            return [""] * len(row)

        st.dataframe(
            top_pairs.style.apply(_row_style, axis=1),
            hide_index=True,
            use_container_width=True,
            height=min(400, top_n * 36 + 40),
        )

    # ── Redundant strategies ──────────────────────────────────────────────────
    high_pairs = all_pairs[all_pairs["correlation"] >= config.corr_normal_threshold]
    if not high_pairs.empty:
        from collections import Counter
        strat_appearances = Counter(
            list(high_pairs["strategy_a"]) + list(high_pairs["strategy_b"])
        )
        st.subheader("Strategies Appearing in Most High-Correlation Pairs")
        redund = pd.DataFrame(
            strat_appearances.most_common(),
            columns=["Strategy", "High-Corr Pair Count"],
        )
        st.dataframe(redund, hide_index=True, use_container_width=True)

else:
    st.warning("Need at least 2 live strategies to compute diversification metrics.")

# ── Section 3: Activity overlap ───────────────────────────────────────────────
with st.expander("Trading Day Overlap Matrix", expanded=False):
    st.caption(
        "Fraction of days where both strategies have non-zero PnL. "
        "High overlap + high correlation = redundant strategies."
    )
    cols_avail = list(portfolio.daily_pnl.columns)
    active = portfolio.daily_pnl != 0

    n = len(cols_avail)
    overlap = np.zeros((n, n))
    for i in range(n):
        for j in range(n):
            both_active = active.iloc[:, i] & active.iloc[:, j]
            either_active = active.iloc[:, i] | active.iloc[:, j]
            denom = either_active.sum()
            overlap[i, j] = both_active.sum() / denom if denom > 0 else 0.0

    overlap_df = pd.DataFrame(overlap, index=cols_avail, columns=cols_avail)
    fig3 = go.Figure(go.Heatmap(
        z=overlap_df.values,
        x=cols_avail,
        y=cols_avail,
        colorscale="Blues",
        zmin=0.0, zmax=1.0,
        text=np.round(overlap_df.values, 2),
        texttemplate="%{text}",
        textfont={"size": 9},
        colorbar=dict(title="Overlap"),
        hovertemplate="%{y} ∩ %{x}: %{z:.2%}<extra></extra>",
    ))
    fig3.update_layout(
        height=max(350, n * 35 + 80),
        xaxis=dict(tickangle=-45 if n > 8 else 0),
        margin=dict(l=10, r=10, t=20, b=10),
    )
    st.plotly_chart(fig3, use_container_width=True)
