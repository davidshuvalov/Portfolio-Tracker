"""
Landing page shown to unauthenticated users.

Displays:
  - Hero / product pitch
  - Pricing cards
  - Login / Signup forms
"""

from __future__ import annotations

import streamlit as st

from ui.auth_ui import render_auth_forms
from ui.pricing import render_pricing_cards


_FEATURES = [
    ("📊", "Portfolio Analytics", "Aggregate equity curves, monthly P&L heatmaps, and KPI tables across all live strategies."),
    ("🎲", "Monte Carlo Simulation", "Resample thousands of trade sequences to estimate risk-of-ruin, drawdown percentiles, and profit distributions."),
    ("🔗", "Correlations & Diversification", "Pairwise strategy correlations and portfolio composition analysis by sector, symbol, and type."),
    ("🔍", "Eligibility Backtest", "Apply walk-forward rules (min Sharpe, max drawdown) to see which strategies qualified in each OOS window."),
    ("⚙️", "Portfolio Optimizer", "Auto-suggest an optimal, diversified portfolio from your eligible strategies. *(Full plan)*"),
    ("📈", "Market Analysis", "ATR, volatility regimes, and market correlations for Buy & Hold markets. *(Full plan)*"),
]


def render_landing() -> None:
    # ── Hero ──────────────────────────────────────────────────────────────────
    st.markdown(
        """
        <div style="text-align:center;padding:52px 0 36px;">
          <h1 style="font-size:2.8rem;font-weight:800;margin-bottom:12px;">
            Portfolio Tracker
          </h1>
          <p style="font-size:1.15rem;color:#94a3b8;max-width:620px;margin:0 auto 28px;">
            Professional analytics for systematic futures traders using
            MultiWalk&nbsp;Pro. Track strategies, run Monte Carlo simulations,
            optimise portfolios, and evaluate correlations — all in one place.
          </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # ── Feature grid ──────────────────────────────────────────────────────────
    cols = st.columns(3, gap="medium")
    for i, (icon, title, desc) in enumerate(_FEATURES):
        with cols[i % 3]:
            st.markdown(
                f"""
                <div style="border:1px solid #1e2d47;border-radius:10px;
                            padding:18px;margin-bottom:12px;background:#0d1626;">
                  <div style="font-size:1.6rem;">{icon}</div>
                  <div style="font-weight:700;margin:6px 0 4px;">{title}</div>
                  <div style="color:#94a3b8;font-size:0.85rem;">{desc}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    st.markdown("---")

    # ── Pricing + Auth ────────────────────────────────────────────────────────
    left, right = st.columns([1.4, 1], gap="large")

    with left:
        st.markdown("### Choose a Plan")
        render_pricing_cards(show_actions=False)

    with right:
        st.markdown("### Get Started")
        render_auth_forms()
