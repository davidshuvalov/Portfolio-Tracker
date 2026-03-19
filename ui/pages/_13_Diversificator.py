"""Diversificator — mirrors T_Diversificator.bas.

Two analysis modes:
  Greedy     — sequentially adds the strategy that maximises the chosen metric,
               showing the optimal build order and incremental benefit.
  Randomised — shuffles strategy order N times; each strategy's median
               contribution is reported, regardless of selection order.

Data source: the current portfolio (Live strategies, contract-scaled).
An optional toggle extends the candidate pool to include Paper / Pass
strategies so you can evaluate potential additions.
"""

from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.analytics.diversification import (
    SortMetric,
    run_greedy_selection,
    run_randomized_analysis,
)
from core.config import AppConfig
from core.data_types import ImportedData, PortfolioData, Strategy
from ui.strategy_labels import build_label_df, render_legend

st.set_page_config(page_title="Diversificator", layout="wide")
st.title("Diversificator")
st.caption(
    "Evaluate the sequential diversification benefit of each strategy. "
    "Greedy mode finds the optimal build order; Randomised mode shows each "
    "strategy's typical contribution across many orderings."
)

config: AppConfig = st.session_state.get("config", AppConfig.load())
portfolio: PortfolioData | None = st.session_state.get("portfolio_data")
imported: ImportedData | None = st.session_state.get("imported_data")

if portfolio is None or imported is None:
    st.info("Portfolio not built yet.")
    st.page_link("ui/pages/03_Portfolio.py", label="Go to Portfolio →")
    st.stop()

if not portfolio.strategies:
    st.warning("No active strategies in portfolio.")
    st.stop()

# ── Sidebar controls ───────────────────────────────────────────────────────────
with st.sidebar:
    st.subheader("Settings")

    mode = st.radio(
        "Analysis mode",
        ["Greedy", "Randomised"],
        help="Greedy: optimal sequential build. Randomised: average contribution across random orderings.",
    )

    metric_label = st.selectbox(
        "Optimise metric",
        ["P / Max DD (RTD)", "P / Avg DD", "Sharpe Ratio"],
        help="The metric used to rank strategies at each step.",
    )
    metric_map = {
        "P / Max DD (RTD)": "rtd",
        "P / Avg DD": "rtd_avg",
        "Sharpe Ratio": "sharpe",
    }
    sort_metric: SortMetric = metric_map[metric_label]  # type: ignore[assignment]

    n_iterations = 100
    if mode == "Randomised":
        n_iterations = st.slider(
            "Iterations",
            min_value=50,
            max_value=1000,
            value=100,
            step=50,
            help="More iterations give smoother median estimates but take longer.",
        )

    include_candidates = st.checkbox(
        "Include candidates (Paper / Pass)",
        value=False,
        help="Extend the candidate pool beyond Live strategies to evaluate "
             "what adding a candidate would do to portfolio metrics.",
    )

    run_btn = st.button("▶ Run Analysis", type="primary", use_container_width=True)

# ── Build daily PnL matrix ─────────────────────────────────────────────────────
# Start with portfolio live strategies (contract-scaled)
pnl_df = portfolio.daily_pnl.copy()

if include_candidates:
    candidate_statuses = {"Paper", "Pass", "Incubating"}
    for s in (imported.strategies or []):
        if s.status in candidate_statuses and s.name not in pnl_df.columns:
            if s.name in imported.daily_m2m.columns:
                raw = imported.daily_m2m[s.name]
                pnl_df[s.name] = raw * (s.contracts or 1)

# Strategy labels for display (only Live strategies get S-codes; candidates stay as-is)
live_strategies = portfolio.strategies
label_map = {s.name: f"S{i+1}" for i, s in enumerate(live_strategies)}

# ── Run analysis ───────────────────────────────────────────────────────────────
if run_btn or st.session_state.get("_diversificator_ran"):

    if run_btn:
        with st.spinner(f"Running {mode} analysis…"):
            if mode == "Greedy":
                greedy_rows = run_greedy_selection(pnl_df, sort_metric=sort_metric)
                st.session_state["_diversificator_greedy"] = greedy_rows
                st.session_state["_diversificator_metric"] = metric_label
            else:
                rand_df = run_randomized_analysis(
                    pnl_df,
                    n_iterations=n_iterations,
                    sort_metric=sort_metric,
                )
                st.session_state["_diversificator_rand"] = rand_df
                st.session_state["_diversificator_metric"] = metric_label

        st.session_state["_diversificator_ran"] = True
        st.session_state["_diversificator_mode"] = mode

    display_mode = st.session_state.get("_diversificator_mode", mode)
    display_metric = st.session_state.get("_diversificator_metric", metric_label)

    # ── Greedy results ─────────────────────────────────────────────────────────
    if display_mode == "Greedy":
        rows = st.session_state.get("_diversificator_greedy", [])
        if not rows:
            st.warning("No results — check that portfolio has strategies.")
            st.stop()

        df = pd.DataFrame(rows)

        st.subheader(f"Greedy Build Order — optimising {display_metric}")
        st.caption(
            "At each step the strategy added is the one that maximises the "
            f"chosen metric ({display_metric}) for the combined portfolio so far."
        )

        # Metric improvement chart
        metric_col = {"P / Max DD (RTD)": "rtd", "P / Avg DD": "rtd_avg", "Sharpe Ratio": "sharpe"}.get(
            display_metric, "rtd"
        )
        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=df["step"],
                y=df[metric_col],
                mode="lines+markers+text",
                text=df["strategy_added"].apply(lambda n: label_map.get(n, n)),
                textposition="top center",
                line=dict(color="#1565C0", width=2),
                marker=dict(size=8),
                name=display_metric,
                hovertemplate=(
                    "Step %{x}<br>"
                    + display_metric + ": %{y:.2f}<br>"
                    + "%{customdata}<extra></extra>"
                ),
                customdata=df["strategy_added"],
            )
        )
        fig.update_layout(
            height=360,
            xaxis_title="Strategies in portfolio",
            yaxis_title=display_metric,
            xaxis=dict(dtick=1),
            hovermode="x unified",
        )
        st.plotly_chart(fig, use_container_width=True)

        # Detailed table
        display_df = df[["step", "strategy_added", "annual_profit", "max_dd", "avg_dd", "sharpe", "rtd", "rtd_avg"]].copy()
        display_df.columns = [
            "Step", "Strategy Added",
            "Annual Profit ($)", "Max DD ($)", "Avg DD ($)",
            "Sharpe", "P/MaxDD", "P/AvgDD",
        ]
        for col in ("Annual Profit ($)", "Max DD ($)", "Avg DD ($)"):
            display_df[col] = display_df[col].map(lambda v: f"${v:,.0f}")
        for col in ("Sharpe", "P/MaxDD", "P/AvgDD"):
            display_df[col] = display_df[col].map(lambda v: f"{v:.2f}")

        st.dataframe(display_df, hide_index=True, use_container_width=True)

    # ── Randomised results ─────────────────────────────────────────────────────
    else:
        rand_df: pd.DataFrame = st.session_state.get("_diversificator_rand", pd.DataFrame())
        if rand_df.empty:
            st.warning("No results.")
            st.stop()

        st.subheader(f"Randomised Analysis ({n_iterations} iterations) — {display_metric}")
        st.caption(
            "Each strategy's contribution is the delta in the portfolio metric "
            "when that strategy was added, averaged across all random orderings. "
            "Higher contribution = more diversification benefit."
        )

        # Bar chart — median contribution
        fig_bar = px.bar(
            rand_df.reset_index(),
            x="strategy",
            y="median_contribution",
            color="pct_positive",
            color_continuous_scale="RdYlGn",
            range_color=[0, 100],
            labels={
                "strategy": "Strategy",
                "median_contribution": f"Median Δ {display_metric}",
                "pct_positive": "% Positive",
            },
            title=f"Median Contribution to {display_metric}",
        )
        fig_bar.update_layout(height=380, xaxis_tickangle=-40, coloraxis_colorbar_title="% +ve")
        st.plotly_chart(fig_bar, use_container_width=True)

        # Table
        display_rand = rand_df.reset_index()[
            ["strategy", "median_rank", "median_contribution", "avg_contribution", "pct_positive"]
        ].copy()
        display_rand.columns = ["Strategy", "Median Rank", "Median Δ", "Avg Δ", "% Positive"]
        for col in ("Median Δ", "Avg Δ"):
            display_rand[col] = display_rand[col].map(lambda v: f"{v:.3f}")
        display_rand["Median Rank"] = display_rand["Median Rank"].map(lambda v: f"{v:.1f}")
        display_rand["% Positive"] = display_rand["% Positive"].map(lambda v: f"{v:.0f}%")

        st.dataframe(display_rand, hide_index=True, use_container_width=True)

    # Strategy legend
    render_legend(live_strategies)

else:
    # ── Pre-run summary ────────────────────────────────────────────────────────
    n_strats = len(pnl_df.columns)
    st.info(
        f"Ready to analyse **{n_strats} strategies** "
        f"({'including candidates' if include_candidates else 'Live only'}). "
        "Configure options in the sidebar and click **▶ Run Analysis**."
    )

    label_df = build_label_df(live_strategies)
    st.dataframe(label_df, hide_index=True, use_container_width=True)
