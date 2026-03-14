"""
Leave-One-Out sensitivity analysis — mirrors J_LOO.bas.

For each strategy in the portfolio, temporarily remove it and re-run
Monte Carlo. The delta metrics show how much each strategy contributes
to (or detracts from) portfolio performance.

With Numba MC, a 20-strategy LOO runs in ~2s vs ~10 minutes in VBA.
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd

from core.analytics.monte_carlo import run_monte_carlo
from core.config import MCConfig
from core.data_types import MCResult, PortfolioData, Strategy
from core.portfolio.aggregator import portfolio_total_pnl


# ── Portfolio manipulation ────────────────────────────────────────────────────

def _remove_strategy(portfolio: PortfolioData, name: str) -> PortfolioData:
    """Return a copy of PortfolioData with the named strategy excluded."""
    strategies = [s for s in portfolio.strategies if s.name != name]
    cols = [s.name for s in strategies]

    daily_pnl = portfolio.daily_pnl[
        [c for c in cols if c in portfolio.daily_pnl.columns]
    ].copy()

    closed_trades = portfolio.closed_trades.copy()
    if not closed_trades.empty and "strategy" in closed_trades.columns:
        closed_trades = closed_trades[closed_trades["strategy"] != name]

    summary = portfolio.summary_metrics
    if not summary.empty:
        summary = summary.drop(index=name, errors="ignore")

    return PortfolioData(
        strategies=strategies,
        daily_pnl=daily_pnl,
        closed_trades=closed_trades,
        summary_metrics=summary,
    )


def _analyse_portfolio(
    portfolio: PortfolioData,
    mc_config: MCConfig,
    margin_threshold: float,
) -> MCResult:
    """
    Run Monte Carlo on the portfolio total PnL.
    Uses IS+OOS period (all data) — period filtering handled by caller if needed.
    """
    total_pnl = portfolio_total_pnl(portfolio)
    if total_pnl.empty:
        return MCResult(
            starting_equity=0.0,
            expected_profit=0.0,
            risk_of_ruin=float("nan"),
            max_drawdown_pct=0.0,
            sharpe_ratio=0.0,
            return_to_drawdown=0.0,
        )

    # LOO always uses IS+OOS (full history) — no per-strategy OOS date filtering
    loo_config = MCConfig(
        simulations=mc_config.simulations,
        period="IS+OOS",
        risk_ruin_target=mc_config.risk_ruin_target,
        risk_ruin_tolerance=mc_config.risk_ruin_tolerance,
        trade_adjustment=mc_config.trade_adjustment,
        trade_option=mc_config.trade_option,
    )
    return run_monte_carlo(
        daily_m2m=total_pnl,
        config=loo_config,
        margin_threshold=margin_threshold,
        return_scenarios=False,
    )


# ── Main LOO function ─────────────────────────────────────────────────────────

def run_leave_one_out(
    portfolio: PortfolioData,
    mc_config: MCConfig,
    margin_threshold: float,
) -> pd.DataFrame:
    """
    Run Leave-One-Out sensitivity analysis.

    For each strategy:
      1. Remove it from the portfolio.
      2. Re-run Monte Carlo on the remaining strategies.
      3. Compute delta vs base portfolio.

    Args:
        portfolio:        PortfolioData (active Live strategies, scaled by contracts).
        mc_config:        MCConfig controlling simulations / trade_adjustment etc.
        margin_threshold: Dollar ruin threshold passed to MC solver.

    Returns:
        DataFrame with columns:
          strategy, delta_profit, delta_sharpe, delta_drawdown, delta_rtd
          delta_ror (delta risk-of-ruin)

        Sorted by delta_profit ascending — strategies whose removal hurts
        most appear first (most valuable to portfolio).
    """
    if not portfolio.strategies:
        return pd.DataFrame(columns=[
            "strategy", "delta_profit", "delta_sharpe",
            "delta_drawdown", "delta_rtd", "delta_ror",
        ])

    base = _analyse_portfolio(portfolio, mc_config, margin_threshold)

    rows = []
    for strategy in portfolio.strategies:
        reduced = _remove_strategy(portfolio, strategy.name)
        result = _analyse_portfolio(reduced, mc_config, margin_threshold)

        rows.append({
            "strategy":       strategy.name,
            "delta_profit":   result.expected_profit - base.expected_profit,
            "delta_sharpe":   result.sharpe_ratio - base.sharpe_ratio,
            "delta_drawdown": result.max_drawdown_pct - base.max_drawdown_pct,
            "delta_rtd":      result.return_to_drawdown - base.return_to_drawdown,
            "delta_ror":      result.risk_of_ruin - base.risk_of_ruin,
            # Absolute values for table display
            "base_profit":    base.expected_profit,
            "result_profit":  result.expected_profit,
            "base_sharpe":    base.sharpe_ratio,
            "result_sharpe":  result.sharpe_ratio,
        })

    df = pd.DataFrame(rows).sort_values("delta_profit", ascending=True)
    return df.reset_index(drop=True)
