"""
Leave-One-Out sensitivity analysis — mirrors J_LOO.bas.

Two modes:
  MC          — remove strategy, re-run Monte Carlo, report delta vs base.
  Chronological — remove strategy, replay actual history in order, report
                  realised delta metrics (no randomisation).

With Numba MC, a 20-strategy MC LOO runs in ~2s vs ~10 minutes in VBA.
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


# ── Chronological (deterministic) LOO ─────────────────────────────────────────

def _chron_stats(pnl: pd.Series) -> dict:
    """
    Compute deterministic portfolio statistics from a daily PnL series.
    Returns dict with: total_profit, annual_profit, sharpe, max_dd_pct, rtd.
    """
    if pnl.empty or pnl.abs().sum() == 0:
        return {
            "total_profit": 0.0,
            "annual_profit": 0.0,
            "sharpe": 0.0,
            "max_dd_pct": 0.0,
            "rtd": 0.0,
        }

    equity = pnl.cumsum()
    n_days = len(pnl)
    annual_factor = 252.0 / max(n_days, 1)

    annual_profit = float(equity.iloc[-1]) * annual_factor
    annual_mean = float(pnl.mean()) * 252.0
    annual_std = float(pnl.std()) * (252.0 ** 0.5)
    sharpe = annual_mean / annual_std if annual_std > 1e-9 else 0.0

    peak = equity.cummax()
    drawdown = (peak - equity)
    max_dd_abs = float(drawdown.max())
    max_dd_pct = float((drawdown / peak.clip(lower=1e-9)).max())
    rtd = annual_profit / max_dd_abs if max_dd_abs > 1e-9 else 10.0

    return {
        "total_profit": float(equity.iloc[-1]),
        "annual_profit": annual_profit,
        "sharpe": sharpe,
        "max_dd_pct": max_dd_pct,
        "rtd": rtd,
    }


def run_leave_one_out_chronological(
    portfolio: PortfolioData,
) -> pd.DataFrame:
    """
    Chronological Leave-One-Out: replay actual history without randomisation.

    For each strategy:
      1. Remove it from portfolio daily_pnl (strategies still scaled by contracts).
      2. Compute realised portfolio metrics over the shared date range.
      3. Report deltas vs the base (full) portfolio.

    Args:
        portfolio: PortfolioData with strategies scaled by contracts.

    Returns:
        DataFrame with columns:
          strategy, delta_profit, delta_annual, delta_sharpe, delta_drawdown, delta_rtd
        Sorted by delta_annual ascending (most valuable strategies first).
    """
    if not portfolio.strategies or portfolio.daily_pnl.empty:
        return pd.DataFrame(columns=[
            "strategy", "delta_profit", "delta_annual",
            "delta_sharpe", "delta_drawdown", "delta_rtd",
        ])

    base_pnl = portfolio.daily_pnl.sum(axis=1)
    base = _chron_stats(base_pnl)

    rows = []
    for strategy in portfolio.strategies:
        name = strategy.name
        if name in portfolio.daily_pnl.columns:
            reduced_pnl = base_pnl - portfolio.daily_pnl[name]
        else:
            reduced_pnl = base_pnl

        stats = _chron_stats(reduced_pnl)
        rows.append({
            "strategy":        name,
            "delta_profit":    stats["total_profit"]  - base["total_profit"],
            "delta_annual":    stats["annual_profit"] - base["annual_profit"],
            "delta_sharpe":    stats["sharpe"]        - base["sharpe"],
            "delta_drawdown":  stats["max_dd_pct"]    - base["max_dd_pct"],
            "delta_rtd":       stats["rtd"]           - base["rtd"],
            # Absolute reference for table
            "base_annual":     base["annual_profit"],
            "result_annual":   stats["annual_profit"],
        })

    df = pd.DataFrame(rows).sort_values("delta_annual", ascending=True)
    return df.reset_index(drop=True)
