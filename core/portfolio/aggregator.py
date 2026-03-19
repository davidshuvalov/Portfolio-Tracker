"""
Portfolio aggregator — mirrors J_Portfolio_Setup.bas CreatePortfolioSummary.

Takes ImportedData + strategies config (from strategies.yaml) and produces:
- portfolio_daily_m2m: daily PnL of each active strategy × contracts
- portfolio_total:     sum of all active strategy PnL × contracts (single series)
- portfolio_rows:      per-strategy portfolio metrics DataFrame (mirrors Portfolio tab)

Only strategies with status == live_status (default "Live") are included.
"""

from __future__ import annotations
from datetime import date
from typing import Any

import numpy as np
import pandas as pd

from core.data_types import ImportedData, PortfolioData, Strategy


def build_portfolio(
    imported: ImportedData,
    strategies_config: list[dict],
    summary_metrics: pd.DataFrame | None = None,
    live_status: str = "Live",
) -> PortfolioData:
    """
    Filter to active (Live) strategies, scale by contracts, build portfolio DataFrames.

    Args:
        imported:           ImportedData from csv_importer.import_all()
        strategies_config:  List of strategy dicts from strategies.yaml
        summary_metrics:    Optional summary DataFrame from summary.compute_summary()
        live_status:        The status value considered "live" (default "Live")

    Returns:
        PortfolioData with active strategies, scaled daily PnL, and portfolio metrics.
    """
    # Build lookup: name → config dict
    config_map: dict[str, dict] = {s["name"]: s for s in strategies_config}

    # Filter to live strategies that also have imported data
    available = set(imported.strategy_names)
    active_strategies: list[Strategy] = []

    for name in imported.strategy_names:
        cfg = config_map.get(name, {})
        status = cfg.get("status", "")
        if status != live_status:
            continue
        contracts = int(cfg.get("contracts", 1) or 1)
        active_strategies.append(Strategy(
            name=name,
            folder=imported.strategies[0].folder if imported.strategies else __import__("pathlib").Path("."),
            status=status,
            contracts=contracts,
            symbol=cfg.get("symbol", ""),
            sector=cfg.get("sector", ""),
            timeframe=cfg.get("timeframe", ""),
            type=cfg.get("type", ""),
            horizon=cfg.get("horizon", ""),
            other=cfg.get("other", ""),
            notes=cfg.get("notes", ""),
        ))

    if not active_strategies:
        # Return empty portfolio
        empty_df = pd.DataFrame(index=imported.daily_m2m.index)
        return PortfolioData(
            strategies=[],
            daily_pnl=empty_df,
            closed_trades=pd.DataFrame(columns=["strategy", "date", "position", "pnl", "mae", "mfe"]),
            summary_metrics=pd.DataFrame(),
        )

    active_names = [s.name for s in active_strategies]
    contracts_map = {s.name: s.contracts for s in active_strategies}

    # ── Scale daily PnL by contracts ──────────────────────────────────────────
    daily_pnl = imported.daily_m2m[active_names].copy()
    for name in active_names:
        daily_pnl[name] = daily_pnl[name] * contracts_map[name]

    # ── Closed trade PnL (scaled) ─────────────────────────────────────────────
    closed_pnl = imported.closed_trade_pnl[active_names].copy()
    for name in active_names:
        closed_pnl[name] = closed_pnl[name] * contracts_map[name]

    # ── Filter trades to active strategies ────────────────────────────────────
    if not imported.trades.empty:
        active_trades = imported.trades[
            imported.trades["strategy"].isin(active_names)
        ].copy()
        # Scale PnL/MAE/MFE by contracts
        for name in active_names:
            mask = active_trades["strategy"] == name
            c = contracts_map[name]
            active_trades.loc[mask, "pnl"] *= c
            active_trades.loc[mask, "mae"] *= c
            active_trades.loc[mask, "mfe"] *= c
    else:
        active_trades = imported.trades.copy()

    # ── Build portfolio metrics DataFrame ─────────────────────────────────────
    if summary_metrics is not None and not summary_metrics.empty:
        # Filter summary to active strategies and add contract-scaled columns
        avail = [n for n in active_names if n in summary_metrics.index]
        port_summary = summary_metrics.loc[avail].copy() if avail else pd.DataFrame()

        # Scale financial metrics by contracts (mirrors VBA portfolio setup)
        _scale_cols = [
            "expected_annual_profit", "actual_annual_profit",
            "annual_sd_is", "annual_sd_isoos",
            "avg_trade", "avg_profitable_trade", "avg_loss_trade",
            "largest_win", "largest_loss",
            "max_drawdown_isoos", "avg_drawdown_isoos",
            "max_drawdown_last_12_months",
            "profit_last_1_month", "profit_last_3_months", "profit_last_6_months",
            "profit_last_9_months", "profit_last_12_months", "profit_since_oos_start",
        ]
        if not port_summary.empty:
            for col in _scale_cols:
                if col in port_summary.columns:
                    for name in avail:
                        port_summary.loc[name, col] = (
                            port_summary.loc[name, col] * contracts_map.get(name, 1)
                        )
            # Add contracts column
            port_summary["contracts"] = [contracts_map.get(n, 1) for n in port_summary.index]
    else:
        port_summary = pd.DataFrame()

    return PortfolioData(
        strategies=active_strategies,
        daily_pnl=daily_pnl,
        closed_trades=active_trades,
        summary_metrics=port_summary,
    )


def portfolio_total_pnl(portfolio: PortfolioData) -> pd.Series:
    """
    Compute the total portfolio daily PnL series (sum across all active strategies).
    Returns a Series indexed by date.
    """
    if portfolio.daily_pnl.empty:
        return pd.Series(dtype=float, name="Portfolio")
    return portfolio.daily_pnl.sum(axis=1).rename("Portfolio")


def portfolio_equity_curve(portfolio: PortfolioData) -> pd.Series:
    """Cumulative equity curve of the total portfolio PnL."""
    return portfolio_total_pnl(portfolio).cumsum().rename("Equity")


def monthly_portfolio_pnl(portfolio: PortfolioData) -> pd.DataFrame:
    """
    Monthly PnL summary — both per-strategy and total.
    Returns DataFrame with index=month-end, columns=strategy names + "Total".
    """
    if portfolio.daily_pnl.empty:
        return pd.DataFrame()

    monthly = portfolio.daily_pnl.resample("ME").sum()
    monthly["Total"] = monthly.sum(axis=1)
    return monthly


def portfolio_summary_stats(portfolio: PortfolioData) -> dict[str, Any]:
    """
    Top-level portfolio stats for the dashboard header metrics.
    Mirrors key J_Portfolio metrics displayed in the Portfolio tab totals row.
    """
    if not portfolio.strategies:
        return {}

    total_pnl = portfolio_total_pnl(portfolio)
    monthly = total_pnl.resample("ME").sum()

    # Total P&L
    total_profit = float(total_pnl.sum())

    # Expected annual profit (sum × contracts from summary)
    expected_annual = 0.0
    if not portfolio.summary_metrics.empty and "expected_annual_profit" in portfolio.summary_metrics.columns:
        expected_annual = float(portfolio.summary_metrics["expected_annual_profit"].sum())

    # Drawdown on total equity curve
    equity = total_pnl.cumsum().values
    if len(equity) > 0:
        peak = np.maximum.accumulate(equity)
        dd = peak - equity
        max_dd = float(np.max(dd))
        avg_dd = float(np.mean(dd[dd > 0])) if np.any(dd > 0) else 0.0
    else:
        max_dd = avg_dd = 0.0

    # Annualised return
    n_days = len(total_pnl)
    years = n_days / 252.0 if n_days > 0 else 1.0
    annualised_profit = total_profit / years if years > 0 else 0.0

    # Win rate (monthly)
    win_rate = float((monthly > 0).mean()) if len(monthly) > 0 else 0.0

    # Last 12 months PnL
    last_12m_start = total_pnl.index[-1] - pd.DateOffset(months=12) if len(total_pnl) > 0 else None
    profit_12m = float(total_pnl.loc[total_pnl.index >= last_12m_start].sum()) if last_12m_start else 0.0

    return {
        "n_strategies": len(portfolio.strategies),
        "total_profit": total_profit,
        "annualised_profit": annualised_profit,
        "expected_annual_profit": expected_annual,
        "max_drawdown": max_dd,
        "avg_drawdown": avg_dd,
        "monthly_win_rate": win_rate,
        "profit_last_12_months": profit_12m,
    }
