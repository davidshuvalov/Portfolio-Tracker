"""Diversification optimizer — mirrors T_Diversificator.bas.

Two analysis modes:
  Greedy      — sequentially adds the strategy that maximises the chosen metric.
  Randomised  — shuffles strategy order N times; reports each strategy's median
                contribution across all orderings.

All functions operate on a daily PnL DataFrame (index = DatetimeIndex,
columns = strategy names, values = dollar P&L per day).
"""

from __future__ import annotations

import random
from typing import Literal

import numpy as np
import pandas as pd

SortMetric = Literal["rtd", "rtd_avg", "sharpe"]

_TRADING_DAYS = 252


# ── Metric helpers ─────────────────────────────────────────────────────────────

def _compute_metrics(series: pd.Series) -> dict[str, float]:
    """Compute diversification metrics for a combined daily PnL series."""
    if series.empty or series.abs().sum() < 1e-9:
        return {
            "annual_profit": 0.0,
            "max_dd": 0.0,
            "avg_dd": 0.0,
            "ann_std": 0.0,
            "sharpe": 0.0,
            "rtd": 0.0,
            "rtd_avg": 0.0,
        }

    years = max(len(series) / _TRADING_DAYS, 1 / _TRADING_DAYS)
    equity = series.cumsum()
    peak = equity.cummax()
    drawdown = peak - equity  # positive = in drawdown

    annual_profit = float(series.sum() / years)
    max_dd = float(drawdown.max())
    pos_dd = drawdown[drawdown > 0]
    avg_dd = float(pos_dd.mean()) if not pos_dd.empty else 0.0
    ann_std = float(series.std() * np.sqrt(_TRADING_DAYS)) if len(series) > 1 else 0.0

    sharpe = annual_profit / ann_std if ann_std > 1e-9 else 0.0
    rtd = annual_profit / max_dd if max_dd > 1e-9 else 0.0
    rtd_avg = annual_profit / avg_dd if avg_dd > 1e-9 else 0.0

    return {
        "annual_profit": annual_profit,
        "max_dd": max_dd,
        "avg_dd": avg_dd,
        "ann_std": ann_std,
        "sharpe": sharpe,
        "rtd": rtd,
        "rtd_avg": rtd_avg,
    }


def _metric_value(m: dict[str, float], sort_metric: SortMetric) -> float:
    return m.get(sort_metric, 0.0)


# ── Greedy selection ───────────────────────────────────────────────────────────

def run_greedy_selection(
    daily_pnl: pd.DataFrame,
    sort_metric: SortMetric = "rtd",
) -> list[dict]:
    """Greedily add the strategy that best improves *sort_metric* at each step.

    Returns a list of dicts (one per step) with:
      step, strategy_added, strategies_in_portfolio,
      annual_profit, max_dd, avg_dd, sharpe, rtd, rtd_avg
    """
    strategies = list(daily_pnl.columns)
    if not strategies:
        return []

    remaining = set(strategies)
    selected: list[str] = []
    running_total: pd.Series | None = None
    results: list[dict] = []

    for step in range(1, len(strategies) + 1):
        best_name: str | None = None
        best_val = -np.inf
        best_metrics: dict[str, float] = {}

        for name in sorted(remaining):
            candidate = daily_pnl[name]
            combined = candidate if running_total is None else running_total + candidate
            m = _compute_metrics(combined)
            v = _metric_value(m, sort_metric)
            if v > best_val:
                best_val = v
                best_name = name
                best_metrics = m

        if best_name is None:
            break

        remaining.remove(best_name)
        selected.append(best_name)
        running_total = (
            daily_pnl[best_name].copy()
            if running_total is None
            else running_total + daily_pnl[best_name]
        )

        results.append(
            {
                "step": step,
                "strategy_added": best_name,
                "strategies_in_portfolio": ", ".join(selected),
                **best_metrics,
            }
        )

    return results


# ── Randomised analysis ────────────────────────────────────────────────────────

def run_randomized_analysis(
    daily_pnl: pd.DataFrame,
    n_iterations: int = 100,
    sort_metric: SortMetric = "rtd",
    seed: int | None = 42,
) -> pd.DataFrame:
    """Run N random strategy orderings and compute each strategy's median
    metric contribution.

    Returns a DataFrame indexed by strategy name with columns:
      median_rank          – median position at which this strategy was added
      median_contribution  – median delta in sort_metric when this strategy was added
      avg_contribution     – mean delta
      pct_positive         – % of iterations where contribution was positive
    """
    strategies = list(daily_pnl.columns)
    n = len(strategies)
    if n == 0:
        return pd.DataFrame()

    rng = random.Random(seed)

    # Per-strategy contribution accumulator  {name: [delta, ...]}
    contributions: dict[str, list[float]] = {s: [] for s in strategies}
    ranks: dict[str, list[int]] = {s: [] for s in strategies}

    for _ in range(n_iterations):
        order = strategies[:]
        rng.shuffle(order)

        running: pd.Series | None = None
        prev_val = 0.0

        for rank, name in enumerate(order, start=1):
            combined = daily_pnl[name] if running is None else running + daily_pnl[name]
            m = _compute_metrics(combined)
            val = _metric_value(m, sort_metric)
            delta = val - prev_val

            contributions[name].append(delta)
            ranks[name].append(rank)

            prev_val = val
            running = combined

    records = []
    for name in strategies:
        deltas = contributions[name]
        rank_list = ranks[name]
        records.append(
            {
                "strategy": name,
                "median_rank": float(np.median(rank_list)),
                "median_contribution": float(np.median(deltas)),
                "avg_contribution": float(np.mean(deltas)),
                "pct_positive": float(np.mean([d > 0 for d in deltas]) * 100),
            }
        )

    df = pd.DataFrame(records).set_index("strategy")
    return df.sort_values("median_contribution", ascending=False)
