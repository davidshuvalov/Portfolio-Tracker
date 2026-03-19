"""
Strategy correlation analysis — mirrors J_Correlations.bas.

Three correlation modes:
  NORMAL    — standard Pearson, filter days where ≥1 strategy has activity
  NEGATIVE  — exclude days where both strategies are simultaneously profitable
  DRAWDOWN  — correlate equity-curve drawdown series (synchronisation)

All modes use scipy.stats.pearsonr and are fully vectorised (no VBA loops).
"""

from __future__ import annotations

from enum import Enum

import numpy as np
import pandas as pd
from scipy.stats import pearsonr


class CorrelationMode(Enum):
    NORMAL = "normal"
    NEGATIVE = "negative"   # Exclude days both strategies are profitable
    DRAWDOWN = "drawdown"   # Equity curve drawdown synchronisation


# ── Core computation ──────────────────────────────────────────────────────────

def compute_correlation_matrix(
    daily_pnl: pd.DataFrame,
    mode: CorrelationMode,
    start_date: pd.Timestamp | None = None,
) -> pd.DataFrame:
    """
    Compute pairwise Pearson correlation matrix for all strategy columns.

    Args:
        daily_pnl:  DataFrame (index=DatetimeIndex, columns=strategy names)
                    with daily PnL values (M2M or closed-trade).
        mode:       CorrelationMode — controls which days are included.
        start_date: If provided, restrict to rows on or after this date.

    Returns:
        Symmetric DataFrame (n×n) with 1.0 on diagonal.
        Returns NaN for pairs with insufficient shared data (< 2 valid days).
    """
    if start_date is not None:
        daily_pnl = daily_pnl[daily_pnl.index >= start_date]
    strats = list(daily_pnl.columns)
    n = len(strats)
    matrix = np.eye(n)

    for i in range(n):
        for j in range(i + 1, n):
            a = daily_pnl.iloc[:, i].values
            b = daily_pnl.iloc[:, j].values
            corr = _pairwise_correlation(a, b, mode)
            matrix[i, j] = corr
            matrix[j, i] = corr

    return pd.DataFrame(matrix, index=strats, columns=strats)


def _pairwise_correlation(
    a: np.ndarray,
    b: np.ndarray,
    mode: CorrelationMode,
) -> float:
    """Compute correlation between two daily PnL arrays under the given mode."""
    if mode == CorrelationMode.NORMAL:
        # Include rows where at least one strategy has non-zero PnL
        mask = (a != 0) | (b != 0)
        a_use, b_use = a[mask], b[mask]

    elif mode == CorrelationMode.NEGATIVE:
        # Exclude rows where both are profitable (focus on joint losses / bad days)
        mask = ~((a > 0) & (b > 0))
        a_use, b_use = a[mask], b[mask]

    else:  # DRAWDOWN
        # Convert cumulative equity to drawdown fraction, use all rows
        a_eq = np.cumsum(a)
        b_eq = np.cumsum(b)
        a_use = _to_drawdown_series(a_eq)
        b_use = _to_drawdown_series(b_eq)

    if len(a_use) < 2:
        return float("nan")

    # Constant arrays → no correlation defined
    if np.std(a_use) < 1e-12 or np.std(b_use) < 1e-12:
        return 0.0

    corr, _ = pearsonr(a_use, b_use)
    return float(corr)


def _to_drawdown_series(equity: np.ndarray) -> np.ndarray:
    """
    Convert equity curve to fractional peak-to-trough drawdown series.
    Values in [0, 1] — 0 = at peak, 1 = total loss.
    """
    peak = np.maximum.accumulate(equity)
    return np.where(peak > 1e-9, (peak - equity) / peak, 0.0)


# ── Analysis helpers ──────────────────────────────────────────────────────────

def get_correlation_pairs(
    matrix: pd.DataFrame,
) -> pd.DataFrame:
    """
    Return all unique strategy pairs with their correlation value,
    sorted descending by |correlation|.

    Columns: strategy_a, strategy_b, correlation
    """
    strats = list(matrix.columns)
    rows = []
    for i in range(len(strats)):
        for j in range(i + 1, len(strats)):
            rows.append({
                "strategy_a": strats[i],
                "strategy_b": strats[j],
                "correlation": float(matrix.iloc[i, j]),
            })
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.sort_values("correlation", ascending=False).reset_index(drop=True)


def flag_high_correlations(
    matrix: pd.DataFrame,
    threshold: float,
) -> list[tuple[str, str, float]]:
    """
    Return pairs whose |correlation| exceeds the threshold.
    Result list of (strategy_a, strategy_b, correlation), sorted descending.
    """
    pairs = get_correlation_pairs(matrix)
    if pairs.empty:
        return []
    high = pairs[pairs["correlation"].abs() >= threshold]
    return list(zip(high["strategy_a"], high["strategy_b"], high["correlation"]))


def average_correlation(matrix: pd.DataFrame) -> float:
    """
    Mean of all unique off-diagonal correlation values.
    Returns NaN for a single-strategy matrix.
    """
    n = len(matrix)
    if n < 2:
        return float("nan")
    values = [
        matrix.iloc[i, j]
        for i in range(n)
        for j in range(i + 1, n)
        if not np.isnan(matrix.iloc[i, j])
    ]
    return float(np.mean(values)) if values else float("nan")


def compute_all_modes(
    daily_pnl: pd.DataFrame,
    start_date: pd.Timestamp | None = None,
) -> dict[str, pd.DataFrame]:
    """
    Convenience: compute all three correlation matrices at once.
    Returns {"normal": df, "negative": df, "drawdown": df}.

    Args:
        daily_pnl:  DataFrame (index=DatetimeIndex, columns=strategy names).
        start_date: If provided, restrict to rows on or after this date
                    (mirrors VBA Correl_Short_Period / Correl_Long_Period).
    """
    return {
        "normal":   compute_correlation_matrix(daily_pnl, CorrelationMode.NORMAL, start_date),
        "negative": compute_correlation_matrix(daily_pnl, CorrelationMode.NEGATIVE, start_date),
        "drawdown": compute_correlation_matrix(daily_pnl, CorrelationMode.DRAWDOWN, start_date),
    }
