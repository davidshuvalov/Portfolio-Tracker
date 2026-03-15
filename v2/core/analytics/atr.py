"""
ATR computation from trade MFE and MAE data.

No OHLC price feed needed. For each trade, the dollar range for that day
(holding 1 contract) is:

    daily_range = abs(MFE) + abs(MAE)

where MFE = Maximum Favourable Excursion and MAE = Maximum Adverse Excursion,
both stored as positive dollar values in TradeData.csv.

Rolling mean of daily_range over N trading days gives the dollar ATR.
ATR is then used for volatility-adjusted contract sizing:

    dollar_risk  = atr * atr_ratio + margin * (1 - atr_ratio)
    contracts    = floor(equity * pct_equity / dollar_risk)

Mirrors VBA Portfolio Contract Settings (Estimated Vol Contract Sizing).
"""

from __future__ import annotations

import math

import pandas as pd


# ── Window definitions ────────────────────────────────────────────────────────
ATR_WINDOWS: dict[str, int] = {
    "ATR Last 3 Months":  63,   # ~3 calendar months of trading days
    "ATR Last 6 Months":  126,
    "ATR Last 12 Months": 252,
}


def compute_daily_range(trades_df: pd.DataFrame) -> pd.DataFrame:
    """
    Compute the dollar daily range per strategy from trade MAE/MFE.

    daily_range = abs(MFE) + abs(MAE)  per trade, then summed per (strategy, date).

    Args:
        trades_df: DataFrame with columns [strategy, date, pnl, mae, mfe].
                   mae/mfe are dollar values for 1 contract.

    Returns:
        DataFrame with DatetimeIndex, columns = strategy names,
        values = summed daily range in dollars. Missing dates/strategies = 0.
    """
    if trades_df is None or trades_df.empty:
        return pd.DataFrame()

    df = trades_df.copy()
    df["_range"] = df["mfe"].abs() + df["mae"].abs()

    grouped = (
        df.groupby(["strategy", "date"])["_range"]
        .sum()
        .reset_index()
    )

    pivoted = grouped.pivot(index="date", columns="strategy", values="_range")
    pivoted.index = pd.DatetimeIndex(pivoted.index)
    pivoted.columns.name = None
    pivoted = pivoted.sort_index().fillna(0.0)

    return pivoted


def compute_atr_series(
    trades_df: pd.DataFrame,
    window: str = "ATR Last 3 Months",
) -> pd.DataFrame:
    """
    Compute rolling dollar ATR time series per strategy.

    Args:
        trades_df: DataFrame with columns [strategy, date, pnl, mae, mfe].
        window:    One of ATR_WINDOWS keys.

    Returns:
        DataFrame(index=DatetimeIndex, columns=strategy names) of rolling ATR.
    """
    if trades_df is None or trades_df.empty:
        return pd.DataFrame()

    window_days = ATR_WINDOWS.get(window, 63)
    daily_range = compute_daily_range(trades_df)

    if daily_range.empty:
        return pd.DataFrame()

    return daily_range.rolling(window=window_days, min_periods=1).mean()


def compute_atr(
    trades_df: pd.DataFrame,
    window: str = "ATR Last 3 Months",
) -> pd.Series:
    """
    Return the current (most recent) dollar ATR per strategy.

    Args:
        trades_df: DataFrame with columns [strategy, date, pnl, mae, mfe].
        window:    ATR window label from ATR_WINDOWS.

    Returns:
        pd.Series indexed by strategy name with latest ATR in dollars.
        Strategies with no trade data return 0.0.
    """
    series = compute_atr_series(trades_df, window)
    if series.empty:
        return pd.Series(dtype=float)
    return series.iloc[-1]


def contract_size_from_atr(
    equity: float,
    contract_size_pct: float,
    atr_dollars: float,
    margin: float,
    ratio: float,
) -> int:
    """
    Compute contract size using a blended ATR/margin dollar-risk approach.

    Mirrors VBA Portfolio Contract Settings:
        dollar_risk = atr_dollars * ratio + margin * (1 - ratio)
        contracts   = floor(equity * contract_size_pct / dollar_risk)

    Falls back to margin-only sizing when ATR is zero or not available.
    Always returns at least 1 contract.

    Args:
        equity:            Portfolio equity in dollars.
        contract_size_pct: Fraction of equity to risk per contract (e.g. 0.01).
        atr_dollars:       Dollar ATR for this strategy at 1-contract basis.
        margin:            Margin requirement per contract in dollars.
        ratio:             ATR weight in the blend (0 = pure margin, 1 = pure ATR).

    Returns:
        Number of contracts (int ≥ 1).
    """
    if not (0.0 <= ratio <= 1.0):
        raise ValueError(f"ratio must be between 0 and 1, got {ratio}")

    effective_risk = abs(atr_dollars) * ratio + abs(margin) * (1.0 - ratio)

    if effective_risk <= 0:
        return 1

    raw = (equity * contract_size_pct) / effective_risk
    return max(1, math.floor(raw))


def reweight_contracts_by_atr(
    base_contracts: pd.Series,
    atr_series: pd.DataFrame,
    current_atr: pd.Series,
) -> pd.DataFrame:
    """
    Scale historical contract counts by (current_ATR / historical_ATR) at each date.

    Mirrors VBA "re-weight portfolio backtest contracts on historical ATR".
    At each date t: adj_contracts(t) = floor(base * current_atr / atr_t)

    Strategies with zero base contracts return 0 throughout.
    Where historical ATR is zero, current ATR is used (ratio = 1).

    Args:
        base_contracts: pd.Series of base contract counts per strategy.
        atr_series:     DataFrame(index=dates, columns=strategies) of rolling ATR.
        current_atr:    pd.Series of current ATR values per strategy.

    Returns:
        DataFrame(index=dates, columns=strategies) of reweighted contract counts (int).
    """
    if atr_series.empty:
        return pd.DataFrame()

    strategies = atr_series.columns
    result = pd.DataFrame(index=atr_series.index, columns=strategies, dtype=float)

    for strat in strategies:
        base = int(base_contracts.get(strat, 1))
        if base == 0:
            result[strat] = 0
            continue

        curr = float(current_atr.get(strat, 0.0))
        if curr <= 0:
            # No current ATR: keep base contracts unchanged
            result[strat] = float(base)
            continue

        hist = atr_series[strat].copy()
        # Where history is zero, fall back to current (ratio = 1 → no rescaling)
        hist = hist.where(hist > 0, curr)

        raw = base * (curr / hist)
        result[strat] = raw.apply(lambda x: float(max(1, math.floor(x))))

    return result.astype(float)
