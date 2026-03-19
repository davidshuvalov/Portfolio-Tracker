"""
Margin tracking and position detection — mirrors M_Margin_Tracking.bas
and V_PositionCheck.bas.

Core concepts:
  - A strategy is IN-MARKET LONG  when in_market_long[strategy] != 0
  - A strategy is IN-MARKET SHORT when in_market_short[strategy] != 0
  - Daily margin used = sum of (contracts × symbol_margin) for all in-market strategies
  - Aggregated by symbol and sector for portfolio-level views
"""

from __future__ import annotations

from enum import Enum

import numpy as np
import pandas as pd

from core.data_types import Strategy


# ── Position status ───────────────────────────────────────────────────────────

class PositionStatus(Enum):
    FLAT  = "Flat"
    LONG  = "Long"
    SHORT = "Short"


# ── Position detection ────────────────────────────────────────────────────────

def detect_positions(
    in_market_long: pd.DataFrame,
    in_market_short: pd.DataFrame,
    as_of_date: pd.Timestamp | None = None,
) -> pd.DataFrame:
    """
    Detect end-of-day position status for each strategy column.

    Args:
        in_market_long:  DataFrame (DatetimeIndex × strategy), daily in-market long PnL.
        in_market_short: DataFrame (DatetimeIndex × strategy), daily in-market short PnL.
        as_of_date:      Evaluate as of this date. Uses last available date if None.

    Returns:
        DataFrame with columns [strategy, is_long, is_short, status, as_of_date].
        `is_long` / `is_short` are bool; `status` is PositionStatus value string.
    """
    if in_market_long.empty and in_market_short.empty:
        return pd.DataFrame(columns=["strategy", "is_long", "is_short", "status", "as_of_date"])

    # Align columns
    strats = list(set(in_market_long.columns) | set(in_market_short.columns))

    # Select as_of_date row
    def _row(df: pd.DataFrame, date: pd.Timestamp) -> pd.Series:
        if df.empty:
            return pd.Series(0.0, index=df.columns)
        idx = df.index[df.index <= date]
        if idx.empty:
            return pd.Series(0.0, index=df.columns)
        return df.loc[idx[-1]]

    eff_date = as_of_date if as_of_date is not None else max(
        in_market_long.index[-1] if not in_market_long.empty else pd.Timestamp.min,
        in_market_short.index[-1] if not in_market_short.empty else pd.Timestamp.min,
    )

    long_row  = _row(in_market_long,  eff_date)
    short_row = _row(in_market_short, eff_date)

    rows = []
    for s in strats:
        long_val  = float(long_row.get(s, 0.0))
        short_val = float(short_row.get(s, 0.0))
        is_long  = long_val  != 0.0
        is_short = short_val != 0.0
        if is_long:
            status = PositionStatus.LONG.value
        elif is_short:
            status = PositionStatus.SHORT.value
        else:
            status = PositionStatus.FLAT.value
        rows.append({
            "strategy":  s,
            "is_long":   is_long,
            "is_short":  is_short,
            "status":    status,
            "as_of_date": eff_date.date() if hasattr(eff_date, "date") else eff_date,
        })

    return pd.DataFrame(rows).sort_values("strategy").reset_index(drop=True)


def get_strategy_position_table(
    in_market_long: pd.DataFrame,
    in_market_short: pd.DataFrame,
    strategies: list[Strategy],
    as_of_date: pd.Timestamp | None = None,
) -> pd.DataFrame:
    """
    Full position table with strategy metadata — mirrors VBA V_PositionCheck table.

    Returns DataFrame with columns:
      strategy, symbol, sector, contracts, status, position_status, last_date
    """
    positions = detect_positions(in_market_long, in_market_short, as_of_date)
    strat_map = {s.name: s for s in strategies}

    rows = []
    for _, pos_row in positions.iterrows():
        name = pos_row["strategy"]
        strat = strat_map.get(name)
        rows.append({
            "strategy":       name,
            "symbol":         strat.symbol    if strat else "",
            "sector":         strat.sector    if strat else "",
            "contracts":      strat.contracts if strat else 1,
            "status":         strat.status    if strat else "",
            "position_status": pos_row["status"],
            "last_date":      pos_row["as_of_date"],
        })

    return pd.DataFrame(rows)


def net_position_by_symbol(position_table: pd.DataFrame) -> pd.DataFrame:
    """
    Net contract position for each symbol: long_contracts - short_contracts.
    Mirrors VBA "Position Summary by Symbol" section.

    Returns DataFrame with columns: symbol, long_count, short_count, net, net_status.
    """
    if position_table.empty:
        return pd.DataFrame(columns=["symbol", "long_count", "short_count", "net", "net_status"])

    rows = []
    for symbol, grp in position_table.groupby("symbol"):
        if not symbol:
            continue
        contracts = grp["contracts"].fillna(1).astype(int)
        long_c  = int(contracts[grp["position_status"] == PositionStatus.LONG.value].sum())
        short_c = int(contracts[grp["position_status"] == PositionStatus.SHORT.value].sum())
        net = long_c - short_c
        if net > 0:
            net_status = PositionStatus.LONG.value
        elif net < 0:
            net_status = PositionStatus.SHORT.value
        else:
            net_status = PositionStatus.FLAT.value
        rows.append({
            "symbol":      symbol,
            "long_count":  long_c,
            "short_count": short_c,
            "net":         net,
            "net_status":  net_status,
        })

    if not rows:
        return pd.DataFrame(columns=["symbol", "long_count", "short_count", "net", "net_status"])
    return pd.DataFrame(rows).sort_values("symbol").reset_index(drop=True)


# ── Margin computation ────────────────────────────────────────────────────────

def _in_market_mask(
    in_market_long: pd.DataFrame,
    in_market_short: pd.DataFrame,
) -> pd.DataFrame:
    """
    Boolean mask: True when a strategy is in-market on a given date (long or short).
    Aligned on union of both DataFrames.
    """
    long_mask  = (in_market_long  != 0) if not in_market_long.empty  else pd.DataFrame()
    short_mask = (in_market_short != 0) if not in_market_short.empty else pd.DataFrame()

    if long_mask.empty and short_mask.empty:
        return pd.DataFrame()
    if long_mask.empty:
        return short_mask
    if short_mask.empty:
        return long_mask

    return long_mask.reindex(
        columns=short_mask.columns.union(long_mask.columns), fill_value=False
    ) | short_mask.reindex(
        columns=short_mask.columns.union(long_mask.columns), fill_value=False
    )


def compute_daily_margin(
    in_market_long: pd.DataFrame,
    in_market_short: pd.DataFrame,
    strategies: list[Strategy],
    symbol_margins: dict[str, float],
    default_margin: float = 5000.0,
) -> pd.Series:
    """
    Daily total margin utilization across all strategies.

    For each day: sum of (contracts × symbol_margin) for in-market strategies.

    Args:
        in_market_long/short: daily in-market PnL DataFrames.
        strategies:           active strategy list (provides contracts + symbol).
        symbol_margins:       symbol → $ margin per contract.
        default_margin:       fallback margin when symbol not in symbol_margins.

    Returns:
        pd.Series indexed by date, values = total margin ($).
    """
    in_market = _in_market_mask(in_market_long, in_market_short)
    if in_market.empty:
        return pd.Series(dtype=float, name="total_margin")

    strat_map = {s.name: s for s in strategies}

    # Build per-strategy margin weight: contracts × margin_per_contract
    weights: dict[str, float] = {}
    for col in in_market.columns:
        strat = strat_map.get(col)
        if strat is None:
            continue
        margin = symbol_margins.get(strat.symbol, default_margin) if strat.symbol else default_margin
        weights[col] = strat.contracts * margin

    if not weights:
        return pd.Series(0.0, index=in_market.index, name="total_margin")

    # Daily margin = dot product of in-market bool × margin weights
    relevant_cols = [c for c in weights if c in in_market.columns]
    weight_series = pd.Series({c: weights[c] for c in relevant_cols})
    daily_margin = in_market[relevant_cols].multiply(weight_series).sum(axis=1)
    daily_margin.name = "total_margin"
    return daily_margin


def margin_by_symbol(
    in_market_long: pd.DataFrame,
    in_market_short: pd.DataFrame,
    strategies: list[Strategy],
    symbol_margins: dict[str, float],
    default_margin: float = 5000.0,
) -> pd.DataFrame:
    """
    Daily margin utilization broken down by symbol.

    Returns DataFrame: index=date, columns=symbol names, values=$ margin.
    """
    in_market = _in_market_mask(in_market_long, in_market_short)
    if in_market.empty:
        return pd.DataFrame()

    strat_map = {s.name: s for s in strategies}

    # Group strategies by symbol
    symbol_groups: dict[str, list[tuple[str, float]]] = {}
    for col in in_market.columns:
        strat = strat_map.get(col)
        if strat is None:
            continue
        sym = strat.symbol or col  # fallback to strategy name if no symbol
        margin = symbol_margins.get(sym, default_margin)
        weight = strat.contracts * margin
        symbol_groups.setdefault(sym, []).append((col, weight))

    result: dict[str, pd.Series] = {}
    for sym, col_weights in symbol_groups.items():
        cols, ws = zip(*col_weights)
        w_series = pd.Series(dict(zip(cols, ws)))
        result[sym] = in_market[[c for c in cols if c in in_market.columns]].multiply(w_series).sum(axis=1)

    if not result:
        return pd.DataFrame()
    return pd.DataFrame(result).fillna(0.0)


def margin_by_sector(
    margin_by_sym: pd.DataFrame,
    strategies: list[Strategy],
    symbol_margins: dict[str, float],
) -> pd.DataFrame:
    """
    Daily margin aggregated by sector.
    Groups symbol columns by sector using strategy metadata.

    Returns DataFrame: index=date, columns=sector names, values=$ margin.
    """
    if margin_by_sym.empty:
        return pd.DataFrame()

    # Build symbol → sector map from strategies
    sym_to_sector: dict[str, str] = {}
    for s in strategies:
        if s.symbol:
            sym_to_sector[s.symbol] = s.sector or "Unknown"

    sector_groups: dict[str, list[str]] = {}
    for sym in margin_by_sym.columns:
        sector = sym_to_sector.get(sym, "Unknown")
        sector_groups.setdefault(sector, []).append(sym)

    result: dict[str, pd.Series] = {}
    for sector, syms in sector_groups.items():
        avail = [s for s in syms if s in margin_by_sym.columns]
        result[sector] = margin_by_sym[avail].sum(axis=1)

    return pd.DataFrame(result).fillna(0.0)


def margin_summary_stats(
    daily_margin: pd.Series,
    margin_by_sym: pd.DataFrame,
) -> dict:
    """
    Summary statistics for display in UI header cards.
    """
    if daily_margin.empty:
        return {}

    current = float(daily_margin.iloc[-1]) if len(daily_margin) > 0 else 0.0
    peak    = float(daily_margin.max())
    avg     = float(daily_margin.mean())
    days_at_peak = int((daily_margin == daily_margin.max()).sum())

    # Symbol with highest average margin
    if not margin_by_sym.empty:
        top_symbol = margin_by_sym.mean().idxmax()
        top_symbol_avg = float(margin_by_sym.mean().max())
    else:
        top_symbol = "N/A"
        top_symbol_avg = 0.0

    return {
        "current_margin":   current,
        "peak_margin":      peak,
        "average_margin":   avg,
        "days_at_peak":     days_at_peak,
        "top_symbol":       top_symbol,
        "top_symbol_avg":   top_symbol_avg,
    }
