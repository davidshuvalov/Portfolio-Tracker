"""
CSV importer — mirrors D_Import_Data.bas OptimizeDataProcessing logic.

Reads EquityData.csv and TradeData.csv files for each strategy and builds
the two primary DataFrames used by all analytics:

    daily_m2m       — dates × strategies, daily mark-to-market PnL
    closed_trade_pnl — dates × strategies, daily closed-trade PnL

EquityData.csv column layout (1-based, matching VBA dataArray indices):
    Col 1: Date
    Col 2: DailyM2MEquity
    Col 3: InMarketLong
    Col 4: InMarketShort
    Col 5: (unused)
    Col 6: ClosedTradePNL

TradeData.csv column layout (Exit rows only):
    Col 1: Date
    Col 4: Type   ("Exit" rows only)
    Col 5: Position ("Long" / "Short")
    Col 6: PNL
    Col 7: MAE
    Col 8: MFE
"""

from __future__ import annotations
from pathlib import Path
from typing import NamedTuple

import numpy as np
import pandas as pd

from core.data_types import ImportedData, Strategy, StrategyFolder
from core.ingestion.date_utils import parse_csv_date


# ── Column indices (0-based, matching VBA 1-based positions) ─────────────────
EQUITY_COL_DATE = 0
EQUITY_COL_M2M = 1
EQUITY_COL_LONG = 2
EQUITY_COL_SHORT = 3
# col index 4 unused
EQUITY_COL_CLOSED = 5

TRADE_COL_DATE = 0
TRADE_COL_TYPE = 3
TRADE_COL_POSITION = 4
TRADE_COL_PNL = 5
TRADE_COL_MAE = 6
TRADE_COL_MFE = 7


class _StrategyEquity(NamedTuple):
    name: str
    dates: list
    m2m: list[float]
    closed: list[float]
    long: list[float]
    short: list[float]


def import_all(
    strategy_folders: list[StrategyFolder],
    date_format: str = "DMY",
    use_cutoff: bool = False,
    cutoff_date=None,
    progress_cb=None,
) -> tuple[ImportedData, list[str]]:
    """
    Import all strategies' CSV data and build aligned DataFrames.

    Returns (ImportedData, warnings).

    Args:
        progress_cb: Optional callable(idx, total, name) called after each
                     strategy CSV is read, for progress-bar updates.

    Mirrors VBA OptimizeDataProcessing:
    1. Read each strategy's EquityData.csv into per-strategy date→value dicts
    2. Collect the full union of dates
    3. Build aligned matrices (zeros for missing dates)
    4. Read TradeData.csv for trade-level analysis
    """
    warnings: list[str] = []
    strategy_data: list[_StrategyEquity] = []
    total = len(strategy_folders)

    for idx, sf in enumerate(strategy_folders):
        if progress_cb is not None:
            progress_cb(idx, total, sf.name)
        equity = _read_equity_csv(sf.equity_csv, sf.name, date_format, warnings)
        if equity is None:
            continue
        strategy_data.append(equity)

    if not strategy_data:
        raise ValueError("No valid EquityData.csv files found in any strategy folder.")

    # Build union of all dates across strategies
    all_dates_set: set = set()
    for sd in strategy_data:
        all_dates_set.update(sd.dates)

    if use_cutoff and cutoff_date is not None:
        all_dates_set = {d for d in all_dates_set if d <= cutoff_date}

    all_dates = sorted(all_dates_set)
    date_index = pd.DatetimeIndex([pd.Timestamp(d) for d in all_dates])

    # Build aligned matrices
    n_dates = len(all_dates)
    n_strats = len(strategy_data)
    date_pos = {d: i for i, d in enumerate(all_dates)}

    m2m_matrix = np.zeros((n_dates, n_strats), dtype=float)
    closed_matrix = np.zeros((n_dates, n_strats), dtype=float)
    long_matrix = np.zeros((n_dates, n_strats), dtype=float)
    short_matrix = np.zeros((n_dates, n_strats), dtype=float)
    col_names: list[str] = []

    for j, sd in enumerate(strategy_data):
        col_names.append(sd.name)
        for date_val, m2m, closed, long, short in zip(
            sd.dates, sd.m2m, sd.closed, sd.long, sd.short
        ):
            if date_val in date_pos:
                i = date_pos[date_val]
                m2m_matrix[i, j] = m2m
                closed_matrix[i, j] = closed
                long_matrix[i, j] = long
                short_matrix[i, j] = short

    daily_m2m = pd.DataFrame(m2m_matrix, index=date_index, columns=col_names)
    closed_pnl = pd.DataFrame(closed_matrix, index=date_index, columns=col_names)
    in_market_long = pd.DataFrame(long_matrix, index=date_index, columns=col_names)
    in_market_short = pd.DataFrame(short_matrix, index=date_index, columns=col_names)

    # Read trade-level data
    trades = _read_all_trades(strategy_folders, date_format, warnings)

    # Build placeholder strategies list (metadata merged later from config)
    strategies = [
        Strategy(name=sf.name, folder=sf.path, status="")
        for sf in strategy_folders
        if any(sd.name == sf.name for sd in strategy_data)
    ]

    imported = ImportedData(
        daily_m2m=daily_m2m,
        closed_trade_pnl=closed_pnl,
        in_market_long=in_market_long,
        in_market_short=in_market_short,
        trades=trades,
        strategies=strategies,
    )
    return imported, warnings


def _read_equity_csv(
    csv_path: Path,
    strategy_name: str,
    date_format: str,
    warnings: list[str],
) -> _StrategyEquity | None:
    """
    Read one EquityData.csv file.

    Returns _StrategyEquity with per-date lists, or None on failure.
    Mirrors VBA ProcessEquityData.
    """
    try:
        # Read raw — no header assumption, all as strings initially
        raw = pd.read_csv(csv_path, header=None, dtype=str, encoding_errors="replace")
    except Exception as e:
        warnings.append(f"'{strategy_name}': could not read EquityData.csv: {e}")
        return None

    if raw.empty:
        warnings.append(f"'{strategy_name}': EquityData.csv has no data rows.")
        return None

    # Skip header row if first row is non-numeric in col 2
    start_row = 1 if not _is_numeric(raw.iloc[0, EQUITY_COL_M2M]) else 0

    dates: list = []
    m2m: list[float] = []
    closed: list[float] = []
    long_vals: list[float] = []
    short_vals: list[float] = []

    for _, row in raw.iloc[start_row:].iterrows():
        d = parse_csv_date(str(row.iloc[EQUITY_COL_DATE]), date_format)
        if d is None:
            continue

        # Ensure enough columns
        n_cols = len(row)
        if n_cols <= EQUITY_COL_CLOSED:
            warnings.append(
                f"'{strategy_name}': row has only {n_cols} columns, "
                f"expected at least {EQUITY_COL_CLOSED + 1}. Skipping row."
            )
            continue

        dates.append(d)
        m2m.append(_to_float(row.iloc[EQUITY_COL_M2M]))
        long_vals.append(_to_float(row.iloc[EQUITY_COL_LONG]))
        short_vals.append(_to_float(row.iloc[EQUITY_COL_SHORT]))
        closed.append(_to_float(row.iloc[EQUITY_COL_CLOSED]))

    if not dates:
        warnings.append(f"'{strategy_name}': no valid rows parsed from EquityData.csv.")
        return None

    return _StrategyEquity(
        name=strategy_name,
        dates=dates,
        m2m=m2m,
        closed=closed,
        long=long_vals,
        short=short_vals,
    )


def _read_all_trades(
    strategy_folders: list[StrategyFolder],
    date_format: str,
    warnings: list[str],
) -> pd.DataFrame:
    """
    Read TradeData.csv for all strategies and combine into one DataFrame.
    Only Exit rows are included (mirrors VBA ProcessTradeData).
    Returns DataFrame with columns: [strategy, date, position, pnl, mae, mfe]
    """
    frames: list[pd.DataFrame] = []

    for sf in strategy_folders:
        if sf.trade_csv is None:
            continue
        df = _read_trade_csv(sf.trade_csv, sf.name, date_format, warnings)
        if df is not None and not df.empty:
            frames.append(df)

    if not frames:
        return pd.DataFrame(columns=["strategy", "date", "position", "pnl", "mae", "mfe"])

    return pd.concat(frames, ignore_index=True)


def _read_trade_csv(
    csv_path: Path,
    strategy_name: str,
    date_format: str,
    warnings: list[str],
) -> pd.DataFrame | None:
    """
    Read one TradeData.csv file, keeping only Exit rows.
    Mirrors VBA ProcessTradeData.
    """
    try:
        raw = pd.read_csv(csv_path, header=None, dtype=str, encoding_errors="replace")
    except Exception as e:
        warnings.append(f"'{strategy_name}': could not read TradeData.csv: {e}")
        return None

    if raw.empty:
        return None

    min_cols = max(TRADE_COL_MFE, TRADE_COL_MAE, TRADE_COL_PNL,
                   TRADE_COL_POSITION, TRADE_COL_TYPE) + 1

    rows = []
    start_row = 1 if not _is_numeric(raw.iloc[0, TRADE_COL_PNL]) else 0

    for _, row in raw.iloc[start_row:].iterrows():
        if len(row) < min_cols:
            continue

        trade_type = str(row.iloc[TRADE_COL_TYPE]).strip()
        if trade_type.lower() != "exit":
            continue

        d = parse_csv_date(str(row.iloc[TRADE_COL_DATE]), date_format)
        if d is None:
            continue

        position = str(row.iloc[TRADE_COL_POSITION]).strip()
        # Normalise: "Long" → "L", "Short" → "S" (mirrors VBA Left(pos,1))
        position_code = "L" if position.upper().startswith("L") else "S"

        rows.append({
            "strategy": strategy_name,
            "date": d,
            "position": position_code,
            "pnl": _to_float(row.iloc[TRADE_COL_PNL]),
            "mae": _to_float(row.iloc[TRADE_COL_MAE]),
            "mfe": _to_float(row.iloc[TRADE_COL_MFE]),
        })

    if not rows:
        return None

    df = pd.DataFrame(rows)
    df["date"] = pd.to_datetime(df["date"])
    return df


# ── Helpers ───────────────────────────────────────────────────────────────────

def _to_float(val) -> float:
    """Convert a cell value to float, returning 0.0 on failure."""
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def _is_numeric(val) -> bool:
    """Return True if val can be converted to float."""
    try:
        float(str(val).replace(",", "").strip())
        return True
    except (ValueError, TypeError):
        return False
