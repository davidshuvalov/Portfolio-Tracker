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
from core.ingestion.date_utils import detect_date_format, parse_csv_date, _sample_date_strings


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

    Supports two EquityData.csv layouts:
      Single-strategy: col 0 = Date, col 1 = M2M, col 2 = Long, col 3 = Short, col 5 = Closed
      Multi-strategy:  col 0 = Name, col 1 = Date, col 2 = M2M, col 3 = Long, col 4 = Short, col 6 = Closed
    Multi-strategy files (e.g. Buy & Hold benchmarks) produce one column per sub-strategy.

    Args:
        progress_cb: Optional callable(idx, total, name) called after each
                     strategy CSV is read, for progress-bar updates.
    """
    warnings: list[str] = []
    strategy_data: list[_StrategyEquity] = []
    total = len(strategy_folders)

    for idx, sf in enumerate(strategy_folders):
        if progress_cb is not None:
            progress_cb(idx, total, sf.name)
        equities = _read_equity_csv(sf.equity_csv, sf.name, date_format, warnings)
        strategy_data.extend(equities)

    if not strategy_data:
        raise ValueError("No valid EquityData.csv files found in any strategy folder.")

    # Deduplicate strategy names (last one wins for multi-file collisions)
    seen: dict[str, _StrategyEquity] = {}
    for sd in strategy_data:
        if sd.name in seen:
            warnings.append(
                f"Duplicate strategy name '{sd.name}' in multiple files — keeping first occurrence."
            )
        else:
            seen[sd.name] = sd
    strategy_data = list(seen.values())

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
    # For multi-strategy files, sub-strategy names differ from the folder name —
    # map them back to the folder that produced them.
    imported_names = {sd.name for sd in strategy_data}
    _folder_by_name: dict[str, Path] = {}
    for sf in strategy_folders:
        # Single-strategy: folder name == strategy name
        if sf.name in imported_names:
            _folder_by_name[sf.name] = sf.path
    # Sub-strategies from multi-strategy files get the parent folder path
    for sd in strategy_data:
        if sd.name not in _folder_by_name:
            # Find the StrategyFolder whose equity_csv produced this sub-strategy
            # (the folder's equity_csv is the same file that produced the sub-strategy)
            # We use the first folder as fallback
            _folder_by_name[sd.name] = strategy_folders[0].path if strategy_folders else Path(".")

    strategies = [
        Strategy(name=sd.name, folder=_folder_by_name[sd.name], status="")
        for sd in strategy_data
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


def _is_multi_strategy_file(csv_path: Path, date_format: str = "DMY") -> bool:
    """
    Quick probe: return True if csv_path is a multi-strategy EquityData.csv
    (col 0 is a name string, col 1 is a date).
    Used by the Import UI to flag multi-strategy folders at scan time.
    """
    try:
        raw = pd.read_csv(csv_path, header=None, dtype=str,
                          nrows=3, encoding_errors="replace")
    except Exception:
        return False
    if raw.shape[0] < 1 or raw.shape[1] < 2:
        return False
    # Skip potential header
    probe_idx = 0
    if not _is_numeric(raw.iloc[0, 1] if raw.shape[1] > 1 else raw.iloc[0, 0]):
        probe_idx = 1
    if probe_idx >= raw.shape[0]:
        return False
    probe = raw.iloc[probe_idx]
    col0_is_date = parse_csv_date(str(probe.iloc[0]), date_format) is not None
    col1_is_date = parse_csv_date(str(probe.iloc[1]), date_format) is not None
    return not col0_is_date and col1_is_date


def _read_equity_csv(
    csv_path: Path,
    strategy_name: str,
    date_format: str,
    warnings: list[str],
) -> list[_StrategyEquity]:
    """
    Read one EquityData.csv file.

    Returns a list of _StrategyEquity (one for single-strategy files,
    one per named sub-strategy for multi-strategy files).

    Multi-strategy detection:
      If the first data column (col 0) is NOT parseable as a date but col 1 IS,
      the file is treated as multi-strategy. Col 0 is the strategy name, and
      the data columns shift right by one:
        col 0 = Name, col 1 = Date, col 2 = M2M, col 3 = Long, col 4 = Short,
        col 5 = (unused), col 6 = Closed

    Mirrors VBA ProcessEquityData.
    """
    try:
        raw = pd.read_csv(csv_path, header=None, dtype=str, encoding_errors="replace")
    except Exception as e:
        warnings.append(f"'{strategy_name}': could not read EquityData.csv: {e}")
        return []

    if raw.empty:
        warnings.append(f"'{strategy_name}': EquityData.csv has no data rows.")
        return []

    # ── Per-file date format detection ────────────────────────────────────────
    # Scan both col 0 and col 1 for unambiguous date signals (day > 12 → DMY,
    # month-position > 12 → MDY).  In a single-strategy file col 0 holds dates;
    # in a multi-strategy file col 1 does — scanning both is safe because numeric
    # data columns never produce false date-format evidence.
    _date_samples = _sample_date_strings(raw, 0) + _sample_date_strings(raw, 1)
    try:
        file_fmt, fmt_source = detect_date_format(_date_samples, fallback=date_format)
    except ValueError as exc:
        warnings.append(f"'{strategy_name}': {exc} — using global setting '{date_format}'.")
        file_fmt = date_format
        fmt_source = "fallback"

    if fmt_source == "detected" and file_fmt != date_format:
        warnings.append(
            f"'{strategy_name}': date format auto-detected as {file_fmt} "
            f"(global setting is {date_format}). Using detected format for this file."
        )

    # ── Detect multi-strategy format ──────────────────────────────────────────
    # Skip a potential header row first, then probe the first data row
    _probe_row_idx = 0
    if raw.shape[0] > 0 and not _is_numeric(raw.iloc[0, 1] if raw.shape[1] > 1 else raw.iloc[0, 0]):
        _probe_row_idx = 1  # row 0 looks like a header

    if raw.shape[0] > _probe_row_idx and raw.shape[1] >= 2:
        _probe = raw.iloc[_probe_row_idx]
        _col0_is_date = parse_csv_date(str(_probe.iloc[0]), file_fmt) is not None
        _col1_is_date = parse_csv_date(str(_probe.iloc[1]), file_fmt) is not None
        is_multi = not _col0_is_date and _col1_is_date
    else:
        is_multi = False

    if is_multi:
        return _read_multi_strategy_csv(raw, strategy_name, file_fmt, warnings)
    else:
        result = _read_single_strategy_csv(raw, strategy_name, file_fmt, warnings)
        return [result] if result is not None else []


def _read_single_strategy_csv(
    raw: pd.DataFrame,
    strategy_name: str,
    date_format: str,
    warnings: list[str],
) -> _StrategyEquity | None:
    """Parse a single-strategy EquityData.csv (original format)."""
    start_row = 1 if not _is_numeric(raw.iloc[0, EQUITY_COL_M2M]) else 0

    dates: list = []
    m2m: list[float] = []
    closed: list[float] = []
    long_vals: list[float] = []
    short_vals: list[float] = []
    skipped_date = 0
    total_rows = 0

    for _, row in raw.iloc[start_row:].iterrows():
        total_rows += 1
        d = parse_csv_date(str(row.iloc[EQUITY_COL_DATE]), date_format)
        if d is None:
            skipped_date += 1
            continue

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

    # Warn if more than 5% of data rows had unparseable dates (possible format mismatch).
    if total_rows > 0 and skipped_date / total_rows > 0.05:
        warnings.append(
            f"'{strategy_name}': {skipped_date}/{total_rows} rows had unparseable dates "
            f"using format {date_format}. Check the date format setting — "
            f"it may be set to {'DMY' if date_format == 'MDY' else 'MDY'} instead."
        )

    return _StrategyEquity(
        name=strategy_name,
        dates=dates,
        m2m=m2m,
        closed=closed,
        long=long_vals,
        short=short_vals,
    )


# Multi-strategy column offsets (col 0 = name, rest shift by +1 vs single-strategy)
_MULTI_COL_NAME   = 0
_MULTI_COL_DATE   = 1
_MULTI_COL_M2M    = 2
_MULTI_COL_LONG   = 3
_MULTI_COL_SHORT  = 4
_MULTI_COL_CLOSED = 6   # col 5 is unused (same as single-strategy pattern)


def _read_multi_strategy_csv(
    raw: pd.DataFrame,
    file_label: str,
    date_format: str,
    warnings: list[str],
) -> list[_StrategyEquity]:
    """
    Parse a multi-strategy EquityData.csv.
    Col 0 = strategy name (groups rows); col 1 = date; cols 2-6 = data.
    Each unique name in col 0 becomes a separate _StrategyEquity.
    """
    # Skip header row if col 1 is non-date in the first row
    start_row = 0
    if raw.shape[0] > 0 and parse_csv_date(str(raw.iloc[0, _MULTI_COL_DATE]), date_format) is None:
        start_row = 1

    # Group rows by strategy name (col 0)
    groups: dict[str, tuple[list, list, list, list, list]] = {}

    min_cols = _MULTI_COL_CLOSED + 1

    for _, row in raw.iloc[start_row:].iterrows():
        if len(row) < min_cols:
            continue

        sub_name = str(row.iloc[_MULTI_COL_NAME]).strip()
        if not sub_name or sub_name.lower() in ("", "nan", "none"):
            continue

        d = parse_csv_date(str(row.iloc[_MULTI_COL_DATE]), date_format)
        if d is None:
            continue

        if sub_name not in groups:
            groups[sub_name] = ([], [], [], [], [])
        dates_l, m2m_l, closed_l, long_l, short_l = groups[sub_name]
        dates_l.append(d)
        m2m_l.append(_to_float(row.iloc[_MULTI_COL_M2M]))
        long_l.append(_to_float(row.iloc[_MULTI_COL_LONG]))
        short_l.append(_to_float(row.iloc[_MULTI_COL_SHORT]))
        closed_l.append(_to_float(row.iloc[_MULTI_COL_CLOSED]))

    if not groups:
        warnings.append(
            f"'{file_label}': multi-strategy file detected but no valid rows parsed."
        )
        return []

    results = []
    for sub_name, (dates_l, m2m_l, closed_l, long_l, short_l) in groups.items():
        results.append(_StrategyEquity(
            name=sub_name,
            dates=dates_l,
            m2m=m2m_l,
            closed=closed_l,
            long=long_l,
            short=short_l,
        ))

    warnings.append(
        f"'{file_label}': multi-strategy file — imported {len(results)} sub-strategies: "
        + ", ".join(r.name for r in results)
    )
    return results


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

    # ── Per-file date format detection ────────────────────────────────────────
    _trade_date_samples = _sample_date_strings(raw, TRADE_COL_DATE)
    try:
        file_fmt, fmt_source = detect_date_format(_trade_date_samples, fallback=date_format)
    except ValueError as exc:
        warnings.append(f"'{strategy_name}' TradeData: {exc} — using global setting '{date_format}'.")
        file_fmt = date_format
        fmt_source = "fallback"

    if fmt_source == "detected" and file_fmt != date_format:
        warnings.append(
            f"'{strategy_name}' TradeData: date format auto-detected as {file_fmt} "
            f"(global setting is {date_format}). Using detected format for this file."
        )

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

        d = parse_csv_date(str(row.iloc[TRADE_COL_DATE]), file_fmt)
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
