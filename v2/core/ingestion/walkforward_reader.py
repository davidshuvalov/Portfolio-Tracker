"""
Walkforward Details CSV reader — mirrors F_Summary_Tab_Setup.bas CreateHeaderArray logic.

Reads the "Walkforward In-Out Periods Analysis Details.csv" exported by MultiWalk and
extracts per-strategy IS/OOS metrics used to populate the Summary tab.

The CSV uses dynamic column headers — we match by name, not position, exactly as the
VBA FindColumnByHeader() function does.
"""

from __future__ import annotations
from datetime import date
from pathlib import Path
from typing import NamedTuple

import pandas as pd

from core.ingestion.date_utils import detect_date_format, parse_csv_date, resolve_oos_dates


# ── Column names as they appear in the MultiWalk export ──────────────────────

# Core identity
COL_STRATEGY_NAME = "Strategy Name"
COL_SYMBOL_1 = "Symbol Data1"
COL_SYMBOL_2 = "Symbol Data2"
COL_SYMBOL_3 = "Symbol Data3"
COL_INTERVAL_1 = "Interval Data1"
COL_INTERVAL_2 = "Interval Data2"
COL_INTERVAL_3 = "Interval Data3"
COL_SESSION = "Session Name Data1"

# WF structure
COL_IN_LENGTH = "IN Period Length"
COL_IN_TYPE = "IN Period Type"
COL_OUT_LENGTH = "OUT Period Length"
COL_OUT_TYPE = "OUT Period Type"
COL_ANCHORED = "Anchored"
COL_FITNESS = "Fitness Function"

# Dates
COL_IS_BEGIN = "IS Begin Date"
COL_OOS_BEGIN = "OOS Begin Date"
COL_OOS_END = "OOS End Date"
COL_REOPT_DATE = "OOS Estimated Reopt Date"

# P&L metrics from WF details
COL_IS_ANN_NET_PROFIT = "IS Annualized Net Profit"
COL_ISOOS_CHANGE_NET_PROFIT = "IS/OOS Change in Net Profit"
COL_IS_NET_PROFIT = "IS Net Profit"
COL_ISOOS_NET_PROFIT = "IS+OOS Net Profit"
COL_ISOOS_ANN_NET_PROFIT = "IS+OOS Annualized Net Profit"

# Win rates
COL_IS_WIN_RATE = "IS Percent Trades Profitable"
COL_ISOOS_WIN_RATE = "IS+OOS Percent Trades Profitable"
COL_ISOOS_TRADES_PROFIT = "IS+OOS Trades Profitable"
COL_IS_TRADES_PROFIT = "IS Trades Profitable"
COL_ISOOS_CHANGE_TRADES = "IS/OOS Change in Total Trades"
COL_ISOOS_TOTAL_TRADES = "IS+OOS Total Trades"

# Trade stats
COL_ISOOS_PCT_TIME = "IS+OOS Percent Time In Market"
COL_ISOOS_AVG_TRADE = "IS+OOS Avg Trade"
COL_ISOOS_AVG_PROFIT_TRADE = "IS+OOS Avg Profitable Trade"
COL_ISOOS_AVG_LOSS_TRADE = "IS+OOS Avg Unprofitable Trade"
COL_ISOOS_LARGEST_WIN = "IS+OOS Largest Profitable Trade"
COL_ISOOS_LARGEST_LOSS = "IS+OOS Largest Unprofitable Trade"

# Drawdown
COL_IS_MAX_DD = "IS Max DD"
COL_ISOOS_MAX_DD = "IS+OOS Max DD"
COL_IS_AVG_DD = "IS Avg DD"
COL_ISOOS_AVG_DD = "IS+OOS Avg DD"

# Calendar/trading days
COL_IS_TRADING_DAYS = "IS Total Trading Days"
COL_IS_CALENDAR_DAYS = "IS Total Calendar Days"
COL_ISOOS_TRADING_DAYS = "IS+OOS Total Trading Days"
COL_ISOOS_CALENDAR_DAYS = "IS+OOS Total Calendar Days"

# Sharpe
COL_IS_SHARPE = "IS Sharpe Ratio"
COL_ISOOS_SHARPE = "IS+OOS Sharpe Ratio"

# Monte Carlo (from MultiWalk)
COL_IS_MC = "IS Monte Carlo"
COL_ISOOS_MC = "IS+OOS Monte Carlo"

# Margin (optional — present in some MultiWalk exports)
COL_MAINT_MARGIN = "Maint Overnight Margin"
COL_INIT_MARGIN  = "Initial Overnight Margin"


class WalkforwardMetrics(NamedTuple):
    """
    Per-strategy metrics extracted from the Walkforward Details CSV.
    None values indicate the column was missing or unparseable.
    """
    strategy_name: str

    # IS/OOS dates (resolved against cutoff)
    is_begin: date | None
    oos_begin: date | None
    oos_end: date | None         # None = open/ongoing or not found

    # WF structure
    in_period: str               # e.g. "12 Month"
    out_period: str              # e.g. "3 Month"
    next_opt_date: date | None
    last_opt_date: date | None
    anchored: str
    fitness: str
    session: str
    symbol: str
    timeframe: str

    # P&L
    expected_annual_profit: float    # IS Annualized Net Profit
    actual_annual_profit: float      # IS/OOS Change in Net Profit / OOS years
    return_efficiency: float         # actual / expected (clamped)
    total_is_profit: float
    total_isoos_profit: float
    annualized_isoos_profit: float
    is_mc: float                     # MultiWalk IS Monte Carlo
    isoos_mc: float                  # MultiWalk IS+OOS Monte Carlo

    # Win rates
    is_win_rate: float
    oos_win_rate: float
    overall_win_rate: float

    # Trades
    trades_per_year: float
    pct_time_in_market: float
    avg_trade_length: float          # trading days per trade
    avg_trade: float
    avg_profitable_trade: float
    avg_loss_trade: float
    largest_win: float
    largest_loss: float

    # Drawdown (from WF details, absolute $)
    max_drawdown_is: float
    max_drawdown_isoos: float
    avg_drawdown_is: float
    avg_drawdown_isoos: float

    # Standard deviation & Sharpe
    sharpe_is: float
    sharpe_isoos: float
    annual_sd_is: float
    annual_sd_isoos: float

    # Period info
    trading_days_is: int
    trading_days_isoos: int
    oos_period_years: float

    # Margin (optional — present in some MultiWalk WF CSV exports)
    maint_overnight_margin: float = 0.0   # Maint Overnight Margin
    init_overnight_margin: float = 0.0    # Initial Overnight Margin


def read_walkforward_csv(
    csv_path: Path,
    strategy_name: str,
    date_format: str = "DMY",
    use_cutoff: bool = False,
    cutoff_date: date | None = None,
) -> WalkforwardMetrics | None:
    """
    Read one strategy's row from the Walkforward Details CSV.

    Returns WalkforwardMetrics, or None if the CSV can't be read or the
    strategy row isn't found.
    """
    try:
        df = pd.read_csv(csv_path, dtype=str, encoding_errors="replace")
    except Exception:
        return None

    if df.empty:
        return None

    # Strip whitespace from all column headers
    df.columns = [c.strip() for c in df.columns]

    # Find the row for this strategy
    name_col = _find_col(df, COL_STRATEGY_NAME)
    if name_col is None:
        # No Strategy Name column (per-strategy export uses MultiWalk Project).
        # If there is exactly one data row it must be the right one.
        if len(df) == 1:
            row = df.iloc[0]
        else:
            return None
    else:
        row_mask = df[name_col].str.strip() == strategy_name
        if not row_mask.any():
            # Fall back to any row if only one data row (single-strategy WF file)
            if len(df) == 1:
                row = df.iloc[0]
            else:
                return None
        else:
            row = df[row_mask].iloc[0]

    g = _RowGetter(row)

    # ── Per-file date format detection ────────────────────────────────────────
    _date_col_names = [COL_IS_BEGIN, COL_OOS_BEGIN, COL_OOS_END, COL_REOPT_DATE]
    _wf_date_samples = [
        str(row[c]).strip()
        for c in _date_col_names
        if c in row.index and str(row[c]).strip() not in ("", "nan", "None")
    ]
    try:
        wf_fmt, _ = detect_date_format(_wf_date_samples, fallback=date_format)
    except ValueError:
        wf_fmt = date_format   # contradictory evidence: fall back silently for WF row

    # ── Dates ────────────────────────────────────────────────────────────────
    is_begin = g.date(COL_IS_BEGIN, wf_fmt)
    oos_begin_raw = g.date(COL_OOS_BEGIN, wf_fmt)
    oos_end_raw = g.date(COL_OOS_END, wf_fmt)

    oos_begin, oos_end = resolve_oos_dates(
        oos_begin_raw, oos_end_raw, use_cutoff, cutoff_date
    )

    # ── OOS period years ─────────────────────────────────────────────────────
    oos_period_years = 0.0
    if oos_begin and oos_end and oos_begin != oos_end:
        oos_period_years = (oos_end - oos_begin).days / 365.25

    # ── Period structure ──────────────────────────────────────────────────────
    in_len = g.str(COL_IN_LENGTH)
    in_type = g.str(COL_IN_TYPE)
    out_len = g.str(COL_OUT_LENGTH)
    out_type = g.str(COL_OUT_TYPE)
    in_period = f"{in_len} {in_type}".strip()
    out_period = f"{out_len} {out_type}".strip()

    # ── Next/Last opt dates ───────────────────────────────────────────────────
    next_opt_date = g.date(COL_REOPT_DATE, wf_fmt)
    last_opt_date = _calc_last_opt_date(next_opt_date, out_len, out_type)

    # ── P&L ──────────────────────────────────────────────────────────────────
    expected_annual_profit = g.flt(COL_IS_ANN_NET_PROFIT)
    isoos_change = g.flt(COL_ISOOS_CHANGE_NET_PROFIT)
    actual_annual_profit = (
        isoos_change / (oos_period_years + 1e-9)
        if oos_period_years > 0 else 0.0
    )
    if abs(expected_annual_profit) < 1e-3:
        return_efficiency = 0.0
    else:
        return_efficiency = actual_annual_profit / expected_annual_profit

    # ── Win rates ─────────────────────────────────────────────────────────────
    is_win_rate = g.flt(COL_IS_WIN_RATE)
    overall_win_rate = g.flt(COL_ISOOS_WIN_RATE)
    isoos_trades_profit = g.flt(COL_ISOOS_TRADES_PROFIT)
    is_trades_profit = g.flt(COL_IS_TRADES_PROFIT)
    isoos_change_trades = g.flt(COL_ISOOS_CHANGE_TRADES)
    oos_win_rate = (
        (isoos_trades_profit - is_trades_profit) / (isoos_change_trades + 1e-9)
        if isoos_change_trades != 0 else 0.0
    )

    # ── Trades ────────────────────────────────────────────────────────────────
    isoos_total_trades = g.flt(COL_ISOOS_TOTAL_TRADES)
    # allPeriodYears = from IS begin to OOS end
    all_period_years = 0.0
    if is_begin and oos_end:
        all_period_years = (oos_end - is_begin).days / 365.25
    trades_per_year = round(isoos_total_trades / (all_period_years + 1e-9), 0) if all_period_years > 0 else 0.0

    pct_time = g.flt(COL_ISOOS_PCT_TIME)
    avg_trade_length = (
        (265.0 * pct_time) / (trades_per_year + 1e-9)
        if trades_per_year > 0 else 0.0
    )

    # ── Calendar & trading days ───────────────────────────────────────────────
    is_trading = g.flt(COL_IS_TRADING_DAYS)
    is_calendar = g.flt(COL_IS_CALENDAR_DAYS)
    isoos_trading = g.flt(COL_ISOOS_TRADING_DAYS)
    isoos_calendar = g.flt(COL_ISOOS_CALENDAR_DAYS)

    trading_days_is = round(is_trading / (is_calendar + 1e-9) * 365.25) if is_calendar > 0 else 0
    trading_days_isoos = round(isoos_trading / (isoos_calendar + 1e-9) * 365.25) if isoos_calendar > 0 else 0

    # ── Sharpe & SD ───────────────────────────────────────────────────────────
    sharpe_is = abs(g.flt(COL_IS_SHARPE))
    sharpe_isoos = abs(g.flt(COL_ISOOS_SHARPE))
    annual_sd_is = abs(
        expected_annual_profit / ((sharpe_is * (365.25 ** 0.5)) + 1e-9)
    )
    avg_trade = abs(g.flt(COL_ISOOS_AVG_TRADE))
    annual_sd_isoos = abs(
        (avg_trade * trades_per_year) / ((sharpe_is * (365.25 ** 0.5)) + 1e-9)
    )

    return WalkforwardMetrics(
        strategy_name=strategy_name,
        is_begin=is_begin,
        oos_begin=oos_begin,
        oos_end=oos_end,
        in_period=in_period,
        out_period=out_period,
        next_opt_date=next_opt_date,
        last_opt_date=last_opt_date,
        anchored=g.str(COL_ANCHORED),
        fitness=g.str(COL_FITNESS),
        session=g.str(COL_SESSION),
        symbol=_clean_symbol(g.str(COL_SYMBOL_1)),
        timeframe=g.str(COL_INTERVAL_1),
        expected_annual_profit=expected_annual_profit,
        actual_annual_profit=actual_annual_profit,
        return_efficiency=return_efficiency,
        total_is_profit=abs(g.flt(COL_IS_NET_PROFIT)),
        total_isoos_profit=abs(g.flt(COL_ISOOS_NET_PROFIT)),
        annualized_isoos_profit=abs(g.flt(COL_ISOOS_ANN_NET_PROFIT)),
        is_mc=g.flt(COL_IS_MC),
        isoos_mc=g.flt(COL_ISOOS_MC),
        is_win_rate=is_win_rate,
        oos_win_rate=oos_win_rate,
        overall_win_rate=overall_win_rate,
        trades_per_year=trades_per_year,
        pct_time_in_market=pct_time,
        avg_trade_length=avg_trade_length,
        avg_trade=avg_trade,
        avg_profitable_trade=abs(g.flt(COL_ISOOS_AVG_PROFIT_TRADE)),
        avg_loss_trade=abs(g.flt(COL_ISOOS_AVG_LOSS_TRADE)),
        largest_win=abs(g.flt(COL_ISOOS_LARGEST_WIN)),
        largest_loss=abs(g.flt(COL_ISOOS_LARGEST_LOSS)),
        max_drawdown_is=abs(g.flt(COL_IS_MAX_DD)),
        max_drawdown_isoos=abs(g.flt(COL_ISOOS_MAX_DD)),
        avg_drawdown_is=abs(g.flt(COL_IS_AVG_DD)),
        avg_drawdown_isoos=abs(g.flt(COL_ISOOS_AVG_DD)),
        sharpe_is=sharpe_is,
        sharpe_isoos=sharpe_isoos,
        annual_sd_is=annual_sd_is,
        annual_sd_isoos=annual_sd_isoos,
        trading_days_is=int(trading_days_is),
        trading_days_isoos=int(trading_days_isoos),
        oos_period_years=oos_period_years,
        maint_overnight_margin=abs(g.flt(COL_MAINT_MARGIN)),
        init_overnight_margin=abs(g.flt(COL_INIT_MARGIN)),
    )


# ── Helpers ───────────────────────────────────────────────────────────────────

class _RowGetter:
    """Safe accessor for a pandas Series row — returns defaults on missing/invalid."""

    def __init__(self, row: pd.Series):
        self._row = row
        # Build case-insensitive lookup
        self._cols = {c.strip().lower(): c.strip() for c in row.index}

    def _get(self, col_name: str) -> str:
        key = col_name.strip().lower()
        actual = self._cols.get(key)
        if actual is None:
            return ""
        val = self._row.get(actual, "")
        return "" if pd.isna(val) else str(val).strip()

    def str(self, col_name: str) -> str:
        return self._get(col_name)

    def flt(self, col_name: str) -> float:
        try:
            return float(self._get(col_name).replace(",", ""))
        except (ValueError, TypeError):
            return 0.0

    def date(self, col_name: str, date_format: str) -> date | None:
        return parse_csv_date(self._get(col_name), date_format)


def _find_col(df: pd.DataFrame, col_name: str) -> str | None:
    """Case-insensitive column search."""
    target = col_name.strip().lower()
    for c in df.columns:
        if c.strip().lower() == target:
            return c
    return None


def _clean_symbol(symbol: str) -> str:
    """Strip special chars from MultiWalk symbol names (mirrors VBA)."""
    return symbol.replace("@", "").replace("$", "").replace(".D", "").upper()


def _calc_last_opt_date(
    next_opt: date | None,
    out_len: str,
    out_type: str,
) -> date | None:
    """Derive last opt date = next opt date minus one OOS period."""
    if next_opt is None:
        return None
    try:
        length = float(out_len)
    except (ValueError, TypeError):
        return None
    t = (out_type or "").strip().lower()
    from datetime import timedelta
    if t == "month":
        days = round(length * 30.5)
    elif t == "year":
        days = round(length * 365.25)
    elif t in ("trading days", "trading day"):
        days = round(length * 365.25 / 252)
    else:
        return None
    return next_opt - timedelta(days=days)
