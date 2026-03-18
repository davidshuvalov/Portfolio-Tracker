"""
Per-strategy summary metrics — mirrors F_Summary_Tab_Setup.bas.

Two computation paths:
1. Static metrics from the Walkforward Details CSV (via walkforward_reader)
2. Dynamic metrics from daily_m2m DataFrame (profit windows, drawdowns, incubation)

Returns a DataFrame with index=strategy_name and columns=all metric fields.
Strategies without WF data get NaN for WF-sourced columns.
"""

from __future__ import annotations
from datetime import date, timedelta
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd

from core.analytics.monte_carlo import closed_trade_mc
from core.config import EligibilityConfig
from core.data_types import ImportedData, StrategyFolder
from core.ingestion.date_utils import resolve_oos_dates
from core.ingestion.walkforward_reader import WalkforwardMetrics, read_walkforward_csv


# ── Days threshold default (mirrors EligibilityDaysThreshold named range) ───
DEFAULT_DAYS_THRESHOLD = 0     # 0 = rolling window mode (not calendar-month mode)
DEFAULT_INCUBATION_MONTHS = 6  # months before incubation check begins
DEFAULT_MIN_INCUBATION_RATIO = 1.0  # profit must reach 1x expected rate
DEFAULT_ELIGIBILITY_MONTHS = 12  # for count_profit_months


def compute_summary(
    imported: ImportedData,
    strategy_folders: list[StrategyFolder],
    date_format: str = "DMY",
    use_cutoff: bool = False,
    cutoff_date: date | None = None,
    days_threshold: int = DEFAULT_DAYS_THRESHOLD,
    incubation_months: int = DEFAULT_INCUBATION_MONTHS,
    min_incubation_ratio: float = DEFAULT_MIN_INCUBATION_RATIO,
    eligibility_months: int = DEFAULT_ELIGIBILITY_MONTHS,
    quitting_method: str = "Drawdown",
    quitting_max_dollars: float = 50_000.0,
    quitting_max_percent: float = 1.5,
    quitting_sd_multiple: float = 1.28,
    data_scope: str = "OOS",
) -> pd.DataFrame:
    """
    Compute per-strategy summary metrics for all strategies in imported.

    Args:
        imported:          Full imported data (daily_m2m, trades, etc.)
        strategy_folders:  StrategyFolder list from scan (for WF CSV paths)
        date_format:       "DMY" or "MDY"
        use_cutoff:        Whether to apply a cutoff date to OOS end
        cutoff_date:       The cutoff date (if use_cutoff)
        days_threshold:    Min days in current month before using it (0=rolling)
        incubation_months: OOS months required before incubation check
        min_incubation_ratio: Profit target as multiple of expected rate
        eligibility_months: Lookback window for count_profit_months

    Returns:
        DataFrame indexed by strategy name with all metric columns.
    """
    # Build lookup: strategy_name → StrategyFolder
    folder_map = {sf.name: sf for sf in strategy_folders}

    # ── Pre-compute direction and trade PnL arrays ────────────────────────────
    _direction_map: dict[str, str] = {}
    _trade_pnls: dict[str, np.ndarray] = {}  # all trades, per strategy
    if imported.trades is not None and not imported.trades.empty:
        for _nm in imported.strategy_names:
            _t = imported.trades[imported.trades["strategy"] == _nm]
            _has_long = bool((_t["position"] == "L").any())
            _has_short = bool((_t["position"] == "S").any())
            if _has_long and _has_short:
                _direction_map[_nm] = "Long & Short"
            elif _has_long:
                _direction_map[_nm] = "Long Only"
            elif _has_short:
                _direction_map[_nm] = "Short Only"
            else:
                _direction_map[_nm] = ""
            if "pnl" in _t.columns and not _t.empty:
                _trade_pnls[_nm] = _t["pnl"].dropna().values.astype(np.float64)

    rows: list[dict[str, Any]] = []

    for name in imported.strategy_names:
        sf = folder_map.get(name)

        # Read WF metrics (may be None if WF CSV absent)
        wf: WalkforwardMetrics | None = None
        if sf and sf.walkforward_csv:
            wf = read_walkforward_csv(
                sf.walkforward_csv, name, date_format, use_cutoff, cutoff_date
            )

        # Daily PnL series for this strategy
        pnl: pd.Series = imported.daily_m2m[name]

        # Determine OOS dates
        oos_begin: date | None = wf.oos_begin if wf else None
        oos_end: date | None = wf.oos_end if wf else None

        # Compute dynamic metrics from daily_m2m
        dynamic = _compute_dynamic_metrics(
            pnl=pnl,
            oos_begin=oos_begin,
            oos_end=oos_end,
            expected_annual_profit=wf.expected_annual_profit if wf else 0.0,
            annual_sd_is=wf.annual_sd_is if wf else 0.0,
            is_max_drawdown=wf.max_drawdown_is if wf else 0.0,
            days_threshold=days_threshold,
            incubation_months=incubation_months,
            min_incubation_ratio=min_incubation_ratio,
            eligibility_months=eligibility_months,
            quitting_method=quitting_method,
            quitting_max_dollars=quitting_max_dollars,
            quitting_max_percent=quitting_max_percent,
            quitting_sd_multiple=quitting_sd_multiple,
            data_scope=data_scope,
        )

        row = _build_row(name, wf, dynamic)

        # ── Direction (auto-detected from trade data) ─────────────────────────
        row["direction"] = _direction_map.get(name, "")

        # ── Last date on file ─────────────────────────────────────────────────
        _last = pnl.dropna()
        row["last_date_on_file"] = _last.index[-1].date() if not _last.empty else None

        # ── Closed-trade Monte Carlo (IS only and IS+OOS) ─────────────────────
        _all_trades = _trade_pnls.get(name, np.array([], dtype=np.float64))
        _tpy = int(round(float(wf.trades_per_year))) if wf and wf.trades_per_year else max(1, len(_all_trades))
        if len(_all_trades) >= 2:
            # IS only: trades before oos_begin
            if oos_begin is not None and not imported.trades.empty:
                _strat_t = imported.trades[imported.trades["strategy"] == name]
                _is_trades = _strat_t.loc[
                    _strat_t["date"] < pd.Timestamp(oos_begin), "pnl"
                ].dropna().values.astype(np.float64)
            else:
                _is_trades = _all_trades
            row["mc_closed_is"] = closed_trade_mc(_is_trades, _tpy) if len(_is_trades) >= 2 else float("nan")
            row["mc_closed_isoos"] = closed_trade_mc(_all_trades, _tpy)
        else:
            row["mc_closed_is"] = float("nan")
            row["mc_closed_isoos"] = float("nan")

        rows.append(row)

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows).set_index("strategy_name")
    return df


# ── Row builder ───────────────────────────────────────────────────────────────

def _build_row(
    name: str,
    wf: WalkforwardMetrics | None,
    dyn: dict[str, Any],
) -> dict[str, Any]:
    """Merge WF metrics and dynamic metrics into a single flat dict."""
    row: dict[str, Any] = {"strategy_name": name}

    if wf:
        row.update({
            # Identity
            "symbol": wf.symbol,
            "timeframe": wf.timeframe,
            "session": wf.session,
            # Dates
            "is_begin": wf.is_begin,
            "oos_begin": wf.oos_begin,
            "oos_end": wf.oos_end,
            "next_opt_date": wf.next_opt_date,
            "last_opt_date": wf.last_opt_date,
            "oos_period_years": wf.oos_period_years,
            # WF structure
            "in_period": wf.in_period,
            "out_period": wf.out_period,
            "anchored": wf.anchored,
            "fitness": wf.fitness,
            # P&L
            "expected_annual_profit": wf.expected_annual_profit,
            "actual_annual_profit": wf.actual_annual_profit,
            "return_efficiency": wf.return_efficiency,
            "total_is_profit": wf.total_is_profit,
            "total_isoos_profit": wf.total_isoos_profit,
            "annualized_isoos_profit": wf.annualized_isoos_profit,
            "mw_mc_is": wf.is_mc,
            "mw_mc_isoos": wf.isoos_mc,
            # Win rates
            "is_win_rate": wf.is_win_rate,
            "oos_win_rate": wf.oos_win_rate,
            "overall_win_rate": wf.overall_win_rate,
            # Trades
            "trades_per_year": wf.trades_per_year,
            "pct_time_in_market": wf.pct_time_in_market,
            "avg_trade_length": wf.avg_trade_length,
            "avg_trade": wf.avg_trade,
            "avg_profitable_trade": wf.avg_profitable_trade,
            "avg_loss_trade": wf.avg_loss_trade,
            "largest_win": wf.largest_win,
            "largest_loss": wf.largest_loss,
            # Drawdown (from WF)
            "max_drawdown_is": wf.max_drawdown_is,
            "max_drawdown_isoos": wf.max_drawdown_isoos,
            "avg_drawdown_is": wf.avg_drawdown_is,
            "avg_drawdown_isoos": wf.avg_drawdown_isoos,
            # Sharpe & SD
            "sharpe_is": wf.sharpe_is,
            "sharpe_isoos": wf.sharpe_isoos,
            "annual_sd_is": wf.annual_sd_is,
            "annual_sd_isoos": wf.annual_sd_isoos,
            "trading_days_is": wf.trading_days_is,
            "trading_days_isoos": wf.trading_days_isoos,
            # Margin data from MultiWalk WF CSV (0.0 when column absent)
            "mw_maint_margin": wf.maint_overnight_margin,
            "mw_init_margin": wf.init_overnight_margin,
        })
    else:
        # Fill WF-sourced columns with NaN
        for col in _WF_FLOAT_COLS:
            row[col] = float("nan")
        for col in _WF_STR_COLS:
            row[col] = ""
        for col in _WF_DATE_COLS:
            row[col] = None
        for col in _WF_INT_COLS:
            row[col] = 0

    # Merge dynamic (always computed)
    row.update(dyn)

    # Derived metrics that combine WF + dynamic
    profit_12m = dyn.get("profit_last_12_months", 0.0) or 0.0
    max_dd_12m = dyn.get("max_drawdown_last_12_months", 0.0) or 0.0
    profit_oos = dyn.get("profit_since_oos_start", 0.0) or 0.0
    max_dd_oos = dyn.get("max_oos_drawdown", 0.0) or 0.0

    row["rtd_12_months"] = (
        profit_12m / (max_dd_12m + 1e-4) if abs(max_dd_12m) >= 10 else 10.0
    )
    row["rtd_oos"] = (
        profit_oos / (max_dd_oos + 1e-4) if abs(max_dd_oos) >= 10 else 10.0
    )

    return row


# ── Dynamic metrics from daily_m2m ────────────────────────────────────────────

def _compute_dynamic_metrics(
    pnl: pd.Series,
    oos_begin: date | None,
    oos_end: date | None,
    expected_annual_profit: float,
    annual_sd_is: float,
    is_max_drawdown: float,
    days_threshold: int,
    incubation_months: int,
    min_incubation_ratio: float,
    eligibility_months: int,
    quitting_method: str = "Drawdown",
    quitting_max_dollars: float = 50_000.0,
    quitting_max_percent: float = 1.5,
    quitting_sd_multiple: float = 1.28,
    data_scope: str = "OOS",
) -> dict[str, Any]:
    """
    Compute time-windowed and drawdown metrics from the daily PnL series.
    Mirrors VBA CalculateProfitAndDrawdown() exactly.
    """
    result: dict[str, Any] = {
        "profit_last_1_month": None,
        "profit_last_3_months": None,
        "profit_last_6_months": None,
        "profit_last_9_months": None,
        "profit_last_12_months": None,
        "efficiency_last_1_month": None,
        "efficiency_last_3_months": None,
        "efficiency_last_6_months": None,
        "efficiency_last_9_months": None,
        "efficiency_last_12_months": None,
        "profit_since_oos_start": 0.0,
        "max_oos_drawdown": 0.0,
        "avg_oos_drawdown": 0.0,
        "max_drawdown_last_12_months": 0.0,
        "count_profit_months": 0,
        "incubation_status": "",
        "incubation_date": None,
        "quitting_status": "",
        "quitting_date": None,
        # Sprint 3.1 metrics
        "k_factor": None,
        "ulcer_index": None,
        "best_month": None,
        "worst_month": None,
        "max_consecutive_loss_months": None,
        # Sprint 3.2 metrics
        "profit_since_quit": None,
    }

    if oos_begin is None or oos_end is None:
        return result
    if pnl.empty:
        return result

    # Slice to OOS period
    oos_begin_ts = pd.Timestamp(oos_begin)
    oos_end_ts = pd.Timestamp(oos_end)
    oos_pnl = pnl.loc[(pnl.index >= oos_begin_ts) & (pnl.index <= oos_end_ts)]

    if oos_pnl.empty:
        return result

    # ── Effective end date & rolling window start dates ───────────────────────
    effective_end_ts, window_starts = _calc_window_starts(
        oos_end_ts, oos_begin_ts, days_threshold
    )

    # ── Profit windows ────────────────────────────────────────────────────────
    # IS+OOS scope: include full history in rolling P&L windows (not just OOS slice)
    _window_base = pnl if data_scope == "IS+OOS" else oos_pnl
    window_pnl = _window_base.loc[_window_base.index <= effective_end_ts]
    windows = {
        "profit_last_1_month":  window_starts["1m"],
        "profit_last_3_months": window_starts["3m"],
        "profit_last_6_months": window_starts["6m"],
        "profit_last_9_months": window_starts["9m"],
        "profit_last_12_months": window_starts["12m"],
    }
    for key, start_ts in windows.items():
        # OOS scope: only compute if window start is within the OOS period
        # IS+OOS scope: allow windows to reach back into IS history
        _in_scope = start_ts is not None and (
            data_scope == "IS+OOS" or start_ts >= oos_begin_ts
        )
        if _in_scope:
            result[key] = float(
                window_pnl.loc[window_pnl.index >= start_ts].sum()
            )
        # else: leave as None (not enough history in scope)

    # ── Efficiency (profit / expected_annual × window_fraction) ──────────────
    # Calendar mode (days_threshold > 0): mirrors VBA — use actual calendar days
    #     efficiency = profit / (expected_annual × actual_days / 365.25)
    # Rolling mode (days_threshold = 0): use months/12
    #     efficiency = profit / (expected_annual × months / 12)
    if expected_annual_profit > 0:
        for wkey, profit_key, eff_key in [
            ("1m",  "profit_last_1_month",   "efficiency_last_1_month"),
            ("3m",  "profit_last_3_months",  "efficiency_last_3_months"),
            ("6m",  "profit_last_6_months",  "efficiency_last_6_months"),
            ("9m",  "profit_last_9_months",  "efficiency_last_9_months"),
            ("12m", "profit_last_12_months", "efficiency_last_12_months"),
        ]:
            p = result[profit_key]
            if p is None:
                continue
            start_ts = window_starts.get(wkey)
            if start_ts is not None and days_threshold > 0:
                # Calendar mode: actual days in window (matches VBA daysN = DateDiff + 1)
                actual_days = (effective_end_ts - start_ts).days + 1
                denom = expected_annual_profit * actual_days / 365.25
            else:
                # Rolling mode: fixed months/12 fraction
                months_map = {"1m": 1, "3m": 3, "6m": 6, "9m": 9, "12m": 12}
                denom = expected_annual_profit * months_map[wkey] / 12.0
            if denom != 0:
                result[eff_key] = p / denom

    # ── Profit since OOS start ────────────────────────────────────────────────
    result["profit_since_oos_start"] = float(oos_pnl.sum())

    # ── OOS drawdown ──────────────────────────────────────────────────────────
    oos_cumsum = oos_pnl.cumsum()
    max_dd, avg_dd = _calc_drawdown(oos_cumsum.values)
    result["max_oos_drawdown"] = max_dd
    result["avg_oos_drawdown"] = avg_dd

    # ── Drawdown last 12 months ───────────────────────────────────────────────
    start_12m = window_starts["12m"]
    if start_12m is not None:
        pnl_12m = window_pnl.loc[window_pnl.index >= start_12m]
        if not pnl_12m.empty:
            cum_12m = pnl_12m.cumsum()
            result["max_drawdown_last_12_months"], _ = _calc_drawdown(cum_12m.values)

    # ── Count profitable months ───────────────────────────────────────────────
    monthly = oos_pnl.resample("ME").sum()
    if eligibility_months > 0 and len(monthly) > 0:
        recent = monthly.iloc[-eligibility_months:]
        result["count_profit_months"] = int((recent > 0).sum())

    # ── Sprint 3.1: additional metrics ───────────────────────────────────────
    monthly = oos_pnl.resample("ME").sum()
    if not monthly.empty:
        result["best_month"] = float(monthly.max())
        result["worst_month"] = float(monthly.min())

        # Max consecutive losing months
        losing_streak = 0
        max_streak = 0
        for v in monthly:
            if v < 0:
                losing_streak += 1
                max_streak = max(max_streak, losing_streak)
            else:
                losing_streak = 0
        result["max_consecutive_loss_months"] = max_streak

        # K-Factor: ratio of (profit_months / total_months) × (avg_win / abs(avg_loss))
        profit_months_vals = monthly[monthly > 0]
        loss_months_vals = monthly[monthly < 0]
        if len(profit_months_vals) > 0 and len(loss_months_vals) > 0:
            win_rate = len(profit_months_vals) / len(monthly)
            avg_win = float(profit_months_vals.mean())
            avg_loss = abs(float(loss_months_vals.mean()))
            result["k_factor"] = (win_rate / (1.0 - win_rate)) * (avg_win / avg_loss) if avg_loss > 0 else None

    # Ulcer Index: RMS of % drawdown over OOS period (Peter Martin, 1987)
    if not oos_pnl.empty:
        oos_eq = oos_pnl.cumsum()
        oos_peak = oos_eq.cummax()
        # % drawdown at each point (avoid division by zero)
        pct_dd = np.where(oos_peak > 1e-9, (oos_peak - oos_eq) / oos_peak * 100.0, 0.0)
        result["ulcer_index"] = float(np.sqrt(np.mean(pct_dd ** 2)))

    # ── Incubation status ─────────────────────────────────────────────────────
    inc_status, inc_date = _calc_incubation(
        oos_pnl,
        expected_annual_profit,
        incubation_months,
        min_incubation_ratio,
    )
    result["incubation_status"] = inc_status
    result["incubation_date"] = inc_date

    # ── Quitting status ───────────────────────────────────────────────────────
    quit_status, quit_date = _calc_quitting_status(
        oos_pnl,
        expected_annual_profit,
        annual_sd_is,
        is_max_drawdown,
        quitting_method,
        quitting_max_dollars,
        quitting_max_percent,
        quitting_sd_multiple,
    )
    result["quitting_status"] = quit_status
    result["quitting_date"] = quit_date

    # Sprint 3.2: profit since last quit point
    if quit_date is not None:
        quit_ts = pd.Timestamp(quit_date)
        after_quit = oos_pnl[oos_pnl.index > quit_ts]
        result["profit_since_quit"] = float(after_quit.sum()) if not after_quit.empty else 0.0

    return result


def _calc_window_starts(
    oos_end_ts: pd.Timestamp,
    oos_begin_ts: pd.Timestamp,
    days_threshold: int,
) -> tuple[pd.Timestamp, dict[str, pd.Timestamp | None]]:
    """
    Compute rolling window start timestamps.
    Mirrors VBA: if daysThreshold > 0 and current month has < threshold days,
    use previous month's end as effective end date.
    """
    # Clamp to valid range: 0 = rolling, 1–31 = calendar snap.
    # Values > 31 would always trigger "snap to previous month" (no month has > 31 days),
    # which is incorrect — guard against stale configs saved with the old max=730 UI bug.
    days_threshold = min(days_threshold, 31)

    if days_threshold > 0:
        month_start = oos_end_ts.replace(day=1)
        days_in_current = (oos_end_ts - month_start).days + 1
        if days_in_current < days_threshold:
            # Use end of previous month
            effective_end = month_start - timedelta(days=1)
        else:
            effective_end = oos_end_ts
    else:
        effective_end = oos_end_ts

    eff = effective_end

    def _months_back(n: int) -> pd.Timestamp:
        """Go back n months from effective end — to start of that month."""
        dt = eff.to_pydatetime() if hasattr(eff, "to_pydatetime") else eff
        m = dt.month - n
        y = dt.year
        while m <= 0:
            m += 12
            y -= 1
        return pd.Timestamp(y, m, 1)

    if days_threshold > 0:
        starts = {
            "1m":  _months_back(0),    # first day of effective end month
            "3m":  _months_back(2),
            "6m":  _months_back(5),
            "9m":  _months_back(8),
            "12m": _months_back(11),
        }
    else:
        # Rolling: last N calendar months from effective end
        starts = {
            "1m":  eff - pd.DateOffset(months=1) + timedelta(days=1),
            "3m":  eff - pd.DateOffset(months=3) + timedelta(days=1),
            "6m":  eff - pd.DateOffset(months=6) + timedelta(days=1),
            "9m":  eff - pd.DateOffset(months=9) + timedelta(days=1),
            "12m": eff - pd.DateOffset(months=12) + timedelta(days=1),
        }

    return effective_end, starts


def _calc_drawdown(equity: np.ndarray) -> tuple[float, float]:
    """
    Compute max and average dollar drawdown from an equity curve array.
    Uses absolute dollar drawdown (not %) to mirror VBA.
    Returns (max_drawdown, avg_drawdown).
    """
    if len(equity) == 0:
        return 0.0, 0.0

    peak = np.maximum.accumulate(equity)
    drawdown = peak - equity  # absolute dollar drawdown

    max_dd = float(np.max(drawdown))
    # Average of non-zero drawdown points
    nonzero = drawdown[drawdown > 0]
    avg_dd = float(np.mean(nonzero)) if len(nonzero) > 0 else 0.0

    return max_dd, avg_dd


def _calc_incubation(
    oos_pnl: pd.Series,
    expected_annual_profit: float,
    incubation_months: int,
    min_incubation_ratio: float,
) -> tuple[str, date | None]:
    """
    Determine incubation status.
    Mirrors VBA: after incubation_months of OOS data, check if cumulative profit
    has reached (expected_daily_rate × days × min_incubation_ratio).

    Returns:
        ("Passed", date)      — target hit; date is when it first passed
        ("Not Passed", None)  — enough OOS history but target never reached
        ("Incubating", None)  — OOS period not long enough yet to evaluate
        ("", None)            — no OOS data / no expected profit to compare
    """
    if oos_pnl.empty or expected_annual_profit <= 0:
        return "", None

    incubation_days = round(incubation_months * 30.5)
    expected_daily = expected_annual_profit / 365.25

    cum = oos_pnl.cumsum()
    reached_threshold = False

    for i, (ts, cum_val) in enumerate(cum.items()):
        days_elapsed = i + 1
        if days_elapsed >= incubation_days:
            reached_threshold = True
            target = expected_daily * days_elapsed * min_incubation_ratio
            if cum_val >= target:
                return "Passed", ts.date()

    if reached_threshold:
        return "Not Passed", None
    return "Incubating", None


def _calc_quitting_status(
    oos_pnl: pd.Series,
    expected_annual_profit: float,
    annual_sd_is: float,
    is_max_drawdown: float,
    quitting_method: str,
    max_dollars: float,
    max_percent: float,
    sd_multiple: float,
) -> tuple[str, date | None]:
    """
    Compute strategy quitting status — mirrors VBA CalculateProfitAndDrawdown()
    quitting state machine.

    States: "" | "Continue" | "Quit" | "Coming Back" | "Recovered" | "N/A"
    Strategy must have > 21 OOS days before quitting is evaluated.

    Method "Drawdown":
        quitting_point = MIN(max_dollars, max_percent × |IS_max_drawdown|)
        Quit when: current_equity < peak_equity - quitting_point

    Method "Standard Deviation":
        quit_equity = expected_daily×days − sqrt(days)×(annual_sd/sqrt(365.25))×sd_multiple
        Quit when: current_equity < quit_equity
    """
    if oos_pnl.empty or quitting_method == "None":
        return "N/A" if quitting_method == "None" else "", None

    if quitting_method not in ("Drawdown", "Standard Deviation"):
        return "N/A", None

    # Pre-compute quitting threshold for Drawdown method
    if quitting_method == "Drawdown":
        if is_max_drawdown <= 0 and max_dollars <= 0:
            return "", None
        quitting_point = min(max_dollars, max_percent * abs(is_max_drawdown))
    else:
        if expected_annual_profit <= 0 or annual_sd_is <= 0:
            return "", None
        expected_daily = expected_annual_profit / 365.25
        sd_daily = annual_sd_is / (365.25 ** 0.5)

    expected_daily_base = expected_annual_profit / 365.25 if expected_annual_profit > 0 else 0.0

    status = "Continue"
    quit_date: date | None = None
    peak_equity = 0.0
    current_equity = 0.0
    last_quit_equity_high = 0.0

    for day_idx, (ts, daily_pnl) in enumerate(oos_pnl.items()):
        current_equity += float(daily_pnl)
        if current_equity > peak_equity:
            peak_equity = current_equity

        days_elapsed = day_idx + 1

        # Compute quit equity threshold
        if quitting_method == "Drawdown":
            quit_equity = peak_equity - quitting_point
        else:  # Standard Deviation
            quit_equity = (
                expected_daily_base * days_elapsed
                - (days_elapsed ** 0.5) * sd_daily * sd_multiple
            )

        # Recovery point (mirrors VBA recoveryPoint logic)
        incubation_target = expected_daily_base * days_elapsed
        recovery_point = (
            incubation_target if peak_equity < incubation_target else last_quit_equity_high
        )

        # Only evaluate quitting after 21 days
        if days_elapsed <= 21:
            continue

        if status == "Continue":
            if current_equity < quit_equity:
                status = "Quit"
                quit_date = ts.date()
                last_quit_equity_high = peak_equity

        elif status == "Quit":
            midpoint = (recovery_point + quit_equity) / 2.0
            if current_equity > recovery_point:
                status = "Recovered"
            elif current_equity > midpoint:
                status = "Coming Back"

        elif status == "Coming Back":
            midpoint = (recovery_point + quit_equity) / 2.0
            if current_equity > recovery_point:
                status = "Recovered"
            elif current_equity < midpoint:
                status = "Quit"
                last_quit_equity_high = peak_equity

        elif status == "Recovered":
            if current_equity < quit_equity:
                status = "Quit"
                quit_date = ts.date()
                last_quit_equity_high = peak_equity

    return status, quit_date


# ── Column group definitions (for NaN-fill when WF missing) ───────────────────

_WF_FLOAT_COLS = [
    "expected_annual_profit", "actual_annual_profit", "return_efficiency",
    "total_is_profit", "total_isoos_profit", "annualized_isoos_profit",
    "mw_mc_is", "mw_mc_isoos",
    "mw_maint_margin", "mw_init_margin",
    "is_win_rate", "oos_win_rate", "overall_win_rate",
    "trades_per_year", "pct_time_in_market", "avg_trade_length",
    "avg_trade", "avg_profitable_trade", "avg_loss_trade",
    "largest_win", "largest_loss",
    "max_drawdown_is", "max_drawdown_isoos", "avg_drawdown_is", "avg_drawdown_isoos",
    "sharpe_is", "sharpe_isoos", "annual_sd_is", "annual_sd_isoos",
    "oos_period_years",
]
_WF_STR_COLS = [
    "symbol", "timeframe", "session", "in_period", "out_period", "anchored", "fitness",
]
_WF_DATE_COLS = ["is_begin", "oos_begin", "oos_end", "next_opt_date", "last_opt_date"]
_WF_INT_COLS = ["trading_days_is", "trading_days_isoos"]

# New dynamic columns added outside WF (always computed)
_DYNAMIC_EXTRA_COLS = ["direction", "last_date_on_file", "mc_closed_is", "mc_closed_isoos"]


# ── Eligibility evaluation ────────────────────────────────────────────────────

def apply_eligibility_rules(
    summary_df: pd.DataFrame,
    eligibility: EligibilityConfig,
) -> pd.Series:
    """
    Apply all configured eligibility rules to the summary DataFrame.
    Returns a boolean Series (True = eligible) indexed by strategy name.

    Mirrors VBA EligibilityCheck() logic from F_Summary_Tab_Setup.bas.
    All enabled qualifiers must be satisfied; any enabled disqualifier voids eligibility.
    """
    eligible = pd.Series(True, index=summary_df.index)
    ratio = eligibility.efficiency_ratio

    # ── Buy & Hold exclusion gate ─────────────────────────────────────────────
    if eligibility.exclude_buy_and_hold and "status" in summary_df.columns:
        _status_lc = summary_df["status"].fillna("").str.lower()
        _is_bh = _status_lc.str.contains("buy") & _status_lc.str.contains("hold")
        eligible &= ~_is_bh

    # ── Previously-quit exclusion gate ───────────────────────────────────────
    # quitting_date being non-null means the strategy ever hit a quit threshold
    if eligibility.exclude_previously_quit and "quitting_date" in summary_df.columns:
        eligible &= summary_df["quitting_date"].isna()

    def _get(col: str) -> pd.Series:
        """Return column values preserving NaN — missing = insufficient OOS history."""
        if col in summary_df.columns:
            return summary_df[col]
        return pd.Series(np.nan, index=summary_df.index)

    def _val(row_series: pd.Series, threshold: float, op: str) -> pd.Series:
        if op == ">":
            return row_series > threshold
        return row_series >= threshold

    # ── Qualifiers (must all pass if enabled) ────────────────────────────────
    # NaN means insufficient OOS history: mirrors VBA IsNumeric() guard which
    # SKIPS the check (treats as passing) when the cell is blank.
    qualifier_map = {
        "profit_1m":  ("profit_last_1_month",  ">", 0.0),
        "profit_3m":  ("profit_last_3_months", ">", 0.0),
        "profit_6m":  ("profit_last_6_months", ">", 0.0),
        "profit_9m":  ("profit_last_9_months", ">", 0.0),
        "profit_12m": ("profit_last_12_months", ">", 0.0),
        "profit_oos": ("profit_since_oos_start", ">", 0.0),
        "efficiency_1m":  ("efficiency_last_1_month",  ">", ratio),
        "efficiency_3m":  ("efficiency_last_3_months", ">", ratio),
        "efficiency_6m":  ("efficiency_last_6_months", ">", ratio),
        "efficiency_9m":  ("efficiency_last_9_months", ">", ratio),
        "efficiency_12m": ("efficiency_last_12_months", ">", ratio),
        "efficiency_oos": ("return_efficiency", ">", ratio),
    }
    for attr, (col, op, threshold) in qualifier_map.items():
        if getattr(eligibility, attr, False):
            col_data = _get(col)
            # NaN → skip check (don't penalise for missing history).
            # pandas NaN comparisons return False (not NaN), so we must use isna().
            eligible &= _val(col_data, threshold, op) | col_data.isna()

    # ── Special: 3M OR 6M profit ──────────────────────────────────────────────
    # Mirrors VBA: disqualify only when BOTH v3M <= 0 AND v6M < 0 (and both numeric).
    # Note VBA uses strict < for 6M (v6M = 0 passes). NaN on either → skip check.
    if eligibility.profit_3or6m:
        p3 = _get("profit_last_3_months")
        p6 = _get("profit_last_6_months")
        both_present = p3.notna() & p6.notna()
        disqualify_3or6 = both_present & (p3 <= 0) & (p6 < 0)
        eligible &= ~disqualify_3or6

    # ── Disqualifiers (any triggered → ineligible) ────────────────────────────
    # NaN = insufficient history = don't disqualify (pandas NaN < 0 → False already).
    disqualifier_map = {
        "loss_1m":  "profit_last_1_month",
        "loss_3m":  "profit_last_3_months",
        "loss_6m":  "profit_last_6_months",
    }
    for attr, col in disqualifier_map.items():
        if getattr(eligibility, attr, False):
            eligible &= ~(_get(col) < 0).fillna(False)

    eff_loss_map = {
        "efficiency_loss_1m": "efficiency_last_1_month",
        "efficiency_loss_3m": "efficiency_last_3_months",
        "efficiency_loss_6m": "efficiency_last_6_months",
    }
    for attr, col in eff_loss_map.items():
        if getattr(eligibility, attr, False):
            eligible &= ~(_get(col) < -ratio).fillna(False)

    # ── Incubation gate ───────────────────────────────────────────────────────
    # Mirrors VBA: only "Passed" is accepted. "" (no expected profit / no OOS data)
    # is treated the same as "Incubating" — ineligible until explicitly passed.
    if eligibility.use_incubation:
        inc_status = summary_df["incubation_status"] if "incubation_status" in summary_df.columns else pd.Series("", index=summary_df.index)
        eligible &= inc_status == "Passed"

    # ── Quitting gate ─────────────────────────────────────────────────────────
    # VBA blocks "Quit" only. Python also blocks "Coming Back" (strategy hit the
    # quit threshold but is tentatively recovering — intentionally more conservative).
    if eligibility.use_quitting:
        quit_status = summary_df["quitting_status"] if "quitting_status" in summary_df.columns else pd.Series("", index=summary_df.index)
        eligible &= ~quit_status.isin(["Quit", "Coming Back"])

    # ── Count profitable months ───────────────────────────────────────────────
    # Mirrors VBA: disqualify when count < min (i.e. eligible when count >= min).
    # monthly_profit_operator controls how count_profit_months was tallied (>0 vs >=0),
    # not the comparison direction here — always use >=.
    if eligibility.use_count_monthly_profits:
        count_months = _get("count_profit_months").fillna(0)
        eligible &= count_months >= eligibility.min_positive_months

    return eligible
