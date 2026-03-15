"""
Monte Carlo simulation — mirrors K_MonteCarlo.bas.

Core loop:  _mc_core()               (Numba JIT, ~600x faster than VBA)
Solver:     solve_starting_equity()  (iterative ROR targeting, 100-iter cap)
Public API: run_monte_carlo()

VBA equivalents:
  K_RunMonteCarlo         → run_monte_carlo()
  K_SolveStartingEquity   → solve_starting_equity()
  K_MCInnerLoop           → _mc_core()
"""

from __future__ import annotations

from datetime import date

import numpy as np
import pandas as pd

from core.config import MCConfig
from core.data_types import MCResult, Strategy

# ── Numba JIT (graceful fallback if not installed) ────────────────────────────
try:
    from numba import njit
    _NUMBA_AVAILABLE = True
except ImportError:  # pragma: no cover
    _NUMBA_AVAILABLE = False

    def njit(*args, **kwargs):  # type: ignore[misc]
        def decorator(func):
            return func
        return decorator


# ── Inner loop ────────────────────────────────────────────────────────────────

@njit(cache=True)
def _mc_core(
    pnl_samples: np.ndarray,
    starting_equity: float,
    margin_threshold: float,
    n_scenarios: int,
    trades_per_year: int,
    trade_adjustment: float,
) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """
    Compiled MC inner loop. ~50ms for 10k scenarios vs VBA's ~30s.

    Returns:
        final_equity      shape (n_scenarios,)  — equity after one year
        max_drawdown_pct  shape (n_scenarios,)  — peak-to-trough fraction (0–1)
        ruined            shape (n_scenarios,)  — bool, equity fell below threshold
    """
    final_equity = np.empty(n_scenarios)
    max_drawdown = np.empty(n_scenarios)
    ruined = np.zeros(n_scenarios, dtype=np.bool_)

    for i in range(n_scenarios):
        equity = starting_equity
        peak = starting_equity
        dd = 0.0

        for j in range(trades_per_year):
            idx = np.random.randint(0, len(pnl_samples))
            equity += pnl_samples[idx] * (1.0 - trade_adjustment)

            if equity > peak:
                peak = equity
            raw_dd = (peak - equity) / peak if peak > 1e-9 else 0.0
            drawdown = raw_dd if raw_dd < 1.0 else 1.0
            if drawdown > dd:
                dd = drawdown
            if equity < margin_threshold:
                ruined[i] = True
                break

        final_equity[i] = equity
        max_drawdown[i] = dd

    return final_equity, max_drawdown, ruined


# ── Helper metrics ────────────────────────────────────────────────────────────

def _calc_sharpe(final_equity: np.ndarray, starting_equity: float) -> float:
    """
    Sharpe ratio from MC scenario outcomes (1-year horizon = already annualised).
    R_i = (final_equity_i - starting_equity) / starting_equity
    Sharpe = mean(R) / std(R)
    """
    if starting_equity < 1e-9 or len(final_equity) < 2:
        return 0.0
    returns = (final_equity - starting_equity) / starting_equity
    std_r = float(np.std(returns))
    if std_r < 1e-9:
        return 0.0
    return float(np.mean(returns)) / std_r


def _calc_rtd(
    expected_profit: float,
    max_drawdown_pct: float,
    starting_equity: float,
) -> float:
    """
    Return-to-drawdown: expected_annual_profit / median_max_drawdown_$.
    Capped at 4.0 when drawdown is negligible (mirrors VBA: IIf(maxDrawdown=0, 4, ...)).
    """
    dd_dollar = max_drawdown_pct * starting_equity
    if dd_dollar < 1e-4:
        return 4.0
    return expected_profit / dd_dollar


# ── Period & sample helpers ───────────────────────────────────────────────────

def _filter_by_period(
    series: pd.Series,
    period: str,
    strategy: Strategy | None,
) -> pd.Series:
    """
    Filter daily PnL series to the requested period (IS / OOS / IS+OOS).
    Falls back to full series when strategy dates are unavailable.
    """
    if strategy is None or period == "IS+OOS":
        return series

    if period == "OOS" and strategy.oos_start is not None:
        return series.loc[series.index >= pd.Timestamp(strategy.oos_start)]

    if period == "IS" and strategy.is_end is not None:
        return series.loc[series.index <= pd.Timestamp(strategy.is_end)]

    return series


def _get_pnl_samples(
    daily_m2m: pd.Series,
    closed_daily: pd.Series | None,
    trade_option: str,
) -> np.ndarray:
    """
    Return the sample array based on trade option.

    M2M:    all daily M2M values (252 days/year, includes zeros)
    Closed: daily closed-trade PnL (non-zero days = trade days)
    """
    if trade_option == "Closed" and closed_daily is not None and not closed_daily.empty:
        return closed_daily.values.astype(np.float64)
    return daily_m2m.values.astype(np.float64)


def _estimate_trades_per_year(pnl_series: pd.Series, trade_option: str) -> int:
    """
    Estimate the number of samples to draw per simulated year.

    M2M:    fixed 252 (trading days)
    Closed: count non-zero trade days per year from historical data
    """
    if trade_option == "M2M" or len(pnl_series) == 0:
        return 252
    n_years = max(len(pnl_series) / 252.0, 1.0)
    trades_per_year = int(round((pnl_series != 0).sum() / n_years))
    return max(trades_per_year, 1)


# ── Iterative solver ──────────────────────────────────────────────────────────

def solve_starting_equity(
    pnl_samples: np.ndarray,
    config: MCConfig,
    margin_threshold: float,
    trades_per_year: int = 252,
) -> tuple[float, float, np.ndarray, np.ndarray]:
    """
    Iterative solver — mirrors VBA's +5% / -0.9% adjustment loop.
    Hard cap at 100 iterations (identical to VBA behaviour).

    Returns:
        (starting_equity, final_ror, final_equity_array, max_drawdown_array)
    """
    equity = margin_threshold * 2.0
    fe = np.array([equity], dtype=np.float64)
    dd = np.array([0.0], dtype=np.float64)
    ror = 1.0

    for _ in range(100):
        fe, dd, ruined = _mc_core(
            pnl_samples,
            equity,
            margin_threshold,
            config.simulations,
            trades_per_year,
            config.trade_adjustment,
        )
        ror = float(ruined.mean())
        if abs(ror - config.risk_ruin_target) < config.risk_ruin_tolerance:
            break
        equity *= 1.05 if ror > config.risk_ruin_target else 0.991

    return float(equity), ror, fe, dd


# ── Public API ────────────────────────────────────────────────────────────────

def run_monte_carlo(
    daily_m2m: pd.Series,
    config: MCConfig,
    margin_threshold: float,
    closed_daily: pd.Series | None = None,
    strategy: Strategy | None = None,
    return_scenarios: bool = False,
) -> MCResult:
    """
    Run Monte Carlo simulation on a daily PnL series.

    Args:
        daily_m2m:        Daily mark-to-market PnL (DatetimeIndex).
                          Pass portfolio total PnL for portfolio MC,
                          or a single strategy's column for per-strategy MC.
        config:           MCConfig with simulation parameters.
        margin_threshold: Dollar amount below which account = "ruined".
        closed_daily:     Daily closed-trade PnL series (for trade_option="Closed").
        strategy:         Strategy object for IS/OOS period filtering.
                          Pass None for portfolio-level MC (uses all data).
        return_scenarios: If True, attach full scenario DataFrame to result.

    Returns:
        MCResult with all summary metrics (+ scenarios_df if requested).
    """
    # Apply period filter
    m2m_filtered = _filter_by_period(daily_m2m, config.period, strategy)
    closed_filtered = (
        _filter_by_period(closed_daily, config.period, strategy)
        if closed_daily is not None
        else None
    )

    if len(m2m_filtered) == 0:
        return MCResult(
            starting_equity=float(margin_threshold * 2),
            expected_profit=0.0,
            risk_of_ruin=float("nan"),
            max_drawdown_pct=0.0,
            sharpe_ratio=0.0,
            return_to_drawdown=0.0,
        )

    pnl_samples = _get_pnl_samples(m2m_filtered, closed_filtered, config.trade_option)
    trades_per_year = _estimate_trades_per_year(m2m_filtered, config.trade_option)

    equity, ror, fe, dd = solve_starting_equity(
        pnl_samples, config, margin_threshold, trades_per_year
    )

    expected_profit = float(np.mean(fe) - equity)
    max_dd_pct = float(np.median(dd))
    sharpe = _calc_sharpe(fe, equity)
    rtd = _calc_rtd(expected_profit, max_dd_pct, equity)

    scenarios_df = None
    if return_scenarios:
        scenarios_df = pd.DataFrame({
            "final_equity": fe,
            "max_drawdown_pct": dd,
            "profit": fe - equity,
        })

    return MCResult(
        starting_equity=float(equity),
        expected_profit=expected_profit,
        risk_of_ruin=ror,
        max_drawdown_pct=max_dd_pct,
        sharpe_ratio=sharpe,
        return_to_drawdown=rtd,
        scenarios_df=scenarios_df,
    )
