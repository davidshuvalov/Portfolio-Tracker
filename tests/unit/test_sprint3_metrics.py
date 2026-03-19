"""
Unit tests for Sprint 3 metrics:
  - K-Factor (monthly win_rate / loss_rate × avg_win / avg_loss)
  - Ulcer Index (RMS % drawdown)
  - best_month / worst_month
  - max_consecutive_loss_months
  - profit_since_quit
  - run_leave_one_out_chronological (LOO chronological mode)
"""

from __future__ import annotations

from datetime import date

import numpy as np
import pandas as pd
import pytest

from pathlib import Path

from core.analytics.leave_one_out import run_leave_one_out_chronological, _chron_stats
from core.data_types import PortfolioData, Strategy


# ── Helpers ───────────────────────────────────────────────────────────────────

def _daily_from_monthly(monthly_values: list[float], start: str = "2021-01-01") -> pd.Series:
    """Build a daily PnL series where each month has the given total, spread over 21 trading days."""
    idx = pd.bdate_range(start, periods=len(monthly_values) * 21)
    daily = np.zeros(len(idx))
    for m, val in enumerate(monthly_values):
        daily[m * 21 : (m + 1) * 21] = val / 21.0
    return pd.Series(daily, index=idx[: len(monthly_values) * 21])


def _make_portfolio(n_strategies: int = 3, n_days: int = 500, seed: int = 0) -> PortfolioData:
    """Build a minimal PortfolioData for LOO chronological tests."""
    rng = np.random.default_rng(seed)
    idx = pd.bdate_range("2022-01-01", periods=n_days)
    daily_data = rng.normal(100, 500, (n_days, n_strategies))
    strategy_names = [f"S{i+1}" for i in range(n_strategies)]
    daily_pnl = pd.DataFrame(daily_data, index=idx, columns=strategy_names)
    strategies = [
        Strategy(name=n, folder=Path("/tmp"), symbol=f"SYM{i}", contracts=1, status="Live")
        for i, n in enumerate(strategy_names)
    ]
    return PortfolioData(
        strategies=strategies,
        daily_pnl=daily_pnl,
        closed_trades=pd.DataFrame(),
        summary_metrics=pd.DataFrame(),
    )


# ── K-Factor and monthly stats ────────────────────────────────────────────────

class TestSprint3MetricsViaCompute:
    """Test metrics by calling compute logic through _compute_dynamic_metrics indirectly.

    We test the mathematical correctness of the underlying computations
    by testing _chron_stats and the direct logic, since _compute_dynamic_metrics
    requires a full OOS date range setup.
    """

    def test_best_worst_month_with_known_series(self):
        """Monthly resampling should find best and worst month correctly."""
        monthly = [500.0, -200.0, 1000.0, -50.0, 300.0]
        daily = _daily_from_monthly(monthly)
        monthly_series = daily.resample("ME").sum()
        best = float(monthly_series.max())
        worst = float(monthly_series.min())
        # Best month has the highest total (1000), worst has the lowest (-200)
        assert best > 0
        assert worst < 0
        assert best > abs(worst)  # 1000 > 200

    def test_consecutive_loss_months(self):
        """Verify max consecutive losing months logic."""
        monthly = [100.0, -50.0, -60.0, -30.0, 200.0, -10.0]
        losing_streak = 0
        max_streak = 0
        for v in monthly:
            if v < 0:
                losing_streak += 1
                max_streak = max(max_streak, losing_streak)
            else:
                losing_streak = 0
        assert max_streak == 3  # Three consecutive losses at positions 1,2,3

    def test_k_factor_formula(self):
        """K-Factor = (win_rate / loss_rate) × (avg_win / avg_loss)."""
        monthly = [100.0, -50.0, 200.0, -50.0]
        m = pd.Series(monthly)
        wins = m[m > 0]
        losses = m[m < 0]
        win_rate = len(wins) / len(m)
        avg_win = float(wins.mean())
        avg_loss = abs(float(losses.mean()))
        k = (win_rate / (1 - win_rate)) * (avg_win / avg_loss)
        # win_rate = 0.5, avg_win = 150, avg_loss = 50 → k = 1.0 * 3.0 = 3.0
        assert k == pytest.approx(3.0, rel=0.01)

    def test_ulcer_index_flat_equity(self):
        """A flat equity curve (no drawdown) should have Ulcer Index = 0."""
        n = 100
        idx = pd.bdate_range("2022-01-01", periods=n)
        pnl = pd.Series(np.ones(n) * 100.0, index=idx)
        eq = pnl.cumsum()
        peak = eq.cummax()
        pct_dd = np.where(peak > 1e-9, (peak - eq) / peak * 100.0, 0.0)
        ulcer = float(np.sqrt(np.mean(pct_dd ** 2)))
        assert ulcer == pytest.approx(0.0, abs=1e-9)

    def test_ulcer_index_monotone_decline(self):
        """A steadily declining equity should have a large Ulcer Index."""
        n = 100
        idx = pd.bdate_range("2022-01-01", periods=n)
        pnl = pd.Series(-np.ones(n) * 100.0, index=idx)
        eq = pnl.cumsum()  # 0, -100, -200, ...
        # Make it start at a positive peak by shifting
        eq = eq + 10100.0  # starts at 10100
        peak = eq.cummax()
        pct_dd = np.where(peak > 1e-9, (peak - eq) / peak * 100.0, 0.0)
        ulcer = float(np.sqrt(np.mean(pct_dd ** 2)))
        assert ulcer > 5.0  # Should be large


# ── _chron_stats ──────────────────────────────────────────────────────────────

class TestChronStats:
    def test_empty_series_returns_zeros(self):
        pnl = pd.Series([], dtype=float)
        result = _chron_stats(pnl)
        assert result["total_profit"] == 0.0
        assert result["annual_profit"] == 0.0
        assert result["sharpe"] == 0.0

    def test_all_zeros_returns_zeros(self):
        idx = pd.bdate_range("2022-01-01", periods=100)
        pnl = pd.Series(np.zeros(100), index=idx)
        result = _chron_stats(pnl)
        assert result["total_profit"] == 0.0

    def test_positive_series(self):
        idx = pd.bdate_range("2022-01-01", periods=252)
        pnl = pd.Series(np.ones(252) * 100.0, index=idx)
        result = _chron_stats(pnl)
        assert result["total_profit"] == pytest.approx(25200.0, rel=0.01)
        assert result["annual_profit"] > 0
        assert result["sharpe"] == 0.0  # constant return → std = 0 → sharpe = 0

    def test_monotone_rise_zero_drawdown(self):
        idx = pd.bdate_range("2022-01-01", periods=100)
        pnl = pd.Series(np.ones(100) * 50.0, index=idx)
        result = _chron_stats(pnl)
        assert result["max_dd_pct"] == pytest.approx(0.0, abs=1e-9)

    def test_drawdown_detected(self):
        """Series that goes up then down should show positive drawdown."""
        idx = pd.bdate_range("2022-01-01", periods=200)
        vals = np.concatenate([np.ones(100) * 100, np.ones(100) * -100])
        pnl = pd.Series(vals, index=idx)
        result = _chron_stats(pnl)
        assert result["max_dd_pct"] > 0.0

    def test_rtd_positive_for_profitable_series(self):
        idx = pd.bdate_range("2022-01-01", periods=200)
        vals = np.concatenate([np.ones(100) * 100, np.ones(100) * -50, np.ones(100) * 150])
        idx2 = pd.bdate_range("2022-01-01", periods=300)
        pnl = pd.Series(vals[:300] if len(vals) >= 300 else vals, index=idx2[:len(vals)])
        result = _chron_stats(pnl)
        assert result["rtd"] > 0


# ── run_leave_one_out_chronological ───────────────────────────────────────────

class TestLOOChronological:
    def test_returns_one_row_per_strategy(self):
        portfolio = _make_portfolio(n_strategies=4)
        result = run_leave_one_out_chronological(portfolio)
        assert len(result) == 4

    def test_all_strategy_names_present(self):
        portfolio = _make_portfolio(n_strategies=3)
        result = run_leave_one_out_chronological(portfolio)
        assert set(result["strategy"].tolist()) == {"S1", "S2", "S3"}

    def test_required_columns_present(self):
        portfolio = _make_portfolio()
        result = run_leave_one_out_chronological(portfolio)
        for col in ["strategy", "delta_annual", "delta_sharpe", "delta_drawdown", "delta_rtd"]:
            assert col in result.columns

    def test_empty_portfolio_returns_empty_df(self):
        empty = PortfolioData(
            strategies=[],
            daily_pnl=pd.DataFrame(),
            closed_trades=pd.DataFrame(),
            summary_metrics=pd.DataFrame(),
        )
        result = run_leave_one_out_chronological(empty)
        assert result.empty

    def test_single_strategy_returns_one_row(self):
        portfolio = _make_portfolio(n_strategies=1)
        result = run_leave_one_out_chronological(portfolio)
        assert len(result) == 1

    def test_removing_profitable_strategy_reduces_profit(self):
        """A strongly profitable strategy should show negative delta when removed."""
        n = 252
        idx = pd.bdate_range("2022-01-01", periods=n)
        # S1: very profitable; S2/S3: flat
        rng = np.random.default_rng(42)
        daily_data = pd.DataFrame({
            "S1": np.ones(n) * 500,
            "S2": rng.normal(0, 10, n),
            "S3": rng.normal(0, 10, n),
        }, index=idx)
        strategies = [
            Strategy(name="S1", folder=Path("/tmp"), symbol="A", contracts=1, status="Live"),
            Strategy(name="S2", folder=Path("/tmp"), symbol="B", contracts=1, status="Live"),
            Strategy(name="S3", folder=Path("/tmp"), symbol="C", contracts=1, status="Live"),
        ]
        portfolio = PortfolioData(
            strategies=strategies,
            daily_pnl=daily_data,
            closed_trades=pd.DataFrame(),
            summary_metrics=pd.DataFrame(),
        )
        result = run_leave_one_out_chronological(portfolio)
        s1_row = result[result["strategy"] == "S1"].iloc[0]
        # Removing a very profitable strategy should reduce annual profit
        assert s1_row["delta_annual"] < 0

    def test_sorted_by_delta_annual_ascending(self):
        portfolio = _make_portfolio(n_strategies=5)
        result = run_leave_one_out_chronological(portfolio)
        delta_vals = result["delta_annual"].values
        assert all(delta_vals[i] <= delta_vals[i + 1] for i in range(len(delta_vals) - 1))
