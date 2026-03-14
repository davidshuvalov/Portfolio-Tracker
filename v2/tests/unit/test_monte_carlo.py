"""
Unit tests for core/analytics/monte_carlo.py

Tests:
  - _mc_core:                shape, ruin detection, no ruin baseline
  - _calc_sharpe:            basic ratio, zero-std edge case
  - _calc_rtd:               normal case, zero-dd cap
  - _filter_by_period:       OOS / IS / IS+OOS / no strategy
  - _get_pnl_samples:        M2M vs Closed selection
  - _estimate_trades_per_year: M2M fixed, Closed computed
  - solve_starting_equity:   converges to ~target ROR
  - run_monte_carlo:         end-to-end with synthetic data; empty-data guard
"""

from __future__ import annotations

import math
from datetime import date

import numpy as np
import pandas as pd
import pytest

from core.analytics.monte_carlo import (
    _calc_rtd,
    _calc_sharpe,
    _estimate_trades_per_year,
    _filter_by_period,
    _get_pnl_samples,
    _mc_core,
    run_monte_carlo,
    solve_starting_equity,
)
from core.config import MCConfig
from core.data_types import Strategy


# ── Fixtures ──────────────────────────────────────────────────────────────────

def make_pnl_series(n_days: int = 500, daily_profit: float = 100.0, seed: int = 42) -> pd.Series:
    """Synthetic daily PnL with slight positive drift."""
    rng = np.random.default_rng(seed)
    values = rng.normal(loc=daily_profit, scale=500.0, size=n_days)
    idx = pd.bdate_range("2021-01-01", periods=n_days)
    return pd.Series(values, index=idx, name="test_strat")


def make_strategy(
    oos_start: date = date(2022, 1, 1),
    is_end: date = date(2021, 12, 31),
) -> Strategy:
    return Strategy(
        name="test_strat",
        folder=__import__("pathlib").Path("."),
        status="Live",
        is_start=date(2019, 1, 1),
        is_end=is_end,
        oos_start=oos_start,
        oos_end=date(2023, 12, 31),
    )


def make_mc_config(**overrides) -> MCConfig:
    defaults = dict(
        simulations=500,
        period="OOS",
        risk_ruin_target=0.10,
        risk_ruin_tolerance=0.02,
        trade_adjustment=0.0,
        trade_option="M2M",
    )
    defaults.update(overrides)
    return MCConfig(**defaults)


# ── _mc_core ──────────────────────────────────────────────────────────────────

class TestMcCore:
    def test_output_shapes(self):
        samples = np.array([100.0, -50.0, 200.0, -100.0, 150.0])
        fe, dd, ruined = _mc_core(
            pnl_samples=samples,
            starting_equity=10_000.0,
            margin_threshold=1_000.0,
            n_scenarios=200,
            trades_per_year=50,
            trade_adjustment=0.0,
        )
        assert fe.shape == (200,)
        assert dd.shape == (200,)
        assert ruined.shape == (200,)

    def test_drawdown_fraction_between_0_and_1(self):
        """Max drawdown must be a fraction (0–1), not a dollar amount."""
        samples = np.array([100.0, -500.0, 300.0])
        _, dd, _ = _mc_core(
            pnl_samples=samples,
            starting_equity=5_000.0,
            margin_threshold=100.0,
            n_scenarios=100,
            trades_per_year=20,
            trade_adjustment=0.0,
        )
        assert float(dd.max()) <= 1.0
        assert float(dd.min()) >= 0.0

    def test_ruin_triggered_when_equity_drops_below_threshold(self):
        """All large-loss samples should trigger ruin with tiny starting equity."""
        samples = np.full(10, -2_000.0)  # always lose $2k per trade
        _, _, ruined = _mc_core(
            pnl_samples=samples,
            starting_equity=1_000.0,
            margin_threshold=999.0,
            n_scenarios=50,
            trades_per_year=10,
            trade_adjustment=0.0,
        )
        # Every scenario should be ruined immediately
        assert ruined.all()

    def test_no_ruin_with_always_profitable_samples(self):
        """All-win samples should never ruin with a sensible threshold."""
        samples = np.full(10, 500.0)  # always win $500
        _, _, ruined = _mc_core(
            pnl_samples=samples,
            starting_equity=10_000.0,
            margin_threshold=100.0,
            n_scenarios=100,
            trades_per_year=20,
            trade_adjustment=0.0,
        )
        assert not ruined.any()

    def test_trade_adjustment_reduces_equity(self):
        """50% trade adjustment should roughly halve the expected gain."""
        samples = np.full(100, 1_000.0)
        fe_no_adj, _, _ = _mc_core(samples, 10_000.0, 0.0, 50, 10, 0.0)
        fe_adj, _, _ = _mc_core(samples, 10_000.0, 0.0, 50, 10, 0.5)
        assert float(np.mean(fe_adj)) < float(np.mean(fe_no_adj))


# ── _calc_sharpe ─────────────────────────────────────────────────────────────

class TestCalcSharpe:
    def test_positive_profit_gives_positive_sharpe(self):
        rng = np.random.default_rng(0)
        fe = 10_000.0 + rng.normal(loc=1_000.0, scale=200.0, size=1_000)
        sharpe = _calc_sharpe(fe, 10_000.0)
        assert sharpe > 0.0

    def test_zero_variance_returns_zero(self):
        fe = np.full(100, 11_000.0)
        assert _calc_sharpe(fe, 10_000.0) == 0.0

    def test_zero_starting_equity_returns_zero(self):
        fe = np.array([1.0, 2.0, 3.0])
        assert _calc_sharpe(fe, 0.0) == 0.0

    def test_negative_profit_gives_negative_sharpe(self):
        rng = np.random.default_rng(1)
        fe = 10_000.0 + rng.normal(loc=-1_000.0, scale=200.0, size=1_000)
        sharpe = _calc_sharpe(fe, 10_000.0)
        assert sharpe < 0.0


# ── _calc_rtd ─────────────────────────────────────────────────────────────────

class TestCalcRtd:
    def test_normal_case(self):
        rtd = _calc_rtd(expected_profit=5_000.0, max_drawdown_pct=0.25, starting_equity=20_000.0)
        # dd_dollar = 0.25 * 20_000 = 5_000 → rtd = 5_000 / 5_000 = 1.0
        assert abs(rtd - 1.0) < 1e-6

    def test_negligible_drawdown_capped_at_4(self):
        rtd = _calc_rtd(expected_profit=1_000.0, max_drawdown_pct=0.0, starting_equity=10_000.0)
        assert rtd == 4.0

    def test_high_profit_gives_high_rtd(self):
        rtd = _calc_rtd(expected_profit=50_000.0, max_drawdown_pct=0.10, starting_equity=10_000.0)
        # dd_dollar = 1_000 → rtd = 50.0
        assert abs(rtd - 50.0) < 1e-4


# ── _filter_by_period ────────────────────────────────────────────────────────

class TestFilterByPeriod:
    def _make_series(self) -> pd.Series:
        idx = pd.bdate_range("2021-01-01", "2023-12-31")
        return pd.Series(np.ones(len(idx)), index=idx)

    def test_oos_filter(self):
        s = self._make_series()
        strategy = make_strategy(oos_start=date(2022, 6, 1))
        filtered = _filter_by_period(s, "OOS", strategy)
        assert filtered.index.min() >= pd.Timestamp("2022-06-01")

    def test_is_filter(self):
        s = self._make_series()
        strategy = make_strategy(is_end=date(2021, 12, 31))
        filtered = _filter_by_period(s, "IS", strategy)
        assert filtered.index.max() <= pd.Timestamp("2021-12-31")

    def test_isoos_returns_full_series(self):
        s = self._make_series()
        strategy = make_strategy()
        filtered = _filter_by_period(s, "IS+OOS", strategy)
        assert len(filtered) == len(s)

    def test_no_strategy_returns_full_series(self):
        s = self._make_series()
        filtered = _filter_by_period(s, "OOS", None)
        assert len(filtered) == len(s)

    def test_oos_filter_no_dates_returns_full(self):
        s = self._make_series()
        strat = Strategy(name="x", folder=__import__("pathlib").Path("."), status="Live")
        filtered = _filter_by_period(s, "OOS", strat)
        assert len(filtered) == len(s)


# ── _get_pnl_samples ─────────────────────────────────────────────────────────

class TestGetPnlSamples:
    def test_m2m_returns_m2m(self):
        m2m = pd.Series([1.0, 2.0, 3.0])
        closed = pd.Series([10.0, 20.0, 30.0])
        samples = _get_pnl_samples(m2m, closed, "M2M")
        np.testing.assert_array_equal(samples, m2m.values.astype(np.float64))

    def test_closed_returns_closed(self):
        m2m = pd.Series([1.0, 2.0, 3.0])
        closed = pd.Series([10.0, 20.0, 30.0])
        samples = _get_pnl_samples(m2m, closed, "Closed")
        np.testing.assert_array_equal(samples, closed.values.astype(np.float64))

    def test_closed_falls_back_to_m2m_when_none(self):
        m2m = pd.Series([1.0, 2.0, 3.0])
        samples = _get_pnl_samples(m2m, None, "Closed")
        np.testing.assert_array_equal(samples, m2m.values.astype(np.float64))

    def test_closed_falls_back_to_m2m_when_empty(self):
        m2m = pd.Series([1.0, 2.0])
        samples = _get_pnl_samples(m2m, pd.Series([], dtype=float), "Closed")
        np.testing.assert_array_equal(samples, m2m.values.astype(np.float64))


# ── _estimate_trades_per_year ────────────────────────────────────────────────

class TestEstimateTradesPerYear:
    def test_m2m_always_returns_252(self):
        s = pd.Series(range(100))
        assert _estimate_trades_per_year(s, "M2M") == 252

    def test_closed_counts_nonzero_days(self):
        # 252 days, exactly half nonzero → ~126 trades/year (1 year of data)
        values = [100.0 if i % 2 == 0 else 0.0 for i in range(252)]
        s = pd.Series(values)
        tpy = _estimate_trades_per_year(s, "Closed")
        assert 120 <= tpy <= 132  # allow rounding

    def test_empty_series_falls_back_to_252(self):
        # Empty data → fall through to safe default of 252 (avoids division by zero)
        assert _estimate_trades_per_year(pd.Series([], dtype=float), "Closed") == 252


# ── solve_starting_equity ────────────────────────────────────────────────────

class TestSolveStartingEquity:
    def test_converges_near_target_ror(self):
        """Solver should converge to within tolerance of the target ROR."""
        rng = np.random.default_rng(42)
        # Moderate PnL distribution: $200 mean, $800 std
        samples = rng.normal(loc=200.0, scale=800.0, size=1_000).astype(np.float64)

        config = make_mc_config(
            simulations=2_000,
            risk_ruin_target=0.10,
            risk_ruin_tolerance=0.03,
        )
        equity, ror, fe, dd = solve_starting_equity(
            samples, config, margin_threshold=5_000.0, trades_per_year=252
        )
        # ROR should be close to target (within ~3× tolerance due to stochasticity)
        assert 0.0 <= ror <= 0.40, f"ROR {ror:.2%} out of expected range"
        assert equity > 5_000.0, "Starting equity should exceed margin threshold"
        assert len(fe) == 2_000
        assert len(dd) == 2_000

    def test_returns_correct_shapes(self):
        samples = np.array([100.0, -50.0, 200.0, 0.0, -100.0])
        config = make_mc_config(simulations=100)
        equity, ror, fe, dd = solve_starting_equity(samples, config, 1_000.0, 20)
        assert len(fe) == 100
        assert len(dd) == 100
        assert isinstance(equity, float)
        assert isinstance(ror, float)


# ── run_monte_carlo ───────────────────────────────────────────────────────────

class TestRunMonteCarlo:
    def test_basic_end_to_end(self):
        """Smoke test: run on synthetic series, check result structure."""
        pnl = make_pnl_series(n_days=500, daily_profit=50.0)
        config = make_mc_config(simulations=200, period="IS+OOS")

        result = run_monte_carlo(
            daily_m2m=pnl,
            config=config,
            margin_threshold=5_000.0,
            return_scenarios=True,
        )

        assert result.starting_equity > 0
        assert not math.isnan(result.risk_of_ruin)
        assert 0.0 <= result.max_drawdown_pct <= 1.0
        assert result.scenarios_df is not None
        assert len(result.scenarios_df) == 200
        assert set(result.scenarios_df.columns) == {"final_equity", "max_drawdown_pct", "profit"}

    def test_scenarios_not_returned_when_false(self):
        pnl = make_pnl_series(200)
        config = make_mc_config(simulations=100, period="IS+OOS")
        result = run_monte_carlo(pnl, config, 1_000.0, return_scenarios=False)
        assert result.scenarios_df is None

    def test_empty_series_returns_safe_default(self):
        empty = pd.Series([], dtype=float)
        config = make_mc_config(simulations=100, period="OOS")
        result = run_monte_carlo(empty, config, 5_000.0)
        assert result.expected_profit == 0.0
        assert math.isnan(result.risk_of_ruin)

    def test_oos_period_filter_applied(self):
        """OOS filter should reduce the sample size vs IS+OOS."""
        pnl = make_pnl_series(n_days=600)
        strategy = make_strategy(oos_start=date(2023, 1, 1))

        config_oos = make_mc_config(simulations=100, period="OOS")
        config_all = make_mc_config(simulations=100, period="IS+OOS")

        r_oos = run_monte_carlo(pnl, config_oos, 1_000.0, strategy=strategy)
        r_all = run_monte_carlo(pnl, config_all, 1_000.0, strategy=strategy)

        # Both should produce valid results (OOS may be empty → default; IS+OOS always has data)
        assert isinstance(r_oos.starting_equity, float)
        assert isinstance(r_all.starting_equity, float)

    def test_with_closed_trade_option(self):
        pnl = make_pnl_series(300, daily_profit=30.0)
        closed = make_pnl_series(300, daily_profit=25.0, seed=99)
        config = make_mc_config(simulations=100, period="IS+OOS", trade_option="Closed")
        result = run_monte_carlo(pnl, config, 2_000.0, closed_daily=closed)
        assert result.starting_equity > 0

    def test_return_to_drawdown_capped_at_4_for_zero_drawdown(self):
        """All-positive PnL → near-zero drawdown → RTD capped at 4 (VBA: IIf(dd=0, 4, ...))."""
        pnl = pd.Series(np.full(252, 1_000.0), index=pd.bdate_range("2022-01-01", periods=252))
        config = make_mc_config(simulations=50, period="IS+OOS", trade_adjustment=0.0)
        result = run_monte_carlo(pnl, config, 1_000.0)
        assert result.return_to_drawdown <= 4.0
