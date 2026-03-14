"""
Unit tests for core/analytics/correlations.py

Tests:
  - compute_correlation_matrix: NORMAL, NEGATIVE, DRAWDOWN modes
  - _pairwise_correlation: edge cases (constant array, insufficient data)
  - _to_drawdown_series: basic shape and value properties
  - get_correlation_pairs: ordering, count
  - flag_high_correlations: threshold filtering
  - average_correlation: mean value, single-strategy guard
  - compute_all_modes: returns all three keys
"""

from __future__ import annotations

import math

import numpy as np
import pandas as pd
import pytest

from core.analytics.correlations import (
    CorrelationMode,
    _pairwise_correlation,
    _to_drawdown_series,
    average_correlation,
    compute_all_modes,
    compute_correlation_matrix,
    flag_high_correlations,
    get_correlation_pairs,
)


# ── Fixtures ──────────────────────────────────────────────────────────────────

def make_daily_pnl(n_days: int = 200, n_strats: int = 3, seed: int = 0) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    idx = pd.bdate_range("2022-01-01", periods=n_days)
    data = rng.normal(0, 500, (n_days, n_strats))
    cols = [f"S{i+1}" for i in range(n_strats)]
    return pd.DataFrame(data, index=idx, columns=cols)


def make_perfect_pos_pair(n: int = 100) -> pd.DataFrame:
    """Two strategies with correlation = +1."""
    idx = pd.bdate_range("2022-01-01", periods=n)
    a = np.linspace(-500, 500, n)
    return pd.DataFrame({"A": a, "B": a * 2}, index=idx)


def make_perfect_neg_pair(n: int = 100) -> pd.DataFrame:
    """Two strategies with correlation = -1."""
    idx = pd.bdate_range("2022-01-01", periods=n)
    a = np.linspace(-500, 500, n)
    return pd.DataFrame({"A": a, "B": -a}, index=idx)


# ── _to_drawdown_series ───────────────────────────────────────────────────────

class TestToDrawdownSeries:
    def test_values_between_0_and_1(self):
        equity = np.array([100.0, 110.0, 105.0, 90.0, 115.0])
        dd = _to_drawdown_series(equity)
        assert dd.min() >= 0.0
        assert dd.max() <= 1.0

    def test_peak_points_are_zero(self):
        equity = np.array([100.0, 110.0, 120.0, 115.0])
        dd = _to_drawdown_series(equity)
        # Index 2 is the all-time peak → drawdown = 0
        assert dd[2] == 0.0

    def test_flat_equity_is_all_zeros(self):
        equity = np.full(10, 1000.0)
        dd = _to_drawdown_series(equity)
        np.testing.assert_array_equal(dd, np.zeros(10))

    def test_monotonically_rising_equity_is_zero_dd(self):
        equity = np.cumsum(np.ones(20) * 100)
        dd = _to_drawdown_series(equity)
        np.testing.assert_allclose(dd, 0.0, atol=1e-9)

    def test_zero_peak_guard(self):
        """Should not raise for equity arrays starting at or below zero."""
        equity = np.array([-100.0, -50.0, -10.0, 10.0])
        dd = _to_drawdown_series(equity)
        assert all(dd >= 0.0)


# ── _pairwise_correlation ─────────────────────────────────────────────────────

class TestPairwiseCorrelation:
    def test_perfect_positive_normal(self):
        a = np.array([1.0, 2.0, 3.0, 4.0, 5.0])
        b = a * 3
        corr = _pairwise_correlation(a, b, CorrelationMode.NORMAL)
        assert abs(corr - 1.0) < 1e-9

    def test_perfect_negative_normal(self):
        a = np.array([1.0, 2.0, 3.0, 4.0, 5.0])
        b = -a
        corr = _pairwise_correlation(a, b, CorrelationMode.NORMAL)
        assert abs(corr + 1.0) < 1e-9

    def test_constant_array_returns_zero(self):
        a = np.ones(20)
        b = np.arange(20, dtype=float)
        corr = _pairwise_correlation(a, b, CorrelationMode.NORMAL)
        assert corr == 0.0

    def test_all_zeros_returns_nan(self):
        """Both arrays zero → mask excludes everything → insufficient data."""
        a = np.zeros(10)
        b = np.zeros(10)
        corr = _pairwise_correlation(a, b, CorrelationMode.NORMAL)
        assert math.isnan(corr)

    def test_negative_mode_excludes_both_positive(self):
        """Days where both > 0 are excluded; check the mask works."""
        a = np.array([100.0, -100.0, 50.0, -50.0])
        b = np.array([100.0, -100.0, -50.0, 50.0])
        corr = _pairwise_correlation(a, b, CorrelationMode.NEGATIVE)
        # Row 0 (both +) excluded; rows 1,2,3 kept
        # a_keep = [-100, 50, -50], b_keep = [-100, -50, 50]
        # corr of [-100,50,-50] with [-100,-50,50] — both in opposing signs
        assert not math.isnan(corr)

    def test_drawdown_mode_returns_float(self):
        rng = np.random.default_rng(5)
        a = rng.normal(50, 300, 100)
        b = rng.normal(30, 250, 100)
        corr = _pairwise_correlation(a, b, CorrelationMode.DRAWDOWN)
        assert isinstance(corr, float)
        assert -1.0 <= corr <= 1.0


# ── compute_correlation_matrix ────────────────────────────────────────────────

class TestComputeCorrelationMatrix:
    def test_shape_equals_n_by_n(self):
        df = make_daily_pnl(n_strats=5)
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        assert m.shape == (5, 5)

    def test_diagonal_is_one(self):
        df = make_daily_pnl()
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        np.testing.assert_allclose(np.diag(m.values), 1.0, atol=1e-9)

    def test_symmetric(self):
        df = make_daily_pnl()
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        np.testing.assert_allclose(m.values, m.values.T, atol=1e-9)

    def test_perfect_positive_pair(self):
        df = make_perfect_pos_pair()
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        assert abs(m.loc["A", "B"] - 1.0) < 1e-6

    def test_perfect_negative_pair(self):
        df = make_perfect_neg_pair()
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        assert abs(m.loc["A", "B"] + 1.0) < 1e-6

    def test_drawdown_mode(self):
        df = make_daily_pnl()
        m = compute_correlation_matrix(df, CorrelationMode.DRAWDOWN)
        assert m.shape == (3, 3)
        np.testing.assert_allclose(np.diag(m.values), 1.0, atol=1e-9)

    def test_negative_mode(self):
        df = make_daily_pnl()
        m = compute_correlation_matrix(df, CorrelationMode.NEGATIVE)
        assert m.shape == (3, 3)
        assert m.values.max() <= 1.001

    def test_columns_and_index_match(self):
        df = make_daily_pnl(n_strats=4)
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        assert list(m.columns) == list(df.columns)
        assert list(m.index) == list(df.columns)


# ── get_correlation_pairs ─────────────────────────────────────────────────────

class TestGetCorrelationPairs:
    def test_pair_count(self):
        df = make_daily_pnl(n_strats=5)
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        pairs = get_correlation_pairs(m)
        # n*(n-1)/2 = 10 pairs for 5 strategies
        assert len(pairs) == 10

    def test_sorted_descending(self):
        df = make_daily_pnl(n_strats=4)
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        pairs = get_correlation_pairs(m)
        corrs = pairs["correlation"].values
        assert all(corrs[i] >= corrs[i + 1] for i in range(len(corrs) - 1))

    def test_columns_present(self):
        df = make_daily_pnl()
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        pairs = get_correlation_pairs(m)
        assert set(pairs.columns) == {"strategy_a", "strategy_b", "correlation"}

    def test_single_strategy_returns_empty(self):
        df = pd.DataFrame({"A": [1.0, 2.0, 3.0]})
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        pairs = get_correlation_pairs(m)
        assert pairs.empty

    def test_no_self_pairs(self):
        df = make_daily_pnl(n_strats=4)
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        pairs = get_correlation_pairs(m)
        assert all(row["strategy_a"] != row["strategy_b"] for _, row in pairs.iterrows())


# ── flag_high_correlations ────────────────────────────────────────────────────

class TestFlagHighCorrelations:
    def test_perfect_positive_exceeds_any_threshold(self):
        df = make_perfect_pos_pair()
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        high = flag_high_correlations(m, 0.99)
        assert len(high) == 1
        a, b, corr = high[0]
        assert abs(corr - 1.0) < 1e-6

    def test_threshold_1_returns_empty(self):
        df = make_daily_pnl()
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        high = flag_high_correlations(m, 1.0)
        # No pair can have |r| exactly == 1 with random data
        assert len(high) == 0

    def test_threshold_0_returns_all_pairs(self):
        df = make_daily_pnl(n_strats=4)
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        high = flag_high_correlations(m, 0.0)
        assert len(high) == 6  # 4*(4-1)/2

    def test_result_is_list_of_tuples(self):
        df = make_perfect_pos_pair()
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        high = flag_high_correlations(m, 0.5)
        assert isinstance(high, list)
        assert all(isinstance(t, tuple) and len(t) == 3 for t in high)


# ── average_correlation ───────────────────────────────────────────────────────

class TestAverageCorrelation:
    def test_single_strategy_returns_nan(self):
        df = pd.DataFrame({"A": [1.0, 2.0, 3.0]})
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        assert math.isnan(average_correlation(m))

    def test_perfect_positive_pair_avg_is_one(self):
        df = make_perfect_pos_pair()
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        assert abs(average_correlation(m) - 1.0) < 1e-6

    def test_perfect_negative_pair_avg_is_minus_one(self):
        df = make_perfect_neg_pair()
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        assert abs(average_correlation(m) + 1.0) < 1e-6

    def test_random_data_within_range(self):
        df = make_daily_pnl(n_strats=5)
        m = compute_correlation_matrix(df, CorrelationMode.NORMAL)
        avg = average_correlation(m)
        assert -1.0 <= avg <= 1.0


# ── compute_all_modes ─────────────────────────────────────────────────────────

class TestComputeAllModes:
    def test_returns_all_three_keys(self):
        df = make_daily_pnl()
        result = compute_all_modes(df)
        assert set(result.keys()) == {"normal", "negative", "drawdown"}

    def test_all_matrices_have_correct_shape(self):
        df = make_daily_pnl(n_strats=4)
        result = compute_all_modes(df)
        for mode, matrix in result.items():
            assert matrix.shape == (4, 4), f"Mode {mode} has wrong shape"

    def test_modes_can_differ(self):
        """NORMAL and NEGATIVE matrices should not always be identical."""
        df = make_daily_pnl(n_strats=3, seed=7)
        result = compute_all_modes(df)
        normal = result["normal"].values
        negative = result["negative"].values
        # Off-diagonal values may differ
        # (they will differ for random data with mixed signs)
        assert not np.allclose(normal, negative, atol=1e-3)
