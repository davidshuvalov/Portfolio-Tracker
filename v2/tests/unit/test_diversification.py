"""Unit tests for core.analytics.diversification."""
from __future__ import annotations

import numpy as np
import pandas as pd
import pytest

from core.analytics.diversification import (
    _compute_metrics,
    run_greedy_selection,
    run_randomized_analysis,
)


def _make_pnl(n_days: int = 500, seed: int = 0, scale: float = 100.0) -> pd.Series:
    rng = np.random.default_rng(seed)
    return pd.Series(rng.normal(0, scale, n_days))


def _make_df(n_strats: int = 4, n_days: int = 500) -> pd.DataFrame:
    data = {f"S{i+1}": _make_pnl(n_days, seed=i) for i in range(n_strats)}
    idx = pd.date_range("2020-01-01", periods=n_days, freq="B")
    return pd.DataFrame(data, index=idx)


# ── _compute_metrics ───────────────────────────────────────────────────────────

class TestComputeMetrics:
    def test_empty_returns_zeros(self):
        m = _compute_metrics(pd.Series([], dtype=float))
        assert m["rtd"] == 0.0
        assert m["sharpe"] == 0.0

    def test_flat_returns_zeros(self):
        m = _compute_metrics(pd.Series([0.0] * 100))
        assert m["rtd"] == 0.0

    def test_positive_trend_has_positive_annual_profit(self):
        pnl = pd.Series([10.0] * 252)
        m = _compute_metrics(pnl)
        assert m["annual_profit"] > 0

    def test_series_with_drawdown_has_positive_rtd(self):
        # Up, then partial pullback → non-zero max_dd so rtd should be positive
        pnl = pd.Series([10.0] * 100 + [-5.0] * 20 + [10.0] * 132)
        m = _compute_metrics(pnl)
        assert m["annual_profit"] > 0
        assert m["max_dd"] > 0
        assert m["rtd"] > 0

    def test_negative_trend_has_negative_rtd(self):
        pnl = pd.Series([-5.0] * 252)
        m = _compute_metrics(pnl)
        assert m["annual_profit"] < 0
        assert m["rtd"] < 0  # annual_profit negative / positive max_dd → negative

    def test_returns_all_keys(self):
        m = _compute_metrics(_make_pnl())
        for k in ("annual_profit", "max_dd", "avg_dd", "ann_std", "sharpe", "rtd", "rtd_avg"):
            assert k in m


# ── run_greedy_selection ───────────────────────────────────────────────────────

class TestGreedySelection:
    def test_returns_one_row_per_strategy(self):
        df = _make_df(4)
        rows = run_greedy_selection(df, sort_metric="rtd")
        assert len(rows) == 4

    def test_step_numbers_sequential(self):
        df = _make_df(3)
        rows = run_greedy_selection(df, sort_metric="rtd")
        assert [r["step"] for r in rows] == [1, 2, 3]

    def test_all_strategies_selected(self):
        df = _make_df(4)
        rows = run_greedy_selection(df, sort_metric="rtd")
        added = [r["strategy_added"] for r in rows]
        assert set(added) == set(df.columns)

    def test_no_duplicates_in_selection(self):
        df = _make_df(5)
        rows = run_greedy_selection(df, sort_metric="rtd")
        added = [r["strategy_added"] for r in rows]
        assert len(added) == len(set(added))

    def test_empty_df_returns_empty(self):
        df = pd.DataFrame()
        rows = run_greedy_selection(df, sort_metric="rtd")
        assert rows == []

    def test_single_strategy(self):
        df = _make_df(1)
        rows = run_greedy_selection(df, sort_metric="rtd")
        assert len(rows) == 1
        assert rows[0]["step"] == 1

    def test_sharpe_metric(self):
        df = _make_df(3)
        rows = run_greedy_selection(df, sort_metric="sharpe")
        assert len(rows) == 3
        for r in rows:
            assert "sharpe" in r

    def test_metrics_keys_present(self):
        df = _make_df(2)
        rows = run_greedy_selection(df, sort_metric="rtd")
        for r in rows:
            for k in ("step", "strategy_added", "annual_profit", "max_dd", "rtd", "rtd_avg", "sharpe"):
                assert k in r


# ── run_randomized_analysis ────────────────────────────────────────────────────

class TestRandomizedAnalysis:
    def test_returns_one_row_per_strategy(self):
        df = _make_df(4)
        result = run_randomized_analysis(df, n_iterations=20, seed=1)
        assert len(result) == 4
        assert set(result.index) == set(df.columns)

    def test_columns_present(self):
        df = _make_df(3)
        result = run_randomized_analysis(df, n_iterations=20, seed=1)
        for col in ("median_rank", "median_contribution", "avg_contribution", "pct_positive"):
            assert col in result.columns

    def test_median_rank_in_range(self):
        n = 5
        df = _make_df(n)
        result = run_randomized_analysis(df, n_iterations=50, seed=2)
        assert (result["median_rank"] >= 1).all()
        assert (result["median_rank"] <= n).all()

    def test_pct_positive_between_0_and_100(self):
        df = _make_df(4)
        result = run_randomized_analysis(df, n_iterations=30, seed=3)
        assert (result["pct_positive"] >= 0).all()
        assert (result["pct_positive"] <= 100).all()

    def test_sorted_descending_by_median_contribution(self):
        df = _make_df(4)
        result = run_randomized_analysis(df, n_iterations=30, seed=4)
        assert list(result["median_contribution"]) == sorted(
            result["median_contribution"], reverse=True
        )

    def test_empty_df_returns_empty(self):
        result = run_randomized_analysis(pd.DataFrame(), n_iterations=10)
        assert result.empty

    def test_single_strategy(self):
        df = _make_df(1)
        result = run_randomized_analysis(df, n_iterations=10, seed=0)
        assert len(result) == 1
        assert result["median_rank"].iloc[0] == 1.0

    def test_deterministic_with_seed(self):
        df = _make_df(4)
        r1 = run_randomized_analysis(df, n_iterations=30, seed=99)
        r2 = run_randomized_analysis(df, n_iterations=30, seed=99)
        pd.testing.assert_frame_equal(r1, r2)
