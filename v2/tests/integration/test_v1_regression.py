"""
v1.24 → v2 regression tests.

These tests load exported CSV data from v1.24 (stored in pt_fixtures/) and
verify that v2's analytics pipeline produces results that match the reference
values from the VBA tool.

Test classes
------------
TestDataIntegrity       — fixture shape, strategy counts, date ranges
TestPortfolioAggregation — total PnL and daily PnL match TotalPortfolioM2M.csv
TestDrawdown            — max drawdown matches v1.24 reference exactly
TestCorrelations        — correlation matrix is valid (shape, range, symmetry)
TestStrategyConfig      — Strategies.csv / Portfolio.csv cross-checks

All tests auto-skip when pt_fixtures/ is absent (see conftest.py).
"""
from __future__ import annotations

import numpy as np
import pandas as pd
import pytest

# ── helpers ───────────────────────────────────────────────────────────────────

TOLERANCE = 0.005   # 0.5% relative tolerance for floating-point comparisons


def _approx(actual: float, expected: float, tol: float = TOLERANCE) -> bool:
    """Return True if actual is within tol of expected (relative)."""
    if abs(expected) < 1e-6:
        return abs(actual) < 1e-6
    return abs(actual - expected) / abs(expected) <= tol


# ── TestDataIntegrity ─────────────────────────────────────────────────────────

class TestDataIntegrity:
    """Verify that fixture files load correctly and have expected shapes."""

    def test_portfolio_m2m_strategy_count(self, v1_portfolio_daily_m2m, v1_strategies):
        """PortfolioDailyM2M.csv has one column per live strategy."""
        live_count = (v1_strategies["Status"] == "Live").sum()
        assert len(v1_portfolio_daily_m2m.columns) == live_count, (
            f"Expected {live_count} live strategies in PortfolioDailyM2M, "
            f"got {len(v1_portfolio_daily_m2m.columns)}"
        )

    def test_portfolio_m2m_strategy_names_match(self, v1_portfolio_daily_m2m, v1_strategies):
        """Every column in PortfolioDailyM2M.csv is a live strategy in Strategies.csv."""
        live_names = set(
            v1_strategies.loc[v1_strategies["Status"] == "Live", "Strategy Name"]
        )
        port_names = set(v1_portfolio_daily_m2m.columns)
        missing = port_names - live_names
        assert not missing, (
            f"{len(missing)} strategy names in PortfolioDailyM2M not found in "
            f"Strategies.csv: {list(missing)[:5]}"
        )

    def test_raw_m2m_contains_all_live_strategies(self, v1_daily_m2m_raw, v1_strategies):
        """DailyM2MEquity.csv (all 430 strategies) includes all 90 live ones."""
        live_names = set(
            v1_strategies.loc[v1_strategies["Status"] == "Live", "Strategy Name"]
        )
        raw_names = set(v1_daily_m2m_raw.columns)
        missing = live_names - raw_names
        assert not missing, (
            f"{len(missing)} live strategies missing from DailyM2MEquity.csv: "
            f"{list(missing)[:5]}"
        )

    def test_raw_m2m_total_strategy_count(self, v1_daily_m2m_raw, v1_strategies):
        """DailyM2MEquity.csv row count in Strategies.csv equals column count."""
        assert len(v1_daily_m2m_raw.columns) == len(v1_strategies), (
            f"DailyM2MEquity has {len(v1_daily_m2m_raw.columns)} cols, "
            f"Strategies.csv has {len(v1_strategies)} rows"
        )

    def test_portfolio_m2m_date_range_covers_total_m2m(
        self, v1_portfolio_daily_m2m, v1_total_m2m
    ):
        """PortfolioDailyM2M spans at least the date range in TotalPortfolioM2M."""
        port_start = v1_portfolio_daily_m2m.index.min()
        port_end = v1_portfolio_daily_m2m.index.max()
        total_start = v1_total_m2m.index.min()
        total_end = v1_total_m2m.index.max()
        assert port_start <= total_start, (
            f"PortfolioDailyM2M starts {port_start} after TotalPortfolioM2M {total_start}"
        )
        assert port_end >= total_end, (
            f"PortfolioDailyM2M ends {port_end} before TotalPortfolioM2M {total_end}"
        )

    def test_total_m2m_no_nan_in_key_columns(self, v1_total_m2m):
        """TotalPortfolioM2M has no NaN in the core financial columns."""
        for col in ["Total Daily Profit", "Total Cumulative P/L", "Total Drawdown"]:
            assert col in v1_total_m2m.columns, f"Column '{col}' missing"
            nan_count = v1_total_m2m[col].isna().sum()
            assert nan_count == 0, f"Column '{col}' has {nan_count} NaN values"

    def test_strategies_have_positive_contracts(self, v1_strategies):
        """All live strategies have a positive contract size (fractional allowed)."""
        live = v1_strategies[v1_strategies["Status"] == "Live"]
        contracts = pd.to_numeric(live["Contracts"], errors="coerce")
        assert (contracts > 0).all(), (
            f"{(contracts <= 0).sum()} live strategies have contracts <= 0"
        )

    def test_latest_positions_match_strategy_names(
        self, v1_latest_positions, v1_strategies
    ):
        """LatestPositionData.csv strategy names overlap with Strategies.csv."""
        all_names = set(v1_strategies["Strategy Name"])
        pos_names = set(v1_latest_positions["Strategy Name"])
        overlap = pos_names & all_names
        assert len(overlap) / len(pos_names) >= 0.95, (
            f"Only {len(overlap)}/{len(pos_names)} position names found in Strategies.csv"
        )


# ── TestPortfolioAggregation ──────────────────────────────────────────────────

class TestPortfolioAggregation:
    """
    Verify that v2's portfolio aggregation reproduces v1.24's total PnL exactly.

    PortfolioDailyM2M.csv is already contract-scaled by v1.24.
    When we feed it into v2's build_portfolio() (contracts=1, all Live),
    portfolio_total_pnl() must produce the same sum as TotalPortfolioM2M.csv.
    """

    def test_total_cumulative_pnl_matches(self, v2_total_pnl, v1_total_m2m):
        """Cumulative portfolio PnL from v2 matches v1.24 reference exactly."""
        common = v2_total_pnl.index.intersection(v1_total_m2m.index)
        assert len(common) > 0, "No common dates between v2 PnL and v1 reference"

        v2_cum = float(v2_total_pnl.loc[common].sum())
        v1_cum = float(v1_total_m2m.loc[common, "Total Cumulative P/L"].iloc[-1])

        assert _approx(v2_cum, v1_cum, tol=0.0001), (
            f"Cumulative PnL mismatch: v2={v2_cum:,.2f}  v1={v1_cum:,.2f}"
        )

    def test_daily_pnl_per_day_matches(self, v2_total_pnl, v1_total_m2m):
        """v2 daily PnL series matches v1.24 Total Daily Profit row-by-row."""
        common = v2_total_pnl.index.intersection(v1_total_m2m.index)
        v2_daily = v2_total_pnl.loc[common]
        v1_daily = v1_total_m2m.loc[common, "Total Daily Profit"]

        max_diff = float((v2_daily - v1_daily).abs().max())
        assert max_diff < 0.02, (
            f"Max daily PnL difference is {max_diff:.4f} (expected 0)"
        )

    def test_common_date_count(self, v2_total_pnl, v1_total_m2m):
        """v2 PnL index overlaps with all dates in TotalPortfolioM2M."""
        common = v2_total_pnl.index.intersection(v1_total_m2m.index)
        assert len(common) == len(v1_total_m2m), (
            f"Only {len(common)} of {len(v1_total_m2m)} TotalPortfolioM2M dates "
            f"present in v2 PnL index"
        )

    def test_active_strategy_count(self, v2_portfolio, v1_strategies):
        """v2 portfolio has exactly the same number of active strategies as v1.24."""
        live_count = int((v1_strategies["Status"] == "Live").sum())
        assert len(v2_portfolio.strategies) == live_count, (
            f"v2 active strategies: {len(v2_portfolio.strategies)}, "
            f"v1 live strategies: {live_count}"
        )


# ── TestDrawdown ──────────────────────────────────────────────────────────────

class TestDrawdown:
    """Verify v2's drawdown computation against v1.24 reference."""

    def test_max_drawdown_matches_v1(self, v2_total_pnl, v1_total_m2m):
        """
        v2 max dollar drawdown matches the maximum value in TotalPortfolioM2M
        'Total Drawdown' column (absolute dollar peak-to-trough).
        """
        common = v2_total_pnl.index.intersection(v1_total_m2m.index)
        equity = v2_total_pnl.loc[common].cumsum()
        peak = equity.cummax()
        v2_max_dd = float((peak - equity).max())

        v1_max_dd = float(v1_total_m2m.loc[common, "Total Drawdown"].max())

        assert _approx(v2_max_dd, v1_max_dd, tol=0.0001), (
            f"Max drawdown mismatch: v2={v2_max_dd:,.2f}  v1={v1_max_dd:,.2f}"
        )

    def test_drawdown_always_non_negative(self, v2_total_pnl, v1_total_m2m):
        """Drawdown series should never be negative (peak - equity >= 0)."""
        common = v2_total_pnl.index.intersection(v1_total_m2m.index)
        equity = v2_total_pnl.loc[common].cumsum()
        peak = equity.cummax()
        drawdown = peak - equity
        neg_count = int((drawdown < -0.01).sum())
        assert neg_count == 0, (
            f"{neg_count} days with negative drawdown (equity above running peak)"
        )

    def test_drawdown_series_matches_v1_per_day(self, v2_total_pnl, v1_total_m2m):
        """Per-day drawdown from v2 matches v1.24 Total Drawdown column."""
        common = v2_total_pnl.index.intersection(v1_total_m2m.index)
        equity = v2_total_pnl.loc[common].cumsum()
        peak = equity.cummax()
        v2_dd = peak - equity
        v1_dd = v1_total_m2m.loc[common, "Total Drawdown"]

        max_diff = float((v2_dd - v1_dd).abs().max())
        assert max_diff < 0.02, (
            f"Max per-day drawdown difference: {max_diff:.4f}"
        )


# ── TestCorrelations ──────────────────────────────────────────────────────────

class TestCorrelations:
    """
    Validate the correlation matrix computed from PortfolioDailyM2M.csv.
    These are structural/sanity tests — not a comparison against a v1.24 output
    (v1.24 doesn't export a correlation matrix CSV).
    """

    def test_matrix_shape_is_square(self, v2_correlation_matrix, v2_portfolio):
        """Correlation matrix is n_strategies × n_strategies."""
        n = len(v2_portfolio.strategies)
        assert v2_correlation_matrix.shape == (n, n), (
            f"Expected ({n}, {n}), got {v2_correlation_matrix.shape}"
        )

    def test_diagonal_is_one(self, v2_correlation_matrix):
        """Each strategy is perfectly correlated with itself."""
        diag = np.diag(v2_correlation_matrix.values)
        assert np.allclose(diag, 1.0, atol=1e-6), (
            f"Diagonal deviates from 1.0: min={diag.min():.6f}, max={diag.max():.6f}"
        )

    def test_values_in_valid_range(self, v2_correlation_matrix):
        """All correlation values are in [-1, 1]."""
        vals = v2_correlation_matrix.values
        assert vals.min() >= -1.0 - 1e-6, f"Min correlation {vals.min():.6f} < -1"
        assert vals.max() <= 1.0 + 1e-6, f"Max correlation {vals.max():.6f} > 1"

    def test_matrix_is_symmetric(self, v2_correlation_matrix):
        """Correlation matrix is symmetric: corr(A,B) == corr(B,A)."""
        m = v2_correlation_matrix.values
        assert np.allclose(m, m.T, atol=1e-10), "Correlation matrix is not symmetric"

    def test_no_nan_values(self, v2_correlation_matrix):
        """No NaN values in the correlation matrix."""
        nan_count = int(np.isnan(v2_correlation_matrix.values).sum())
        assert nan_count == 0, f"{nan_count} NaN values in correlation matrix"


# ── TestStrategyConfig ────────────────────────────────────────────────────────

class TestStrategyConfig:
    """Cross-check Strategies.csv and Portfolio.csv for consistency."""

    def test_portfolio_csv_live_strategy_count(self, v1_portfolio, v1_strategies):
        """Portfolio.csv row count matches live strategy count in Strategies.csv."""
        live_count = int((v1_strategies["Status"] == "Live").sum())
        assert len(v1_portfolio) == live_count, (
            f"Portfolio.csv has {len(v1_portfolio)} rows, "
            f"Strategies.csv has {live_count} live strategies"
        )

    def test_portfolio_names_subset_of_strategies(self, v1_portfolio, v1_strategies):
        """Every strategy in Portfolio.csv is listed in Strategies.csv."""
        all_names = set(v1_strategies["Strategy Name"])
        port_names = set(v1_portfolio.index)
        missing = port_names - all_names
        assert not missing, (
            f"{len(missing)} Portfolio.csv strategies not in Strategies.csv: "
            f"{list(missing)[:5]}"
        )

    def test_unique_strategy_names_in_strategies_csv(self, v1_strategies):
        """Strategy names are unique in Strategies.csv."""
        dupes = v1_strategies["Strategy Name"].duplicated()
        assert not dupes.any(), (
            f"{dupes.sum()} duplicate strategy names: "
            f"{v1_strategies.loc[dupes, 'Strategy Name'].tolist()[:5]}"
        )

    def test_all_positions_are_valid(self, v1_latest_positions):
        """LatestPositionData positions are numeric (0, 1, -1, etc.)."""
        positions = pd.to_numeric(v1_latest_positions["Position"], errors="coerce")
        nan_count = positions.isna().sum()
        assert nan_count == 0, (
            f"{nan_count} non-numeric position values in LatestPositionData.csv"
        )
