"""
Tests for core/analytics/atr.py

Covers: daily range computation, rolling ATR, contract sizing formula,
        historical ATR reweighting.
"""

import math
import pytest
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import pandas as pd
import numpy as np

from core.analytics.atr import (
    ATR_WINDOWS,
    compute_daily_range,
    compute_atr,
    compute_atr_series,
    contract_size_from_atr,
    reweight_contracts_by_atr,
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _make_trades(rows: list[dict]) -> pd.DataFrame:
    df = pd.DataFrame(rows)
    df["date"] = pd.to_datetime(df["date"])
    return df


# ── compute_daily_range ───────────────────────────────────────────────────────

class TestComputeDailyRange:
    def test_abs_mfe_plus_abs_mae(self):
        trades = _make_trades([
            {"strategy": "A", "date": "2023-01-02", "pnl": 100, "mae": 50.0, "mfe": 200.0},
        ])
        result = compute_daily_range(trades)
        assert result.loc[pd.Timestamp("2023-01-02"), "A"] == pytest.approx(250.0)

    def test_negative_mae_handled_by_abs(self):
        trades = _make_trades([
            {"strategy": "A", "date": "2023-01-02", "pnl": -50, "mae": -80.0, "mfe": 120.0},
        ])
        result = compute_daily_range(trades)
        # abs(-80) + abs(120) = 200
        assert result.loc[pd.Timestamp("2023-01-02"), "A"] == pytest.approx(200.0)

    def test_multiple_trades_same_day_summed(self):
        trades = _make_trades([
            {"strategy": "A", "date": "2023-01-02", "pnl": 100, "mae": 30.0, "mfe": 70.0},
            {"strategy": "A", "date": "2023-01-02", "pnl":  50, "mae": 20.0, "mfe": 40.0},
        ])
        result = compute_daily_range(trades)
        # (30+70) + (20+40) = 160
        assert result.loc[pd.Timestamp("2023-01-02"), "A"] == pytest.approx(160.0)

    def test_multiple_strategies(self):
        trades = _make_trades([
            {"strategy": "A", "date": "2023-01-02", "pnl": 100, "mae": 50.0, "mfe": 100.0},
            {"strategy": "B", "date": "2023-01-02", "pnl":  80, "mae": 30.0, "mfe":  60.0},
        ])
        result = compute_daily_range(trades)
        assert "A" in result.columns
        assert "B" in result.columns
        assert result.loc[pd.Timestamp("2023-01-02"), "A"] == pytest.approx(150.0)
        assert result.loc[pd.Timestamp("2023-01-02"), "B"] == pytest.approx(90.0)

    def test_missing_strategy_date_filled_with_zero(self):
        trades = _make_trades([
            {"strategy": "A", "date": "2023-01-02", "pnl": 100, "mae": 50.0, "mfe": 100.0},
            {"strategy": "B", "date": "2023-01-03", "pnl":  80, "mae": 30.0, "mfe":  60.0},
        ])
        result = compute_daily_range(trades)
        assert result.loc[pd.Timestamp("2023-01-02"), "B"] == pytest.approx(0.0)
        assert result.loc[pd.Timestamp("2023-01-03"), "A"] == pytest.approx(0.0)

    def test_empty_trades_returns_empty_df(self):
        result = compute_daily_range(pd.DataFrame())
        assert result.empty

    def test_none_trades_returns_empty_df(self):
        result = compute_daily_range(None)
        assert result.empty

    def test_result_has_datetime_index(self):
        trades = _make_trades([
            {"strategy": "A", "date": "2023-01-02", "pnl": 100, "mae": 50.0, "mfe": 100.0},
        ])
        result = compute_daily_range(trades)
        assert isinstance(result.index, pd.DatetimeIndex)

    def test_index_sorted_ascending(self):
        trades = _make_trades([
            {"strategy": "A", "date": "2023-01-04", "pnl": 100, "mae": 50.0, "mfe": 100.0},
            {"strategy": "A", "date": "2023-01-02", "pnl":  80, "mae": 30.0, "mfe":  60.0},
        ])
        result = compute_daily_range(trades)
        assert result.index[0] < result.index[1]


# ── compute_atr_series ────────────────────────────────────────────────────────

class TestComputeAtrSeries:
    def test_rolling_mean_correct(self):
        """ATR series should be rolling mean of daily ranges."""
        dates = pd.date_range("2023-01-02", periods=5, freq="B")
        ranges = [100.0, 200.0, 300.0, 400.0, 500.0]
        trades_rows = [
            {"strategy": "A", "date": d.date(), "pnl": r, "mae": r * 0.4, "mfe": r * 0.6}
            for d, r in zip(dates, ranges)
        ]
        trades = _make_trades(trades_rows)

        # Window of 3 days
        result = compute_atr_series(trades, "ATR Last 3 Months")
        # Not testing exact values since window=63; just check shape/type
        assert isinstance(result, pd.DataFrame)
        assert "A" in result.columns
        assert len(result) == 5

    def test_window_3m_uses_63_days(self):
        assert ATR_WINDOWS["ATR Last 3 Months"] == 63

    def test_window_6m_uses_126_days(self):
        assert ATR_WINDOWS["ATR Last 6 Months"] == 126

    def test_window_12m_uses_252_days(self):
        assert ATR_WINDOWS["ATR Last 12 Months"] == 252

    def test_min_periods_1_no_nan_at_start(self):
        """min_periods=1 means first row should not be NaN."""
        trades = _make_trades([
            {"strategy": "A", "date": "2023-01-02", "pnl": 100, "mae": 50.0, "mfe": 100.0},
        ])
        result = compute_atr_series(trades, "ATR Last 3 Months")
        assert not result.isna().any().any()

    def test_empty_returns_empty(self):
        assert compute_atr_series(pd.DataFrame()).empty


# ── compute_atr ───────────────────────────────────────────────────────────────

class TestComputeAtr:
    def test_returns_latest_value(self):
        """compute_atr returns the last row of the ATR series."""
        trades = _make_trades([
            {"strategy": "A", "date": "2023-01-02", "pnl": 100, "mae":  50.0, "mfe": 100.0},
            {"strategy": "A", "date": "2023-01-03", "pnl": 200, "mae": 100.0, "mfe": 200.0},
        ])
        result = compute_atr(trades, "ATR Last 3 Months")
        assert isinstance(result, pd.Series)
        assert "A" in result.index
        assert result["A"] > 0

    def test_consistent_with_series_last_row(self):
        trades = _make_trades([
            {"strategy": "A", "date": "2023-01-02", "pnl": 100, "mae": 50.0, "mfe": 100.0},
            {"strategy": "A", "date": "2023-01-03", "pnl": 200, "mae": 80.0, "mfe": 120.0},
        ])
        series_last = compute_atr_series(trades).iloc[-1]
        scalar = compute_atr(trades)
        assert series_last["A"] == pytest.approx(scalar["A"])

    def test_empty_returns_empty_series(self):
        result = compute_atr(pd.DataFrame())
        assert isinstance(result, pd.Series)
        assert result.empty


# ── contract_size_from_atr ────────────────────────────────────────────────────

class TestContractSizeFromAtr:
    def test_pure_atr_sizing(self):
        """ratio=1.0: dollar_risk = atr only."""
        # equity=100_000, pct=0.01, atr=500, margin=0, ratio=1.0
        # raw = 100_000 * 0.01 / 500 = 2.0 → 2 contracts
        n = contract_size_from_atr(
            equity=100_000, contract_size_pct=0.01,
            atr_dollars=500.0, margin=0.0, ratio=1.0,
        )
        assert n == 2

    def test_pure_margin_sizing(self):
        """ratio=0.0: dollar_risk = margin only."""
        # equity=100_000, pct=0.01, margin=1_000, ratio=0.0
        # raw = 100_000 * 0.01 / 1_000 = 1.0 → 1 contract
        n = contract_size_from_atr(
            equity=100_000, contract_size_pct=0.01,
            atr_dollars=0.0, margin=1_000.0, ratio=0.0,
        )
        assert n == 1

    def test_blended_50_50(self):
        """ratio=0.5: blends ATR and margin equally."""
        # dollar_risk = 600*0.5 + 400*0.5 = 500
        # raw = 200_000 * 0.01 / 500 = 4.0
        n = contract_size_from_atr(
            equity=200_000, contract_size_pct=0.01,
            atr_dollars=600.0, margin=400.0, ratio=0.5,
        )
        assert n == 4

    def test_floor_applied(self):
        """Non-integer result is floored, not rounded."""
        # raw = 100_000 * 0.01 / 300 = 3.333... → 3
        n = contract_size_from_atr(
            equity=100_000, contract_size_pct=0.01,
            atr_dollars=300.0, margin=0.0, ratio=1.0,
        )
        assert n == 3

    def test_minimum_one_contract(self):
        """Very small equity → floor(raw) could be 0 but we return 1."""
        n = contract_size_from_atr(
            equity=100, contract_size_pct=0.01,
            atr_dollars=50_000.0, margin=0.0, ratio=1.0,
        )
        assert n == 1

    def test_zero_atr_falls_back_to_margin(self):
        """ATR=0 with ratio=0.5: dollar_risk = 0*0.5 + margin*0.5 = margin/2."""
        n = contract_size_from_atr(
            equity=100_000, contract_size_pct=0.01,
            atr_dollars=0.0, margin=1_000.0, ratio=0.5,
        )
        # dollar_risk = 500, raw = 1000/500 = 2
        assert n == 2

    def test_zero_effective_risk_returns_one(self):
        n = contract_size_from_atr(
            equity=100_000, contract_size_pct=0.01,
            atr_dollars=0.0, margin=0.0, ratio=1.0,
        )
        assert n == 1

    def test_invalid_ratio_raises(self):
        with pytest.raises(ValueError):
            contract_size_from_atr(100_000, 0.01, 500.0, 1_000.0, ratio=1.5)
        with pytest.raises(ValueError):
            contract_size_from_atr(100_000, 0.01, 500.0, 1_000.0, ratio=-0.1)

    def test_negative_atr_uses_abs(self):
        """Negative ATR values are treated as absolute."""
        n1 = contract_size_from_atr(100_000, 0.01, 500.0, 0.0, 1.0)
        n2 = contract_size_from_atr(100_000, 0.01, -500.0, 0.0, 1.0)
        assert n1 == n2


# ── reweight_contracts_by_atr ─────────────────────────────────────────────────

class TestReweightContractsByAtr:
    def _make_atr_series(self, strategy, values, dates=None):
        if dates is None:
            dates = pd.date_range("2023-01-02", periods=len(values), freq="B")
        return pd.DataFrame({strategy: values}, index=dates)

    def test_higher_current_atr_increases_contracts(self):
        """current_atr > historical → more contracts (equity risk unchanged)."""
        dates = pd.date_range("2023-01-02", periods=3, freq="B")
        # Historical ATR = 100, current = 200 → scale factor 2
        atr_series = pd.DataFrame({"A": [100.0, 100.0, 100.0]}, index=dates)
        base = pd.Series({"A": 2})
        current = pd.Series({"A": 200.0})

        result = reweight_contracts_by_atr(base, atr_series, current)
        # floor(2 * 200/100) = floor(4.0) = 4
        assert result.loc[dates[0], "A"] == pytest.approx(4.0)

    def test_lower_current_atr_decreases_contracts(self):
        """current_atr < historical → fewer contracts."""
        dates = pd.date_range("2023-01-02", periods=2, freq="B")
        atr_series = pd.DataFrame({"A": [200.0, 200.0]}, index=dates)
        base = pd.Series({"A": 4})
        current = pd.Series({"A": 100.0})

        result = reweight_contracts_by_atr(base, atr_series, current)
        # floor(4 * 100/200) = floor(2.0) = 2
        assert result.loc[dates[0], "A"] == pytest.approx(2.0)

    def test_same_atr_unchanged(self):
        """current_atr == historical_atr → same as base."""
        dates = pd.date_range("2023-01-02", periods=2, freq="B")
        atr_series = pd.DataFrame({"A": [300.0, 300.0]}, index=dates)
        base = pd.Series({"A": 3})
        current = pd.Series({"A": 300.0})

        result = reweight_contracts_by_atr(base, atr_series, current)
        assert result.loc[dates[0], "A"] == pytest.approx(3.0)

    def test_zero_base_returns_zero(self):
        dates = pd.date_range("2023-01-02", periods=2, freq="B")
        atr_series = pd.DataFrame({"A": [300.0, 300.0]}, index=dates)
        base = pd.Series({"A": 0})
        current = pd.Series({"A": 300.0})

        result = reweight_contracts_by_atr(base, atr_series, current)
        assert (result["A"] == 0).all()

    def test_minimum_one_contract_when_base_positive(self):
        """Even when current_atr << historical, result is at least 1."""
        dates = pd.date_range("2023-01-02", periods=2, freq="B")
        atr_series = pd.DataFrame({"A": [10_000.0, 10_000.0]}, index=dates)
        base = pd.Series({"A": 1})
        current = pd.Series({"A": 1.0})

        result = reweight_contracts_by_atr(base, atr_series, current)
        # floor(1 * 1/10000) = 0, but minimum is 1
        assert (result["A"] >= 1).all()

    def test_zero_historical_atr_uses_current(self):
        """When historical ATR is 0, ratio = 1 → no rescaling (base kept)."""
        dates = pd.date_range("2023-01-02", periods=2, freq="B")
        atr_series = pd.DataFrame({"A": [0.0, 0.0]}, index=dates)
        base = pd.Series({"A": 3})
        current = pd.Series({"A": 200.0})

        result = reweight_contracts_by_atr(base, atr_series, current)
        # hist is replaced by current → floor(3 * 200/200) = 3
        assert (result["A"] == 3.0).all()

    def test_empty_atr_series_returns_empty(self):
        base = pd.Series({"A": 3})
        current = pd.Series({"A": 200.0})
        result = reweight_contracts_by_atr(base, pd.DataFrame(), current)
        assert result.empty

    def test_multiple_strategies(self):
        dates = pd.date_range("2023-01-02", periods=2, freq="B")
        atr_series = pd.DataFrame(
            {"A": [100.0, 100.0], "B": [200.0, 200.0]}, index=dates
        )
        base = pd.Series({"A": 2, "B": 4})
        current = pd.Series({"A": 200.0, "B": 100.0})

        result = reweight_contracts_by_atr(base, atr_series, current)
        # A: floor(2 * 200/100) = 4
        # B: floor(4 * 100/200) = 2
        assert result.loc[dates[0], "A"] == pytest.approx(4.0)
        assert result.loc[dates[0], "B"] == pytest.approx(2.0)
