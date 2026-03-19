"""
Unit tests for core/analytics/margin.py

Tests:
  - detect_positions: Long/Short/Flat detection, as_of_date filtering, alignment
  - get_strategy_position_table: strategy metadata merge
  - net_position_by_symbol: long/short/flat aggregation
  - compute_daily_margin: weight calculation, empty data guards
  - margin_by_symbol: symbol grouping
  - margin_by_sector: sector aggregation
  - margin_summary_stats: stat values
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import pytest

from core.analytics.margin import (
    PositionStatus,
    _in_market_mask,
    compute_daily_margin,
    detect_positions,
    get_strategy_position_table,
    margin_by_sector,
    margin_by_symbol,
    margin_summary_stats,
    net_position_by_symbol,
)
from core.data_types import Strategy


# ── Fixtures ──────────────────────────────────────────────────────────────────

def make_strategy(
    name: str,
    symbol: str = "",
    sector: str = "",
    contracts: int = 1,
    status: str = "Live",
) -> Strategy:
    return Strategy(
        name=name,
        folder=Path("."),
        status=status,
        contracts=contracts,
        symbol=symbol,
        sector=sector,
    )


def make_in_market_frames(
    n_days: int = 10,
    long_strats: dict[str, list[float]] | None = None,   # name → daily values
    short_strats: dict[str, list[float]] | None = None,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Build in_market_long and in_market_short DataFrames."""
    idx = pd.bdate_range("2023-01-01", periods=n_days)
    long_data  = long_strats  or {}
    short_data = short_strats or {}

    all_strats = list(set(long_data) | set(short_data))
    long_df  = pd.DataFrame(
        {s: long_data.get(s, [0.0] * n_days) for s in all_strats}, index=idx
    )
    short_df = pd.DataFrame(
        {s: short_data.get(s, [0.0] * n_days) for s in all_strats}, index=idx
    )
    return long_df, short_df


# ── detect_positions ──────────────────────────────────────────────────────────

class TestDetectPositions:
    def test_long_detected(self):
        idx = pd.bdate_range("2023-01-01", periods=3)
        long  = pd.DataFrame({"A": [0.0, 100.0, 150.0]}, index=idx)
        short = pd.DataFrame({"A": [0.0, 0.0, 0.0]}, index=idx)
        result = detect_positions(long, short)
        assert result.loc[result["strategy"] == "A", "status"].iloc[0] == PositionStatus.LONG.value

    def test_short_detected(self):
        idx = pd.bdate_range("2023-01-01", periods=3)
        long  = pd.DataFrame({"A": [0.0, 0.0, 0.0]}, index=idx)
        short = pd.DataFrame({"A": [0.0, -100.0, -150.0]}, index=idx)
        result = detect_positions(long, short)
        assert result.loc[result["strategy"] == "A", "status"].iloc[0] == PositionStatus.SHORT.value

    def test_flat_detected(self):
        idx = pd.bdate_range("2023-01-01", periods=3)
        long  = pd.DataFrame({"A": [0.0, 0.0, 0.0]}, index=idx)
        short = pd.DataFrame({"A": [0.0, 0.0, 0.0]}, index=idx)
        result = detect_positions(long, short)
        assert result.loc[result["strategy"] == "A", "status"].iloc[0] == PositionStatus.FLAT.value

    def test_multiple_strategies(self):
        idx = pd.bdate_range("2023-01-01", periods=3)
        long  = pd.DataFrame({"A": [100.0, 100.0, 100.0], "B": [0.0, 0.0, 0.0]}, index=idx)
        short = pd.DataFrame({"A": [0.0, 0.0, 0.0],       "B": [-50.0, -50.0, -50.0]}, index=idx)
        result = detect_positions(long, short)
        pos = dict(zip(result["strategy"], result["status"]))
        assert pos["A"] == PositionStatus.LONG.value
        assert pos["B"] == PositionStatus.SHORT.value

    def test_as_of_date_filtering(self):
        """as_of_date should use most recent row on or before that date."""
        idx = pd.bdate_range("2023-01-01", periods=5)
        # Strategy is long on day 3, flat after
        long_vals  = [100.0, 100.0, 100.0, 0.0, 0.0]
        short_vals = [0.0] * 5
        long  = pd.DataFrame({"A": long_vals}, index=idx)
        short = pd.DataFrame({"A": short_vals}, index=idx)

        # Query as of day 3 → long
        result_early = detect_positions(long, short, as_of_date=idx[2])
        assert result_early.loc[0, "status"] == PositionStatus.LONG.value

        # Query as of day 5 → flat
        result_late = detect_positions(long, short, as_of_date=idx[4])
        assert result_late.loc[0, "status"] == PositionStatus.FLAT.value

    def test_empty_dataframes_returns_empty(self):
        result = detect_positions(pd.DataFrame(), pd.DataFrame())
        assert result.empty

    def test_long_takes_priority_over_short_in_same_day(self):
        """If both long and short are non-zero, long wins."""
        idx = pd.bdate_range("2023-01-01", periods=1)
        long  = pd.DataFrame({"A": [100.0]}, index=idx)
        short = pd.DataFrame({"A": [-50.0]}, index=idx)
        result = detect_positions(long, short)
        assert result.loc[0, "is_long"] == True


# ── get_strategy_position_table ───────────────────────────────────────────────

class TestGetStrategyPositionTable:
    def test_symbol_and_sector_merged(self):
        idx = pd.bdate_range("2023-01-01", periods=2)
        long  = pd.DataFrame({"ES": [100.0, 100.0]}, index=idx)
        short = pd.DataFrame({"ES": [0.0, 0.0]}, index=idx)
        strats = [make_strategy("ES", symbol="ES", sector="Equity", contracts=2)]
        result = get_strategy_position_table(long, short, strats)
        row = result.loc[result["strategy"] == "ES"].iloc[0]
        assert row["symbol"] == "ES"
        assert row["sector"] == "Equity"
        assert row["contracts"] == 2
        assert row["position_status"] == PositionStatus.LONG.value

    def test_unknown_strategy_gets_empty_metadata(self):
        idx = pd.bdate_range("2023-01-01", periods=2)
        long  = pd.DataFrame({"UNKNOWN": [100.0, 100.0]}, index=idx)
        short = pd.DataFrame({"UNKNOWN": [0.0, 0.0]}, index=idx)
        result = get_strategy_position_table(long, short, [])
        row = result.loc[result["strategy"] == "UNKNOWN"].iloc[0]
        assert row["symbol"] == ""
        assert row["sector"] == ""


# ── net_position_by_symbol ────────────────────────────────────────────────────

class TestNetPositionBySymbol:
    def _make_table(self, rows):
        return pd.DataFrame(rows)

    def test_net_long(self):
        table = self._make_table([
            {"strategy": "S1", "symbol": "ES", "contracts": 1, "position_status": "Long"},
            {"strategy": "S2", "symbol": "ES", "contracts": 2, "position_status": "Long"},
        ])
        net = net_position_by_symbol(table)
        row = net.loc[net["symbol"] == "ES"].iloc[0]
        assert row["net"] == 3
        assert row["net_status"] == PositionStatus.LONG.value

    def test_net_short(self):
        table = self._make_table([
            {"strategy": "S1", "symbol": "NQ", "contracts": 1, "position_status": "Short"},
        ])
        net = net_position_by_symbol(table)
        row = net.loc[net["symbol"] == "NQ"].iloc[0]
        assert row["net"] == -1
        assert row["net_status"] == PositionStatus.SHORT.value

    def test_mixed_nets_to_flat(self):
        table = self._make_table([
            {"strategy": "S1", "symbol": "CL", "contracts": 2, "position_status": "Long"},
            {"strategy": "S2", "symbol": "CL", "contracts": 2, "position_status": "Short"},
        ])
        net = net_position_by_symbol(table)
        row = net.loc[net["symbol"] == "CL"].iloc[0]
        assert row["net"] == 0
        assert row["net_status"] == PositionStatus.FLAT.value

    def test_empty_table(self):
        assert net_position_by_symbol(pd.DataFrame()).empty

    def test_empty_symbol_excluded(self):
        table = self._make_table([
            {"strategy": "S1", "symbol": "", "contracts": 1, "position_status": "Long"},
        ])
        net = net_position_by_symbol(table)
        assert net.empty


# ── compute_daily_margin ──────────────────────────────────────────────────────

class TestComputeDailyMargin:
    def test_basic_margin_calculation(self):
        """1 strategy in-market all 5 days, 1 contract, $10k margin → $10k/day."""
        n = 5
        idx = pd.bdate_range("2023-01-01", periods=n)
        long  = pd.DataFrame({"ES": [100.0] * n}, index=idx)
        short = pd.DataFrame({"ES": [0.0]   * n}, index=idx)
        strats = [make_strategy("ES", symbol="ES", contracts=1)]
        margin = compute_daily_margin(long, short, strats, {"ES": 10_000.0})
        assert float(margin.mean()) == pytest.approx(10_000.0)

    def test_contracts_multiplied(self):
        """2 contracts × $5k = $10k per day."""
        n = 3
        idx = pd.bdate_range("2023-01-01", periods=n)
        long  = pd.DataFrame({"NQ": [50.0] * n}, index=idx)
        short = pd.DataFrame({"NQ": [0.0]  * n}, index=idx)
        strats = [make_strategy("NQ", symbol="NQ", contracts=2)]
        margin = compute_daily_margin(long, short, strats, {"NQ": 5_000.0})
        assert float(margin.iloc[0]) == pytest.approx(10_000.0)

    def test_flat_days_contribute_zero(self):
        n = 4
        idx = pd.bdate_range("2023-01-01", periods=n)
        # In-market only on day 0 and 2
        long  = pd.DataFrame({"A": [100.0, 0.0, 100.0, 0.0]}, index=idx)
        short = pd.DataFrame({"A": [0.0]   * n},               index=idx)
        strats = [make_strategy("A", symbol="A", contracts=1)]
        margin = compute_daily_margin(long, short, strats, {"A": 5_000.0})
        assert float(margin.iloc[1]) == 0.0
        assert float(margin.iloc[2]) == pytest.approx(5_000.0)

    def test_default_margin_used_for_unknown_symbol(self):
        n = 2
        idx = pd.bdate_range("2023-01-01", periods=n)
        long  = pd.DataFrame({"X": [100.0] * n}, index=idx)
        short = pd.DataFrame({"X": [0.0]   * n}, index=idx)
        strats = [make_strategy("X", symbol="UNKNOWN_SYM", contracts=1)]
        margin = compute_daily_margin(long, short, strats, {}, default_margin=3_000.0)
        assert float(margin.iloc[0]) == pytest.approx(3_000.0)

    def test_empty_in_market_returns_empty_series(self):
        strats = [make_strategy("A")]
        margin = compute_daily_margin(pd.DataFrame(), pd.DataFrame(), strats, {})
        assert margin.empty

    def test_multiple_strategies_summed(self):
        n = 2
        idx = pd.bdate_range("2023-01-01", periods=n)
        long  = pd.DataFrame({"A": [100.0]*n, "B": [50.0]*n}, index=idx)
        short = pd.DataFrame({"A": [0.0]*n,   "B": [0.0]*n}, index=idx)
        strats = [
            make_strategy("A", symbol="ES", contracts=1),
            make_strategy("B", symbol="NQ", contracts=1),
        ]
        margin = compute_daily_margin(long, short, strats, {"ES": 10_000.0, "NQ": 20_000.0})
        assert float(margin.iloc[0]) == pytest.approx(30_000.0)


# ── margin_by_symbol ─────────────────────────────────────────────────────────

class TestMarginBySymbol:
    def test_groups_by_symbol(self):
        """Two strategies on same symbol → their margins sum under that symbol."""
        n = 3
        idx = pd.bdate_range("2023-01-01", periods=n)
        long  = pd.DataFrame({"S1": [100.0]*n, "S2": [50.0]*n}, index=idx)
        short = pd.DataFrame({"S1": [0.0]*n,   "S2": [0.0]*n}, index=idx)
        strats = [
            make_strategy("S1", symbol="ES", contracts=1),
            make_strategy("S2", symbol="ES", contracts=1),
        ]
        result = margin_by_symbol(long, short, strats, {"ES": 5_000.0})
        assert "ES" in result.columns
        assert float(result["ES"].iloc[0]) == pytest.approx(10_000.0)

    def test_different_symbols_separate_columns(self):
        n = 2
        idx = pd.bdate_range("2023-01-01", periods=n)
        long  = pd.DataFrame({"S1": [100.0]*n, "S2": [50.0]*n}, index=idx)
        short = pd.DataFrame({"S1": [0.0]*n,   "S2": [0.0]*n}, index=idx)
        strats = [
            make_strategy("S1", symbol="ES"),
            make_strategy("S2", symbol="NQ"),
        ]
        result = margin_by_symbol(long, short, strats, {"ES": 1_000.0, "NQ": 2_000.0})
        assert "ES" in result.columns
        assert "NQ" in result.columns

    def test_empty_returns_empty(self):
        result = margin_by_symbol(pd.DataFrame(), pd.DataFrame(), [], {})
        assert result.empty


# ── margin_by_sector ─────────────────────────────────────────────────────────

class TestMarginBySector:
    def test_aggregates_by_sector(self):
        n = 2
        idx = pd.bdate_range("2023-01-01", periods=n)
        sym_df = pd.DataFrame({"ES": [10_000.0]*n, "NQ": [20_000.0]*n}, index=idx)
        strats = [
            make_strategy("s1", symbol="ES", sector="Equity"),
            make_strategy("s2", symbol="NQ", sector="Equity"),
        ]
        result = margin_by_sector(sym_df, strats, {})
        assert "Equity" in result.columns
        assert float(result["Equity"].iloc[0]) == pytest.approx(30_000.0)

    def test_empty_returns_empty(self):
        result = margin_by_sector(pd.DataFrame(), [], {})
        assert result.empty


# ── margin_summary_stats ──────────────────────────────────────────────────────

class TestMarginSummaryStats:
    def test_basic_stats(self):
        idx = pd.bdate_range("2023-01-01", periods=5)
        series = pd.Series([1000.0, 2000.0, 5000.0, 3000.0, 5000.0], index=idx, name="total_margin")
        sym = pd.DataFrame({"ES": series})
        stats = margin_summary_stats(series, sym)

        assert stats["current_margin"]  == pytest.approx(5000.0)
        assert stats["peak_margin"]     == pytest.approx(5000.0)
        assert stats["average_margin"]  == pytest.approx(3200.0)
        assert stats["days_at_peak"]    == 2  # days 3 and 5 both at 5000
        assert stats["top_symbol"]      == "ES"

    def test_empty_series_returns_empty(self):
        stats = margin_summary_stats(pd.Series(dtype=float), pd.DataFrame())
        assert stats == {}
