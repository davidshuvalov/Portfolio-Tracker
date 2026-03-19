"""
Unit tests for core/analytics/leave_one_out.py

Tests:
  - _remove_strategy: correct exclusion, shape preservation
  - _analyse_portfolio: returns MCResult for non-empty portfolio
  - run_leave_one_out: delta signs, row count, column names, edge cases
"""

from __future__ import annotations

from pathlib import Path

import numpy as np
import pandas as pd
import pytest

from core.analytics.leave_one_out import (
    _analyse_portfolio,
    _remove_strategy,
    run_leave_one_out,
)
from core.config import MCConfig
from core.data_types import MCResult, PortfolioData, Strategy


# ── Fixtures ──────────────────────────────────────────────────────────────────

def make_strategy(name: str, contracts: int = 1) -> Strategy:
    return Strategy(
        name=name,
        folder=Path("."),
        status="Live",
        contracts=contracts,
    )


def make_portfolio(
    n_strats: int = 3,
    n_days: int = 300,
    daily_profit: float = 100.0,
    seed: int = 42,
) -> PortfolioData:
    rng = np.random.default_rng(seed)
    idx = pd.bdate_range("2021-01-01", periods=n_days)
    strat_names = [f"S{i+1}" for i in range(n_strats)]
    strategies = [make_strategy(name) for name in strat_names]

    daily_pnl = pd.DataFrame(
        rng.normal(daily_profit, 500, (n_days, n_strats)),
        index=idx,
        columns=strat_names,
    )
    closed_trades = pd.DataFrame({
        "strategy": strat_names * 10,
        "date": [idx[i % n_days] for i in range(n_strats * 10)],
        "pnl": rng.normal(daily_profit * 5, 1000, n_strats * 10),
        "position": ["Long"] * (n_strats * 10),
        "mae": np.zeros(n_strats * 10),
        "mfe": np.zeros(n_strats * 10),
    })
    return PortfolioData(
        strategies=strategies,
        daily_pnl=daily_pnl,
        closed_trades=closed_trades,
        summary_metrics=pd.DataFrame(),
    )


def make_mc_config(simulations: int = 100) -> MCConfig:
    return MCConfig(
        simulations=simulations,
        period="IS+OOS",
        risk_ruin_target=0.10,
        risk_ruin_tolerance=0.03,
        trade_adjustment=0.0,
        trade_option="M2M",
    )


# ── _remove_strategy ─────────────────────────────────────────────────────────

class TestRemoveStrategy:
    def test_removes_correct_strategy(self):
        p = make_portfolio(n_strats=3)
        reduced = _remove_strategy(p, "S2")
        strat_names = [s.name for s in reduced.strategies]
        assert "S2" not in strat_names
        assert "S1" in strat_names
        assert "S3" in strat_names

    def test_daily_pnl_columns_reduced(self):
        p = make_portfolio(n_strats=4)
        reduced = _remove_strategy(p, "S3")
        assert "S3" not in reduced.daily_pnl.columns
        assert len(reduced.daily_pnl.columns) == 3

    def test_strategy_count_decreases_by_one(self):
        p = make_portfolio(n_strats=5)
        reduced = _remove_strategy(p, "S2")
        assert len(reduced.strategies) == 4

    def test_date_index_preserved(self):
        p = make_portfolio(n_strats=3)
        reduced = _remove_strategy(p, "S1")
        assert len(reduced.daily_pnl) == len(p.daily_pnl)
        pd.testing.assert_index_equal(reduced.daily_pnl.index, p.daily_pnl.index)

    def test_closed_trades_filtered(self):
        p = make_portfolio(n_strats=3)
        reduced = _remove_strategy(p, "S2")
        assert "S2" not in reduced.closed_trades["strategy"].values

    def test_removing_nonexistent_strategy_is_noop(self):
        p = make_portfolio(n_strats=3)
        reduced = _remove_strategy(p, "MISSING")
        assert len(reduced.strategies) == 3
        assert len(reduced.daily_pnl.columns) == 3

    def test_removing_last_strategy_gives_empty(self):
        p = make_portfolio(n_strats=1)
        reduced = _remove_strategy(p, "S1")
        assert len(reduced.strategies) == 0
        assert reduced.daily_pnl.empty


# ── _analyse_portfolio ────────────────────────────────────────────────────────

class TestAnalysePortfolio:
    def test_returns_mcresult(self):
        p = make_portfolio()
        config = make_mc_config()
        result = _analyse_portfolio(p, config, margin_threshold=5_000.0)
        assert isinstance(result, MCResult)

    def test_empty_portfolio_returns_zero_profit(self):
        empty = PortfolioData(
            strategies=[],
            daily_pnl=pd.DataFrame(),
            closed_trades=pd.DataFrame(),
            summary_metrics=pd.DataFrame(),
        )
        config = make_mc_config()
        result = _analyse_portfolio(empty, config, 5_000.0)
        assert result.expected_profit == 0.0
        assert result.starting_equity == 0.0

    def test_positive_drift_gives_positive_expected_profit(self):
        """A strongly positive PnL series should yield positive expected profit."""
        rng = np.random.default_rng(1)
        idx = pd.bdate_range("2022-01-01", periods=252)
        strats = [make_strategy("A"), make_strategy("B")]
        daily_pnl = pd.DataFrame(
            rng.normal(200, 100, (252, 2)),  # strong positive drift
            index=idx,
            columns=["A", "B"],
        )
        p = PortfolioData(
            strategies=strats,
            daily_pnl=daily_pnl,
            closed_trades=pd.DataFrame(),
            summary_metrics=pd.DataFrame(),
        )
        config = make_mc_config(simulations=200)
        result = _analyse_portfolio(p, config, 5_000.0)
        assert result.expected_profit > 0


# ── run_leave_one_out ─────────────────────────────────────────────────────────

class TestRunLeaveOneOut:
    def test_returns_correct_number_of_rows(self):
        p = make_portfolio(n_strats=4)
        config = make_mc_config()
        result = run_leave_one_out(p, config, 5_000.0)
        assert len(result) == 4

    def test_all_strategy_names_present(self):
        p = make_portfolio(n_strats=3)
        config = make_mc_config()
        result = run_leave_one_out(p, config, 5_000.0)
        assert set(result["strategy"]) == {"S1", "S2", "S3"}

    def test_required_columns_present(self):
        p = make_portfolio(n_strats=3)
        config = make_mc_config()
        result = run_leave_one_out(p, config, 5_000.0)
        required = {"strategy", "delta_profit", "delta_sharpe", "delta_drawdown", "delta_rtd", "delta_ror"}
        assert required.issubset(set(result.columns))

    def test_empty_portfolio_returns_empty_df(self):
        empty = PortfolioData(
            strategies=[],
            daily_pnl=pd.DataFrame(),
            closed_trades=pd.DataFrame(),
            summary_metrics=pd.DataFrame(),
        )
        config = make_mc_config()
        result = run_leave_one_out(empty, config, 5_000.0)
        assert result.empty

    def test_single_strategy_gives_one_row(self):
        p = make_portfolio(n_strats=1)
        config = make_mc_config()
        result = run_leave_one_out(p, config, 1_000.0)
        assert len(result) == 1

    def test_sorted_by_delta_profit_ascending(self):
        """Result should be sorted delta_profit ascending (most valuable first)."""
        p = make_portfolio(n_strats=4)
        config = make_mc_config()
        result = run_leave_one_out(p, config, 5_000.0)
        deltas = result["delta_profit"].values
        assert all(deltas[i] <= deltas[i + 1] for i in range(len(deltas) - 1))

    def test_delta_types_are_float(self):
        p = make_portfolio(n_strats=3)
        config = make_mc_config()
        result = run_leave_one_out(p, config, 5_000.0)
        for col in ["delta_profit", "delta_sharpe", "delta_drawdown", "delta_rtd"]:
            assert result[col].dtype in (np.float64, float, np.float32)
