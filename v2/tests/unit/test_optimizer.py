"""
Tests for core/portfolio/optimizer.py

Covers each workflow step function and the run_workflow orchestrator.
"""

from __future__ import annotations

import sys
from pathlib import Path
import math

import pytest
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from core.portfolio.optimizer import (
    OptimizerState,
    ExclusionRecord,
    build_candidates,
    run_workflow,
    step_filter_eligibility,
    step_filter_excluded_symbols,
    step_filter_contract_size,
    step_rank,
    step_size_contracts,
    step_select_strategies,
    step_adjust_correlations,
    step_adjust_gross_margins,
    step_adjust_drawdowns,
    portfolio_summary,
    _round_to_fraction,
)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _state(strategies: list[dict], equity: float = 500_000.0) -> OptimizerState:
    """Build a minimal OptimizerState from strategy dicts."""
    return OptimizerState(
        candidates=list(strategies),
        contracts={s["name"]: float(s.get("contracts", 1)) for s in strategies},
        equity=equity,
    )


def _strat(name: str, symbol: str = "ES", sector: str = "Index",
           contracts: float = 1.0, **kwargs) -> dict:
    return {"name": name, "symbol": symbol, "sector": sector,
            "contracts": contracts, **kwargs}


_MARGINS = {"ES": 10_000.0, "NQ": 15_000.0, "CL": 5_000.0, "GC": 8_000.0}
_MARGIN_MULT = 1.0


# ── _round_to_fraction ─────────────────────────────────────────────────────────

class TestRoundToFraction:
    def test_rounds_down(self):
        assert _round_to_fraction(3.75, 0.1) == pytest.approx(3.7)

    def test_exact_multiple_unchanged(self):
        assert _round_to_fraction(3.5, 0.5) == pytest.approx(3.5)

    def test_zero_fraction_returns_float(self):
        assert _round_to_fraction(3.75, 0.0) == pytest.approx(3.75)

    def test_one_fraction(self):
        assert _round_to_fraction(4.9, 1.0) == pytest.approx(4.0)


# ── step_filter_eligibility ────────────────────────────────────────────────────

class TestStepFilterEligibility:
    def test_removes_ineligible(self):
        state = _state([_strat("A"), _strat("B")])
        state = step_filter_eligibility(state, {"A": True, "B": False})
        assert {c["name"] for c in state.candidates} == {"A"}
        assert any(e.name == "B" for e in state.excluded)

    def test_keeps_all_when_all_eligible(self):
        strats = [_strat("A"), _strat("B")]
        state = _state(strats)
        state = step_filter_eligibility(state, {"A": True, "B": True})
        assert len(state.candidates) == 2
        assert len(state.excluded) == 0

    def test_missing_from_mask_kept(self):
        """Strategies absent from mask are treated as eligible (True default)."""
        state = _state([_strat("A"), _strat("B")])
        state = step_filter_eligibility(state, {"B": False})
        assert {c["name"] for c in state.candidates} == {"A"}

    def test_empty_mask_keeps_all(self):
        state = _state([_strat("A"), _strat("B")])
        state = step_filter_eligibility(state, {})
        assert len(state.candidates) == 2

    def test_log_entry_added(self):
        state = _state([_strat("A")])
        state = step_filter_eligibility(state, {"A": False})
        assert any("eligibility" in line for line in state.log)


# ── step_filter_excluded_symbols ───────────────────────────────────────────────

class TestStepFilterExcludedSymbols:
    def test_removes_matching_symbol(self):
        state = _state([_strat("A", symbol="ES"), _strat("B", symbol="S")])
        state = step_filter_excluded_symbols(state, ["S"])
        assert {c["name"] for c in state.candidates} == {"A"}

    def test_case_insensitive(self):
        state = _state([_strat("A", symbol="ym")])
        state = step_filter_excluded_symbols(state, ["YM"])
        assert len(state.candidates) == 0

    def test_no_exclusions_unchanged(self):
        state = _state([_strat("A"), _strat("B")])
        before = len(state.candidates)
        state = step_filter_excluded_symbols(state, [])
        assert len(state.candidates) == before

    def test_multiple_exclusions(self):
        state = _state([
            _strat("A", symbol="ES"),
            _strat("B", symbol="S"),
            _strat("C", symbol="YM"),
        ])
        state = step_filter_excluded_symbols(state, ["S", "YM"])
        assert {c["name"] for c in state.candidates} == {"A"}


# ── step_filter_contract_size ──────────────────────────────────────────────────

class TestStepFilterContractSize:
    def test_removes_below_threshold(self):
        state = _state([_strat("A"), _strat("B")])
        state.contracts = {"A": 0.5, "B": 1.0}
        state = step_filter_contract_size(state, min_threshold=0.65)
        assert {c["name"] for c in state.candidates} == {"B"}

    def test_keeps_at_or_above_threshold(self):
        state = _state([_strat("A"), _strat("B")])
        state.contracts = {"A": 0.65, "B": 2.0}
        state = step_filter_contract_size(state, min_threshold=0.65)
        assert len(state.candidates) == 2

    def test_zero_contracts_removed(self):
        state = _state([_strat("A")])
        state.contracts = {"A": 0.0}
        state = step_filter_contract_size(state, min_threshold=0.1)
        assert len(state.candidates) == 0


# ── step_rank ─────────────────────────────────────────────────────────────────

class TestStepRank:
    def test_descending_default(self):
        strats = [
            _strat("A", rtd_oos=1.0),
            _strat("B", rtd_oos=3.0),
            _strat("C", rtd_oos=2.0),
        ]
        state = _state(strats)
        state = step_rank(state, metric="rtd_oos", ascending=False)
        names = [c["name"] for c in state.candidates]
        assert names == ["B", "C", "A"]

    def test_ascending(self):
        strats = [
            _strat("A", ulcer=5.0),
            _strat("B", ulcer=1.0),
            _strat("C", ulcer=3.0),
        ]
        state = _state(strats)
        state = step_rank(state, metric="ulcer", ascending=True)
        names = [c["name"] for c in state.candidates]
        assert names == ["B", "C", "A"]

    def test_none_values_sort_last_descending(self):
        strats = [_strat("A", rtd_oos=None), _strat("B", rtd_oos=2.0)]
        state = _state(strats)
        state = step_rank(state, metric="rtd_oos", ascending=False)
        assert state.candidates[0]["name"] == "B"

    def test_missing_metric_key_sorts_last(self):
        strats = [_strat("A"), _strat("B", rtd_oos=5.0)]
        state = _state(strats)
        state = step_rank(state, metric="rtd_oos", ascending=False)
        assert state.candidates[0]["name"] == "B"


# ── step_size_contracts ────────────────────────────────────────────────────────

class TestStepSizeContracts:
    def _atr(self) -> dict[str, float]:
        return {"A": 1_000.0, "B": 500.0}

    def test_pure_atr_sizing(self):
        # equity=500k, pct=0.01, atr=1000, ratio=1.0 → raw=5, rounded to 0.1 → 5.0
        state = _state([_strat("A", symbol="ES")])
        state = step_size_contracts(
            state, equity=500_000, contract_size_pct=0.01,
            atr={"A": 1_000.0}, margins=_MARGINS, ratio=1.0,
            contract_margin_multiple=1.0, min_fraction=0.1,
        )
        assert state.contracts["A"] == pytest.approx(5.0)

    def test_rounds_down_to_fraction(self):
        # raw = 500_000 * 0.01 / 750 = 6.666... → floor to 0.1 → 6.6
        state = _state([_strat("A", symbol="ES")])
        state = step_size_contracts(
            state, equity=500_000, contract_size_pct=0.01,
            atr={"A": 750.0}, margins=_MARGINS, ratio=1.0,
            contract_margin_multiple=1.0, min_fraction=0.1,
        )
        assert state.contracts["A"] == pytest.approx(6.6)

    def test_equity_updated_on_state(self):
        state = _state([_strat("A")])
        state = step_size_contracts(
            state, equity=999_999, contract_size_pct=0.01,
            atr={}, margins={}, ratio=0.5, contract_margin_multiple=1.0,
        )
        assert state.equity == pytest.approx(999_999)

    def test_zero_atr_uses_margin_only(self):
        # atr=0, ratio=0.5 → dollar_risk = 0*0.5 + 10000*1.0*0.5 = 5000
        # raw = 500_000 * 0.01 / 5000 = 1.0
        state = _state([_strat("A", symbol="ES")])
        state = step_size_contracts(
            state, equity=500_000, contract_size_pct=0.01,
            atr={"A": 0.0}, margins=_MARGINS, ratio=0.5,
            contract_margin_multiple=1.0, min_fraction=0.1,
        )
        assert state.contracts["A"] == pytest.approx(1.0)


# ── step_select_strategies ────────────────────────────────────────────────────

class TestStepSelectStrategies:
    def _sized_state(self, strategies: list[dict], equity: float = 500_000.0) -> OptimizerState:
        state = _state(strategies, equity)
        # Give each strategy 1.0 contract
        state.contracts = {s["name"]: 1.0 for s in strategies}
        return state

    def test_respects_max_strategies(self):
        strats = [_strat(f"S{i}", symbol=f"X{i}") for i in range(10)]
        state = self._sized_state(strats)
        state = step_select_strategies(
            state, margins={f"X{i}": 100.0 for i in range(10)},
            contract_margin_multiple=1.0,
            max_margin_pct=1.0, max_strategies=5,
        )
        assert len(state.candidates) == 5

    def test_respects_margin_cap(self):
        # Each strategy: 1 contract × 50k margin = 50k
        # equity=500k, max_margin_pct=0.30 → 150k → 3 strategies max
        strats = [_strat(f"S{i}", symbol=f"Y{i}") for i in range(10)]
        state = self._sized_state(strats, equity=500_000)
        state = step_select_strategies(
            state, margins={f"Y{i}": 50_000.0 for i in range(10)},
            contract_margin_multiple=1.0,
            max_margin_pct=0.30, max_strategies=60,
            per_symbol_first=False,
        )
        assert len(state.candidates) <= 3

    def test_per_symbol_first_includes_unique_symbols(self):
        # 3 strategies on ES, 1 on NQ — with per_symbol_first, best-ES + NQ first
        strats = [
            _strat("ES_A", symbol="ES"),
            _strat("ES_B", symbol="ES"),
            _strat("ES_C", symbol="ES"),
            _strat("NQ_A", symbol="NQ"),
        ]
        state = self._sized_state(strats, equity=200_000)
        state = step_select_strategies(
            state, margins={"ES": 10_000.0, "NQ": 15_000.0},
            contract_margin_multiple=1.0,
            max_margin_pct=0.30, max_strategies=60,
            per_symbol_first=True,
        )
        selected = {c["name"] for c in state.candidates}
        # ES_A is best ES, NQ_A is only NQ — both should be in
        assert "ES_A" in selected
        assert "NQ_A" in selected

    def test_not_selected_added_to_excluded(self):
        strats = [_strat(f"S{i}", symbol=f"Z{i}") for i in range(5)]
        state = self._sized_state(strats, equity=100_000)
        state = step_select_strategies(
            state, margins={f"Z{i}": 100_000.0 for i in range(5)},
            contract_margin_multiple=1.0,
            max_margin_pct=0.10, max_strategies=60,
        )
        # Only 1 strategy fits (100k margin = 100% of 100k equity, but cap=10%)
        assert len(state.excluded) >= 1


# ── step_adjust_correlations ──────────────────────────────────────────────────

class TestStepAdjustCorrelations:
    def _corr_matrix(self, data: dict) -> pd.DataFrame:
        return pd.DataFrame(data)

    def test_removes_high_correlation(self):
        # A and B are 0.9 correlated — B (lower ranked = index 1) is removed
        corr = pd.DataFrame(
            {"A": [1.0, 0.9], "B": [0.9, 1.0]},
            index=["A", "B"],
        )
        state = _state([_strat("A"), _strat("B")])
        state = step_adjust_correlations(state, corr_matrix=corr, max_corr=0.70)
        assert {c["name"] for c in state.candidates} == {"A"}

    def test_removes_high_negative_correlation(self):
        corr = pd.DataFrame(
            {"A": [1.0, -0.8], "B": [-0.8, 1.0]},
            index=["A", "B"],
        )
        state = _state([_strat("A"), _strat("B")])
        state = step_adjust_correlations(
            state, corr_matrix=corr, max_corr=0.70, max_neg_corr=0.50
        )
        assert {c["name"] for c in state.candidates} == {"A"}

    def test_low_correlation_no_removal(self):
        corr = pd.DataFrame(
            {"A": [1.0, 0.3], "B": [0.3, 1.0]},
            index=["A", "B"],
        )
        state = _state([_strat("A"), _strat("B")])
        state = step_adjust_correlations(state, corr_matrix=corr, max_corr=0.70)
        assert len(state.candidates) == 2

    def test_none_corr_matrix_skipped(self):
        state = _state([_strat("A"), _strat("B")])
        before = len(state.candidates)
        state = step_adjust_correlations(state, corr_matrix=None)
        assert len(state.candidates) == before

    def test_higher_ranked_strategy_kept(self):
        """A (index 0 = highest rank) is kept; B (index 1) is removed."""
        corr = pd.DataFrame(
            {"A": [1.0, 0.95], "B": [0.95, 1.0]},
            index=["A", "B"],
        )
        state = _state([_strat("A"), _strat("B")])  # A is rank 1, B is rank 2
        state = step_adjust_correlations(state, corr_matrix=corr, max_corr=0.70)
        assert any(c["name"] == "A" for c in state.candidates)
        assert not any(c["name"] == "B" for c in state.candidates)


# ── step_adjust_gross_margins ─────────────────────────────────────────────────

class TestStepAdjustGrossMargins:
    def _state_with_contracts(self, strategies: list[dict]) -> OptimizerState:
        state = _state(strategies, equity=500_000)
        state.contracts = {s["name"]: float(s.get("contracts", 1)) for s in strategies}
        return state

    def test_removes_when_symbol_exceeds_limit(self):
        # Both on ES, 1 contract × 10k each = 20k total
        # max_single_pct=0.30 → 6k cap per symbol → ES=20k / 20k = 100% > 30%
        # So remove worst (last in candidates)
        strats = [_strat("A", symbol="ES", contracts=1), _strat("B", symbol="ES", contracts=1)]
        state = self._state_with_contracts(strats)
        state = step_adjust_gross_margins(
            state, margins={"ES": 10_000.0}, contract_margin_multiple=1.0,
            equity=500_000, max_single_pct=0.30, max_sector_pct=1.0,
        )
        # One of them should be removed
        assert len(state.candidates) < 2

    def test_compliant_portfolio_unchanged(self):
        strats = [
            _strat("A", symbol="ES", sector="Index", contracts=1),
            _strat("B", symbol="NQ", sector="Index", contracts=1),
            _strat("C", symbol="CL", sector="Energy", contracts=1),
        ]
        state = self._state_with_contracts(strats)
        # Total margin = 10k + 15k + 5k = 30k
        # ES share = 10k/30k = 33% > 12.5% → will trim
        state = step_adjust_gross_margins(
            state, margins=_MARGINS, contract_margin_multiple=1.0,
            equity=500_000, max_single_pct=0.50, max_sector_pct=1.0,
        )
        # With 50% single limit, no removals needed
        assert len(state.candidates) == 3

    def test_removes_when_sector_exceeds_limit(self):
        strats = [
            _strat("A", symbol="ES", sector="Index", contracts=1),
            _strat("B", symbol="NQ", sector="Index", contracts=1),
        ]
        state = self._state_with_contracts(strats)
        # Total = 25k, Index = 25k = 100% > max_sector_pct=0.40
        state = step_adjust_gross_margins(
            state, margins=_MARGINS, contract_margin_multiple=1.0,
            equity=500_000, max_single_pct=1.0, max_sector_pct=0.40,
        )
        assert len(state.candidates) < 2


# ── step_adjust_drawdowns ─────────────────────────────────────────────────────

class TestStepAdjustDrawdowns:
    def test_reduces_contracts_when_single_dd_too_large(self):
        # max_oos_drawdown=50k per contract, 2 contracts → 100k
        # equity=500k, max_single=0.05 → limit=25k → must reduce to 0.5 contracts
        strat = _strat("A", max_oos_drawdown=50_000.0, contracts=2)
        state = _state([strat], equity=500_000)
        state.contracts = {"A": 2.0}
        state = step_adjust_drawdowns(
            state, equity=500_000, max_avg_pct=0.20, max_single_pct=0.05,
            max_single_trade_pct=0.10, min_fraction=0.1,
        )
        assert state.contracts["A"] < 2.0
        assert state.contracts["A"] * 50_000 <= 0.05 * 500_000 + 0.1  # within limit

    def test_compliant_drawdown_unchanged(self):
        # 1 contract × 5k drawdown = 5k < 0.05 × 500k = 25k → OK
        strat = _strat("A", max_oos_drawdown=5_000.0, contracts=1)
        state = _state([strat], equity=500_000)
        state.contracts = {"A": 1.0}
        state = step_adjust_drawdowns(
            state, equity=500_000, max_avg_pct=0.05, max_single_pct=0.125,
            max_single_trade_pct=0.05, min_fraction=0.1,
        )
        assert state.contracts["A"] == pytest.approx(1.0)

    def test_missing_drawdown_not_adjusted(self):
        """Strategies with no drawdown data are left unchanged."""
        strat = _strat("A")  # no max_oos_drawdown
        state = _state([strat], equity=500_000)
        state.contracts = {"A": 5.0}
        state = step_adjust_drawdowns(
            state, equity=500_000, max_avg_pct=0.01, max_single_pct=0.01,
            max_single_trade_pct=0.01, min_fraction=0.1,
        )
        assert state.contracts["A"] == pytest.approx(5.0)


# ── build_candidates ──────────────────────────────────────────────────────────

class TestBuildCandidates:
    def test_merges_summary_metrics(self):
        strats = [{"name": "A", "symbol": "ES", "sector": "Index"}]
        summary = pd.DataFrame({"rtd_oos": [2.5]}, index=["A"])
        result = build_candidates(strats, summary, None, {}, 5_000.0)
        assert result[0]["rtd_oos"] == pytest.approx(2.5)

    def test_atr_added(self):
        strats = [{"name": "A", "symbol": "ES", "sector": "Index"}]
        atr = pd.Series({"A": 1_200.0})
        result = build_candidates(strats, None, atr, {}, 5_000.0)
        assert result[0]["atr"] == pytest.approx(1_200.0)

    def test_atr_zero_when_not_available(self):
        strats = [{"name": "A", "symbol": "ES", "sector": "Index"}]
        result = build_candidates(strats, None, None, {}, 5_000.0)
        assert result[0]["atr"] == pytest.approx(0.0)

    def test_margin_per_contract_set(self):
        strats = [{"name": "A", "symbol": "ES", "sector": "Index"}]
        result = build_candidates(strats, None, None, {"ES": 12_000.0}, 5_000.0)
        assert result[0]["margin_per_contract"] == pytest.approx(12_000.0)

    def test_default_margin_when_symbol_missing(self):
        strats = [{"name": "A", "symbol": "ZZ", "sector": "Other"}]
        result = build_candidates(strats, None, None, {"ES": 12_000.0}, 7_500.0)
        assert result[0]["margin_per_contract"] == pytest.approx(7_500.0)

    def test_existing_strategy_keys_not_overwritten_by_summary(self):
        """Strategy config keys (name, symbol, sector) aren't overwritten."""
        strats = [{"name": "A", "symbol": "ES", "sector": "Index", "contracts": 5}]
        summary = pd.DataFrame({"sector": ["Energy"]}, index=["A"])  # conflict
        result = build_candidates(strats, summary, None, {}, 5_000.0)
        assert result[0]["sector"] == "Index"


# ── run_workflow ───────────────────────────────────────────────────────────────

class TestRunWorkflow:
    def test_empty_workflow_returns_all_candidates(self):
        strats = [_strat("A"), _strat("B")]
        state = run_workflow([], strats, equity=500_000)
        assert len(state.candidates) == 2

    def test_step_errors_logged_not_raised(self):
        def _bad_step(state, **kwargs):
            raise RuntimeError("Boom!")

        strats = [_strat("A")]
        state = run_workflow([(_bad_step, {})], strats, equity=500_000)
        assert any("ERROR" in line for line in state.log)
        assert len(state.candidates) == 1  # not mutated

    def test_steps_executed_in_order(self):
        """rank then select: rank should affect which strategies are kept."""
        strats = [
            _strat("A", symbol="ES", rtd_oos=1.0, max_oos_drawdown=1_000.0),
            _strat("B", symbol="NQ", rtd_oos=5.0, max_oos_drawdown=1_000.0),
        ]
        steps = [
            (step_rank, {"metric": "rtd_oos", "ascending": False}),
            (step_select_strategies, {
                "margins": {"ES": 10_000.0, "NQ": 15_000.0},
                "contract_margin_multiple": 1.0,
                "max_margin_pct": 0.05,  # tight cap → only 1 fits per 500k equity
                "max_strategies": 60,
                "per_symbol_first": False,
            }),
        ]
        state = run_workflow(steps, strats, equity=500_000)
        # B has higher rtd_oos and fits within 0.05×500k=25k with NQ margin=15k
        assert state.candidates[0]["name"] == "B"


# ── portfolio_summary ─────────────────────────────────────────────────────────

class TestPortfolioSummary:
    def test_basic_stats(self):
        strats = [
            _strat("A", symbol="ES", sector="Index"),
            _strat("B", symbol="CL", sector="Energy"),
        ]
        state = _state(strats, equity=100_000)
        state.contracts = {"A": 1.0, "B": 2.0}
        stats = portfolio_summary(state, _MARGINS, _MARGIN_MULT)
        # total margin = 1×10k + 2×5k = 20k
        assert stats["total_margin"] == pytest.approx(20_000.0)
        assert stats["n_strategies"] == 2
        assert stats["margin_pct_equity"] == pytest.approx(0.20)

    def test_zero_equity_no_crash(self):
        state = _state([], equity=0)
        stats = portfolio_summary(state, {}, 1.0)
        assert stats["n_strategies"] == 0
        assert stats["total_margin"] == pytest.approx(0.0)


# ── PortfolioOptimizerConfig round-trip ───────────────────────────────────────

class TestPortfolioOptimizerConfig:
    def test_defaults(self):
        from core.config import PortfolioOptimizerConfig
        cfg = PortfolioOptimizerConfig()
        assert "filter_eligibility" in cfg.workflow_steps
        assert "rank" in cfg.workflow_steps
        assert cfg.max_margin_pct == pytest.approx(0.75)
        assert cfg.max_correlation == pytest.approx(0.70)
        assert cfg.min_contract_size_threshold == pytest.approx(0.65)

    def test_appconfig_has_optimizer(self):
        from core.config import AppConfig, PortfolioOptimizerConfig
        cfg = AppConfig()
        assert hasattr(cfg, "optimizer")
        assert isinstance(cfg.optimizer, PortfolioOptimizerConfig)

    def test_yaml_round_trip(self, tmp_path):
        import yaml
        from core.config import AppConfig

        cfg = AppConfig()
        cfg.optimizer.max_strategies = 45
        cfg.optimizer.excluded_symbols = ["S", "YM"]
        cfg.optimizer.rank_metric = "profit_last_12_months"

        data = cfg.model_dump(mode="json")
        data["folders"] = [str(p) for p in cfg.folders]
        yaml_str = yaml.dump(data, default_flow_style=False)
        reloaded = AppConfig.model_validate(yaml.safe_load(yaml_str))

        assert reloaded.optimizer.max_strategies == 45
        assert "S" in reloaded.optimizer.excluded_symbols
        assert reloaded.optimizer.rank_metric == "profit_last_12_months"
