"""
Unit tests for apply_eligibility_rules() in core/portfolio/summary.py

Verifies each toggle category independently:
  - Profit qualifiers (1m / 3m / 6m / 3or6m / 9m / 12m / oos)
  - Efficiency qualifiers
  - Profit disqualifiers (loss_1m / loss_3m / loss_6m)
  - Incubation gate
  - Quitting gate
  - Count monthly profits
"""

from __future__ import annotations

import pandas as pd
import pytest

from core.config import EligibilityConfig
from core.portfolio.summary import apply_eligibility_rules


def _make_df(**kwargs) -> pd.DataFrame:
    """
    Build a single-row summary DataFrame with sensible defaults.
    Override any column via kwargs.
    """
    defaults = {
        "profit_last_1_month":  1000.0,
        "profit_last_3_months": 3000.0,
        "profit_last_6_months": 6000.0,
        "profit_last_9_months": 9000.0,
        "profit_last_12_months": 12000.0,
        "profit_since_oos_start": 50000.0,
        "efficiency_last_1_month":  0.5,
        "efficiency_last_3_months": 0.5,
        "efficiency_last_6_months": 0.5,
        "efficiency_last_9_months": 0.5,
        "efficiency_last_12_months": 0.5,
        "return_efficiency": 0.5,
        "incubation_status": "Passed",
        "quitting_status": "Continue",
        "count_profit_months": 10,
    }
    defaults.update(kwargs)
    return pd.DataFrame([defaults], index=["strat_a"])


def _bare_config(**kwargs) -> EligibilityConfig:
    """All toggles OFF; override specific ones via kwargs."""
    cfg = EligibilityConfig(
        profit_1m=False, profit_3m=False, profit_6m=False,
        profit_3or6m=False, profit_9m=False, profit_12m=False, profit_oos=False,
        efficiency_1m=False, efficiency_3m=False, efficiency_6m=False,
        efficiency_9m=False, efficiency_12m=False, efficiency_oos=False,
        loss_1m=False, loss_3m=False, loss_6m=False,
        efficiency_loss_1m=False, efficiency_loss_3m=False, efficiency_loss_6m=False,
        use_incubation=False, use_quitting=False,
        use_count_monthly_profits=False,
        efficiency_ratio=0.15,
    )
    for k, v in kwargs.items():
        setattr(cfg, k, v)
    return cfg


class TestProfitQualifiers:
    def test_all_off_always_eligible(self):
        df = _make_df(profit_last_1_month=-9999.0)
        result = apply_eligibility_rules(df, _bare_config())
        assert result["strat_a"] == True

    def test_profit_1m_pass(self):
        df = _make_df(profit_last_1_month=1.0)
        result = apply_eligibility_rules(df, _bare_config(profit_1m=True))
        assert result["strat_a"] == True

    def test_profit_1m_fail(self):
        df = _make_df(profit_last_1_month=-1.0)
        result = apply_eligibility_rules(df, _bare_config(profit_1m=True))
        assert result["strat_a"] == False

    def test_profit_12m_pass(self):
        df = _make_df(profit_last_12_months=100.0)
        result = apply_eligibility_rules(df, _bare_config(profit_12m=True))
        assert result["strat_a"] == True

    def test_profit_12m_fail(self):
        df = _make_df(profit_last_12_months=-100.0)
        result = apply_eligibility_rules(df, _bare_config(profit_12m=True))
        assert result["strat_a"] == False

    def test_profit_oos_pass(self):
        df = _make_df(profit_since_oos_start=1.0)
        result = apply_eligibility_rules(df, _bare_config(profit_oos=True))
        assert result["strat_a"] == True

    def test_profit_oos_fail(self):
        df = _make_df(profit_since_oos_start=-1.0)
        result = apply_eligibility_rules(df, _bare_config(profit_oos=True))
        assert result["strat_a"] == False

    def test_profit_3or6m_pass_via_3m(self):
        df = _make_df(profit_last_3_months=1.0, profit_last_6_months=-1.0)
        result = apply_eligibility_rules(df, _bare_config(profit_3or6m=True))
        assert result["strat_a"] == True

    def test_profit_3or6m_pass_via_6m(self):
        df = _make_df(profit_last_3_months=-1.0, profit_last_6_months=1.0)
        result = apply_eligibility_rules(df, _bare_config(profit_3or6m=True))
        assert result["strat_a"] == True

    def test_profit_3or6m_fail_both_negative(self):
        df = _make_df(profit_last_3_months=-1.0, profit_last_6_months=-1.0)
        result = apply_eligibility_rules(df, _bare_config(profit_3or6m=True))
        assert result["strat_a"] == False


class TestEfficiencyQualifiers:
    def test_efficiency_oos_pass(self):
        df = _make_df(return_efficiency=0.5)
        result = apply_eligibility_rules(df, _bare_config(efficiency_oos=True, efficiency_ratio=0.15))
        assert result["strat_a"] == True

    def test_efficiency_oos_fail(self):
        df = _make_df(return_efficiency=0.1)
        result = apply_eligibility_rules(df, _bare_config(efficiency_oos=True, efficiency_ratio=0.15))
        assert result["strat_a"] == False

    def test_efficiency_12m_pass(self):
        df = _make_df(efficiency_last_12_months=0.5)
        result = apply_eligibility_rules(df, _bare_config(efficiency_12m=True, efficiency_ratio=0.15))
        assert result["strat_a"] == True

    def test_efficiency_12m_fail(self):
        df = _make_df(efficiency_last_12_months=0.05)
        result = apply_eligibility_rules(df, _bare_config(efficiency_12m=True, efficiency_ratio=0.15))
        assert result["strat_a"] == False


class TestDisqualifiers:
    def test_loss_1m_disqualifies_on_negative_profit(self):
        df = _make_df(profit_last_1_month=-500.0)
        result = apply_eligibility_rules(df, _bare_config(loss_1m=True))
        assert result["strat_a"] == False

    def test_loss_1m_does_not_disqualify_on_zero(self):
        df = _make_df(profit_last_1_month=0.0)
        result = apply_eligibility_rules(df, _bare_config(loss_1m=True))
        assert result["strat_a"] == True

    def test_loss_3m_disqualifies(self):
        df = _make_df(profit_last_3_months=-1.0)
        result = apply_eligibility_rules(df, _bare_config(loss_3m=True))
        assert result["strat_a"] == False

    def test_loss_6m_disqualifies(self):
        df = _make_df(profit_last_6_months=-1.0)
        result = apply_eligibility_rules(df, _bare_config(loss_6m=True))
        assert result["strat_a"] == False

    def test_efficiency_loss_disqualifies(self):
        df = _make_df(efficiency_last_3_months=-0.5)
        result = apply_eligibility_rules(df, _bare_config(efficiency_loss_3m=True, efficiency_ratio=0.15))
        assert result["strat_a"] == False

    def test_efficiency_loss_does_not_disqualify_mildly_negative(self):
        # efficiency = -0.05, threshold = -0.15 → should NOT disqualify
        df = _make_df(efficiency_last_3_months=-0.05)
        result = apply_eligibility_rules(df, _bare_config(efficiency_loss_3m=True, efficiency_ratio=0.15))
        assert result["strat_a"] == True


class TestIncubationGate:
    def test_passed_incubation_eligible(self):
        df = _make_df(incubation_status="Passed")
        result = apply_eligibility_rules(df, _bare_config(use_incubation=True))
        assert result["strat_a"] == True

    def test_empty_incubation_eligible(self):
        # "" means no expected profit / no OOS data → don't block
        df = _make_df(incubation_status="")
        result = apply_eligibility_rules(df, _bare_config(use_incubation=True))
        assert result["strat_a"] == True

    def test_not_passed_incubation_ineligible(self):
        df = _make_df(incubation_status="Not Passed")
        result = apply_eligibility_rules(df, _bare_config(use_incubation=True))
        assert result["strat_a"] == False

    def test_incubating_ineligible(self):
        df = _make_df(incubation_status="Incubating")
        result = apply_eligibility_rules(df, _bare_config(use_incubation=True))
        assert result["strat_a"] == False

    def test_incubation_gate_off_does_not_filter(self):
        df = _make_df(incubation_status="Not Passed")
        result = apply_eligibility_rules(df, _bare_config(use_incubation=False))
        assert result["strat_a"] == True


class TestQuittingGate:
    def test_continue_eligible(self):
        df = _make_df(quitting_status="Continue")
        result = apply_eligibility_rules(df, _bare_config(use_quitting=True))
        assert result["strat_a"] == True

    def test_recovered_eligible(self):
        df = _make_df(quitting_status="Recovered")
        result = apply_eligibility_rules(df, _bare_config(use_quitting=True))
        assert result["strat_a"] == True

    def test_quit_ineligible(self):
        df = _make_df(quitting_status="Quit")
        result = apply_eligibility_rules(df, _bare_config(use_quitting=True))
        assert result["strat_a"] == False

    def test_coming_back_ineligible(self):
        df = _make_df(quitting_status="Coming Back")
        result = apply_eligibility_rules(df, _bare_config(use_quitting=True))
        assert result["strat_a"] == False

    def test_quitting_gate_off_does_not_filter(self):
        df = _make_df(quitting_status="Quit")
        result = apply_eligibility_rules(df, _bare_config(use_quitting=False))
        assert result["strat_a"] == True


class TestCountMonthlyProfits:
    def test_enough_months_eligible(self):
        df = _make_df(count_profit_months=9)
        result = apply_eligibility_rules(
            df,
            _bare_config(
                use_count_monthly_profits=True,
                min_positive_months=8,
                monthly_profit_operator=">0",
            ),
        )
        assert result["strat_a"] == True

    def test_not_enough_months_ineligible(self):
        df = _make_df(count_profit_months=7)
        result = apply_eligibility_rules(
            df,
            _bare_config(
                use_count_monthly_profits=True,
                min_positive_months=8,
                monthly_profit_operator=">0",
            ),
        )
        assert result["strat_a"] == False

    def test_ge_operator_exact_boundary_eligible(self):
        df = _make_df(count_profit_months=8)
        result = apply_eligibility_rules(
            df,
            _bare_config(
                use_count_monthly_profits=True,
                min_positive_months=8,
                monthly_profit_operator=">=0",
            ),
        )
        assert result["strat_a"] == True


class TestMultipleRules:
    def test_all_default_toggles_eligible_strategy(self):
        """The default config toggles (profit_3or6m, profit_12m, efficiency_oos) all pass."""
        df = _make_df(
            profit_last_3_months=1000.0,
            profit_last_12_months=12000.0,
            return_efficiency=0.5,
            incubation_status="Passed",
            quitting_status="Continue",
        )
        # Simulate the default config
        cfg = EligibilityConfig()  # uses all defaults from config.py
        result = apply_eligibility_rules(df, cfg)
        assert result["strat_a"] == True

    def test_multiple_rows(self):
        df = pd.DataFrame([
            {"profit_last_12_months": 1000.0, "return_efficiency": 0.5,
             "profit_last_3_months": 100.0, "profit_last_6_months": 200.0,
             "incubation_status": "Passed", "quitting_status": "Continue",
             "count_profit_months": 10},
            {"profit_last_12_months": -100.0, "return_efficiency": 0.5,
             "profit_last_3_months": 100.0, "profit_last_6_months": 200.0,
             "incubation_status": "Passed", "quitting_status": "Continue",
             "count_profit_months": 10},
        ], index=["strat_pass", "strat_fail"])
        cfg = _bare_config(profit_12m=True)
        result = apply_eligibility_rules(df, cfg)
        assert result["strat_pass"] == True
        assert result["strat_fail"] == False
