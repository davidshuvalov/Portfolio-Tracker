"""
Unit tests for core/analytics/eligibility/rules.py

Tests:
  - Rule primitives: _last_n_sum, _consecutive_positive, _count_positive_ge,
    _momentum_positive, _acceleration_positive, _threshold_annual, _recovery
  - evaluate_rule: each RuleType with pass/fail cases
  - require_oos_profitable gate
  - build_rule_catalogue: count, unique ids, all section types present
"""

from __future__ import annotations

import numpy as np
import pytest

from core.analytics.eligibility.rules import (
    EligibilityRule,
    RuleType,
    _acceleration_positive,
    _consecutive_positive,
    _count_positive_ge,
    _last_n_sum,
    _momentum_positive,
    _recovery,
    _threshold_annual,
    build_rule_catalogue,
    evaluate_rule,
)


# ── _last_n_sum ───────────────────────────────────────────────────────────────

class TestLastNSum:
    def test_basic_sum(self):
        arr = np.array([1.0, 2.0, 3.0, 4.0, 5.0])
        assert _last_n_sum(arr, 3) == pytest.approx(12.0)

    def test_n_equals_length(self):
        arr = np.array([1.0, -2.0, 3.0])
        assert _last_n_sum(arr, 3) == pytest.approx(2.0)

    def test_n_greater_than_length_returns_zero(self):
        arr = np.array([1.0, 2.0])
        assert _last_n_sum(arr, 5) == 0.0

    def test_n_zero_returns_zero(self):
        arr = np.array([1.0, 2.0, 3.0])
        assert _last_n_sum(arr, 0) == 0.0

    def test_empty_array(self):
        assert _last_n_sum(np.array([]), 3) == 0.0


# ── _consecutive_positive ─────────────────────────────────────────────────────

class TestConsecutivePositive:
    def test_all_positive(self):
        arr = np.array([1.0, 2.0, 3.0, 4.0])
        assert _consecutive_positive(arr, 3) is True

    def test_last_negative(self):
        arr = np.array([1.0, 2.0, 3.0, -1.0])
        assert _consecutive_positive(arr, 3) is False

    def test_zero_fails(self):
        arr = np.array([1.0, 0.0, 1.0])
        assert _consecutive_positive(arr, 3) is False  # 0 is not > 0

    def test_insufficient_history(self):
        arr = np.array([1.0, 2.0])
        assert _consecutive_positive(arr, 5) is False

    def test_n_zero_returns_false(self):
        arr = np.array([1.0, 2.0, 3.0])
        assert _consecutive_positive(arr, 0) is False


# ── _count_positive_ge ────────────────────────────────────────────────────────

class TestCountPositiveGe:
    def test_2_of_3_passes(self):
        arr = np.array([1.0, -1.0, 1.0])
        assert _count_positive_ge(arr, 2, 3) is True

    def test_1_of_3_fails_when_need_2(self):
        arr = np.array([1.0, -1.0, -1.0])
        assert _count_positive_ge(arr, 2, 3) is False

    def test_exact_k(self):
        arr = np.array([1.0, 1.0, -1.0, -1.0, -1.0, 1.0])
        # last 6: 3 positive → passes k=3 n=6
        assert _count_positive_ge(arr, 3, 6) is True

    def test_insufficient_history(self):
        arr = np.array([1.0])
        assert _count_positive_ge(arr, 1, 3) is False


# ── _momentum_positive ────────────────────────────────────────────────────────

class TestMomentumPositive:
    def test_recent_greater(self):
        # Last 3M = [200, 200, 200] = 600; prior 3M = [100, 100, 100] = 300
        arr = np.array([100.0, 100.0, 100.0, 200.0, 200.0, 200.0])
        assert _momentum_positive(arr, 3, 6) is True

    def test_recent_less(self):
        arr = np.array([200.0, 200.0, 200.0, 100.0, 100.0, 100.0])
        assert _momentum_positive(arr, 3, 6) is False

    def test_insufficient_history(self):
        arr = np.array([1.0, 2.0])
        assert _momentum_positive(arr, 3, 6) is False

    def test_short_ge_long_returns_false(self):
        arr = np.array([1.0] * 10)
        assert _momentum_positive(arr, 6, 3) is False  # short >= long


# ── _acceleration_positive ────────────────────────────────────────────────────

class TestAccelerationPositive:
    def test_accelerating(self):
        # ann short (3M × 4 = 12M equiv): 500×4=2000 > ann long (6M × 2): 600×2=1200
        arr = np.array([100.0, 100.0, 100.0, 500.0, 500.0, 500.0])
        assert _acceleration_positive(arr, 3, 6) is True

    def test_decelerating(self):
        arr = np.array([500.0, 500.0, 500.0, 100.0, 100.0, 100.0])
        assert _acceleration_positive(arr, 3, 6) is False

    def test_insufficient_history(self):
        arr = np.array([100.0, 200.0])
        assert _acceleration_positive(arr, 3, 6) is False


# ── _threshold_annual ─────────────────────────────────────────────────────────

class TestThresholdAnnual:
    def test_meets_threshold(self):
        # Last 6M = 600 → ann = 600 × 2 = 1200; expected_annual=1000, ratio=0.5 → target=500
        arr = np.full(6, 100.0)
        assert _threshold_annual(arr, 6, 1000.0, 0.5) is True

    def test_below_threshold(self):
        # Last 6M = 60 → ann = 120; target = 500
        arr = np.full(6, 10.0)
        assert _threshold_annual(arr, 6, 1000.0, 0.5) is False

    def test_zero_expected_returns_false(self):
        arr = np.full(6, 100.0)
        assert _threshold_annual(arr, 6, 0.0, 0.5) is False

    def test_insufficient_history(self):
        arr = np.array([100.0, 200.0])
        assert _threshold_annual(arr, 6, 1000.0, 0.5) is False


# ── _recovery ─────────────────────────────────────────────────────────────────

class TestRecovery:
    def test_recovery_passes(self):
        # pos_n=1 (last 1M positive), neg_n=1 (prior 1M negative)
        arr = np.array([-200.0, 100.0])
        assert _recovery(arr, 1, 1) is True

    def test_no_prior_negative(self):
        arr = np.array([100.0, 100.0])
        assert _recovery(arr, 1, 1) is False

    def test_recent_negative(self):
        arr = np.array([-200.0, -100.0])
        assert _recovery(arr, 1, 1) is False

    def test_3m_recovery(self):
        arr = np.array([-100.0, -100.0, -100.0, 50.0, 50.0, 50.0])
        assert _recovery(arr, 3, 3) is True

    def test_insufficient_history(self):
        arr = np.array([100.0])
        assert _recovery(arr, 1, 1) is False


# ── evaluate_rule ─────────────────────────────────────────────────────────────

class TestEvaluateRule:
    def _rule(self, rule_type, p1=0, p2=0, p3=0, oos=False):
        return EligibilityRule(
            id=0, label="test", rule_type=rule_type,
            param1=p1, param2=p2, param3=p3,
            require_oos_profitable=oos,
        )

    def test_baseline_always_true(self):
        arr = np.zeros(10)
        rule = self._rule(RuleType.BASELINE)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is True

    def test_oos_profitable_pass(self):
        arr = np.array([100.0, 200.0])
        rule = self._rule(RuleType.OOS_PROFITABLE)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is True

    def test_oos_profitable_fail(self):
        arr = np.array([-100.0, -200.0])
        rule = self._rule(RuleType.OOS_PROFITABLE)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is False

    def test_simple_positive_pass(self):
        arr = np.array([100.0, 200.0, 300.0])
        rule = self._rule(RuleType.SIMPLE_POSITIVE, p1=3)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is True

    def test_simple_positive_fail(self):
        arr = np.array([-100.0, -200.0, -300.0])
        rule = self._rule(RuleType.SIMPLE_POSITIVE, p1=3)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is False

    def test_simple_negative_pass(self):
        arr = np.array([-100.0, -200.0, -300.0])
        rule = self._rule(RuleType.SIMPLE_NEGATIVE, p1=3)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is True

    def test_consecutive_pass(self):
        arr = np.array([100.0, 200.0, 300.0])
        rule = self._rule(RuleType.CONSECUTIVE, p1=3)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is True

    def test_count_positive_pass(self):
        arr = np.array([100.0, -50.0, 200.0])
        rule = self._rule(RuleType.COUNT_POSITIVE, p1=2, p2=3)  # 2 of 3
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is True

    def test_and_combo_both_positive(self):
        arr = np.array([100.0, 100.0, 100.0, 100.0, 100.0, 100.0])
        rule = self._rule(RuleType.AND_COMBO, p1=3, p2=6)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is True

    def test_and_combo_one_negative(self):
        arr = np.array([-100.0, -100.0, -100.0, 100.0, 100.0, 100.0])
        rule = self._rule(RuleType.AND_COMBO, p1=3, p2=6)
        # last 3M positive, but last 6M negative overall → combo fails
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is False

    def test_any_of_3_one_positive(self):
        arr = np.array([0.0] * 12)
        arr[-3:] = 100.0  # last 3M positive
        rule = self._rule(RuleType.ANY_OF_3, p1=3, p2=6, p3=12)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is True

    def test_all_of_3_one_negative(self):
        arr = np.array([100.0] * 12)
        arr[-3:] = -100.0  # last 3M negative
        rule = self._rule(RuleType.ALL_OF_3, p1=3, p2=6, p3=12)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is False

    def test_threshold_annual_pass(self):
        arr = np.full(6, 200.0)  # 200×6=1200, ann=2400 ≥ 0.5 × 1000 = 500
        rule = self._rule(RuleType.THRESHOLD_ANNUAL, p1=6)
        assert evaluate_rule(rule, arr, 0, 1000.0, 0.5) is True

    def test_recovery_pass(self):
        arr = np.array([-200.0, 100.0])
        rule = self._rule(RuleType.RECOVERY, p1=1, p2=1)
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is True

    def test_require_oos_profitable_blocks_failing_oos(self):
        # OOS portion (from idx 1) is negative → require_oos_profitable blocks
        arr = np.array([1000.0, -100.0, -100.0])
        rule = self._rule(RuleType.SIMPLE_POSITIVE, p1=1, oos=True)
        # last 1M is negative → simple_positive fails anyway, but oos check also fails
        assert evaluate_rule(rule, arr, oos_start_idx=1, expected_annual=0.0, efficiency_ratio=0.5) is False

    def test_require_oos_profitable_allows_passing_oos(self):
        # OOS portion (from idx 1) is strongly positive; last 1M also positive
        arr = np.array([0.0, 100.0, 200.0])
        rule = self._rule(RuleType.SIMPLE_POSITIVE, p1=1, oos=True)
        assert evaluate_rule(rule, arr, oos_start_idx=1, expected_annual=0.0, efficiency_ratio=0.5) is True

    def test_unknown_rule_type_returns_false(self):
        """Any unrecognised rule type should fail safe (return False)."""
        arr = np.array([100.0])
        rule = EligibilityRule(id=999, label="fake", rule_type=RuleType.BASELINE)
        # Monkey-patch to an invalid value for coverage
        rule.rule_type = "NOT_A_RULE_TYPE"  # type: ignore[assignment]
        assert evaluate_rule(rule, arr, 0, 0.0, 0.5) is False


# ── build_rule_catalogue ──────────────────────────────────────────────────────

class TestBuildRuleCatalogue:
    def setup_method(self):
        self.rules = build_rule_catalogue()

    def test_produces_rules(self):
        assert len(self.rules) > 0

    def test_ids_are_unique(self):
        ids = [r.id for r in self.rules]
        assert len(ids) == len(set(ids))

    def test_ids_are_sequential_from_zero(self):
        ids = sorted(r.id for r in self.rules)
        assert ids == list(range(len(self.rules)))

    def test_all_section_types_present(self):
        types_present = {r.rule_type for r in self.rules}
        expected_types = {
            RuleType.BASELINE,
            RuleType.OOS_PROFITABLE,
            RuleType.SIMPLE_POSITIVE,
            RuleType.SIMPLE_NEGATIVE,
            RuleType.CONSECUTIVE,
            RuleType.COUNT_POSITIVE,
            RuleType.MOMENTUM,
            RuleType.ACCELERATION,
            RuleType.AND_COMBO,
            RuleType.ANY_OF_3,
            RuleType.ALL_OF_3,
            RuleType.THRESHOLD_ANNUAL,
            RuleType.RECOVERY,
        }
        assert expected_types.issubset(types_present)

    def test_oos_variants_present(self):
        oos_rules = [r for r in self.rules if r.require_oos_profitable]
        assert len(oos_rules) >= 10  # each section has OOS variants

    def test_labels_are_non_empty_strings(self):
        assert all(isinstance(r.label, str) and r.label for r in self.rules)

    def test_baseline_rule_is_first(self):
        assert self.rules[0].rule_type == RuleType.BASELINE

    def test_at_least_70_rules(self):
        """Catalogue should have all sections A-H = 70 rules."""
        assert len(self.rules) >= 70
