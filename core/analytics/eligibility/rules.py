"""
Eligibility rule definitions and evaluation — mirrors U_BackTest_Eligibility.bas.

Sections A–H produce 70 parameterised rules covering all types in the VBA module:
  A  Baseline (2)        — All Eligible, OOS Profitable
  B  Simple Period (20)  — Last 1/3/6/9/12M >0 or <0, ±OOS-profitable variant
  C  Consecutive (8)     — Last 3/4/5/6 months all positive, ±OOS variant
  D  Count-Based (10)    — K-of-N positive, ±OOS variant
  E  Momentum (8)        — Recent > prior / acceleration (annualised), ±OOS
  F  Combination (10)    — AND combos, ANY_OF_3, ALL_OF_3, ±OOS variant
  G  Threshold (6)       — Annualised return ≥ efficiency × expected, ±OOS
  H  Recovery (6)        — Positive after negative window, ±OOS variant

All rule functions receive pre-sliced monthly PnL arrays (up to and including
the evaluation month) to prevent look-ahead bias.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum

import numpy as np


# ── Types ─────────────────────────────────────────────────────────────────────

class RuleType(Enum):
    BASELINE         = "baseline"
    OOS_PROFITABLE   = "oos_profitable"
    SIMPLE_POSITIVE  = "simple_positive"
    SIMPLE_NEGATIVE  = "simple_negative"
    CONSECUTIVE      = "consecutive"
    COUNT_POSITIVE   = "count_positive"
    MOMENTUM         = "momentum"
    ACCELERATION     = "acceleration"
    AND_COMBO        = "and_combo"
    ANY_OF_3         = "any_of_3"
    ALL_OF_3         = "all_of_3"
    THRESHOLD_ANNUAL = "threshold_annual"
    RECOVERY         = "recovery"


@dataclass
class EligibilityRule:
    """Single eligibility rule definition (mirrors VBA EligibilityRule type)."""
    id: int
    label: str
    rule_type: RuleType
    param1: float = 0.0   # Primary param   (e.g. N months)
    param2: float = 0.0   # Secondary param  (e.g. K for K-of-N)
    param3: float = 0.0   # Tertiary param   (e.g. 3rd window)
    require_oos_profitable: bool = False  # AND: cumulative OOS PnL > 0


# ── Rule evaluation ───────────────────────────────────────────────────────────

def evaluate_rule(
    rule: EligibilityRule,
    monthly_pnl: np.ndarray,
    oos_start_idx: int,
    expected_annual: float,
    efficiency_ratio: float,
) -> bool:
    """
    Evaluate one rule for one strategy at one point in time.

    Args:
        rule:             Rule to evaluate.
        monthly_pnl:      Monthly PnL array sliced UP TO AND INCLUDING the
                          current evaluation month (no future data).
        oos_start_idx:    Index in monthly_pnl where OOS begins.
        expected_annual:  Expected annual profit from walkforward CSV.
        efficiency_ratio: Threshold multiplier from EligibilityConfig.

    Returns:
        True if strategy passes the rule, False otherwise.
    """
    # Guard: OOS profitable prerequisite
    if rule.require_oos_profitable:
        oos = monthly_pnl[oos_start_idx:]
        if len(oos) == 0 or float(oos.sum()) <= 0.0:
            return False

    rt = rule.rule_type

    if rt == RuleType.BASELINE:
        return True

    if rt == RuleType.OOS_PROFITABLE:
        oos = monthly_pnl[oos_start_idx:]
        return len(oos) > 0 and float(oos.sum()) > 0.0

    if rt == RuleType.SIMPLE_POSITIVE:
        return _last_n_sum(monthly_pnl, int(rule.param1)) > 0.0

    if rt == RuleType.SIMPLE_NEGATIVE:
        return _last_n_sum(monthly_pnl, int(rule.param1)) < 0.0

    if rt == RuleType.CONSECUTIVE:
        return _consecutive_positive(monthly_pnl, int(rule.param1))

    if rt == RuleType.COUNT_POSITIVE:
        return _count_positive_ge(monthly_pnl, int(rule.param1), int(rule.param2))

    if rt == RuleType.MOMENTUM:
        return _momentum_positive(monthly_pnl, int(rule.param1), int(rule.param2))

    if rt == RuleType.ACCELERATION:
        return _acceleration_positive(monthly_pnl, int(rule.param1), int(rule.param2))

    if rt == RuleType.AND_COMBO:
        m1, m2 = int(rule.param1), int(rule.param2)
        return _last_n_sum(monthly_pnl, m1) > 0.0 and _last_n_sum(monthly_pnl, m2) > 0.0

    if rt == RuleType.ANY_OF_3:
        m1, m2, m3 = int(rule.param1), int(rule.param2), int(rule.param3)
        return any(_last_n_sum(monthly_pnl, m) > 0.0 for m in (m1, m2, m3))

    if rt == RuleType.ALL_OF_3:
        m1, m2, m3 = int(rule.param1), int(rule.param2), int(rule.param3)
        return all(_last_n_sum(monthly_pnl, m) > 0.0 for m in (m1, m2, m3))

    if rt == RuleType.THRESHOLD_ANNUAL:
        n = int(rule.param1)
        return _threshold_annual(monthly_pnl, n, expected_annual, efficiency_ratio)

    if rt == RuleType.RECOVERY:
        return _recovery(monthly_pnl, int(rule.param1), int(rule.param2))

    return False  # unknown rule type — fail safe


# ── Rule primitive helpers ────────────────────────────────────────────────────

def _last_n_sum(arr: np.ndarray, n: int) -> float:
    """Sum of the last n values. Returns 0 if array has fewer than n elements."""
    if len(arr) < n or n <= 0:
        return 0.0
    return float(arr[-n:].sum())


def _consecutive_positive(arr: np.ndarray, n: int) -> bool:
    """True if the last n monthly values are ALL strictly positive."""
    if len(arr) < n or n <= 0:
        return False
    return bool((arr[-n:] > 0).all())


def _count_positive_ge(arr: np.ndarray, k: int, n: int) -> bool:
    """True if at least k of the last n monthly values are strictly positive."""
    if len(arr) < n or n <= 0:
        return False
    return int((arr[-n:] > 0).sum()) >= k


def _momentum_positive(arr: np.ndarray, short_n: int, long_n: int) -> bool:
    """
    True if sum of last short_n months > sum of prior (long_n - short_n) months.
    Requires at least long_n months of history.
    """
    if len(arr) < long_n or short_n >= long_n:
        return False
    recent = float(arr[-short_n:].sum())
    prior  = float(arr[-(long_n):-short_n].sum())
    return recent > prior


def _acceleration_positive(arr: np.ndarray, short_n: int, long_n: int) -> bool:
    """
    True if annualised return over short_n > annualised return over long_n.
    Mirrors VBA acceleration check.
    """
    if len(arr) < long_n or short_n <= 0 or long_n <= 0:
        return False
    ann_short = float(arr[-short_n:].sum()) * (12.0 / short_n)
    ann_long  = float(arr[-long_n:].sum())  * (12.0 / long_n)
    return ann_short > ann_long


def _threshold_annual(
    arr: np.ndarray,
    n: int,
    expected_annual: float,
    efficiency_ratio: float,
) -> bool:
    """
    True if annualised return over last n months ≥ efficiency_ratio × expected_annual.
    Returns False if expected_annual <= 0 (no comparison possible).
    """
    if len(arr) < n or n <= 0 or expected_annual <= 0.0:
        return False
    ann = float(arr[-n:].sum()) * (12.0 / n)
    return ann >= efficiency_ratio * expected_annual


def _recovery(arr: np.ndarray, pos_n: int, neg_n: int) -> bool:
    """
    True if last pos_n months sum > 0 AND prior neg_n months sum < 0.
    Requires at least pos_n + neg_n months of history.
    """
    total = pos_n + neg_n
    if len(arr) < total or pos_n <= 0 or neg_n <= 0:
        return False
    recent = float(arr[-pos_n:].sum())
    prior  = float(arr[-(total):-pos_n].sum())
    return recent > 0.0 and prior < 0.0


# ── Rule catalogue ────────────────────────────────────────────────────────────

def build_rule_catalogue() -> list[EligibilityRule]:
    """
    Build the full 70-rule catalogue (mirrors VBA sections A–H).
    Rules are returned in ascending id order.
    """
    rules: list[EligibilityRule] = []
    rid = 0

    def add(
        label: str,
        rule_type: RuleType,
        p1: float = 0.0,
        p2: float = 0.0,
        p3: float = 0.0,
        oos: bool = False,
    ) -> None:
        nonlocal rid
        rules.append(EligibilityRule(
            id=rid, label=label, rule_type=rule_type,
            param1=p1, param2=p2, param3=p3,
            require_oos_profitable=oos,
        ))
        rid += 1

    # ── A: Baseline ───────────────────────────────────────────────────────────
    add("Baseline (All Eligible)",        RuleType.BASELINE)
    add("OOS Profitable",                 RuleType.OOS_PROFITABLE)

    # ── B: Simple Period ──────────────────────────────────────────────────────
    for m, mstr in [(1, "1M"), (3, "3M"), (6, "6M"), (9, "9M"), (12, "12M")]:
        add(f"Last {mstr} > 0",           RuleType.SIMPLE_POSITIVE, p1=m)
        add(f"Last {mstr} < 0",           RuleType.SIMPLE_NEGATIVE, p1=m)
        add(f"Last {mstr} > 0 + OOS+",   RuleType.SIMPLE_POSITIVE, p1=m, oos=True)
        add(f"Last {mstr} < 0 + OOS+",   RuleType.SIMPLE_NEGATIVE, p1=m, oos=True)

    # ── C: Consecutive ────────────────────────────────────────────────────────
    for n in [3, 4, 5, 6]:
        add(f"Consecutive {n}M+",         RuleType.CONSECUTIVE, p1=n)
        add(f"Consecutive {n}M+ + OOS+",  RuleType.CONSECUTIVE, p1=n, oos=True)

    # ── D: Count-Based ────────────────────────────────────────────────────────
    for k, n in [(2, 3), (3, 4), (3, 5), (4, 6), (5, 6)]:
        add(f"{k} of {n}M+",              RuleType.COUNT_POSITIVE, p1=k, p2=n)
        add(f"{k} of {n}M+ + OOS+",       RuleType.COUNT_POSITIVE, p1=k, p2=n, oos=True)

    # ── E: Momentum / Acceleration ────────────────────────────────────────────
    for short, long_ in [(3, 6), (6, 12)]:
        prior = long_ - short
        add(f"Momentum {short}M > {prior}M prior",     RuleType.MOMENTUM,      p1=short, p2=long_)
        add(f"Momentum {short}M > {prior}M + OOS+",    RuleType.MOMENTUM,      p1=short, p2=long_, oos=True)
        add(f"Accel {short}M ann > {long_}M ann",      RuleType.ACCELERATION,  p1=short, p2=long_)
        add(f"Accel {short}M ann > {long_}M + OOS+",   RuleType.ACCELERATION,  p1=short, p2=long_, oos=True)

    # ── F: Combination ────────────────────────────────────────────────────────
    for m1, m2 in [(3, 6), (6, 12), (3, 9)]:
        add(f"Last {m1}M > 0 AND Last {m2}M > 0",      RuleType.AND_COMBO, p1=m1, p2=m2)
        add(f"Last {m1}M+{m2}M > 0 + OOS+",             RuleType.AND_COMBO, p1=m1, p2=m2, oos=True)
    add("Any of 3M/6M/12M > 0",                         RuleType.ANY_OF_3, p1=3, p2=6, p3=12)
    add("Any of 3M/6M/12M > 0 + OOS+",                  RuleType.ANY_OF_3, p1=3, p2=6, p3=12, oos=True)
    add("All of 3M/6M/12M > 0",                         RuleType.ALL_OF_3, p1=3, p2=6, p3=12)
    add("All of 3M/6M/12M > 0 + OOS+",                  RuleType.ALL_OF_3, p1=3, p2=6, p3=12, oos=True)

    # ── G: Threshold ──────────────────────────────────────────────────────────
    for n in [6, 12, 3]:
        add(f"Ann Return {n}M ≥ Eff×Exp",                RuleType.THRESHOLD_ANNUAL, p1=n)
        add(f"Ann Return {n}M ≥ Eff×Exp + OOS+",         RuleType.THRESHOLD_ANNUAL, p1=n, oos=True)

    # ── H: Recovery ───────────────────────────────────────────────────────────
    for pos_n, neg_n in [(1, 1), (1, 3), (3, 3)]:
        add(f"Recovery +{pos_n}M after -{neg_n}M",       RuleType.RECOVERY, p1=pos_n, p2=neg_n)
        add(f"Recovery +{pos_n}M after -{neg_n}M + OOS+",RuleType.RECOVERY, p1=pos_n, p2=neg_n, oos=True)

    return rules
