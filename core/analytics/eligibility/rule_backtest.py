"""
Walk-forward eligibility rule statistics — mirrors U_BackTest_Eligibility.bas.

For each month m in history:
  1. Determine which strategies were eligible (status, OOS days, DD cap)
  2. Evaluate each rule → passing set
  3. For each horizon h (1–12): record actual next-h-months PnL per strategy

Aggregate across all months → N / Win% / $/Month / vs-Baseline for each
(rule × horizon) cell, returned as a flat DataFrame.

Performance target: ~3s for 70 rules on a 5-year dataset with 20 strategies.
"""

from __future__ import annotations

from datetime import date

import numpy as np
import pandas as pd

from core.analytics.eligibility.rules import (
    EligibilityRule,
    build_rule_catalogue,
    evaluate_rule,
)
from core.config import EligibilityConfig


# ── Eligibility helpers ───────────────────────────────────────────────────────

def _is_eligible(
    strat: str,
    month_end: pd.Timestamp,
    summary: pd.DataFrame,
    config: EligibilityConfig,
    status_col: str = "status",
) -> bool:
    """
    Return True if strategy passes base eligibility at the given month-end date.

    Checks:
      1. Status in config.status_include
      2. Days of OOS data >= config.days_threshold_oos
      3. OOS max drawdown <= oos_dd_vs_is_cap × IS max drawdown (if cap > 0)
    """
    if strat not in summary.index:
        return False

    row = summary.loc[strat]

    # ── 1. Status ─────────────────────────────────────────────────────────────
    status = str(row.get(status_col, ""))
    if status not in config.status_include:
        return False

    # ── 2. OOS days threshold ─────────────────────────────────────────────────
    # Determine eligibility date based on date_type setting
    if config.date_type == "Incubation Pass Date":
        elig_date = row.get("incubation_date")
        if elig_date is None:
            return False
        elig_ts = pd.Timestamp(elig_date)
    else:  # OOS Start Date
        oos_begin = row.get("oos_begin")
        if oos_begin is None:
            return False
        elig_ts = pd.Timestamp(oos_begin)

    days_in_oos = (month_end - elig_ts).days
    if days_in_oos < config.days_threshold_oos:
        return False

    # ── 3. OOS drawdown cap ───────────────────────────────────────────────────
    if config.oos_dd_vs_is_cap > 0:
        is_dd  = abs(float(row.get("max_drawdown_is", 0) or 0))
        oos_dd = abs(float(row.get("max_oos_drawdown", 0) or 0))
        if is_dd > 1e-4 and oos_dd > config.oos_dd_vs_is_cap * is_dd:
            return False

    return True


def _oos_start_idx(strat: str, summary: pd.DataFrame, months_index: pd.DatetimeIndex) -> int:
    """
    Return the integer index in months_index where the strategy's OOS begins.
    Returns 0 if OOS start is before the series or not found.
    """
    if strat not in summary.index:
        return 0
    oos_begin = summary.loc[strat].get("oos_begin")
    if oos_begin is None:
        return 0
    oos_ts = pd.Timestamp(oos_begin)
    idx = months_index.searchsorted(oos_ts, side="left")
    return int(min(idx, len(months_index) - 1))


# ── Main function ─────────────────────────────────────────────────────────────

def run_rule_backtest(
    daily_pnl: pd.DataFrame,
    summary: pd.DataFrame,
    config: EligibilityConfig,
    rules: list[EligibilityRule] | None = None,
    max_horizon: int = 12,
) -> pd.DataFrame:
    """
    Walk-forward rule statistics.

    Args:
        daily_pnl:    Daily PnL (DatetimeIndex × strategy columns).
        summary:      Per-strategy summary metrics (from compute_summary()).
                      Must include a 'status' column; add it before calling:
                          summary['status'] = [status_map[n] for n in summary.index]
        config:       EligibilityConfig.
        rules:        Rule catalogue. Uses build_rule_catalogue() if None.
        max_horizon:  Max lookahead horizon in months.

    Returns:
        DataFrame with one row per rule. Columns:
          rule_id, label, rule_type,
          N_{h}, win_count_{h}, win_pct_{h}, avg_pnl_{h}, vs_base_{h}
          for h in 1..max_horizon.
    """
    if rules is None:
        rules = build_rule_catalogue()

    # ── Pre-compute monthly PnL ───────────────────────────────────────────────
    monthly = daily_pnl.resample("ME").sum()
    months = monthly.index  # DatetimeIndex of month-end dates
    n_months = len(months)
    strats = list(monthly.columns)

    if n_months < 2:
        return _empty_result(rules, max_horizon)

    # ── Pre-compute OOS start indices ─────────────────────────────────────────
    oos_idx_map: dict[str, int] = {
        s: _oos_start_idx(s, summary, months) for s in strats
    }

    # ── Pre-compute expected annual profit and efficiency ratio ───────────────
    exp_annual_map: dict[str, float] = {}
    for s in strats:
        if s in summary.index:
            exp_annual_map[s] = float(summary.loc[s].get("expected_annual_profit", 0) or 0)
        else:
            exp_annual_map[s] = 0.0

    eff_ratio = config.efficiency_ratio

    # ── Accumulate stats per rule ─────────────────────────────────────────────
    # stats[rule_idx][horizon] = {"n": int, "wins": int, "sum_pnl": float}
    n_rules = len(rules)
    h_range = list(range(1, max_horizon + 1))

    # Use list of dicts for speed (avoid DataFrame overhead in inner loop)
    stats: list[dict] = [
        {h: {"n": 0, "wins": 0, "sum_pnl": 0.0} for h in h_range}
        for _ in rules
    ]
    base_stats: dict[int, dict] = {h: {"n": 0, "wins": 0, "sum_pnl": 0.0} for h in h_range}

    # Walk-forward: for each evaluation month m (need m + max_horizon months ahead)
    last_eval_month = n_months - 1 - max_horizon

    for m_idx in range(n_months):
        month_end = months[m_idx]

        # Determine eligible strategies at this month
        eligible = [
            s for s in strats
            if _is_eligible(s, month_end, summary, config)
        ]
        if not eligible:
            continue

        # For each eligible strategy, build monthly PnL slice up to m_idx (inclusive)
        # and evaluate all rules
        for rule_idx, rule in enumerate(rules):
            passing: list[str] = []
            for s in eligible:
                pnl_slice = monthly[s].values[: m_idx + 1]
                oos_i = oos_idx_map[s]
                if evaluate_rule(
                    rule, pnl_slice, oos_i,
                    exp_annual_map[s], eff_ratio
                ):
                    passing.append(s)

            if not passing:
                continue

            # Accumulate stats for each horizon
            for h in h_range:
                future_end = m_idx + h
                if future_end >= n_months:
                    continue
                # Sum of each passing strategy's PnL over months m+1 to m+h
                for s in passing:
                    future_pnl = float(monthly[s].values[m_idx + 1: future_end + 1].sum())
                    st = stats[rule_idx][h]
                    st["n"] += 1
                    if future_pnl > 0:
                        st["wins"] += 1
                    st["sum_pnl"] += future_pnl

        # Baseline: all eligible strategies (for vs_base computation)
        for h in h_range:
            future_end = m_idx + h
            if future_end >= n_months:
                continue
            for s in eligible:
                future_pnl = float(monthly[s].values[m_idx + 1: future_end + 1].sum())
                bs = base_stats[h]
                bs["n"] += 1
                if future_pnl > 0:
                    bs["wins"] += 1
                bs["sum_pnl"] += future_pnl

    # ── Build results DataFrame ───────────────────────────────────────────────
    rows = []
    for rule_idx, rule in enumerate(rules):
        row: dict = {
            "rule_id":   rule.id,
            "label":     rule.label,
            "rule_type": rule.rule_type.value,
        }
        for h in h_range:
            st = stats[rule_idx][h]
            n = st["n"]
            bs = base_stats[h]
            base_avg = bs["sum_pnl"] / bs["n"] if bs["n"] > 0 else 0.0
            rule_avg = st["sum_pnl"] / n if n > 0 else 0.0
            row[f"N_{h}"]        = n
            row[f"win_pct_{h}"]  = round(st["wins"] / n * 100, 1) if n > 0 else 0.0
            row[f"avg_pnl_{h}"]  = round(rule_avg, 0)
            row[f"vs_base_{h}"]  = (
                round((rule_avg / base_avg - 1) * 100, 1)
                if abs(base_avg) > 1e-4 else 0.0
            )
        rows.append(row)

    return pd.DataFrame(rows)


def _empty_result(rules: list[EligibilityRule], max_horizon: int) -> pd.DataFrame:
    """Return an empty-stats DataFrame with correct columns."""
    rows = []
    h_range = list(range(1, max_horizon + 1))
    for rule in rules:
        row: dict = {
            "rule_id":   rule.id,
            "label":     rule.label,
            "rule_type": rule.rule_type.value,
        }
        for h in h_range:
            row[f"N_{h}"] = 0
            row[f"win_pct_{h}"] = 0.0
            row[f"avg_pnl_{h}"] = 0.0
            row[f"vs_base_{h}"] = 0.0
        rows.append(row)
    return pd.DataFrame(rows)
