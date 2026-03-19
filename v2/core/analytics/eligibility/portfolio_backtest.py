"""
Walk-forward portfolio construction backtest — v2 extension of the VBA module.

At each month m:
  1. Determine eligible strategies
  2. Apply selected eligibility rule → passing set
  3. Optionally rank by metric and cap to top-N
  4. Portfolio PnL for month m+1 = mean of selected strategies' next-month PnL
     (equal weight) or weighted by contracts

Returns a dict mapping rule_label → PortfolioBacktestResult,
always including "Baseline (All Eligible)" for comparison.
"""

from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from core.analytics.eligibility.rule_backtest import _is_eligible, _oos_start_idx
from core.analytics.eligibility.rules import (
    EligibilityRule,
    build_rule_catalogue,
    evaluate_rule,
)
from core.config import EligibilityConfig


# ── Config & Result types ─────────────────────────────────────────────────────

@dataclass
class PortfolioBacktestConfig:
    """User configuration for portfolio construction backtest."""
    rule_ids: list[int]                     # Which rules to test (by rule.id)
    max_strategies: int | None = None       # Cap on count (None = all passing)
    ranking_metric: str = "oos_pnl"        # oos_pnl | momentum_3m | momentum_6m | expected_return
    weighting: str = "equal"               # equal | by_contracts
    include_baseline: bool = True


@dataclass
class PortfolioBacktestResult:
    """Results for one rule's portfolio construction backtest."""
    label: str
    monthly_pnl: pd.Series                 # index=month-end, values=portfolio PnL
    monthly_strategy_count: pd.Series      # strategies selected each month
    monthly_selected: pd.DataFrame         # bool mask: rows=months, cols=strategies
    equity_curve: pd.Series               # cumulative sum of monthly_pnl
    win_rate: float                        # % of months with positive PnL
    avg_monthly_pnl: float                 # mean monthly PnL
    max_drawdown: float                    # dollar max drawdown on equity curve
    sharpe_ratio: float                    # annualised Sharpe (monthly data)
    vs_baseline_pct: float                 # % improvement vs baseline avg monthly PnL


# ── Ranking helpers ───────────────────────────────────────────────────────────

def _rank_strategies(
    passing: list[str],
    ranking_metric: str,
    monthly_pnl_up_to_m: pd.DataFrame,
    summary: pd.DataFrame,
) -> list[str]:
    """
    Rank passing strategies by the chosen metric (descending — best first).
    Strategies without data fall to the end.
    """
    scores: dict[str, float] = {}
    for s in passing:
        if ranking_metric == "expected_return":
            scores[s] = float(summary.loc[s].get("expected_annual_profit", 0) or 0) if s in summary.index else 0.0
        elif ranking_metric == "momentum_3m":
            col = monthly_pnl_up_to_m[s] if s in monthly_pnl_up_to_m.columns else None
            scores[s] = float(col.iloc[-3:].sum()) if col is not None and len(col) >= 3 else 0.0
        elif ranking_metric == "momentum_6m":
            col = monthly_pnl_up_to_m[s] if s in monthly_pnl_up_to_m.columns else None
            scores[s] = float(col.iloc[-6:].sum()) if col is not None and len(col) >= 6 else 0.0
        else:  # oos_pnl (default)
            col = monthly_pnl_up_to_m[s] if s in monthly_pnl_up_to_m.columns else None
            oos_begin = summary.loc[s].get("oos_begin") if s in summary.index else None
            if col is not None and oos_begin is not None:
                oos_ts = pd.Timestamp(oos_begin)
                oos_vals = col.loc[col.index >= oos_ts]
                scores[s] = float(oos_vals.sum()) if not oos_vals.empty else 0.0
            else:
                scores[s] = 0.0

    return sorted(passing, key=lambda s: scores.get(s, 0.0), reverse=True)


# ── Result computation helpers ────────────────────────────────────────────────

def _compute_result(
    label: str,
    monthly_pnl_series: pd.Series,
    monthly_count: pd.Series,
    monthly_selected: pd.DataFrame,
    baseline_avg: float,
) -> PortfolioBacktestResult:
    equity = monthly_pnl_series.cumsum()

    n = len(monthly_pnl_series)
    win_rate = float((monthly_pnl_series > 0).mean()) if n > 0 else 0.0
    avg_monthly = float(monthly_pnl_series.mean()) if n > 0 else 0.0

    # Max drawdown (dollar)
    peak = equity.cummax()
    dd = peak - equity
    max_dd = float(dd.max()) if len(dd) > 0 else 0.0

    # Monthly Sharpe (annualised: × √12)
    if n > 1:
        std = float(monthly_pnl_series.std())
        sharpe = (avg_monthly / std * np.sqrt(12)) if std > 1e-9 else 0.0
    else:
        sharpe = 0.0

    vs_base = (
        (avg_monthly / baseline_avg - 1) * 100
        if abs(baseline_avg) > 1e-4 else 0.0
    )

    return PortfolioBacktestResult(
        label=label,
        monthly_pnl=monthly_pnl_series,
        monthly_strategy_count=monthly_count,
        monthly_selected=monthly_selected,
        equity_curve=equity,
        win_rate=win_rate,
        avg_monthly_pnl=avg_monthly,
        max_drawdown=max_dd,
        sharpe_ratio=sharpe,
        vs_baseline_pct=vs_base,
    )


# ── Main function ─────────────────────────────────────────────────────────────

def run_portfolio_backtest(
    daily_pnl: pd.DataFrame,
    summary: pd.DataFrame,
    config: EligibilityConfig,
    backtest_config: PortfolioBacktestConfig,
    rules: list[EligibilityRule] | None = None,
) -> dict[str, PortfolioBacktestResult]:
    """
    Walk-forward portfolio construction backtest.

    Args:
        daily_pnl:        Daily PnL (DatetimeIndex × strategy columns).
        summary:          Per-strategy summary metrics; must include 'status'.
        config:           EligibilityConfig.
        backtest_config:  Which rules, max_strategies, ranking, weighting.
        rules:            Full rule catalogue. Uses build_rule_catalogue() if None.

    Returns:
        dict mapping rule_label → PortfolioBacktestResult.
        Always includes "Baseline (All Eligible)" when include_baseline=True.
    """
    if rules is None:
        rules = build_rule_catalogue()

    rule_map: dict[int, EligibilityRule] = {r.id: r for r in rules}

    monthly = daily_pnl.resample("ME").sum()
    months = monthly.index
    n_months = len(months)
    strats = list(monthly.columns)

    if n_months < 2:
        return {}

    contracts_map: dict[str, int] = {}
    for s in strats:
        if s in summary.index:
            contracts_map[s] = int(summary.loc[s].get("contracts", 1) or 1)
        else:
            contracts_map[s] = 1

    exp_annual_map = {
        s: float(summary.loc[s].get("expected_annual_profit", 0) or 0)
        if s in summary.index else 0.0
        for s in strats
    }
    oos_idx_map = {s: _oos_start_idx(s, summary, months) for s in strats}
    eff_ratio = config.efficiency_ratio

    # Rules to evaluate (plus baseline)
    selected_rule_ids = set(backtest_config.rule_ids)
    baseline_rule = next((r for r in rules if r.id == 0), rules[0])  # rule_id=0 is Baseline

    eval_rules: list[EligibilityRule] = []
    if backtest_config.include_baseline:
        eval_rules.append(baseline_rule)
    for rid in backtest_config.rule_ids:
        r = rule_map.get(rid)
        if r and r.id != 0:
            eval_rules.append(r)

    # Per-rule accumulators
    # monthly_pnl_acc[rule_label][month_ts] = portfolio_pnl
    pnl_acc: dict[str, dict] = {r.label: {} for r in eval_rules}
    count_acc: dict[str, dict] = {r.label: {} for r in eval_rules}
    selected_acc: dict[str, dict] = {r.label: {} for r in eval_rules}  # month → list[str]

    # Walk-forward (portfolio PnL comes from month m+1 onwards)
    for m_idx in range(n_months - 1):
        month_end = months[m_idx]
        next_month = months[m_idx + 1]

        eligible = [
            s for s in strats
            if _is_eligible(s, month_end, summary, config)
        ]
        if not eligible:
            continue

        for rule in eval_rules:
            # Evaluate rule on each eligible strategy
            passing: list[str] = []
            for s in eligible:
                pnl_slice = monthly[s].values[: m_idx + 1]
                oos_i = oos_idx_map[s]
                if evaluate_rule(rule, pnl_slice, oos_i, exp_annual_map[s], eff_ratio):
                    passing.append(s)

            if not passing:
                pnl_acc[rule.label][next_month]      = 0.0
                count_acc[rule.label][next_month]    = 0
                selected_acc[rule.label][next_month] = []
                continue

            # Optionally rank and cap
            if backtest_config.max_strategies is not None:
                monthly_slice = monthly.iloc[: m_idx + 1]
                passing = _rank_strategies(
                    passing, backtest_config.ranking_metric, monthly_slice, summary
                )
                passing = passing[: backtest_config.max_strategies]

            # Portfolio PnL for next month
            if backtest_config.weighting == "by_contracts":
                total_contracts = sum(contracts_map.get(s, 1) for s in passing)
                portfolio_pnl = sum(
                    monthly.loc[next_month, s] * contracts_map.get(s, 1)
                    for s in passing
                ) / (total_contracts or 1)
            else:  # equal weight
                portfolio_pnl = float(monthly.loc[next_month, passing].mean())

            pnl_acc[rule.label][next_month]      = portfolio_pnl
            count_acc[rule.label][next_month]    = len(passing)
            selected_acc[rule.label][next_month] = passing

    # ── Compute baseline avg for vs_base calculation ──────────────────────────
    baseline_label = baseline_rule.label
    baseline_pnl_vals = list(pnl_acc.get(baseline_label, {}).values())
    baseline_avg = float(np.mean(baseline_pnl_vals)) if baseline_pnl_vals else 0.0

    # ── Build results ─────────────────────────────────────────────────────────
    results: dict[str, PortfolioBacktestResult] = {}
    for rule in eval_rules:
        lbl = rule.label
        pnl_dict = pnl_acc[lbl]
        cnt_dict = count_acc[lbl]
        sel_dict = selected_acc[lbl]

        if not pnl_dict:
            continue

        pnl_series = pd.Series(pnl_dict, name=lbl).sort_index()
        count_series = pd.Series(cnt_dict, name="count").sort_index()

        # Build selected boolean DataFrame
        all_months = pnl_series.index
        sel_df = pd.DataFrame(False, index=all_months, columns=strats, dtype=bool)
        for m, sel_list in sel_dict.items():
            if m in sel_df.index:
                for s in sel_list:
                    if s in sel_df.columns:
                        sel_df.loc[m, s] = True

        results[lbl] = _compute_result(
            label=lbl,
            monthly_pnl_series=pnl_series,
            monthly_count=count_series,
            monthly_selected=sel_df,
            baseline_avg=baseline_avg,
        )

    return results
