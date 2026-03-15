"""
Portfolio Optimizer — composable workflow engine.

Each ``step_*`` function takes an :class:`OptimizerState` and keyword
parameters, mutates the state in place, and returns it.  A *workflow* is
simply an ordered list of ``(step_fn, kwargs)`` tuples run by
:func:`run_workflow`.

Typical workflow (mirrors the user's documented process):

1. filter_eligibility       — drop strategies that failed eligibility gates
2. filter_excluded_symbols  — drop large/unwanted symbols (e.g. "S")
3. size_contracts           — compute ATR/margin-blended contract counts
4. filter_contract_size     — drop strategies where raw count < 0.65
5. rank                     — sort by chosen metric
6. select_strategies        — greedy selection respecting margin cap
7. adjust_correlations      — remove/reduce correlated pairs
8. adjust_gross_margins     — enforce per-symbol/sector margin caps
9. adjust_drawdowns         — reduce contracts if drawdown limit breached
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field

import pandas as pd


# ── State ─────────────────────────────────────────────────────────────────────

@dataclass
class ExclusionRecord:
    name: str
    step: str
    reason: str


@dataclass
class OptimizerState:
    """
    Mutable snapshot of the optimizer's working set.

    ``candidates`` is the *active* list of strategy dicts (merged config +
    summary metrics).  Each dict must have at minimum a ``"name"`` key.
    ``contracts`` maps strategy name → fractional contract count (multiples
    of ``min_fraction``).  ``equity`` is the current account balance (may be
    updated by an MC capital step).
    """
    candidates: list[dict]
    contracts: dict[str, float]
    equity: float
    excluded: list[ExclusionRecord] = field(default_factory=list)
    log: list[str] = field(default_factory=list)

    # ── Helpers ────────────────────────────────────────────────────────────

    @property
    def active_names(self) -> set[str]:
        return {c["name"] for c in self.candidates}

    def exclude_strategy(self, name: str, step: str, reason: str) -> None:
        self.excluded.append(ExclusionRecord(name=name, step=step, reason=reason))
        self.candidates = [c for c in self.candidates if c["name"] != name]
        self.contracts.pop(name, None)
        self.log.append(f"[{step}] Removed '{name}': {reason}")

    def reduce_contracts(
        self, name: str, new_n: float, step: str, reason: str
    ) -> None:
        old = self.contracts.get(name, 0.0)
        self.contracts[name] = new_n
        self.log.append(
            f"[{step}] Reduced '{name}' {old:.1f} → {new_n:.1f} contracts: {reason}"
        )

    def strategy_margin(
        self,
        name: str,
        symbol: str,
        margins: dict[str, float],
        margin_multiple: float,
    ) -> float:
        """Absolute margin usage for one strategy (contracts × margin × multiple)."""
        n = self.contracts.get(name, 0.0)
        m = margins.get(symbol, 5_000.0) * margin_multiple
        return n * m

    def total_margin_used(
        self, margins: dict[str, float], margin_multiple: float
    ) -> float:
        return sum(
            self.strategy_margin(c["name"], c.get("symbol", ""), margins, margin_multiple)
            for c in self.candidates
        )


# ── Helpers ────────────────────────────────────────────────────────────────────

def _round_to_fraction(value: float, fraction: float = 0.1) -> float:
    """Round *down* to the nearest multiple of ``fraction``."""
    if fraction <= 0:
        return float(value)
    return round(math.floor(value / fraction) * fraction, 10)


# ── Step functions ─────────────────────────────────────────────────────────────

def step_filter_eligibility(
    state: OptimizerState,
    eligible_mask: dict[str, bool],
) -> OptimizerState:
    """Remove strategies that are not eligible per the eligibility engine."""
    before = len(state.candidates)
    for name in list(state.active_names):
        if not eligible_mask.get(name, True):
            state.exclude_strategy(name, "eligibility", "Not eligible")
    removed = before - len(state.candidates)
    state.log.append(
        f"[eligibility] Removed {removed}, {len(state.candidates)} remain"
    )
    return state


def step_filter_excluded_symbols(
    state: OptimizerState,
    excluded_symbols: list[str],
) -> OptimizerState:
    """Remove strategies whose symbol appears in ``excluded_symbols``."""
    if not excluded_symbols:
        state.log.append("[excluded_symbols] No symbols excluded — skipped")
        return state
    excluded_upper = {s.strip().upper() for s in excluded_symbols if s.strip()}
    to_remove = [
        c for c in state.candidates
        if c.get("symbol", "").upper() in excluded_upper
    ]
    for c in to_remove:
        state.exclude_strategy(
            c["name"], "excluded_symbols",
            f"Symbol '{c.get('symbol', '')}' is in exclusion list",
        )
    state.log.append(
        f"[excluded_symbols] Removed {len(to_remove)}, {len(state.candidates)} remain"
    )
    return state


def step_filter_contract_size(
    state: OptimizerState,
    min_threshold: float = 0.65,
) -> OptimizerState:
    """
    Remove strategies where the computed contract count is below
    ``min_threshold``.  A count below this means the contract is too large
    to trade with a meaningful fraction of equity (no micro/mini available).
    """
    to_remove = [
        c for c in state.candidates
        if state.contracts.get(c["name"], 0.0) < min_threshold
    ]
    for c in to_remove:
        raw = state.contracts.get(c["name"], 0.0)
        state.exclude_strategy(
            c["name"], "contract_size",
            f"Contract count {raw:.2f} < minimum threshold {min_threshold}",
        )
    state.log.append(
        f"[contract_size] Removed {len(to_remove)}, {len(state.candidates)} remain"
    )
    return state


def step_rank(
    state: OptimizerState,
    metric: str = "rtd_oos",
    ascending: bool = False,
) -> OptimizerState:
    """Sort candidates by ``metric``.  Strategies with no/None value sort last."""
    nan_val = float("inf") if ascending else float("-inf")

    def _key(c: dict) -> float:
        v = c.get(metric)
        if v is None:
            return nan_val
        try:
            return float(v)
        except (TypeError, ValueError):
            return nan_val

    state.candidates.sort(key=_key, reverse=not ascending)
    top3 = [
        f"{c['name']}={c.get(metric, 'N/A')}"
        for c in state.candidates[:3]
    ]
    state.log.append(
        f"[rank] Sorted by '{metric}' ({'asc' if ascending else 'desc'}). "
        f"Top 3: {', '.join(top3)}"
    )
    return state


def step_size_contracts(
    state: OptimizerState,
    equity: float,
    contract_size_pct: float,
    atr: dict[str, float],
    margins: dict[str, float],
    ratio: float,
    contract_margin_multiple: float = 0.33,
    min_fraction: float = 0.1,
) -> OptimizerState:
    """
    Compute fractional contract counts for all active candidates.

    Formula (mirrors VBA Estimated Vol Contract Sizing)::

        dollar_risk = atr × ratio + (margin × margin_multiple) × (1 - ratio)
        raw_contracts = equity × contract_size_pct / dollar_risk
        contracts = floor(raw_contracts / min_fraction) × min_fraction

    ``ratio=0.5`` and ``contract_margin_multiple=0.33`` replicates the
    user's workflow: "average of 3M ATR and 33% of maintenance margin".

    Unlike :func:`~core.analytics.atr.contract_size_from_atr`, this step
    preserves fractional precision (rounds to ``min_fraction``, not to int).
    """
    new_contracts: dict[str, float] = {}
    for c in state.candidates:
        name = c["name"]
        symbol = c.get("symbol", "")
        atr_val = abs(float(atr.get(name, 0.0)))
        raw_margin = float(margins.get(symbol, 5_000.0))
        margin = abs(raw_margin) * contract_margin_multiple

        effective_risk = atr_val * ratio + margin * (1.0 - ratio)
        if effective_risk <= 0:
            raw_float = 0.0
        else:
            raw_float = (equity * contract_size_pct) / effective_risk
        new_contracts[name] = _round_to_fraction(raw_float, min_fraction)

    state.contracts = new_contracts
    state.equity = equity
    state.log.append(
        f"[size_contracts] equity=${equity:,.0f}, pct={contract_size_pct:.1%}, "
        f"ATR ratio={ratio:.0%}. Sized {len(state.candidates)} strategies."
    )
    return state


def step_select_strategies(
    state: OptimizerState,
    margins: dict[str, float],
    contract_margin_multiple: float,
    max_margin_pct: float = 0.75,
    max_strategies: int = 60,
    per_symbol_first: bool = True,
) -> OptimizerState:
    """
    Greedy strategy selection that respects a total-margin cap.

    When ``per_symbol_first=True`` (the default, matching the user's workflow):

    - **Pass 1**: Add the best-ranked strategy from each unique symbol.
    - **Pass 2**: Fill remaining slots by rank until the margin cap or
      ``max_strategies`` is reached.

    Strategies not selected are excluded with reason "not selected".
    """
    equity = state.equity
    ranked = list(state.candidates)  # already sorted by rank step

    def _margin(c: dict) -> float:
        n = state.contracts.get(c["name"], 0.0)
        m = margins.get(c.get("symbol", ""), 5_000.0) * contract_margin_multiple
        return n * m

    selected: list[dict] = []
    total_m = 0.0

    if per_symbol_first:
        # Pass 1: one per unique symbol
        seen_symbols: set[str] = set()
        pass1: list[dict] = []
        for c in ranked:
            sym = c.get("symbol", "")
            if sym and sym not in seen_symbols:
                seen_symbols.add(sym)
                pass1.append(c)

        for c in pass1:
            m = _margin(c)
            if len(selected) >= max_strategies:
                break
            if total_m + m <= max_margin_pct * equity:
                selected.append(c)
                total_m += m

        # Pass 2: remaining by rank
        selected_names = {c["name"] for c in selected}
        for c in ranked:
            if c["name"] in selected_names:
                continue
            if len(selected) >= max_strategies:
                break
            m = _margin(c)
            if total_m + m <= max_margin_pct * equity:
                selected.append(c)
                total_m += m
    else:
        # Simple greedy by rank
        for c in ranked:
            if len(selected) >= max_strategies:
                break
            m = _margin(c)
            if total_m + m <= max_margin_pct * equity:
                selected.append(c)
                total_m += m

    # Exclude not selected
    selected_names = {c["name"] for c in selected}
    for c in list(state.candidates):
        if c["name"] not in selected_names:
            state.exclude_strategy(c["name"], "select", "Not selected (margin or count limit)")

    state.log.append(
        f"[select] Selected {len(state.candidates)} strategies, "
        f"total margin={total_m:,.0f} ({total_m/equity:.1%} of equity)"
    )
    return state


def step_adjust_correlations(
    state: OptimizerState,
    corr_matrix: "pd.DataFrame | None",
    max_corr: float = 0.70,
    max_neg_corr: float = 0.50,
) -> OptimizerState:
    """
    For each pair of strategies that violates a correlation threshold, remove
    the *lower-ranked* one (i.e. later in ``state.candidates``).

    - ``|corr| > max_corr``     → too positively correlated
    - ``corr < -max_neg_corr``  → too negatively correlated
    """
    if corr_matrix is None or corr_matrix.empty:
        state.log.append("[correlations] No correlation matrix — skipped")
        return state

    active = [c["name"] for c in state.candidates]
    avail = [n for n in active if n in corr_matrix.index and n in corr_matrix.columns]
    if len(avail) < 2:
        return state

    mat = corr_matrix.loc[avail, avail]
    removed: set[str] = set()

    for i, n1 in enumerate(avail):
        if n1 in removed:
            continue
        for j, n2 in enumerate(avail):
            if j <= i or n2 in removed:
                continue
            try:
                corr = float(mat.loc[n1, n2])
            except (KeyError, TypeError, ValueError):
                continue

            reason = ""
            if corr > max_corr:
                reason = f"Corr with '{n1}' = {corr:.2f} > {max_corr}"
            elif corr < -max_neg_corr:
                reason = f"Neg-corr with '{n1}' = {corr:.2f} < -{max_neg_corr}"

            if reason:
                # Remove the lower-ranked (later index in active) strategy
                to_remove = n2  # n2 has higher index in avail list → lower rank
                state.exclude_strategy(to_remove, "correlations", reason)
                removed.add(to_remove)

    state.log.append(
        f"[correlations] Removed {len(removed)}, {len(state.candidates)} remain"
    )
    return state


def step_adjust_gross_margins(
    state: OptimizerState,
    margins: dict[str, float],
    contract_margin_multiple: float,
    equity: float,
    max_single_pct: float = 0.125,
    max_sector_pct: float = 0.25,
) -> OptimizerState:
    """
    Enforce per-symbol (12.5%) and per-sector (25%) gross margin caps.

    Gross margin share = strategy_margin / total_portfolio_margin.
    When a group exceeds its cap, the lowest-ranked strategy in that group
    (last in ``state.candidates``) is removed.  The loop repeats until
    compliant or no candidates remain.
    """
    adjusted = 0

    for _ in range(200):
        total = state.total_margin_used(margins, contract_margin_multiple)
        if total <= 0 or not state.candidates:
            break

        # Build per-symbol and per-sector tallies
        sym_margin: dict[str, float] = {}
        sym_strats: dict[str, list[dict]] = {}
        sec_margin: dict[str, float] = {}
        sec_strats: dict[str, list[dict]] = {}

        for c in state.candidates:
            name = c["name"]
            sym = c.get("symbol", "?")
            sec = c.get("sector", "Other") or "Other"
            m = state.strategy_margin(name, sym, margins, contract_margin_multiple)
            sym_margin[sym] = sym_margin.get(sym, 0.0) + m
            sym_strats.setdefault(sym, []).append(c)
            sec_margin[sec] = sec_margin.get(sec, 0.0) + m
            sec_strats.setdefault(sec, []).append(c)

        violated = False

        for sym, m_sym in sym_margin.items():
            if m_sym / total > max_single_pct:
                worst = sym_strats[sym][-1]  # last = lowest ranked
                state.exclude_strategy(
                    worst["name"], "gross_margin_symbol",
                    f"Symbol '{sym}' margin {m_sym/total:.1%} > {max_single_pct:.1%}",
                )
                adjusted += 1
                violated = True
                break

        if violated:
            continue

        for sec, m_sec in sec_margin.items():
            if m_sec / total > max_sector_pct:
                worst = sec_strats[sec][-1]
                state.exclude_strategy(
                    worst["name"], "gross_margin_sector",
                    f"Sector '{sec}' margin {m_sec/total:.1%} > {max_sector_pct:.1%}",
                )
                adjusted += 1
                violated = True
                break

        if not violated:
            break

    state.log.append(
        f"[gross_margins] Made {adjusted} adjustments, "
        f"{len(state.candidates)} strategies remain"
    )
    return state


def step_adjust_drawdowns(
    state: OptimizerState,
    equity: float,
    max_avg_pct: float = 0.05,
    max_single_pct: float = 0.125,
    max_single_trade_pct: float = 0.05,
    min_fraction: float = 0.1,
) -> OptimizerState:
    """
    Reduce contracts (or remove strategies) where drawdown limits are breached.

    Checks per active strategy (using ``max_oos_drawdown`` from summary):

    - **Single strategy**: ``n × max_dd_per_contract`` ≤ ``max_single_pct × equity``
      → reduce contracts to the largest valid ``n`` in steps of ``min_fraction``
    - **Average drawdown** across all strategies vs ``max_avg_pct × equity``
      → reduce the highest-drawdown strategy by one ``min_fraction`` step,
      repeat until compliant.

    ``max_single_trade_pct`` is checked against the per-contract ``max_oos_drawdown``
    as a proxy (exact trade-level data not always available here).
    """
    adjusted = 0

    # --- per-strategy max drawdown ---
    for c in list(state.candidates):
        name = c["name"]
        n = state.contracts.get(name, 1.0)
        raw_dd = abs(float(c.get("max_oos_drawdown") or 0.0))
        scaled_dd = raw_dd * n
        if scaled_dd > max_single_pct * equity and raw_dd > 0:
            max_allowed = max_single_pct * equity
            new_n = _round_to_fraction(max_allowed / raw_dd, min_fraction)
            new_n = max(min_fraction, new_n)
            if new_n < n:
                state.reduce_contracts(
                    name, new_n, "drawdown_single",
                    f"Drawdown {scaled_dd:,.0f} > {max_single_pct:.1%} equity "
                    f"({max_single_pct * equity:,.0f})",
                )
                adjusted += 1

    # --- average drawdown across portfolio ---
    for _ in range(200):
        if not state.candidates:
            break
        avg = sum(
            abs(float(c.get("max_oos_drawdown") or 0)) * state.contracts.get(c["name"], 1.0)
            for c in state.candidates
        ) / len(state.candidates)
        if avg <= max_avg_pct * equity:
            break
        # Reduce the strategy with the largest scaled drawdown
        worst = max(
            state.candidates,
            key=lambda c: (
                abs(float(c.get("max_oos_drawdown") or 0))
                * state.contracts.get(c["name"], 1.0)
            ),
        )
        name = worst["name"]
        curr = state.contracts.get(name, min_fraction)
        new_n = round(curr - min_fraction, 10)
        if new_n >= min_fraction:
            state.reduce_contracts(
                name, new_n, "drawdown_avg",
                f"Portfolio avg drawdown {avg:,.0f} > {max_avg_pct:.1%} equity",
            )
            adjusted += 1
        else:
            state.exclude_strategy(
                name, "drawdown_avg",
                f"Portfolio avg drawdown too high; cannot reduce contracts further",
            )
            adjusted += 1

    state.log.append(
        f"[drawdowns] Made {adjusted} adjustments, "
        f"{len(state.candidates)} strategies remain"
    )
    return state


# ── Candidate builder ─────────────────────────────────────────────────────────

def build_candidates(
    strategies: list[dict],
    summary_df: "pd.DataFrame | None",
    atr_series: "pd.Series | None",
    margins: dict[str, float],
    default_margin: float,
) -> list[dict]:
    """
    Merge strategy config dicts with summary metrics and ATR into unified
    candidate dicts suitable for the optimizer.

    Args:
        strategies:   List of strategy config dicts (name, symbol, sector, …).
        summary_df:   DataFrame indexed by strategy name with computed metrics.
        atr_series:   Series indexed by strategy name with current ATR in $.
        margins:      Per-symbol margin lookup (from AppConfig.symbol_margins).
        default_margin: Fallback margin when symbol not in ``margins``.

    Returns:
        List of merged dicts, one per strategy.
    """
    result = []
    for s in strategies:
        row = dict(s)
        name = s.get("name", "")
        symbol = s.get("symbol", "")

        # Merge summary metrics
        if summary_df is not None and not summary_df.empty:
            if name in summary_df.index:
                for col, val in summary_df.loc[name].items():
                    if col not in row:
                        row[col] = val

        # ATR
        row["atr"] = float(atr_series.get(name, 0.0)) if atr_series is not None else 0.0

        # Margin
        row["margin_per_contract"] = margins.get(symbol, default_margin)

        result.append(row)
    return result


# ── Workflow runner ────────────────────────────────────────────────────────────

def run_workflow(
    steps: list[tuple],  # [(step_fn, kwargs_dict), ...]
    candidates: list[dict],
    equity: float,
) -> OptimizerState:
    """
    Execute an ordered list of workflow steps against the initial candidate pool.

    Args:
        steps:      List of ``(step_function, kwargs)`` tuples.
        candidates: Initial merged candidate dicts (from :func:`build_candidates`).
        equity:     Starting account balance.

    Returns:
        Final :class:`OptimizerState` after all steps have run.
    """
    state = OptimizerState(
        candidates=list(candidates),
        contracts={
            c.get("name", ""): float(c.get("contracts", 1) or 1)
            for c in candidates
        },
        equity=equity,
    )

    for step_fn, kwargs in steps:
        try:
            state = step_fn(state, **kwargs)
        except Exception as exc:
            state.log.append(f"[ERROR] {step_fn.__name__}: {exc}")

    return state


# ── Summary helpers ────────────────────────────────────────────────────────────

def portfolio_summary(
    state: OptimizerState,
    margins: dict[str, float],
    contract_margin_multiple: float,
) -> dict:
    """Compute aggregate stats for the final portfolio."""
    total_margin = state.total_margin_used(margins, contract_margin_multiple)
    equity = state.equity

    # Per-symbol and per-sector margin shares
    sym_margin: dict[str, float] = {}
    sec_margin: dict[str, float] = {}
    for c in state.candidates:
        name = c["name"]
        sym = c.get("symbol", "?")
        sec = c.get("sector", "Other") or "Other"
        m = state.strategy_margin(name, sym, margins, contract_margin_multiple)
        sym_margin[sym] = sym_margin.get(sym, 0.0) + m
        sec_margin[sec] = sec_margin.get(sec, 0.0) + m

    return {
        "n_strategies": len(state.candidates),
        "n_excluded": len(state.excluded),
        "total_margin": total_margin,
        "margin_pct_equity": total_margin / equity if equity else 0.0,
        "top_symbol_pct": max(sym_margin.values()) / total_margin if total_margin else 0.0,
        "top_sector_pct": max(sec_margin.values()) / total_margin if total_margin else 0.0,
        "symbol_margin": sym_margin,
        "sector_margin": sec_margin,
    }
