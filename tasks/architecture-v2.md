# Portfolio Tracker v2 — Python Architecture Spec

**Date:** 2026-03-14
**Version:** 1.1 (updated with product decisions)
**Status:** Approved for implementation

---

## 0. Product Decisions (Resolved)

| Question | Decision |
|----------|----------|
| MultiWalk dependency | Required — MultiWalk must be installed; registry lookup retained |
| Pricing model | Annual subscription with server-side license validation |
| Platform target | Windows-only (.exe distribution via PyInstaller) |
| CSV format | Microsoft Excel CSV (standard comma-delimited, UTF-8/ANSI) |
| Data migration | Not required — clean install only |
| Feature scope | v1.24 parity + new **Portfolio Eligibility Backtest** feature |

---

## 1. Goals & Non-Goals

### Goals
- Feature-parity with v1.24 (all analytics: MC, correlations, diversification, LOO, backtest)
- Significant performance improvement on compute-heavy modules (MC, LOO, correlations)
- Annual subscription licensing with server-side validation (replaces MultiWalk DLL)
- Maintainable, testable codebase with clear separation of concerns
- Packaged as Windows `.exe` — no Python install required for end users
- New feature: **Portfolio Eligibility Backtest** (walk-forward portfolio construction)

### Non-Goals (v2.0)
- Real-time data feeds or broker integration
- Multi-user collaboration or cloud sync
- Mac/Linux support
- Data migration from v1.24 Excel workbook

---

## 2. Platform: Python + Streamlit (Windows Desktop)

Streamlit runs locally as a web UI served from the packaged `.exe`. Customers open
`http://localhost:8501` in their browser (or it auto-opens). No internet required for
operation — only license validation on startup requires a connection.

```
PyInstaller .exe
    └── Launches Streamlit server on localhost:8501
    └── Auto-opens default browser
    └── Reads MultiWalk CSV folders from local filesystem
    └── Validates license key against subscription API on startup
```

### Why Not Alternatives
| Option | Reason Rejected |
|--------|----------------|
| PyQt6 / tkinter | Significant UI code; charts require manual wiring; slower to iterate |
| FastAPI + React | 3-6x more frontend code; premature before analytics validated |
| Jupyter Notebook | Not shippable as a product |
| Keep VBA | No testing, no vectorization, DLL licensing fragile, VBA has no future |

---

## 3. Technology Stack

| Layer | Library | Replaces |
|-------|---------|---------|
| Data manipulation | `pandas 2.x` | Worksheet arrays, Dictionary lookups |
| Numerical computing | `numpy` | VBA Variant arrays, For loops |
| MC acceleration | `numba` (JIT, cached) | VBA's critical inner loops |
| Statistics | `scipy.stats` | Manual correlation/Pearson code |
| Visualization | `plotly` | Excel charts |
| UI | `streamlit` | Excel sheets + buttons |
| Configuration | `pydantic v2` + YAML | Named ranges, Strategies tab |
| File I/O | `pathlib`, `pandas` | FileSystemObject, Open() |
| Licensing | `cryptography` (JWT) + subscription API | MultiWalkLicense64.dll |
| Packaging | `PyInstaller` | .xlsb distribution |
| Testing | `pytest` + `hypothesis` | None (currently untested) |

**Python version: 3.11+**

---

## 4. Project Structure

```
portfolio-tracker/
├── app.py                          # Streamlit entrypoint + license gate
├── pyproject.toml                  # Dependencies (uv / Poetry)
├── config/
│   └── default_settings.yaml       # Pydantic model defaults
│
├── core/                           # Pure analytics — zero UI dependency
│   ├── config.py                   # Pydantic settings model
│   ├── data_types.py               # Domain dataclasses
│   │
│   ├── ingestion/
│   │   ├── folder_scanner.py       # C_Retrieve_Folder_Locations
│   │   ├── csv_importer.py         # D_Import_Data
│   │   └── date_utils.py           # I_MISC date helpers, IS/OOS resolution
│   │
│   ├── portfolio/
│   │   ├── aggregator.py           # J_Portfolio_Setup
│   │   ├── summary.py              # F_Summary_Tab_Setup (80+ metrics)
│   │   └── strategies.py           # O_Strategies_Tab
│   │
│   ├── analytics/
│   │   ├── monte_carlo.py          # K_MonteCarlo (vectorized + Numba JIT)
│   │   ├── correlations.py         # L_Correlations (3 modes)
│   │   ├── diversification.py      # T_Diversificator (randomized + greedy)
│   │   ├── leave_one_out.py        # S_LeaveOneOut
│   │   ├── backtest.py             # N_BackTest
│   │   ├── margin.py               # M_Margin_Tracking
│   │   └── eligibility/            # NEW FEATURE — see Section 8
│   │       ├── rules.py            # Rule definitions & evaluation engine
│   │       ├── rule_backtest.py    # Walk-forward rule statistics (replaces U module)
│   │       └── portfolio_backtest.py  # NEW: walk-forward portfolio construction
│   │
│   ├── reporting/
│   │   ├── strategy_report.py      # G_Create_Strategy_Tab
│   │   ├── position_check.py       # V_PositionCheck
│   │   └── excel_export.py         # Optional .xlsx export
│   │
│   └── licensing/
│       ├── license_manager.py      # JWT validation + API check
│       └── hardware_id.py          # Machine fingerprint
│
├── ui/
│   ├── pages/
│   │   ├── 01_Import.py
│   │   ├── 02_Portfolio.py
│   │   ├── 03_Monte_Carlo.py
│   │   ├── 04_Correlations.py
│   │   ├── 05_Diversification.py
│   │   ├── 06_Leave_One_Out.py
│   │   ├── 07_Backtest.py
│   │   ├── 08_Margin_Tracking.py
│   │   ├── 09_Position_Check.py
│   │   └── 10_Eligibility_Backtest.py   # NEW FEATURE
│   └── components/
│       ├── strategy_selector.py
│       ├── metrics_table.py
│       └── equity_chart.py
│
└── tests/
    ├── unit/
    │   ├── test_monte_carlo.py
    │   ├── test_correlations.py
    │   ├── test_csv_importer.py
    │   ├── test_eligibility_rules.py
    │   └── test_portfolio_backtest.py
    ├── integration/
    │   └── test_full_pipeline.py
    └── fixtures/
        └── sample_data/            # Anonymized MultiWalk CSVs
```

---

## 5. Data Layer Design

### Replacing Excel Sheets with DataFrames

| Excel Sheet | Python Equivalent | Notes |
|-------------|------------------|-------|
| `DailyM2MEquity` | `daily_pnl: pd.DataFrame` | index=date, columns=strategy names |
| `ClosedTradePNL` | `closed_trades: pd.DataFrame` | columns=[strategy, date, pnl] |
| `Summary` | `summary: pd.DataFrame` | index=strategy, columns=80+ metrics |
| `Portfolio` | `portfolio: pd.DataFrame` | aggregated metrics |
| `Strategies` | `strategies.yaml` + DataFrame | user-editable config file |
| `MW Folder Locations` | `list[Path]` | in-memory only |
| `PortInMarketLong/Short` | `positions: pd.DataFrame` | dates × symbols |

### Core Data Types

```python
# core/data_types.py

from dataclasses import dataclass, field
from pathlib import Path
from datetime import date
import pandas as pd

@dataclass
class Strategy:
    name: str
    folder: Path
    status: str          # Live | Paper | Retired | Pass | etc.
    contracts: int
    symbol: str
    sector: str
    is_start: date
    is_end: date
    oos_start: date
    oos_end: date
    incubation_passed_date: date | None = None
    is_max_drawdown: float = 0.0
    expected_annual_return: float = 0.0

@dataclass
class PortfolioData:
    strategies: list[Strategy]
    daily_pnl: pd.DataFrame        # index=date, columns=strategy names
    closed_trades: pd.DataFrame
    summary_metrics: pd.DataFrame

@dataclass
class MCResult:
    starting_equity: float
    expected_profit: float
    risk_of_ruin: float
    max_drawdown_pct: float
    sharpe_ratio: float
    return_to_drawdown: float
```

### Configuration (Replacing Named Ranges)

```python
# core/config.py — Pydantic v2

from pydantic import BaseModel, Field
from pathlib import Path

class MCConfig(BaseModel):
    simulations: int = 10_000
    period: str = "OOS"                   # IS | OOS | IS+OOS
    risk_ruin_target: float = 0.10
    risk_ruin_tolerance: float = 0.01
    trade_adjustment: float = 0.0
    trade_option: str = "Closed"          # Closed | M2M

class EligibilityConfig(BaseModel):
    days_threshold_oos: int = 0           # NM_INCUBATE
    oos_dd_vs_is_cap: float = 1.5         # NM_DDCAP (0 = disabled)
    status_include: list[str] = ["Live"]  # NM_STATUS_INCLUDE
    efficiency_ratio: float = 0.5         # NM_EFFICIENCY_RATIO
    date_type: str = "OOS Start Date"     # or "Incubation Pass Date"
    enable_sector_analysis: bool = True
    enable_symbol_analysis: bool = True
    max_horizon: int = 12

class PortfolioConfig(BaseModel):
    period_years: float = 3.0
    cutoff_date: str | None = None
    use_cutoff: bool = False
    buy_and_hold_status: str = "Buy&Hold"
    live_status: str = "Live"
    pass_status: str = "Pass"

class AppConfig(BaseModel):
    folders: list[Path] = Field(default_factory=list)
    date_format: str = "DMY"
    portfolio: PortfolioConfig = Field(default_factory=PortfolioConfig)
    monte_carlo: MCConfig = Field(default_factory=MCConfig)
    eligibility: EligibilityConfig = Field(default_factory=EligibilityConfig)
    corr_normal_threshold: float = 0.70
    corr_negative_threshold: float = 0.30
    corr_drawdown_threshold: float = 0.70
```

---

## 6. Analytics: Existing Modules

### 6.1 Monte Carlo (vectorized + Numba)

VBA inner loop replaced with Numba JIT. Speedup: ~600x.

```python
# core/analytics/monte_carlo.py
from numba import njit
import numpy as np

@njit(cache=True)
def _mc_core(pnl_samples, starting_equity, margin_threshold,
             n_scenarios, trades_per_year, trade_adjustment):
    """
    Compiled inner loop. ~50ms vs VBA's ~30s.
    Returns: (final_equity, max_drawdown, ruined) shape (n_scenarios,)
    """
    final_equity = np.empty(n_scenarios)
    max_drawdown = np.empty(n_scenarios)
    ruined = np.zeros(n_scenarios, dtype=np.bool_)
    for i in range(n_scenarios):
        equity = starting_equity
        peak = starting_equity
        dd = 0.0
        for j in range(trades_per_year):
            idx = np.random.randint(0, len(pnl_samples))
            equity += pnl_samples[idx] * (1.0 - trade_adjustment)
            if equity > peak:
                peak = equity
            drawdown = (peak - equity) / peak if peak > 1e-9 else 0.0
            if drawdown > dd:
                dd = drawdown
            if equity < margin_threshold:
                ruined[i] = True
                break
        final_equity[i] = equity
        max_drawdown[i] = dd
    return final_equity, max_drawdown, ruined


def solve_starting_equity(pnl_samples, config, margin_threshold):
    """
    Iterative solver mirrors VBA's +5% / -0.9% adjustment loop.
    Hard cap at 100 iterations (identical to VBA behaviour).
    """
    equity = margin_threshold * 2.0
    ror = 1.0
    for _ in range(100):
        fe, dd, ruined = _mc_core(
            pnl_samples, equity, margin_threshold,
            config.simulations, 252, config.trade_adjustment
        )
        ror = ruined.mean()
        if abs(ror - config.risk_ruin_target) < config.risk_ruin_tolerance:
            break
        equity *= 1.05 if ror > config.risk_ruin_target else 0.991
    return MCResult(
        starting_equity=equity,
        risk_of_ruin=float(ror),
        expected_profit=float(np.mean(fe) - equity),
        max_drawdown_pct=float(np.median(dd)),
        sharpe_ratio=_calc_sharpe(fe, equity),
        return_to_drawdown=_calc_rtd(fe, dd, equity),
    )
```

### 6.2 Correlations (3 modes, vectorized)

```python
# core/analytics/correlations.py
from enum import Enum
from scipy.stats import pearsonr
import numpy as np, pandas as pd

class CorrelationMode(Enum):
    NORMAL = "normal"
    NEGATIVE = "negative"     # Exclude days both strategies profitable
    DRAWDOWN = "drawdown"     # Equity curve synchronization

def compute_correlation_matrix(daily_pnl: pd.DataFrame,
                                mode: CorrelationMode) -> pd.DataFrame:
    strats = daily_pnl.columns
    n = len(strats)
    matrix = np.eye(n)
    for i in range(n):
        for j in range(i + 1, n):
            a = daily_pnl.iloc[:, i].values
            b = daily_pnl.iloc[:, j].values
            if mode == CorrelationMode.NORMAL:
                mask = (a != 0) | (b != 0)
            elif mode == CorrelationMode.NEGATIVE:
                mask = ~((a > 0) & (b > 0))
            else:  # DRAWDOWN
                a = _to_drawdown_series(np.cumsum(a))
                b = _to_drawdown_series(np.cumsum(b))
                mask = np.ones(len(a), dtype=bool)
            corr = pearsonr(a[mask], b[mask])[0] if mask.sum() > 1 else 0.0
            matrix[i, j] = matrix[j, i] = corr
    return pd.DataFrame(matrix, index=strats, columns=strats)

def _to_drawdown_series(equity: np.ndarray) -> np.ndarray:
    peak = np.maximum.accumulate(equity)
    return np.where(peak > 0, (peak - equity) / peak, 0.0)
```

### 6.3 Leave-One-Out (fast because MC is fast)

With vectorized MC, LOO for 20 strategies runs in ~2 seconds vs ~10 minutes in VBA.

```python
# core/analytics/leave_one_out.py
def run_leave_one_out(portfolio_data, config, method="monte_carlo"):
    base = _run_analysis(portfolio_data, config, method)
    rows = []
    for s in portfolio_data.strategies:
        reduced = _remove_strategy(portfolio_data, s.name)
        result = _run_analysis(reduced, config, method)
        rows.append({
            "strategy": s.name,
            "delta_profit": result.expected_profit - base.expected_profit,
            "delta_sharpe": result.sharpe_ratio - base.sharpe_ratio,
            "delta_drawdown": result.max_drawdown_pct - base.max_drawdown_pct,
            "delta_rtd": result.return_to_drawdown - base.return_to_drawdown,
        })
    return pd.DataFrame(rows).sort_values("delta_profit")
```

---

## 7. New Feature: Portfolio Eligibility Backtest

This is the most important new feature and the most complex new module. It extends the
existing `U_BackTest_Eligibility` module from **rule statistics analysis** into a full
**walk-forward portfolio construction backtest**.

### 7.1 What the VBA Module (U) Currently Does

At each month-end, for each of ~160 eligibility rules, it records:
- How many strategies passed the rule (`N`)
- How many of those were profitable in the next 1-12 months (`Win%`)
- Average $ per strategy per month (`$/Month`)
- % difference vs "Baseline (All Eligible)" (`vs Base`)

This tells you *which rules historically predicted better forward performance*,
but it doesn't tell you what happens when you actually USE a rule to construct
a portfolio each month and track the resulting equity curve.

### 7.2 What v2 Adds: Walk-Forward Portfolio Construction

v2 adds a second layer on top of the existing rule analysis:

```
For each month in history:
    1. Evaluate all eligibility rules (existing U module logic)
    2. For a user-selected rule (or ranked set of rules):
       a. Select strategies that passed the rule
       b. Optionally apply ranking to pick top-N
       c. Construct portfolio for that month (equal weight or by contracts)
    3. Record the portfolio's actual P&L for that month
    4. Repeat → full walk-forward equity curve

Output:
    - Equity curve: Rule-selected portfolio vs All-strategies baseline
    - Monthly win rate, average $, drawdown for the selected rule
    - Side-by-side comparison of multiple rules
    - Heatmap: rule performance across all months and horizons
```

### 7.3 Rule System Design

The VBA module hardcodes 160 rules across 8 sections. v2 makes rules fully
configurable — users can define custom rules in the UI or YAML.

```python
# core/analytics/eligibility/rules.py

from dataclasses import dataclass
from enum import Enum
import numpy as np

class RuleType(str, Enum):
    BASELINE = "BASELINE"
    OOS_PROFITABLE = "OOS_PROFITABLE"
    SIMPLE_POSITIVE = "SIMPLE_POSITIVE"        # Last NM > 0
    SIMPLE_NEGATIVE = "SIMPLE_NEGATIVE"        # Last NM < 0
    CONSECUTIVE = "CONSECUTIVE"                # Last NM all positive
    COUNT_POSITIVE = "COUNT_POSITIVE"          # K+ of last NM positive
    MOMENTUM = "MOMENTUM"                      # Recent period > prior period
    ACCELERATION = "ACCELERATION"              # Short ann. > long ann.
    AND_COMBO = "AND_COMBO"                    # Two period windows both positive
    ANY_OF_3 = "ANY_OF_3"                      # Any of 3 windows positive
    ALL_OF_3 = "ALL_OF_3"                      # All of 3 windows positive
    THRESHOLD_ANNUAL = "THRESHOLD_ANNUAL"      # Ann. return > efficiency * expected
    RECOVERY = "RECOVERY"                      # Recent positive after prior negative
    # OOS variants of all above (add _OOS suffix)

@dataclass
class EligibilityRule:
    id: int
    label: str
    rule_type: RuleType
    param1: float = 0.0   # Primary param (e.g. months)
    param2: float = 0.0   # Secondary param
    param3: float = 0.0   # Tertiary param
    is_active: bool = True
    require_oos_profitable: bool = False   # OOS variant flag


def evaluate_rule(rule: EligibilityRule,
                  monthly_pnl: np.ndarray,    # pre-calculated monthly PnL array
                  month_idx: int,             # current month index
                  oos_total_pnl: float,       # total OOS PnL for OOS variants
                  expected_annual: float,      # from Summary sheet
                  efficiency_ratio: float) -> bool:
    """
    Pure function. Evaluates one rule for one strategy at one month.
    No Excel dependencies — fully testable.
    """
    p1 = _trailing_sum(monthly_pnl, month_idx, int(rule.param1))
    p3 = _trailing_sum(monthly_pnl, month_idx, 3)
    p6 = _trailing_sum(monthly_pnl, month_idx, 6)
    p9 = _trailing_sum(monthly_pnl, month_idx, 9)
    p12 = _trailing_sum(monthly_pnl, month_idx, 12)

    passed = _evaluate_base(rule, p1, p3, p6, p9, p12, monthly_pnl, month_idx,
                            expected_annual, efficiency_ratio)

    if rule.require_oos_profitable:
        passed = passed and (oos_total_pnl > 0)

    return passed
```

### 7.4 Walk-Forward Rule Statistics (replaces U module)

```python
# core/analytics/eligibility/rule_backtest.py

import pandas as pd
import numpy as np
from .rules import EligibilityRule, evaluate_rule, build_default_rules

def run_rule_backtest(
    daily_pnl: pd.DataFrame,          # dates × strategies
    summary: pd.DataFrame,             # strategy metadata
    config: EligibilityConfig,
) -> pd.DataFrame:
    """
    Replicates U_BackTest_Eligibility logic in Python.
    Returns: DataFrame with rows=rules, columns=N/WinPct/AvgPnL/vsBase for horizons 1-12.
    """
    rules = build_default_rules()
    monthly_pnl = _aggregate_to_monthly(daily_pnl)    # date × strategy monthly sums
    month_dates = monthly_pnl.index.tolist()
    n_months = len(month_dates)

    results = {r.id: _init_rule_stats(r, config.max_horizon) for r in rules}

    for m_idx in range(n_months - 1):
        for strat in daily_pnl.columns:
            if not _is_eligible_at_month(strat, month_dates[m_idx], summary, config):
                continue
            strat_monthly = monthly_pnl[strat].values
            oos_pnl = _get_oos_pnl(strat, daily_pnl, summary)
            expected = summary.loc[strat, "expected_annual_return"]

            rule_mask = np.array([
                evaluate_rule(r, strat_monthly, m_idx, oos_pnl, expected,
                              config.efficiency_ratio)
                for r in rules
            ])

            for h in range(1, min(config.max_horizon + 1, n_months - m_idx)):
                fwd_pnl = _forward_sum(strat_monthly, m_idx + 1, h)
                for r_idx, rule in enumerate(rules):
                    if rule_mask[r_idx]:
                        results[rule.id]["n"][h] += 1
                        results[rule.id]["wins"][h] += int(fwd_pnl > 0)
                        results[rule.id]["sum_pnl"][h] += fwd_pnl

    return _format_results(results, rules, config.max_horizon)
```

### 7.5 Portfolio Construction Backtest (NEW)

This is the key new capability. Takes the rule analysis one step further:
simulate actually trading using a rule each month and show the resulting equity curve.

```python
# core/analytics/eligibility/portfolio_backtest.py

from dataclasses import dataclass
import pandas as pd
import numpy as np

@dataclass
class PortfolioBacktestConfig:
    rule_id: int | list[int]      # Which rule(s) to use for selection
    max_strategies: int | None    # Cap on number of strategies (None = all passing)
    ranking_metric: str = "oos_pnl"    # How to rank when capping: oos_pnl | momentum_3m | momentum_6m | expected_return
    weighting: str = "equal"           # equal | by_contracts
    rebalance_frequency: str = "monthly"   # monthly (only option in v2.0)
    comparison_rules: list[int] = None     # Additional rules to compare side-by-side


@dataclass
class PortfolioBacktestResult:
    equity_curve: pd.Series              # index=month, values=cumulative PnL
    monthly_pnl: pd.Series
    monthly_strategy_count: pd.Series    # How many strategies active each month
    monthly_selected: pd.DataFrame       # Which strategies were selected each month
    win_rate: float
    avg_monthly_pnl: float
    max_drawdown: float
    sharpe_ratio: float
    vs_baseline_pct: float               # % improvement vs all-strategies baseline


def run_portfolio_backtest(
    daily_pnl: pd.DataFrame,
    summary: pd.DataFrame,
    config_app: EligibilityConfig,
    backtest_config: PortfolioBacktestConfig,
) -> dict[str, PortfolioBacktestResult]:
    """
    Walk-forward portfolio construction.
    Returns dict keyed by rule label → PortfolioBacktestResult.
    Always includes 'Baseline (All Eligible)' as a comparison.

    Algorithm:
        for each month m:
            1. Determine eligible strategies (incubation, DD cap, status filter)
            2. Evaluate selected rule(s) → passing_strategies
            3. If max_strategies: rank passing_strategies by ranking_metric → top-N
            4. Portfolio PnL for month m+1 = mean(daily_pnl[m+1, selected_strategies])
               (equal weight) or weighted by contracts
            5. Record selection, PnL, strategy count
    """
    rules = build_default_rules()
    rule_map = {r.id: r for r in rules}
    monthly_pnl = _aggregate_to_monthly(daily_pnl)
    month_dates = monthly_pnl.index.tolist()

    rule_ids_to_run = _resolve_rule_ids(backtest_config, rules)
    # Always include baseline for comparison
    baseline_id = next(r.id for r in rules if r.rule_type.value == "BASELINE")
    if baseline_id not in rule_ids_to_run:
        rule_ids_to_run = [baseline_id] + rule_ids_to_run

    portfolio_monthly: dict[int, list[float]] = {rid: [] for rid in rule_ids_to_run}
    selected_monthly: dict[int, list[set]] = {rid: [] for rid in rule_ids_to_run}

    for m_idx in range(len(month_dates) - 1):
        eval_date = month_dates[m_idx]

        # Step 1: Base eligibility filter (same for all rules)
        eligible = [
            s for s in daily_pnl.columns
            if _is_eligible_at_month(s, eval_date, summary, config_app)
        ]

        for rid in rule_ids_to_run:
            rule = rule_map[rid]

            # Step 2: Apply rule filter
            passing = [
                s for s in eligible
                if evaluate_rule(
                    rule,
                    monthly_pnl[s].values,
                    m_idx,
                    _get_oos_pnl(s, daily_pnl, summary),
                    summary.loc[s, "expected_annual_return"],
                    config_app.efficiency_ratio,
                )
            ]

            # Step 3: Optional ranking + cap
            if backtest_config.max_strategies and len(passing) > backtest_config.max_strategies:
                passing = _rank_strategies(
                    passing, monthly_pnl, m_idx, summary,
                    backtest_config.ranking_metric,
                    backtest_config.max_strategies,
                )

            selected_monthly[rid].append(set(passing))

            # Step 4: Portfolio PnL for next month
            if not passing:
                portfolio_monthly[rid].append(0.0)
            else:
                next_month_pnl = monthly_pnl.iloc[m_idx + 1][passing]
                if backtest_config.weighting == "equal":
                    portfolio_monthly[rid].append(next_month_pnl.sum())
                else:
                    contracts = summary.loc[passing, "contracts"]
                    portfolio_monthly[rid].append(
                        (next_month_pnl * contracts).sum() / contracts.sum()
                    )

    return {
        rule_map[rid].label: _build_result(
            portfolio_monthly[rid],
            selected_monthly[rid],
            month_dates[1:],
            portfolio_monthly[baseline_id],
        )
        for rid in rule_ids_to_run
    }


def _rank_strategies(
    strategies: list[str],
    monthly_pnl: pd.DataFrame,
    m_idx: int,
    summary: pd.DataFrame,
    metric: str,
    top_n: int,
) -> list[str]:
    """Rank strategies by chosen metric and return top N."""
    scores = {}
    for s in strategies:
        if metric == "oos_pnl":
            scores[s] = summary.loc[s, "oos_total_pnl"]
        elif metric == "momentum_3m":
            scores[s] = _trailing_sum(monthly_pnl[s].values, m_idx, 3)
        elif metric == "momentum_6m":
            scores[s] = _trailing_sum(monthly_pnl[s].values, m_idx, 6)
        elif metric == "expected_return":
            scores[s] = summary.loc[s, "expected_annual_return"]
        else:
            scores[s] = 0.0
    return sorted(strategies, key=lambda s: scores[s], reverse=True)[:top_n]
```

### 7.6 UI: Eligibility Backtest Page

```python
# ui/pages/10_Eligibility_Backtest.py

import streamlit as st
import plotly.graph_objects as go
from core.analytics.eligibility.rule_backtest import run_rule_backtest
from core.analytics.eligibility.portfolio_backtest import (
    run_portfolio_backtest, PortfolioBacktestConfig
)

st.title("Portfolio Eligibility Backtest")

tab1, tab2 = st.tabs(["Rule Statistics", "Portfolio Construction Backtest"])

# ── Tab 1: Rule Statistics (existing U module functionality) ──────────────
with tab1:
    st.markdown("""
    **How it works:** At each month-end, each rule is evaluated for all eligible
    strategies. Forward performance (next 1-12 months) is recorded for strategies
    that passed. This shows which rules historically predicted better performance.
    """)
    if st.button("Run Rule Analysis", type="primary"):
        with st.spinner("Evaluating 160 rules across all months..."):
            results = run_rule_backtest(
                st.session_state.portfolio.daily_pnl,
                st.session_state.portfolio.summary_metrics,
                st.session_state.config.eligibility,
            )
        st.dataframe(
            results.style
                .background_gradient(subset=[c for c in results.columns if "vs Base" in c],
                                     cmap="RdYlGn", vmin=-0.5, vmax=0.5)
                .format({c: "{:.0f}" for c in results.columns if "$/Month" in c})
                .format({c: "{:.1%}" for c in results.columns if "Win%" in c or "vs Base" in c}),
            use_container_width=True,
            height=600,
        )

# ── Tab 2: Portfolio Construction Backtest (NEW) ──────────────────────────
with tab2:
    st.markdown("""
    **How it works:** Simulates using a rule each month to select which strategies
    are in the portfolio. Shows the resulting walk-forward equity curve vs baseline.
    """)

    col1, col2 = st.columns(2)
    with col1:
        selected_rules = st.multiselect(
            "Rules to backtest",
            options=get_rule_labels(),
            default=["Last 3M > 0", "Last 6M > 0", "3M AND 6M > 0"],
            max_selections=5,
        )
        max_strats = st.number_input("Max strategies per month (0 = no limit)", 0, 50, 0)
    with col2:
        ranking = st.selectbox("Ranking metric (when capping)",
                               ["oos_pnl", "momentum_3m", "momentum_6m", "expected_return"])
        weighting = st.selectbox("Portfolio weighting", ["equal", "by_contracts"])

    if st.button("Run Portfolio Backtest", type="primary"):
        cfg = PortfolioBacktestConfig(
            rule_id=[get_rule_id(r) for r in selected_rules],
            max_strategies=max_strats or None,
            ranking_metric=ranking,
            weighting=weighting,
            comparison_rules=None,
        )
        with st.spinner("Running walk-forward backtest..."):
            results = run_portfolio_backtest(
                st.session_state.portfolio.daily_pnl,
                st.session_state.portfolio.summary_metrics,
                st.session_state.config.eligibility,
                cfg,
            )

        # Equity curves
        fig = go.Figure()
        for rule_label, result in results.items():
            fig.add_trace(go.Scatter(
                x=result.equity_curve.index,
                y=result.equity_curve.values,
                name=rule_label,
                line=dict(width=2 if "Baseline" not in rule_label else 1,
                          dash="dash" if "Baseline" in rule_label else "solid"),
            ))
        fig.update_layout(title="Walk-Forward Equity Curves by Rule",
                          xaxis_title="Month", yaxis_title="Cumulative P&L ($)")
        st.plotly_chart(fig, use_container_width=True)

        # Summary metrics table
        summary_rows = []
        for rule_label, result in results.items():
            summary_rows.append({
                "Rule": rule_label,
                "Win Rate": f"{result.win_rate:.1%}",
                "Avg Monthly P&L": f"${result.avg_monthly_pnl:,.0f}",
                "Max Drawdown": f"{result.max_drawdown:.1%}",
                "Sharpe": f"{result.sharpe_ratio:.2f}",
                "vs Baseline": f"{result.vs_baseline_pct:+.1%}",
            })
        st.dataframe(pd.DataFrame(summary_rows), use_container_width=True)

        # Monthly strategy count
        st.subheader("Strategies Selected Per Month")
        count_df = pd.DataFrame({
            label: result.monthly_strategy_count
            for label, result in results.items()
            if "Baseline" not in label
        })
        st.area_chart(count_df)
```

---

## 8. Licensing: Annual Subscription

Replace the MultiWalk DLL with a proper subscription system.

### Architecture

```
Customer buys annual subscription → receives license key (JWT, RS256-signed)
License key contains: { customer_name, hardware_id, expiry, tier, issued_at }
App validates on startup:
    1. Decode JWT with embedded public key (offline check)
    2. Verify hardware_id matches current machine
    3. If expiry < 30 days away → prompt renewal
    4. Optional: phone-home to subscription API for revocation check
```

```python
# core/licensing/license_manager.py

import jwt, hashlib, uuid
from datetime import datetime, timedelta
from pathlib import Path

LICENSE_FILE = Path.home() / ".portfolio_tracker" / "license.key"
PUBLIC_KEY = "-----BEGIN PUBLIC KEY-----\n..."  # Embedded at build time

def validate_license() -> dict | None:
    if not LICENSE_FILE.exists():
        return None
    key = LICENSE_FILE.read_text().strip()
    try:
        claims = jwt.decode(key, PUBLIC_KEY, algorithms=["RS256"])
    except jwt.ExpiredSignatureError:
        return None
    except jwt.InvalidTokenError:
        return None
    if claims.get("hardware_id") != get_hardware_id():
        return None
    return claims

def days_until_expiry(claims: dict) -> int:
    expiry = datetime.fromisoformat(claims["expiry"])
    return (expiry - datetime.now()).days

def get_hardware_id() -> str:
    mac = uuid.getnode()
    return hashlib.sha256(f"{mac}".encode()).hexdigest()[:16]
```

### Subscription API (minimal server, e.g. on Railway/Render)

```python
# server/api.py — simple FastAPI service
# POST /activate   { license_key, hardware_id } → validates + logs activation
# GET  /check      { license_key }              → returns valid/revoked/expired
# POST /issue      { customer_email, tier }     → admin: generates new key
```

This server holds the private key and customer database. The app embeds only the
public key — so license files cannot be forged even if the app is decompiled.

---

## 9. Performance Summary

| Module | VBA Time | Python Time | Speedup |
|--------|----------|-------------|---------|
| Monte Carlo (10k scenarios) | ~30s | ~50ms | 600x |
| Leave-One-Out (20 strats) | ~10 min | ~1s | 600x |
| Correlations (30×30 matrix) | ~2 min | <10ms | 12,000x |
| Rule Backtest (160 rules) | ~5 min | ~3s | 100x |
| Portfolio Backtest (5 rules) | n/a | ~5s | new |
| CSV import (20 strategies) | ~30s | ~1s | 30x |

---

## 10. Testing Strategy

### Test Pyramid

```
Unit tests (fast, isolated, deterministic):
    test_monte_carlo.py        — fixed RNG seed → identical output every run
    test_correlations.py       — known data → verify against scipy.stats
    test_eligibility_rules.py  — each of 160 rules with synthetic monthly data
    test_portfolio_backtest.py — 3-strategy toy portfolio, verify equity curve math
    test_csv_importer.py       — edge cases: missing files, bad dates, encoding
    test_date_utils.py         — IS/OOS resolution, cutoff date logic, trading days

Integration tests:
    test_full_pipeline.py      — scan → import → portfolio → MC → rule backtest
    test_settings_roundtrip.py — export config → import → identical settings

Regression tests (golden dataset):
    Run v1.24 on anonymized test data → capture all metrics as JSON fixtures
    v2 must match within ±0.5% tolerance (MC differs due to RNG; document this)
```

### Look-Ahead Bias Validation
The portfolio backtest must be verified as free of look-ahead bias. Each month's
portfolio selection must use only data available on or before that month-end date.
Dedicated test: construct a toy case with known correct output, verify no future data
bleeds into the selection decision.

---

## 11. Build Order (Phased Delivery)

### Phase 1 — Data Foundation (2-3 weeks)
- [ ] `core/ingestion/folder_scanner.py` — scan MultiWalk folders, validate CSVs
- [ ] `core/ingestion/csv_importer.py` — parse EquityData + TradeData CSVs
- [ ] `core/ingestion/date_utils.py` — IS/OOS period logic, cutoff dates, trading calendar
- [ ] `core/config.py` + `core/data_types.py`
- [ ] `ui/pages/01_Import.py` — folder picker, import progress, data preview table
- [ ] Unit tests for all ingestion code

### Phase 2 — Portfolio & Summary (1-2 weeks)
- [ ] `core/portfolio/aggregator.py` — aggregate strategies into portfolio
- [ ] `core/portfolio/summary.py` — 80+ metrics per strategy
- [ ] `ui/pages/02_Portfolio.py` — strategy table with filter/sort
- [ ] Integration test: metrics match v1.24

### Phase 3 — Monte Carlo (1 week)
- [ ] `core/analytics/monte_carlo.py` — Numba JIT inner loop
- [ ] `ui/pages/03_Monte_Carlo.py`
- [ ] Regression test vs v1.24 output

### Phase 4 — Correlations + Diversification (1-2 weeks)
- [ ] `core/analytics/correlations.py` — 3 modes
- [ ] `core/analytics/diversification.py` — randomized + greedy
- [ ] `ui/pages/04_Correlations.py` + `05_Diversification.py`

### Phase 5 — Eligibility Backtest (2-3 weeks) ← NEW FEATURE
- [ ] `core/analytics/eligibility/rules.py` — all 160 rules as pure functions
- [ ] `core/analytics/eligibility/rule_backtest.py` — walk-forward rule statistics
- [ ] `core/analytics/eligibility/portfolio_backtest.py` — portfolio construction
- [ ] `ui/pages/10_Eligibility_Backtest.py` — both tabs
- [ ] Unit tests: each rule type, look-ahead bias validation

### Phase 6 — Advanced Analytics (2 weeks)
- [ ] `core/analytics/leave_one_out.py`
- [ ] `core/analytics/backtest.py`
- [ ] `core/analytics/margin.py`
- [ ] Remaining UI pages

### Phase 7 — Licensing + Packaging (1-2 weeks)
- [ ] `core/licensing/` — JWT hardware-bound keys
- [ ] Subscription API server (minimal FastAPI)
- [ ] `core/reporting/excel_export.py` — optional .xlsx export
- [ ] Settings export/import (replaces Q module)
- [ ] PyInstaller packaging + code signing
- [ ] Auto-update check on startup
- [ ] End-to-end regression test suite

---

## 12. Key Risks & Mitigations

| Risk | Impact | Mitigation |
|------|--------|-----------|
| Portfolio backtest has look-ahead bias | Critical — invalidates results | Dedicated unit tests; code review gate before release |
| MC output doesn't match v1.24 exactly | High | Document RNG difference; golden dataset tests within ±0.5% |
| IS/OOS date logic bugs | High | Mirror VBA's ResolveOOSDates exactly; unit test every edge case |
| Eligibility rule 160 definitions mismatch | Medium | Port each rule type from VBA line-by-line; unit test each |
| PyInstaller + Numba JIT conflict | Medium | Test on clean Windows VM; pre-compile with cache=True |
| Subscription API downtime blocks startup | Medium | Cache last valid check for 7 days offline grace period |
| MultiWalk CSV format variations across MW versions | Medium | Header detection; log warnings on parse failure |
