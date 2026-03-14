# Portfolio Tracker v2 — Python Architecture Spec

**Date:** 2026-03-14
**Status:** Draft
**Scope:** Full rewrite of the Excel/VBA Portfolio Tracker as a Python application

---

## 1. Goals & Non-Goals

### Goals
- Feature-parity with v1.24 (all analytics: MC, correlations, diversification, LOO, backtest)
- Significant performance improvement on compute-heavy modules (MC, LOO, correlations)
- Professional licensing and delivery mechanism (no DLL dependency)
- Maintainable, testable codebase with clear separation of concerns
- Deployable as a local desktop app today; upgradeable to SaaS later

### Non-Goals (v2.0)
- Real-time data feeds or broker integration
- Multi-user collaboration
- Cloud storage (v2.0 reads local MultiWalk CSV folders)
- Rebuilding the Excel workbook UI (replaced entirely)

---

## 2. Recommended Platform: Python + Streamlit (Local Desktop)

### Why Streamlit
- Zero frontend code required; interactive widgets, charts, and tables are built-in
- Runs locally — customers point it at their MultiWalk folders (no upload friction)
- Can be packaged as a Windows `.exe` via PyInstaller for non-technical users
- The same codebase can be deployed as a hosted web app later with no changes
- Streamlit's session state maps cleanly to the Excel workbook's named-range config

### Why Not the Alternatives
| Option | Reason Rejected |
|--------|----------------|
| Excel VBA (evolve) | VBA has no future; no testing, no vectorization, no packaging |
| PyQt6 / tkinter | Significant UI code to write; charts require matplotlib wiring |
| FastAPI + React | 3-6x more code; premature before analytics layer is validated |
| Jupyter Notebook | Not shippable to non-technical customers |

### Migration Path
```
v2.0: Streamlit local app (Python analytics + Streamlit UI)
v2.5: Optional hosted mode (same code, hosted on Railway/Render/AWS)
v3.0: FastAPI backend + React frontend (if SaaS growth justifies it)
```

---

## 3. Technology Stack

| Layer | Library | Replaces |
|-------|---------|---------|
| Data manipulation | `pandas` | Worksheet arrays, Dictionary lookups |
| Numerical computing | `numpy` | VBA Variant arrays, For loops |
| Monte Carlo acceleration | `numba` (JIT) | VBA's slow inner loops |
| Statistics | `scipy.stats` | Manual Pearson correlation code |
| Visualization | `plotly` | Excel charts |
| UI | `streamlit` | Excel sheets + buttons |
| Configuration | `pydantic` + YAML | Named ranges, Strategies tab |
| File I/O | `pathlib`, `pandas` | FileSystemObject, Open() |
| Licensing | `cryptography` + license server | MultiWalkLicense64.dll |
| Packaging | `PyInstaller` | .xlsb distribution |
| Testing | `pytest` + `hypothesis` | None (untested currently) |

### Python Version
- **Python 3.11+** (required for Numba compatibility and performance)

---

## 4. Project Structure

```
portfolio-tracker/
├── app.py                          # Streamlit entrypoint
├── pyproject.toml                  # Dependencies (Poetry or uv)
├── config/
│   └── default_settings.yaml       # Default named range equivalents
│
├── core/                           # Pure Python analytics — no UI dependency
│   ├── __init__.py
│   ├── config.py                   # Pydantic settings model
│   ├── data_types.py               # Dataclasses / TypedDicts for domain objects
│   │
│   ├── ingestion/                  # Replaces C + D modules
│   │   ├── folder_scanner.py       # C_Retrieve_Folder_Locations
│   │   ├── csv_importer.py         # D_Import_Data
│   │   └── date_utils.py           # I_MISC date helpers
│   │
│   ├── portfolio/                  # Replaces J + F + O modules
│   │   ├── aggregator.py           # J_Portfolio_Setup
│   │   ├── summary.py              # F_Summary_Tab_Setup
│   │   └── strategies.py           # O_Strategies_Tab
│   │
│   ├── analytics/                  # Replaces K + L + T + S + N modules
│   │   ├── monte_carlo.py          # K_MonteCarlo (vectorized)
│   │   ├── correlations.py         # L_Correlations
│   │   ├── diversification.py      # T_Diversificator
│   │   ├── leave_one_out.py        # S_LeaveOneOut
│   │   ├── backtest.py             # N_BackTest
│   │   └── margin.py              # M_Margin_Tracking
│   │
│   ├── reporting/                  # Replaces G + H + V modules
│   │   ├── strategy_report.py      # G_Create_Strategy_Tab
│   │   ├── position_check.py       # V_PositionCheck
│   │   └── excel_export.py         # Export to .xlsx for customers who want it
│   │
│   └── licensing/                  # Replaces A + B modules
│       ├── license_manager.py
│       └── hardware_id.py
│
├── ui/                             # Streamlit pages — thin wrappers over core/
│   ├── pages/
│   │   ├── 01_Import.py
│   │   ├── 02_Portfolio.py
│   │   ├── 03_Monte_Carlo.py
│   │   ├── 04_Correlations.py
│   │   ├── 05_Diversification.py
│   │   ├── 06_Leave_One_Out.py
│   │   ├── 07_Backtest.py
│   │   ├── 08_Margin_Tracking.py
│   │   └── 09_Position_Check.py
│   └── components/                 # Shared UI widgets
│       ├── strategy_selector.py
│       ├── date_range_picker.py
│       └── metrics_table.py
│
└── tests/
    ├── unit/
    │   ├── test_monte_carlo.py
    │   ├── test_correlations.py
    │   ├── test_csv_importer.py
    │   └── test_diversification.py
    ├── integration/
    │   └── test_full_pipeline.py
    └── fixtures/
        └── sample_data/            # Anonymized MultiWalk CSVs for testing
```

---

## 5. Data Layer Design

### Replacing the Excel Sheet System

The VBA tool stores all state in Excel worksheets. Python replaces this with in-memory pandas DataFrames, persisted as Parquet files for session caching.

| Excel Sheet | Python Equivalent | Format |
|-------------|------------------|--------|
| `DailyM2MEquity` | `daily_pnl: pd.DataFrame` | dates × strategies matrix |
| `ClosedTradePNL` | `closed_trades: pd.DataFrame` | trades × strategies |
| `Summary` | `summary: pd.DataFrame` | strategies × metrics |
| `Portfolio` | `portfolio: pd.DataFrame` | aggregated metrics |
| `Strategies` | `strategies.yaml` + `pd.DataFrame` | user-editable config |
| `MW Folder Locations` | `folders: list[Path]` | in-memory only |
| `PortInMarketLong/Short` | `positions: pd.DataFrame` | dates × symbols |

### Core Data Types

```python
# core/data_types.py

from dataclasses import dataclass, field
from pathlib import Path
import pandas as pd

@dataclass
class Strategy:
    name: str
    folder: Path
    status: str                     # Live | Paper | Retired | etc.
    contracts: int
    symbol: str
    sector: str
    is_start: pd.Timestamp
    is_end: pd.Timestamp
    oos_start: pd.Timestamp
    oos_end: pd.Timestamp

@dataclass
class PortfolioData:
    strategies: list[Strategy]
    daily_pnl: pd.DataFrame         # index=date, columns=strategy names
    closed_trades: pd.DataFrame     # index=trade_id, columns=[strategy, date, pnl]
    summary_metrics: pd.DataFrame   # index=strategy name, columns=80+ metrics

@dataclass
class MCResult:
    expected_profit: float
    starting_equity: float
    risk_of_ruin: float
    max_drawdown_pct: float
    sharpe_ratio: float
    return_to_drawdown: float
    scenarios: pd.DataFrame         # Full scenario array if needed
```

### Replacing Named Ranges (Configuration)

```python
# core/config.py

from pydantic import BaseModel, Field
from pathlib import Path

class MCConfig(BaseModel):
    simulations: int = 10_000
    period: str = "OOS"             # IS | OOS | IS+OOS
    risk_ruin_target: float = 0.10
    risk_ruin_tolerance: float = 0.01
    trade_adjustment: float = 0.0
    trade_option: str = "Closed"    # Closed | M2M

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
    # Correlation thresholds
    corr_normal_threshold: float = 0.70
    corr_negative_threshold: float = 0.30
    corr_drawdown_threshold: float = 0.70
```

---

## 6. Analytics Layer — Module-by-Module Design

### 6.1 Monte Carlo (`core/analytics/monte_carlo.py`)

The single biggest performance win. The VBA inner loop is replaced with NumPy vectorized operations.

**VBA approach:** nested For loop over scenarios × trades (2.5M iterations, ~30s in VBA)

**Python approach:** fully vectorized with NumPy random sampling

```python
import numpy as np
from numba import njit

@njit(cache=True)
def _mc_core(pnl_samples: np.ndarray,
             starting_equity: float,
             margin_threshold: float,
             n_scenarios: int,
             trades_per_year: int,
             trade_adjustment: float) -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    """
    JIT-compiled inner loop. Runs in ~50ms vs VBA's ~30s.
    Returns: (final_equity, max_drawdown, ruined_flag) arrays, shape (n_scenarios,)
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
            drawdown = (peak - equity) / peak if peak > 0 else 0.0
            if drawdown > dd:
                dd = drawdown
            if equity < margin_threshold:
                ruined[i] = True
                break
        final_equity[i] = equity
        max_drawdown[i] = dd

    return final_equity, max_drawdown, ruined


def solve_starting_equity(pnl_samples: np.ndarray,
                           config: MCConfig,
                           margin_threshold: float) -> MCResult:
    """
    Iterative solver: adjusts starting equity until RoR hits target.
    Mirrors VBA's Do..Loop with +5%/-0.9% adjustment steps.
    """
    equity = margin_threshold * 2
    for _ in range(100):
        final_equity, max_dd, ruined = _mc_core(
            pnl_samples, equity, margin_threshold,
            config.simulations, 252, config.trade_adjustment
        )
        ror = ruined.mean()
        if abs(ror - config.risk_ruin_target) < config.risk_ruin_tolerance:
            break
        elif ror > config.risk_ruin_target:
            equity *= 1.05
        else:
            equity *= 0.991  # matches VBA's 0.9% decrease

    return MCResult(
        starting_equity=equity,
        risk_of_ruin=ror,
        expected_profit=np.mean(final_equity - equity),
        max_drawdown_pct=np.median(max_dd),
        sharpe_ratio=_calc_sharpe(final_equity, equity),
        return_to_drawdown=_calc_rtd(final_equity, max_dd, equity),
    )
```

**Performance:** ~50ms per strategy run vs ~30s in VBA. Leave-One-Out with 20 strategies drops from ~10 minutes to ~1 second.

---

### 6.2 Correlations (`core/analytics/correlations.py`)

Three correlation modes, all vectorized with NumPy/SciPy.

```python
import numpy as np
from scipy.stats import pearsonr
from enum import Enum

class CorrelationMode(Enum):
    NORMAL = "normal"           # Standard Pearson
    NEGATIVE = "negative"       # Exclude days both strategies profitable
    DRAWDOWN = "drawdown"       # Equity curve synchronization

def compute_correlation_matrix(daily_pnl: pd.DataFrame,
                                mode: CorrelationMode) -> pd.DataFrame:
    strategies = daily_pnl.columns
    n = len(strategies)
    matrix = np.eye(n)

    for i in range(n):
        for j in range(i + 1, n):
            a = daily_pnl.iloc[:, i].values
            b = daily_pnl.iloc[:, j].values

            if mode == CorrelationMode.NORMAL:
                mask = (a != 0) | (b != 0)
            elif mode == CorrelationMode.NEGATIVE:
                # Exclude days where BOTH are profitable (matches VBA logic)
                mask = ~((a > 0) & (b > 0))
            elif mode == CorrelationMode.DRAWDOWN:
                # Convert to cumulative equity curves, then correlate drawdowns
                a = _to_drawdown_series(a)
                b = _to_drawdown_series(b)
                mask = np.ones(len(a), dtype=bool)

            if mask.sum() < 2:
                corr = 0.0
            else:
                corr, _ = pearsonr(a[mask], b[mask])

            matrix[i, j] = matrix[j, i] = corr

    return pd.DataFrame(matrix, index=strategies, columns=strategies)

def _to_drawdown_series(pnl: np.ndarray) -> np.ndarray:
    equity = np.cumsum(pnl)
    peak = np.maximum.accumulate(equity)
    drawdown = np.where(peak > 0, (peak - equity) / peak, 0)
    return drawdown
```

---

### 6.3 Diversification (`core/analytics/diversification.py`)

Two algorithms; both operate over the correlation matrix.

```python
def greedy_diversification(
    summary: pd.DataFrame,
    corr_matrix: pd.DataFrame,
    n_strategies: int,
    sort_metric: str = "sharpe"
) -> list[str]:
    """
    Greedy algorithm: add strategies one at a time, always picking the one
    that minimizes average correlation with already-selected strategies.
    O(s^2) - fast in Python.
    """
    candidates = summary.sort_values(sort_metric, ascending=False).index.tolist()
    selected = [candidates[0]]

    while len(selected) < n_strategies and candidates:
        best, best_score = None, float("inf")
        for candidate in candidates:
            if candidate in selected:
                continue
            avg_corr = corr_matrix.loc[candidate, selected].mean()
            if avg_corr < best_score:
                best_score = avg_corr
                best = candidate
        if best:
            selected.append(best)
            candidates.remove(best)

    return selected


def randomized_diversification(
    summary: pd.DataFrame,
    corr_matrix: pd.DataFrame,
    n_strategies: int,
    iterations: int = 500,
    sort_metric: str = "sharpe"
) -> pd.DataFrame:
    """
    Monte Carlo over strategy subsets: average diversification benefit.
    O(iterations * s) - vectorized with NumPy.
    """
    names = summary.index.tolist()
    results = []
    for _ in range(iterations):
        subset = np.random.choice(names, size=n_strategies, replace=False)
        sub_corr = corr_matrix.loc[subset, subset].values
        avg_corr = (sub_corr.sum() - n_strategies) / (n_strategies * (n_strategies - 1))
        sub_metrics = summary.loc[subset, sort_metric]
        results.append({"subset": subset, "avg_corr": avg_corr,
                         "avg_metric": sub_metrics.mean()})
    return pd.DataFrame(results).sort_values("avg_metric", ascending=False)
```

---

### 6.4 Leave-One-Out (`core/analytics/leave_one_out.py`)

```python
def run_leave_one_out(
    portfolio_data: PortfolioData,
    config: AppConfig,
    method: str = "monte_carlo"    # "monte_carlo" | "chronological"
) -> pd.DataFrame:
    """
    For each strategy, remove it and recompute portfolio metrics.
    With vectorized MC this runs in <2s for 20 strategies.
    """
    base_result = run_portfolio_mc(portfolio_data, config)
    results = []

    for strategy in portfolio_data.strategies:
        reduced = _remove_strategy(portfolio_data, strategy.name)
        if method == "monte_carlo":
            result = run_portfolio_mc(reduced, config)
        else:
            result = run_chronological_backtest(reduced, config)

        results.append({
            "strategy": strategy.name,
            "delta_profit": result.expected_profit - base_result.expected_profit,
            "delta_sharpe": result.sharpe_ratio - base_result.sharpe_ratio,
            "delta_drawdown": result.max_drawdown_pct - base_result.max_drawdown_pct,
            "delta_rtd": result.return_to_drawdown - base_result.return_to_drawdown,
        })

    return pd.DataFrame(results).sort_values("delta_profit", ascending=True)
```

---

### 6.5 Folder Scanner (`core/ingestion/folder_scanner.py`)

Replaces the Windows FileSystemObject with `pathlib` — cross-platform.

```python
from pathlib import Path
from dataclasses import dataclass

REQUIRED_FILES = ["EquityData.csv", "Walkforward In-Out Periods Analysis Details.csv"]

def scan_folders(base_folders: list[Path]) -> tuple[list[StrategyFolder], list[str]]:
    """
    Returns (valid_strategies, warnings).
    Warnings include duplicates, missing CSVs, empty folders.
    """
    seen_names: dict[str, Path] = {}
    strategies, warnings = [], []

    for base in base_folders:
        if not base.exists():
            warnings.append(f"Folder not found: {base}")
            continue
        for folder in sorted(base.iterdir()):
            if not folder.is_dir():
                continue
            missing = [f for f in REQUIRED_FILES if not (folder / f).exists()]
            if missing:
                warnings.append(f"{folder.name}: missing {missing}")
                continue
            if folder.name in seen_names:
                warnings.append(f"Duplicate: {folder.name} in {base} and {seen_names[folder.name]}")
                continue
            seen_names[folder.name] = base
            strategies.append(StrategyFolder(name=folder.name, path=folder))

    return strategies, warnings
```

---

## 7. Licensing System

Replace the fragile DLL dependency with a proper software licensing system.

### Approach: Hardware-Bound License Keys

```
Customer purchases → receives license key (JWT signed with private key)
App validates key locally: key contains hardware_id, expiry, customer_name
Hardware ID = SHA256(MAC address + CPU serial + hostname)
Optional: phone-home validation for stricter enforcement
```

```python
# core/licensing/license_manager.py

import jwt
import hashlib
import platform
from datetime import datetime

def get_hardware_id() -> str:
    """Generates a stable machine fingerprint."""
    components = [
        platform.node(),
        platform.processor(),
        _get_mac_address(),
    ]
    return hashlib.sha256("|".join(components).encode()).hexdigest()[:16]

def validate_license(license_key: str, public_key: str) -> dict | None:
    """
    Returns decoded license claims if valid, None if invalid/expired.
    Claims: { customer_name, hardware_id, expiry, features }
    """
    try:
        claims = jwt.decode(license_key, public_key, algorithms=["RS256"])
        hw_id = get_hardware_id()
        if claims["hardware_id"] != hw_id:
            return None
        if datetime.fromisoformat(claims["expiry"]) < datetime.now():
            return None
        return claims
    except Exception:
        return None
```

**Benefits over DLL:**
- No dependency on MultiWalk installation
- Works offline
- Revocable (short-lived keys with renewal)
- Cross-platform
- Keys can encode feature tiers (Basic / Pro / Enterprise)

---

## 8. UI Layer (Streamlit)

Each page is a thin wrapper over the `core/` analytics. No business logic in the UI.

```python
# ui/pages/03_Monte_Carlo.py

import streamlit as st
from core.analytics.monte_carlo import solve_starting_equity
from core.config import MCConfig

st.title("Monte Carlo Simulation")

with st.sidebar:
    st.subheader("Settings")
    simulations = st.slider("Simulations", 1_000, 50_000, 10_000, step=1_000)
    period = st.selectbox("Period", ["OOS", "IS", "IS+OOS"])
    ror_target = st.slider("Risk of Ruin Target", 0.01, 0.25, 0.10)
    trade_adj = st.slider("Trade Adjustment", 0.0, 0.5, 0.0)

config = MCConfig(simulations=simulations, period=period,
                  risk_ruin_target=ror_target, trade_adjustment=trade_adj)

if st.button("Run Monte Carlo", type="primary"):
    portfolio = st.session_state.get("portfolio_data")
    if not portfolio:
        st.error("Import data first.")
    else:
        with st.spinner("Running simulations..."):
            results = solve_starting_equity(portfolio.get_pnl_for_period(period), config)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Starting Equity", f"${results.starting_equity:,.0f}")
        col2.metric("Expected Annual Profit", f"${results.expected_profit:,.0f}")
        col3.metric("Risk of Ruin", f"{results.risk_of_ruin:.1%}")
        col4.metric("Max Drawdown", f"{results.max_drawdown_pct:.1%}")

        st.plotly_chart(results.plot_distribution())
```

---

## 9. Testing Strategy

The current VBA codebase has zero tests. This is the biggest quality risk. v2 must have:

### Test Pyramid

```
Unit tests (fast, isolated):
  - MC with fixed RNG seed → reproducible output
  - Correlation modes with known data → verify against scipy
  - Date range / IS-OOS period logic
  - CSV parsing edge cases (missing files, bad dates, encoding)

Integration tests:
  - Full pipeline: scan folders → import → portfolio → MC → LOO
  - Settings export/import round-trip
  - Results match known-good v1.24 Excel output (regression tests)

Property-based tests (hypothesis):
  - MC never returns RoR outside [0, 1]
  - Correlation matrix is always symmetric
  - LOO delta is consistent with individual MC runs
```

### Golden Dataset Strategy
Run v1.24 on a set of anonymized test strategies, capture all output metrics as JSON fixtures. v2 integration tests must match within acceptable tolerance (±0.5% for MC results due to RNG differences).

---

## 10. Packaging & Distribution

### Windows .exe (PyInstaller)

```bash
pyinstaller app.py \
  --name "Portfolio Tracker v2" \
  --onefile \
  --add-data "config/default_settings.yaml:config" \
  --icon assets/icon.ico \
  --hidden-import numba \
  --hidden-import streamlit
```

### Auto-Update
Use `pyupdater` or a simple GitHub Releases check on startup:
```python
def check_for_updates(current_version: str) -> str | None:
    """Returns download URL if newer version available, else None."""
    latest = requests.get("https://api.github.com/repos/.../releases/latest").json()
    if latest["tag_name"] > current_version:
        return latest["assets"][0]["browser_download_url"]
    return None
```

---

## 11. Build Order (Phased Delivery)

### Phase 1 — Data Foundation (2-3 weeks)
- [ ] `core/ingestion/folder_scanner.py` — scan MultiWalk folders
- [ ] `core/ingestion/csv_importer.py` — import EquityData + TradeData CSVs
- [ ] `core/ingestion/date_utils.py` — IS/OOS period logic, trading day calendar
- [ ] `core/config.py` — Pydantic config model
- [ ] `core/data_types.py` — domain dataclasses
- [ ] `ui/pages/01_Import.py` — folder selection, import progress, data preview
- [ ] Unit tests for all ingestion code

### Phase 2 — Portfolio & Summary (1-2 weeks)
- [ ] `core/portfolio/aggregator.py` — replaces J_Portfolio_Setup
- [ ] `core/portfolio/summary.py` — 80+ metrics per strategy
- [ ] `ui/pages/02_Portfolio.py` — strategy table with filtering/sorting
- [ ] Integration test: compare metrics to v1.24 output

### Phase 3 — Monte Carlo (1 week)
- [ ] `core/analytics/monte_carlo.py` — vectorized + Numba JIT
- [ ] `ui/pages/03_Monte_Carlo.py`
- [ ] Unit tests with fixed RNG seed
- [ ] Regression test vs v1.24

### Phase 4 — Correlations + Diversification (1-2 weeks)
- [ ] `core/analytics/correlations.py` — 3 modes
- [ ] `core/analytics/diversification.py` — randomized + greedy
- [ ] `ui/pages/04_Correlations.py`
- [ ] `ui/pages/05_Diversification.py`

### Phase 5 — Advanced Analytics (2-3 weeks)
- [ ] `core/analytics/leave_one_out.py`
- [ ] `core/analytics/backtest.py`
- [ ] `core/analytics/margin.py`
- [ ] `ui/pages/06_Leave_One_Out.py`
- [ ] `ui/pages/07_Backtest.py`
- [ ] `ui/pages/08_Margin_Tracking.py`

### Phase 6 — Polish & Distribution (1-2 weeks)
- [ ] `core/licensing/` — hardware-bound JWT keys
- [ ] `core/reporting/excel_export.py` — optional .xlsx export
- [ ] Settings export/import (replaces Q module)
- [ ] PyInstaller packaging
- [ ] Auto-update check
- [ ] End-to-end regression test suite

---

## 12. Key Risks & Mitigations

| Risk | Impact | Mitigation |
|------|--------|-----------|
| MC output doesn't match v1.24 | High — customers notice | Golden dataset regression tests; document RNG difference |
| IS/OOS date logic bugs | High — corrupts all metrics | Unit test every edge case; mirror VBA's ResolveOOSDates exactly |
| Correlation mode 3 (drawdown) definition unclear | Medium | Read VBA L_Correlations in full; add integration test |
| MultiWalk CSV format changes across MW versions | Medium | Version-detect CSV headers; log warnings on parse failure |
| PyInstaller bloat / antivirus false positives | Medium | Code-sign the executable; test on clean Windows VM |
| Numba JIT compile time on first run | Low | Pre-compile with `cache=True`; show spinner in UI |
| Customers uncomfortable leaving Excel | Low | Provide optional .xlsx export from any screen |

---

## 13. Open Questions

1. **MultiWalk DLL licensing** — Do you want v2 to still require MultiWalk, or stand alone?
2. **Pricing model** — One-time purchase with hardware lock, or annual subscription?
3. **Platform** — Windows-only (PyInstaller .exe) or also Mac/Linux?
4. **Feature scope** — Any new features to add in v2 beyond parity?
5. **Data migration** — Do customers need to migrate their Strategies tab config from v1.24?
6. **MultiWalk CSV format** — Which MultiWalk versions need to be supported?
