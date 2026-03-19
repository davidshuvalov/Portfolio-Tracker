"""
Session-scoped fixtures for v1.24 regression integration tests.

All fixtures auto-skip if pt_fixtures/ is absent (fresh checkout without
real data).  When the fixtures are present the full suite runs and validates
v2's analytics against v1.24's exported reference data.

Fixture layout in pt_fixtures/ (repo root):
    PortfolioDailyM2M.csv   — per-strategy daily M2M, already contract-scaled
                              (90 live strategies × all trading days)
    TotalPortfolioM2M.csv   — aggregated daily portfolio totals (v1.24 reference)
    DailyM2MEquity.csv      — raw 1-contract daily M2M for all 430 strategies
    ClosedTradePNL.csv      — raw 1-contract daily closed PnL for all 430 strategies
    Strategies.csv          — strategy config: name, status, contracts, symbol, …
    Portfolio.csv           — portfolio-level metrics per live strategy (v1.24 output)
    Summary.csv             — per-strategy summary metrics (v1.24 output)
    LatestPositionData.csv  — most-recent position per strategy
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
import pytest

# Path to fixture directory (repo root / pt_fixtures)
PT_FIXTURES = Path(__file__).parent.parent.parent.parent / "pt_fixtures"
SKIP_MSG = "pt_fixtures/ not found — skipping integration tests (run from real dataset)"


def _require_fixtures():
    """Raise pytest.skip if the fixture directory is missing or empty."""
    if not PT_FIXTURES.is_dir():
        pytest.skip(SKIP_MSG)
    sentinel = PT_FIXTURES / "PortfolioDailyM2M.csv"
    if not sentinel.exists():
        pytest.skip(SKIP_MSG)


# ── Raw CSV loaders ───────────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def v1_portfolio_daily_m2m() -> pd.DataFrame:
    """
    PortfolioDailyM2M.csv — per-strategy daily M2M already scaled by contracts.
    index=DatetimeIndex, columns=strategy names.
    """
    _require_fixtures()
    df = pd.read_csv(
        PT_FIXTURES / "PortfolioDailyM2M.csv",
        index_col=0,
        parse_dates=True,
    )
    df.index = pd.to_datetime(df.index)
    return df


@pytest.fixture(scope="session")
def v1_total_m2m() -> pd.DataFrame:
    """
    TotalPortfolioM2M.csv — daily portfolio totals from v1.24.
    Columns include: Total Daily Profit, Total Cumulative P/L, Total Drawdown.
    """
    _require_fixtures()
    df = pd.read_csv(PT_FIXTURES / "TotalPortfolioM2M.csv")
    df["Date"] = pd.to_datetime(df["Date"])
    return df.set_index("Date")


@pytest.fixture(scope="session")
def v1_daily_m2m_raw() -> pd.DataFrame:
    """
    DailyM2MEquity.csv — raw 1-contract daily M2M for all 430 strategies.
    index=DatetimeIndex (DD/MM/YYYY), columns=strategy names.
    """
    _require_fixtures()
    df = pd.read_csv(
        PT_FIXTURES / "DailyM2MEquity.csv",
        index_col=0,
        parse_dates=True,
        dayfirst=True,
    )
    df.index = pd.to_datetime(df.index)
    return df


@pytest.fixture(scope="session")
def v1_strategies() -> pd.DataFrame:
    """
    Strategies.csv — strategy config from v1.24.
    Columns: Strategy Number, Status, Strategy Name, Contracts, Symbol, …
    """
    _require_fixtures()
    return pd.read_csv(PT_FIXTURES / "Strategies.csv")


@pytest.fixture(scope="session")
def v1_summary() -> pd.DataFrame:
    """
    Summary.csv — per-strategy summary metrics from v1.24.
    Returns DataFrame with Strategy Name as index.
    """
    _require_fixtures()
    df = pd.read_csv(PT_FIXTURES / "Summary.csv")
    if "Strategy Name" in df.columns:
        df = df.set_index("Strategy Name")
    return df


@pytest.fixture(scope="session")
def v1_portfolio() -> pd.DataFrame:
    """
    Portfolio.csv — portfolio-level per-strategy metrics from v1.24.
    Returns DataFrame with Strategy Name as index.
    """
    _require_fixtures()
    df = pd.read_csv(PT_FIXTURES / "Portfolio.csv")
    if "Strategy Name" in df.columns:
        df = df.set_index("Strategy Name")
    return df


@pytest.fixture(scope="session")
def v1_latest_positions() -> pd.DataFrame:
    """LatestPositionData.csv — most-recent position per strategy."""
    _require_fixtures()
    return pd.read_csv(PT_FIXTURES / "LatestPositionData.csv")


# ── v2 pipeline fixtures ──────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def imported_from_portfolio_m2m(v1_portfolio_daily_m2m):
    """
    Build an ImportedData object directly from PortfolioDailyM2M.csv.

    Since PortfolioDailyM2M is already contract-scaled, we treat every
    strategy as contracts=1 and status='Live'.  This lets us exercise the
    full v2 aggregation pipeline against the v1.24 reference data without
    needing per-strategy EquityData.csv folders.
    """
    import sys
    sys.path.insert(0, str(Path(__file__).parent.parent.parent))

    from pathlib import Path as P
    from core.data_types import ImportedData, Strategy

    port = v1_portfolio_daily_m2m
    dummy = pd.DataFrame(0.0, index=port.index, columns=port.columns)
    trades = pd.DataFrame(columns=["strategy", "date", "position", "pnl", "mae", "mfe"])
    strategies = [
        Strategy(name=col, folder=P("."), status="Live", contracts=1)
        for col in port.columns
    ]
    return ImportedData(
        daily_m2m=port,
        closed_trade_pnl=dummy,
        in_market_long=dummy,
        in_market_short=dummy,
        trades=trades,
        strategies=strategies,
    )


@pytest.fixture(scope="session")
def v2_portfolio(imported_from_portfolio_m2m):
    """
    PortfolioData built from the v1.24 PortfolioDailyM2M fixture via
    v2's build_portfolio() — all strategies treated as Live, contracts=1.
    """
    import sys
    sys.path.insert(0, str(Path(__file__).parent.parent.parent))

    from core.portfolio.aggregator import build_portfolio

    imported = imported_from_portfolio_m2m
    config = [
        {"name": s.name, "status": "Live", "contracts": 1, "symbol": "", "sector": ""}
        for s in imported.strategies
    ]
    return build_portfolio(imported, config, live_status="Live")


@pytest.fixture(scope="session")
def v2_total_pnl(v2_portfolio):
    """Total portfolio daily PnL Series from v2 (sum of all active strategies)."""
    import sys
    sys.path.insert(0, str(Path(__file__).parent.parent.parent))

    from core.portfolio.aggregator import portfolio_total_pnl

    return portfolio_total_pnl(v2_portfolio)


@pytest.fixture(scope="session")
def v2_correlation_matrix(v2_portfolio):
    """Pearson correlation matrix from v2 computed on the portfolio daily PnL."""
    import sys
    sys.path.insert(0, str(Path(__file__).parent.parent.parent))

    from core.analytics.correlations import CorrelationMode, compute_correlation_matrix

    return compute_correlation_matrix(v2_portfolio.daily_pnl, CorrelationMode.NORMAL)
