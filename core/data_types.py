"""
Domain dataclasses for Portfolio Tracker v2.
All data flowing through the system is typed via these classes.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

import pandas as pd


@dataclass
class Strategy:
    """
    Configuration and metadata for a single strategy.
    Sourced from the Strategies tab (user-editable) and
    the Walkforward Details CSV (auto-populated from MultiWalk).
    """
    name: str
    folder: Path
    status: str                         # Live | Paper | Retired | Pass | Not Loaded | etc.
    contracts: int = 1
    symbol: str = ""
    sector: str = ""
    timeframe: str = ""
    type: str = ""                      # Trend | Mean Reversion | Seasonal | etc.
    horizon: str = ""                   # Short | Medium | Long
    other: str = ""
    notes: str = ""

    # Walkforward dates (from Walkforward Details CSV)
    is_start: date | None = None        # IS begin date
    is_end: date | None = None          # IS end date
    oos_start: date | None = None       # OOS begin date
    oos_end: date | None = None         # OOS end date (None = open / ongoing)

    # Computed fields (from Summary calculation)
    incubation_passed_date: date | None = None
    is_max_drawdown: float = 0.0
    expected_annual_return: float = 0.0
    last_date_on_file: date | None = None

    @property
    def is_live(self) -> bool:
        return self.status == "Live"

    @property
    def is_buy_and_hold(self) -> bool:
        return "buy" in self.status.lower() and "hold" in self.status.lower()


@dataclass
class StrategyFolder:
    """Result of folder scanning — a discovered strategy folder with its files."""
    name: str
    path: Path
    equity_csv: Path
    trade_csv: Path | None              # Optional
    walkforward_csv: Path | None        # Optional
    base_folder: Path | None = None     # The root base folder this strategy was found under


@dataclass
class ScanResult:
    """Output of folder_scanner.scan_folders()."""
    strategies: list[StrategyFolder]
    warnings: list[str]                 # Non-fatal issues (missing files, duplicates etc.)
    errors: list[str]                   # Fatal issues (base folder not found etc.)


@dataclass
class ImportedData:
    """
    All data imported from MultiWalk CSV files.
    This is the primary data object passed between ingestion and analytics.
    """
    # Matrix sheets — dates × strategies
    # index = pd.DatetimeIndex, columns = strategy names
    daily_m2m: pd.DataFrame             # Daily mark-to-market PnL (col 2 of EquityData)
    closed_trade_pnl: pd.DataFrame      # Daily closed-trade PnL (col 6 of EquityData)
    in_market_long: pd.DataFrame        # In-market long PnL (col 3 of EquityData)
    in_market_short: pd.DataFrame       # In-market short PnL (col 4 of EquityData)

    # Trade-level data (from TradeData.csv, Exit rows only)
    # columns: [strategy, date, position, pnl, mae, mfe]
    trades: pd.DataFrame

    # Strategy metadata list (merged: Strategies config + Walkforward dates)
    strategies: list[Strategy]

    @property
    def strategy_names(self) -> list[str]:
        return list(self.daily_m2m.columns)

    @property
    def date_range(self) -> tuple[date, date]:
        idx = self.daily_m2m.index
        return idx.min().date(), idx.max().date()


@dataclass
class PortfolioData:
    """
    Aggregated portfolio data — the ImportedData filtered to active strategies
    and combined into portfolio-level series.
    """
    strategies: list[Strategy]
    daily_pnl: pd.DataFrame             # Active strategies daily M2M
    closed_trades: pd.DataFrame         # Active strategies closed trade PnL
    summary_metrics: pd.DataFrame       # index=strategy name, columns=80+ metrics


@dataclass
class MCResult:
    starting_equity: float
    expected_profit: float
    risk_of_ruin: float
    max_drawdown_pct: float
    sharpe_ratio: float
    return_to_drawdown: float
    scenarios_df: pd.DataFrame | None = None  # Full scenario array if requested
