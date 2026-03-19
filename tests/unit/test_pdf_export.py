"""
Unit tests for core.reporting.pdf_export.

Validates that the PDF is generated without errors and the output is a
recognisable PDF byte stream (starts with %PDF-).
"""
from __future__ import annotations

import io
import numpy as np
import pandas as pd
import pytest

from core.data_types import PortfolioData, MCResult, Strategy
from core.config import AppConfig
from pathlib import Path


def _make_portfolio(n_strats=3, n_days=126):
    dates = pd.date_range("2022-01-03", periods=n_days, freq="B")
    names = [f"S_{i}" for i in range(n_strats)]
    rng   = np.random.default_rng(1)
    pnl   = pd.DataFrame(rng.standard_normal((n_days, n_strats)) * 300, index=dates, columns=names)
    strategies = [
        Strategy(n, Path("/fake"), "Live", 1, f"ES{i}", f"Sector{i%3}")
        for i, n in enumerate(names)
    ]
    summary = pd.DataFrame(
        {"profit_since_oos_start": [5000.0] * n_strats},
        index=names,
    )
    return PortfolioData(
        strategies=strategies,
        daily_pnl=pnl,
        closed_trades=pd.DataFrame(),
        summary_metrics=summary,
    )


def _make_mc():
    return MCResult(
        starting_equity=80_000.0,
        expected_profit=15_000.0,
        risk_of_ruin=0.07,
        max_drawdown_pct=0.22,
        sharpe_ratio=1.1,
        return_to_drawdown=2.2,
    )


class TestExportPortfolioPdf:
    def test_returns_bytes(self):
        from core.reporting.pdf_export import export_portfolio_pdf
        portfolio = _make_portfolio()
        result = export_portfolio_pdf(portfolio, AppConfig())
        assert isinstance(result, bytes)
        assert len(result) > 500

    def test_output_is_pdf(self):
        from core.reporting.pdf_export import export_portfolio_pdf
        portfolio = _make_portfolio()
        pdf = export_portfolio_pdf(portfolio, AppConfig())
        assert pdf[:5] == b"%PDF-"

    def test_with_mc_result(self):
        from core.reporting.pdf_export import export_portfolio_pdf
        portfolio = _make_portfolio()
        pdf = export_portfolio_pdf(portfolio, AppConfig(), mc_result=_make_mc())
        assert pdf[:5] == b"%PDF-"
        assert len(pdf) > 1000

    def test_empty_portfolio(self):
        from core.reporting.pdf_export import export_portfolio_pdf
        portfolio = PortfolioData(
            strategies=[],
            daily_pnl=pd.DataFrame(),
            closed_trades=pd.DataFrame(),
            summary_metrics=pd.DataFrame(),
        )
        pdf = export_portfolio_pdf(portfolio, AppConfig())
        assert pdf[:5] == b"%PDF-"

    def test_many_strategies(self):
        """Should paginate without error when strategies exceed one page."""
        from core.reporting.pdf_export import export_portfolio_pdf
        portfolio = _make_portfolio(n_strats=50, n_days=252)
        pdf = export_portfolio_pdf(portfolio, AppConfig())
        assert pdf[:5] == b"%PDF-"


class TestPdfFilenameHelper:
    def test_ends_with_pdf(self):
        from core.reporting.pdf_export import pdf_export_filename
        assert pdf_export_filename().endswith(".pdf")

    def test_contains_date(self):
        from core.reporting.pdf_export import pdf_export_filename
        import re
        assert re.search(r"\d{4}-\d{2}-\d{2}", pdf_export_filename())
