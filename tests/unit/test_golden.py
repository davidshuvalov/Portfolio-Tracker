"""
Golden fixture regression tests.

These tests only run when golden fixture JSON files exist in tests/fixtures/.
If the files don't exist (fresh checkout before first capture), the tests
are skipped automatically.

To capture fixtures:
    python tests/fixtures/capture_golden.py --folder /path/to/MultiWalk

To re-capture after an intentional change:
    python tests/fixtures/capture_golden.py --folder /path/to/MultiWalk
    git diff tests/fixtures/golden_*.json  # review changes
    git add tests/fixtures/golden_*.json && git commit -m "Update golden fixtures"
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

FIXTURE_DIR = Path(__file__).parent.parent / "fixtures"
TOLERANCE   = 0.005   # 0.5% tolerance for floating-point values


def _load(name: str) -> dict | None:
    path = FIXTURE_DIR / f"{name}.json"
    if not path.exists():
        return None
    return json.loads(path.read_text())


def _approx(actual: float, expected: float, tol: float = TOLERANCE) -> bool:
    if expected == 0:
        return abs(actual) < 1e-6
    return abs(actual - expected) / abs(expected) <= tol


# ── Import fixtures ───────────────────────────────────────────────────────────

class TestGoldenImport:
    @pytest.fixture(autouse=True)
    def _load_fixture(self):
        self.golden = _load("golden_import")
        if self.golden is None:
            pytest.skip("golden_import.json not captured yet")

    def test_strategy_count(self, shared_imported):
        assert len(shared_imported.strategy_names) == self.golden["n_strategies"]

    def test_strategy_names(self, shared_imported):
        assert sorted(shared_imported.strategy_names) == sorted(self.golden["strategy_names"])

    def test_date_range(self, shared_imported):
        start, end = shared_imported.date_range
        assert str(start) == self.golden["date_start"]
        assert str(end)   == self.golden["date_end"]

    def test_trade_count(self, shared_imported):
        assert len(shared_imported.trades) == self.golden["n_trades"]


# ── Summary fixtures ──────────────────────────────────────────────────────────

class TestGoldenSummary:
    @pytest.fixture(autouse=True)
    def _load_fixture(self):
        self.golden = _load("golden_summary")
        if self.golden is None:
            pytest.skip("golden_summary.json not captured yet")

    def test_all_strategies_present(self, shared_summary):
        for name in self.golden:
            assert name in shared_summary.index, f"Strategy {name} missing from summary"

    def test_expected_annual_profit(self, shared_summary):
        for name, expected_row in self.golden.items():
            if "expected_annual_profit" not in expected_row:
                continue
            exp_val = expected_row["expected_annual_profit"]
            if exp_val is None:
                continue
            actual = shared_summary.loc[name, "expected_annual_profit"]
            if actual != actual:  # NaN
                continue
            assert _approx(float(actual), float(exp_val)), (
                f"{name}: expected_annual_profit {actual:.2f} vs golden {exp_val:.2f}"
            )

    def test_sharpe_isoos(self, shared_summary):
        for name, expected_row in self.golden.items():
            if "sharpe_isoos" not in expected_row or expected_row["sharpe_isoos"] is None:
                continue
            actual = shared_summary.loc[name, "sharpe_isoos"]
            if actual != actual:
                continue
            golden_val = expected_row["sharpe_isoos"]
            assert _approx(float(actual), float(golden_val)), (
                f"{name}: sharpe_isoos {actual:.3f} vs golden {golden_val:.3f}"
            )


# ── Portfolio fixtures ────────────────────────────────────────────────────────

class TestGoldenPortfolio:
    @pytest.fixture(autouse=True)
    def _load_fixture(self):
        self.golden = _load("golden_portfolio")
        if self.golden is None:
            pytest.skip("golden_portfolio.json not captured yet")

    def test_active_strategy_count(self, shared_portfolio):
        assert len(shared_portfolio.strategies) == self.golden["n_active_strategies"]

    def test_total_pnl(self, shared_portfolio):
        actual = float(shared_portfolio.daily_pnl.sum().sum())
        assert _approx(actual, self.golden["total_pnl"])


# ── conftest hooks (shared fixtures) ─────────────────────────────────────────
# shared_imported / shared_summary / shared_portfolio are defined in
# conftest.py at the tests/unit level if a real dataset is available.
# If not defined, any test that uses them will error; the autouse fixture
# above skips those tests when the JSON is missing, so no false failures.
