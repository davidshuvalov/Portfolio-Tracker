"""
Tests for core/ingestion/csv_importer.py

Mirrors the exact behaviour of VBA D_Import_Data.bas OptimizeDataProcessing.
"""

import pytest
from pathlib import Path
from datetime import date
import textwrap

import sys
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

import pandas as pd
import numpy as np

from core.ingestion.csv_importer import (
    import_all,
    _read_equity_csv,
    _read_trade_csv,
    _to_float,
    _is_numeric,
    EQUITY_COL_DATE, EQUITY_COL_M2M, EQUITY_COL_LONG,
    EQUITY_COL_SHORT, EQUITY_COL_CLOSED,
    TRADE_COL_DATE, TRADE_COL_TYPE, TRADE_COL_POSITION,
    TRADE_COL_PNL, TRADE_COL_MAE, TRADE_COL_MFE,
)
from core.data_types import StrategyFolder


# ── CSV content builders ──────────────────────────────────────────────────────

def _equity_csv(rows: list[tuple], header=True) -> str:
    """Build EquityData.csv content (6 columns, VBA layout)."""
    lines = []
    if header:
        lines.append("Date,DailyM2M,Long,Short,Unused,Closed")
    for row in rows:
        lines.append(",".join(str(v) for v in row))
    return "\n".join(lines) + "\n"


def _trade_csv(rows: list[tuple], header=True) -> str:
    """Build TradeData.csv content (8 columns minimum)."""
    lines = []
    if header:
        lines.append("Date,Col2,Col3,Type,Position,PNL,MAE,MFE")
    for row in rows:
        lines.append(",".join(str(v) for v in row))
    return "\n".join(lines) + "\n"


def _make_strategy_folder(
    tmp_path: Path,
    name: str,
    equity_content: str,
    trade_content: str | None = None,
) -> StrategyFolder:
    """Write CSV files and return a StrategyFolder pointing to them."""
    wf_dir = tmp_path / name / "Walkforward Files"
    wf_dir.mkdir(parents=True)

    equity_csv = wf_dir / f"{name} EquityData.csv"
    equity_csv.write_text(equity_content)

    trade_csv = None
    if trade_content is not None:
        trade_csv = wf_dir / f"{name} TradeData.csv"
        trade_csv.write_text(trade_content)

    return StrategyFolder(
        name=name,
        path=tmp_path / name,
        equity_csv=equity_csv,
        trade_csv=trade_csv,
        walkforward_csv=None,
    )


# ── _to_float / _is_numeric ───────────────────────────────────────────────────

class TestHelpers:
    def test_to_float_plain_number(self):
        assert _to_float("123.45") == pytest.approx(123.45)

    def test_to_float_comma_thousands(self):
        assert _to_float("1,234.56") == pytest.approx(1234.56)

    def test_to_float_whitespace(self):
        assert _to_float("  42  ") == pytest.approx(42.0)

    def test_to_float_invalid_returns_zero(self):
        assert _to_float("not_a_number") == 0.0

    def test_to_float_empty_returns_zero(self):
        assert _to_float("") == 0.0

    def test_to_float_none_returns_zero(self):
        assert _to_float(None) == 0.0

    def test_is_numeric_true_for_number(self):
        assert _is_numeric("123.45")

    def test_is_numeric_false_for_text(self):
        assert not _is_numeric("Date")

    def test_is_numeric_false_for_empty(self):
        assert not _is_numeric("")


# ── _read_equity_csv ──────────────────────────────────────────────────────────

class TestReadEquityCsv:
    def test_reads_basic_dmy_csv(self, tmp_path):
        content = _equity_csv([
            ("15/06/2023", 100.0, 1, 0, 0, 50.0),
            ("16/06/2023", 200.0, 1, 0, 0, 75.0),
        ])
        csv_path = tmp_path / "test.csv"
        csv_path.write_text(content)
        warnings = []
        result = _read_equity_csv(csv_path, "StratA", "DMY", warnings)
        assert result is not None
        assert len(result.dates) == 2
        assert result.dates[0] == date(2023, 6, 15)
        assert result.m2m[0] == pytest.approx(100.0)
        assert result.closed[0] == pytest.approx(50.0)

    def test_reads_mdy_csv(self, tmp_path):
        content = _equity_csv([
            ("06/15/2023", 100.0, 1, 0, 0, 50.0),
        ])
        csv_path = tmp_path / "test.csv"
        csv_path.write_text(content)
        warnings = []
        result = _read_equity_csv(csv_path, "StratA", "MDY", warnings)
        assert result is not None
        assert result.dates[0] == date(2023, 6, 15)

    def test_skips_header_row(self, tmp_path):
        content = _equity_csv([("15/06/2023", 100.0, 1, 0, 0, 50.0)], header=True)
        csv_path = tmp_path / "test.csv"
        csv_path.write_text(content)
        warnings = []
        result = _read_equity_csv(csv_path, "StratA", "DMY", warnings)
        assert result is not None
        assert len(result.dates) == 1

    def test_no_header_row(self, tmp_path):
        content = _equity_csv([("15/06/2023", 100.0, 1, 0, 0, 50.0)], header=False)
        csv_path = tmp_path / "test.csv"
        csv_path.write_text(content)
        warnings = []
        result = _read_equity_csv(csv_path, "StratA", "DMY", warnings)
        assert result is not None
        assert len(result.dates) == 1

    def test_reads_all_columns(self, tmp_path):
        content = _equity_csv([("15/06/2023", 100.0, 1.0, 0.5, 0, 75.0)])
        csv_path = tmp_path / "test.csv"
        csv_path.write_text(content)
        warnings = []
        result = _read_equity_csv(csv_path, "StratA", "DMY", warnings)
        assert result.m2m[0] == pytest.approx(100.0)
        assert result.long[0] == pytest.approx(1.0)
        assert result.short[0] == pytest.approx(0.5)
        assert result.closed[0] == pytest.approx(75.0)

    def test_invalid_date_row_skipped(self, tmp_path):
        content = _equity_csv([
            ("not-a-date", 100.0, 1, 0, 0, 50.0),
            ("15/06/2023", 200.0, 1, 0, 0, 75.0),
        ])
        csv_path = tmp_path / "test.csv"
        csv_path.write_text(content)
        warnings = []
        result = _read_equity_csv(csv_path, "StratA", "DMY", warnings)
        # Only the valid row is kept (invalid date row skipped)
        assert result is not None
        assert len(result.dates) == 1
        assert result.dates[0] == date(2023, 6, 15)

    def test_nonexistent_file_returns_none_with_warning(self, tmp_path):
        warnings = []
        result = _read_equity_csv(tmp_path / "missing.csv", "StratA", "DMY", warnings)
        assert result is None
        assert len(warnings) == 1

    def test_empty_file_returns_none_with_warning(self, tmp_path):
        csv_path = tmp_path / "empty.csv"
        csv_path.write_text("")
        warnings = []
        result = _read_equity_csv(csv_path, "StratA", "DMY", warnings)
        assert result is None

    def test_negative_pnl_values(self, tmp_path):
        content = _equity_csv([("15/06/2023", -100.0, 0, 1, 0, -50.0)])
        csv_path = tmp_path / "test.csv"
        csv_path.write_text(content)
        warnings = []
        result = _read_equity_csv(csv_path, "StratA", "DMY", warnings)
        assert result.m2m[0] == pytest.approx(-100.0)
        assert result.closed[0] == pytest.approx(-50.0)

    def test_to_float_strips_commas_from_thousands(self):
        # Verify _to_float handles comma-formatted numbers (e.g. "1,500.00")
        # Note: comma-separated CSV can't contain unquoted commas in values;
        # this tests the helper in isolation.
        assert _to_float("1,500.00") == pytest.approx(1500.0)
        assert _to_float("2,000.00") == pytest.approx(2000.0)


# ── _read_trade_csv ───────────────────────────────────────────────────────────

class TestReadTradeCsv:
    def _make_trade_csv(self, tmp_path, content):
        p = tmp_path / "trades.csv"
        p.write_text(content)
        return p

    def test_reads_exit_rows_only(self, tmp_path):
        content = _trade_csv([
            ("15/06/2023", "", "", "Entry", "Long",  -100, 0, 0),
            ("15/06/2023", "", "", "Exit",  "Long",   250, 50, 300),
        ])
        p = self._make_trade_csv(tmp_path, content)
        warnings = []
        df = _read_trade_csv(p, "StratA", "DMY", warnings)
        assert df is not None
        assert len(df) == 1
        assert df.iloc[0]["pnl"] == pytest.approx(250.0)

    def test_long_position_normalised_to_L(self, tmp_path):
        content = _trade_csv([
            ("15/06/2023", "", "", "Exit", "Long", 250, 50, 300),
        ])
        p = self._make_trade_csv(tmp_path, content)
        warnings = []
        df = _read_trade_csv(p, "StratA", "DMY", warnings)
        assert df.iloc[0]["position"] == "L"

    def test_short_position_normalised_to_S(self, tmp_path):
        content = _trade_csv([
            ("15/06/2023", "", "", "Exit", "Short", 250, 50, 300),
        ])
        p = self._make_trade_csv(tmp_path, content)
        warnings = []
        df = _read_trade_csv(p, "StratA", "DMY", warnings)
        assert df.iloc[0]["position"] == "S"

    def test_output_columns(self, tmp_path):
        content = _trade_csv([
            ("15/06/2023", "", "", "Exit", "Long", 250.0, 50.0, 300.0),
        ])
        p = self._make_trade_csv(tmp_path, content)
        warnings = []
        df = _read_trade_csv(p, "StratA", "DMY", warnings)
        assert set(df.columns) == {"strategy", "date", "position", "pnl", "mae", "mfe"}
        assert df.iloc[0]["strategy"] == "StratA"
        assert df.iloc[0]["mae"] == pytest.approx(50.0)
        assert df.iloc[0]["mfe"] == pytest.approx(300.0)

    def test_no_exit_rows_returns_none(self, tmp_path):
        content = _trade_csv([
            ("15/06/2023", "", "", "Entry", "Long", -100, 0, 0),
        ])
        p = self._make_trade_csv(tmp_path, content)
        warnings = []
        df = _read_trade_csv(p, "StratA", "DMY", warnings)
        assert df is None

    def test_date_column_is_datetime(self, tmp_path):
        content = _trade_csv([
            ("15/06/2023", "", "", "Exit", "Long", 250.0, 50.0, 300.0),
        ])
        p = self._make_trade_csv(tmp_path, content)
        warnings = []
        df = _read_trade_csv(p, "StratA", "DMY", warnings)
        assert pd.api.types.is_datetime64_any_dtype(df["date"])

    def test_multiple_exit_rows(self, tmp_path):
        content = _trade_csv([
            ("15/06/2023", "", "", "Exit", "Long",  250.0, 50.0, 300.0),
            ("20/06/2023", "", "", "Exit", "Short", -75.0, 20.0, 100.0),
        ])
        p = self._make_trade_csv(tmp_path, content)
        warnings = []
        df = _read_trade_csv(p, "StratA", "DMY", warnings)
        assert len(df) == 2


# ── import_all ────────────────────────────────────────────────────────────────

class TestImportAll:
    def test_single_strategy_basic(self, tmp_path):
        content = _equity_csv([
            ("15/06/2023", 100.0, 1, 0, 0, 50.0),
            ("16/06/2023", 200.0, 1, 0, 0, 75.0),
        ])
        sf = _make_strategy_folder(tmp_path, "StratA", content)
        imported, warnings = import_all([sf], date_format="DMY")
        assert "StratA" in imported.daily_m2m.columns
        assert len(imported.daily_m2m) == 2

    def test_two_strategies_aligned(self, tmp_path):
        """Strategies with different date ranges → union index, zeros for missing."""
        content_a = _equity_csv([
            ("15/06/2023", 100.0, 1, 0, 0, 50.0),
            ("16/06/2023", 200.0, 1, 0, 0, 75.0),
        ])
        content_b = _equity_csv([
            ("16/06/2023", 300.0, 1, 0, 0, 100.0),
            ("17/06/2023", 400.0, 1, 0, 0, 125.0),
        ])
        sf_a = _make_strategy_folder(tmp_path, "StratA", content_a)
        sf_b = _make_strategy_folder(tmp_path, "StratB", content_b)

        imported, warnings = import_all([sf_a, sf_b], date_format="DMY")

        # Union of 3 dates
        assert len(imported.daily_m2m) == 3

        # Missing dates filled with zero
        # StratA on 17/06/2023 should be 0
        ts_17 = pd.Timestamp("2023-06-17")
        assert imported.daily_m2m.loc[ts_17, "StratA"] == pytest.approx(0.0)

        # StratB on 15/06/2023 should be 0
        ts_15 = pd.Timestamp("2023-06-15")
        assert imported.daily_m2m.loc[ts_15, "StratB"] == pytest.approx(0.0)

    def test_cutoff_filters_dates(self, tmp_path):
        content = _equity_csv([
            ("15/06/2023", 100.0, 1, 0, 0, 50.0),
            ("16/06/2023", 200.0, 1, 0, 0, 75.0),
            ("17/06/2023", 300.0, 1, 0, 0, 100.0),
        ])
        sf = _make_strategy_folder(tmp_path, "StratA", content)
        imported, _ = import_all(
            [sf],
            date_format="DMY",
            use_cutoff=True,
            cutoff_date=date(2023, 6, 16),
        )
        assert len(imported.daily_m2m) == 2
        assert pd.Timestamp("2023-06-17") not in imported.daily_m2m.index

    def test_no_cutoff_includes_all_dates(self, tmp_path):
        content = _equity_csv([
            ("15/06/2023", 100.0, 1, 0, 0, 50.0),
            ("17/06/2023", 300.0, 1, 0, 0, 100.0),
        ])
        sf = _make_strategy_folder(tmp_path, "StratA", content)
        imported, _ = import_all([sf], date_format="DMY", use_cutoff=False)
        assert len(imported.daily_m2m) == 2

    def test_imported_data_has_all_dataframes(self, tmp_path):
        content = _equity_csv([("15/06/2023", 100.0, 1, 0, 0, 50.0)])
        sf = _make_strategy_folder(tmp_path, "StratA", content)
        imported, _ = import_all([sf], date_format="DMY")
        assert imported.daily_m2m is not None
        assert imported.closed_trade_pnl is not None
        assert imported.in_market_long is not None
        assert imported.in_market_short is not None
        assert imported.trades is not None

    def test_date_range_property(self, tmp_path):
        content = _equity_csv([
            ("15/06/2023", 100.0, 1, 0, 0, 50.0),
            ("20/06/2023", 200.0, 1, 0, 0, 75.0),
        ])
        sf = _make_strategy_folder(tmp_path, "StratA", content)
        imported, _ = import_all([sf], date_format="DMY")
        start, end = imported.date_range
        assert str(start)[:10] == "2023-06-15"
        assert str(end)[:10] == "2023-06-20"

    def test_strategy_names_property(self, tmp_path):
        content = _equity_csv([("15/06/2023", 100.0, 1, 0, 0, 50.0)])
        sf = _make_strategy_folder(tmp_path, "StratA", content)
        imported, _ = import_all([sf], date_format="DMY")
        assert "StratA" in imported.strategy_names

    def test_no_valid_equity_csvs_raises(self, tmp_path):
        # Empty CSV → import_all should raise
        csv_path = tmp_path / "StratA" / "Walkforward Files" / "StratA EquityData.csv"
        csv_path.parent.mkdir(parents=True)
        csv_path.write_text("")
        sf = StrategyFolder(
            name="StratA",
            path=tmp_path / "StratA",
            equity_csv=csv_path,
            trade_csv=None,
            walkforward_csv=None,
        )
        with pytest.raises(ValueError, match="No valid"):
            import_all([sf], date_format="DMY")

    def test_trades_populated_when_trade_csv_present(self, tmp_path):
        equity = _equity_csv([("15/06/2023", 100.0, 1, 0, 0, 50.0)])
        trade = _trade_csv([
            ("15/06/2023", "", "", "Exit", "Long", 250.0, 50.0, 300.0),
        ])
        sf = _make_strategy_folder(tmp_path, "StratA", equity, trade)
        imported, _ = import_all([sf], date_format="DMY")
        assert len(imported.trades) == 1
        assert imported.trades.iloc[0]["strategy"] == "StratA"

    def test_trades_empty_when_no_trade_csv(self, tmp_path):
        equity = _equity_csv([("15/06/2023", 100.0, 1, 0, 0, 50.0)])
        sf = _make_strategy_folder(tmp_path, "StratA", equity, trade_content=None)
        imported, _ = import_all([sf], date_format="DMY")
        assert imported.trades.empty

    def test_index_is_datetimeindex(self, tmp_path):
        content = _equity_csv([("15/06/2023", 100.0, 1, 0, 0, 50.0)])
        sf = _make_strategy_folder(tmp_path, "StratA", content)
        imported, _ = import_all([sf], date_format="DMY")
        assert isinstance(imported.daily_m2m.index, pd.DatetimeIndex)

    def test_values_match_csv_data(self, tmp_path):
        content = _equity_csv([
            ("15/06/2023", 123.45, 1, 0, 0, 67.89),
        ])
        sf = _make_strategy_folder(tmp_path, "StratA", content)
        imported, _ = import_all([sf], date_format="DMY")
        ts = pd.Timestamp("2023-06-15")
        assert imported.daily_m2m.loc[ts, "StratA"] == pytest.approx(123.45)
        assert imported.closed_trade_pnl.loc[ts, "StratA"] == pytest.approx(67.89)

    def test_in_market_columns_populated(self, tmp_path):
        content = _equity_csv([
            ("15/06/2023", 100.0, 1.0, 0.5, 0, 50.0),
        ])
        sf = _make_strategy_folder(tmp_path, "StratA", content)
        imported, _ = import_all([sf], date_format="DMY")
        ts = pd.Timestamp("2023-06-15")
        assert imported.in_market_long.loc[ts, "StratA"] == pytest.approx(1.0)
        assert imported.in_market_short.loc[ts, "StratA"] == pytest.approx(0.5)

    def test_dates_sorted_ascending(self, tmp_path):
        # Provide dates out of order in CSV
        content = _equity_csv([
            ("20/06/2023", 200.0, 1, 0, 0, 75.0),
            ("15/06/2023", 100.0, 1, 0, 0, 50.0),
        ])
        sf = _make_strategy_folder(tmp_path, "StratA", content)
        imported, _ = import_all([sf], date_format="DMY")
        assert imported.daily_m2m.index[0] < imported.daily_m2m.index[1]
