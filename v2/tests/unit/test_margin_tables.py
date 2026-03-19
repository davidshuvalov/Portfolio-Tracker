"""
Unit tests for xlsb_importer MarginTables logic.

Tests the MarginTables dataclass, get_margin(), resolve_for_symbols(),
save/load round-trip, and the raw-data parsing helpers (_float, _avg).
The xlsb file itself is not used — all tests use synthetic data.
"""

from __future__ import annotations

import tempfile
from pathlib import Path

import pytest
import yaml

from core.ingestion.xlsb_importer import (
    MarginEntry,
    MarginTables,
    MARGIN_TABLES_FILE,
    _avg,
    _float,
    load_margin_tables,
    save_margin_tables,
)


# ── Fixtures ──────────────────────────────────────────────────────────────────

def _sample_tables() -> MarginTables:
    return MarginTables(
        ts={
            "ES": MarginEntry(description="E-Mini S&P 500", initial=22545.0, maintenance=20495.5),
            "NQ": MarginEntry(description="E-Mini Nasdaq 100", initial=32920.0, maintenance=29927.5),
            "NK": MarginEntry(description="Nikkei 225 USD", initial=16556.0, maintenance=15051.0),
            "CL": MarginEntry(description="Crude Oil", initial=6428.0, maintenance=5843.5),
        },
        ib={
            "ES":  MarginEntry(initial=15000.0, maintenance=12000.0),
            "NQ":  MarginEntry(initial=25000.0, maintenance=20000.0),
            "NKD": MarginEntry(initial=17000.0, maintenance=13500.0),
            "CL":  MarginEntry(initial=7000.0,  maintenance=5500.0),
        },
        lookup={
            "ES": "ES",
            "NQ": "NQ",
            "NK": "NKD",   # TS NK → IB NKD
            "CL": "CL",
            "FV": "ZF",    # TS FV → IB ZF (not in ib dict, tests fallback)
        },
        sector_lookup={
            "ES": "Index",
            "NQ": "Index",
            "YM": "Index",
            "CL": "Energy",
            "NG": "Energy",
            "GC": "Metals",
            "SI": "Metals",
            "VX": "Volatility",
            "TY": "Interest Rate",
            "EC": "Currencies",
        },
    )


# ── MarginEntry defaults ──────────────────────────────────────────────────────

class TestMarginEntry:
    def test_defaults(self):
        e = MarginEntry()
        assert e.description == ""
        assert e.initial == 0.0
        assert e.maintenance == 0.0

    def test_values(self):
        e = MarginEntry(description="ES", initial=22000.0, maintenance=18000.0)
        assert e.initial == 22000.0
        assert e.maintenance == 18000.0


# ── MarginTables.get_margin ───────────────────────────────────────────────────

class TestMarginTablesGetMargin:
    def test_ts_maintenance(self):
        tables = _sample_tables()
        val = tables.get_margin("ES", "TradeStation", "Maintenance")
        assert val == pytest.approx(20495.5)

    def test_ts_initial(self):
        tables = _sample_tables()
        val = tables.get_margin("ES", "TradeStation", "Initial")
        assert val == pytest.approx(22545.0)

    def test_ib_maintenance_via_lookup(self):
        """NK (TS) maps to NKD (IB) via symbol_lookup."""
        tables = _sample_tables()
        val = tables.get_margin("NK", "InteractiveBrokers", "Maintenance")
        assert val == pytest.approx(13500.0)

    def test_ib_initial_via_lookup(self):
        tables = _sample_tables()
        val = tables.get_margin("NK", "InteractiveBrokers", "Initial")
        assert val == pytest.approx(17000.0)

    def test_ib_same_symbol(self):
        """ES maps to ES (same symbol) via lookup."""
        tables = _sample_tables()
        val = tables.get_margin("ES", "InteractiveBrokers", "Maintenance")
        assert val == pytest.approx(12000.0)

    def test_ib_fallback_when_ts_symbol_not_in_lookup(self):
        """Symbol not in lookup dict: fall back to using ts_symbol directly as ib_symbol."""
        tables = _sample_tables()
        # CL is in ib dict, and lookup["CL"] = "CL"
        val = tables.get_margin("CL", "InteractiveBrokers", "Maintenance")
        assert val == pytest.approx(5500.0)

    def test_returns_none_for_unknown_ts_symbol(self):
        tables = _sample_tables()
        assert tables.get_margin("UNKNOWN", "TradeStation", "Maintenance") is None

    def test_returns_none_for_unknown_ib_symbol(self):
        """FV → ZF via lookup, but ZF is not in ib dict."""
        tables = _sample_tables()
        assert tables.get_margin("FV", "InteractiveBrokers", "Maintenance") is None

    def test_returns_none_when_ts_dict_empty(self):
        tables = MarginTables(ts={}, ib={}, lookup={})
        assert tables.get_margin("ES", "TradeStation", "Maintenance") is None


# ── MarginTables.resolve_for_symbols ─────────────────────────────────────────

class TestResolveForSymbols:
    def test_returns_matching_symbols(self):
        tables = _sample_tables()
        result = tables.resolve_for_symbols(["ES", "NQ"], "TradeStation", "Maintenance")
        assert "ES" in result
        assert "NQ" in result
        assert result["ES"] == pytest.approx(20495.5)

    def test_skips_zero_or_missing(self):
        tables = _sample_tables()
        result = tables.resolve_for_symbols(["ES", "UNKNOWN"], "TradeStation", "Maintenance")
        assert "UNKNOWN" not in result
        assert "ES" in result

    def test_empty_symbol_list(self):
        tables = _sample_tables()
        result = tables.resolve_for_symbols([], "TradeStation", "Maintenance")
        assert result == {}

    def test_ib_source(self):
        tables = _sample_tables()
        result = tables.resolve_for_symbols(["ES", "NK"], "InteractiveBrokers", "Initial")
        assert result["ES"] == pytest.approx(15000.0)
        assert result["NK"] == pytest.approx(17000.0)  # NK → NKD

    def test_skips_zero_margin_entries(self):
        tables = MarginTables(
            ts={"ZZ": MarginEntry(initial=0.0, maintenance=0.0)},
            ib={}, lookup={},
        )
        result = tables.resolve_for_symbols(["ZZ"], "TradeStation", "Maintenance")
        assert result == {}


# ── sector_lookup ─────────────────────────────────────────────────────────────

class TestSectorLookup:
    def test_known_symbols(self):
        tables = _sample_tables()
        assert tables.sector_lookup["ES"] == "Index"
        assert tables.sector_lookup["CL"] == "Energy"
        assert tables.sector_lookup["GC"] == "Metals"
        assert tables.sector_lookup["VX"] == "Volatility"

    def test_missing_symbol_returns_key_error(self):
        tables = _sample_tables()
        with pytest.raises(KeyError):
            _ = tables.sector_lookup["UNKNOWN"]

    def test_resolve_for_symbols_uses_sector_via_get_margin(self):
        """sector_lookup is independent — get_margin is not affected by it."""
        tables = _sample_tables()
        result = tables.resolve_for_symbols(["ES", "CL"], "TradeStation", "Maintenance")
        assert "ES" in result  # margin lookup unaffected by sector_lookup

    def test_empty_sector_lookup(self):
        tables = MarginTables(ts={}, ib={}, lookup={}, sector_lookup={})
        assert tables.sector_lookup == {}

    def test_all_common_symbols_have_sectors(self):
        """Smoke-test the breadth of coverage in the sample data."""
        tables = _sample_tables()
        # symbols that should map to specific sectors
        expected = {
            "ES": "Index", "NQ": "Index",
            "CL": "Energy",
            "GC": "Metals", "SI": "Metals",
            "VX": "Volatility",
            "TY": "Interest Rate",
            "EC": "Currencies",
        }
        for sym, sector in expected.items():
            assert tables.sector_lookup.get(sym) == sector, (
                f"{sym}: expected {sector}, got {tables.sector_lookup.get(sym)}"
            )


# ── save / load round-trip ────────────────────────────────────────────────────

class TestSaveLoadMarginTables:
    def test_round_trip_includes_sector_lookup(self, tmp_path, monkeypatch):
        """sector_lookup persists through save/load."""
        margin_file = tmp_path / "margin_tables.yaml"
        monkeypatch.setattr("core.ingestion.xlsb_importer.MARGIN_TABLES_FILE", margin_file)
        monkeypatch.setattr("core.ingestion.xlsb_importer.CONFIG_DIR", tmp_path)

        tables = _sample_tables()
        save_margin_tables(tables)
        loaded = load_margin_tables()

        assert loaded is not None
        assert loaded.sector_lookup == tables.sector_lookup

    def test_round_trip(self, tmp_path, monkeypatch):
        """save_margin_tables / load_margin_tables round-trip."""
        margin_file = tmp_path / "margin_tables.yaml"
        monkeypatch.setattr(
            "core.ingestion.xlsb_importer.MARGIN_TABLES_FILE", margin_file
        )
        monkeypatch.setattr(
            "core.ingestion.xlsb_importer.CONFIG_DIR", tmp_path
        )

        tables = _sample_tables()
        save_margin_tables(tables)
        assert margin_file.exists()

        loaded = load_margin_tables()
        assert loaded is not None
        assert set(loaded.ts.keys()) == set(tables.ts.keys())
        assert loaded.ts["ES"].initial == pytest.approx(tables.ts["ES"].initial)
        assert loaded.ts["ES"].maintenance == pytest.approx(tables.ts["ES"].maintenance)
        assert loaded.ts["ES"].description == tables.ts["ES"].description
        assert set(loaded.ib.keys()) == set(tables.ib.keys())
        assert loaded.ib["NQ"].maintenance == pytest.approx(tables.ib["NQ"].maintenance)
        assert loaded.lookup == tables.lookup

    def test_load_returns_none_when_file_missing(self, tmp_path, monkeypatch):
        monkeypatch.setattr(
            "core.ingestion.xlsb_importer.MARGIN_TABLES_FILE", tmp_path / "nonexistent.yaml"
        )
        assert load_margin_tables() is None

    def test_load_returns_none_on_corrupt_yaml(self, tmp_path, monkeypatch):
        margin_file = tmp_path / "margin_tables.yaml"
        margin_file.write_text(":: not valid yaml ::\n{{{")
        monkeypatch.setattr(
            "core.ingestion.xlsb_importer.MARGIN_TABLES_FILE", margin_file
        )
        assert load_margin_tables() is None

    def test_save_creates_parent_dir(self, tmp_path, monkeypatch):
        nested = tmp_path / "a" / "b" / "margin_tables.yaml"
        monkeypatch.setattr("core.ingestion.xlsb_importer.MARGIN_TABLES_FILE", nested)
        monkeypatch.setattr("core.ingestion.xlsb_importer.CONFIG_DIR", tmp_path / "a" / "b")
        save_margin_tables(_sample_tables())
        assert nested.exists()


# ── _float helper ─────────────────────────────────────────────────────────────

class TestFloatHelper:
    def test_positive_number(self):
        assert _float(23013.0) == pytest.approx(23013.0)

    def test_zero_returns_none(self):
        assert _float(0) is None

    def test_negative_returns_none(self):
        assert _float(-100) is None

    def test_none_returns_none(self):
        assert _float(None) is None

    def test_string_number(self):
        assert _float("5000") == pytest.approx(5000.0)

    def test_non_numeric_string_returns_none(self):
        assert _float("NONE") is None


# ── _avg helper ───────────────────────────────────────────────────────────────

class TestAvgHelper:
    def test_both_values(self):
        assert _avg(10.0, 20.0) == pytest.approx(15.0)

    def test_one_none(self):
        assert _avg(10.0, None) == pytest.approx(10.0)

    def test_both_none(self):
        assert _avg(None, None) is None

    def test_equal_values(self):
        assert _avg(5000.0, 5000.0) == pytest.approx(5000.0)
