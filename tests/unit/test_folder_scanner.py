"""
Tests for core/ingestion/folder_scanner.py

Mirrors the exact behaviour of VBA C_Retrieve_Folder_Locations.bas.
"""

import pytest
from pathlib import Path

import sys
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from core.ingestion.folder_scanner import (
    scan_folders,
    reconcile_statuses,
    apply_not_loaded_prefix,
    strip_not_loaded_prefix,
    NOT_LOADED_PREFIX,
    WALKFORWARD_DIR,
    EQUITY_SUFFIX,
    TRADE_SUFFIX,
    WALKFORWARD_DETAILS,
)


# ── Helpers ───────────────────────────────────────────────────────────────────

def _make_strategy(tmp_path: Path, strategy_name: str, *, trade=True, wf=True) -> Path:
    """Create a minimal strategy folder structure under tmp_path."""
    subfolder = tmp_path / strategy_name
    wf_dir = subfolder / WALKFORWARD_DIR
    wf_dir.mkdir(parents=True)

    (wf_dir / f"{strategy_name}{EQUITY_SUFFIX}").write_text("date,m2m\n01/01/2024,100\n")

    if trade:
        (wf_dir / f"{strategy_name}{TRADE_SUFFIX}").write_text("date,type\n01/01/2024,Exit\n")

    if wf:
        (wf_dir / WALKFORWARD_DETAILS).write_text("IS,OOS\n")

    return subfolder


# ── scan_folders: base folder validation ─────────────────────────────────────

class TestScanFoldersBaseValidation:
    def test_nonexistent_base_folder_gives_error(self, tmp_path):
        missing = tmp_path / "does_not_exist"
        result = scan_folders([missing])
        assert len(result.errors) == 1
        assert "not found" in result.errors[0]
        assert result.strategies == []

    def test_file_instead_of_directory_gives_error(self, tmp_path):
        f = tmp_path / "not_a_dir.txt"
        f.write_text("x")
        result = scan_folders([f])
        assert len(result.errors) == 1
        assert "not a directory" in result.errors[0].lower()

    def test_empty_base_folder_returns_no_strategies(self, tmp_path):
        result = scan_folders([tmp_path])
        assert result.strategies == []
        assert result.errors == []

    def test_empty_list_returns_empty_result(self):
        result = scan_folders([])
        assert result.strategies == []
        assert result.errors == []
        assert result.warnings == []


# ── scan_folders: happy path ──────────────────────────────────────────────────

class TestScanFoldersHappyPath:
    def test_single_strategy_found(self, tmp_path):
        _make_strategy(tmp_path, "StratA")
        result = scan_folders([tmp_path])
        assert len(result.strategies) == 1
        assert result.strategies[0].name == "StratA"
        assert result.errors == []

    def test_multiple_strategies_found(self, tmp_path):
        _make_strategy(tmp_path, "StratA")
        _make_strategy(tmp_path, "StratB")
        _make_strategy(tmp_path, "StratC")
        result = scan_folders([tmp_path])
        names = {s.name for s in result.strategies}
        assert names == {"StratA", "StratB", "StratC"}

    def test_strategy_folder_has_correct_paths(self, tmp_path):
        _make_strategy(tmp_path, "StratA")
        result = scan_folders([tmp_path])
        sf = result.strategies[0]
        assert sf.path == tmp_path / "StratA"
        assert sf.equity_csv.name == f"StratA{EQUITY_SUFFIX}"
        assert sf.trade_csv.name == f"StratA{TRADE_SUFFIX}"
        assert sf.walkforward_csv.name == WALKFORWARD_DETAILS

    def test_results_sorted_alphabetically(self, tmp_path):
        for name in ["Zebra", "Apple", "Mango"]:
            _make_strategy(tmp_path, name)
        result = scan_folders([tmp_path])
        names = [s.name for s in result.strategies]
        assert names == sorted(names)

    def test_multiple_base_folders(self, tmp_path):
        base1 = tmp_path / "base1"
        base2 = tmp_path / "base2"
        base1.mkdir(); base2.mkdir()
        _make_strategy(base1, "StratA")
        _make_strategy(base2, "StratB")
        result = scan_folders([base1, base2])
        names = {s.name for s in result.strategies}
        assert names == {"StratA", "StratB"}


# ── scan_folders: optional files ─────────────────────────────────────────────

class TestScanFoldersOptionalFiles:
    def test_missing_trade_csv_gives_warning_not_error(self, tmp_path):
        _make_strategy(tmp_path, "StratA", trade=False, wf=True)
        result = scan_folders([tmp_path])
        assert len(result.strategies) == 1
        assert result.strategies[0].trade_csv is None
        assert any("TradeData" in w for w in result.warnings)

    def test_missing_wf_details_gives_warning_not_error(self, tmp_path):
        _make_strategy(tmp_path, "StratA", trade=True, wf=False)
        result = scan_folders([tmp_path])
        assert len(result.strategies) == 1
        assert result.strategies[0].walkforward_csv is None
        assert any(WALKFORWARD_DETAILS in w for w in result.warnings)

    def test_both_optional_missing(self, tmp_path):
        _make_strategy(tmp_path, "StratA", trade=False, wf=False)
        result = scan_folders([tmp_path])
        assert len(result.strategies) == 1
        assert len(result.warnings) == 2

    def test_subfolder_without_walkforward_dir_is_silently_skipped(self, tmp_path):
        # No Walkforward Files/ directory
        (tmp_path / "RandomFolder").mkdir()
        result = scan_folders([tmp_path])
        assert result.strategies == []
        assert result.warnings == []

    def test_equity_csv_missing_gives_warning_and_skip(self, tmp_path):
        subfolder = tmp_path / "StratNoEquity"
        wf_dir = subfolder / WALKFORWARD_DIR
        wf_dir.mkdir(parents=True)
        # No equity CSV
        result = scan_folders([tmp_path])
        assert result.strategies == []
        assert any("EquityData" in w for w in result.warnings)

    def test_files_in_base_folder_are_ignored(self, tmp_path):
        """Non-directory items in base folder should not cause errors."""
        (tmp_path / "some_file.txt").write_text("x")
        _make_strategy(tmp_path, "StratA")
        result = scan_folders([tmp_path])
        assert len(result.strategies) == 1


# ── scan_folders: multiple strategies per subfolder ──────────────────────────

class TestMultipleStrategiesPerSubfolder:
    """
    One subfolder can yield multiple strategies when its Walkforward Files/
    directory contains several *EquityData.csv files (common for Buy & Hold
    folders where many instruments share one directory).
    Mirrors VBA GetFolderData Dir() loop behaviour.
    """

    def _make_multi_strategy_subfolder(
        self, tmp_path: Path, subfolder_name: str, strategy_names: list[str], *, wf: bool = True
    ) -> Path:
        """Create a subfolder with multiple *EquityData.csv files."""
        subfolder = tmp_path / subfolder_name
        wf_dir = subfolder / WALKFORWARD_DIR
        wf_dir.mkdir(parents=True)
        for name in strategy_names:
            (wf_dir / f"{name}{EQUITY_SUFFIX}").write_text("date,m2m\n01/01/2024,100\n")
            (wf_dir / f"{name}{TRADE_SUFFIX}").write_text("date,type\n01/01/2024,Exit\n")
        if wf:
            (wf_dir / WALKFORWARD_DETAILS).write_text("IS,OOS\n")
        return subfolder

    def test_multiple_equity_files_yield_multiple_strategies(self, tmp_path):
        names = ["BnH NQ", "BnH ES", "BnH CL"]
        self._make_multi_strategy_subfolder(tmp_path, "BuyAndHold", names)
        result = scan_folders([tmp_path])
        found = {s.name for s in result.strategies}
        assert found == set(names)
        assert result.errors == []

    def test_44_strategies_in_one_subfolder(self, tmp_path):
        names = [f"BnH {i:02d}" for i in range(44)]
        self._make_multi_strategy_subfolder(tmp_path, "BuyAndHold", names)
        result = scan_folders([tmp_path])
        assert len(result.strategies) == 44
        assert result.errors == []

    def test_each_strategy_gets_correct_equity_csv(self, tmp_path):
        names = ["BnH NQ", "BnH ES"]
        self._make_multi_strategy_subfolder(tmp_path, "BuyAndHold", names)
        result = scan_folders([tmp_path])
        by_name = {s.name: s for s in result.strategies}
        assert by_name["BnH NQ"].equity_csv.name == f"BnH NQ{EQUITY_SUFFIX}"
        assert by_name["BnH ES"].equity_csv.name == f"BnH ES{EQUITY_SUFFIX}"

    def test_each_strategy_gets_correct_trade_csv(self, tmp_path):
        names = ["BnH NQ", "BnH ES"]
        self._make_multi_strategy_subfolder(tmp_path, "BuyAndHold", names)
        result = scan_folders([tmp_path])
        by_name = {s.name: s for s in result.strategies}
        assert by_name["BnH NQ"].trade_csv is not None
        assert by_name["BnH NQ"].trade_csv.name == f"BnH NQ{TRADE_SUFFIX}"
        assert by_name["BnH ES"].trade_csv.name == f"BnH ES{TRADE_SUFFIX}"

    def test_walkforward_details_shared_across_strategies(self, tmp_path):
        names = ["BnH NQ", "BnH ES", "BnH CL"]
        self._make_multi_strategy_subfolder(tmp_path, "BuyAndHold", names)
        result = scan_folders([tmp_path])
        # All strategies share the same walkforward CSV from the same folder
        wf_csvs = {s.walkforward_csv for s in result.strategies}
        assert len(wf_csvs) == 1
        assert next(iter(wf_csvs)).name == WALKFORWARD_DETAILS

    def test_missing_wf_details_warns_once_per_subfolder(self, tmp_path):
        names = ["BnH NQ", "BnH ES", "BnH CL"]
        self._make_multi_strategy_subfolder(tmp_path, "BuyAndHold", names, wf=False)
        result = scan_folders([tmp_path])
        assert len(result.strategies) == 3
        wf_warnings = [w for w in result.warnings if WALKFORWARD_DETAILS in w]
        # One warning per subfolder, not per strategy
        assert len(wf_warnings) == 1

    def test_missing_trade_csv_warns_per_strategy(self, tmp_path):
        subfolder = tmp_path / "BuyAndHold"
        wf_dir = subfolder / WALKFORWARD_DIR
        wf_dir.mkdir(parents=True)
        for name in ["BnH NQ", "BnH ES"]:
            (wf_dir / f"{name}{EQUITY_SUFFIX}").write_text("date,m2m\n01/01/2024,100\n")
            # No TradeData.csv
        result = scan_folders([tmp_path])
        assert len(result.strategies) == 2
        trade_warnings = [w for w in result.warnings if "TradeData" in w]
        assert len(trade_warnings) == 2

    def test_multi_strategy_combined_with_single_strategy_folders(self, tmp_path):
        # Standard single-strategy folders alongside a multi-strategy BnH folder
        _make_strategy(tmp_path, "StratA")
        _make_strategy(tmp_path, "StratB")
        self._make_multi_strategy_subfolder(tmp_path, "BuyAndHold", ["BnH NQ", "BnH ES"])
        result = scan_folders([tmp_path])
        names = {s.name for s in result.strategies}
        assert names == {"StratA", "StratB", "BnH NQ", "BnH ES"}

    def test_results_sorted_alphabetically_across_strategies(self, tmp_path):
        names = ["BnH ZZ", "BnH AA", "BnH MM"]
        self._make_multi_strategy_subfolder(tmp_path, "BuyAndHold", names)
        result = scan_folders([tmp_path])
        returned_names = [s.name for s in result.strategies]
        assert returned_names == sorted(returned_names)

    def test_all_strategies_share_same_subfolder_path(self, tmp_path):
        names = ["BnH NQ", "BnH ES"]
        self._make_multi_strategy_subfolder(tmp_path, "BuyAndHold", names)
        result = scan_folders([tmp_path])
        paths = {s.path for s in result.strategies}
        assert len(paths) == 1  # all point to the same subfolder


# ── scan_folders: duplicate detection ────────────────────────────────────────

class TestScanFoldersDuplicates:
    def test_duplicate_name_across_base_folders_gives_warning(self, tmp_path):
        base1 = tmp_path / "base1"
        base2 = tmp_path / "base2"
        base1.mkdir(); base2.mkdir()
        _make_strategy(base1, "StratA")
        _make_strategy(base2, "StratA")  # same name
        result = scan_folders([base1, base2])
        # Only one strategy kept
        assert len(result.strategies) == 1
        assert any("Duplicate" in w or "duplicate" in w for w in result.warnings)

    def test_first_occurrence_kept_on_duplicate(self, tmp_path):
        base1 = tmp_path / "base1"
        base2 = tmp_path / "base2"
        base1.mkdir(); base2.mkdir()
        _make_strategy(base1, "StratA")
        _make_strategy(base2, "StratA")
        result = scan_folders([base1, base2])
        assert result.strategies[0].path.parent == base1


# ── reconcile_statuses ────────────────────────────────────────────────────────

class TestReconcileStatuses:
    def test_new_strategy_added_with_new_status(self):
        found = {"NewStrat"}
        configured = []
        result = reconcile_statuses(found, configured)
        assert len(result) == 1
        assert result[0]["name"] == "NewStrat"
        assert result[0]["status"] == "New"

    def test_found_strategy_strips_not_loaded_prefix(self):
        found = {"StratA"}
        configured = [{"name": "StratA", "status": "Not Loaded - Active"}]
        result = reconcile_statuses(found, configured)
        assert result[0]["status"] == "Active"

    def test_missing_strategy_gets_not_loaded_prefix(self):
        found = set()
        configured = [{"name": "StratA", "status": "Active"}]
        result = reconcile_statuses(found, configured)
        assert result[0]["status"] == "Not Loaded - Active"

    def test_already_not_loaded_stays_not_loaded(self):
        found = set()
        configured = [{"name": "StratA", "status": "Not Loaded - Active"}]
        result = reconcile_statuses(found, configured)
        # Should not double-prefix
        assert result[0]["status"] == "Not Loaded - Active"
        assert result[0]["status"].count("Not Loaded -") == 1

    def test_found_strategy_keeps_other_fields(self):
        found = {"StratA"}
        configured = [{"name": "StratA", "status": "Active", "contracts": 3, "symbol": "ES"}]
        result = reconcile_statuses(found, configured)
        assert result[0]["contracts"] == 3
        assert result[0]["symbol"] == "ES"

    def test_new_strategy_has_default_fields(self):
        found = {"BrandNew"}
        configured = []
        result = reconcile_statuses(found, configured)
        row = result[0]
        assert row["contracts"] == 1
        assert row["symbol"] == ""
        assert row["sector"] == ""

    def test_mixed_found_missing_new(self):
        found = {"StratA", "StratC"}
        configured = [
            {"name": "StratA", "status": "Active"},
            {"name": "StratB", "status": "Inactive"},
        ]
        result = reconcile_statuses(found, configured)
        by_name = {r["name"]: r for r in result}
        assert by_name["StratA"]["status"] == "Active"
        assert by_name["StratB"]["status"] == "Not Loaded - Inactive"
        assert by_name["StratC"]["status"] == "New"

    def test_empty_found_and_configured(self):
        result = reconcile_statuses(set(), [])
        assert result == []

    def test_new_strategies_sorted_alphabetically(self):
        found = {"Zebra", "Apple", "Mango"}
        configured = []
        result = reconcile_statuses(found, configured)
        names = [r["name"] for r in result]
        assert names == sorted(names)


# ── Not Loaded prefix helpers ─────────────────────────────────────────────────

class TestNotLoadedHelpers:
    def test_apply_prefix_adds_it(self):
        assert apply_not_loaded_prefix("Active") == "Not Loaded - Active"

    def test_apply_prefix_does_not_double_add(self):
        s = apply_not_loaded_prefix("Active")
        s2 = apply_not_loaded_prefix(s)
        assert s2 == "Not Loaded - Active"

    def test_strip_prefix_removes_it(self):
        assert strip_not_loaded_prefix("Not Loaded - Active") == "Active"

    def test_strip_prefix_no_op_if_not_present(self):
        assert strip_not_loaded_prefix("Active") == "Active"

    def test_not_loaded_prefix_constant(self):
        assert NOT_LOADED_PREFIX == "Not Loaded - "
