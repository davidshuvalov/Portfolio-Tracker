"""Unit tests for core.portfolio.snapshot."""
from __future__ import annotations

from pathlib import Path

import pytest

from core.portfolio.snapshot import (
    CompareResult,
    compare_portfolios,
    delete_snapshot,
    list_snapshots,
    load_snapshot,
    save_snapshot,
)


def _strats(*names, status="Live", contracts=1) -> list[dict]:
    return [{"name": n, "status": status, "contracts": contracts, "symbol": "ES", "sector": "Index"} for n in names]


# ── save / load / list / delete ────────────────────────────────────────────────

class TestSnapshotPersistence:
    def test_save_creates_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.portfolio.snapshot.SNAPSHOT_DIR", tmp_path)
        path = save_snapshot(_strats("A", "B"), label="test1")
        assert path.exists()

    def test_save_only_live_strategies(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.portfolio.snapshot.SNAPSHOT_DIR", tmp_path)
        mixed = _strats("A", "B") + _strats("C", status="Paper") + _strats("D", status="Retired")
        path = save_snapshot(mixed, label="liveonly")
        loaded = load_snapshot(path.name)
        assert [s["name"] for s in loaded] == ["A", "B"]

    def test_list_snapshots_returns_metadata(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.portfolio.snapshot.SNAPSHOT_DIR", tmp_path)
        save_snapshot(_strats("A"), label="snap1")
        save_snapshot(_strats("A", "B"), label="snap2")
        snaps = list_snapshots()
        assert len(snaps) == 2
        for s in snaps:
            for k in ("filename", "label", "saved_at", "n_strategies"):
                assert k in s

    def test_list_snapshots_empty_dir(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.portfolio.snapshot.SNAPSHOT_DIR", tmp_path)
        assert list_snapshots() == []

    def test_list_snapshots_no_dir(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.portfolio.snapshot.SNAPSHOT_DIR", tmp_path / "nonexistent")
        assert list_snapshots() == []

    def test_delete_removes_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.portfolio.snapshot.SNAPSHOT_DIR", tmp_path)
        path = save_snapshot(_strats("A"), label="del_me")
        delete_snapshot(path.name)
        assert not path.exists()

    def test_delete_missing_file_no_error(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.portfolio.snapshot.SNAPSHOT_DIR", tmp_path)
        delete_snapshot("does_not_exist.yaml")  # should not raise

    def test_slug_sanitises_special_chars(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.portfolio.snapshot.SNAPSHOT_DIR", tmp_path)
        path = save_snapshot(_strats("A"), label="Pre March/2026 - test!")
        assert path.exists()
        # filename should not contain special chars
        assert "/" not in path.name
        assert "!" not in path.name


# ── compare_portfolios ─────────────────────────────────────────────────────────

class TestComparePortfolios:
    def test_identical_portfolios_no_changes(self):
        current = _strats("A", "B", "C")
        reference = _strats("A", "B", "C")
        result = compare_portfolios(current, reference)
        assert not result.has_changes
        assert len(result.unchanged) == 3

    def test_new_strategy_detected(self):
        current = _strats("A", "B", "C")
        reference = _strats("A", "B")
        result = compare_portfolios(current, reference)
        assert len(result.new_strategies) == 1
        assert result.new_strategies[0]["name"] == "C"

    def test_removed_strategy_detected(self):
        current = _strats("A", "B")
        reference = _strats("A", "B", "C")
        result = compare_portfolios(current, reference)
        assert len(result.removed_strategies) == 1
        assert result.removed_strategies[0]["name"] == "C"

    def test_contract_change_detected(self):
        current = [{"name": "A", "status": "Live", "contracts": 3, "symbol": "ES", "sector": ""}]
        reference = [{"name": "A", "status": "Live", "contracts": 1, "symbol": "ES", "sector": ""}]
        result = compare_portfolios(current, reference)
        assert len(result.contract_changes) == 1
        chg = result.contract_changes[0]
        assert chg["old_contracts"] == 1
        assert chg["new_contracts"] == 3
        assert chg["delta"] == 2

    def test_paper_strategy_not_in_new(self):
        # Paper strategies in current should not appear as new
        current = _strats("A") + _strats("B", status="Paper")
        reference = _strats("A")
        result = compare_portfolios(current, reference)
        assert result.new_strategies == []

    def test_all_removed_when_current_empty(self):
        current = _strats("A", status="Retired")
        reference = _strats("A", "B")
        result = compare_portfolios(current, reference)
        assert len(result.removed_strategies) == 2

    def test_has_changes_false_when_no_changes(self):
        current = _strats("A", "B")
        reference = _strats("A", "B")
        result = compare_portfolios(current, reference)
        assert not result.has_changes

    def test_has_changes_true_with_new(self):
        result = compare_portfolios(_strats("A", "B"), _strats("A"))
        assert result.has_changes

    def test_summary_string(self):
        current = _strats("A", "B", "C")
        reference = _strats("A", "B")
        result = compare_portfolios(current, reference)
        assert "new" in result.summary

    def test_no_changes_summary(self):
        result = compare_portfolios(_strats("A"), _strats("A"))
        assert "no changes" in result.summary

    def test_custom_live_status(self):
        current = [{"name": "A", "status": "Active", "contracts": 1, "symbol": "ES", "sector": ""}]
        reference = [{"name": "B", "status": "Live", "contracts": 1, "symbol": "ES", "sector": ""}]
        result = compare_portfolios(current, reference, live_status="Active")
        assert len(result.new_strategies) == 1
        assert result.new_strategies[0]["name"] == "A"
        # B was in reference but current has no "Active" B → removed
        assert len(result.removed_strategies) == 1
