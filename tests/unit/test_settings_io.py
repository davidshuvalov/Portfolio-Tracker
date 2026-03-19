"""
Unit tests for core.reporting.settings_io.

Tests export/import round-trip, manifest validation, and error handling.
All tests use temporary directories — no ~/.portfolio_tracker modification.
"""
from __future__ import annotations

import json
import zipfile
from pathlib import Path

import pytest
import yaml

from core.reporting.settings_io import (
    _MANIFEST_KEY,
    _VERSION,
    default_export_filename,
    export_settings,
    import_settings,
)


# ── Helpers ────────────────────────────────────────────────────────────────────

def _make_zip(files: dict[str, bytes | str], manifest: dict | None = None) -> bytes:
    """Build a synthetic ZIP archive for import testing."""
    import io
    if manifest is None:
        manifest = {"type": _MANIFEST_KEY, "version": _VERSION, "exported_at": "2026-01-01T00:00:00+00:00"}
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("manifest.json", json.dumps(manifest))
        for name, data in files.items():
            if isinstance(data, str):
                zf.writestr(name, data)
            else:
                zf.writestr(name, data)
    return buf.getvalue()


# ── export_settings ────────────────────────────────────────────────────────────

class TestExportSettings:
    def test_returns_bytes(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.reporting.settings_io.CONFIG_DIR", tmp_path)
        monkeypatch.setattr("core.reporting.settings_io.USER_CONFIG_FILE", tmp_path / "settings.yaml")
        monkeypatch.setattr("core.reporting.settings_io._STRATEGIES_FILE", tmp_path / "strategies.yaml")
        monkeypatch.setattr("core.reporting.settings_io._MARGIN_FILE", tmp_path / "margin_tables.yaml")
        result = export_settings()
        assert isinstance(result, bytes)
        assert len(result) > 0

    def test_zip_contains_manifest(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.reporting.settings_io.CONFIG_DIR", tmp_path)
        monkeypatch.setattr("core.reporting.settings_io.USER_CONFIG_FILE", tmp_path / "s.yaml")
        monkeypatch.setattr("core.reporting.settings_io._STRATEGIES_FILE", tmp_path / "st.yaml")
        monkeypatch.setattr("core.reporting.settings_io._MARGIN_FILE", tmp_path / "m.yaml")
        raw = export_settings()
        import io
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            assert "manifest.json" in zf.namelist()
            manifest = json.loads(zf.read("manifest.json"))
            assert manifest["type"] == _MANIFEST_KEY
            assert manifest["version"] == _VERSION

    def test_includes_existing_files(self, tmp_path, monkeypatch):
        settings_file = tmp_path / "settings.yaml"
        strategies_file = tmp_path / "strategies.yaml"
        settings_file.write_text("folders: []\n")
        strategies_file.write_text("- name: ES_Trend\n")
        monkeypatch.setattr("core.reporting.settings_io.CONFIG_DIR", tmp_path)
        monkeypatch.setattr("core.reporting.settings_io.USER_CONFIG_FILE", settings_file)
        monkeypatch.setattr("core.reporting.settings_io._STRATEGIES_FILE", strategies_file)
        monkeypatch.setattr("core.reporting.settings_io._MARGIN_FILE", tmp_path / "nonexistent.yaml")
        import io
        raw = export_settings()
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            names = zf.namelist()
        assert "settings.yaml" in names
        assert "strategies.yaml" in names
        assert "margin_tables.yaml" not in names  # didn't exist

    def test_skips_missing_files(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.reporting.settings_io.CONFIG_DIR", tmp_path)
        monkeypatch.setattr("core.reporting.settings_io.USER_CONFIG_FILE", tmp_path / "missing.yaml")
        monkeypatch.setattr("core.reporting.settings_io._STRATEGIES_FILE", tmp_path / "missing2.yaml")
        monkeypatch.setattr("core.reporting.settings_io._MARGIN_FILE", tmp_path / "missing3.yaml")
        import io
        raw = export_settings()
        with zipfile.ZipFile(io.BytesIO(raw)) as zf:
            names = zf.namelist()
        assert "manifest.json" in names
        assert "settings.yaml" not in names


# ── import_settings ────────────────────────────────────────────────────────────

class TestImportSettings:
    def test_round_trip(self, tmp_path, monkeypatch):
        settings_content = "folders: []\ndate_format: DMY\n"
        strategies_content = "- name: NQ_Trend\n  status: Live\n"
        monkeypatch.setattr("core.reporting.settings_io.CONFIG_DIR", tmp_path)
        monkeypatch.setattr("core.reporting.settings_io.USER_CONFIG_FILE", tmp_path / "settings.yaml")
        monkeypatch.setattr("core.reporting.settings_io._STRATEGIES_FILE", tmp_path / "strategies.yaml")
        monkeypatch.setattr("core.reporting.settings_io._MARGIN_FILE", tmp_path / "margin.yaml")

        raw = _make_zip({
            "settings.yaml": settings_content,
            "strategies.yaml": strategies_content,
        })
        ok, err, restored = import_settings(raw)
        assert ok is True
        assert err == ""
        assert "settings.yaml" in restored
        assert "strategies.yaml" in restored
        assert (tmp_path / "settings.yaml").read_text() == settings_content

    def test_missing_manifest_fails(self):
        import io
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr("settings.yaml", "folders: []\n")
        ok, err, _ = import_settings(buf.getvalue())
        assert ok is False
        assert "manifest.json" in err

    def test_wrong_manifest_type_fails(self):
        raw = _make_zip({}, manifest={"type": "wrong", "version": "1.0"})
        ok, err, _ = import_settings(raw)
        assert ok is False
        assert "valid" in err.lower()

    def test_corrupt_zip_fails(self):
        ok, err, _ = import_settings(b"not a zip file at all")
        assert ok is False
        assert "ZIP" in err or "zip" in err.lower()

    def test_invalid_yaml_fails(self, tmp_path, monkeypatch):
        monkeypatch.setattr("core.reporting.settings_io.CONFIG_DIR", tmp_path)
        monkeypatch.setattr("core.reporting.settings_io.USER_CONFIG_FILE", tmp_path / "s.yaml")
        monkeypatch.setattr("core.reporting.settings_io._STRATEGIES_FILE", tmp_path / "st.yaml")
        monkeypatch.setattr("core.reporting.settings_io._MARGIN_FILE", tmp_path / "m.yaml")
        raw = _make_zip({"settings.yaml": "{invalid: yaml: format"})
        ok, err, _ = import_settings(raw)
        assert ok is False
        assert "YAML" in err or "yaml" in err.lower()

    def test_empty_zip_no_config_files_fails(self):
        raw = _make_zip({})  # only manifest, no config files
        ok, err, _ = import_settings(raw)
        assert ok is False
        assert "no recognisable" in err.lower()

    def test_partial_restore(self, tmp_path, monkeypatch):
        """Only strategies.yaml in archive → only that file restored."""
        monkeypatch.setattr("core.reporting.settings_io.CONFIG_DIR", tmp_path)
        monkeypatch.setattr("core.reporting.settings_io.USER_CONFIG_FILE", tmp_path / "settings.yaml")
        monkeypatch.setattr("core.reporting.settings_io._STRATEGIES_FILE", tmp_path / "strategies.yaml")
        monkeypatch.setattr("core.reporting.settings_io._MARGIN_FILE", tmp_path / "margin.yaml")
        raw = _make_zip({"strategies.yaml": "- name: CL_Mean\n"})
        ok, err, restored = import_settings(raw)
        assert ok is True
        assert restored == ["strategies.yaml"]
        assert not (tmp_path / "settings.yaml").exists()


# ── default_export_filename ────────────────────────────────────────────────────

class TestDefaultExportFilename:
    def test_contains_date(self):
        name = default_export_filename()
        import re
        assert re.search(r"\d{4}-\d{2}-\d{2}", name)

    def test_ends_with_zip(self):
        assert default_export_filename().endswith(".zip")
