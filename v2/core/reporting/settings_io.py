"""
Settings export / import — mirrors Q_Export_Import_Settings.bas.

Exports a ZIP archive containing all user configuration:
  manifest.json       — version + timestamp
  settings.yaml       — AppConfig (margin, MC params, folders, etc.)
  strategies.yaml     — strategy list
  margin_tables.yaml  — TS/IB margin reference tables (if present)

Import validates the manifest and restores files to CONFIG_DIR.
"""

from __future__ import annotations

import io
import json
from datetime import datetime, timezone
from pathlib import Path

import yaml

from core.config import CONFIG_DIR, USER_CONFIG_FILE, AppConfig

_STRATEGIES_FILE = CONFIG_DIR / "strategies.yaml"
_MARGIN_FILE     = CONFIG_DIR / "margin_tables.yaml"

_VERSION = "2.0.0"
_MANIFEST_KEY = "portfolio_tracker_config"


# ── Export ────────────────────────────────────────────────────────────────────

def export_settings() -> bytes:
    """
    Bundle all user config files into a ZIP archive.
    Returns the raw ZIP bytes suitable for st.download_button.
    """
    import zipfile

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # Manifest
        manifest = {
            "type": _MANIFEST_KEY,
            "version": _VERSION,
            "exported_at": datetime.now(timezone.utc).isoformat(),
        }
        zf.writestr("manifest.json", json.dumps(manifest, indent=2))

        # settings.yaml
        if USER_CONFIG_FILE.exists():
            zf.write(USER_CONFIG_FILE, "settings.yaml")

        # strategies.yaml
        if _STRATEGIES_FILE.exists():
            zf.write(_STRATEGIES_FILE, "strategies.yaml")

        # margin_tables.yaml (optional)
        if _MARGIN_FILE.exists():
            zf.write(_MARGIN_FILE, "margin_tables.yaml")

    return buf.getvalue()


# ── Import ────────────────────────────────────────────────────────────────────

def import_settings(zip_bytes: bytes) -> tuple[bool, str, list[str]]:
    """
    Restore configuration from a previously exported ZIP archive.

    Returns:
        (success, error_message, restored_files)
        On success: error_message is "", restored_files lists what was restored.
        On failure: error_message describes the problem.
    """
    import zipfile

    try:
        buf = io.BytesIO(zip_bytes)
        with zipfile.ZipFile(buf, "r") as zf:
            names = zf.namelist()

            # Validate manifest
            if "manifest.json" not in names:
                return False, "Not a valid Portfolio Tracker config archive (missing manifest.json).", []

            manifest = json.loads(zf.read("manifest.json"))
            if manifest.get("type") != _MANIFEST_KEY:
                return False, "Not a valid Portfolio Tracker config archive.", []

            # Ensure config dir exists
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)

            restored: list[str] = []
            _FILE_MAP = {
                "settings.yaml":       USER_CONFIG_FILE,
                "strategies.yaml":     _STRATEGIES_FILE,
                "margin_tables.yaml":  _MARGIN_FILE,
            }

            for archive_name, dest_path in _FILE_MAP.items():
                if archive_name in names:
                    data = zf.read(archive_name)
                    # Validate YAML before overwriting
                    try:
                        yaml.safe_load(data)
                    except yaml.YAMLError as e:
                        return False, f"{archive_name} contains invalid YAML: {e}", []
                    dest_path.write_bytes(data)
                    restored.append(archive_name)

            if not restored:
                return False, "Archive contained no recognisable config files.", []

            return True, "", restored

    except zipfile.BadZipFile:
        return False, "The file is not a valid ZIP archive.", []
    except Exception as e:
        return False, f"Unexpected error during import: {e}", []


# ── Helpers ───────────────────────────────────────────────────────────────────

def default_export_filename() -> str:
    """Return a timestamped default filename for the export."""
    ts = datetime.now().strftime("%Y-%m-%d")
    return f"PortfolioTrackerConfig_{ts}.zip"
