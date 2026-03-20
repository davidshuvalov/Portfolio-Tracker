"""
Strategies configuration persistence.
Stores strategy metadata (status, sector, contracts etc.) as YAML.
"""

from __future__ import annotations
from pathlib import Path
import yaml

STRATEGIES_FILE = Path.home() / ".portfolio_tracker" / "strategies.yaml"


def load_strategies() -> list[dict]:
    """Load strategies from the user config file. Returns empty list if not found."""
    if not STRATEGIES_FILE.exists():
        return []
    with open(STRATEGIES_FILE) as f:
        data = yaml.safe_load(f) or []
    return data if isinstance(data, list) else []


def save_strategies(strategies: list[dict]) -> None:
    """Persist strategies list to the local YAML file and sync to cloud if logged in."""
    STRATEGIES_FILE.parent.mkdir(parents=True, exist_ok=True)
    # Ensure all values are serialisable (convert None to "")
    cleaned = [_clean_row(s) for s in strategies]
    with open(STRATEGIES_FILE, "w") as f:
        yaml.dump(cleaned, f, default_flow_style=False, sort_keys=False, allow_unicode=True)
    # Best-effort cloud sync — never raises
    try:
        import streamlit as st
        if st.session_state.get("_sb_session"):
            from core.cloud_sync import save_strategies_to_cloud
            save_strategies_to_cloud(cleaned)
    except Exception:
        pass


def _clean_row(row: dict) -> dict:
    return {
        k: ("" if v is None else v)
        for k, v in row.items()
    }
