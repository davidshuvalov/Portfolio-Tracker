"""Portfolio snapshot — save, load and compare Live strategy configurations.

Mirrors R_Check_New_Strategies.bas, replacing the VBA "pick a reference .xlsb
file" approach with a lightweight YAML snapshot stored alongside the user's
other config files.

Snapshots live at:  ~/.portfolio_tracker/snapshots/<timestamp>_<slug>.yaml

Each snapshot file stores:
  saved_at:   ISO-8601 datetime string
  label:      user-supplied description
  strategies: list of {name, contracts, symbol, sector, status}
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

import yaml

SNAPSHOT_DIR = Path.home() / ".portfolio_tracker" / "snapshots"

_KEEP_FIELDS = ("name", "contracts", "symbol", "sector", "status")


# ── Persistence helpers ────────────────────────────────────────────────────────

def _slug(label: str) -> str:
    """Convert a label to a filesystem-safe slug."""
    return re.sub(r"[^a-zA-Z0-9]+", "_", label).strip("_")[:40] or "snapshot"


def save_snapshot(strategies: list[dict], label: str) -> Path:
    """Persist a snapshot of the supplied strategies list.

    Only Live strategies are stored (status == 'Live'); pass the full
    strategies list — filtering is done here.

    Returns the path of the created file.
    """
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{ts}_{_slug(label)}.yaml"
    path = SNAPSHOT_DIR / filename

    live = [
        {k: s.get(k, "") for k in _KEEP_FIELDS}
        for s in strategies
        if s.get("status") == "Live"
    ]

    payload: dict[str, Any] = {
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "label": label,
        "strategies": live,
    }
    with open(path, "w") as f:
        yaml.dump(payload, f, default_flow_style=False, sort_keys=False, allow_unicode=True)

    return path


def list_snapshots() -> list[dict]:
    """Return metadata for all saved snapshots, newest first.

    Each entry: {filename, label, saved_at, n_strategies, path}
    """
    if not SNAPSHOT_DIR.exists():
        return []

    results = []
    for p in sorted(SNAPSHOT_DIR.glob("*.yaml"), reverse=True):
        try:
            with open(p) as f:
                data = yaml.safe_load(f) or {}
            results.append(
                {
                    "filename": p.name,
                    "label": data.get("label", p.stem),
                    "saved_at": data.get("saved_at", ""),
                    "n_strategies": len(data.get("strategies", [])),
                    "path": p,
                }
            )
        except Exception:
            pass

    return results


def load_snapshot(filename: str) -> list[dict]:
    """Load and return the strategies list from a snapshot file."""
    path = SNAPSHOT_DIR / filename
    with open(path) as f:
        data = yaml.safe_load(f) or {}
    return data.get("strategies", [])


def delete_snapshot(filename: str) -> None:
    """Delete a snapshot file."""
    (SNAPSHOT_DIR / filename).unlink(missing_ok=True)


# ── Comparison ─────────────────────────────────────────────────────────────────

@dataclass
class CompareResult:
    new_strategies: list[dict] = field(default_factory=list)
    """In current Live but absent from the snapshot."""

    removed_strategies: list[dict] = field(default_factory=list)
    """In snapshot but no longer Live (or removed entirely)."""

    contract_changes: list[dict] = field(default_factory=list)
    """Strategy present in both but with different contracts.
    Each entry: {name, symbol, old_contracts, new_contracts, delta}"""

    unchanged: list[dict] = field(default_factory=list)
    """Same name and contracts in both."""

    @property
    def has_changes(self) -> bool:
        return bool(
            self.new_strategies or self.removed_strategies or self.contract_changes
        )

    @property
    def summary(self) -> str:
        parts = []
        if self.new_strategies:
            parts.append(f"{len(self.new_strategies)} new")
        if self.removed_strategies:
            parts.append(f"{len(self.removed_strategies)} removed")
        if self.contract_changes:
            parts.append(f"{len(self.contract_changes)} contract changes")
        if not parts:
            parts.append("no changes")
        return ", ".join(parts)


def compare_portfolios(
    current: list[dict],
    reference: list[dict],
    live_status: str = "Live",
) -> CompareResult:
    """Compare the current Live strategies against a reference snapshot.

    *current*   — full strategies list (all statuses); only Live are evaluated.
    *reference* — strategies list from a saved snapshot (already Live-filtered).
    """
    current_live = {
        s["name"]: s for s in current if s.get("status") == live_status
    }
    ref_map = {s["name"]: s for s in reference}

    result = CompareResult()

    # Strategies in current Live
    for name, s in current_live.items():
        if name not in ref_map:
            result.new_strategies.append(dict(s))
        else:
            ref = ref_map[name]
            cur_c = int(s.get("contracts", 1) or 1)
            ref_c = int(ref.get("contracts", 1) or 1)
            if cur_c != ref_c:
                result.contract_changes.append(
                    {
                        "name": name,
                        "symbol": s.get("symbol", ""),
                        "old_contracts": ref_c,
                        "new_contracts": cur_c,
                        "delta": cur_c - ref_c,
                    }
                )
            else:
                result.unchanged.append(dict(s))

    # Strategies in snapshot but no longer Live
    for name, ref in ref_map.items():
        if name not in current_live:
            result.removed_strategies.append(dict(ref))

    return result
