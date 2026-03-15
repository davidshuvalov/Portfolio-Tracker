"""
Folder scanner — mirrors C_Retrieve_Folder_Locations.bas.

Scans MultiWalk base folders for strategy subfolders, validates required
files exist, detects duplicates, and returns a ScanResult.

Expected folder structure (from VBA analysis):
    {base_folder}/
    └── {strategy_subfolder}/
        └── Walkforward Files/
            ├── {StrategyName} EquityData.csv      ← REQUIRED
            ├── {StrategyName} TradeData.csv        ← optional
            └── Walkforward In-Out Periods Analysis Details.csv  ← optional
"""

from __future__ import annotations
import re
from pathlib import Path

from core.data_types import ScanResult, StrategyFolder

# ── Auto-population helpers ───────────────────────────────────────────────────

# Common futures symbol → sector mapping
_SYMBOL_SECTOR: dict[str, str] = {
    # Indices
    "ES": "Index", "NQ": "Index", "YM": "Index", "RTY": "Index", "EMD": "Index",
    "MES": "Index", "MNQ": "Index", "MYM": "Index", "M2K": "Index",
    # Interest Rates
    "ZB": "Interest Rate", "ZN": "Interest Rate", "ZF": "Interest Rate",
    "ZT": "Interest Rate", "ZQ": "Interest Rate", "UB": "Interest Rate",
    "TN": "Interest Rate", "TU": "Interest Rate", "FV": "Interest Rate",
    "TY": "Interest Rate", "US": "Interest Rate",
    # Energy
    "CL": "Energy", "NG": "Energy", "RB": "Energy", "HO": "Energy",
    "BZ": "Energy", "QM": "Energy",
    # Metals
    "GC": "Metals", "SI": "Metals", "HG": "Metals", "PL": "Metals", "PA": "Metals",
    "MGC": "Metals",
    # Currencies
    "6E": "Currencies", "6J": "Currencies", "6B": "Currencies", "6A": "Currencies",
    "6C": "Currencies", "6S": "Currencies", "6N": "Currencies", "6M": "Currencies",
    "DX": "Currencies",
    # Agriculture
    "ZC": "Agriculture", "ZW": "Agriculture", "ZS": "Agriculture",
    "ZL": "Agriculture", "ZM": "Agriculture", "ZO": "Agriculture", "ZR": "Agriculture",
    # Softs
    "KC": "Soft", "CT": "Soft", "CC": "Soft", "SB": "Soft", "OJ": "Soft",
    # Meats
    "LE": "Meats", "GF": "Meats", "HE": "Meats",
    # Crypto
    "BTC": "Crypto", "ETH": "Crypto", "MBT": "Crypto", "MET": "Crypto",
    # Volatility
    "VX": "Volatility",
    # Eurex
    "FDAX": "Eurex Index", "FESX": "Eurex Index", "FDXM": "Eurex Index",
    "FGBL": "Eurex Interest Rate", "FGBM": "Eurex Interest Rate",
    "FGBS": "Eurex Interest Rate", "FGBX": "Eurex Interest Rate",
}


def parse_name_parts(name: str) -> tuple[str, str]:
    """
    Extract symbol and timeframe from a MultiWalk strategy name.
    Expects the pattern [@SYMBOL-TIMEFRAME], e.g. [@TU-60min] → ('TU', '60min').
    Returns ('', '') if pattern not found.
    """
    m = re.search(r"\[@([A-Z0-9]+)-([^\]]+)\]", name)
    if m:
        return m.group(1), m.group(2)
    return "", ""

WALKFORWARD_DIR = "Walkforward Files"
EQUITY_SUFFIX = " EquityData.csv"
TRADE_SUFFIX = " TradeData.csv"
WALKFORWARD_DETAILS = "Walkforward In-Out Periods Analysis Details.csv"


def scan_folders(base_folders: list[Path]) -> ScanResult:
    """
    Scan all base folders for strategy subfolders.

    Mirrors VBA RetrieveAllFolderData + GetFolderData logic:
    - Scans each base folder for subfolders
    - In each subfolder looks for 'Walkforward Files/' subdirectory
    - Validates EquityData.csv exists (required)
    - Notes missing TradeData.csv or Walkforward Details.csv (warnings only)
    - Detects duplicate strategy names across base folders
    """
    strategies: list[StrategyFolder] = []
    warnings: list[str] = []
    errors: list[str] = []
    seen_names: dict[str, Path] = {}   # name → first base folder (duplicate detection)

    for base in base_folders:
        if not base.exists():
            errors.append(f"Base folder not found: {base}")
            continue
        if not base.is_dir():
            errors.append(f"Path is not a directory: {base}")
            continue

        for subfolder in sorted(base.iterdir()):
            if not subfolder.is_dir():
                continue

            result = _scan_strategy_folder(subfolder, warnings)
            if result is None:
                continue

            # Duplicate detection (mirrors VBA dict.Exists check)
            if result.name in seen_names:
                warnings.append(
                    f"Duplicate strategy '{result.name}' found in '{base.name}' "
                    f"(already seen in '{seen_names[result.name].name}'). Skipping duplicate."
                )
                continue

            seen_names[result.name] = base
            strategies.append(result)

    return ScanResult(strategies=strategies, warnings=warnings, errors=errors)


def _scan_strategy_folder(
    subfolder: Path, warnings: list[str]
) -> StrategyFolder | None:
    """
    Inspect one strategy subfolder.
    Returns StrategyFolder if valid, None if EquityData.csv is missing.
    """
    wf_dir = subfolder / WALKFORWARD_DIR

    if not wf_dir.exists():
        # No Walkforward Files directory — skip silently (could be an unrelated folder)
        return None

    # Find EquityData.csv — required
    equity_csv = _find_file(wf_dir, EQUITY_SUFFIX)
    if equity_csv is None:
        warnings.append(
            f"'{subfolder.name}': no *EquityData.csv found in '{WALKFORWARD_DIR}/'. "
            f"Skipping."
        )
        return None

    # Derive strategy name from filename: "{name} EquityData.csv" → "{name}"
    strategy_name = equity_csv.name.removesuffix(EQUITY_SUFFIX)

    # Find TradeData.csv — optional
    trade_csv = _find_file(wf_dir, TRADE_SUFFIX)
    if trade_csv is None:
        warnings.append(
            f"'{strategy_name}': no *TradeData.csv found. "
            f"Trade-level analysis will not be available."
        )

    # Find Walkforward Details CSV — optional
    wf_details = wf_dir / WALKFORWARD_DETAILS
    if not wf_details.exists():
        warnings.append(
            f"'{strategy_name}': '{WALKFORWARD_DETAILS}' not found. "
            f"IS/OOS dates will not be available from this folder."
        )
        wf_details = None

    return StrategyFolder(
        name=strategy_name,
        path=subfolder,
        equity_csv=equity_csv,
        trade_csv=trade_csv,
        walkforward_csv=wf_details,
    )


def _find_file(directory: Path, suffix: str) -> Path | None:
    """Find the first file in directory whose name ends with suffix."""
    for f in directory.iterdir():
        if f.is_file() and f.name.endswith(suffix):
            return f
    return None


# ── "Not Loaded" prefix helpers (mirrors VBA logic) ───────────────────────────

NOT_LOADED_PREFIX = "Not Loaded - "


def apply_not_loaded_prefix(status: str) -> str:
    """Add 'Not Loaded - ' prefix if not already present."""
    if NOT_LOADED_PREFIX not in status:
        return NOT_LOADED_PREFIX + status
    return status


def strip_not_loaded_prefix(status: str) -> str:
    """Remove 'Not Loaded - ' prefix if present."""
    return status.replace(NOT_LOADED_PREFIX, "")


def reconcile_statuses(
    found_names: set[str],
    configured_strategies: list[dict],
) -> list[dict]:
    """
    Cross-reference found strategy folders against the configured strategies list.

    - Strategies found in folders but not in config → added with status 'New'
    - Strategies in config but not found in folders → prefixed with 'Not Loaded - '
    - Strategies in both → 'Not Loaded - ' prefix stripped if present

    Mirrors VBA GetStrategyStatus + CheckMissingStrategies logic.

    Args:
        found_names: set of strategy names discovered by scan_folders
        configured_strategies: list of strategy dicts from strategies config

    Returns updated list of strategy dicts.
    """
    configured_names = {s["name"] for s in configured_strategies}
    result = []

    # Update existing configured strategies
    for s in configured_strategies:
        name = s["name"]
        updated = dict(s)
        if name in found_names:
            # Found — strip Not Loaded prefix if present
            updated["status"] = strip_not_loaded_prefix(updated.get("status", ""))
        else:
            # Missing from folders — apply Not Loaded prefix
            updated["status"] = apply_not_loaded_prefix(updated.get("status", ""))
        result.append(updated)

    # Add newly discovered strategies not yet in config
    for name in sorted(found_names - configured_names):
        sym, tf = parse_name_parts(name)
        sector = _SYMBOL_SECTOR.get(sym, "")
        result.append({
            "name": name,
            "status": "New",
            "contracts": 1,
            "symbol": sym,
            "sector": sector,
            "timeframe": tf,
            "type": "",
            "horizon": "",
            "other": "",
            "notes": "",
        })

    # Back-fill symbol/timeframe/sector for existing entries that are still empty
    for s in result:
        if not s.get("symbol") or not s.get("timeframe"):
            sym, tf = parse_name_parts(s["name"])
            if not s.get("symbol"):
                s["symbol"] = sym
            if not s.get("timeframe"):
                s["timeframe"] = tf
        if not s.get("sector") and s.get("symbol"):
            s["sector"] = _SYMBOL_SECTOR.get(s["symbol"], "")

    return result
