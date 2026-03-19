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
    # ── TradeStation legacy symbol aliases ────────────────────────────────────
    # Interest Rates (TS codes map to standard CBOT symbols)
    # FV/TY/US/TU already present above; duplicated here for clarity
    # Currencies — TS uses two-letter codes instead of CME "6X" convention
    "EC": "Currencies",   # Euro FX          (CME: 6E)
    "AD": "Currencies",   # Australian Dollar (CME: 6A)
    "JY": "Currencies",   # Japanese Yen      (CME: 6J)
    "BP": "Currencies",   # British Pound     (CME: 6B)
    "CD": "Currencies",   # Canadian Dollar   (CME: 6C)
    "SF": "Currencies",   # Swiss Franc       (CME: 6S)
    "NE1": "Currencies",  # New Zealand Dollar (CME: 6N)
    # Agriculture — TS uses single-letter codes for CBOT grains
    "C":  "Agriculture",  # Corn              (CBOT: ZC)
    "S":  "Agriculture",  # Soybeans          (CBOT: ZS)
    "W":  "Agriculture",  # Chicago SRW Wheat (CBOT: ZW)
    "KW": "Agriculture",  # Hard Red Winter Wheat (CBOT: KE)
    # Micro Ag (launched Feb 2025)
    "MZC": "Agriculture", "MZS": "Agriculture", "MZW": "Agriculture",
    # Micro FX (direct-pair micros — same quote direction as standard contract)
    "M6E": "Currencies", "M6A": "Currencies", "M6J": "Currencies",
    "M6B": "Currencies", "MCD": "Currencies", "MSF": "Currencies",
    "M6N": "Currencies",
    # Micro Energy
    "MCL": "Energy",
    # Micro/Mini Interest Rates
    "MWN": "Interest Rate",   # Micro Ultra T-Bond (0.1× UB)
    "MTN": "Interest Rate",   # Micro Ultra 10-Year (0.1× TN)
    # Eurex DAX variants
    "FDXS": "Eurex Index",    # Micro-DAX (1/25 FDAX)
    # Agriculture mini
    "MKC": "Agriculture",     # Mini HRW Wheat (0.2× KE)
}


def parse_name_parts(name: str) -> tuple[str, str]:
    """
    Extract symbol and timeframe from a MultiWalk strategy name.
    Expects the pattern [@SYMBOL-TIMEFRAME], e.g. [@TU-60min] → ('TU', '60min').
    The symbol may include a session suffix like '.D' (day session) which is stripped.
    Returns ('', '') if pattern not found.
    """
    m = re.search(r"\[@([A-Z0-9]+)(?:\.[A-Z]+)?-([^\]]+)\]", name)
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
    - One subfolder can yield multiple strategies (one per *EquityData.csv found),
      mirroring the VBA Dir() loop in GetFolderData.
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

            for result in _scan_strategy_folder(subfolder, warnings, base_folder=base):
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
    subfolder: Path, warnings: list[str], base_folder: Path | None = None
) -> list[StrategyFolder]:
    """
    Inspect one strategy subfolder.

    Returns a list of StrategyFolder — one per *EquityData.csv found in the
    Walkforward Files directory.  A single subfolder can yield multiple strategies,
    which is the normal layout for Buy & Hold folders where many instruments share
    one Walkforward Files directory.  Mirrors the VBA Dir() loop in GetFolderData.
    """
    wf_dir = subfolder / WALKFORWARD_DIR

    if not wf_dir.exists():
        # No Walkforward Files directory — skip silently (could be an unrelated folder)
        return []

    # Find ALL EquityData.csv files — each becomes a separate strategy (sorted for
    # deterministic ordering, mirroring VBA's Dir() alphabetical enumeration).
    equity_csvs = sorted(
        f for f in wf_dir.iterdir() if f.is_file() and f.name.endswith(EQUITY_SUFFIX)
    )

    if not equity_csvs:
        warnings.append(
            f"'{subfolder.name}': no *EquityData.csv found in '{WALKFORWARD_DIR}/'. "
            f"Skipping."
        )
        return []

    # Find optional Walkforward Details CSV — shared across all strategies in this
    # subfolder (there is typically one per folder, not one per instrument).
    wf_details_path = wf_dir / WALKFORWARD_DETAILS
    wf_details: Path | None = wf_details_path if wf_details_path.exists() else None
    if wf_details is None:
        warnings.append(
            f"'{subfolder.name}': '{WALKFORWARD_DETAILS}' not found. "
            f"IS/OOS dates will not be available from this folder."
        )

    results: list[StrategyFolder] = []
    for equity_csv in equity_csvs:
        strategy_name = equity_csv.name.removesuffix(EQUITY_SUFFIX)

        # Find the matching TradeData.csv for this specific strategy by name.
        trade_csv_path = wf_dir / (strategy_name + TRADE_SUFFIX)
        trade_csv: Path | None = trade_csv_path if trade_csv_path.exists() else None
        if trade_csv is None:
            warnings.append(
                f"'{strategy_name}': no *TradeData.csv found. "
                f"Trade-level analysis will not be available."
            )

        results.append(StrategyFolder(
            name=strategy_name,
            path=subfolder,
            equity_csv=equity_csv,
            trade_csv=trade_csv,
            walkforward_csv=wf_details,
            base_folder=base_folder,
        ))

    return results


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
    strategy_folders: list[StrategyFolder] | None = None,
    folder_default_status: dict[str, str] | None = None,
) -> list[dict]:
    """
    Cross-reference found strategy folders against the configured strategies list.

    - Strategies found in folders but not in config → added with folder default status (or 'New')
    - Strategies in config but not found in folders → prefixed with 'Not Loaded - '
    - Strategies in both → 'Not Loaded - ' prefix stripped if present

    Mirrors VBA GetStrategyStatus + CheckMissingStrategies logic.

    Args:
        found_names:            set of strategy names discovered by scan_folders or import
        configured_strategies:  list of strategy dicts from strategies config
        strategy_folders:       optional list of StrategyFolder (for base_folder lookup)
        folder_default_status:  optional dict of base_folder_path_str → default status

    Returns updated list of strategy dicts.
    """
    # Build lookup: strategy name → default status from its folder
    _name_to_default: dict[str, str] = {}
    if strategy_folders and folder_default_status:
        for sf in strategy_folders:
            if sf.base_folder is not None:
                default = folder_default_status.get(str(sf.base_folder))
                if default:
                    _name_to_default[sf.name] = default

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
        default_status = _name_to_default.get(name, "New")
        result.append({
            "name": name,
            "status": default_status,
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
