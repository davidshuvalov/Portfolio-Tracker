"""
One-time importer for existing Portfolio Tracker v1.24 .xlsb workbooks.

Reads the Strategies tab and extracts configuration so customers don't
need to re-enter their strategy settings manually.

Uses pyxlsb (read-only, no Excel installation required).

Strategies tab column layout (from E_ColumnConstants.bas, 1-based):
    Col 1: Strategy Number
    Col 2: Status
    Col 3: Strategy Name
    Col 4: Contracts
    Col 5: Symbol
    Col 6: Timeframe
    Col 7: Type
    Col 8: Horizon
    Col 9: Other
    Col 10: ClosedTradeMC
    Col 11: Folder
    Col 12: Notes
"""

from __future__ import annotations
from pathlib import Path

try:
    import pyxlsb
    PYXLSB_AVAILABLE = True
except ImportError:
    PYXLSB_AVAILABLE = False


# Column indices (0-based) matching VBA 1-based positions
COL_STATUS = 1
COL_NAME = 2
COL_CONTRACTS = 3
COL_SYMBOL = 4
COL_TIMEFRAME = 5
COL_TYPE = 6
COL_HORIZON = 7
COL_OTHER = 8
COL_NOTES = 11

STRATEGIES_SHEET = "Strategies"


def import_strategies_from_xlsb(xlsb_path: Path) -> tuple[list[dict], list[str]]:
    """
    Read the Strategies tab from a v1.24 .xlsb file.

    Returns:
        (strategies, warnings)
        strategies — list of dicts with keys matching Strategy dataclass fields
        warnings   — list of non-fatal issues encountered
    """
    if not PYXLSB_AVAILABLE:
        raise ImportError(
            "pyxlsb is required for .xlsb import. "
            "Install it with: pip install pyxlsb"
        )

    if not xlsb_path.exists():
        raise FileNotFoundError(f"File not found: {xlsb_path}")

    strategies: list[dict] = []
    warnings: list[str] = []

    try:
        with pyxlsb.open_workbook(str(xlsb_path)) as wb:
            sheet_names = wb.sheets
            if STRATEGIES_SHEET not in sheet_names:
                raise ValueError(
                    f"Sheet '{STRATEGIES_SHEET}' not found in workbook. "
                    f"Available sheets: {sheet_names}"
                )

            with wb.get_sheet(STRATEGIES_SHEET) as ws:
                rows = list(ws.rows())

            if len(rows) < 2:
                warnings.append("Strategies tab has no data rows.")
                return strategies, warnings

            # Skip header row (row 0)
            for row_idx, row in enumerate(rows[1:], start=2):
                values = [cell.v for cell in row]

                # Pad row to expected length
                while len(values) <= max(COL_NAME, COL_STATUS, COL_NOTES):
                    values.append(None)

                name = _str(values[COL_NAME])
                if not name:
                    continue  # Skip blank rows

                status = _str(values[COL_STATUS])
                if not status:
                    warnings.append(f"Row {row_idx}: strategy '{name}' has no status.")

                strategies.append({
                    "name": name,
                    "status": status or "Live",
                    "contracts": _int(values[COL_CONTRACTS]),
                    "symbol": _str(values[COL_SYMBOL]),
                    "timeframe": _str(values[COL_TIMEFRAME]),
                    "type": _str(values[COL_TYPE]),
                    "horizon": _str(values[COL_HORIZON]),
                    "other": _str(values[COL_OTHER]),
                    "notes": _str(values[COL_NOTES]),
                    "sector": "",  # Not in Strategies tab — user fills in v2
                })

    except ValueError:
        raise
    except Exception as e:
        raise RuntimeError(f"Failed to read {xlsb_path.name}: {e}") from e

    return strategies, warnings


# ── Helpers ───────────────────────────────────────────────────────────────────

def _str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


def _int(val, default: int = 1) -> int:
    try:
        return max(1, int(float(str(val))))
    except (ValueError, TypeError):
        return default
