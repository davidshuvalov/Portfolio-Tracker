"""
One-time importer for existing Portfolio Tracker v1.24 .xlsb workbooks.

Reads the Strategies tab and margin reference tables, extracting
configuration so customers don't need to re-enter settings manually.

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

TradeStation Margins sheet column layout (0-based):
    Col 0: Product Description
    Col 1: Symbol Root (TS symbol)
    Col 2: Intraday Initial
    Col 3: Intraday Maintenance
    Col 4: Long Overnight Margin  (initial)
    Col 5: Short Overnight Margin (initial)
    Col 6: Long Maintenance Margin
    Col 7: Short Maintenance Margin
    Col 8: Intraday Rate

InteractiveBrokers Margins sheet column layout (0-based):
    Col 0: Trading Class (IB symbol)
    Col 1: Intraday Initial
    Col 2: Overnight Initial
    Col 3: Overnight Maintenance
    Col 4: Short Overnight Initial
    Col 5: Short Overnight Maintenance
    Col 6: Currency

TS Symbol Lookup sheet column layout (0-based):
    Col 0: Product Description
    Col 1: TradeStation Code
    Col 2: Interactive Brokers Code

Margin computation mirrors VBA LookupMarginRequirements() in F_Summary_Tab_Setup.bas:
    TS Initial     = avg(Long Overnight,       Short Overnight)
    TS Maintenance = avg(Long Maintenance,     Short Maintenance)
    IB Initial     = avg(Overnight Initial,    Short Overnight Initial)
    IB Maintenance = avg(Overnight Maintenance, Short Overnight Maintenance)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

try:
    import pyxlsb
    PYXLSB_AVAILABLE = True
except ImportError:
    PYXLSB_AVAILABLE = False

CONFIG_DIR = Path.home() / ".portfolio_tracker"
MARGIN_TABLES_FILE = CONFIG_DIR / "margin_tables.yaml"


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

STRATEGIES_SHEET  = "Strategies"
TS_MARGINS_SHEET  = "TradeStation Margins"
IB_MARGINS_SHEET  = "InteractiveBrokers Margins"
SYMBOL_LOOKUP_SHEET = "TS Symbol Lookup"


# ── Margin tables dataclass ───────────────────────────────────────────────────

@dataclass
class MarginEntry:
    """Margin values for a single futures symbol."""
    description: str = ""
    initial: float = 0.0
    maintenance: float = 0.0


@dataclass
class MarginTables:
    """
    Reference tables imported from the v1.24 xlsb workbook.

    ts:     TS symbol  → MarginEntry (using TradeStation overnight values)
    ib:     IB symbol  → MarginEntry (using Interactive Brokers overnight values)
    lookup: TS symbol  → IB symbol  (from TS Symbol Lookup sheet)
    """
    ts: dict[str, MarginEntry] = field(default_factory=dict)
    ib: dict[str, MarginEntry] = field(default_factory=dict)
    lookup: dict[str, str] = field(default_factory=dict)

    def get_margin(
        self,
        ts_symbol: str,
        source: str = "TradeStation",
        margin_type: str = "Maintenance",
    ) -> float | None:
        """
        Return dollar margin for a TS symbol.

        Args:
            ts_symbol:   Symbol as used by TradeStation (e.g. "ES", "NK").
            source:      "TradeStation" | "InteractiveBrokers"
            margin_type: "Initial" | "Maintenance"

        Returns:
            Dollar margin, or None if the symbol is not in the reference data.
        """
        key = margin_type.lower()  # "initial" | "maintenance"
        if source == "TradeStation":
            entry = self.ts.get(ts_symbol)
        else:
            ib_sym = self.lookup.get(ts_symbol, ts_symbol)
            entry = self.ib.get(ib_sym)
        if entry is None:
            return None
        return getattr(entry, key, None)

    def resolve_for_symbols(
        self,
        ts_symbols: list[str],
        source: str = "TradeStation",
        margin_type: str = "Maintenance",
    ) -> dict[str, float]:
        """
        Return {ts_symbol: margin} for a list of TS symbols.
        Symbols not found in the reference data are skipped.
        """
        result: dict[str, float] = {}
        for sym in ts_symbols:
            val = self.get_margin(sym, source, margin_type)
            if val is not None and val > 0:
                result[sym] = val
        return result


# ── Import margin tables from xlsb ───────────────────────────────────────────

def import_margin_tables(xlsb_path: Path) -> MarginTables:
    """
    Read TradeStation Margins, InteractiveBrokers Margins, and TS Symbol Lookup
    sheets from a v1.24 .xlsb file.

    Margin calculations mirror VBA LookupMarginRequirements() in
    F_Summary_Tab_Setup.bas.

    Raises:
        ImportError   — pyxlsb not installed
        FileNotFoundError — xlsb file not found
        RuntimeError  — unexpected read error
    """
    if not PYXLSB_AVAILABLE:
        raise ImportError("pyxlsb is required. Install with: pip install pyxlsb")
    if not xlsb_path.exists():
        raise FileNotFoundError(f"File not found: {xlsb_path}")

    ts_margins: dict[str, MarginEntry] = {}
    ib_margins: dict[str, MarginEntry] = {}
    symbol_lookup: dict[str, str] = {}

    try:
        with pyxlsb.open_workbook(str(xlsb_path)) as wb:
            available = wb.sheets

            # ── TradeStation Margins ──────────────────────────────────────
            if TS_MARGINS_SHEET in available:
                with wb.get_sheet(TS_MARGINS_SHEET) as ws:
                    for i, row in enumerate(ws.rows()):
                        if i == 0:
                            continue  # header
                        vals = [c.v for c in row]
                        if len(vals) < 8:
                            continue
                        sym = _str(vals[1])
                        if not sym:
                            continue
                        # Skip section-header rows (all numeric cols are None)
                        long_overnight  = _float(vals[4])
                        short_overnight = _float(vals[5])
                        long_maint      = _float(vals[6])
                        short_maint     = _float(vals[7])
                        if long_overnight is None and long_maint is None:
                            continue  # section header like "Index", "Energy"
                        initial     = _avg(long_overnight, short_overnight)
                        maintenance = _avg(long_maint, short_maint)
                        if initial is None and maintenance is None:
                            continue
                        ts_margins[sym] = MarginEntry(
                            description=_str(vals[0]),
                            initial=initial or 0.0,
                            maintenance=maintenance or 0.0,
                        )

            # ── InteractiveBrokers Margins ────────────────────────────────
            if IB_MARGINS_SHEET in available:
                with wb.get_sheet(IB_MARGINS_SHEET) as ws:
                    for i, row in enumerate(ws.rows()):
                        if i == 0:
                            continue  # header
                        vals = [c.v for c in row]
                        if len(vals) < 6:
                            continue
                        sym = _str(vals[0])
                        if not sym:
                            continue
                        overnight_init     = _float(vals[2])
                        overnight_maint    = _float(vals[3])
                        short_init         = _float(vals[4])
                        short_maint        = _float(vals[5])
                        initial     = _avg(overnight_init, short_init)
                        maintenance = _avg(overnight_maint, short_maint)
                        if initial is None and maintenance is None:
                            continue
                        ib_margins[sym] = MarginEntry(
                            initial=initial or 0.0,
                            maintenance=maintenance or 0.0,
                        )

            # ── TS Symbol Lookup ──────────────────────────────────────────
            if SYMBOL_LOOKUP_SHEET in available:
                with wb.get_sheet(SYMBOL_LOOKUP_SHEET) as ws:
                    for i, row in enumerate(ws.rows()):
                        if i == 0:
                            continue  # header
                        vals = [c.v for c in row]
                        if len(vals) < 3:
                            continue
                        ts_sym = _str(vals[1])
                        ib_sym = _str(vals[2])
                        if ts_sym and ib_sym:
                            symbol_lookup[ts_sym] = ib_sym

    except (ImportError, FileNotFoundError):
        raise
    except Exception as e:
        raise RuntimeError(f"Failed to read margin tables from {xlsb_path.name}: {e}") from e

    return MarginTables(ts=ts_margins, ib=ib_margins, lookup=symbol_lookup)


# ── Persist / load margin tables ──────────────────────────────────────────────

def save_margin_tables(tables: MarginTables) -> None:
    """Save margin tables to ~/.portfolio_tracker/margin_tables.yaml."""
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    data = {
        "ts": {
            sym: {"description": e.description, "initial": e.initial, "maintenance": e.maintenance}
            for sym, e in tables.ts.items()
        },
        "ib": {
            sym: {"initial": e.initial, "maintenance": e.maintenance}
            for sym, e in tables.ib.items()
        },
        "lookup": tables.lookup,
    }
    with open(MARGIN_TABLES_FILE, "w") as f:
        yaml.dump(data, f, default_flow_style=False, sort_keys=False)


def load_margin_tables() -> MarginTables | None:
    """
    Load margin tables from ~/.portfolio_tracker/margin_tables.yaml.
    Returns None if the file does not exist.
    """
    if not MARGIN_TABLES_FILE.exists():
        return None
    try:
        with open(MARGIN_TABLES_FILE) as f:
            data = yaml.safe_load(f) or {}
        ts = {
            sym: MarginEntry(
                description=d.get("description", ""),
                initial=float(d.get("initial", 0)),
                maintenance=float(d.get("maintenance", 0)),
            )
            for sym, d in data.get("ts", {}).items()
        }
        ib = {
            sym: MarginEntry(
                initial=float(d.get("initial", 0)),
                maintenance=float(d.get("maintenance", 0)),
            )
            for sym, d in data.get("ib", {}).items()
        }
        lookup = {str(k): str(v) for k, v in data.get("lookup", {}).items()}
        return MarginTables(ts=ts, ib=ib, lookup=lookup)
    except Exception:
        return None


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


def _float(val) -> float | None:
    if val is None:
        return None
    try:
        f = float(val)
        return f if f > 0 else None
    except (ValueError, TypeError):
        return None


def _avg(a: float | None, b: float | None) -> float | None:
    """Average of two values, ignoring None. Returns None if both are None."""
    vals = [v for v in (a, b) if v is not None]
    return sum(vals) / len(vals) if vals else None
