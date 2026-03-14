"""
Excel export — generates .xlsx reports from in-memory portfolio data.

Functions:
  export_portfolio(portfolio_data, config) -> bytes
      Exports strategy metrics + summary to an .xlsx workbook.

  export_correlations(corr_df) -> bytes
      Exports a correlation matrix to an .xlsx workbook.

Requires openpyxl (listed in pyproject.toml dependencies).
"""

from __future__ import annotations

import io
from datetime import date
from typing import TYPE_CHECKING

import pandas as pd

if TYPE_CHECKING:
    from core.config import AppConfig
    from core.data_types import PortfolioData


# ── Helpers ───────────────────────────────────────────────────────────────────

def _make_workbook():
    """Return a new openpyxl Workbook (raises ImportError if not installed)."""
    try:
        from openpyxl import Workbook
        return Workbook()
    except ImportError:
        raise ImportError(
            "openpyxl is required for Excel export. "
            "Install it with: pip install openpyxl"
        )


def _apply_header_style(ws, row: int = 1) -> None:
    """Bold the header row and freeze panes below it."""
    from openpyxl.styles import Font, PatternFill, Alignment
    fill = PatternFill("solid", fgColor="1565C0")
    font = Font(bold=True, color="FFFFFF")
    for cell in ws[row]:
        cell.font = font
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def _autofit(ws) -> None:
    """Approximate column widths from content."""
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)


def _df_to_sheet(ws, df: pd.DataFrame, title: str | None = None) -> None:
    """Write a DataFrame to an openpyxl worksheet."""
    row_offset = 0
    if title:
        ws.append([title])
        from openpyxl.styles import Font
        ws.cell(row=1, column=1).font = Font(bold=True, size=13)
        row_offset = 1
        ws.append([])
        row_offset += 1

    # Header
    ws.append(list(df.columns))
    _apply_header_style(ws, row=row_offset + 1)

    # Data rows
    for _, row in df.iterrows():
        ws.append([
            v.item() if hasattr(v, "item") else v
            for v in row
        ])

    _autofit(ws)


# ── Portfolio export ──────────────────────────────────────────────────────────

def export_portfolio(
    portfolio_data: "PortfolioData",
    config: "AppConfig",
) -> bytes:
    """
    Export full portfolio metrics to an .xlsx workbook.

    Sheets:
      Summary      — one row per strategy with all computed metrics
      Portfolio    — aggregate portfolio equity curve (daily)

    Returns raw bytes suitable for st.download_button.
    """
    from core.portfolio.summary import compute_strategy_metrics

    wb = _make_workbook()
    wb.remove(wb.active)  # drop default empty sheet

    # ── Summary sheet ─────────────────────────────────────────────────────────
    rows = []
    for strat in portfolio_data.strategies:
        eq = portfolio_data.equity.get(strat.name)
        if eq is None or eq.empty:
            continue
        m = compute_strategy_metrics(eq, strat, config.portfolio)
        row = {
            "Strategy":       strat.name,
            "Status":         strat.status,
            "Symbol":         strat.symbol,
            "Sector":         strat.sector,
            "Timeframe":      strat.timeframe,
            "Contracts":      strat.contracts,
            "Total P&L ($)":  round(m.get("total_pnl", 0), 2),
            "Ann. Return (%)": round(m.get("annualised_return_pct", 0), 2),
            "Max DD (%)":     round(m.get("max_drawdown_pct", 0) * 100, 2),
            "Sharpe":         round(m.get("sharpe_ratio", 0), 3),
            "Win Rate (%)":   round(m.get("win_rate_pct", 0), 2),
            "Avg Trade ($)":  round(m.get("avg_trade", 0), 2),
            "Trades":         m.get("total_trades", 0),
            "IS Start":       str(m.get("is_start", "")),
            "OOS Start":      str(m.get("oos_start", "")),
        }
        rows.append(row)

    ws_summary = wb.create_sheet("Summary")
    if rows:
        _df_to_sheet(ws_summary, pd.DataFrame(rows), title="Strategy Summary")
    else:
        ws_summary.append(["No strategy data available."])

    # ── Portfolio equity curve sheet ─────────────────────────────────────────
    ws_equity = wb.create_sheet("Portfolio Equity")
    if portfolio_data.portfolio_equity is not None and not portfolio_data.portfolio_equity.empty:
        eq_df = portfolio_data.portfolio_equity.reset_index()
        eq_df.columns = ["Date", "Equity ($)"]
        eq_df["Date"] = eq_df["Date"].astype(str)
        _df_to_sheet(ws_equity, eq_df, title="Portfolio Equity Curve")
    else:
        ws_equity.append(["No portfolio equity data available."])

    # ── Metadata sheet ────────────────────────────────────────────────────────
    ws_meta = wb.create_sheet("Export Info")
    ws_meta.append(["Portfolio Tracker v2 Export"])
    ws_meta.append(["Generated", str(date.today())])
    ws_meta.append(["Strategies", len(portfolio_data.strategies)])
    ws_meta.append(["Period (years)", config.portfolio.period_years])

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Correlations export ───────────────────────────────────────────────────────

def export_correlations(corr_df: pd.DataFrame, mode: str = "Normal") -> bytes:
    """
    Export a correlation matrix DataFrame to .xlsx.

    Returns raw bytes suitable for st.download_button.
    """
    wb = _make_workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(f"Correlations ({mode})")

    from openpyxl.styles import Font, PatternFill, Alignment, numbers
    from openpyxl.utils import get_column_letter

    header_fill = PatternFill("solid", fgColor="1565C0")
    header_font = Font(bold=True, color="FFFFFF")

    symbols = list(corr_df.columns)

    # Top-left corner blank
    ws.cell(row=1, column=1).value = ""

    # Column headers
    for col_idx, sym in enumerate(symbols, start=2):
        cell = ws.cell(row=1, column=col_idx, value=sym)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Row headers + data
    for row_idx, sym in enumerate(symbols, start=2):
        header_cell = ws.cell(row=row_idx, column=1, value=sym)
        header_cell.font = header_font
        header_cell.fill = header_fill
        header_cell.alignment = Alignment(horizontal="center")

        for col_idx, col_sym in enumerate(symbols, start=2):
            val = corr_df.loc[sym, col_sym]
            cell = ws.cell(row=row_idx, column=col_idx, value=round(float(val), 4))
            # Colour-code: high positive = red, high negative = blue, near-zero = white
            from openpyxl.styles import numbers as num_styles
            cell.number_format = "0.00"

    ws.freeze_panes = "B2"
    _autofit(ws)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Filename helpers ──────────────────────────────────────────────────────────

def portfolio_export_filename() -> str:
    return f"portfolio_export_{date.today()}.xlsx"


def correlations_export_filename(mode: str = "Normal") -> str:
    return f"correlations_{mode.lower()}_{date.today()}.xlsx"
