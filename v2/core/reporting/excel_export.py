"""
Excel export — generates .xlsx reports from in-memory portfolio data.

Functions:
  export_portfolio(portfolio_data, config)   -> bytes   Computed-output report
  export_raw_data(imported_data)             -> bytes   Raw imported DataFrames
  export_mc_result(mc_result, label)         -> bytes   Monte Carlo results
  export_loo_result(loo_result, base_*args)  -> bytes   Leave-One-Out table
  export_correlations(corr_df, mode)         -> bytes   Correlation matrix

Requires openpyxl (listed in pyproject.toml dependencies).
"""

from __future__ import annotations

import io
from datetime import date
from typing import TYPE_CHECKING

import pandas as pd

if TYPE_CHECKING:
    from core.config import AppConfig
    from core.data_types import ImportedData, MCResult, PortfolioData


# ── Style helpers ─────────────────────────────────────────────────────────────

def _make_workbook():
    try:
        from openpyxl import Workbook
        return Workbook()
    except ImportError:
        raise ImportError(
            "openpyxl is required for Excel export. "
            "Install it with: pip install openpyxl"
        )


def _header_style(ws, row: int = 1) -> None:
    from openpyxl.styles import Font, PatternFill, Alignment
    fill = PatternFill("solid", fgColor="1565C0")
    font = Font(bold=True, color="FFFFFF")
    for cell in ws[row]:
        cell.font = font
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = ws.cell(row=row + 1, column=1)


def _autofit(ws) -> None:
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=8)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)


def _df_to_sheet(ws, df: pd.DataFrame, title: str | None = None) -> None:
    from openpyxl.styles import Font
    row_offset = 0
    if title:
        ws.append([title])
        ws.cell(row=1, column=1).font = Font(bold=True, size=13)
        ws.append([])
        row_offset = 2

    ws.append(list(df.columns))
    _header_style(ws, row=row_offset + 1)

    for _, row in df.iterrows():
        ws.append([v.item() if hasattr(v, "item") else v for v in row])

    _autofit(ws)


def _save(wb) -> bytes:
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Computed-output portfolio export ─────────────────────────────────────────

def export_portfolio(
    portfolio_data: "PortfolioData",
    config: "AppConfig",
) -> bytes:
    """
    Export computed portfolio analytics to .xlsx.

    Sheets:
      Summary         — summary_metrics (80+ columns) for every strategy
      Portfolio Equity — cumulative portfolio P&L curve
      Export Info     — metadata
    """
    wb = _make_workbook()
    wb.remove(wb.active)

    # ── Summary sheet ─────────────────────────────────────────────────────────
    ws_summary = wb.create_sheet("Summary")
    summary = portfolio_data.summary_metrics

    if not summary.empty:
        # Prepend Status + Contracts from Strategy objects
        status_map    = {s.name: s.status    for s in portfolio_data.strategies}
        contracts_map = {s.name: s.contracts for s in portfolio_data.strategies}
        df = summary.copy().reset_index().rename(columns={"index": "Strategy"})
        df.insert(1, "Status",    [status_map.get(n, "")    for n in df["Strategy"]])
        df.insert(2, "Contracts", [contracts_map.get(n, 1) for n in df["Strategy"]])

        # Convert date columns to strings for Excel compatibility
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime("%Y-%m-%d")
            elif df[col].dtype == object:
                # Try to detect date objects
                try:
                    sample = df[col].dropna().iloc[0]
                    if hasattr(sample, "strftime"):
                        df[col] = df[col].apply(lambda v: v.strftime("%Y-%m-%d") if v is not None and hasattr(v, "strftime") else v)
                except (IndexError, TypeError):
                    pass

        _df_to_sheet(ws_summary, df, title="Strategy Summary")
    else:
        ws_summary.append(["No strategy metrics available."])

    # ── Portfolio equity curve ────────────────────────────────────────────────
    ws_equity = wb.create_sheet("Portfolio Equity")
    if not portfolio_data.daily_pnl.empty:
        port_pnl = portfolio_data.daily_pnl.sum(axis=1)
        port_equity = port_pnl.cumsum()
        eq_df = pd.DataFrame({
            "Date":       port_equity.index.strftime("%Y-%m-%d"),
            "Daily PnL ($)":  port_pnl.values.round(2),
            "Equity ($)": port_equity.values.round(2),
        })
        _df_to_sheet(ws_equity, eq_df, title="Portfolio Equity Curve")
    else:
        ws_equity.append(["No portfolio data available."])

    # ── Metadata ─────────────────────────────────────────────────────────────
    ws_meta = wb.create_sheet("Export Info")
    ws_meta.append(["Portfolio Tracker v2 — Computed Output Export"])
    ws_meta.append(["Generated",         str(date.today())])
    ws_meta.append(["Live strategies",   len(portfolio_data.strategies)])
    ws_meta.append(["Period (years)",    config.portfolio.period_years])
    _autofit(ws_meta)

    return _save(wb)


# ── Raw data export ───────────────────────────────────────────────────────────

def export_raw_data(imported_data: "ImportedData") -> bytes:
    """
    Export the raw imported DataFrames to .xlsx.

    Sheets:
      Daily M2M     — daily mark-to-market PnL (dates × strategies)
      Closed Trades — daily closed-trade PnL   (dates × strategies)
      In Market Long
      In Market Short
      Trades        — individual trade records
    """
    wb = _make_workbook()
    wb.remove(wb.active)

    def _matrix_to_sheet(ws, df: pd.DataFrame, title: str) -> None:
        out = df.copy()
        out.index = out.index.strftime("%Y-%m-%d")
        out = out.reset_index().rename(columns={"index": "Date"})
        out = out.round(2)
        _df_to_sheet(ws, out, title=title)

    ws1 = wb.create_sheet("Daily M2M")
    _matrix_to_sheet(ws1, imported_data.daily_m2m, "Daily Mark-to-Market PnL")

    ws2 = wb.create_sheet("Closed Trades")
    _matrix_to_sheet(ws2, imported_data.closed_trade_pnl, "Daily Closed-Trade PnL")

    ws3 = wb.create_sheet("In Market Long")
    _matrix_to_sheet(ws3, imported_data.in_market_long, "In-Market Long PnL")

    ws4 = wb.create_sheet("In Market Short")
    _matrix_to_sheet(ws4, imported_data.in_market_short, "In-Market Short PnL")

    ws5 = wb.create_sheet("Trades")
    if not imported_data.trades.empty:
        trades = imported_data.trades.copy()
        if pd.api.types.is_datetime64_any_dtype(trades.get("date", pd.Series(dtype="object"))):
            trades["date"] = trades["date"].dt.strftime("%Y-%m-%d")
        trades = trades.round(2)
        _df_to_sheet(ws5, trades, title="Individual Trade Records")
    else:
        ws5.append(["No trade-level data available."])

    ws_meta = wb.create_sheet("Export Info")
    start, end = imported_data.date_range
    ws_meta.append(["Portfolio Tracker v2 — Raw Data Export"])
    ws_meta.append(["Generated",    str(date.today())])
    ws_meta.append(["Strategies",   len(imported_data.strategy_names)])
    ws_meta.append(["Date Range",   f"{start} → {end}"])
    ws_meta.append(["Trading Days", len(imported_data.daily_m2m)])
    ws_meta.append(["Trades",       len(imported_data.trades)])
    _autofit(ws_meta)

    return _save(wb)


# ── Monte Carlo export ────────────────────────────────────────────────────────

def export_mc_result(mc_result: "MCResult", label: str = "Portfolio") -> bytes:
    """Export MC summary + scenario distribution to .xlsx."""
    wb = _make_workbook()
    wb.remove(wb.active)

    ws_summary = wb.create_sheet("MC Summary")
    ws_summary.append(["Portfolio Tracker v2 — Monte Carlo Results"])
    ws_summary.append(["Target",              label])
    ws_summary.append(["Generated",           str(date.today())])
    ws_summary.append([])
    ws_summary.append(["Metric", "Value"])
    from openpyxl.styles import Font
    ws_summary.cell(row=5, column=1).font = Font(bold=True)
    ws_summary.cell(row=5, column=2).font = Font(bold=True)

    rows = [
        ("Starting Equity ($)",    round(mc_result.starting_equity, 2)),
        ("Expected Annual Profit ($)", round(mc_result.expected_profit, 2)),
        ("Risk of Ruin",           f"{mc_result.risk_of_ruin:.3%}"),
        ("Max Drawdown (median)",  f"{mc_result.max_drawdown_pct:.3%}"),
        ("Sharpe Ratio",           round(mc_result.sharpe_ratio, 3)),
        ("Return / Drawdown",      round(mc_result.return_to_drawdown, 3)),
    ]
    for r in rows:
        ws_summary.append(list(r))
    _autofit(ws_summary)

    if mc_result.scenarios_df is not None and not mc_result.scenarios_df.empty:
        ws_scenarios = wb.create_sheet("Scenario Distribution")
        df = mc_result.scenarios_df.copy().round(4)
        # Format pct columns
        for col in df.columns:
            if "pct" in col.lower():
                df[col] = df[col].apply(lambda v: f"{v:.2%}")
        _df_to_sheet(ws_scenarios, df, title="Monte Carlo Scenarios")

    return _save(wb)


# ── Leave-One-Out export ──────────────────────────────────────────────────────

def export_loo_result(
    loo_result: "pd.DataFrame",
    base_profit: float = 0.0,
    base_sharpe: float = 0.0,
) -> bytes:
    """Export LOO analysis table to .xlsx."""
    wb = _make_workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet("Leave-One-Out")
    ws.append(["Portfolio Tracker v2 — Leave-One-Out Analysis"])
    ws.append(["Generated",         str(date.today())])
    ws.append(["Base Exp. Profit",  round(base_profit, 2)])
    ws.append(["Base Sharpe",       round(base_sharpe, 3)])
    ws.append([])

    df = loo_result.copy()
    for col in df.select_dtypes(include="number").columns:
        df[col] = df[col].round(4)

    _df_to_sheet(ws, df, title=None)
    _autofit(ws)

    return _save(wb)


# ── Correlations export ───────────────────────────────────────────────────────

def export_correlations(corr_df: "pd.DataFrame", mode: str = "Normal") -> bytes:
    """Export a correlation matrix DataFrame to .xlsx."""
    wb = _make_workbook()
    wb.remove(wb.active)

    ws = wb.create_sheet(f"Correlations ({mode})")
    from openpyxl.styles import Font, PatternFill, Alignment

    header_fill = PatternFill("solid", fgColor="1565C0")
    header_font = Font(bold=True, color="FFFFFF")
    symbols = list(corr_df.columns)

    ws.cell(row=1, column=1).value = ""
    for col_idx, sym in enumerate(symbols, start=2):
        cell = ws.cell(row=1, column=col_idx, value=sym)
        cell.font  = header_font
        cell.fill  = header_fill
        cell.alignment = Alignment(horizontal="center")

    for row_idx, sym in enumerate(symbols, start=2):
        h = ws.cell(row=row_idx, column=1, value=sym)
        h.font  = header_font
        h.fill  = header_fill
        h.alignment = Alignment(horizontal="center")
        for col_idx, col_sym in enumerate(symbols, start=2):
            val = corr_df.loc[sym, col_sym]
            cell = ws.cell(row=row_idx, column=col_idx, value=round(float(val), 4))
            cell.number_format = "0.00"

    ws.freeze_panes = "B2"
    _autofit(ws)

    return _save(wb)


# ── Filename helpers ──────────────────────────────────────────────────────────

def portfolio_export_filename() -> str:
    return f"portfolio_output_{date.today()}.xlsx"

def raw_data_export_filename() -> str:
    return f"portfolio_raw_data_{date.today()}.xlsx"

def mc_export_filename(label: str = "portfolio") -> str:
    return f"mc_{label.lower().replace(' ', '_')}_{date.today()}.xlsx"

def loo_export_filename() -> str:
    return f"leave_one_out_{date.today()}.xlsx"

def correlations_export_filename(mode: str = "Normal") -> str:
    return f"correlations_{mode.lower()}_{date.today()}.xlsx"
