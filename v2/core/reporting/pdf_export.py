"""
PDF export — generates a portfolio summary report as a PDF.

Uses reportlab (pure Python, no binary dependencies).
Requires: pip install reportlab

Functions:
  export_portfolio_pdf(portfolio_data, config, mc_result=None) -> bytes
      Returns raw PDF bytes for download.
"""

from __future__ import annotations

import io
from datetime import date
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from core.config import AppConfig
    from core.data_types import MCResult, PortfolioData


# ── Page layout constants ────────────────────────────────────────────────────
_A4_W, _A4_H = 595, 842          # points (A4 portrait)
_LETTER_W, _LETTER_H = 612, 792  # US Letter

_BLUE  = (0.082, 0.392, 0.753)   # #1565C0
_GREEN = (0.184, 0.490, 0.196)   # #2e7d32
_RED   = (0.773, 0.102, 0.141)   # #c51922


def _make_canvas(buf: io.BytesIO, width: float, height: float):
    try:
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import cm
    except ImportError:
        raise ImportError(
            "reportlab is required for PDF export. "
            "Install it with: pip install reportlab"
        )
    return canvas.Canvas(buf, pagesize=(width, height))


def _rgb(r, g, b):
    from reportlab.lib.colors import Color
    return Color(r, g, b)


# ── Cover page ────────────────────────────────────────────────────────────────

def _cover_page(c, width: float, height: float, n_strategies: int, generated: str) -> None:
    # Background header band
    c.setFillColor(_rgb(*_BLUE))
    c.rect(0, height - 120, width, 120, fill=True, stroke=False)

    c.setFillColor(_rgb(1, 1, 1))
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(width / 2, height - 60, "Portfolio Tracker v2")
    c.setFont("Helvetica", 14)
    c.drawCentredString(width / 2, height - 90, "Portfolio Summary Report")

    c.setFillColor(_rgb(0.1, 0.1, 0.1))
    c.setFont("Helvetica", 11)
    c.drawCentredString(width / 2, height - 145, f"Generated: {generated}")
    c.drawCentredString(width / 2, height - 165, f"Active strategies: {n_strategies}")


# ── Section heading ───────────────────────────────────────────────────────────

def _section(c, x: float, y: float, title: str) -> float:
    c.setFillColor(_rgb(*_BLUE))
    c.rect(x, y - 4, 450, 18, fill=True, stroke=False)
    c.setFillColor(_rgb(1, 1, 1))
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x + 6, y, title)
    c.setFillColor(_rgb(0, 0, 0))
    return y - 28


# ── Metrics table ─────────────────────────────────────────────────────────────

def _metrics_table(
    c,
    x: float,
    y: float,
    rows: list[tuple[str, str]],
    col1_w: float = 220,
    col2_w: float = 180,
    row_h: float = 16,
) -> float:
    from reportlab.lib.colors import HexColor
    c.setFont("Helvetica-Bold", 9)
    c.setFillColor(_rgb(*_BLUE))
    c.rect(x, y, col1_w + col2_w, row_h, fill=True, stroke=False)
    c.setFillColor(_rgb(1, 1, 1))
    c.drawString(x + 4, y + 4, "Metric")
    c.drawString(x + col1_w + 4, y + 4, "Value")
    y -= row_h

    c.setFont("Helvetica", 9)
    for i, (metric, value) in enumerate(rows):
        bg = 0.93 if i % 2 == 0 else 1.0
        c.setFillColorRGB(bg, bg, bg)
        c.rect(x, y, col1_w + col2_w, row_h, fill=True, stroke=False)
        c.setFillColor(_rgb(0.1, 0.1, 0.1))
        c.drawString(x + 4, y + 4, str(metric))
        c.drawString(x + col1_w + 4, y + 4, str(value))
        y -= row_h

    return y - 8


# ── Strategy summary table ────────────────────────────────────────────────────

_STRAT_COLS = [
    ("Strategy",      120, "name"),
    ("Status",         50, "status"),
    ("Symbol",         40, "symbol"),
    ("Sector",         70, "sector"),
    ("Contracts",      55, "contracts"),
]

def _strategy_table(
    c,
    x: float,
    y: float,
    strategies,
    summary_df,
    page_width: float,
    page_height: float,
    margin: float = 50,
) -> None:
    row_h = 14.0
    header_h = 16.0
    min_y = margin + 20

    def _draw_header(y_pos):
        c.setFont("Helvetica-Bold", 8)
        c.setFillColor(_rgb(*_BLUE))
        total_w = sum(w for _, w, _ in _STRAT_COLS) + 80  # extra for OOS Profit col
        c.rect(x, y_pos, total_w, header_h, fill=True, stroke=False)
        c.setFillColor(_rgb(1, 1, 1))
        cx = x + 4
        for label, w, _ in _STRAT_COLS:
            c.drawString(cx, y_pos + 4, label)
            cx += w
        c.drawString(cx, y_pos + 4, "OOS Profit")
        return y_pos - header_h

    y = _draw_header(y)

    c.setFont("Helvetica", 8)
    for i, strat in enumerate(strategies):
        if y < min_y:
            c.showPage()
            y = page_height - margin - 20
            y = _draw_header(y)
            c.setFont("Helvetica", 8)

        bg = 0.93 if i % 2 == 0 else 1.0
        total_w = sum(w for _, w, _ in _STRAT_COLS) + 80
        c.setFillColorRGB(bg, bg, bg)
        c.rect(x, y, total_w, row_h, fill=True, stroke=False)
        c.setFillColor(_rgb(0.1, 0.1, 0.1))

        cx = x + 4
        for _, w, field in _STRAT_COLS:
            val = str(getattr(strat, field, ""))
            c.drawString(cx, y + 3, val[:int(w / 5.5)])  # rough char limit
            cx += w

        # OOS profit from summary_metrics if available
        oos_profit = ""
        if summary_df is not None and strat.name in summary_df.index:
            v = summary_df.loc[strat.name].get("profit_since_oos_start", None)
            if v is not None and not (hasattr(v, "__class__") and v.__class__.__name__ == "float" and str(v) == "nan"):
                try:
                    oos_profit = f"${float(v):,.0f}"
                except Exception:
                    pass
        c.drawString(cx, y + 3, oos_profit)
        y -= row_h


# ── Main export function ──────────────────────────────────────────────────────

def export_portfolio_pdf(
    portfolio_data: "PortfolioData",
    config: "AppConfig",
    mc_result: "MCResult | None" = None,
) -> bytes:
    """
    Generate a PDF portfolio summary report.

    Returns raw PDF bytes suitable for st.download_button.
    """
    import numpy as np

    buf = io.BytesIO()
    W, H = _LETTER_W, _LETTER_H
    margin = 50.0
    c = _make_canvas(buf, W, H)

    generated = date.today().isoformat()
    strategies = portfolio_data.strategies
    summary = portfolio_data.summary_metrics
    n = len(strategies)

    # ── Page 1: Cover + Portfolio Overview ───────────────────────────────────
    _cover_page(c, W, H, n, generated)

    y = H - 210
    y = _section(c, margin, y, "Portfolio Overview")

    # Compute headline metrics from daily_pnl
    port_pnl = portfolio_data.daily_pnl.sum(axis=1)
    total_pnl = port_pnl.sum()
    equity = port_pnl.cumsum()
    rolling_max = equity.cummax()
    dd = equity - rolling_max
    max_dd = dd.min()
    max_dd_pct = (max_dd / rolling_max[dd.idxmin()]) if not equity.empty else 0

    ann_return = 0.0
    n_years = len(port_pnl) / 252
    if n_years > 0 and port_pnl.std() > 0:
        sharpe = (port_pnl.mean() / port_pnl.std()) * (252 ** 0.5)
    else:
        sharpe = 0.0

    overview_rows = [
        ("Active strategies",        str(n)),
        ("Period (years)",            f"{config.portfolio.period_years:.1f}"),
        ("Total P&L",                 f"${total_pnl:,.0f}"),
        ("Max Drawdown",              f"{max_dd_pct:.1%}"),
        ("Sharpe Ratio (daily M2M)",  f"{sharpe:.2f}"),
    ]

    if not equity.empty:
        overview_rows.append(("Data range", f"{equity.index.min().date()} → {equity.index.max().date()}"))

    y = _metrics_table(c, margin, y, overview_rows)

    # ── MC results section ────────────────────────────────────────────────────
    if mc_result is not None:
        import math
        if y < 200:
            c.showPage()
            y = H - margin
        y = _section(c, margin, y, "Monte Carlo Summary")
        mc_rows = [
            ("Starting Equity",        f"${mc_result.starting_equity:,.0f}"),
            ("Expected Annual Profit", f"${mc_result.expected_profit:,.0f}"),
            ("Risk of Ruin",           f"{mc_result.risk_of_ruin:.1%}"),
            ("Max Drawdown (median)",  f"{mc_result.max_drawdown_pct:.1%}"),
            ("Sharpe Ratio",           f"{mc_result.sharpe_ratio:.2f}"),
            ("Return / Drawdown",      f"{mc_result.return_to_drawdown:.2f}"),
        ]
        y = _metrics_table(c, margin, y, mc_rows)

    # ── Strategy table ────────────────────────────────────────────────────────
    if y < 180:
        c.showPage()
        y = H - margin

    y -= 10
    y = _section(c, margin, y, "Strategy List")
    _strategy_table(c, margin, y, strategies, summary if not summary.empty else None, W, H, margin)

    c.save()
    return buf.getvalue()


# ── Filename helper ───────────────────────────────────────────────────────────

def pdf_export_filename() -> str:
    return f"portfolio_report_{date.today()}.pdf"
