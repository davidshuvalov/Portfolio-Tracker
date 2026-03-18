"""
Strategy Detail — full drill-down for one strategy.

Mirrors the individual strategy sheets in the spreadsheet.

Sections:
  1. Metadata header (status, symbol, contracts, dates…)
  2. IS vs OOS side-by-side metric cards
  3. Equity curve with IS / OOS shading
  4. Drawdown chart
  5. Monthly PnL heatmap
  6. Trade list (if available)
"""

from __future__ import annotations

import os
import platform
import subprocess
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.config import AppConfig
from core.data_types import ImportedData, PortfolioData

st.set_page_config(page_title="Strategy Detail", layout="wide")

config: AppConfig = st.session_state.get("config", AppConfig.load())
imported: ImportedData | None = st.session_state.get("imported_data")
portfolio: PortfolioData | None = st.session_state.get("portfolio_data")

if imported is None:
    st.info("No data loaded yet.")
    st.page_link("ui/pages/01_Import.py", label="Go to Import →")
    st.stop()

all_strategy_names = sorted(imported.strategy_names)

# ── Sidebar — strategy selector ────────────────────────────────────────────────
with st.sidebar:
    st.header("Strategy")

    # Pre-select whatever was set by the calling page
    preselected = st.session_state.get("selected_strategy")
    default_idx = (
        all_strategy_names.index(preselected)
        if preselected and preselected in all_strategy_names
        else 0
    )

    selected_name = st.selectbox(
        "Choose strategy",
        all_strategy_names,
        index=default_idx,
    )
    st.session_state.selected_strategy = selected_name

    # Navigate back
    st.divider()
    st.page_link("ui/pages/02_Strategies.py", label="← Strategies")
    st.page_link("ui/pages/03_Portfolio.py",  label="← Portfolio")

# ── Resolve strategy metadata ──────────────────────────────────────────────────
# Try portfolio first (has full Strategy objects), fall back to imported
_strat_obj = None
if portfolio:
    _strat_obj = next((s for s in portfolio.strategies if s.name == selected_name), None)
if _strat_obj is None:
    _strat_obj = next((s for s in imported.strategies if s.name == selected_name), None)

# Pull summary row (metrics computed by compute_summary)
_summary_row: pd.Series | None = None
if portfolio is not None and not portfolio.summary_metrics.empty:
    if selected_name in portfolio.summary_metrics.index:
        _summary_row = portfolio.summary_metrics.loc[selected_name]
# Also check all-strategies cache (from Strategies page)
if _summary_row is None:
    _all_sm = st.session_state.get("all_strategies_summary_cache")
    if _all_sm is not None and not _all_sm.empty and selected_name in _all_sm.index:
        _summary_row = _all_sm.loc[selected_name]

def _sm(key: str, default=None):
    """Safe getter from summary row."""
    if _summary_row is None:
        return default
    val = _summary_row.get(key, default)
    try:
        is_nan = not isinstance(val, (str, bool)) and pd.isna(val)
    except Exception:
        is_nan = False
    return default if is_nan else val


# ── Collect key dates ──────────────────────────────────────────────────────────
is_start  = _strat_obj.is_start  if _strat_obj else None
is_end    = _strat_obj.is_end    if _strat_obj else None
oos_start = _strat_obj.oos_start if _strat_obj else None
oos_end   = _strat_obj.oos_end   if _strat_obj else None

# Fall back to summary row dates
if oos_start is None:
    _ob = _sm("oos_begin")
    if _ob is not None:
        try:
            oos_start = pd.Timestamp(_ob).date()
        except Exception:
            pass

# ── Title ─────────────────────────────────────────────────────────────────────
st.title(selected_name)

# Fall back to strategies.yaml config when imported.strategies stubs are empty
from core.portfolio.strategies import load_strategies as _load_strats
from core.ingestion.folder_scanner import parse_name_parts as _parse_name_parts, _SYMBOL_SECTOR as _SYM_SEC
_configured_meta = {s["name"]: s for s in _load_strats()}.get(selected_name, {})

def _meta(field: str, obj_val: str) -> str:
    """Return obj_val if non-empty, else fall back to strategies.yaml config."""
    if obj_val:
        return obj_val
    return _configured_meta.get(field, "") or ""

status    = _meta("status",    (_strat_obj.status    if _strat_obj else "")) or "—"
symbol    = _meta("symbol",    (_strat_obj.symbol    if _strat_obj else ""))
timeframe = _meta("timeframe", (_strat_obj.timeframe if _strat_obj else ""))
sector    = _meta("sector",    (_strat_obj.sector    if _strat_obj else ""))
s_type    = _meta("type",      (_strat_obj.type      if _strat_obj else "")) or "—"
horizon   = _meta("horizon",   (_strat_obj.horizon   if _strat_obj else "")) or "—"
contracts = int(_strat_obj.contracts if _strat_obj and _strat_obj.contracts else
                _configured_meta.get("contracts", 1) or 1)
notes     = _meta("notes",     (_strat_obj.notes     if _strat_obj else "")) or ""

# Last-resort: parse symbol/timeframe from strategy name if still empty
if not symbol or not timeframe:
    _parsed_sym, _parsed_tf = _parse_name_parts(selected_name)
    if not symbol:
        symbol = _parsed_sym
    if not timeframe:
        timeframe = _parsed_tf
# Last-resort: sector from symbol lookup
if not sector and symbol:
    sector = _SYM_SEC.get(symbol, "")

symbol    = symbol    or "—"
timeframe = timeframe or "—"
sector    = sector    or "—"

# Metadata strip
_direction = _sm("direction", "") or "—"

# Eligibility: use pre-computed value if available, otherwise derive it now
_raw_elig = _sm("eligibility_status", None)
if _raw_elig is not None and str(_raw_elig) not in ("", "nan"):
    _elig_status = str(_raw_elig)
elif _summary_row is not None:
    from core.portfolio.summary import apply_eligibility_rules
    _elig_df = pd.DataFrame([_summary_row], index=[selected_name])
    _elig_mask = apply_eligibility_rules(_elig_df, config.eligibility)
    _elig_status = "Eligible" if bool(_elig_mask.iloc[0]) else "Ineligible"
else:
    _elig_status = "—"
_next_opt = _sm("next_opt_date")
_last_opt = _sm("last_opt_date")
_next_opt_str = str(_next_opt) if _next_opt else "—"
_last_opt_str = str(_last_opt) if _last_opt else "—"

mc1, mc2, mc3, mc4, mc5 = st.columns(5)
mc1.metric("Status",      status)
mc2.metric("Symbol",      symbol)
mc3.metric("Sector",      sector)
mc4.metric("Direction",   _direction)
mc5.metric("Eligibility", _elig_status)

mc6, mc7, mc8, mc9, mc10 = st.columns(5)
mc6.metric("Contracts",   contracts)
mc7.metric("Timeframe",   timeframe)
mc8.metric("Type",        s_type)
mc9.metric("Next Opt",    _next_opt_str)
mc10.metric("Last Opt",   _last_opt_str)

if notes:
    st.caption(f"Notes: {notes}")

st.divider()

# ── Build daily PnL series (1 contract, raw) ──────────────────────────────────
if selected_name not in imported.daily_m2m.columns:
    st.warning(f"No daily PnL data found for **{selected_name}**.")
    st.stop()

raw_pnl = imported.daily_m2m[selected_name].dropna()

# Contract-scaled version
scaled_pnl = raw_pnl * contracts

# ── IS / OOS splits ────────────────────────────────────────────────────────────
oos_ts = pd.Timestamp(oos_start) if oos_start else None

is_pnl  = scaled_pnl[scaled_pnl.index <  oos_ts] if oos_ts is not None else scaled_pnl
oos_pnl = scaled_pnl[scaled_pnl.index >= oos_ts] if oos_ts is not None else pd.Series(dtype=float)

def _metrics_for(pnl: pd.Series, label: str) -> dict:
    if pnl.empty:
        return {}
    eq = pnl.cumsum()
    peak = eq.cummax()
    dd = peak - eq
    n_years = max((pnl.index[-1] - pnl.index[0]).days / 365.25, 1e-3)
    total = float(pnl.sum())
    ann = total / n_years
    max_dd = float(dd.max())
    monthly = pnl.resample("ME").sum()
    win_rate = float((monthly > 0).mean()) if len(monthly) > 0 else 0.0
    std_m = float(monthly.std()) if len(monthly) > 1 else 0.0
    sharpe = (float(monthly.mean()) / std_m * np.sqrt(12)) if std_m > 1e-9 else 0.0
    rtd = (total / max_dd) if max_dd > 0 else 0.0
    return {
        "label":     label,
        "total":     total,
        "ann":       ann,
        "max_dd":    max_dd,
        "win_rate":  win_rate,
        "sharpe":    sharpe,
        "rtd":       rtd,
        "start":     pnl.index[0].date(),
        "end":       pnl.index[-1].date(),
        "n_days":    len(pnl),
    }

is_m  = _metrics_for(is_pnl,  "In-Sample (IS)")
oos_m = _metrics_for(oos_pnl, "Out-of-Sample (OOS)")

# ── IS / OOS metric cards ──────────────────────────────────────────────────────
def _metric_block(m: dict, colour: str) -> None:
    if not m:
        st.info("No data for this period.")
        return
    st.markdown(
        f"<div style='background:{colour};padding:10px 14px;border-radius:8px;"
        f"margin-bottom:6px;font-size:0.8rem;color:#555;'>"
        f"{m['start']} → {m['end']} &nbsp;·&nbsp; {m['n_days']} days</div>",
        unsafe_allow_html=True,
    )
    r1c1, r1c2, r1c3 = st.columns(3)
    r1c1.metric("Total P&L",  f"${m['total']:,.0f}")
    r1c2.metric("Ann. P&L",   f"${m['ann']:,.0f}")
    r1c3.metric("Max DD",     f"${m['max_dd']:,.0f}")
    r2c1, r2c2, r2c3 = st.columns(3)
    r2c1.metric("Win Rate",   f"{m['win_rate']:.1%}")
    r2c2.metric("Sharpe",     f"{m['sharpe']:.2f}")
    r2c3.metric("R:DD",       f"{m['rtd']:.2f}")

col_is, col_oos = st.columns(2)
with col_is:
    st.subheader("In-Sample")
    _metric_block(is_m, "#e3f2fd")
with col_oos:
    st.subheader("Out-of-Sample")
    _metric_block(oos_m, "#e8f5e9")

# Supplementary OOS metrics from walkforward summary
if _summary_row is not None:
    with st.expander("Walkforward Detail Metrics", expanded=False):
        exp_ann  = _sm("expected_annual_profit", 0)
        act_ann  = _sm("actual_annual_profit",   0)
        eff      = _sm("return_efficiency",       0)
        wf_rtd   = _sm("rtd_oos",                0)
        wf_dd    = _sm("max_oos_drawdown",        0)
        incub    = _sm("incubation_status",       "—")
        trades_y = _sm("trades_per_year",         0)
        win_r    = _sm("overall_win_rate",        0)

        wc1, wc2, wc3, wc4 = st.columns(4)
        wc1.metric("Exp. Annual ($)", f"${exp_ann:,.0f}" if exp_ann else "—",
                   help="IS Annualized Net Profit from WF CSV (expected OOS rate)")
        wc2.metric("WF Act. Annual ($)", f"${act_ann:,.0f}" if act_ann else "—",
                   help="(IS+OOS Change in Net Profit) ÷ OOS years, from WF CSV")
        wc3.metric("Efficiency",      f"{eff:.1%}"       if eff      else "—")
        wc4.metric("Incubation",      str(incub))
        wc5, wc6, wc7, _ = st.columns(4)
        wc5.metric("Trades/Year",     f"{trades_y:.1f}"  if trades_y else "—")
        wc6.metric("Win Rate (all)",  f"{win_r:.1%}"     if win_r    else "—")
        wc7.metric("OOS Max DD ($)",  f"${wf_dd:,.0f}"   if wf_dd    else "—")

_sd_exp_col, _ = st.columns([1, 5])
with _sd_exp_col:
    if st.button("Export to Excel", key="sd_export_btn"):
        from core.reporting.excel_export import (
            export_strategy_detail,
            strategy_detail_export_filename,
        )
        _xlsx = export_strategy_detail(
            selected_name, scaled_pnl, oos_start, _summary_row
        )
        st.download_button(
            "📥 Download",
            data=_xlsx,
            file_name=strategy_detail_export_filename(selected_name),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_sd_xlsx",
        )

st.divider()

# ── Equity curve ───────────────────────────────────────────────────────────────
st.subheader("Equity Curve")

fig_eq = go.Figure()

if not is_pnl.empty:
    is_eq = is_pnl.cumsum()
    fig_eq.add_trace(go.Scatter(
        x=is_eq.index, y=is_eq.values,
        name="IS", line=dict(color="#1565C0", width=2),
    ))

if not oos_pnl.empty:
    # Carry IS end value as OOS base so the curve is continuous
    is_base = float(is_pnl.sum()) if not is_pnl.empty else 0.0
    oos_eq  = oos_pnl.cumsum() + is_base
    fig_eq.add_trace(go.Scatter(
        x=oos_eq.index, y=oos_eq.values,
        name="OOS", line=dict(color="#2E7D32", width=2.5),
    ))

# OOS start vertical line
if oos_ts is not None:
    fig_eq.add_vline(
        x=int(oos_ts.timestamp() * 1000), line_dash="dash", line_color="#B71C1C",
        annotation_text="OOS Start", annotation_position="top right",
    )

fig_eq.update_layout(
    height=380, xaxis_title="Date", yaxis_title="Cumulative P&L ($)",
    hovermode="x unified",
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
)
st.plotly_chart(fig_eq, use_container_width=True)

# ── Drawdown ───────────────────────────────────────────────────────────────────
with st.expander("Drawdown", expanded=False):
    full_eq = scaled_pnl.cumsum()
    peak    = full_eq.cummax()
    dd_ser  = -(peak - full_eq)

    fig_dd = go.Figure()
    if oos_ts is not None:
        fig_dd.add_trace(go.Scatter(
            x=dd_ser[dd_ser.index < oos_ts].index,
            y=dd_ser[dd_ser.index < oos_ts].values,
            fill="tozeroy", name="IS DD",
            line=dict(color="#90CAF9"),
        ))
        fig_dd.add_trace(go.Scatter(
            x=dd_ser[dd_ser.index >= oos_ts].index,
            y=dd_ser[dd_ser.index >= oos_ts].values,
            fill="tozeroy", name="OOS DD",
            line=dict(color="#F44336"),
        ))
        fig_dd.add_vline(x=int(oos_ts.timestamp() * 1000), line_dash="dash", line_color="#B71C1C")
    else:
        fig_dd.add_trace(go.Scatter(
            x=dd_ser.index, y=dd_ser.values,
            fill="tozeroy", name="Drawdown",
            line=dict(color="#F44336"),
        ))

    fig_dd.update_layout(
        height=240, yaxis_title="Drawdown ($)", hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    st.plotly_chart(fig_dd, use_container_width=True)

st.divider()

# ── Monthly PnL heatmap ────────────────────────────────────────────────────────
st.subheader("Monthly P&L")

monthly_pnl = scaled_pnl.resample("ME").sum()
if not monthly_pnl.empty:
    mdf = pd.DataFrame({
        "year":  monthly_pnl.index.year,
        "month": monthly_pnl.index.month,
        "pnl":   monthly_pnl.values,
    })
    pivot = mdf.pivot(index="year", columns="month", values="pnl")
    pivot.columns = [
        ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][c - 1]
        for c in pivot.columns
    ]
    pivot = pivot.sort_index(ascending=False)

    # Label each year as IS or OOS based on oos_start
    _oos_year = oos_start.year if oos_start else None
    _oos_month = oos_start.month if oos_start else None
    def _year_label(yr: int) -> str:
        if _oos_year is None:
            return str(yr)
        if yr > _oos_year:
            return f"{yr} OOS"
        if yr < _oos_year:
            return f"{yr} IS"
        # Same year as OOS start — label based on which half dominates
        if _oos_month is not None and _oos_month <= 6:
            return f"{yr} OOS"
        return f"{yr} IS"

    pivot.index = [_year_label(y) for y in pivot.index]

    fig_hm = px.imshow(
        pivot,
        color_continuous_scale="RdYlGn",
        color_continuous_midpoint=0,
        text_auto=".0f",
        aspect="auto",
    )
    fig_hm.update_layout(
        height=max(200, len(pivot) * 40 + 80),
        coloraxis_showscale=False,
    )
    # Add a horizontal line between the last IS year and first OOS year
    if _oos_year is not None:
        _year_labels_ordered = list(pivot.index)  # already sorted descending
        _oos_rows = [i for i, lbl in enumerate(_year_labels_ordered) if "OOS" in lbl]
        _is_rows  = [i for i, lbl in enumerate(_year_labels_ordered) if "IS"  in lbl]
        if _oos_rows and _is_rows:
            # The boundary is between the last OOS row and first IS row (descending order)
            _boundary = max(_oos_rows) + 0.5
            fig_hm.add_hline(
                y=_boundary,
                line_dash="dash",
                line_color="#B71C1C",
                line_width=2,
                annotation_text="← OOS start",
                annotation_position="right",
            )
    st.plotly_chart(fig_hm, use_container_width=True)

    # Annual totals
    annual = monthly_pnl.resample("YE").sum()
    ann_df = pd.DataFrame({"Year": annual.index.year, "Total P&L ($)": annual.values.round(0).astype(int)})
    st.dataframe(ann_df.sort_values("Year", ascending=False), hide_index=True, use_container_width=True)

# ── Files & Folder ─────────────────────────────────────────────────────────────
# Mirror of VBA H_Open_Code_Tab.bas — open strategy folder or code files

def _open_path(path: Path) -> None:
    """Open a file or folder in the OS default application (local desktop app)."""
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as exc:
        st.error(f"Could not open: {exc}")

# Find the strategy's folder from scan_result or Strategy.folder
_folder_path: Path | None = None
_scan_result = st.session_state.get("scan_result")
if _scan_result:
    for _sf in _scan_result.strategies:
        if _sf.name == selected_name:
            _folder_path = _sf.path
            break
if _folder_path is None and _strat_obj is not None and _strat_obj.folder:
    _folder_path = Path(_strat_obj.folder)

# Known MultiWalk / MultiCharts code file extensions
_CODE_EXTENSIONS = {".mex", ".eld", ".els", ".pla", ".c", ".cpp", ".py"}

def _is_code_file(f: Path) -> bool:
    if f.suffix.lower() in _CODE_EXTENSIONS:
        return True
    # MultiWalk EasyLanguage export: "... ELCode.txt"
    if f.suffix.lower() == ".txt" and f.name.endswith("ELCode.txt"):
        return True
    return False

_code_files: list[Path] = []
_data_files: dict[str, Path] = {}
if _folder_path and _folder_path.exists():
    for _f in sorted(_folder_path.iterdir()):
        if _is_code_file(_f):
            _code_files.append(_f)
        elif _f.suffix.lower() == ".csv":
            _data_files[_f.name] = _f

with st.expander("Files & Folder", expanded=False):
    if _folder_path is None:
        st.info("Folder path not available. Import data first.")
    else:
        # Folder path (copy-paste friendly)
        st.caption("Strategy folder")
        st.code(str(_folder_path), language=None)
        if _folder_path.exists():
            if st.button("📂 Open Folder", key="open_folder_btn"):
                _open_path(_folder_path)
        else:
            st.warning(f"Folder not found on disk: `{_folder_path}`")

        # Code files
        if _code_files:
            st.caption("Code files")
            for _cf in _code_files:
                col_name, col_btn = st.columns([5, 1])
                col_name.markdown(f"`{_cf.name}`")
                if col_btn.button("Open", key=f"open_code_{_cf.name}"):
                    _open_path(_cf)
        else:
            st.caption("No code files (.mex / .eld / .pla / .els / ELCode.txt) found in this folder.")

        # Data files summary
        if _data_files:
            st.caption("Data files (CSV)")
            for _name, _path in _data_files.items():
                try:
                    size_kb = _path.stat().st_size / 1024
                    st.markdown(f"- `{_name}` — {size_kb:.0f} KB")
                except Exception:
                    st.markdown(f"- `{_name}`")

st.divider()

# ── Trade list ─────────────────────────────────────────────────────────────────
if not imported.trades.empty and "strategy" in imported.trades.columns:
    strat_trades = imported.trades[imported.trades["strategy"] == selected_name].copy()
    if not strat_trades.empty:
        with st.expander(f"Trade List ({len(strat_trades)} trades)", expanded=False):
            strat_trades = strat_trades.drop(columns=["strategy"], errors="ignore")
            if "pnl" in strat_trades.columns:
                strat_trades = strat_trades.sort_values("date", ascending=False) if "date" in strat_trades.columns else strat_trades
                def _trade_style(row):
                    try:
                        v = float(row.get("pnl", 0))
                        return ["color: #2e7d32" if v > 0 else "color: #c62828"] * len(row)
                    except Exception:
                        return [""] * len(row)
                st.dataframe(
                    strat_trades.style.apply(_trade_style, axis=1),
                    hide_index=True,
                    use_container_width=True,
                    height=min(500, len(strat_trades) * 36 + 40),
                )
