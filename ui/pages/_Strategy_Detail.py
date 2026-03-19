"""
Strategy Detail — full drill-down for one strategy.

Mirrors the individual strategy sheets in the spreadsheet.

Tabs:
  1. Overview       — equity curve, monthly PnL heatmap
  2. WF Dashboard   — per-period IS vs OOS performance across all WF windows
  3. Health Monitor — RAG cockpit: days since opt, OOS P&L, losing months, drawdown
  4. Drawdown       — underwater curve, event table, duration histogram, rolling max DD
  5. Files & Trades — folder / code files, trade list
"""

from __future__ import annotations

import os
import platform
import subprocess
from datetime import date, timedelta
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

    st.divider()
    st.page_link("ui/pages/02_Strategies.py", label="← Strategies")
    st.page_link("ui/pages/03_Portfolio.py",  label="← Portfolio")

# ── Resolve strategy metadata ──────────────────────────────────────────────────
_strat_obj = None
if portfolio:
    _strat_obj = next((s for s in portfolio.strategies if s.name == selected_name), None)
if _strat_obj is None:
    _strat_obj = next((s for s in imported.strategies if s.name == selected_name), None)

_summary_row: pd.Series | None = None
if portfolio is not None and not portfolio.summary_metrics.empty:
    if selected_name in portfolio.summary_metrics.index:
        _summary_row = portfolio.summary_metrics.loc[selected_name]
if _summary_row is None:
    _all_sm = st.session_state.get("all_strategies_summary_cache")
    if _all_sm is not None and not _all_sm.empty and selected_name in _all_sm.index:
        _summary_row = _all_sm.loc[selected_name]

def _sm(key: str, default=None):
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

if oos_start is None:
    _ob = _sm("oos_begin")
    if _ob is not None:
        try:
            oos_start = pd.Timestamp(_ob).date()
        except Exception:
            pass

# ── Title ──────────────────────────────────────────────────────────────────────
st.title(selected_name)

from core.portfolio.strategies import load_strategies as _load_strats
from core.ingestion.folder_scanner import parse_name_parts as _parse_name_parts, _SYMBOL_SECTOR as _SYM_SEC
_configured_meta = {s["name"]: s for s in _load_strats()}.get(selected_name, {})

def _meta(field: str, obj_val: str) -> str:
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

if not symbol or not timeframe:
    _parsed_sym, _parsed_tf = _parse_name_parts(selected_name)
    if not symbol:
        symbol = _parsed_sym
    if not timeframe:
        timeframe = _parsed_tf
if not sector and symbol:
    sector = _SYM_SEC.get(symbol, "")

symbol    = symbol    or "—"
timeframe = timeframe or "—"
sector    = sector    or "—"

_direction = _sm("direction", "") or "—"

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

# ── Build daily PnL series ─────────────────────────────────────────────────────
if selected_name not in imported.daily_m2m.columns:
    st.warning(f"No daily PnL data found for **{selected_name}**.")
    st.stop()

raw_pnl    = imported.daily_m2m[selected_name].dropna()
scaled_pnl = raw_pnl * contracts

oos_ts  = pd.Timestamp(oos_start) if oos_start else None
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
        "label":    label,
        "total":    total,
        "ann":      ann,
        "max_dd":   max_dd,
        "win_rate": win_rate,
        "sharpe":   sharpe,
        "rtd":      rtd,
        "start":    pnl.index[0].date(),
        "end":      pnl.index[-1].date(),
        "n_days":   len(pnl),
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

# ── Locate WF CSV for this strategy ───────────────────────────────────────────
_wf_csv_path: Path | None = None
_folder_path: Path | None = None
_scan_result = st.session_state.get("scan_result")
if _scan_result:
    for _sf in _scan_result.strategies:
        if _sf.name == selected_name:
            _folder_path  = _sf.path
            _wf_csv_path  = _sf.walkforward_csv
            break
if _folder_path is None and _strat_obj is not None and _strat_obj.folder:
    _folder_path = Path(_strat_obj.folder)

# ── Load all WF periods (for WF Dashboard tab) ────────────────────────────────
_wf_periods: list[dict] = []
if _wf_csv_path is not None and _wf_csv_path.exists():
    from core.ingestion.walkforward_reader import read_all_walkforward_periods
    _wf_periods = read_all_walkforward_periods(
        _wf_csv_path,
        selected_name,
        date_format=config.date_format if hasattr(config, "date_format") else "DMY",
    )


# ══════════════════════════════════════════════════════════════════════════════
# TABS
# ══════════════════════════════════════════════════════════════════════════════
tab_overview, tab_wf, tab_health, tab_dd, tab_files = st.tabs([
    "📊 Overview",
    "🔄 Walk-Forward Dashboard",
    "🏥 Health Monitor",
    "📉 Drawdown Deep-Dive",
    "📁 Files & Trades",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════
with tab_overview:

    # Equity curve ─────────────────────────────────────────────────────────────
    st.subheader("Equity Curve")
    fig_eq = go.Figure()

    if not is_pnl.empty:
        is_eq = is_pnl.cumsum()
        fig_eq.add_trace(go.Scatter(
            x=is_eq.index, y=is_eq.values,
            name="IS", line=dict(color="#1565C0", width=2),
        ))

    if not oos_pnl.empty:
        is_base = float(is_pnl.sum()) if not is_pnl.empty else 0.0
        oos_eq  = oos_pnl.cumsum() + is_base
        fig_eq.add_trace(go.Scatter(
            x=oos_eq.index, y=oos_eq.values,
            name="OOS", line=dict(color="#2E7D32", width=2.5),
        ))

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

    # Monthly PnL heatmap ──────────────────────────────────────────────────────
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

        _oos_year  = oos_start.year  if oos_start else None
        _oos_month = oos_start.month if oos_start else None
        def _year_label(yr: int) -> str:
            if _oos_year is None:
                return str(yr)
            if yr > _oos_year:
                return f"{yr} OOS"
            if yr < _oos_year:
                return f"{yr} IS"
            return f"{yr} OOS" if (_oos_month is not None and _oos_month <= 6) else f"{yr} IS"

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

        if _oos_year is not None:
            _yr_labels = list(pivot.index)
            _oos_rows  = [i for i, lbl in enumerate(_yr_labels) if "OOS" in lbl]
            _is_rows   = [i for i, lbl in enumerate(_yr_labels) if "IS"  in lbl]
            if _oos_rows and _is_rows:
                _boundary = max(_oos_rows) + 0.5
                fig_hm.add_hline(
                    y=_boundary, line_dash="dash", line_color="#B71C1C", line_width=2,
                    annotation_text="← OOS start", annotation_position="right",
                )

        st.plotly_chart(fig_hm, use_container_width=True)

        annual = monthly_pnl.resample("YE").sum()
        ann_df = pd.DataFrame({
            "Year": annual.index.year,
            "Total P&L ($)": annual.values.round(0).astype(int),
        })
        st.dataframe(ann_df.sort_values("Year", ascending=False), hide_index=True, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — WALK-FORWARD DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
with tab_wf:
    st.subheader("Walk-Forward Performance Dashboard")
    st.caption("IS vs OOS across every optimization window — the primary curve-fit diagnostic.")

    if not _wf_periods:
        st.info(
            "No walkforward period data found. "
            "Import a **Walkforward Details CSV** (MultiWalk export) to enable this view."
        )
    else:
        # ── Build per-period stats ─────────────────────────────────────────────
        period_rows = []
        for i, p in enumerate(_wf_periods):
            oos_b = p["oos_begin"]
            oos_e = p["oos_end"]

            # Slice equity data to this OOS window
            if oos_b and oos_e:
                _mask = (scaled_pnl.index >= pd.Timestamp(oos_b)) & (scaled_pnl.index <= pd.Timestamp(oos_e))
            elif oos_b:
                _mask = scaled_pnl.index >= pd.Timestamp(oos_b)
            else:
                _mask = pd.Series(False, index=scaled_pnl.index)

            oos_slice = scaled_pnl[_mask]
            oos_years = max(len(oos_slice) / 252, 1e-3) if not oos_slice.empty else 1e-3
            oos_total = float(oos_slice.sum()) if not oos_slice.empty else 0.0
            oos_ann_actual = oos_total / oos_years

            is_ann  = p["is_ann_profit"]
            oos_ann_wf = p["isoos_change"] / oos_years if p["isoos_change"] != 0 else 0.0

            is_sh  = p["is_sharpe"]
            oos_sh = p["isoos_sharpe"]
            sh_ratio = oos_sh / max(is_sh, 1e-9)

            period_rows.append({
                "period":         f"P{i + 1}",
                "oos_begin":      str(oos_b) if oos_b else "—",
                "oos_end":        str(oos_e) if oos_e else "open",
                "is_ann":         is_ann,
                "oos_ann_actual": oos_ann_actual,
                "oos_ann_wf":     oos_ann_wf,
                "is_sharpe":      is_sh,
                "oos_sharpe":     oos_sh,
                "sh_ratio":       sh_ratio,
                "is_max_dd":      p["is_max_dd"],
                "isoos_max_dd":   p["isoos_max_dd"],
                "oos_eq":         oos_slice.cumsum() if not oos_slice.empty else pd.Series(dtype=float),
            })

        # ── Chart 1: OOS equity per WF window ─────────────────────────────────
        st.markdown("#### OOS Equity per Walk-Forward Window")
        st.caption("Each line shows the cumulative P&L within that OOS window, starting from zero.")

        fig_oos = go.Figure()
        palette = px.colors.qualitative.Set2
        for idx, pr in enumerate(period_rows):
            eq = pr["oos_eq"]
            if not eq.empty:
                eq_reset = eq - eq.iloc[0]
                fig_oos.add_trace(go.Scatter(
                    x=eq_reset.index, y=eq_reset.values,
                    name=pr["period"],
                    line=dict(color=palette[idx % len(palette)], width=2),
                    hovertemplate=f"{pr['period']} — %{{y:$,.0f}}<extra></extra>",
                ))

        fig_oos.update_layout(
            height=360, xaxis_title="Date", yaxis_title="OOS Cumulative P&L ($)",
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig_oos, use_container_width=True)

        # ── Chart 2: IS vs OOS annual return bar chart ────────────────────────
        st.markdown("#### IS vs OOS Annualised Return")
        st.caption("Bars show IS expected return (from WF CSV) vs actual OOS return from equity data.")

        periods_lbl = [pr["period"] for pr in period_rows]
        is_vals     = [pr["is_ann"]         for pr in period_rows]
        oos_vals    = [pr["oos_ann_actual"]  for pr in period_rows]

        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(
            name="IS Ann. Return",
            x=periods_lbl, y=is_vals,
            marker_color="#1565C0",
            text=[f"${v:,.0f}" for v in is_vals], textposition="outside",
        ))
        fig_bar.add_trace(go.Bar(
            name="OOS Ann. Return (actual)",
            x=periods_lbl, y=oos_vals,
            marker_color=[("#2E7D32" if v >= 0 else "#C62828") for v in oos_vals],
            text=[f"${v:,.0f}" for v in oos_vals], textposition="outside",
        ))
        fig_bar.add_hline(y=0, line_color="black", line_width=1)
        fig_bar.update_layout(
            barmode="group", height=360,
            xaxis_title="WF Period", yaxis_title="Annualised P&L ($)",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

        # ── Chart 3: Sharpe IS vs ISOOS degradation ───────────────────────────
        st.markdown("#### IS vs IS+OOS Sharpe Ratio — Degradation per Period")
        st.caption("OOS/IS Sharpe ratio < 1 indicates curve-fit degradation. Closer to 1 = more robust.")

        is_sh_vals  = [pr["is_sharpe"]  for pr in period_rows]
        oos_sh_vals = [pr["oos_sharpe"] for pr in period_rows]
        sh_ratios   = [pr["sh_ratio"]   for pr in period_rows]

        fig_sh = go.Figure()
        fig_sh.add_trace(go.Bar(
            name="IS Sharpe", x=periods_lbl, y=is_sh_vals,
            marker_color="#1565C0",
        ))
        fig_sh.add_trace(go.Bar(
            name="IS+OOS Sharpe", x=periods_lbl, y=oos_sh_vals,
            marker_color="#43A047",
        ))
        # Degradation ratio line on secondary y
        fig_sh.add_trace(go.Scatter(
            name="OOS/IS Ratio", x=periods_lbl, y=sh_ratios,
            mode="lines+markers",
            line=dict(color="#FF6F00", width=2, dash="dot"),
            marker=dict(size=8),
            yaxis="y2",
        ))
        fig_sh.update_layout(
            barmode="group", height=360,
            xaxis_title="WF Period", yaxis_title="Sharpe Ratio",
            yaxis2=dict(title="OOS/IS Ratio", overlaying="y", side="right",
                        range=[0, max(max(sh_ratios, default=2), 2)]),
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        fig_sh.add_hline(y=1.0, line_color="#FF6F00", line_dash="dash",
                         line_width=1, yref="y2",
                         annotation_text="Ratio = 1", annotation_position="right")
        st.plotly_chart(fig_sh, use_container_width=True)

        # ── Period summary table ───────────────────────────────────────────────
        st.markdown("#### Period Summary Table")
        tbl = pd.DataFrame([{
            "Period":           pr["period"],
            "OOS Start":        pr["oos_begin"],
            "OOS End":          pr["oos_end"],
            "IS Ann. ($)":      f"${pr['is_ann']:,.0f}",
            "OOS Ann. ($)":     f"${pr['oos_ann_actual']:,.0f}",
            "IS Sharpe":        f"{pr['is_sharpe']:.2f}",
            "IS+OOS Sharpe":    f"{pr['oos_sharpe']:.2f}",
            "Sharpe Ratio":     f"{pr['sh_ratio']:.2f}",
            "IS Max DD ($)":    f"${pr['is_max_dd']:,.0f}",
            "IS+OOS Max DD ($)": f"${pr['isoos_max_dd']:,.0f}",
        } for pr in period_rows])
        st.dataframe(tbl, hide_index=True, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — HEALTH MONITOR
# ══════════════════════════════════════════════════════════════════════════════
with tab_health:
    st.subheader("Strategy Health Monitor")
    st.caption("Daily cockpit check — RAG status across key OOS health indicators.")

    TODAY = date.today()

    # ── Helper: RAG badge ──────────────────────────────────────────────────────
    def _rag(status_str: str) -> str:
        colours = {"GREEN": "#2E7D32", "AMBER": "#F57F17", "RED": "#C62828"}
        icons   = {"GREEN": "●", "AMBER": "●", "RED": "●"}
        c = colours.get(status_str, "#888")
        i = icons.get(status_str, "●")
        return f"<span style='color:{c};font-size:1.4rem;'>{i}</span> **{status_str}**"

    # ── Metric 1: Days since last optimisation ─────────────────────────────────
    _last_opt_date = _sm("last_opt_date")
    if _last_opt_date is not None:
        try:
            _last_opt_dt = pd.Timestamp(_last_opt_date).date()
            _days_since  = (TODAY - _last_opt_dt).days
        except Exception:
            _last_opt_dt = None
            _days_since  = None
    else:
        _last_opt_dt = None
        _days_since  = None

    if _days_since is None:
        _opt_rag = "GREEN"
        _opt_detail = "Last opt date not recorded"
    elif _days_since > 180:
        _opt_rag = "RED"
        _opt_detail = f"{_days_since} days since last optimisation (> 6 months)"
    elif _days_since > 90:
        _opt_rag = "AMBER"
        _opt_detail = f"{_days_since} days since last optimisation (> 3 months)"
    else:
        _opt_rag = "GREEN"
        _opt_detail = f"{_days_since} days since last optimisation"

    # ── Metric 2: OOS P&L vs Expected ─────────────────────────────────────────
    _exp_ann = _sm("expected_annual_profit", 0) or 0.0
    if not oos_m or _exp_ann == 0:
        _pnl_rag = "GREEN"
        _pnl_detail = "Expected annual profit not set or no OOS data"
        _oos_efficiency = None
    else:
        _oos_years = max(oos_m["n_days"] / 365.25, 1e-3)
        _exp_total = _exp_ann * _oos_years
        _act_total = oos_m["total"]
        _oos_efficiency = _act_total / _exp_total if abs(_exp_total) > 1 else None

        if _oos_efficiency is None:
            _pnl_rag = "GREEN"
            _pnl_detail = "Cannot compute efficiency"
        elif _oos_efficiency < 0.0:
            _pnl_rag = "RED"
            _pnl_detail = f"OOS P&L is negative (efficiency: {_oos_efficiency:.0%} of expected)"
        elif _oos_efficiency < 0.5:
            _pnl_rag = "RED"
            _pnl_detail = f"OOS efficiency {_oos_efficiency:.0%} of expected (< 50%)"
        elif _oos_efficiency < 0.75:
            _pnl_rag = "AMBER"
            _pnl_detail = f"OOS efficiency {_oos_efficiency:.0%} of expected (< 75%)"
        else:
            _pnl_rag = "GREEN"
            _pnl_detail = f"OOS efficiency {_oos_efficiency:.0%} of expected"

    # ── Metric 3: Consecutive losing months in OOS ─────────────────────────────
    if oos_pnl.empty:
        _consec_rag = "GREEN"
        _consec_detail = "No OOS data"
        _consec_losing = 0
    else:
        _oos_monthly = oos_pnl.resample("ME").sum()
        _oos_monthly_arr = _oos_monthly.values[::-1]  # most recent first
        _consec_losing = 0
        for v in _oos_monthly_arr:
            if v < 0:
                _consec_losing += 1
            else:
                break

        if _consec_losing >= 4:
            _consec_rag = "RED"
            _consec_detail = f"{_consec_losing} consecutive losing months in OOS (≥ 4)"
        elif _consec_losing >= 2:
            _consec_rag = "AMBER"
            _consec_detail = f"{_consec_losing} consecutive losing months in OOS (≥ 2)"
        else:
            _consec_rag = "GREEN"
            _consec_detail = f"{_consec_losing} consecutive losing months in OOS"

    # ── Metric 4: Current drawdown vs historical max ───────────────────────────
    _full_eq      = scaled_pnl.cumsum()
    _peak_all     = _full_eq.cummax()
    _current_eq   = float(_full_eq.iloc[-1]) if not _full_eq.empty else 0.0
    _peak_val     = float(_peak_all.iloc[-1]) if not _peak_all.empty else 0.0
    _current_dd   = max(_peak_val - _current_eq, 0.0)
    _hist_max_dd  = float((_peak_all - _full_eq).max()) if not _full_eq.empty else 0.0

    if _hist_max_dd < 1:
        _dd_rag = "GREEN"
        _dd_detail = "No drawdown on record"
        _dd_pct = 0.0
    else:
        _dd_pct = _current_dd / _hist_max_dd
        if _dd_pct > 0.9:
            _dd_rag = "RED"
            _dd_detail = f"Current DD ${_current_dd:,.0f} is {_dd_pct:.0%} of historical max ${_hist_max_dd:,.0f}"
        elif _dd_pct > 0.6:
            _dd_rag = "AMBER"
            _dd_detail = f"Current DD ${_current_dd:,.0f} is {_dd_pct:.0%} of historical max ${_hist_max_dd:,.0f}"
        else:
            _dd_rag = "GREEN"
            _dd_detail = f"Current DD ${_current_dd:,.0f} is {_dd_pct:.0%} of historical max ${_hist_max_dd:,.0f}"

    # ── Overall RAG ────────────────────────────────────────────────────────────
    _all_rags = [_opt_rag, _pnl_rag, _consec_rag, _dd_rag]
    if "RED" in _all_rags:
        _overall_rag = "RED"
    elif "AMBER" in _all_rags:
        _overall_rag = "AMBER"
    else:
        _overall_rag = "GREEN"

    # ── Render ─────────────────────────────────────────────────────────────────
    ov_col, _ = st.columns([1, 3])
    with ov_col:
        bg = {"GREEN": "#E8F5E9", "AMBER": "#FFF8E1", "RED": "#FFEBEE"}.get(_overall_rag, "#F5F5F5")
        st.markdown(
            f"<div style='background:{bg};padding:14px 20px;border-radius:10px;"
            f"font-size:1.1rem;font-weight:600;'>"
            f"Overall Status &nbsp; {_rag(_overall_rag)}</div>",
            unsafe_allow_html=True,
        )

    st.divider()

    rag_data = [
        ("Days Since Last Optimisation", _opt_rag,    _opt_detail),
        ("OOS P&L vs Expected",          _pnl_rag,    _pnl_detail),
        ("Consecutive Losing Months",    _consec_rag, _consec_detail),
        ("Current Drawdown vs Max",      _dd_rag,     _dd_detail),
    ]

    for label, rag, detail in rag_data:
        rc1, rc2, rc3 = st.columns([2, 1, 5])
        rc1.markdown(f"**{label}**")
        rc2.markdown(_rag(rag), unsafe_allow_html=True)
        rc3.markdown(f"<span style='color:#555;'>{detail}</span>", unsafe_allow_html=True)
        st.divider()

    # ── OOS rolling monthly P&L trend ─────────────────────────────────────────
    if not oos_pnl.empty:
        st.markdown("#### OOS Monthly P&L — Rolling Trend")
        _oos_mo = oos_pnl.resample("ME").sum()
        fig_mo = go.Figure()
        fig_mo.add_trace(go.Bar(
            x=_oos_mo.index, y=_oos_mo.values,
            marker_color=[("#2E7D32" if v >= 0 else "#C62828") for v in _oos_mo.values],
            name="Monthly P&L",
        ))
        # 3-month rolling average
        if len(_oos_mo) >= 3:
            _roll3 = _oos_mo.rolling(3).mean()
            fig_mo.add_trace(go.Scatter(
                x=_roll3.index, y=_roll3.values,
                name="3-month MA", line=dict(color="#FF6F00", width=2),
            ))
        fig_mo.add_hline(y=0, line_color="black", line_width=1)
        fig_mo.update_layout(
            height=300, xaxis_title="Month", yaxis_title="Monthly P&L ($)",
            hovermode="x unified",
            legend=dict(orientation="h", yanchor="bottom", y=1.02),
        )
        st.plotly_chart(fig_mo, use_container_width=True)

    # ── Key numbers panel ──────────────────────────────────────────────────────
    st.markdown("#### Key Numbers")
    kn1, kn2, kn3, kn4 = st.columns(4)
    kn1.metric("Days Since Last Opt", f"{_days_since}" if _days_since is not None else "—")
    kn2.metric("OOS Efficiency",
               f"{_oos_efficiency:.0%}" if _oos_efficiency is not None else "—",
               help="Actual OOS P&L ÷ Expected OOS P&L (based on IS annual rate × OOS years)")
    kn3.metric("Consec. Losing Months", str(_consec_losing))
    kn4.metric("Current DD vs Max",
               f"{_dd_pct:.0%}" if _hist_max_dd > 1 else "—",
               help="Current drawdown as % of all-time maximum drawdown")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — DRAWDOWN DEEP-DIVE
# ══════════════════════════════════════════════════════════════════════════════
with tab_dd:
    st.subheader("Drawdown Deep-Dive")
    st.caption("Full underwater equity curve, drawdown events, duration histogram, and rolling max drawdown.")

    _full_eq2 = scaled_pnl.cumsum()
    _peak2    = _full_eq2.cummax()
    _dd_ser   = _peak2 - _full_eq2           # positive = underwater

    # ── Chart 1: Full underwater equity curve ─────────────────────────────────
    st.markdown("#### Underwater Equity Curve")
    fig_uw = go.Figure()

    if oos_ts is not None:
        fig_uw.add_trace(go.Scatter(
            x=_dd_ser[_dd_ser.index < oos_ts].index,
            y=-_dd_ser[_dd_ser.index < oos_ts].values,
            fill="tozeroy", name="IS Drawdown",
            line=dict(color="#90CAF9"), fillcolor="rgba(144,202,249,0.4)",
        ))
        fig_uw.add_trace(go.Scatter(
            x=_dd_ser[_dd_ser.index >= oos_ts].index,
            y=-_dd_ser[_dd_ser.index >= oos_ts].values,
            fill="tozeroy", name="OOS Drawdown",
            line=dict(color="#F44336"), fillcolor="rgba(244,67,54,0.4)",
        ))
        fig_uw.add_vline(
            x=int(oos_ts.timestamp() * 1000), line_dash="dash", line_color="#B71C1C",
            annotation_text="OOS Start", annotation_position="top right",
        )
    else:
        fig_uw.add_trace(go.Scatter(
            x=_dd_ser.index, y=-_dd_ser.values,
            fill="tozeroy", name="Drawdown",
            line=dict(color="#F44336"), fillcolor="rgba(244,67,54,0.4)",
        ))

    fig_uw.update_layout(
        height=300, xaxis_title="Date", yaxis_title="Drawdown ($)",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    st.plotly_chart(fig_uw, use_container_width=True)

    # ── Detect discrete drawdown events ───────────────────────────────────────
    def _dd_events(eq: pd.Series) -> list[dict]:
        pk = eq.cummax()
        dds = pk - eq
        events, in_dd, start, peak_at_start, max_dd_val = [], False, None, 0.0, 0.0
        for ts, val in dds.items():
            if val > 0 and not in_dd:
                in_dd, start, max_dd_val = True, ts, val
            elif in_dd and val > max_dd_val:
                max_dd_val = val
            elif in_dd and val == 0.0:
                in_dd = False
                events.append({
                    "start":    start,
                    "recovery": ts,
                    "duration": (ts - start).days,
                    "max_dd":   max_dd_val,
                    "open":     False,
                })
                max_dd_val = 0.0
        if in_dd:
            events.append({
                "start":    start,
                "recovery": None,
                "duration": (dds.index[-1] - start).days,
                "max_dd":   max_dd_val,
                "open":     True,
            })
        return events

    events = _dd_events(_full_eq2)

    # ── Chart 2: Drawdown duration histogram ──────────────────────────────────
    st.markdown("#### Drawdown Duration Histogram")
    st.caption("How long drawdowns typically last (calendar days).")

    if events:
        durations = [e["duration"] for e in events]
        fig_hist = px.histogram(
            x=durations, nbins=20,
            color_discrete_sequence=["#EF5350"],
            labels={"x": "Duration (days)", "y": "Count"},
        )
        fig_hist.update_layout(height=300, bargap=0.05)
        st.plotly_chart(fig_hist, use_container_width=True)
    else:
        st.info("No discrete drawdown events detected.")

    # ── Chart 3: Time-to-recovery per event ───────────────────────────────────
    st.markdown("#### Time to Recovery per Drawdown Event")
    st.caption("Each bubble shows a drawdown event: X = start date, Y = duration (days), size = max DD depth.")

    closed_events = [e for e in events if not e["open"]]
    if closed_events:
        ev_df = pd.DataFrame({
            "start":    [e["start"] for e in closed_events],
            "duration": [e["duration"] for e in closed_events],
            "max_dd":   [e["max_dd"] for e in closed_events],
            "label":    [f"DD: ${e['max_dd']:,.0f}<br>Duration: {e['duration']}d" for e in closed_events],
        })
        fig_recov = go.Figure()
        fig_recov.add_trace(go.Scatter(
            x=ev_df["start"], y=ev_df["duration"],
            mode="markers",
            marker=dict(
                size=[max(8, min(40, d / 5)) for d in ev_df["max_dd"]],
                color=ev_df["max_dd"], colorscale="Reds",
                showscale=True,
                colorbar=dict(title="Max DD ($)"),
            ),
            text=ev_df["label"],
            hoverinfo="text",
        ))
        fig_recov.update_layout(
            height=350, xaxis_title="Drawdown Start Date",
            yaxis_title="Days to Recovery",
        )
        st.plotly_chart(fig_recov, use_container_width=True)
    elif events:
        st.info("Current drawdown still open — no completed recovery events yet.")
    else:
        st.info("No drawdown events detected.")

    # ── Chart 4: Rolling max drawdown ─────────────────────────────────────────
    st.markdown("#### Rolling Maximum Drawdown (1-Year Window)")
    st.caption("Worst drawdown within each trailing 252-day window — shows whether current conditions are historically normal.")

    _window = 252
    _eq_arr = _full_eq2.values
    _roll_mdd = np.empty(len(_eq_arr))
    for i in range(len(_eq_arr)):
        _start = max(0, i - _window + 1)
        _sub   = _eq_arr[_start: i + 1]
        _pk    = np.maximum.accumulate(_sub)
        _roll_mdd[i] = float((_pk - _sub).max())

    _roll_mdd_series = pd.Series(_roll_mdd, index=_full_eq2.index)

    fig_roll = go.Figure()
    if oos_ts is not None:
        fig_roll.add_trace(go.Scatter(
            x=_roll_mdd_series[_roll_mdd_series.index < oos_ts].index,
            y=_roll_mdd_series[_roll_mdd_series.index < oos_ts].values,
            name="IS", line=dict(color="#90CAF9", width=2),
        ))
        fig_roll.add_trace(go.Scatter(
            x=_roll_mdd_series[_roll_mdd_series.index >= oos_ts].index,
            y=_roll_mdd_series[_roll_mdd_series.index >= oos_ts].values,
            name="OOS", line=dict(color="#F44336", width=2),
        ))
        fig_roll.add_vline(
            x=int(oos_ts.timestamp() * 1000), line_dash="dash", line_color="#B71C1C",
        )
    else:
        fig_roll.add_trace(go.Scatter(
            x=_roll_mdd_series.index, y=_roll_mdd_series.values,
            name="Rolling Max DD", line=dict(color="#F44336", width=2),
        ))

    fig_roll.update_layout(
        height=300, xaxis_title="Date", yaxis_title="Rolling Max Drawdown ($)",
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
    )
    st.plotly_chart(fig_roll, use_container_width=True)

    # ── Event table ────────────────────────────────────────────────────────────
    if events:
        st.markdown("#### Drawdown Event Table")
        ev_tbl = pd.DataFrame([{
            "Start":          str(e["start"].date()) if hasattr(e["start"], "date") else str(e["start"]),
            "Recovery":       str(e["recovery"].date()) if e["recovery"] and hasattr(e["recovery"], "date") else ("Open" if e["open"] else str(e["recovery"])),
            "Duration (days)": e["duration"],
            "Max Drawdown ($)": f"${e['max_dd']:,.0f}",
            "Status":          "Open" if e["open"] else "Recovered",
        } for e in sorted(events, key=lambda x: x["max_dd"], reverse=True)])
        st.dataframe(ev_tbl, hide_index=True, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — FILES & TRADES
# ══════════════════════════════════════════════════════════════════════════════
with tab_files:

    # Files & Folder ───────────────────────────────────────────────────────────
    def _open_path(path: Path) -> None:
        try:
            if platform.system() == "Windows":
                os.startfile(str(path))
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            st.error(f"Could not open: {exc}")

    _CODE_EXTENSIONS = {".mex", ".eld", ".els", ".pla", ".c", ".cpp", ".py"}

    def _is_code_file(f: Path) -> bool:
        if f.suffix.lower() in _CODE_EXTENSIONS:
            return True
        return f.suffix.lower() == ".txt" and f.name.endswith("ELCode.txt")

    _code_files: list[Path] = []
    _data_files: dict[str, Path] = {}
    if _folder_path and _folder_path.exists():
        for _f in sorted(_folder_path.iterdir()):
            if _is_code_file(_f):
                _code_files.append(_f)
            elif _f.suffix.lower() == ".csv":
                _data_files[_f.name] = _f

    st.subheader("Files & Folder")
    if _folder_path is None:
        st.info("Folder path not available. Import data first.")
    else:
        st.caption("Strategy folder")
        st.code(str(_folder_path), language=None)
        if _folder_path.exists():
            if st.button("📂 Open Folder", key="open_folder_btn"):
                _open_path(_folder_path)
        else:
            st.warning(f"Folder not found on disk: `{_folder_path}`")

        if _code_files:
            st.caption("Code files")
            for _cf in _code_files:
                col_name, col_btn = st.columns([5, 1])
                col_name.markdown(f"`{_cf.name}`")
                if col_btn.button("Open", key=f"open_code_{_cf.name}"):
                    _open_path(_cf)
        else:
            st.caption("No code files (.mex / .eld / .pla / .els / ELCode.txt) found.")

        if _data_files:
            st.caption("Data files (CSV)")
            for _name, _path in _data_files.items():
                try:
                    size_kb = _path.stat().st_size / 1024
                    st.markdown(f"- `{_name}` — {size_kb:.0f} KB")
                except Exception:
                    st.markdown(f"- `{_name}`")

    st.divider()

    # Trade list ───────────────────────────────────────────────────────────────
    if not imported.trades.empty and "strategy" in imported.trades.columns:
        strat_trades = imported.trades[imported.trades["strategy"] == selected_name].copy()
        if not strat_trades.empty:
            st.subheader(f"Trade List ({len(strat_trades)} trades)")
            strat_trades = strat_trades.drop(columns=["strategy"], errors="ignore")
            if "date" in strat_trades.columns:
                strat_trades = strat_trades.sort_values("date", ascending=False)
            if "pnl" in strat_trades.columns:
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
        else:
            st.info("No trades found for this strategy.")
    else:
        st.info("No trade data loaded.")
