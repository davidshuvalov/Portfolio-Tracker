"""
Market Analysis — Buy & Hold market intelligence terminal.

Uses Buy & Hold strategy equity data as market proxies to build:
1. Market Overview    — all traded markets, current ATR & volatility regime
2. ATR & Vol Regimes  — 50-day vs 200-day rolling range with regime shading
3. Correlations       — market & sector correlation heatmap with date slider
4. Market News        — live RSS headlines for traded instruments
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import requests
import streamlit as st

from core.config import AppConfig
from core.portfolio.strategies import load_strategies

st.set_page_config(page_title="Market Analysis", layout="wide")

# ── Session state / data loading ─────────────────────────────────────────────

config: AppConfig = st.session_state.get("config", AppConfig.load())
imported = st.session_state.get("imported_data")
strategies_config = load_strategies()

st.title("Market Analysis")
st.caption("Buy & Hold market intelligence — ATR, volatility regimes, correlations, and market news.")

_nav_l, _ = st.columns([1, 7])
with _nav_l:
    st.page_link("ui/pages/03_Portfolio.py", label="← Portfolio")

if imported is None:
    st.info("No data loaded yet — import your strategy data first.")
    st.page_link("ui/pages/01_Import.py", label="Go to Import →")
    st.stop()

# ── Identify Buy & Hold strategies ────────────────────────────────────────────

_bh_cfg = [s for s in strategies_config if "buy" in s.get("status", "").lower() and "hold" in s.get("status", "").lower()]
_bh_names_cfg = {s["name"] for s in _bh_cfg}

# Also include any strategy whose name appears in imported data and is tagged B&H
_all_imported = set(imported.strategy_names)
_bh_names = sorted(_bh_names_cfg & _all_imported)

# Fallback: if none configured, try to detect by presence in data + no matching Live/Pass etc.
if not _bh_names:
    _non_bh_statuses = {"Live", "Paper", "Retired", "Pass", "Incubating", "New"}
    _cfg_names = {s["name"]: s.get("status", "") for s in strategies_config}
    _bh_names = sorted([
        n for n in _all_imported
        if _cfg_names.get(n, "Buy&Hold") not in _non_bh_statuses
        and "buy" in _cfg_names.get(n, "Buy&Hold").lower()
    ])

if not _bh_names:
    st.warning(
        "No **Buy & Hold** strategies found. "
        "Mark strategies as **Buy&Hold** status on the Strategies page to enable Market Analysis."
    )
    st.page_link("ui/pages/02_Strategies.py", label="Go to Strategies →")
    st.stop()

# ── Build market metadata lookup ──────────────────────────────────────────────

_cfg_map = {s["name"]: s for s in strategies_config}

def _market_symbol(name: str) -> str:
    return _cfg_map.get(name, {}).get("symbol", name)

def _market_sector(name: str) -> str:
    return _cfg_map.get(name, {}).get("sector", "Other")

# ── Core data: daily M2M for B&H strategies ───────────────────────────────────

bh_cols = [n for n in _bh_names if n in imported.daily_m2m.columns]
if not bh_cols:
    st.error("Buy & Hold strategy data not found in imported equity data.")
    st.stop()

bh_m2m: pd.DataFrame = imported.daily_m2m[bh_cols].copy()
bh_m2m = bh_m2m.loc[(bh_m2m != 0).any(axis=1)]  # drop all-zero rows

# Dollar range proxy = abs(daily M2M) — captures market movement in dollars
bh_range = bh_m2m.abs()

# ── ATR computation (rolling mean of dollar range) ────────────────────────────

def _compute_atr(df: pd.DataFrame, window: int) -> pd.DataFrame:
    """Rolling mean of abs(daily_m2m) = dollar ATR proxy."""
    return df.rolling(window=window, min_periods=max(1, window // 4)).mean()


atr_50d  = _compute_atr(bh_range, 50)
atr_200d = _compute_atr(bh_range, 200)
atr_63d  = _compute_atr(bh_range, 63)   # ~3 month
atr_126d = _compute_atr(bh_range, 126)  # ~6 month
atr_252d = _compute_atr(bh_range, 252)  # ~12 month

# ── Volatility regime ─────────────────────────────────────────────────────────

def _regime(ratio: float) -> str:
    if pd.isna(ratio):
        return "Unknown"
    if ratio > 1.15:
        return "High"
    if ratio < 0.85:
        return "Low"
    return "Normal"

# Current regime per market
_regime_ratio = (atr_50d.iloc[-1] / atr_200d.iloc[-1]).fillna(1.0)
_regime_labels = _regime_ratio.map(_regime)

_REGIME_COLOR = {"High": "#ef4444", "Normal": "#f59e0b", "Low": "#10b981", "Unknown": "#6b7280"}
_REGIME_BG    = {"High": "#2d1212", "Normal": "#2d2208", "Low": "#0d2d1a", "Unknown": "#1c1c2e"}

# ── Yahoo Finance symbol map (for news) ───────────────────────────────────────

_YAHOO_MAP: dict[str, str] = {
    "ES": "ES=F", "NQ": "NQ=F", "YM": "YM=F", "RTY": "RTY=F",
    "CL": "CL=F", "NG": "NG=F", "RB": "RB=F", "HO": "HO=F",
    "GC": "GC=F", "SI": "SI=F", "HG": "HG=F", "PL": "PL=F",
    "ZB": "ZB=F", "ZN": "ZN=F", "ZF": "ZF=F", "ZT": "ZT=F",
    "6E": "EURUSD=X", "EC": "EURUSD=X",
    "6J": "JPY=X",   "JY": "JPY=X",
    "6B": "GBPUSD=X", "BP": "GBPUSD=X",
    "6A": "AUDUSD=X", "AD": "AUDUSD=X",
    "6C": "CADUSD=X", "CD": "CADUSD=X",
    "6S": "CHFUSD=X",
    "ZC": "ZC=F", "ZW": "ZW=F", "ZS": "ZS=F", "ZM": "ZM=F", "ZL": "ZL=F",
    "LE": "LE=F", "GF": "GF=F", "HE": "HE=F",
    "BTC": "BTC-USD", "ETH": "ETH-USD",
    "FDAX": "^GDAXI", "FGBL": "GBL=F",
    "VX": "^VIX",
}

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────

tab_overview, tab_atr, tab_corr, tab_news = st.tabs([
    "📊 Market Overview",
    "📈 ATR & Volatility Regimes",
    "🔗 Correlations",
    "📰 Market News",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1: MARKET OVERVIEW
# ══════════════════════════════════════════════════════════════════════════════

with tab_overview:
    st.subheader("Market Overview")
    st.caption(
        f"Showing **{len(bh_cols)}** Buy & Hold markets. "
        "ATR = rolling mean of |daily M2M| (dollar range proxy). "
        "Regime = 50-day ATR vs 200-day ATR ratio."
    )

    # ── Overview metrics row ───────────────────────────────────────────────
    _high_vol_n  = (_regime_labels == "High").sum()
    _low_vol_n   = (_regime_labels == "Low").sum()
    _norm_vol_n  = (_regime_labels == "Normal").sum()

    _ov_c1, _ov_c2, _ov_c3, _ov_c4 = st.columns(4)
    _ov_c1.metric("Total Markets", len(bh_cols))
    _ov_c2.metric("High Volatility", _high_vol_n, delta=None)
    _ov_c3.metric("Normal Volatility", _norm_vol_n, delta=None)
    _ov_c4.metric("Low Volatility", _low_vol_n, delta=None)

    st.divider()

    # ── Sector grouping ────────────────────────────────────────────────────
    sectors = sorted(set(_market_sector(n) for n in bh_cols))
    _all_sectors_opt = ["All Sectors"] + sectors
    _sel_sector = st.selectbox("Filter by sector", _all_sectors_opt, key="ov_sector")

    _ov_names = bh_cols if _sel_sector == "All Sectors" else [
        n for n in bh_cols if _market_sector(n) == _sel_sector
    ]

    # ── Build overview table ───────────────────────────────────────────────
    _ov_rows = []
    for name in _ov_names:
        sym    = _market_symbol(name)
        sector = _market_sector(name)
        regime = _regime_labels.get(name, "Unknown")
        ratio  = float(_regime_ratio.get(name, 1.0))

        c_atr_3m  = float(atr_63d[name].iloc[-1])  if name in atr_63d.columns  else 0.0
        c_atr_6m  = float(atr_126d[name].iloc[-1]) if name in atr_126d.columns else 0.0
        c_atr_12m = float(atr_252d[name].iloc[-1]) if name in atr_252d.columns else 0.0

        # Recent trend: last 20 days of ATR vs prior 20 days
        _recent = float(atr_63d[name].iloc[-20:].mean())  if len(atr_63d[name]) >= 40 else c_atr_3m
        _prior  = float(atr_63d[name].iloc[-40:-20].mean()) if len(atr_63d[name]) >= 40 else c_atr_3m
        trend = "↑" if _recent > _prior * 1.05 else ("↓" if _recent < _prior * 0.95 else "→")

        last_date = bh_m2m[name].last_valid_index()
        last_pnl  = float(bh_m2m[name].iloc[-1]) if len(bh_m2m[name]) > 0 else 0.0

        _ov_rows.append({
            "Market": name,
            "Symbol": sym,
            "Sector": sector,
            "Regime": regime,
            "50d/200d": round(ratio, 2),
            "ATR 3M ($)": round(c_atr_3m, 0),
            "ATR 6M ($)": round(c_atr_6m, 0),
            "ATR 12M ($)": round(c_atr_12m, 0),
            "Vol Trend": trend,
            "Last Date": last_date.date() if last_date else None,
            "Last PnL ($)": round(last_pnl, 0),
        })

    _ov_df = pd.DataFrame(_ov_rows)

    def _color_regime(row):
        c = _REGIME_COLOR.get(row["Regime"], "#6b7280")
        bg = _REGIME_BG.get(row["Regime"], "#1c1c2e")
        style = [f"color:{c};background-color:{bg};font-weight:600"
                 if col == "Regime" else "" for col in _ov_df.columns]
        return style

    st.dataframe(
        _ov_df.style.apply(_color_regime, axis=1),
        use_container_width=True,
        hide_index=True,
        column_config={
            "ATR 3M ($)":  st.column_config.NumberColumn(format="$%.0f"),
            "ATR 6M ($)":  st.column_config.NumberColumn(format="$%.0f"),
            "ATR 12M ($)": st.column_config.NumberColumn(format="$%.0f"),
            "Last PnL ($)": st.column_config.NumberColumn(format="$%.0f"),
            "Last Date":   st.column_config.DateColumn("Last Date"),
            "50d/200d":    st.column_config.NumberColumn("50d/200d ATR", format="%.2f"),
        },
    )

    # ── ATR sparklines grid ────────────────────────────────────────────────
    st.divider()
    st.subheader("ATR Sparklines — Last 12 Months")
    _spark_lookback = bh_m2m.index[-252:] if len(bh_m2m) >= 252 else bh_m2m.index

    _ncols = 4
    _spark_names = [n for n in _ov_names if n in atr_63d.columns]
    for _row_i in range(0, len(_spark_names), _ncols):
        _spark_cols = st.columns(_ncols)
        for _ci, _name in enumerate(_spark_names[_row_i:_row_i + _ncols]):
            with _spark_cols[_ci]:
                _sa = atr_63d[_name].reindex(_spark_lookback).dropna()
                if len(_sa) < 2:
                    st.caption(f"**{_name}** — insufficient data")
                    continue
                regime_now = _regime_labels.get(_name, "Normal")
                rc = _REGIME_COLOR.get(regime_now, "#f59e0b")
                _sfig = go.Figure()
                _sfig.add_trace(go.Scatter(
                    x=_sa.index, y=_sa.values,
                    mode="lines", line=dict(color=rc, width=1.5),
                    fill="tozeroy", fillcolor=rc.replace(")", ",0.12)").replace("rgb", "rgba"),
                    hovertemplate="%{x|%d %b}: $%{y:,.0f}<extra></extra>",
                    showlegend=False,
                ))
                _sfig.update_layout(
                    margin=dict(l=0, r=0, t=24, b=0),
                    height=100,
                    xaxis=dict(visible=False),
                    yaxis=dict(visible=False),
                    plot_bgcolor="rgba(0,0,0,0)",
                    paper_bgcolor="rgba(0,0,0,0)",
                    title=dict(
                        text=f"<b>{_market_symbol(_name)}</b> <span style='color:{rc}'>{regime_now}</span>"
                             f"  <span style='color:#94a3b8;font-size:0.7rem'>${float(_sa.iloc[-1]):,.0f}</span>",
                        font=dict(size=11), x=0, xanchor="left", y=1, yanchor="top",
                    ),
                )
                st.plotly_chart(_sfig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2: ATR & VOLATILITY REGIMES
# ══════════════════════════════════════════════════════════════════════════════

with tab_atr:
    st.subheader("ATR & Volatility Regime Analysis")

    # ── Controls ──────────────────────────────────────────────────────────
    _av_c1, _av_c2, _av_c3 = st.columns([3, 2, 2])
    with _av_c1:
        _sel_markets = st.multiselect(
            "Markets",
            options=bh_cols,
            default=bh_cols[:min(4, len(bh_cols))],
            format_func=lambda n: f"{_market_symbol(n)} ({n})",
            key="av_markets",
        )
    with _av_c2:
        _date_presets = {"1 Year": 252, "3 Years": 756, "5 Years": 1260, "All": -1}
        _preset = st.selectbox("Date range", list(_date_presets.keys()), index=2, key="av_preset")
    with _av_c3:
        _show_raw = st.checkbox("Overlay raw daily range", value=False, key="av_raw")

    if not _sel_markets:
        st.info("Select at least one market above.")
    else:
        _lookback_days = _date_presets[_preset]
        for _mkt in _sel_markets:
            sym = _market_symbol(_mkt)
            regime_now = _regime_labels.get(_mkt, "Normal")
            rc = _REGIME_COLOR.get(regime_now, "#f59e0b")

            _raw = bh_range[_mkt] if _mkt in bh_range.columns else pd.Series(dtype=float)
            _a50  = atr_50d[_mkt]  if _mkt in atr_50d.columns  else pd.Series(dtype=float)
            _a200 = atr_200d[_mkt] if _mkt in atr_200d.columns else pd.Series(dtype=float)

            if _lookback_days > 0:
                _raw  = _raw.iloc[-_lookback_days:]
                _a50  = _a50.iloc[-_lookback_days:]
                _a200 = _a200.iloc[-_lookback_days:]

            _a50  = _a50.dropna()
            _a200 = _a200.dropna()

            _fig = go.Figure()

            # Regime background shading (fill between 50d and 200d)
            if len(_a50) > 0 and len(_a200) > 0:
                _common_idx = _a50.index.intersection(_a200.index)
                _hi = _a50.reindex(_common_idx)
                _lo = _a200.reindex(_common_idx)

                # High vol zone (50d > 200d)
                _fig.add_trace(go.Scatter(
                    x=list(_common_idx) + list(_common_idx[::-1]),
                    y=list(_hi.clip(lower=_lo).values) + list(_lo.values[::-1]),
                    fill="toself",
                    fillcolor="rgba(239,68,68,0.10)",
                    line=dict(width=0),
                    showlegend=False,
                    hoverinfo="skip",
                ))
                # Low vol zone (50d < 200d)
                _fig.add_trace(go.Scatter(
                    x=list(_common_idx) + list(_common_idx[::-1]),
                    y=list(_lo.clip(lower=_hi).values) + list(_hi.values[::-1]),
                    fill="toself",
                    fillcolor="rgba(16,185,129,0.10)",
                    line=dict(width=0),
                    showlegend=False,
                    hoverinfo="skip",
                ))

            if _show_raw and len(_raw) > 0:
                _fig.add_trace(go.Bar(
                    x=_raw.index, y=_raw.values,
                    name="Daily Range",
                    marker_color="rgba(100,116,139,0.3)",
                    showlegend=True,
                ))

            if len(_a50) > 0:
                _fig.add_trace(go.Scatter(
                    x=_a50.index, y=_a50.values,
                    mode="lines", name="50-day MA",
                    line=dict(color="#3b82f6", width=1.8),
                ))
            if len(_a200) > 0:
                _fig.add_trace(go.Scatter(
                    x=_a200.index, y=_a200.values,
                    mode="lines", name="200-day MA",
                    line=dict(color="#f59e0b", width=2.2, dash="dash"),
                ))

            _fig.update_layout(
                title=dict(
                    text=f"<b>{sym}</b> — Regime: "
                         f"<span style='color:{rc}'>{regime_now}</span>  "
                         f"(50d/200d = {float(_regime_ratio.get(_mkt, 1.0)):.2f}×)",
                    font=dict(size=13), x=0,
                ),
                xaxis_title="Date",
                yaxis_title="Dollar ATR ($)",
                hovermode="x unified",
                height=320,
                margin=dict(l=0, r=0, t=40, b=0),
                legend=dict(orientation="h", yanchor="bottom", y=1.02),
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(_fig, use_container_width=True)

    # ── Regime history chart ───────────────────────────────────────────────
    st.divider()
    st.subheader("Volatility Regime Ratios — All Markets")

    _all_ratios = (atr_50d / atr_200d.replace(0, np.nan))
    if _lookback_days > 0:
        _all_ratios = _all_ratios.iloc[-_lookback_days:]
    _all_ratios = _all_ratios[bh_cols].dropna(how="all")

    _rr_fig = go.Figure()
    _rr_fig.add_hline(y=1.15, line=dict(color="rgba(239,68,68,0.4)", dash="dot"), annotation_text="High vol threshold")
    _rr_fig.add_hline(y=0.85, line=dict(color="rgba(16,185,129,0.4)", dash="dot"), annotation_text="Low vol threshold")
    _rr_fig.add_hline(y=1.0,  line=dict(color="rgba(100,116,139,0.3)", width=1))

    _palette = px.colors.qualitative.Plotly
    for _ki, _mkt in enumerate(bh_cols):
        if _mkt not in _all_ratios.columns:
            continue
        _s = _all_ratios[_mkt].dropna()
        _rr_fig.add_trace(go.Scatter(
            x=_s.index, y=_s.values,
            mode="lines", name=_market_symbol(_mkt),
            line=dict(width=1.5, color=_palette[_ki % len(_palette)]),
            opacity=0.85,
        ))

    _rr_fig.update_layout(
        xaxis_title="Date",
        yaxis_title="50d ATR / 200d ATR ratio",
        hovermode="x unified",
        height=360,
        margin=dict(l=0, r=0, t=10, b=0),
        legend=dict(orientation="h", yanchor="bottom", y=1.02),
        plot_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(_rr_fig, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3: CORRELATIONS
# ══════════════════════════════════════════════════════════════════════════════

with tab_corr:
    st.subheader("Market & Sector Correlations")

    # ── Date range slider ──────────────────────────────────────────────────
    _all_dates = bh_m2m.index
    _min_date = _all_dates[0].date() if len(_all_dates) > 0 else datetime.today().date() - timedelta(days=365*5)
    _max_date = _all_dates[-1].date() if len(_all_dates) > 0 else datetime.today().date()

    _cr_c1, _cr_c2 = st.columns([3, 1])
    with _cr_c1:
        _presets_corr = {
            "Last 1 Year":  (_max_date - timedelta(days=365),  _max_date),
            "Last 3 Years": (_max_date - timedelta(days=365*3), _max_date),
            "Last 5 Years": (_max_date - timedelta(days=365*5), _max_date),
            "All History":  (_min_date, _max_date),
        }
        _corr_preset = st.selectbox("Date range preset", list(_presets_corr.keys()), index=1, key="corr_preset")
        _preset_start, _preset_end = _presets_corr[_corr_preset]

        _corr_start, _corr_end = st.slider(
            "Adjust date range",
            min_value=_min_date,
            max_value=_max_date,
            value=(_preset_start, _preset_end),
            format="YYYY-MM-DD",
            key="corr_slider",
        )
    with _cr_c2:
        _corr_method = st.radio("Method", ["Pearson", "Spearman"], key="corr_method")
        _min_overlap = st.number_input("Min days overlap", min_value=20, max_value=252, value=60, step=10, key="corr_overlap")

    # Filter data to selected range
    _m2m_filtered = bh_m2m.loc[
        (bh_m2m.index.date >= _corr_start) & (bh_m2m.index.date <= _corr_end)
    ]

    if len(_m2m_filtered) < _min_overlap:
        st.warning(f"Only {len(_m2m_filtered)} days in selected range — need at least {_min_overlap}.")
    else:
        _method = _corr_method.lower()
        _corr_matrix = _m2m_filtered.corr(method=_method, min_periods=_min_overlap)

        # ── Market correlation heatmap ─────────────────────────────────────
        st.markdown(f"**Market Correlations** — {_corr_start} to {_corr_end} ({len(_m2m_filtered)} days)")

        _labels = [_market_symbol(n) for n in _corr_matrix.columns]
        _vals   = _corr_matrix.values

        _heatmap_text = [[f"{v:.2f}" if not np.isnan(v) else "" for v in row] for row in _vals]

        _fig_corr = go.Figure(go.Heatmap(
            z=_vals,
            x=_labels,
            y=_labels,
            colorscale="RdBu_r",
            zmid=0, zmin=-1, zmax=1,
            text=_heatmap_text,
            texttemplate="%{text}",
            textfont={"size": 9},
            hovertemplate="%{y} / %{x}: %{text}<extra></extra>",
            colorbar=dict(title="Corr"),
        ))
        _fig_corr.update_layout(
            height=max(400, len(_corr_matrix) * 28 + 80),
            margin=dict(l=0, r=0, t=10, b=0),
        )
        st.plotly_chart(_fig_corr, use_container_width=True)

        # ── Sector correlation ─────────────────────────────────────────────
        st.divider()
        st.markdown("**Sector Average Correlations**")

        _sector_map = {n: _market_sector(n) for n in bh_cols}
        _unique_sectors = sorted(set(_sector_map.values()))

        if len(_unique_sectors) > 1:
            # Build sector-level correlation (average of all market pairs within/between sectors)
            _sec_labels = _unique_sectors
            _sec_size   = len(_sec_labels)
            _sec_corr   = pd.DataFrame(np.nan, index=_sec_labels, columns=_sec_labels)

            for _s1 in _sec_labels:
                for _s2 in _sec_labels:
                    _mkts1 = [n for n in bh_cols if _sector_map[n] == _s1 and n in _corr_matrix.columns]
                    _mkts2 = [n for n in bh_cols if _sector_map[n] == _s2 and n in _corr_matrix.columns]
                    if not _mkts1 or not _mkts2:
                        continue
                    _sub = _corr_matrix.loc[_mkts1, _mkts2].values.flatten()
                    _sub = _sub[~np.isnan(_sub)]
                    if _s1 == _s2 and len(_sub) > len(_mkts1):
                        # Remove self-correlations (diag)
                        _diag = np.diag(_corr_matrix.loc[_mkts1, _mkts2].values)
                        _sub = _sub[_sub != 1.0]  # exclude perfect self-correlation
                    _sec_corr.loc[_s1, _s2] = float(np.nanmean(_sub)) if len(_sub) > 0 else np.nan

            _sec_vals = _sec_corr.values.astype(float)
            _sec_text = [[f"{v:.2f}" if not np.isnan(v) else "" for v in row] for row in _sec_vals]

            _fig_sec = go.Figure(go.Heatmap(
                z=_sec_vals,
                x=_sec_labels,
                y=_sec_labels,
                colorscale="RdBu_r",
                zmid=0, zmin=-1, zmax=1,
                text=_sec_text,
                texttemplate="%{text}",
                textfont={"size": 11},
                hovertemplate="%{y} / %{x}: %{text}<extra></extra>",
                colorbar=dict(title="Avg Corr"),
            ))
            _fig_sec.update_layout(
                height=max(300, _sec_size * 40 + 80),
                margin=dict(l=0, r=0, t=10, b=0),
            )
            st.plotly_chart(_fig_sec, use_container_width=True)

        # ── Highest / lowest correlated pairs ─────────────────────────────
        st.divider()
        st.markdown("**Most and Least Correlated Market Pairs**")

        _pairs = []
        _corr_cols = list(_corr_matrix.columns)
        for _i in range(len(_corr_cols)):
            for _j in range(_i + 1, len(_corr_cols)):
                _n1, _n2 = _corr_cols[_i], _corr_cols[_j]
                _v = _corr_matrix.loc[_n1, _n2]
                if not np.isnan(_v):
                    _pairs.append({
                        "Market A": _market_symbol(_n1),
                        "Market B": _market_symbol(_n2),
                        "Correlation": round(float(_v), 3),
                        "Sector A": _market_sector(_n1),
                        "Sector B": _market_sector(_n2),
                    })

        if _pairs:
            _pairs_df = pd.DataFrame(_pairs).sort_values("Correlation", ascending=False)
            _p1, _p2 = st.columns(2)
            with _p1:
                st.caption("Most correlated (top 10)")
                st.dataframe(
                    _pairs_df.head(10).reset_index(drop=True),
                    use_container_width=True, hide_index=True,
                    column_config={"Correlation": st.column_config.NumberColumn(format="%.3f")},
                )
            with _p2:
                st.caption("Least correlated (bottom 10)")
                st.dataframe(
                    _pairs_df.tail(10).reset_index(drop=True),
                    use_container_width=True, hide_index=True,
                    column_config={"Correlation": st.column_config.NumberColumn(format="%.3f")},
                )

        # ── Rolling correlation chart ──────────────────────────────────────
        st.divider()
        st.subheader("Rolling Correlation Between Two Markets")

        _rc_c1, _rc_c2, _rc_c3 = st.columns([2, 2, 1])
        with _rc_c1:
            _rc_mkt1 = st.selectbox(
                "Market A",
                bh_cols, format_func=lambda n: f"{_market_symbol(n)} ({n})",
                key="rc_mkt1",
            )
        with _rc_c2:
            _rc_default2 = bh_cols[1] if len(bh_cols) > 1 else bh_cols[0]
            _rc_mkt2 = st.selectbox(
                "Market B",
                bh_cols, index=min(1, len(bh_cols)-1),
                format_func=lambda n: f"{_market_symbol(n)} ({n})",
                key="rc_mkt2",
            )
        with _rc_c3:
            _rc_window = st.number_input("Window (days)", min_value=20, max_value=252, value=63, step=1, key="rc_window")

        if _rc_mkt1 != _rc_mkt2 and _rc_mkt1 in bh_m2m.columns and _rc_mkt2 in bh_m2m.columns:
            _roll_corr = (
                bh_m2m[[_rc_mkt1, _rc_mkt2]]
                .loc[(bh_m2m.index.date >= _corr_start) & (bh_m2m.index.date <= _corr_end)]
                .rolling(_rc_window, min_periods=_min_overlap)
                .corr()
                .unstack()[_rc_mkt2][_rc_mkt1]
                .dropna()
            )

            _fig_rc = go.Figure()
            _fig_rc.add_hline(y=0, line=dict(color="rgba(100,116,139,0.4)", width=1))
            _fig_rc.add_hline(y=0.7, line=dict(color="rgba(239,68,68,0.3)", dash="dot"))
            _fig_rc.add_hline(y=-0.7, line=dict(color="rgba(16,185,129,0.3)", dash="dot"))

            _fig_rc.add_trace(go.Scatter(
                x=_roll_corr.index, y=_roll_corr.values,
                mode="lines",
                name=f"{_rc_window}-day rolling corr",
                line=dict(color="#3b82f6", width=1.8),
                fill="tozeroy",
                fillcolor="rgba(59,130,246,0.08)",
            ))
            _fig_rc.update_layout(
                title=f"{_market_symbol(_rc_mkt1)} vs {_market_symbol(_rc_mkt2)} — {_rc_window}-day rolling correlation",
                xaxis_title="Date", yaxis_title="Correlation",
                yaxis=dict(range=[-1.05, 1.05]),
                height=300,
                margin=dict(l=0, r=0, t=40, b=0),
                plot_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(_fig_rc, use_container_width=True)


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4: MARKET NEWS
# ══════════════════════════════════════════════════════════════════════════════

with tab_news:
    st.subheader("Market News")
    st.caption("Live headlines for traded instruments via Yahoo Finance RSS. Headlines refresh on page load.")

    # ── Build unique symbols list ──────────────────────────────────────────
    _news_symbols: dict[str, str] = {}  # yahoo_sym → display_name
    for _n in bh_cols:
        _sym = _market_symbol(_n)
        _yahoo = _YAHOO_MAP.get(_sym, _YAHOO_MAP.get(_sym.upper(), ""))
        if _yahoo and _yahoo not in _news_symbols:
            _news_symbols[_yahoo] = f"{_sym} ({_market_sector(_n)})"

    # Always include broad market news
    _DEFAULT_FEEDS = {
        "https://finance.yahoo.com/news/rssindex": "Market Headlines (Yahoo Finance)",
    }

    # ── Symbol selector ────────────────────────────────────────────────────
    _news_c1, _news_c2 = st.columns([3, 1])
    with _news_c1:
        _all_yahoo_syms = list(_news_symbols.keys())
        _sel_news_syms = st.multiselect(
            "Fetch news for symbols",
            options=_all_yahoo_syms,
            default=_all_yahoo_syms[:min(4, len(_all_yahoo_syms))],
            format_func=lambda s: f"{s} — {_news_symbols.get(s, s)}",
            key="news_syms",
        )
    with _news_c2:
        _max_items = st.number_input("Max items per feed", min_value=3, max_value=20, value=6, step=1)
        _fetch_news = st.button("🔄 Refresh News", type="primary", key="refresh_news")

    # ── Fetch & parse RSS ─────────────────────────────────────────────────
    @st.cache_data(ttl=900, show_spinner=False)
    def _fetch_rss(url: str, max_items: int) -> list[dict]:
        """Fetch and parse an RSS feed. Returns list of {title, link, date, summary}."""
        try:
            headers = {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/120.0.0.0 Safari/537.36"
                )
            }
            resp = requests.get(url, headers=headers, timeout=8)
            if resp.status_code != 200:
                return []
            root = ET.fromstring(resp.content)
            ns = {"media": "http://search.yahoo.com/mrss/"}
            items = []
            for item in root.iter("item"):
                title   = (item.findtext("title") or "").strip()
                link    = (item.findtext("link")  or "").strip()
                pubdate = (item.findtext("pubDate") or "").strip()
                desc    = (item.findtext("description") or "").strip()
                # Strip HTML tags from description
                import re
                desc = re.sub(r"<[^>]+>", "", desc)[:200]
                if title:
                    items.append({
                        "title": title,
                        "link":  link,
                        "date":  pubdate,
                        "summary": desc,
                    })
                if len(items) >= max_items:
                    break
            return items
        except Exception:
            return []

    if _fetch_news:
        st.cache_data.clear()

    # ── General market news ────────────────────────────────────────────────
    _general_items = _fetch_rss(
        "https://finance.yahoo.com/news/rssindex",
        int(_max_items),
    )
    if _general_items:
        st.markdown("### Market Headlines")
        for _item in _general_items:
            with st.container(border=True):
                _dl, _dr = st.columns([5, 1])
                with _dl:
                    if _item["link"]:
                        st.markdown(f"**[{_item['title']}]({_item['link']})**")
                    else:
                        st.markdown(f"**{_item['title']}**")
                    if _item["summary"]:
                        st.caption(_item["summary"])
                with _dr:
                    st.caption(_item["date"][:16] if _item["date"] else "")

    # ── Per-symbol news ────────────────────────────────────────────────────
    for _ysym in _sel_news_syms:
        _display = _news_symbols.get(_ysym, _ysym)
        _url = f"https://finance.yahoo.com/rss/headline?s={_ysym}"
        _items = _fetch_rss(_url, int(_max_items))

        with st.expander(f"📰 {_display}", expanded=False):
            if not _items:
                st.caption("No headlines available or feed temporarily unavailable.")
            else:
                for _item in _items:
                    _tl, _tr = st.columns([5, 1])
                    with _tl:
                        if _item["link"]:
                            st.markdown(f"• **[{_item['title']}]({_item['link']})**")
                        else:
                            st.markdown(f"• {_item['title']}")
                        if _item["summary"]:
                            st.caption(_item["summary"])
                    with _tr:
                        st.caption(_item["date"][:16] if _item["date"] else "")

    if not _general_items and not _sel_news_syms:
        st.info("Select symbols above and click **Refresh News** to load headlines.")
