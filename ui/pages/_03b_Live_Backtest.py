"""
Live Backtest page — reconstruct live trading performance.

Lets you define which strategies you were trading (with contract counts and
date ranges) and builds the combined equity curve from the imported daily M2M
data. Two modes:
  - Manual: specify each strategy + date range individually
  - Portfolio Periods: stitch together saved portfolio snapshots
"""

from __future__ import annotations

import pandas as pd
import plotly.express as px
import streamlit as st

from core.config import AppConfig
from core.portfolio.strategies import load_strategies

st.set_page_config(page_title="Live Backtest", layout="wide")
st.title("Live Backtest")

# ── Top navigation ─────────────────────────────────────────────────────────────
_nav_l, _nav_r, _ = st.columns([1, 1, 6])
with _nav_l:
    st.page_link("ui/pages/03_Portfolio.py", label="← Build Portfolio")
with _nav_r:
    st.page_link("ui/pages/_05_Correlations.py", label="→ Correlations")

st.caption(
    "Reconstruct your live trading performance by defining which strategies you were "
    "trading and for which periods. The equity curve is built from the imported daily M2M data."
)

_bt_imported = st.session_state.get("imported_data")
if _bt_imported is None:
    st.info("Import data first to use the Live Backtest.")
    st.page_link("ui/pages/01_Import.py", label="Go to Import →")
    st.stop()

config: AppConfig = st.session_state.get("config", AppConfig.load())
strategies = load_strategies()

_bt_all_names = sorted({s.get("name", "") for s in strategies if s.get("name")})
_bt_m2m: pd.DataFrame = _bt_imported.daily_m2m

_bt_mode = st.radio(
    "Backtest mode",
    ["Manual Entries", "Portfolio Periods"],
    horizontal=True,
    help=(
        "**Manual**: Add individual strategies with custom contracts and date ranges.  \n"
        "**Portfolio Periods**: Use saved portfolio snapshots to stitch together your "
        "actual live-trading history across multiple portfolio versions."
    ),
)

st.divider()

# ── Manual mode ────────────────────────────────────────────────────────────────
if _bt_mode == "Manual Entries":
    st.markdown("### Add Strategy Entries")
    st.caption(
        "Each row represents a strategy you were trading with a specific contract count "
        "over a date range. Add multiple rows — they are summed to produce the combined curve."
    )

    if "live_bt_entries" not in st.session_state:
        st.session_state.live_bt_entries = []

    with st.form("add_bt_entry_form", clear_on_submit=True):
        _fc1, _fc2, _fc3, _fc4, _fc5 = st.columns([4, 1, 2, 2, 1])
        _bt_name = _fc1.selectbox("Strategy", _bt_all_names, key="bt_entry_name")
        _bt_contracts = _fc2.number_input("Contracts", min_value=1, max_value=999, value=1, step=1, key="bt_entry_contracts")
        _bt_start = _fc3.date_input("Start date", key="bt_entry_start")
        _bt_end   = _fc4.date_input("End date",   key="bt_entry_end")
        if _fc5.form_submit_button("Add", use_container_width=True):
            if _bt_end >= _bt_start:
                st.session_state.live_bt_entries.append({
                    "name":      _bt_name,
                    "contracts": int(_bt_contracts),
                    "start":     str(_bt_start),
                    "end":       str(_bt_end),
                })
                st.rerun()
            else:
                st.warning("End date must be on or after start date.")

    if st.session_state.live_bt_entries:
        _entries_df = pd.DataFrame(st.session_state.live_bt_entries)
        st.dataframe(_entries_df, use_container_width=True, hide_index=True)

        _rm_col, _ = st.columns([1, 4])
        if _rm_col.button("Clear all entries"):
            st.session_state.live_bt_entries = []
            st.rerun()

# ── Portfolio Periods mode ─────────────────────────────────────────────────────
else:
    st.markdown("### Define Portfolio Periods")
    st.caption(
        "Each period maps a saved portfolio snapshot to a date range. "
        "The Live strategies in each snapshot are used for that period. "
        "Periods may overlap — P&L is summed across all active strategies."
    )

    from core.portfolio.snapshot import list_snapshots as _list_snaps_bt, load_snapshot as _load_snap_bt

    if "live_bt_periods" not in st.session_state:
        st.session_state.live_bt_periods = []

    _snaps_bt = _list_snaps_bt()
    if not _snaps_bt:
        st.info(
            "No portfolio snapshots saved yet. "
            "Go to **Review Strategies → Performance Summary → 📸 Set Live Portfolio** to save your first snapshot."
        )
    else:
        _snap_labels = [s["label"] for s in _snaps_bt]

        with st.form("add_bt_period_form", clear_on_submit=True):
            _pc1, _pc2, _pc3, _pc4 = st.columns([3, 2, 2, 1])
            _bt_snap  = _pc1.selectbox("Portfolio snapshot", _snap_labels, key="bt_period_snap")
            _bt_pstart = _pc2.date_input("Start date", key="bt_period_start")
            _bt_pend   = _pc3.date_input("End date",   key="bt_period_end")
            if _pc4.form_submit_button("Add", use_container_width=True):
                if _bt_pend >= _bt_pstart:
                    st.session_state.live_bt_periods.append({
                        "snapshot": _bt_snap,
                        "start":    str(_bt_pstart),
                        "end":      str(_bt_pend),
                    })
                    st.rerun()
                else:
                    st.warning("End date must be on or after start date.")

        if st.session_state.live_bt_periods:
            for _pi, _per in enumerate(st.session_state.live_bt_periods):
                _pa, _pb, _pc = st.columns([4, 3, 1])
                _pa.markdown(f"**{_per['snapshot']}**")
                _pb.markdown(f"{_per['start']} → {_per['end']}")
                if _pc.button("Remove", key=f"rm_period_{_pi}"):
                    st.session_state.live_bt_periods.pop(_pi)
                    st.rerun()

            if st.button("Clear all periods"):
                st.session_state.live_bt_periods = []
                st.rerun()

# ── Compute & show equity curve ────────────────────────────────────────────────
st.divider()
_bt_equity_start = st.number_input(
    "Starting equity ($)", min_value=1_000.0, max_value=100_000_000.0,
    value=100_000.0, step=10_000.0, format="%.0f", key="bt_starting_equity",
)

if st.button("Compute Equity Curve", type="primary", key="bt_run"):
    _combined_pnl: pd.Series = pd.Series(dtype=float)

    if _bt_mode == "Manual Entries":
        _work_entries = st.session_state.get("live_bt_entries", [])
    else:
        from core.portfolio.snapshot import load_snapshot as _load_snap_bt2
        _work_entries = []
        for _per in st.session_state.get("live_bt_periods", []):
            _snap_strats = _load_snap_bt2(_per["snapshot"])
            _live_status = config.portfolio.live_status
            for _ss in _snap_strats:
                if _ss.get("status") == _live_status:
                    _work_entries.append({
                        "name":      _ss.get("name", ""),
                        "contracts": int(_ss.get("contracts") or 1),
                        "start":     _per["start"],
                        "end":       _per["end"],
                    })

    _missing: list[str] = []
    for _ent in _work_entries:
        _nm = _ent["name"]
        if _nm not in _bt_m2m.columns:
            _missing.append(_nm)
            continue
        _ts = pd.Timestamp(_ent["start"])
        _te = pd.Timestamp(_ent["end"])
        _pnl_slice = _bt_m2m.loc[_ts:_te, _nm] * int(_ent["contracts"])
        _combined_pnl = _combined_pnl.add(_pnl_slice, fill_value=0)

    if _missing:
        st.warning(f"No M2M data found for: {', '.join(set(_missing))}")

    if _combined_pnl.empty:
        st.info("No data to plot — check that your strategies and date ranges match the imported data.")
    else:
        _combined_pnl = _combined_pnl.sort_index()
        _equity_curve = _bt_equity_start + _combined_pnl.cumsum()

        # Key stats
        _total_pnl    = float(_combined_pnl.sum())
        _peak         = float(_equity_curve.cummax().max())
        _drawdown     = float((_equity_curve - _equity_curve.cummax()).min())
        _max_dd_pct   = abs(_drawdown / _bt_equity_start) * 100 if _bt_equity_start else 0

        _s1, _s2, _s3, _s4 = st.columns(4)
        _s1.metric("Total P&L",    f"${_total_pnl:,.0f}", delta=f"{_total_pnl/_bt_equity_start*100:.1f}%")
        _s2.metric("Peak Equity",  f"${_peak:,.0f}")
        _s3.metric("Max Drawdown", f"${_drawdown:,.0f}")
        _s4.metric("Max DD %",     f"{_max_dd_pct:.1f}%")

        _fig = px.line(
            _equity_curve,
            title="Live Portfolio Equity Curve",
            labels={"value": "Equity ($)", "index": "Date"},
        )
        _fig.update_layout(showlegend=False, height=400)
        _fig.add_hline(y=_bt_equity_start, line_dash="dash", line_color="grey", opacity=0.5)
        st.plotly_chart(_fig, use_container_width=True)

        # Per-strategy contribution table
        _contrib_rows = []
        for _ent in _work_entries:
            _nm = _ent["name"]
            if _nm not in _bt_m2m.columns:
                continue
            _ts = pd.Timestamp(_ent["start"])
            _te = pd.Timestamp(_ent["end"])
            _slice_pnl = _bt_m2m.loc[_ts:_te, _nm] * int(_ent["contracts"])
            _contrib_rows.append({
                "Strategy":     _nm,
                "Contracts":    _ent["contracts"],
                "Start":        _ent["start"],
                "End":          _ent["end"],
                "P&L ($)":      round(float(_slice_pnl.sum()), 0),
                "Trading Days": len(_slice_pnl.dropna()),
            })
        if _contrib_rows:
            st.subheader("Strategy Contributions")
            _contrib_df = pd.DataFrame(_contrib_rows).sort_values("P&L ($)", ascending=False)
            st.dataframe(_contrib_df, use_container_width=True, hide_index=True)
