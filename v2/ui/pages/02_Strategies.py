"""
Strategies page — editable strategy configuration table + performance summary.
Mirrors the VBA Strategies tab (config) and Summary tab (performance metrics).
"""

import streamlit as st
import pandas as pd

from core.portfolio.strategies import load_strategies, save_strategies
from ui.strategy_labels import render_strategy_picker

st.set_page_config(page_title="Strategies", layout="wide")

# ── Sidebar workflow status ────────────────────────────────────────────────────
try:
    from ui.workflow import render_workflow_sidebar
    with st.sidebar:
        render_workflow_sidebar()
except Exception:
    pass

st.title("Strategies")

# ── Top navigation ─────────────────────────────────────────────────────────────
_nav_l, _nav_r, _ = st.columns([1, 1, 6])
with _nav_l:
    st.page_link("ui/pages/01_Import.py", label="← Import")
with _nav_r:
    st.page_link("ui/pages/03_Portfolio.py", label="→ Build Portfolio")

# ── Load strategies ────────────────────────────────────────────────────────────
strategies = load_strategies()

if not strategies:
    st.info("No strategies found yet. Scan your folders on the Import page first.")
    st.page_link("ui/pages/01_Import.py", label="Go to Import →")
    st.stop()

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab_config, tab_summary = st.tabs(["⚙ Configure", "📊 Performance Summary"])


# ── Column definitions (shared) ────────────────────────────────────────────────
_COLUMNS = [
    "name", "status", "contracts", "symbol", "sector",
    "timeframe", "type", "horizon", "other", "notes",
]

_COLUMN_CONFIG = {
    "name": st.column_config.TextColumn(
        "Strategy", disabled=True, width="large"
    ),
    "status": st.column_config.SelectboxColumn(
        "Status",
        options=[
            "Live", "Paper", "Retired", "Pass",
            "Buy&Hold", "Incubating", "New",
            "Not Loaded - Live", "Not Loaded - Paper",
            "Not Loaded - Retired", "Not Loaded - Pass",
        ],
        required=True,
        width="medium",
    ),
    "contracts": st.column_config.NumberColumn(
        "Contracts", min_value=0, max_value=999, step=1,
        format="%d", width="small",
    ),
    "symbol": st.column_config.TextColumn("Symbol", width="small"),
    "sector": st.column_config.SelectboxColumn(
        "Sector",
        options=[
            "", "Index", "Energy", "Metals", "Currencies", "Interest Rate",
            "Agriculture", "Soft", "Meats", "Crypto", "Volatility",
            "Eurex Index", "Eurex Interest Rate", "Euronext LIFFE", "Other",
        ],
        width="medium",
    ),
    "timeframe": st.column_config.TextColumn("Timeframe", width="small"),
    "type": st.column_config.SelectboxColumn(
        "Type",
        options=["", "Trend", "Mean Reversion", "Seasonal", "Arbitrage", "Other"],
        width="medium",
    ),
    "horizon": st.column_config.SelectboxColumn(
        "Horizon",
        options=["", "Short", "Medium", "Long"],
        width="small",
    ),
    "other": st.column_config.TextColumn("Other", width="small"),
    "notes": st.column_config.TextColumn("Notes", width="large"),
}


def _to_df(strats: list[dict]) -> pd.DataFrame:
    rows = []
    for s in strats:
        row = {col: s.get(col, "") for col in _COLUMNS}
        try:
            row["contracts"] = int(row["contracts"] or 1)
        except (ValueError, TypeError):
            row["contracts"] = 1
        rows.append(row)
    return pd.DataFrame(rows, columns=_COLUMNS)


# ══════════════════════════════════════════════════════════════════════════════
# Configure tab
# ══════════════════════════════════════════════════════════════════════════════
with tab_config:
    st.caption("Step 3 of 4 — set each strategy's status, contracts, symbol, and sector. Mark active strategies as Live.")

    # Filters
    with st.expander("Filter / Search", expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            search = st.text_input("Search name", placeholder="e.g. ES_Trend")
        with col2:
            all_statuses = sorted({s.get("status", "") for s in strategies if s.get("status")})
            filter_status = st.multiselect("Filter by status", options=all_statuses)
        with col3:
            all_sectors = sorted({s.get("sector", "") for s in strategies if s.get("sector")})
            filter_sector = st.multiselect("Filter by sector", options=all_sectors)

    filtered = strategies
    if search:
        filtered = [s for s in filtered if search.lower() in s.get("name", "").lower()]
    if filter_status:
        filtered = [s for s in filtered if s.get("status") in filter_status]
    if filter_sector:
        filtered = [s for s in filtered if s.get("sector") in filter_sector]

    is_filtered = bool(search or filter_status or filter_sector)
    st.caption(
        f"Showing **{len(filtered)}** of **{len(strategies)}** strategies"
        + (" (filtered)" if is_filtered else "")
    )

    df = _to_df(filtered)
    edited_df = st.data_editor(
        df,
        column_config=_COLUMN_CONFIG,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="strategies_editor",
    )

    # Save
    if st.button("Save Changes", type="primary"):
        edited_rows = edited_df.to_dict(orient="records")
        edited_by_name = {r["name"]: r for r in edited_rows}
        merged = []
        for s in strategies:
            name = s.get("name", "")
            if name in edited_by_name:
                updated = dict(s)
                updated.update(edited_by_name[name])
                merged.append(updated)
            else:
                merged.append(s)
        save_strategies(merged)
        st.session_state.portfolio_data = None
        live_count = sum(1 for r in edited_rows if r.get("status") == "Live")
        st.success(
            f"Saved {len(edited_rows)} records. "
            + (f"{live_count} Live strategies — portfolio will need a rebuild." if live_count else "")
        )
        if live_count:
            st.page_link("ui/pages/03_Portfolio.py", label="→ Rebuild Portfolio")
        st.rerun()

    st.divider()

    # Bulk operations
    st.subheader("Bulk Operations")
    col_bulk1, col_bulk2, col_bulk3 = st.columns(3)

    with col_bulk1:
        bulk_status = st.selectbox(
            "Set status for all *New* strategies",
            options=["", "Live", "Paper", "Retired", "Pass"],
            index=0,
            key="bulk_status",
        )
        if bulk_status and st.button("Apply to New strategies"):
            updated = []
            changed = 0
            for s in strategies:
                if s.get("status") == "New":
                    s = dict(s)
                    s["status"] = bulk_status
                    changed += 1
                updated.append(s)
            save_strategies(updated)
            st.success(f"Updated {changed} strategies to '{bulk_status}'.")
            st.rerun()

    with col_bulk2:
        if st.button("Reset all contracts to 1"):
            updated = [dict(s, contracts=1) for s in strategies]
            save_strategies(updated)
            st.success("All contracts reset to 1.")
            st.rerun()

    with col_bulk3:
        not_loaded = [s for s in strategies if "Not Loaded" in s.get("status", "")]
        if not_loaded:
            st.warning(f"{len(not_loaded)} strategies not found in folders.")
            if st.button("Remove Not Loaded strategies"):
                kept = [s for s in strategies if "Not Loaded" not in s.get("status", "")]
                save_strategies(kept)
                st.success(f"Removed {len(not_loaded)} not-loaded strategies.")
                st.rerun()

    # Auto-fill sectors
    st.divider()
    st.subheader("Auto-fill Sectors from v1.24 Reference")

    from core.ingestion.xlsb_importer import load_margin_tables  # noqa: E402
    if "margin_tables" not in st.session_state:
        st.session_state["margin_tables"] = load_margin_tables()
    _mt = st.session_state.get("margin_tables")

    if _mt is not None and _mt.sector_lookup:
        missing_sector = [s for s in strategies if not s.get("sector") and s.get("symbol")]
        fillable = [s for s in missing_sector if s.get("symbol") in _mt.sector_lookup]
        if fillable:
            st.caption(
                f"**{len(fillable)}** strategies have a symbol but no sector, and their "
                f"symbol is in the v1.24 reference data."
            )
            if st.button(f"Auto-fill sectors for {len(fillable)} strategies", type="primary"):
                updated = []
                for s in strategies:
                    sym = s.get("symbol", "")
                    if not s.get("sector") and sym and sym in _mt.sector_lookup:
                        s = dict(s, sector=_mt.sector_lookup[sym])
                    updated.append(s)
                save_strategies(updated)
                st.success(f"Filled sectors for {len(fillable)} strategies.")
                st.rerun()
        else:
            st.caption(
                "All strategies with symbols already have sectors, "
                "or their symbols are not in the reference data."
            )
    else:
        st.caption(
            "No v1.24 reference data found. Run the **Migrate** page with your "
            "`.xlsb` file to import the Sector reference table."
        )


# ── Quick stats sidebar ────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Summary")

    status_counts: dict[str, int] = {}
    for s in strategies:
        stat = s.get("status", "Unknown")
        status_counts[stat] = status_counts.get(stat, 0) + 1

    total = len(strategies)
    live = status_counts.get("Live", 0)
    paper = status_counts.get("Paper", 0)
    new_ = status_counts.get("New", 0)

    st.metric("Total Strategies", total)
    st.metric("Live", live)
    if paper:
        st.metric("Paper", paper)
    if new_:
        st.metric("New (unconfirmed)", new_)

    st.divider()
    st.caption("Status breakdown")
    for _stat, _count in sorted(status_counts.items()):
        pct = _count / total * 100 if total else 0
        st.write(f"**{_stat}**: {_count} ({pct:.0f}%)")

    _imported_for_picker = st.session_state.get("imported_data")
    if _imported_for_picker and _imported_for_picker.strategies:
        st.divider()
        render_strategy_picker(_imported_for_picker.strategies, key="strats_strat_picker")


# ══════════════════════════════════════════════════════════════════════════════
# Performance Summary tab — mirrors the Excel Summary tab
# Shows per-strategy WF metrics for ALL strategies (not filtered to Live)
# ══════════════════════════════════════════════════════════════════════════════
with tab_summary:
    from core.portfolio.snapshot import (
        compare_portfolios as _cmp_portfolios,
        list_snapshots as _list_snaps,
        load_snapshot as _load_snap,
        save_snapshot as _save_snap,
    )

    st.caption(
        "All strategies with performance metrics. "
        "Tick **In Portfolio** to include in the Live portfolio. "
        "Click **Compute / Refresh** after each import to update metrics."
    )

    _imported = st.session_state.get("imported_data")

    if _imported is None:
        st.info("Import data first to see performance metrics.")
        st.page_link("ui/pages/01_Import.py", label="→ Go to Import")
    else:
        _SUMMARY_KEY = "all_strategies_summary_cache"
        _LIVE_STATUS = config.portfolio.live_status  # typically "Live"

        # ── Toolbar row ────────────────────────────────────────────────────────
        _tb1, _tb2, _tb3, _tb4 = st.columns([1, 1, 1, 4])
        with _tb1:
            _do_compute = st.button("Compute / Refresh", type="primary", key="compute_summary_btn")
        with _tb2:
            if st.session_state.get(_SUMMARY_KEY) is not None:
                if st.button("Clear Cache", key="clear_summary_btn"):
                    st.session_state[_SUMMARY_KEY] = None
                    st.rerun()
        with _tb3:
            _snap_btn = st.button(
                "📸 Set Live Portfolio",
                key="set_live_portfolio_btn",
                help="Save the current Live strategies as a snapshot and show trading instructions vs the previous baseline.",
            )

        if _do_compute:
            from core.config import AppConfig as _AC
            from core.ingestion.folder_scanner import scan_folders as _scan_f
            from core.portfolio.summary import compute_summary as _cs
            from datetime import date as _date

            _cfg = st.session_state.get("config", _AC.load())
            with st.spinner("Computing strategy metrics..."):
                _scan_res = _scan_f(_cfg.folders) if _cfg.folders else None
                _sfl = _scan_res.strategies if _scan_res else []
                _cutoff = None
                if _cfg.portfolio.use_cutoff and _cfg.portfolio.cutoff_date:
                    try:
                        _cutoff = _date.fromisoformat(_cfg.portfolio.cutoff_date)
                    except ValueError:
                        pass
                st.session_state[_SUMMARY_KEY] = _cs(
                    imported=_imported,
                    strategy_folders=_sfl,
                    date_format=_cfg.date_format,
                    use_cutoff=_cfg.portfolio.use_cutoff,
                    cutoff_date=_cutoff,
                )
            st.success(f"Summary computed for {len(st.session_state[_SUMMARY_KEY])} strategies.")

        # ── Set Live Portfolio handler ──────────────────────────────────────────
        if _snap_btn:
            _current_strats = load_strategies()
            _snaps = _list_snaps()
            _prev_ref = _load_snap(_snaps[0]["filename"]) if _snaps else []
            _result = _cmp_portfolios(_current_strats, _prev_ref, live_status=_LIVE_STATUS)

            from datetime import datetime as _dt
            _label = _dt.now().strftime("%Y-%m-%d %H:%M")
            _save_snap(_current_strats, _label)

            _live_now = [s for s in _current_strats if s.get("status") == _LIVE_STATUS]
            st.success(f"📸 Live portfolio saved — {len(_live_now)} strategies ({_label})")

            if _result.has_changes:
                st.subheader("Trading Instructions vs previous baseline")
                st.caption(
                    "These are the changes you need to make in your trading system "
                    "to bring it in line with the new portfolio."
                )

                if _result.new_strategies:
                    st.markdown("#### ✅ Enable in trading system (new Live strategies)")
                    for _ns in _result.new_strategies:
                        _sym = _ns.get("symbol", "")
                        _c = _ns.get("contracts", 1)
                        st.markdown(f"- **{_ns['name']}**  ({_sym}, {_c} contract{'s' if _c != 1 else ''})")

                if _result.removed_strategies:
                    st.markdown("#### ❌ Disable in trading system (removed from portfolio)")
                    for _rs in _result.removed_strategies:
                        st.markdown(f"- **{_rs['name']}**  ({_rs.get('symbol', '')})")

                if _result.contract_changes:
                    st.markdown("#### 🔄 Adjust contracts in trading system")
                    for _chg in _result.contract_changes:
                        _arrow = "▲" if _chg["delta"] > 0 else "▼"
                        st.markdown(
                            f"- **{_chg['name']}**  ({_chg['symbol']})  "
                            f"{_chg['old_contracts']} → **{_chg['new_contracts']} contracts** {_arrow}"
                        )
            else:
                if _snaps:
                    st.info("No changes vs previous baseline — trading system already matches.")
                else:
                    st.info("Baseline saved for the first time. Future comparisons will reference this.")

        # ── Quick Portfolio Editor ─────────────────────────────────────────────
        _all_strats_now = load_strategies()
        _live_names = [s["name"] for s in _all_strats_now if s.get("status") == _LIVE_STATUS]
        _non_live_names = [s["name"] for s in _all_strats_now if s.get("status") != _LIVE_STATUS]

        with st.expander(
            f"Quick Portfolio Editor — {len(_live_names)} Live strategies", expanded=False
        ):
            st.caption(
                "Quickly add or remove strategies from the portfolio. "
                "Changes are saved immediately. Removed strategies are set to **Pass**."
            )
            _qc1, _qc2 = st.columns(2)

            with _qc1:
                st.markdown("**Add to portfolio**")
                _to_add = st.multiselect(
                    "Select strategies to add as Live",
                    options=_non_live_names,
                    key="quick_add_select",
                    placeholder="Search strategies…",
                )
                if st.button("➕ Add selected", key="quick_add_btn", disabled=not _to_add):
                    _updated = []
                    for _s in _all_strats_now:
                        if _s.get("name") in _to_add:
                            _s = dict(_s, status=_LIVE_STATUS)
                        _updated.append(_s)
                    save_strategies(_updated)
                    st.session_state.portfolio_data = None
                    st.success(f"Added {len(_to_add)} strategies to portfolio.")
                    st.rerun()

            with _qc2:
                st.markdown("**Remove from portfolio**")
                _to_remove = st.multiselect(
                    "Select Live strategies to remove",
                    options=_live_names,
                    key="quick_remove_select",
                    placeholder="Search strategies…",
                )
                if st.button("➖ Remove selected", key="quick_remove_btn", disabled=not _to_remove):
                    _updated = []
                    for _s in _all_strats_now:
                        if _s.get("name") in _to_remove:
                            _s = dict(_s, status="Pass")
                        _updated.append(_s)
                    save_strategies(_updated)
                    st.session_state.portfolio_data = None
                    st.success(f"Removed {len(_to_remove)} strategies from portfolio (set to Pass).")
                    st.rerun()

        # ── Summary metrics table ──────────────────────────────────────────────
        _sm = st.session_state.get(_SUMMARY_KEY)

        if _sm is not None and not _sm.empty:
            # Merge status + contracts from config
            _strats_map = {s.get("name"): s for s in strategies}
            _sm2 = _sm.copy()

            # in_portfolio: True if status == Live (the key editable toggle)
            _sm2.insert(0, "in_portfolio", _sm2.index.map(
                lambda n: _strats_map.get(n, {}).get("status", "") == _LIVE_STATUS
            ))
            _sm2.insert(1, "contracts", _sm2.index.map(
                lambda n: int(_strats_map.get(n, {}).get("contracts") or 1)
            ))
            _sm2.insert(2, "status", _sm2.index.map(
                lambda n: _strats_map.get(n, {}).get("status", "")
            ))

            _disp_cols = [c for c in [
                "in_portfolio", "contracts", "status", "symbol", "sector",
                "oos_begin", "oos_end",
                "expected_annual_profit", "actual_annual_profit", "return_efficiency",
                "profit_last_1_month", "profit_last_3_months",
                "profit_last_6_months", "profit_last_12_months",
                "profit_since_oos_start", "max_oos_drawdown", "rtd_oos",
                "incubation_status",
            ] if c in _sm2.columns]

            _disp = _sm2[_disp_cols].reset_index()
            _disp.rename(columns={"strategy_name": "Strategy"}, inplace=True)

            _readonly_cols = [c for c in _disp.columns if c not in ("in_portfolio", "contracts", "status")]

            _edited_summary = st.data_editor(
                _disp,
                use_container_width=True,
                hide_index=True,
                disabled=_readonly_cols,
                key="summary_editor",
                column_config={
                    "in_portfolio": st.column_config.CheckboxColumn(
                        "In Portfolio",
                        help="Tick to include in the Live portfolio. Untick to set to Pass.",
                        width="small",
                    ),
                    "status": st.column_config.SelectboxColumn(
                        "Status",
                        options=[
                            "Live", "Paper", "Retired", "Pass",
                            "Buy&Hold", "Incubating", "New",
                            "Not Loaded - Live", "Not Loaded - Paper",
                            "Not Loaded - Retired", "Not Loaded - Pass",
                        ],
                        width="small",
                    ),
                    "contracts": st.column_config.NumberColumn(
                        "Contr.", format="%d", min_value=0, max_value=999, step=1, width="small"
                    ),
                    "oos_begin": st.column_config.DateColumn("OOS Start"),
                    "oos_end": st.column_config.DateColumn("OOS End"),
                    "expected_annual_profit": st.column_config.NumberColumn(
                        "Exp. Annual ($)", format="$%.0f"
                    ),
                    "actual_annual_profit": st.column_config.NumberColumn(
                        "Act. Annual ($)", format="$%.0f"
                    ),
                    "return_efficiency": st.column_config.NumberColumn(
                        "Efficiency", format="%.1%%"
                    ),
                    "profit_last_1_month": st.column_config.NumberColumn(
                        "Last 1M ($)", format="$%.0f"
                    ),
                    "profit_last_3_months": st.column_config.NumberColumn(
                        "Last 3M ($)", format="$%.0f"
                    ),
                    "profit_last_6_months": st.column_config.NumberColumn(
                        "Last 6M ($)", format="$%.0f"
                    ),
                    "profit_last_12_months": st.column_config.NumberColumn(
                        "Last 12M ($)", format="$%.0f"
                    ),
                    "profit_since_oos_start": st.column_config.NumberColumn(
                        "OOS P&L ($)", format="$%.0f"
                    ),
                    "max_oos_drawdown": st.column_config.NumberColumn(
                        "OOS Max DD ($)", format="$%.0f"
                    ),
                    "rtd_oos": st.column_config.NumberColumn("R:DD OOS", format="%.2f"),
                    "incubation_status": st.column_config.TextColumn("Incubation"),
                },
            )

            _save_col, _export_col, _nav_col = st.columns([1, 1, 4])
            with _save_col:
                _do_save = st.button("Save Changes", type="primary", key="save_summary_btn")
            with _export_col:
                if st.button("Export to Excel", key="export_summary_xlsx_btn"):
                    from core.reporting.excel_export import (
                        export_summary_metrics,
                        summary_metrics_export_filename,
                    )
                    _xlsx = export_summary_metrics(_sm2.drop(columns=["in_portfolio"], errors="ignore"), strategies)
                    st.download_button(
                        "📥 Download",
                        data=_xlsx,
                        file_name=summary_metrics_export_filename(),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_summary_xlsx",
                    )

            if _do_save:
                _edits_by_name = {
                    r["Strategy"]: r for r in _edited_summary.to_dict(orient="records")
                }
                _merged = []
                for _s in strategies:
                    _nm = _s.get("name", "")
                    if _nm in _edits_by_name:
                        _e = _edits_by_name[_nm]
                        _s = dict(_s)
                        _in_port = _e.get("in_portfolio", False)
                        _explicit_status = _e.get("status", _s.get("status", ""))
                        # in_portfolio checkbox takes precedence:
                        #   checked  → Live
                        #   unchecked & was Live  → Pass
                        #   unchecked & was not Live → keep explicit status
                        if _in_port:
                            _s["status"] = _LIVE_STATUS
                        elif _explicit_status == _LIVE_STATUS and not _in_port:
                            _s["status"] = "Pass"
                        else:
                            _s["status"] = _explicit_status
                        try:
                            _s["contracts"] = int(_e.get("contracts") or 1)
                        except (ValueError, TypeError):
                            pass
                    _merged.append(_s)
                save_strategies(_merged)
                st.session_state.portfolio_data = None
                _live_n = sum(1 for _s in _merged if _s.get("status") == _LIVE_STATUS)
                st.success(
                    f"Saved — **{_live_n} Live** strategies. "
                    "Click **📸 Set Live Portfolio** to record this as your trading baseline."
                )
                st.rerun()

            st.divider()
            st.page_link("ui/pages/03_Portfolio.py", label="→ Build Portfolio with Live strategies")

        elif _sm is not None:
            st.info(
                "No Walkforward data found. "
                "Make sure your import includes Walkforward Details CSVs."
            )
        else:
            st.info("Click **Compute / Refresh** to load per-strategy performance metrics.")
