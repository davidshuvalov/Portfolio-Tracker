"""
Strategies page — editable strategy configuration table + performance summary.
Mirrors the VBA Strategies tab (config) and Summary tab (performance metrics).
"""

import os
import platform
import subprocess
from pathlib import Path

import streamlit as st
import pandas as pd

from core.portfolio.strategies import load_strategies, save_strategies
from ui.strategy_labels import render_strategy_picker

st.set_page_config(page_title="Strategies", layout="wide")

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
tab_config, tab_summary, tab_backtest = st.tabs([
    "⚙ Configure", "📊 Performance Summary", "📈 Live Backtest"
])


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


# ── Strategy row-action helpers ────────────────────────────────────────────────
_CODE_EXTENSIONS = {".mex", ".eld", ".els", ".pla", ".c", ".cpp", ".py"}


def _open_path(path: Path) -> None:
    """Open a file or folder in the OS default application."""
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))
        elif platform.system() == "Darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])
    except Exception as exc:
        st.error(f"Could not open: {exc}")


def _strategy_folder(name: str) -> Path | None:
    """Return the folder Path for a strategy from the session scan_result."""
    scan = st.session_state.get("scan_result")
    if scan:
        for sf in scan.strategies:
            if sf.name == name:
                return sf.path
    return None


def _render_strategy_actions(names: list[str], key: str) -> None:
    """
    Compact action bar rendered below a strategies table.
    Lets the user pick a strategy and open its detail page, folder, or code file.
    """
    if not names:
        return

    st.markdown("**Open strategy**")
    col_pick, col_detail, col_folder, col_code = st.columns([4, 1, 1, 1])

    with col_pick:
        current = st.session_state.get("selected_strategy")
        default_idx = names.index(current) if current in names else 0
        chosen = st.selectbox(
            "Strategy", names, index=default_idx,
            key=f"{key}_picker", label_visibility="collapsed",
        )
        st.session_state.selected_strategy = chosen

    folder = _strategy_folder(chosen)
    code_files = (
        sorted([f for f in folder.iterdir() if f.suffix.lower() in _CODE_EXTENSIONS])
        if folder and folder.exists()
        else []
    )

    with col_detail:
        if st.button("📊 Detail", key=f"{key}_detail", use_container_width=True):
            st.session_state.selected_strategy = chosen
            st.switch_page("ui/pages/_Strategy_Detail.py")

    with col_folder:
        folder_exists = folder is not None and folder.exists()
        if st.button(
            "📂 Folder", key=f"{key}_folder",
            disabled=not folder_exists, use_container_width=True,
            help=str(folder) if folder else "Folder not found — import data first",
        ):
            _open_path(folder)

    with col_code:
        if len(code_files) == 1:
            if st.button("📄 Code", key=f"{key}_code", use_container_width=True,
                         help=code_files[0].name):
                _open_path(code_files[0])
        elif len(code_files) > 1:
            # Multiple files: show them in a small expander
            if st.button("📄 Code ▾", key=f"{key}_code", use_container_width=True,
                         help=f"{len(code_files)} code files — click to expand"):
                st.session_state[f"{key}_show_code"] = not st.session_state.get(f"{key}_show_code", False)
        else:
            st.button("📄 Code", key=f"{key}_code", disabled=True,
                      use_container_width=True, help="No code files found")

    # Multi-file picker (shown after clicking Code ▾)
    if len(code_files) > 1 and st.session_state.get(f"{key}_show_code"):
        with st.container():
            for i, cf in enumerate(code_files):
                c1, c2 = st.columns([5, 1])
                c1.markdown(f"`{cf.name}`")
                if c2.button("Open", key=f"{key}_cf_{i}"):
                    _open_path(cf)


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
    # Merge auto-detected direction from summary cache (read-only)
    _dir_cache = st.session_state.get("all_strategies_summary_cache")
    if _dir_cache is not None and "direction" in _dir_cache.columns:
        df["direction"] = df["name"].map(lambda n: _dir_cache["direction"].get(n, ""))
    else:
        df["direction"] = ""
    _cfg_cols = list(_COLUMNS) + ["direction"]
    _cfg_col_config = dict(_COLUMN_CONFIG)
    _cfg_col_config["direction"] = st.column_config.TextColumn("Direction", disabled=True, width="small")
    edited_df = st.data_editor(
        df[_cfg_cols],
        column_config=_cfg_col_config,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="strategies_editor",
    )

    # Save
    if st.button("Save Changes", type="primary"):
        edited_rows = edited_df.to_dict(orient="records")
        # Strip read-only computed columns before saving
        for _r in edited_rows:
            _r.pop("direction", None)
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

    # Strategy row actions
    _render_strategy_actions([s["name"] for s in filtered], key="cfg")

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
        bulk_contract_mode = st.radio(
            "Contracts for new strategies",
            options=["Estimated", "1", "Custom"],
            horizontal=True,
            key="bulk_contract_mode",
            help="Estimated uses ATR/margin blend from Contract Sizing settings.",
        )
        bulk_custom_contracts = 1
        if bulk_contract_mode == "Custom":
            bulk_custom_contracts = st.number_input(
                "Contracts", min_value=1, max_value=999, value=1, step=1,
                key="bulk_custom_contracts",
            )
        if bulk_status and st.button("Apply to New strategies"):
            _cfg_bulk = st.session_state.get("config", None)
            if _cfg_bulk is None:
                from core.config import AppConfig as _ACB
                _cfg_bulk = _ACB.load()
            _imported_bulk = st.session_state.get("imported_data")
            _new_strats = [s for s in strategies if s.get("status") == "New"]
            if bulk_contract_mode == "Estimated" and _imported_bulk is not None and _new_strats:
                from core.analytics.atr import estimate_contracts as _est_c
                _bulk_estimated = _est_c(_imported_bulk.trades, _new_strats, _cfg_bulk)
            else:
                _bulk_estimated = {}
            updated = []
            changed = 0
            for s in strategies:
                if s.get("status") == "New":
                    s = dict(s)
                    s["status"] = bulk_status
                    if bulk_contract_mode == "Estimated":
                        s["contracts"] = _bulk_estimated.get(s.get("name", ""), 1)
                    elif bulk_contract_mode == "1":
                        s["contracts"] = 1
                    else:
                        s["contracts"] = int(bulk_custom_contracts)
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
    from core.config import AppConfig as _AC
    from core.portfolio.snapshot import (
        compare_portfolios as _cmp_portfolios,
        list_snapshots as _list_snaps,
        load_snapshot as _load_snap,
        save_snapshot as _save_snap,
    )

    config = st.session_state.get("config", _AC.load())

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
            from core.ingestion.folder_scanner import scan_folders as _scan_f
            from core.portfolio.summary import compute_summary as _cs
            from datetime import date as _date

            _cfg = config
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
                    days_threshold=_cfg.eligibility.days_threshold_oos,
                    strategy_mc_config=_cfg.strategy_mc,
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
                _contract_mode = st.radio(
                    "Contract count for new strategies",
                    options=["Estimated", "1", "Custom"],
                    horizontal=True,
                    key="quick_add_contract_mode",
                    help="Estimated uses the ATR/margin blend from Contract Sizing settings.",
                )
                _custom_contracts = 1
                if _contract_mode == "Custom":
                    _custom_contracts = st.number_input(
                        "Contracts", min_value=1, max_value=999, value=1, step=1,
                        key="quick_add_custom_contracts",
                    )
                if st.button("➕ Add selected", key="quick_add_btn", disabled=not _to_add):
                    _cfg_now = st.session_state.get("config", _AC.load())
                    _imported_now = st.session_state.get("imported_data")
                    if _contract_mode == "Estimated" and _imported_now is not None:
                        from core.analytics.atr import estimate_contracts as _est_contracts
                        _to_add_dicts = [s for s in _all_strats_now if s.get("name") in _to_add]
                        _estimated = _est_contracts(_imported_now.trades, _to_add_dicts, _cfg_now)
                    else:
                        _estimated = {}
                    _updated = []
                    for _s in _all_strats_now:
                        if _s.get("name") in _to_add:
                            _s = dict(_s, status=_LIVE_STATUS)
                            if _contract_mode == "Estimated":
                                _s["contracts"] = _estimated.get(_s["name"], 1)
                            elif _contract_mode == "1":
                                _s["contracts"] = 1
                            else:
                                _s["contracts"] = int(_custom_contracts)
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

            # ── Exclude Buy & Hold strategies (they appear in Market Analysis) ──
            _bh_mask = _sm2["status"].str.lower().str.contains("buy", na=False) & \
                       _sm2["status"].str.lower().str.contains("hold", na=False)
            _bh_count = _bh_mask.sum()
            if _bh_count > 0:
                _sm2 = _sm2[~_bh_mask]
                st.info(
                    f"ℹ {_bh_count} Buy & Hold strategy{'s are' if _bh_count != 1 else ' is'} "
                    "excluded from this table — see **Market Analysis** for ATR and volatility data.",
                    icon=None,
                )
                st.page_link("ui/pages/_16_Market_Analysis.py", label="Open Market Analysis →")

            # ── Eligibility computation ────────────────────────────────────────
            from core.portfolio.summary import apply_eligibility_rules as _apply_elig
            _elig_mask = _apply_elig(_sm[~_bh_mask] if _bh_count > 0 else _sm, config.eligibility)
            _sm2["eligibility_status"] = _elig_mask.map({True: "Eligible", False: "Ineligible"})

            # All available summary columns with friendly labels
            _SUMM_ALL_COLS: dict[str, str] = {
                "symbol": "Symbol",
                "sector": "Sector",
                "direction": "Direction",
                "eligibility_status": "Eligibility",
                "last_date_on_file": "Last Date on File",
                "oos_begin": "OOS Start",
                "oos_end": "OOS End",
                "next_opt_date": "Next Opt Date",
                "last_opt_date": "Last Opt Date",
                "oos_period_years": "OOS Years",
                "expected_annual_profit": "Exp. Annual ($)",
                "actual_annual_profit": "Act. Annual ($)",
                "return_efficiency": "Efficiency",
                "mw_mc_is": "MW MC IS (%)",
                "mw_mc_isoos": "MW MC IS+OOS (%)",
                "mc_closed_is": "Closed MC IS (10%)",
                "mc_closed_isoos": "Closed MC IS+OOS (10%)",
                "strategy_mc_equity": "Strategy MC Equity ($)",
                "strategy_mc_max_dd": "Strategy MC Max DD %",
                "trades_per_year": "Trades/Yr",
                "overall_win_rate": "Win Rate",
                "sharpe_isoos": "Sharpe IS+OOS",
                "sharpe_is": "Sharpe IS",
                "max_drawdown_isoos": "Max DD IS+OOS ($)",
                "max_drawdown_is": "Max DD IS ($)",
                "profit_last_1_month": "Last 1M ($)",
                "profit_last_3_months": "Last 3M ($)",
                "profit_last_6_months": "Last 6M ($)",
                "profit_last_9_months": "Last 9M ($)",
                "profit_last_12_months": "Last 12M ($)",
                "profit_since_oos_start": "OOS P&L ($)",
                "max_oos_drawdown": "OOS Max DD ($)",
                "avg_oos_drawdown": "Avg OOS DD ($)",
                "rtd_oos": "R:DD OOS",
                "rtd_12_months": "R:DD 12M",
                "count_profit_months": "Profit Months",
                "incubation_status": "Incubation",
                "incubation_date": "Incub. Date",
                "quitting_status": "Quit Status",
                "quitting_date": "Quit Date",
                "profit_since_quit": "P&L Since Quit ($)",
                "k_factor": "K-Factor",
                "ulcer_index": "Ulcer Index",
                "best_month": "Best Month ($)",
                "worst_month": "Worst Month ($)",
                "max_consecutive_loss_months": "Max Loss Streak",
            }
            # Columns always shown (not in picker)
            _SUMM_FIXED = ["in_portfolio", "contracts", "status"]
            _SUMM_DEFAULT = [
                "symbol", "sector",
                "expected_annual_profit", "actual_annual_profit", "return_efficiency",
                "profit_last_3_months", "profit_last_12_months",
                "max_oos_drawdown", "rtd_oos", "incubation_status",
                "eligibility_status",
            ]

            _avail_summ_cols = [c for c in _SUMM_ALL_COLS if c in _sm2.columns]
            _avail_summ_labels = {c: _SUMM_ALL_COLS[c] for c in _avail_summ_cols}
            _summ_default_sel = [c for c in _SUMM_DEFAULT if c in _avail_summ_cols]

            _summ_col_key = "summ_metrics_col_picker"
            if _summ_col_key not in st.session_state:
                st.session_state[_summ_col_key] = _summ_default_sel

            with st.expander("⚙ Columns", expanded=False):
                _summ_sel_cols = st.multiselect(
                    "Select columns to display",
                    options=_avail_summ_cols,
                    format_func=lambda c: _avail_summ_labels.get(c, c),
                    key=_summ_col_key,
                )
                if st.button("Reset to defaults", key="summ_metrics_cols_reset"):
                    st.session_state[_summ_col_key] = _summ_default_sel
                    st.rerun()

            _picked = st.session_state.get(_summ_col_key) or _summ_default_sel
            _disp_cols = [c for c in _SUMM_FIXED + _picked if c in _sm2.columns]

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
                        "Efficiency", format="%.1f%%"
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
                    "eligibility_status": st.column_config.TextColumn("Eligibility"),
                    "direction": st.column_config.TextColumn("Direction"),
                    "last_date_on_file": st.column_config.DateColumn("Last Date on File"),
                    "next_opt_date": st.column_config.DateColumn("Next Opt Date"),
                    "last_opt_date": st.column_config.DateColumn("Last Opt Date"),
                    "mw_mc_is": st.column_config.NumberColumn("MW MC IS (%)", format="%.1f%%"),
                    "mw_mc_isoos": st.column_config.NumberColumn("MW MC IS+OOS (%)", format="%.1f%%"),
                    "mc_closed_is": st.column_config.NumberColumn(
                        "Closed MC IS (10%)", format="%.1f%%",
                        help="Closed-trade Monte Carlo max drawdown at 10% risk of ruin (IS only).",
                    ),
                    "mc_closed_isoos": st.column_config.NumberColumn(
                        "Closed MC IS+OOS (10%)", format="%.1f%%",
                        help="Closed-trade Monte Carlo max drawdown at 10% risk of ruin (IS+OOS).",
                    ),
                    # Additional columns available via column picker
                    "symbol": st.column_config.TextColumn("Symbol"),
                    "sector": st.column_config.TextColumn("Sector"),
                    "oos_period_years": st.column_config.NumberColumn("OOS Years", format="%.1f"),
                    "trades_per_year": st.column_config.NumberColumn("Trades/Yr", format="%.1f"),
                    "overall_win_rate": st.column_config.NumberColumn("Win Rate", format="%.1f%%"),
                    "sharpe_isoos": st.column_config.NumberColumn("Sharpe IS+OOS", format="%.2f"),
                    "sharpe_is": st.column_config.NumberColumn("Sharpe IS", format="%.2f"),
                    "max_drawdown_isoos": st.column_config.NumberColumn("Max DD IS+OOS ($)", format="$%.0f"),
                    "max_drawdown_is": st.column_config.NumberColumn("Max DD IS ($)", format="$%.0f"),
                    "profit_last_9_months": st.column_config.NumberColumn("Last 9M ($)", format="$%.0f"),
                    "avg_oos_drawdown": st.column_config.NumberColumn("Avg OOS DD ($)", format="$%.0f"),
                    "rtd_12_months": st.column_config.NumberColumn("R:DD 12M", format="%.2f"),
                    "count_profit_months": st.column_config.NumberColumn(
                        "Profit Months", format="%d",
                        help="# of profitable months in the eligibility lookback window.",
                    ),
                    "incubation_date": st.column_config.DateColumn("Incub. Date"),
                    "quitting_status": st.column_config.TextColumn(
                        "Quit Status",
                        help="Continue / Quit / Coming Back / Recovered / N/A",
                    ),
                    "quitting_date": st.column_config.DateColumn("Quit Date"),
                    "profit_since_quit": st.column_config.NumberColumn(
                        "P&L Since Quit ($)", format="$%.0f",
                        help="Cumulative OOS P&L from the date the strategy entered 'Quit' status.",
                    ),
                    "k_factor": st.column_config.NumberColumn(
                        "K-Factor", format="%.2f",
                        help="(Win rate / Loss rate) × (Avg win / Avg loss) — monthly P&L.",
                    ),
                    "ulcer_index": st.column_config.NumberColumn(
                        "Ulcer Index", format="%.2f",
                        help="RMS % drawdown over OOS period. Lower = smoother equity curve.",
                    ),
                    "best_month": st.column_config.NumberColumn("Best Month ($)", format="$%.0f"),
                    "worst_month": st.column_config.NumberColumn("Worst Month ($)", format="$%.0f"),
                    "max_consecutive_loss_months": st.column_config.NumberColumn(
                        "Max Loss Streak", format="%d",
                        help="Maximum consecutive losing months in the OOS period.",
                    ),
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

            # Strategy row actions
            _sm_names = list(_sm2.index)
            _render_strategy_actions(_sm_names, key="summ")

            # ── Re-optimisation alerts ─────────────────────────────────────────
            if "next_opt_date" in _sm2.columns:
                from datetime import date as _today_cls, timedelta as _td
                _today = _today_cls.today()
                _reopt_rows = []
                for _idx, _nod in _sm2["next_opt_date"].items():
                    if _nod is None or (hasattr(_nod, "__class__") and str(_nod) in ("NaT", "None", "nan")):
                        continue
                    try:
                        _nod_d = _nod if isinstance(_nod, _today_cls) else pd.Timestamp(_nod).date()
                        _days_left = (_nod_d - _today).days
                        if _days_left <= 21:
                            _lod = _sm2.at[_idx, "last_opt_date"] if "last_opt_date" in _sm2.columns else None
                            _reopt_rows.append({
                                "Strategy": _idx,
                                "Next Opt Date": str(_nod_d),
                                "Last Opt Date": str(_lod) if _lod else "—",
                                "Days Until": _days_left,
                                "Alert": "🔴 Overdue / ≤1 week" if _days_left <= 7 else "🟡 Within 3 weeks",
                            })
                    except Exception:
                        pass
                if _reopt_rows:
                    st.divider()
                    st.subheader("Re-optimisation Due")
                    _reopt_df = pd.DataFrame(_reopt_rows).sort_values("Days Until")

                    def _reopt_style(row):
                        if row.get("Days Until", 99) <= 7:
                            return ["background-color: #ffcdd2; color: #7f0000; font-weight: 600;"] * len(row)
                        return ["background-color: #fff9c4; color: #5d4037; font-weight: 500;"] * len(row)

                    st.dataframe(
                        _reopt_df.style.apply(_reopt_style, axis=1),
                        hide_index=True,
                        use_container_width=True,
                    )

            st.divider()
            st.page_link("ui/pages/03_Portfolio.py", label="→ Build Portfolio with Live strategies")

        elif _sm is not None:
            st.info(
                "No Walkforward data found. "
                "Make sure your import includes Walkforward Details CSVs."
            )
        else:
            st.info("Click **Compute / Refresh** to load per-strategy performance metrics.")


# ══════════════════════════════════════════════════════════════════════════════
# Live Backtest tab
# ══════════════════════════════════════════════════════════════════════════════
with tab_backtest:
    import plotly.express as px
    from core.portfolio.snapshot import list_snapshots as _list_snaps_bt, load_snapshot as _load_snap_bt

    st.caption(
        "Reconstruct your live trading performance by defining which strategies you were "
        "trading and for which periods. The equity curve is built from the imported daily M2M data."
    )

    _bt_imported = st.session_state.get("imported_data")
    if _bt_imported is None:
        st.info("Import data first to use the Live Backtest.")
        st.stop()

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

    # ── Manual mode ───────────────────────────────────────────────────────────
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

    # ── Portfolio Periods mode ────────────────────────────────────────────────
    else:
        st.markdown("### Define Portfolio Periods")
        st.caption(
            "Each period maps a saved portfolio snapshot to a date range. "
            "The Live strategies in each snapshot are used for that period. "
            "Periods may overlap — P&L is summed across all active strategies."
        )

        if "live_bt_periods" not in st.session_state:
            st.session_state.live_bt_periods = []

        _snaps_bt = _list_snaps_bt()
        if not _snaps_bt:
            st.info(
                "No portfolio snapshots saved yet. "
                "Go to **Performance Summary → 📸 Set Live Portfolio** to save your first snapshot."
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

    # ── Compute & show equity curve ───────────────────────────────────────────
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
            # Expand portfolio periods into entries
            _work_entries = []
            for _per in st.session_state.get("live_bt_periods", []):
                _snap_strats = _load_snap_bt(_per["snapshot"])
                _live_status = st.session_state.get("config", _AC.load()).portfolio.live_status
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
            _s1.metric("Total P&L",       f"${_total_pnl:,.0f}",  delta=f"{_total_pnl/_bt_equity_start*100:.1f}%")
            _s2.metric("Peak Equity",     f"${_peak:,.0f}")
            _s3.metric("Max Drawdown",    f"${_drawdown:,.0f}")
            _s4.metric("Max DD %",        f"{_max_dd_pct:.1f}%")

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
                    "Strategy":    _nm,
                    "Contracts":   _ent["contracts"],
                    "Start":       _ent["start"],
                    "End":         _ent["end"],
                    "P&L ($)":     round(float(_slice_pnl.sum()), 0),
                    "Trading Days": len(_slice_pnl.dropna()),
                })
            if _contrib_rows:
                st.subheader("Strategy Contributions")
                _contrib_df = pd.DataFrame(_contrib_rows).sort_values("P&L ($)", ascending=False)
                st.dataframe(_contrib_df, use_container_width=True, hide_index=True)
