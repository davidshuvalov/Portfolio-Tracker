"""
Import page — scan MultiWalk folders, load strategy CSV data, and configure strategies.
Mirrors the VBA 'Retrieve Folder Data' + 'Import Data' + Strategy Configure workflow.
"""

import datetime
import streamlit as st
import pandas as pd
from pathlib import Path

from core.config import AppConfig
from core.ingestion.folder_scanner import scan_folders, reconcile_statuses
from core.ingestion.csv_importer import import_all
from core.portfolio.strategies import load_strategies, save_strategies

st.set_page_config(page_title="Import", layout="wide")
st.title("Import")

config: AppConfig = st.session_state.get("config", AppConfig.load())

tab_import, tab_configure = st.tabs(["📥 Import Data", "⚙ Configure Strategies"])


# ══════════════════════════════════════════════════════════════════
# Tab 1 — Import Data
# ══════════════════════════════════════════════════════════════════

with tab_import:
    st.caption("Steps 1 & 2 of 4 — add folders, then load strategy CSV data.")

    # ─── Section 1: Base Folder Configuration ─────────────────────
    st.subheader("MultiWalk Base Folders")
    st.caption(
        "Add the folders that contain your MultiWalk strategy subfolders. "
        "Equivalent to Folder1–Folder10 and FolderBH in v1.24."
    )

    _STATUS_OPTIONS = ["New", "Live", "Paper", "Pass", "Retired", "Buy&Hold"]

    if config.folders:
        _fdr_hdr = st.columns([5, 2, 1])
        _fdr_hdr[0].caption("Folder path")
        _fdr_hdr[1].caption("Default status for new strategies")
        _fdr_hdr[2].caption("Remove")
        for folder in config.folders:
            exists = folder.exists()
            icon = "✓" if exists else "✗"
            colour = "green" if exists else "red"
            _fc1, _fc2, _fc3 = st.columns([5, 2, 1])
            with _fc1:
                st.markdown(
                    f":{colour}[{icon}] `{folder}`",
                    help="Folder exists" if exists else "Folder not found on disk",
                )
            with _fc2:
                _cur_default = config.folder_default_status.get(str(folder), "New")
                _safe_idx = _STATUS_OPTIONS.index(_cur_default) if _cur_default in _STATUS_OPTIONS else 0
                _new_default = st.selectbox(
                    "Default",
                    _STATUS_OPTIONS,
                    index=_safe_idx,
                    key=f"folder_status_{folder}",
                    label_visibility="collapsed",
                    help="Status assigned to strategies from this folder when first discovered.",
                )
                if _new_default != _cur_default:
                    config.set_folder_default_status(folder, _new_default)
                    st.session_state.config = config
            with _fc3:
                if st.button("✕", key=f"remove_{folder}", help=f"Remove {folder}"):
                    config.remove_folder(folder)
                    st.session_state.config = config
                    st.rerun()
    else:
        st.info("No folders configured yet. Add a folder below.")

    with st.form("add_folder_form", clear_on_submit=True):
        _af1, _af2 = st.columns([4, 2])
        with _af1:
            new_folder = st.text_input(
                "Add folder path",
                placeholder=r"C:\MultiWalk\Strategies",
            )
        with _af2:
            new_folder_status = st.selectbox(
                "Default status",
                _STATUS_OPTIONS,
                index=0,
                help="Status assigned to strategies found in this folder when they are first discovered.",
            )
        submitted = st.form_submit_button("Add Folder")
        if submitted and new_folder:
            p = Path(new_folder.strip())
            if not p.exists():
                st.error(f"Folder not found: {p}")
            elif p in config.folders:
                st.warning("Folder already in list.")
            else:
                config.add_folder(p, default_status=new_folder_status)
                st.session_state.config = config
                st.success(f"Added: {p} (default status: **{new_folder_status}**)")
                st.rerun()

    if config.folders:
        st.success("**Step 1 complete** — folders configured. Now scan and import data below.")

    st.divider()

    # ─── Section 2: Data Settings ─────────────────────────────────
    st.subheader("Data Settings")

    col1, col2, col3 = st.columns(3)
    with col1:
        date_format = st.selectbox(
            "CSV Date Format",
            options=["DMY", "MDY"],
            index=0 if config.date_format == "DMY" else 1,
            help="DMY = day/month/year (EU/UK/AU). MDY = month/day/year (US).",
        )
    with col2:
        use_cutoff = st.toggle("Use Cutoff Date", value=config.portfolio.use_cutoff)
    with col3:
        cutoff_date = None
        if use_cutoff:
            cutoff_date = st.date_input(
                "Cutoff Date",
                value=pd.Timestamp(config.portfolio.cutoff_date).date()
                if config.portfolio.cutoff_date
                else datetime.date.today(),
            )

    if (
        date_format != config.date_format
        or use_cutoff != config.portfolio.use_cutoff
        or (use_cutoff and str(cutoff_date) != config.portfolio.cutoff_date)
    ):
        config.date_format = date_format
        config.portfolio.use_cutoff = use_cutoff
        config.portfolio.cutoff_date = str(cutoff_date) if use_cutoff else None
        config.save()
        st.session_state.config = config

    st.divider()

    # ─── Section 3: Scan & Import ──────────────────────────────────
    st.subheader("Import Data")

    if not config.folders:
        st.warning("Add at least one base folder above before scanning.")
        st.stop()

    col_scan, col_import, _ = st.columns([1, 1, 4])
    with col_scan:
        scan_clicked = st.button("Scan Folders", use_container_width=True)
    with col_import:
        import_clicked = st.button(
            "Import Data",
            type="primary",
            use_container_width=True,
            disabled=not bool(config.folders),
        )

    if scan_clicked:
        with st.spinner("Scanning folders..."):
            result = scan_folders(config.folders)
        st.session_state.scan_result = result

        if result.errors:
            for e in result.errors:
                st.error(e)
        if result.warnings:
            with st.expander(f"{len(result.warnings)} warning(s)"):
                for w in result.warnings:
                    st.warning(w)

        if result.strategies:
            st.success(f"Found **{len(result.strategies)}** strategies.")
            configured = load_strategies()
            found_names = {sf.name for sf in result.strategies}
            reconciled = reconcile_statuses(
                found_names,
                configured,
                strategy_folders=result.strategies,
                folder_default_status=config.folder_default_status,
            )
            save_strategies(reconciled)

            from core.ingestion.csv_importer import _is_multi_strategy_file
            rows = []
            for sf in result.strategies:
                is_multi = _is_multi_strategy_file(sf.equity_csv, config.date_format)
                rows.append({
                    "Strategy": sf.name,
                    "Folder": sf.path.name,
                    "Multi-Strategy": "Yes" if is_multi else "",
                    "Trade Data": "Yes" if sf.trade_csv else "No",
                    "WF Details": "Yes" if sf.walkforward_csv else "No",
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

            _paths = "\n".join(str(sf.path) for sf in result.strategies)
            with st.expander("Export folder paths (for MultiWalk)"):
                st.text_area(
                    "Strategy folder paths — one per line",
                    value=_paths,
                    height=200,
                    label_visibility="collapsed",
                    help="Copy and paste these paths into MultiWalk.",
                )
                st.download_button(
                    "Download as .txt",
                    data=_paths,
                    file_name="strategy_folders.txt",
                    mime="text/plain",
                    use_container_width=True,
                )
        else:
            st.warning("No strategies found. Check your folder paths.")

    if import_clicked:
        cutoff = None
        if use_cutoff and cutoff_date:
            cutoff = cutoff_date if isinstance(cutoff_date, datetime.date) else None

        if st.session_state.get("scan_result") is None:
            with st.spinner("Scanning folders…"):
                _auto = scan_folders(config.folders)
            st.session_state.scan_result = _auto
        _strategy_folders = st.session_state.scan_result.strategies

        n_strats_to_import = len(_strategy_folders)
        progress = st.progress(0, text=f"Importing 0 / {n_strats_to_import} strategies…")
        status_text = st.empty()

        def _on_progress(idx: int, total: int, name: str) -> None:
            frac = idx / max(total, 1)
            progress.progress(frac, text=f"Reading {name} ({idx}/{total})…")

        try:
            imported, warnings = import_all(
                _strategy_folders,
                date_format=config.date_format,
                use_cutoff=use_cutoff,
                cutoff_date=cutoff,
                progress_cb=_on_progress,
            )

            if warnings:
                with st.expander(f"{len(warnings)} import warning(s)"):
                    for w in warnings:
                        st.warning(w)

            _configured_map = {s["name"]: s for s in load_strategies()}
            for _strat in imported.strategies:
                _cfg = _configured_map.get(_strat.name, {})
                if _cfg.get("symbol"):    _strat.symbol    = _cfg["symbol"]
                if _cfg.get("sector"):    _strat.sector    = _cfg["sector"]
                if _cfg.get("timeframe"): _strat.timeframe = _cfg["timeframe"]
                if _cfg.get("type"):      _strat.type      = _cfg["type"]
                if _cfg.get("horizon"):   _strat.horizon   = _cfg["horizon"]
                if _cfg.get("status"):    _strat.status    = _cfg["status"]
                if _cfg.get("contracts"): _strat.contracts = int(_cfg["contracts"])
                if _cfg.get("notes"):     _strat.notes     = _cfg["notes"]

            _actual_names = set(imported.strategy_names)
            _post_configured = load_strategies()
            _post_reconciled = reconcile_statuses(
                _actual_names,
                _post_configured,
                strategy_folders=_strategy_folders,
                folder_default_status=config.folder_default_status,
            )
            save_strategies(_post_reconciled)

            st.session_state.imported_data = imported
            st.session_state.portfolio_data = None
            st.session_state._import_summary = {
                "n_strats": len(imported.strategy_names),
                "start":    str(imported.date_range[0]),
                "end":      str(imported.date_range[1]),
                "n_days":   len(imported.daily_m2m),
                "n_trades": len(imported.trades),
            }

            progress.progress(1.0, text="Import complete")
            st.rerun()

        except Exception as e:
            st.error(f"Import failed: {e}")
            progress.empty()
            raise

    _summary = st.session_state.pop("_import_summary", None)
    if _summary:
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Strategies",   _summary["n_strats"])
        col2.metric("Trading Days", f"{_summary['n_days']:,}")
        col3.metric("Date Range",   f"{_summary['start']} → {_summary['end']}")
        col4.metric("Trades",       f"{_summary['n_trades']:,}")
        st.success("**Step 2 complete** — data imported. Configure strategies in the next tab or continue to Strategy Tracker.")
        st.page_link("ui/pages/02_Strategies.py", label="Next: Strategy Tracker →")
    elif st.session_state.get("imported_data") and not import_clicked:
        data = st.session_state.imported_data
        start, end = data.date_range
        st.info(
            f"Currently loaded: **{len(data.strategy_names)}** strategies, "
            f"{len(data.daily_m2m):,} trading days ({start} → {end}). "
            f"Click **Import Data** to reload."
        )
        st.page_link("ui/pages/02_Strategies.py", label="Next: Strategy Tracker →")


# ══════════════════════════════════════════════════════════════════
# Tab 2 — Configure Strategies
# ══════════════════════════════════════════════════════════════════

with tab_configure:
    st.caption(
        "Set status, contracts, symbol and sector for each strategy. "
        "Only **Live** strategies are included in the portfolio."
    )

    _cfg_strategies = load_strategies()

    if not _cfg_strategies:
        st.info("No strategies found yet. Scan and import data first.")
        st.stop()

    _CFG_COLUMNS = [
        "name", "status", "contracts", "symbol", "sector",
        "timeframe", "type", "horizon", "other", "notes",
    ]
    _CFG_COLUMN_CONFIG = {
        "name": st.column_config.TextColumn("Strategy", disabled=True, width="large"),
        "status": st.column_config.SelectboxColumn(
            "Status",
            options=[
                "Live", "Paper", "Retired", "Pass",
                "Buy&Hold", "Incubating", "New",
                "Not Loaded - Live", "Not Loaded - Paper",
                "Not Loaded - Retired", "Not Loaded - Pass",
            ],
            required=True, width="medium",
        ),
        "contracts": st.column_config.NumberColumn(
            "Contracts", min_value=0, max_value=999, step=1, format="%d", width="small"
        ),
        "symbol":    st.column_config.TextColumn("Symbol", width="small"),
        "sector":    st.column_config.SelectboxColumn(
            "Sector",
            options=[
                "", "Index", "Energy", "Metals", "Currencies", "Interest Rate",
                "Agriculture", "Soft", "Meats", "Crypto", "Volatility",
                "Eurex Index", "Eurex Interest Rate", "Euronext LIFFE", "Other",
            ],
            width="medium",
        ),
        "timeframe": st.column_config.TextColumn("Timeframe", width="small"),
        "type":      st.column_config.SelectboxColumn(
            "Type",
            options=["", "Trend", "Mean Reversion", "Seasonal", "Arbitrage", "Other"],
            width="medium",
        ),
        "horizon":   st.column_config.SelectboxColumn(
            "Horizon", options=["", "Short", "Medium", "Long"], width="small"
        ),
        "other":     st.column_config.TextColumn("Other", width="small"),
        "notes":     st.column_config.TextColumn("Notes", width="large"),
    }

    # Filters
    with st.expander("Filter / Search", expanded=False):
        _cf1, _cf2, _cf3 = st.columns(3)
        with _cf1:
            _cfg_search = st.text_input("Search name", placeholder="e.g. ES_Trend", key="imp_cfg_search")
        with _cf2:
            _cfg_all_statuses = sorted({s.get("status", "") for s in _cfg_strategies if s.get("status")})
            _cfg_filter_status = st.multiselect("Filter by status", options=_cfg_all_statuses, key="imp_cfg_status")
        with _cf3:
            _cfg_all_sectors = sorted({s.get("sector", "") for s in _cfg_strategies if s.get("sector")})
            _cfg_filter_sector = st.multiselect("Filter by sector", options=_cfg_all_sectors, key="imp_cfg_sector")

    _cfg_filtered = _cfg_strategies
    if _cfg_search:
        _cfg_filtered = [s for s in _cfg_filtered if _cfg_search.lower() in s.get("name", "").lower()]
    if _cfg_filter_status:
        _cfg_filtered = [s for s in _cfg_filtered if s.get("status") in _cfg_filter_status]
    if _cfg_filter_sector:
        _cfg_filtered = [s for s in _cfg_filtered if s.get("sector") in _cfg_filter_sector]

    st.caption(f"Showing **{len(_cfg_filtered)}** of **{len(_cfg_strategies)}** strategies")

    def _cfg_to_df(strats):
        rows = []
        for s in strats:
            row = {col: s.get(col, "") for col in _CFG_COLUMNS}
            try:
                row["contracts"] = int(row["contracts"] or 1)
            except (ValueError, TypeError):
                row["contracts"] = 1
            rows.append(row)
        return pd.DataFrame(rows, columns=_CFG_COLUMNS)

    _cfg_df = _cfg_to_df(_cfg_filtered)
    _cfg_edited = st.data_editor(
        _cfg_df,
        column_config=_CFG_COLUMN_CONFIG,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        key="import_cfg_editor",
    )

    if st.button("Save Changes", type="primary", key="import_cfg_save"):
        _edited_rows = _cfg_edited.to_dict(orient="records")
        _by_name = {r["name"]: r for r in _edited_rows}
        _merged = []
        for _s in _cfg_strategies:
            _nm = _s.get("name", "")
            if _nm in _by_name:
                _updated = dict(_s)
                _updated.update(_by_name[_nm])
                _merged.append(_updated)
            else:
                _merged.append(_s)
        save_strategies(_merged)
        st.session_state.portfolio_data = None
        _live_n = sum(1 for r in _edited_rows if r.get("status") == "Live")
        st.success(
            f"Saved {len(_edited_rows)} records. "
            + (f"{_live_n} Live strategies — portfolio will need a rebuild." if _live_n else "")
        )
        if _live_n:
            st.page_link("ui/pages/03_Portfolio.py", label="→ Build Portfolio")
        st.rerun()

    st.divider()
    st.page_link("ui/pages/02_Strategies.py", label="→ Strategy Tracker (Performance Summary)")
