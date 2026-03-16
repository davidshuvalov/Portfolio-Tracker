"""
Import page — scan MultiWalk folders and load strategy CSV data.
Mirrors the VBA 'Retrieve Folder Data' + 'Import Data' workflow.
"""

import streamlit as st
import pandas as pd
from pathlib import Path

from core.config import AppConfig
from core.ingestion.folder_scanner import scan_folders, reconcile_statuses
from core.ingestion.csv_importer import import_all
from core.portfolio.strategies import load_strategies, save_strategies

st.set_page_config(page_title="Import", layout="wide")

st.title("Import")
st.caption("Steps 1 & 2 of 4 — add folders, then scan and load strategy CSV data.")

config: AppConfig = st.session_state.get("config", AppConfig.load())


# ═══════════════════════════════════════════════════════════════
# Section 1 — Base Folder Configuration
# ═══════════════════════════════════════════════════════════════

st.subheader("MultiWalk Base Folders")
st.caption(
    "Add the folders that contain your MultiWalk strategy subfolders. "
    "Equivalent to Folder1–Folder10 and FolderBH in v1.24."
)

_STATUS_OPTIONS = ["New", "Live", "Paper", "Pass", "Retired", "Buy&Hold"]

# Display current folders
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

# Add new folder
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

# ── Step 1 completion callout ──────────────────────────────────
if config.folders:
    st.success("**Step 1 complete** — folders configured. Now scan and import data below.")

st.divider()


# ═══════════════════════════════════════════════════════════════
# Section 2 — Cutoff Date Settings
# ═══════════════════════════════════════════════════════════════

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
        import datetime
        cutoff_date = st.date_input(
            "Cutoff Date",
            value=pd.Timestamp(config.portfolio.cutoff_date).date()
            if config.portfolio.cutoff_date
            else datetime.date.today(),
        )

# Save settings if changed
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


# ═══════════════════════════════════════════════════════════════
# Section 3 — Scan & Import
# ═══════════════════════════════════════════════════════════════

st.subheader("Scan & Import")

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

# ── Scan ──────────────────────────────────────────────────────
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

        # Reconcile with configured strategies (uses folder default status)
        configured = load_strategies()
        found_names = {sf.name for sf in result.strategies}
        reconciled = reconcile_statuses(
            found_names,
            configured,
            strategy_folders=result.strategies,
            folder_default_status=config.folder_default_status,
        )
        save_strategies(reconciled)

        # Show scan results table
        # Quick-peek multi-strategy detection for display only
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
    else:
        st.warning("No strategies found. Check your folder paths.")

# ── Import ────────────────────────────────────────────────────
if import_clicked:
    # Auto-scan if no cached scan result yet
    if st.session_state.get("scan_result") is None:
        with st.spinner("Scanning folders…"):
            _auto = scan_folders(config.folders)
        st.session_state.scan_result = _auto
    result = st.session_state.scan_result

    cutoff = None
    if use_cutoff and cutoff_date:
        import datetime
        cutoff = cutoff_date if isinstance(cutoff_date, datetime.date) else None

    n_strats_to_import = len(result.strategies)
    progress = st.progress(0, text=f"Importing 0 / {n_strats_to_import} strategies…")
    status_text = st.empty()

    def _on_progress(idx: int, total: int, name: str) -> None:
        frac = idx / max(total, 1)
        progress.progress(frac, text=f"Reading {name} ({idx}/{total})…")

    try:
        imported, warnings = import_all(
            result.strategies,
            date_format=config.date_format,
            use_cutoff=use_cutoff,
            cutoff_date=cutoff,
            progress_cb=_on_progress,
        )

        if warnings:
            with st.expander(f"{len(warnings)} import warning(s)"):
                for w in warnings:
                    st.warning(w)

        # Merge configured metadata (symbol, sector, timeframe, etc.) into
        # the stub Strategy objects created by import_all
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

        # Post-import reconciliation: actual strategy names may differ from scan names
        # (e.g. multi-strategy files produce sub-strategy names not known at scan time)
        _actual_names = set(imported.strategy_names)
        _post_configured = load_strategies()
        _scan_result = st.session_state.get("scan_result")
        _sf_list = _scan_result.strategies if _scan_result else []
        _post_reconciled = reconcile_statuses(
            _actual_names,
            _post_configured,
            strategy_folders=_sf_list,
            folder_default_status=config.folder_default_status,
        )
        save_strategies(_post_reconciled)

        st.session_state.imported_data = imported
        st.session_state.portfolio_data = None  # Reset downstream cache

        progress.progress(1.0, text="Import complete")

        # Summary stats
        n_strats = len(imported.strategy_names)
        start, end = imported.date_range
        n_days = len(imported.daily_m2m)
        n_trades = len(imported.trades)

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Strategies", n_strats)
        col2.metric("Trading Days", f"{n_days:,}")
        col3.metric("Date Range", f"{start} → {end}")
        col4.metric("Trades", f"{n_trades:,}")

        st.success("**Step 2 complete** — data imported successfully.")
        st.page_link("ui/pages/02_Strategies.py", label="Next: Review Strategies →")

    except Exception as e:
        st.error(f"Import failed: {e}")
        progress.empty()
        raise

# ── Show currently loaded data (if any) ──────────────────────
if st.session_state.get("imported_data") and not import_clicked:
    data = st.session_state.imported_data
    start, end = data.date_range
    st.info(
        f"Currently loaded: **{len(data.strategy_names)}** strategies, "
        f"{len(data.daily_m2m):,} trading days ({start} → {end}). "
        f"Click **Import Data** to reload."
    )
    st.page_link("ui/pages/02_Strategies.py", label="Next: Review Strategies →")
