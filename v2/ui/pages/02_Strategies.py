"""
Strategies page — editable strategy configuration table.
Mirrors the VBA Strategies tab: user sets status, contracts, symbol, sector, etc.
Uses st.data_editor for inline editing with immediate persistence to strategies.yaml.
"""

import streamlit as st
import pandas as pd
from pathlib import Path

from core.portfolio.strategies import load_strategies, save_strategies

st.set_page_config(page_title="Strategies", layout="wide")
st.title("Strategies")

st.caption(
    "Configure strategy status, contracts, and metadata. "
    "Changes are saved automatically. Run an **Import** first to populate new strategies."
)


# ── Load strategies ────────────────────────────────────────────────────────────

strategies = load_strategies()

if not strategies:
    st.info(
        "No strategies configured yet. "
        "Go to **Import** → Scan Folders to discover your MultiWalk strategies."
    )
    st.stop()


# ── Build editable DataFrame ───────────────────────────────────────────────────

# Column display order and types for st.data_editor
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
        # Ensure contracts is numeric
        try:
            row["contracts"] = int(row["contracts"] or 1)
        except (ValueError, TypeError):
            row["contracts"] = 1
        rows.append(row)
    return pd.DataFrame(rows, columns=_COLUMNS)


# ── Filters ────────────────────────────────────────────────────────────────────

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


# ── Editable table ─────────────────────────────────────────────────────────────

df = _to_df(filtered)

edited_df = st.data_editor(
    df,
    column_config=_COLUMN_CONFIG,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    key="strategies_editor",
)

# ── Save changes ───────────────────────────────────────────────────────────────

if st.button("Save Changes", type="primary"):
    edited_rows = edited_df.to_dict(orient="records")

    # Merge edits back into the full strategies list (preserving unfiltered rows)
    edited_by_name = {r["name"]: r for r in edited_rows}

    merged = []
    for s in strategies:
        name = s.get("name", "")
        if name in edited_by_name:
            # Use edited values, preserve any extra keys not in the editor
            updated = dict(s)
            updated.update(edited_by_name[name])
            merged.append(updated)
        else:
            merged.append(s)

    save_strategies(merged)
    st.success(f"Saved {len(edited_rows)} strategy records.")
    st.rerun()


# ── Quick stats sidebar ────────────────────────────────────────────────────────

with st.sidebar:
    st.header("Summary")

    status_counts = {}
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

    # Status breakdown
    st.caption("Status breakdown")
    for status, count in sorted(status_counts.items()):
        pct = count / total * 100 if total else 0
        st.write(f"**{status}**: {count} ({pct:.0f}%)")


# ── Bulk operations ────────────────────────────────────────────────────────────

st.divider()
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

# ── Auto-fill sectors from v1.24 reference data ────────────────────────────────

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
