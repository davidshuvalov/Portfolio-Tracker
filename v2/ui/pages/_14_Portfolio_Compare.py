"""Portfolio Compare — mirrors R_Check_New_Strategies.bas.

Replaces the VBA "pick a reference .xlsb file" workflow with a lightweight
snapshot system:

  1. Save Snapshot  — captures the current Live portfolio (name + contracts)
                      to ~/.portfolio_tracker/snapshots/<timestamp>_<label>.yaml
  2. Compare        — loads any saved snapshot and shows what changed:
                        • New strategies   (green)
                        • Removed          (red)
                        • Contract changes (orange)
                        • Unchanged        (grey / no highlight)

Snapshots are independent of the 'status' column so historical comparisons
are unaffected by later status edits.
"""

from __future__ import annotations

from datetime import datetime

import pandas as pd
import streamlit as st

from core.config import AppConfig
from core.data_types import PortfolioData
from core.portfolio.snapshot import (
    CompareResult,
    compare_portfolios,
    delete_snapshot,
    list_snapshots,
    load_snapshot,
    save_snapshot,
)
from core.portfolio.strategies import load_strategies

st.set_page_config(page_title="Portfolio Compare", layout="wide")

# ── Sidebar workflow status ────────────────────────────────────────────────────
try:
    from ui.workflow import render_workflow_sidebar
    with st.sidebar:
        render_workflow_sidebar()
except Exception:
    pass

# ── Top navigation ─────────────────────────────────────────────────────────────
_nav_l, _ = st.columns([1, 7])
with _nav_l:
    st.page_link("ui/pages/03_Portfolio.py", label="← Portfolio")

st.title("Portfolio Compare")
st.caption(
    "Save a snapshot of the current Live portfolio, then compare any future "
    "state against it to see what changed."
)

config: AppConfig = st.session_state.get("config", AppConfig.load())
portfolio: PortfolioData | None = st.session_state.get("portfolio_data")
strategies_list = load_strategies()

live_strategies = [s for s in strategies_list if s.get("status") == "Live"]

tab_save, tab_compare, tab_manage = st.tabs(
    ["💾 Save Snapshot", "🔍 Compare", "🗂 Manage Snapshots"]
)

# ── Tab 1: Save Snapshot ───────────────────────────────────────────────────────
with tab_save:
    st.subheader("Save current Live portfolio as a snapshot")

    if not live_strategies:
        st.warning("No Live strategies found. Load data and set strategies to Live first.")
    else:
        today_label = datetime.now().strftime("%Y-%m-%d")
        label = st.text_input(
            "Snapshot label",
            value=f"Pre-{today_label}",
            help="A description to identify this snapshot later (e.g. 'Pre-March-2026', 'After rebal').",
        )

        st.write(f"**{len(live_strategies)} Live strategies** will be saved:")
        preview_df = pd.DataFrame(
            [
                {
                    "Strategy": s.get("name", ""),
                    "Symbol": s.get("symbol", ""),
                    "Sector": s.get("sector", ""),
                    "Contracts": s.get("contracts", 1),
                }
                for s in live_strategies
            ]
        )
        st.dataframe(preview_df, hide_index=True, use_container_width=True, height=min(400, len(live_strategies) * 36 + 40))

        if st.button("💾 Save Snapshot", type="primary"):
            if not label.strip():
                st.error("Please enter a label.")
            else:
                path = save_snapshot(strategies_list, label.strip())
                st.success(f"Snapshot saved: `{path.name}`")
                st.rerun()

# ── Tab 2: Compare ─────────────────────────────────────────────────────────────
with tab_compare:
    snapshots = list_snapshots()

    if not snapshots:
        st.info("No snapshots saved yet. Go to the **Save Snapshot** tab to create one.")
    else:
        snap_options = {
            f"{s['label']}  ({s['saved_at'][:10]}, {s['n_strategies']} strategies)": s["filename"]
            for s in snapshots
        }
        selected_label = st.selectbox(
            "Compare current Live portfolio against",
            list(snap_options.keys()),
        )
        selected_filename = snap_options[selected_label]

        if st.button("🔍 Compare", type="primary"):
            reference = load_snapshot(selected_filename)
            result: CompareResult = compare_portfolios(
                strategies_list,
                reference,
                live_status=config.portfolio.live_status,
            )
            st.session_state["_compare_result"] = result
            st.session_state["_compare_label"] = selected_label

        result: CompareResult | None = st.session_state.get("_compare_result")
        if result is not None:
            label_used = st.session_state.get("_compare_label", "")
            st.subheader(f"Changes vs: {label_used}")

            # Summary cards
            cols = st.columns(4)
            cols[0].metric("New", len(result.new_strategies), delta=f"+{len(result.new_strategies)}" if result.new_strategies else "0")
            cols[1].metric("Removed", len(result.removed_strategies), delta=f"-{len(result.removed_strategies)}" if result.removed_strategies else "0", delta_color="inverse")
            cols[2].metric("Contract changes", len(result.contract_changes))
            cols[3].metric("Unchanged", len(result.unchanged))

            if not result.has_changes:
                st.success("No changes — current Live portfolio matches the snapshot exactly.")
            else:
                # New strategies
                if result.new_strategies:
                    st.markdown("#### 🟢 New strategies (not in snapshot)")
                    new_df = pd.DataFrame(
                        [{"Strategy": s.get("name", ""), "Symbol": s.get("symbol", ""), "Sector": s.get("sector", ""), "Contracts": s.get("contracts", 1)} for s in result.new_strategies]
                    )
                    st.dataframe(
                        new_df.style.applymap(lambda _: "background-color: #C8E6C9", subset=new_df.columns),
                        hide_index=True, use_container_width=True,
                    )

                # Removed strategies
                if result.removed_strategies:
                    st.markdown("#### 🔴 Removed strategies (were Live in snapshot)")
                    rem_df = pd.DataFrame(
                        [{"Strategy": s.get("name", ""), "Symbol": s.get("symbol", ""), "Sector": s.get("sector", ""), "Contracts": s.get("contracts", 1)} for s in result.removed_strategies]
                    )
                    st.dataframe(
                        rem_df.style.applymap(lambda _: "background-color: #FFCDD2", subset=rem_df.columns),
                        hide_index=True, use_container_width=True,
                    )

                # Contract changes
                if result.contract_changes:
                    st.markdown("#### 🟠 Contract changes")
                    chg_df = pd.DataFrame(result.contract_changes)[
                        ["name", "symbol", "old_contracts", "new_contracts", "delta"]
                    ]
                    chg_df.columns = ["Strategy", "Symbol", "Was", "Now", "Δ"]

                    def _chg_style(row):
                        color = "#FFE0B2" if row["Δ"] != 0 else ""
                        return [f"background-color: {color}"] * len(row)

                    st.dataframe(
                        chg_df.style.apply(_chg_style, axis=1),
                        hide_index=True, use_container_width=True,
                    )

                # Unchanged (collapsed)
                if result.unchanged:
                    with st.expander(f"Unchanged ({len(result.unchanged)} strategies)", expanded=False):
                        unch_df = pd.DataFrame(
                            [{"Strategy": s.get("name", ""), "Symbol": s.get("symbol", ""), "Contracts": s.get("contracts", 1)} for s in result.unchanged]
                        )
                        st.dataframe(unch_df, hide_index=True, use_container_width=True)

# ── Tab 3: Manage Snapshots ────────────────────────────────────────────────────
with tab_manage:
    snapshots = list_snapshots()

    if not snapshots:
        st.info("No snapshots yet.")
    else:
        st.write(f"**{len(snapshots)} saved snapshots** in `~/.portfolio_tracker/snapshots/`")

        for snap in snapshots:
            col_info, col_del = st.columns([6, 1])
            col_info.markdown(
                f"**{snap['label']}** — {snap['saved_at'][:16].replace('T', ' ')} "
                f"· {snap['n_strategies']} strategies  \n"
                f"`{snap['filename']}`"
            )
            if col_del.button("🗑", key=f"del_{snap['filename']}", help="Delete this snapshot"):
                delete_snapshot(snap["filename"])
                # Clear compare result if it was from this snapshot
                if st.session_state.get("_compare_result") is not None:
                    del st.session_state["_compare_result"]
                st.rerun()
