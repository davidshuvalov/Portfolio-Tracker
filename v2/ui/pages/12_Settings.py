"""
Settings page — license management, configuration backup/restore, Excel export.

Sections:
  1. License — show status, update TradeStation Customer ID
  2. Export / Import — ZIP backup and restore of all config files
  3. Excel Export — download portfolio summary or correlation matrix as .xlsx
  4. App Preferences — date format and other AppConfig settings
"""

from __future__ import annotations

import streamlit as st
import pandas as pd

from core.config import AppConfig
from core.reporting.settings_io import (
    default_export_filename,
    export_settings,
    import_settings,
)

st.set_page_config(page_title="Settings", layout="wide")
st.title("Settings")

config: AppConfig = st.session_state.get("config", AppConfig.load())
portfolio = st.session_state.get("portfolio_data")


# ── 1. License ────────────────────────────────────────────────────────────────
st.header("License")

from core.licensing.license_manager import is_known_customer, validate_full

current_id = config.customer_id
if current_id:
    col_a, col_b = st.columns([1, 3])
    with col_a:
        st.metric("Current Customer ID", current_id)
    with col_b:
        known = is_known_customer(current_id)
        if known:
            st.success("Customer ID is in the licensed-customer list.")
        else:
            st.error("Customer ID is NOT in the licensed-customer list.")
else:
    st.info("No TradeStation Customer ID configured yet.")

with st.expander("Change / verify Customer ID"):
    with st.form("update_license"):
        new_id = st.number_input(
            "TradeStation Customer ID",
            min_value=1,
            max_value=9_999_999,
            value=current_id or 1,
            step=1,
        )
        col1, col2 = st.columns(2)
        save_btn  = col1.form_submit_button("Save", type="primary", use_container_width=True)
        check_btn = col2.form_submit_button("Save & Check DLL", use_container_width=True)

        if save_btn or check_btn:
            config.customer_id = int(new_id)
            config.save()
            st.session_state.config = config
            if check_btn:
                with st.spinner("Checking license via MultiWalk DLL…"):
                    valid, msg = validate_full(int(new_id))
                if valid:
                    st.success("License validated successfully.")
                else:
                    st.error(f"License check failed: {msg}")
            else:
                st.success("Customer ID saved.")
            st.rerun()

st.divider()


# ── 2. Export / Import Settings ───────────────────────────────────────────────
st.header("Export / Import Settings")

col_exp, col_imp = st.columns(2)

with col_exp:
    st.subheader("Export")
    st.caption(
        "Downloads a ZIP archive containing your strategies, margins, and app settings. "
        "Use this to back up your configuration or move it to another machine."
    )
    if st.button("Generate Export", use_container_width=True):
        zip_bytes = export_settings()
        st.download_button(
            label="⬇ Download config ZIP",
            data=zip_bytes,
            file_name=default_export_filename(),
            mime="application/zip",
            use_container_width=True,
        )

with col_imp:
    st.subheader("Import")
    st.caption(
        "Upload a previously exported ZIP to restore your configuration. "
        "**This overwrites your current settings.**"
    )
    uploaded = st.file_uploader(
        "Select config ZIP",
        type=["zip"],
        label_visibility="collapsed",
    )
    if uploaded is not None:
        if st.button("Restore from ZIP", type="primary", use_container_width=True):
            ok, err, restored = import_settings(uploaded.read())
            if ok:
                st.success(f"Restored: {', '.join(restored)}")
                # Reload config from disk
                st.session_state.config = AppConfig.load()
                # Clear cached margin tables so they reload
                st.session_state.pop("margin_tables", None)
                st.rerun()
            else:
                st.error(f"Import failed: {err}")

st.divider()


# ── 3. Excel Export ───────────────────────────────────────────────────────────
st.header("Excel Export")

tab_portfolio, tab_correlations = st.tabs(["Portfolio Summary", "Correlations"])

with tab_portfolio:
    if portfolio is None:
        st.info("No portfolio loaded. Go to **Portfolio** page first.")
    else:
        st.caption(
            f"Export strategy metrics for **{len(portfolio.strategies)}** strategies "
            "to an Excel workbook (Summary + Portfolio Equity sheets)."
        )
        if st.button("Export Portfolio to Excel", use_container_width=True):
            try:
                from core.reporting.excel_export import (
                    export_portfolio,
                    portfolio_export_filename,
                )
                xlsx_bytes = export_portfolio(portfolio, config)
                st.download_button(
                    label="⬇ Download portfolio.xlsx",
                    data=xlsx_bytes,
                    file_name=portfolio_export_filename(),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

with tab_correlations:
    corr_data = st.session_state.get("correlation_result")
    if corr_data is None:
        st.info("No correlation results cached. Run the **Correlations** page first.")
    else:
        # corr_data may be a DataFrame or a dict keyed by mode
        if isinstance(corr_data, dict):
            mode = st.selectbox("Mode", list(corr_data.keys()))
            corr_df = corr_data[mode]
        else:
            mode = "Normal"
            corr_df = corr_data

        st.caption(f"Export **{mode}** correlation matrix ({len(corr_df)} × {len(corr_df.columns)}) to Excel.")
        if st.button("Export Correlations to Excel", use_container_width=True):
            try:
                from core.reporting.excel_export import (
                    export_correlations,
                    correlations_export_filename,
                )
                xlsx_bytes = export_correlations(corr_df, mode)
                st.download_button(
                    label="⬇ Download correlations.xlsx",
                    data=xlsx_bytes,
                    file_name=correlations_export_filename(mode),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

st.divider()


# ── 4. App Preferences ────────────────────────────────────────────────────────
st.header("App Preferences")

with st.form("preferences_form"):
    date_fmt = st.selectbox(
        "Date format in CSV files",
        ["DMY", "MDY"],
        index=["DMY", "MDY"].index(config.date_format),
        help="DMY = DD/MM/YYYY (TradeStation default). MDY = MM/DD/YYYY.",
    )
    corr_normal = st.number_input(
        "Correlation threshold — Normal mode",
        min_value=0.0, max_value=1.0, step=0.05,
        value=config.corr_normal_threshold,
        format="%.2f",
        help="Pairs above this are highlighted as highly correlated.",
    )
    corr_drawdown = st.number_input(
        "Correlation threshold — Drawdown mode",
        min_value=0.0, max_value=1.0, step=0.05,
        value=config.corr_drawdown_threshold,
        format="%.2f",
    )
    corr_negative = st.number_input(
        "Correlation threshold — Negative mode",
        min_value=0.0, max_value=1.0, step=0.05,
        value=config.corr_negative_threshold,
        format="%.2f",
        help="Pairs BELOW this are considered negatively correlated (diversifying).",
    )

    if st.form_submit_button("Save Preferences", type="primary"):
        config.date_format = date_fmt
        config.corr_normal_threshold   = corr_normal
        config.corr_drawdown_threshold = corr_drawdown
        config.corr_negative_threshold = corr_negative
        config.save()
        st.session_state.config = config
        st.success("Preferences saved.")
