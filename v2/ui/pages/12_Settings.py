"""
Settings page — license, config backup/restore, and all export options.

Sections:
  1. License        — status + update TradeStation Customer ID
  2. Export / Import — ZIP config backup/restore
  3. Excel Export   — Raw Data | Computed Output | Correlations
  4. PDF Export     — Portfolio summary report
  5. App Preferences — date format, correlation thresholds
"""

from __future__ import annotations

import streamlit as st

from core.config import AppConfig
from core.reporting.settings_io import (
    default_export_filename,
    export_settings,
    import_settings,
)

st.set_page_config(page_title="Settings", layout="wide")
st.title("Settings")

config: AppConfig          = st.session_state.get("config", AppConfig.load())
imported                   = st.session_state.get("imported_data")
portfolio                  = st.session_state.get("portfolio_data")
mc_result                  = st.session_state.get("mc_result")
mc_label                   = st.session_state.get("mc_target_label", "Portfolio")
loo_result                 = st.session_state.get("loo_result")
loo_base_profit            = st.session_state.get("loo_base_profit", 0.0)
loo_base_sharpe            = st.session_state.get("loo_base_sharpe", 0.0)
corr_data                  = st.session_state.get("corr_matrices") or st.session_state.get("correlation_result")


# ── 1. License ────────────────────────────────────────────────────────────────
st.header("License")

from core.licensing.license_manager import is_known_customer, validate_full

current_id = config.customer_id
if current_id:
    col_a, col_b = st.columns([1, 3])
    with col_a:
        st.metric("Current Customer ID", current_id)
    with col_b:
        if is_known_customer(current_id):
            st.success("Customer ID is in the licensed-customer list.")
        else:
            st.error("Customer ID is NOT in the licensed-customer list.")
else:
    st.info("No TradeStation Customer ID configured yet.")

with st.expander("Change / verify Customer ID"):
    with st.form("update_license"):
        new_id = st.number_input(
            "TradeStation Customer ID",
            min_value=1, max_value=9_999_999,
            value=current_id or 1, step=1,
        )
        c1, c2 = st.columns(2)
        save_btn  = c1.form_submit_button("Save", type="primary", use_container_width=True)
        check_btn = c2.form_submit_button("Save & Check DLL", use_container_width=True)
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


# ── 2. Export / Import Config ─────────────────────────────────────────────────
st.header("Export / Import Settings")

col_exp, col_imp = st.columns(2)

with col_exp:
    st.subheader("Export")
    st.caption("Download a ZIP of your strategies, margin tables, and app settings.")
    if st.button("Generate Config ZIP", use_container_width=True):
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
    st.caption("Upload a previously exported ZIP to restore your configuration. **Overwrites current settings.**")
    uploaded = st.file_uploader("Select config ZIP", type=["zip"], label_visibility="collapsed")
    if uploaded is not None:
        if st.button("Restore from ZIP", type="primary", use_container_width=True):
            ok, err, restored = import_settings(uploaded.read())
            if ok:
                st.success(f"Restored: {', '.join(restored)}")
                st.session_state.config = AppConfig.load()
                st.session_state.pop("margin_tables", None)
                st.rerun()
            else:
                st.error(f"Import failed: {err}")

st.divider()


# ── 3. Excel Export ───────────────────────────────────────────────────────────
st.header("Excel Export")

tab_raw, tab_output, tab_mc, tab_loo, tab_corr = st.tabs([
    "Raw Data", "Computed Output", "Monte Carlo", "Leave-One-Out", "Correlations"
])

# ── 3a. Raw Data ──────────────────────────────────────────────────────────────
with tab_raw:
    st.caption(
        "Export the raw imported DataFrames: Daily M2M, Closed Trades, "
        "In-Market Long/Short, and individual trade records."
    )
    if imported is None:
        st.info("No data loaded. Go to **Import** first.")
    else:
        start, end = imported.date_range
        n_strats = len(imported.strategy_names)
        st.metric("Strategies", n_strats)
        st.caption(f"Date range: {start} → {end} | Trades: {len(imported.trades):,}")
        if st.button("Export Raw Data to Excel", use_container_width=True, key="raw_xlsx"):
            try:
                from core.reporting.excel_export import export_raw_data, raw_data_export_filename
                xlsx = export_raw_data(imported)
                st.download_button(
                    "⬇ Download raw_data.xlsx", data=xlsx,
                    file_name=raw_data_export_filename(),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

# ── 3b. Computed Output ───────────────────────────────────────────────────────
with tab_output:
    st.caption(
        "Export computed strategy metrics (80+ columns) plus the "
        "portfolio cumulative equity curve."
    )
    if portfolio is None:
        st.info("No portfolio built. Go to **Portfolio** page first.")
    else:
        st.metric("Active strategies", len(portfolio.strategies))
        if st.button("Export Portfolio Output to Excel", use_container_width=True, key="output_xlsx"):
            try:
                from core.reporting.excel_export import export_portfolio, portfolio_export_filename
                xlsx = export_portfolio(portfolio, config)
                st.download_button(
                    "⬇ Download portfolio_output.xlsx", data=xlsx,
                    file_name=portfolio_export_filename(),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

# ── 3c. Monte Carlo ───────────────────────────────────────────────────────────
with tab_mc:
    st.caption("Export MC summary metrics and scenario distribution.")
    if mc_result is None:
        st.info("No MC results cached. Run **Monte Carlo** page first.")
    else:
        c1, c2, c3 = st.columns(3)
        c1.metric("Starting Equity",    f"${mc_result.starting_equity:,.0f}")
        c2.metric("Expected Profit",    f"${mc_result.expected_profit:,.0f}")
        c3.metric("Risk of Ruin",       f"{mc_result.risk_of_ruin:.1%}")
        if st.button("Export MC Results to Excel", use_container_width=True, key="mc_xlsx"):
            try:
                from core.reporting.excel_export import export_mc_result, mc_export_filename
                xlsx = export_mc_result(mc_result, label=mc_label)
                st.download_button(
                    "⬇ Download mc_results.xlsx", data=xlsx,
                    file_name=mc_export_filename(mc_label),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

# ── 3d. Leave-One-Out ─────────────────────────────────────────────────────────
with tab_loo:
    st.caption("Export the Leave-One-Out analysis table.")
    if loo_result is None:
        st.info("No LOO results cached. Run **Leave One Out** page first.")
    else:
        st.metric("Strategies analysed", len(loo_result))
        if st.button("Export LOO Results to Excel", use_container_width=True, key="loo_xlsx"):
            try:
                from core.reporting.excel_export import export_loo_result, loo_export_filename
                xlsx = export_loo_result(loo_result, loo_base_profit, loo_base_sharpe)
                st.download_button(
                    "⬇ Download leave_one_out.xlsx", data=xlsx,
                    file_name=loo_export_filename(),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

# ── 3e. Correlations ─────────────────────────────────────────────────────────
with tab_corr:
    st.caption("Export a correlation matrix to Excel.")
    if corr_data is None:
        st.info("No correlation results cached. Run **Correlations** page first.")
    else:
        if isinstance(corr_data, dict):
            mode = st.selectbox("Mode", list(corr_data.keys()), key="corr_mode_sel")
            corr_df = corr_data[mode]
        else:
            mode = "Normal"
            corr_df = corr_data
        st.caption(f"{len(corr_df)} × {len(corr_df.columns)} matrix")
        if st.button("Export Correlations to Excel", use_container_width=True, key="corr_xlsx"):
            try:
                from core.reporting.excel_export import export_correlations, correlations_export_filename
                xlsx = export_correlations(corr_df, mode)
                st.download_button(
                    "⬇ Download correlations.xlsx", data=xlsx,
                    file_name=correlations_export_filename(mode),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except ImportError as e:
                st.error(str(e))

st.divider()


# ── 4. PDF Export ─────────────────────────────────────────────────────────────
st.header("PDF Export")

st.caption(
    "Export a portfolio summary report as a PDF — cover page, overview metrics, "
    "MC summary (if available), and strategy list."
)

if portfolio is None:
    st.info("No portfolio built. Go to **Portfolio** page first.")
else:
    n_active = len(portfolio.strategies)
    col1, col2 = st.columns(2)
    col1.metric("Active strategies", n_active)
    if mc_result is not None:
        col2.metric("MC results included", "Yes")

    if st.button("Export Portfolio Report to PDF", type="primary", use_container_width=True):
        try:
            from core.reporting.pdf_export import export_portfolio_pdf, pdf_export_filename
            with st.spinner("Generating PDF…"):
                pdf_bytes = export_portfolio_pdf(portfolio, config, mc_result=mc_result)
            st.download_button(
                "⬇ Download portfolio_report.pdf", data=pdf_bytes,
                file_name=pdf_export_filename(),
                mime="application/pdf",
                use_container_width=True,
            )
        except ImportError as e:
            st.error(str(e))

st.divider()


# ── 5. App Preferences ────────────────────────────────────────────────────────
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
        value=config.corr_normal_threshold, format="%.2f",
    )
    corr_drawdown = st.number_input(
        "Correlation threshold — Drawdown mode",
        min_value=0.0, max_value=1.0, step=0.05,
        value=config.corr_drawdown_threshold, format="%.2f",
    )
    corr_negative = st.number_input(
        "Correlation threshold — Negative mode",
        min_value=0.0, max_value=1.0, step=0.05,
        value=config.corr_negative_threshold, format="%.2f",
        help="Pairs BELOW this are considered negatively correlated (diversifying).",
    )
    if st.form_submit_button("Save Preferences", type="primary"):
        config.date_format            = date_fmt
        config.corr_normal_threshold  = corr_normal
        config.corr_drawdown_threshold = corr_drawdown
        config.corr_negative_threshold = corr_negative
        config.save()
        st.session_state.config = config
        st.success("Preferences saved.")
