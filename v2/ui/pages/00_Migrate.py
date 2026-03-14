"""
One-time migration wizard from Portfolio Tracker v1.24 (.xlsb).
Shown automatically on first run if no strategies are configured yet.
"""

import streamlit as st
import pandas as pd
from pathlib import Path

from core.ingestion.xlsb_importer import (
    import_strategies_from_xlsb,
    import_margin_tables,
    save_margin_tables,
    PYXLSB_AVAILABLE,
)
from core.config import AppConfig

st.set_page_config(page_title="Migrate from v1.24", layout="wide")
st.title("Migrate from Portfolio Tracker v1.24")

config: AppConfig = st.session_state.get("config", AppConfig.load())

st.markdown("""
This one-time import reads your existing **v1.24 `.xlsb` file** and copies key
configuration into Portfolio Tracker v2.

**What gets imported:**
- Strategy name, status, contracts, symbol, timeframe, type, horizon, notes
- **Margin reference tables** (TradeStation Margins, IB Margins, Symbol Lookup)
  — used by the Margin Tracking page to auto-populate margin requirements

**What does NOT get imported:** All computed metrics — these are recalculated
from your MultiWalk CSV files when you run an import.

Once complete, you won't need to use this page again.
""")

if not PYXLSB_AVAILABLE:
    st.error(
        "`pyxlsb` is not installed. Run `pip install pyxlsb` then restart the app."
    )
    st.stop()

st.divider()

uploaded = st.file_uploader(
    "Select your Portfolio Tracker v1.24 .xlsb file",
    type=["xlsb"],
    help="This is the file named something like 'Portfolio Tracker - A Tool for MultiWalk v1.24.xlsb'",
)

if uploaded is None:
    st.info("Upload your .xlsb file above to begin.")
    st.stop()

# Write uploaded file to a temp path so pyxlsb can open it
import tempfile, os

with tempfile.NamedTemporaryFile(suffix=".xlsb", delete=False) as tmp:
    tmp.write(uploaded.read())
    tmp_path = Path(tmp.name)

try:
    with st.spinner("Reading Strategies tab and margin tables from workbook…"):
        strategies, warnings = import_strategies_from_xlsb(tmp_path)
        try:
            margin_tables = import_margin_tables(tmp_path)
        except Exception as e:
            margin_tables = None
            warnings.append(f"Could not read margin tables: {e}")
finally:
    os.unlink(tmp_path)

if warnings:
    with st.expander(f"{len(warnings)} warning(s)"):
        for w in warnings:
            st.warning(w)

if not strategies:
    st.error("No strategies found in the Strategies tab. Nothing to import.")
    st.stop()

# Auto-fill sectors from Sector sheet lookup before showing preview
n_filled = 0
if margin_tables is not None and margin_tables.sector_lookup:
    for strat in strategies:
        sym = strat.get("symbol", "")
        if sym and not strat.get("sector") and sym in margin_tables.sector_lookup:
            strat["sector"] = margin_tables.sector_lookup[sym]
            n_filled += 1

col_s, col_m = st.columns(2)
col_s.success(f"Found **{len(strategies)}** strategies.")
if margin_tables is not None:
    sector_note = f" · **{n_filled}** sectors auto-filled" if n_filled else ""
    col_m.success(
        f"Reference tables: **{len(margin_tables.ts)}** TS margins, "
        f"**{len(margin_tables.ib)}** IB margins, "
        f"**{len(margin_tables.sector_lookup)}** sector mappings"
        + sector_note + "."
    )
else:
    col_m.warning("Reference tables could not be read (see warnings above).")

# Preview table
df = pd.DataFrame(strategies)
display_cols = ["name", "status", "contracts", "symbol", "sector",
                "timeframe", "type", "horizon", "notes"]
df_display = df[[c for c in display_cols if c in df.columns]]

st.subheader("Preview — strategies to be imported")
st.dataframe(df_display, use_container_width=True, height=400)

st.caption(
    "The **sector** column is blank — it's not stored in v1.24. "
    "You can fill it in on the Strategies page after importing."
)

st.divider()
col1, col2 = st.columns([1, 5])
with col1:
    confirm = st.button("Confirm & Import", type="primary", use_container_width=True)
with col2:
    st.caption("This will save the strategies above as your v2 configuration.")

if confirm:
    from core.portfolio.strategies import save_strategies
    save_strategies(strategies)
    st.session_state.config = config
    if margin_tables is not None:
        save_margin_tables(margin_tables)
        st.session_state["margin_tables"] = margin_tables
    st.success(
        f"Imported {len(strategies)} strategies"
        + (f" and margin tables for {len(margin_tables.ts)} symbols" if margin_tables else "")
        + ". Go to the **Strategies** page to review and add sector information, "
        "then go to **Import** to load your MultiWalk data."
    )
    st.balloons()
