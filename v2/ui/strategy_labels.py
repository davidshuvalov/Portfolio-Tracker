"""
Shared utility: short strategy labels for use on chart axes and in tables.

Strategy names are often long (e.g. "ES_Trend_H4_v3_WF2024").  When used
as axis tick labels on heatmaps or bar charts they become unreadable.

This module provides:
  - build_label_map()  → { full_name: "S01 ×2" }
  - build_label_df()   → DataFrame legend table (# / Code / Name / Symbol / Contracts)
  - render_legend()    → renders the legend table inside a Streamlit expander
  - relabel_matrix()   → renames both axes of a square DataFrame
  - relabel_series()   → renames the index of a Series
"""

from __future__ import annotations

import pandas as pd
import streamlit as st

from core.data_types import Strategy


def build_label_map(strategies: list[Strategy]) -> dict[str, str]:
    """
    Return {strategy_name: short_code} for every strategy.

    Short code format:  S{n} {symbol}   (symbol omitted when not set)
      n = 1-based position

    Example: "ES_Trend_H4_v3" trading TU → "S3 TU"
    """
    return {
        s.name: f"S{i + 1} {s.symbol}" if s.symbol else f"S{i + 1}"
        for i, s in enumerate(strategies)
    }


def build_label_df(strategies: list[Strategy]) -> pd.DataFrame:
    """
    Return a DataFrame suitable for use as a legend table.
    Columns: #, Code, Strategy, Symbol, Contracts
    """
    rows = []
    for i, s in enumerate(strategies):
        code = f"S{i + 1} {s.symbol}" if s.symbol else f"S{i + 1}"
        rows.append({
            "#":          i + 1,
            "Code":       code,
            "Strategy":   s.name,
            "Symbol":     s.symbol or "—",
            "Contracts":  s.contracts if s.contracts else 1,
        })
    return pd.DataFrame(rows)


def render_legend(strategies: list[Strategy], *, expanded: bool = False) -> None:
    """
    Render a collapsible legend table mapping short codes → strategy names.
    Call this after any chart that uses short labels.
    """
    df = build_label_df(strategies)
    with st.expander("Strategy Legend", expanded=expanded):
        st.dataframe(df, hide_index=True, use_container_width=True)


def relabel_matrix(matrix: pd.DataFrame, label_map: dict[str, str]) -> pd.DataFrame:
    """
    Rename both index and columns of a square correlation/overlap DataFrame
    using the provided label_map.  Unknown names are left unchanged.
    """
    mapper = lambda name: label_map.get(name, name)
    return matrix.rename(index=mapper, columns=mapper)


def relabel_series(series: pd.Series, label_map: dict[str, str]) -> pd.Series:
    """Rename the index of a Series using label_map."""
    return series.rename(index=lambda name: label_map.get(name, name))


def render_strategy_picker(
    strategies: list[Strategy],
    *,
    label: str = "Open strategy detail",
    key: str = "strategy_picker",
) -> None:
    """
    Render a compact strategy selector + link inside a sidebar section.

    Sets ``st.session_state.selected_strategy`` and shows a page_link to
    the Strategy Detail page so the user can navigate there with one click.

    Usage (inside a ``with st.sidebar:`` block)::

        render_strategy_picker(portfolio.strategies)
    """
    if not strategies:
        return

    names = [s.name for s in strategies]
    current = st.session_state.get("selected_strategy")
    default_idx = names.index(current) if current in names else 0

    chosen = st.selectbox(label, names, index=default_idx, key=key)
    st.session_state.selected_strategy = chosen
    st.page_link(
        "ui/pages/_Strategy_Detail.py",
        label="Open detail →",
    )
