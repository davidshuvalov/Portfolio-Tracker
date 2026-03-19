"""
Plan enforcement helpers.

Usage in any page:

    from ui.plan_gate import enforce_strategy_limit, gate_full_plan

    # At the top of a page that handles strategies:
    enforce_strategy_limit()

    # To gate a Full-only feature:
    if not gate_full_plan("Portfolio Optimizer"):
        st.stop()
"""

from __future__ import annotations

import streamlit as st

from auth.session import (
    get_plan,
    get_strategy_limit,
    is_subscribed,
    within_strategy_limit,
)


# ── Upgrade prompt ────────────────────────────────────────────────────────────
def render_upgrade_prompt(reason: str = "") -> None:
    """Show a visually distinct upgrade call-to-action."""
    from ui.pricing import render_pricing_cards  # local import to avoid circular

    st.markdown("---")
    st.warning(
        f"**Plan limit reached.** {reason}"
        if reason
        else "**Plan limit reached.** Upgrade to continue."
    )

    with st.expander("View plans", expanded=True):
        render_pricing_cards(show_actions=True)


# ── Strategy count enforcement ────────────────────────────────────────────────
def enforce_strategy_limit() -> None:
    """
    Check the current strategy count against the active plan limit.
    If the user is over their limit, render an upgrade prompt and stop the page.

    Call this at the top of Import, Strategies, and Portfolio pages.
    """
    imported = st.session_state.get("imported_data")
    if imported is None:
        return  # Nothing loaded yet — nothing to gate

    try:
        count = len(imported.strategy_names)
    except AttributeError:
        return

    limit = get_strategy_limit()
    if limit is None:  # unlimited
        return

    if count > limit:
        render_upgrade_prompt(
            f"You are tracking **{count} strategies** but your "
            f"**{get_plan() or 'current'} plan** allows up to **{limit}**. "
            "Upgrade to Full for unlimited strategies."
        )
        st.stop()


def check_strategy_count_before_import(new_count: int) -> bool:
    """
    Returns True if importing `new_count` strategies is allowed.
    Call before committing an import if you know the count in advance.
    Renders an inline warning and returns False if the limit would be exceeded.
    """
    if within_strategy_limit(new_count):
        return True

    limit = get_strategy_limit()
    st.error(
        f"Importing {new_count} strategies would exceed your plan limit of {limit}. "
        "Upgrade to Full for unlimited strategies or reduce the number of strategy folders."
    )
    return False


# ── Feature-level gates ───────────────────────────────────────────────────────
FULL_ONLY_FEATURES = {
    "Portfolio Optimizer",
    "Leave One Out",
    "Margin Tracking",
    "Market Analysis",
    "Portfolio Compare",
    "Portfolio History",
}


def gate_full_plan(feature_name: str) -> bool:
    """
    Returns True if the user has a Full plan (gate passes).
    Otherwise renders an upgrade prompt and returns False.

    Usage:
        if not gate_full_plan("Portfolio Optimizer"):
            st.stop()
    """
    if get_plan() == "full":
        return True

    st.warning(
        f"**{feature_name}** is available on the **Full plan** ($49/mo). "
        "Upgrade to unlock unlimited strategies and all analytics modules."
    )

    from ui.pricing import render_pricing_cards

    with st.expander("View plans", expanded=True):
        render_pricing_cards(show_actions=True)

    return False


# ── Sidebar plan badge ────────────────────────────────────────────────────────
def render_plan_badge() -> None:
    """Render the current plan and usage as a compact sidebar widget."""
    if not is_subscribed():
        st.caption("No active subscription")
        from ui.pricing import fetch_checkout_url, js_redirect

        if st.button("Subscribe", use_container_width=True, type="primary"):
            url = fetch_checkout_url("lite")
            if url:
                js_redirect(url)
        return

    plan = get_plan() or "unknown"
    limit = get_strategy_limit()
    imported = st.session_state.get("imported_data")

    count_str = ""
    if imported is not None:
        try:
            count = len(imported.strategy_names)
            limit_str = str(limit) if limit is not None else "∞"
            count_str = f"{count} / {limit_str} strategies"
        except AttributeError:
            pass

    plan_color = "#3b82f6" if plan == "lite" else "#8b5cf6"
    st.markdown(
        f'<span style="background:{plan_color}22;color:{plan_color};'
        f'border-radius:4px;padding:2px 8px;font-size:0.72rem;font-weight:700;">'
        f"{plan.upper()} PLAN</span>",
        unsafe_allow_html=True,
    )
    if count_str:
        st.caption(count_str)
