"""
Pricing cards and Stripe session helpers for the Streamlit frontend.

Checkout flow:
  1. User clicks Subscribe button
  2. We POST to the FastAPI backend → get a Stripe Checkout URL
  3. We inject a JS redirect to send the browser to Stripe
  4. After payment, Stripe redirects back to FRONTEND_URL?checkout=success
  5. The webhook fires → profile is updated → user is now subscribed
"""

from __future__ import annotations

import os

import requests
import streamlit as st
import streamlit.components.v1 as components

from auth.session import get_access_token, get_plan, is_subscribed

# Backend URL — set via Streamlit secrets or env var
BACKEND_URL: str = (
    st.secrets.get("BACKEND_URL", "")
    or os.getenv("BACKEND_URL", "http://localhost:8000")
).rstrip("/")

# ── Plan definitions ───────────────────────────────────────────────────────────
PLANS = [
    {
        "key": "lite",
        "name": "Lite",
        "price": "$19",
        "period": "/month",
        "limit_label": "Up to 20 strategies",
        "color": "#3b82f6",
        "highlighted": False,
        "features": [
            "Track up to 20 strategies",
            "Monte Carlo simulation",
            "Correlations & diversification",
            "Portfolio analytics",
            "Eligibility backtest",
            "Excel & PDF export",
        ],
    },
    {
        "key": "full",
        "name": "Full",
        "price": "$49",
        "period": "/month",
        "limit_label": "Unlimited strategies",
        "color": "#8b5cf6",
        "highlighted": True,
        "features": [
            "Unlimited strategies",
            "Everything in Lite",
            "Portfolio optimizer",
            "Leave-one-out analysis",
            "Margin tracking",
            "Market analysis terminal",
            "Portfolio compare & history",
            "What-If backtest",
        ],
    },
]


# ── Backend calls ─────────────────────────────────────────────────────────────
def _auth_headers() -> dict:
    token = get_access_token()
    return {"Authorization": f"Bearer {token}"} if token else {}


def fetch_checkout_url(plan: str) -> str | None:
    """Call the backend to create a Stripe Checkout session. Returns the URL."""
    try:
        r = requests.post(
            f"{BACKEND_URL}/api/create-checkout-session",
            json={"plan": plan},
            headers=_auth_headers(),
            timeout=10,
        )
        if r.ok:
            return r.json()["url"]
        st.error(f"Could not start checkout: {r.json().get('detail', r.text)}")
    except Exception as exc:
        st.error(f"Backend unreachable: {exc}")
    return None


def fetch_billing_portal_url() -> str | None:
    """Call the backend to create a Stripe Billing Portal session. Returns the URL."""
    try:
        r = requests.post(
            f"{BACKEND_URL}/api/create-billing-portal-session",
            headers=_auth_headers(),
            timeout=10,
        )
        if r.ok:
            return r.json()["url"]
        st.error(f"Could not open billing portal: {r.json().get('detail', r.text)}")
    except Exception as exc:
        st.error(f"Backend unreachable: {exc}")
    return None


def js_redirect(url: str) -> None:
    """Redirect the browser to an external URL using a JS snippet."""
    # window.parent targets the top frame even when Streamlit is iframed
    # (e.g. on Streamlit Community Cloud)
    components.html(
        f'<script>window.parent.location.href = "{url}";</script>',
        height=0,
    )


# ── Pricing card renderer ─────────────────────────────────────────────────────
def render_pricing_cards(show_actions: bool = True) -> None:
    """
    Render two pricing cards side-by-side.

    show_actions=True  → show Subscribe / Manage Billing buttons
    show_actions=False → card display only (e.g. on the landing page hero)
    """
    cols = st.columns(2, gap="large")
    current_plan = get_plan() if show_actions else None
    subscribed = is_subscribed() if show_actions else False

    for i, plan in enumerate(PLANS):
        with cols[i]:
            _render_single_card(plan, current_plan, subscribed, show_actions)


def _render_single_card(
    plan: dict,
    current_plan: str | None,
    subscribed: bool,
    show_actions: bool,
) -> None:
    border = plan["color"]
    bg = "#1e1b4b" if plan["highlighted"] else "#0d1626"
    features_html = "".join(
        f'<li style="padding:3px 0;color:#cbd5e1;">✓ {f}</li>'
        for f in plan["features"]
    )

    st.markdown(
        f"""
        <div style="border:2px solid {border};border-radius:12px;padding:24px;
                    background:{bg};margin-bottom:8px;">
          <div style="font-size:1.25rem;font-weight:700;color:{border};">
            {plan['name']}
          </div>
          <div style="font-size:2.4rem;font-weight:800;margin:6px 0 2px;">
            {plan['price']}<span style="font-size:1rem;color:#94a3b8;">&nbsp;{plan['period']}</span>
          </div>
          <div style="color:#64748b;font-size:0.85rem;margin-bottom:14px;">
            {plan['limit_label']}
          </div>
          <ul style="list-style:none;padding:0;margin:0 0 18px 0;font-size:0.9rem;">
            {features_html}
          </ul>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if not show_actions:
        return

    plan_key = plan["key"]
    is_current = subscribed and current_plan == plan_key

    if is_current:
        st.success("✓ Your current plan")
        if st.button("Manage Billing", key=f"portal_{plan_key}", use_container_width=True):
            url = fetch_billing_portal_url()
            if url:
                js_redirect(url)
    else:
        label = (
            f"Upgrade to {plan['name']} — {plan['price']}/mo"
            if subscribed
            else f"Subscribe — {plan['price']}/mo"
        )
        btn_type = "primary" if plan["highlighted"] else "secondary"
        if st.button(label, key=f"sub_{plan_key}", use_container_width=True, type=btn_type):
            url = fetch_checkout_url(plan_key)
            if url:
                js_redirect(url)
