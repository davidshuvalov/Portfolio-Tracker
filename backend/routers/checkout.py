"""
Stripe Checkout and Billing Portal session creation.

All endpoints require a valid Supabase JWT in the Authorization header.
The JWT is validated server-side using the Supabase service-role client.
"""

from __future__ import annotations

import os

import stripe
from fastapi import APIRouter, Depends, Header, HTTPException
from pydantic import BaseModel
from supabase import Client, create_client

router = APIRouter()

# ── Plan → Stripe price mapping ───────────────────────────────────────────────
VALID_PLANS = {"lite", "full"}


def _stripe() -> None:
    """Set Stripe API key from env (call before any stripe.* call)."""
    stripe.api_key = os.environ["STRIPE_SECRET_KEY"]


def _supabase_admin() -> Client:
    return create_client(
        os.environ["SUPABASE_URL"],
        os.environ["SUPABASE_SERVICE_ROLE_KEY"],
    )


async def _current_user(authorization: str = Header(default=None)):
    """Validate the Supabase JWT and return the User object."""
    if not authorization or not authorization.startswith("Bearer "):
        raise HTTPException(status_code=401, detail="Missing or invalid Authorization header")
    token = authorization.removeprefix("Bearer ").strip()
    try:
        admin = _supabase_admin()
        response = admin.auth.get_user(token)
        return response.user
    except Exception as exc:
        raise HTTPException(status_code=401, detail=f"Invalid token: {exc}") from exc


# ── Request models ─────────────────────────────────────────────────────────────
class CreateCheckoutRequest(BaseModel):
    plan: str  # "lite" | "full"


# ── Helpers ───────────────────────────────────────────────────────────────────
def _get_or_create_stripe_customer(user_id: str, email: str, admin: Client) -> str:
    """Return an existing stripe_customer_id or create a new Stripe customer."""
    _stripe()
    result = (
        admin.table("profiles")
        .select("stripe_customer_id")
        .eq("user_id", user_id)
        .single()
        .execute()
    )
    customer_id: str | None = (result.data or {}).get("stripe_customer_id")

    if not customer_id:
        customer = stripe.Customer.create(
            email=email,
            metadata={"supabase_user_id": user_id},
        )
        customer_id = customer.id
        admin.table("profiles").update({"stripe_customer_id": customer_id}).eq(
            "user_id", user_id
        ).execute()

    return customer_id


# ── Endpoints ─────────────────────────────────────────────────────────────────
@router.post("/create-checkout-session")
async def create_checkout_session(
    body: CreateCheckoutRequest,
    user=Depends(_current_user),
):
    """
    Create a Stripe Checkout Session for a subscription plan.
    Returns {"url": "<stripe hosted page URL>"}.
    """
    if body.plan not in VALID_PLANS:
        raise HTTPException(status_code=400, detail=f"Invalid plan '{body.plan}'. Must be one of {VALID_PLANS}")

    price_id = os.environ.get(f"STRIPE_PRICE_ID_{body.plan.upper()}")
    if not price_id:
        raise HTTPException(status_code=500, detail=f"Price ID not configured for plan '{body.plan}'")

    admin = _supabase_admin()
    customer_id = _get_or_create_stripe_customer(user.id, user.email, admin)

    _stripe()
    session = stripe.checkout.Session.create(
        customer=customer_id,
        payment_method_types=["card"],
        line_items=[{"price": price_id, "quantity": 1}],
        mode="subscription",
        success_url=f"{os.environ['FRONTEND_URL']}?checkout=success",
        cancel_url=f"{os.environ['FRONTEND_URL']}?checkout=cancel",
        # Embed user ID so the webhook can find the profile even if
        # checkout.session.completed fires before subscription.created
        client_reference_id=user.id,
        subscription_data={
            "metadata": {"supabase_user_id": user.id, "plan": body.plan}
        },
    )

    return {"url": session.url}


@router.post("/create-billing-portal-session")
async def create_billing_portal_session(user=Depends(_current_user)):
    """
    Create a Stripe Billing Portal Session so the user can manage their
    subscription, update payment methods, or cancel.
    Returns {"url": "<stripe portal URL>"}.
    """
    admin = _supabase_admin()
    result = (
        admin.table("profiles")
        .select("stripe_customer_id")
        .eq("user_id", user.id)
        .single()
        .execute()
    )
    customer_id: str | None = (result.data or {}).get("stripe_customer_id")

    if not customer_id:
        raise HTTPException(
            status_code=400,
            detail="No Stripe customer found. Complete a checkout first.",
        )

    _stripe()
    session = stripe.billing_portal.Session.create(
        customer=customer_id,
        return_url=os.environ["FRONTEND_URL"],
    )
    return {"url": session.url}
