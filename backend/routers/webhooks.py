"""
Stripe webhook handler.

Handles these events:
  - checkout.session.completed
  - customer.subscription.created
  - customer.subscription.updated
  - customer.subscription.deleted
  - invoice.paid
  - invoice.payment_failed

All profile writes use the Supabase service-role key (bypasses RLS).
"""

from __future__ import annotations

import os
from datetime import datetime, timezone

import stripe
from fastapi import APIRouter, Header, HTTPException, Request
from supabase import Client, create_client

router = APIRouter()

# ── Plan definitions ──────────────────────────────────────────────────────────
# Maps Stripe price IDs → (plan_name, strategy_limit)
# strategy_limit=None means unlimited.
def _plan_from_price_id(price_id: str) -> tuple[str, int | None]:
    if price_id == os.environ.get("STRIPE_PRICE_ID_LITE"):
        return "lite", 20
    if price_id == os.environ.get("STRIPE_PRICE_ID_FULL"):
        return "full", None
    return "unknown", None


def _supabase_admin() -> Client:
    return create_client(
        os.environ["SUPABASE_URL"],
        os.environ["SUPABASE_SERVICE_ROLE_KEY"],
    )


def _ts_to_iso(epoch: int | None) -> str | None:
    if epoch is None:
        return None
    return datetime.fromtimestamp(epoch, tz=timezone.utc).isoformat()


def _update_profile_by_customer(admin: Client, customer_id: str, patch: dict) -> None:
    """Update profiles row identified by stripe_customer_id."""
    patch["updated_at"] = datetime.now(tz=timezone.utc).isoformat()
    admin.table("profiles").update(patch).eq("stripe_customer_id", customer_id).execute()


def _update_profile_by_user_id(admin: Client, user_id: str, patch: dict) -> None:
    patch["updated_at"] = datetime.now(tz=timezone.utc).isoformat()
    admin.table("profiles").update(patch).eq("user_id", user_id).execute()


# ── Webhook endpoint ──────────────────────────────────────────────────────────
@router.post("/stripe/webhook")
async def stripe_webhook(
    request: Request,
    stripe_signature: str = Header(default=None, alias="stripe-signature"),
):
    payload = await request.body()
    webhook_secret = os.environ["STRIPE_WEBHOOK_SECRET"]

    try:
        event = stripe.Webhook.construct_event(payload, stripe_signature, webhook_secret)
    except stripe.errors.SignatureVerificationError as exc:
        raise HTTPException(status_code=400, detail=f"Invalid Stripe signature: {exc}") from exc

    stripe.api_key = os.environ["STRIPE_SECRET_KEY"]
    admin = _supabase_admin()
    event_type: str = event["type"]
    obj = event["data"]["object"]

    # ── checkout.session.completed ────────────────────────────────────────────
    if event_type == "checkout.session.completed":
        subscription_id: str | None = obj.get("subscription")
        user_id: str | None = obj.get("client_reference_id") or obj.get("metadata", {}).get("supabase_user_id")
        customer_id: str = obj["customer"]

        if subscription_id:
            sub = stripe.Subscription.retrieve(subscription_id)
            price_id = sub["items"]["data"][0]["price"]["id"]
            plan, limit = _plan_from_price_id(price_id)
            patch = {
                "stripe_customer_id": customer_id,
                "subscription_status": sub["status"],
                "subscription_plan": plan,
                "strategy_limit": limit,
                "subscription_current_period_end": _ts_to_iso(sub.get("current_period_end")),
            }
            if user_id:
                _update_profile_by_user_id(admin, user_id, patch)
            else:
                _update_profile_by_customer(admin, customer_id, patch)

    # ── customer.subscription.created / updated ───────────────────────────────
    elif event_type in ("customer.subscription.created", "customer.subscription.updated"):
        customer_id = obj["customer"]
        price_id = obj["items"]["data"][0]["price"]["id"]
        plan, limit = _plan_from_price_id(price_id)
        patch = {
            "subscription_status": obj["status"],
            "subscription_plan": plan,
            "strategy_limit": limit,
            "subscription_current_period_end": _ts_to_iso(obj.get("current_period_end")),
        }
        # Also carry the customer ID in case it wasn't set yet
        patch["stripe_customer_id"] = customer_id

        user_id = obj.get("metadata", {}).get("supabase_user_id")
        if user_id:
            _update_profile_by_user_id(admin, user_id, patch)
        else:
            _update_profile_by_customer(admin, customer_id, patch)

    # ── customer.subscription.deleted ─────────────────────────────────────────
    elif event_type == "customer.subscription.deleted":
        customer_id = obj["customer"]
        _update_profile_by_customer(admin, customer_id, {
            "subscription_status": "canceled",
            "subscription_plan": None,
            "strategy_limit": 0,
            "subscription_current_period_end": None,
        })

    # ── invoice.paid ──────────────────────────────────────────────────────────
    elif event_type == "invoice.paid":
        customer_id = obj["customer"]
        subscription_id = obj.get("subscription")
        patch: dict = {"subscription_status": "active"}
        if subscription_id:
            sub = stripe.Subscription.retrieve(subscription_id)
            patch["subscription_current_period_end"] = _ts_to_iso(sub.get("current_period_end"))
        _update_profile_by_customer(admin, customer_id, patch)

    # ── invoice.payment_failed ────────────────────────────────────────────────
    elif event_type == "invoice.payment_failed":
        customer_id = obj["customer"]
        _update_profile_by_customer(admin, customer_id, {"subscription_status": "past_due"})

    return {"received": True}
