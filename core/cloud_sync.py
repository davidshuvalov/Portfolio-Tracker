"""
Per-user cloud persistence via Supabase.

Settings (AppConfig + strategies) are stored in the `user_settings` Postgres
table.  Users install and run locally; on login the app checks Supabase to
verify an active subscription.  Settings/strategies sync to Supabase so the
user can restore them on reinstall.

All public functions are designed to fail silently so that a Supabase outage
never breaks the local-only workflow — they return None / False on any error.

Prerequisites (run supabase/schema.sql in the Supabase SQL editor):
  - public.user_settings table with RLS
"""

from __future__ import annotations

from typing import Optional

from auth.supabase_client import get_supabase
from auth.session import get_user, get_access_token

_SETTINGS_TABLE = "user_settings"


# ── Internal helpers ───────────────────────────────────────────────────────────

def _authed_postgrest():
    """Supabase client with the current user's JWT set for PostgREST (table) access."""
    sb = get_supabase()
    token = get_access_token()
    if token:
        sb.postgrest.auth(token)
    return sb


# ── App settings ──────────────────────────────────────────────────────────────

def save_settings_to_cloud(settings_dict: dict) -> bool:
    """
    Upsert the serialised AppConfig dict to Supabase for the current user.
    Returns True on success, False on any error.
    """
    user = get_user()
    if not user:
        return False
    try:
        sb = _authed_postgrest()
        sb.table(_SETTINGS_TABLE).upsert(
            {"user_id": user.id, "settings_json": settings_dict},
            on_conflict="user_id",
        ).execute()
        return True
    except Exception:
        return False


def load_settings_from_cloud() -> Optional[dict]:
    """
    Load the serialised AppConfig dict from Supabase for the current user.
    Returns None if not found or on any error.
    """
    user = get_user()
    if not user:
        return None
    try:
        sb = _authed_postgrest()
        result = (
            sb.table(_SETTINGS_TABLE)
            .select("settings_json")
            .eq("user_id", user.id)
            .maybe_single()
            .execute()
        )
        if result.data:
            data = result.data.get("settings_json")
            # Only return if it's a non-empty dict (avoid overwriting with empty defaults)
            return data if isinstance(data, dict) and data else None
        return None
    except Exception:
        return None


# ── Strategy configuration ─────────────────────────────────────────────────────

def save_strategies_to_cloud(strategies: list) -> bool:
    """
    Upsert the strategies list to Supabase for the current user.
    Returns True on success, False on any error.
    """
    user = get_user()
    if not user:
        return False
    try:
        sb = _authed_postgrest()
        sb.table(_SETTINGS_TABLE).upsert(
            {"user_id": user.id, "strategies_json": strategies},
            on_conflict="user_id",
        ).execute()
        return True
    except Exception:
        return False


def load_strategies_from_cloud() -> Optional[list]:
    """
    Load the strategies list from Supabase for the current user.
    Returns None if not found or on any error.
    """
    user = get_user()
    if not user:
        return None
    try:
        sb = _authed_postgrest()
        result = (
            sb.table(_SETTINGS_TABLE)
            .select("strategies_json")
            .eq("user_id", user.id)
            .maybe_single()
            .execute()
        )
        if result.data:
            data = result.data.get("strategies_json")
            return data if isinstance(data, list) else None
        return None
    except Exception:
        return None
