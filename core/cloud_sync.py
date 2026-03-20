"""
Per-user cloud persistence via Supabase.

Settings (AppConfig + strategies) are stored in the `user_settings` Postgres
table.  CSV files are stored in the `user-csv-data` Supabase Storage bucket
under the path  <user_id>/<filename>.

All public functions are designed to fail silently so that a Supabase outage
never breaks the local-only workflow — they return None / False / [] on any
error.

Prerequisites (run supabase/schema.sql in the Supabase SQL editor):
  - public.user_settings table with RLS
  - Storage bucket "user-csv-data" (private) with per-user RLS policies
"""

from __future__ import annotations

import os
from typing import Optional

import streamlit as st

from auth.supabase_client import get_supabase
from auth.session import get_user, get_access_token

_SETTINGS_TABLE = "user_settings"
_STORAGE_BUCKET = "user-csv-data"


# ── Internal helpers ───────────────────────────────────────────────────────────

def _authed_postgrest():
    """Supabase client with the current user's JWT set for PostgREST (table) access."""
    sb = get_supabase()
    token = get_access_token()
    if token:
        sb.postgrest.auth(token)
    return sb


def _authed_storage():
    """
    Supabase Storage client authenticated with the current user's JWT.

    Creates a SyncStorageClient directly (from the storage3 package, which is
    a dependency of supabase-py) so we can pass the user's access token as the
    Authorization header without touching the shared anon-key client.
    """
    url: str = st.secrets.get("SUPABASE_URL", "") or os.getenv("SUPABASE_URL", "")
    anon_key: str = st.secrets.get("SUPABASE_ANON_KEY", "") or os.getenv("SUPABASE_ANON_KEY", "")
    token = get_access_token() or anon_key

    from storage3 import SyncStorageClient  # ships with supabase>=2.4.0
    return SyncStorageClient(
        url=f"{url}/storage/v1",
        headers={"apikey": anon_key, "Authorization": f"Bearer {token}"},
    )


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


# ── CSV file storage ───────────────────────────────────────────────────────────

def upload_csv(file_name: str, file_bytes: bytes) -> bool:
    """
    Upload a CSV file to Supabase Storage at  <user_id>/<file_name>.
    Overwrites any existing file with the same name (upsert=true).
    Returns True on success, False on any error.
    """
    user = get_user()
    if not user:
        return False
    try:
        storage = _authed_storage()
        path = f"{user.id}/{file_name}"
        storage.from_(_STORAGE_BUCKET).upload(
            path,
            file_bytes,
            file_options={"content-type": "text/csv", "upsert": "true"},
        )
        return True
    except Exception:
        return False


def list_user_csvs() -> list[str]:
    """
    Return a list of CSV file names in the current user's storage folder.
    Returns an empty list on any error.
    """
    user = get_user()
    if not user:
        return []
    try:
        storage = _authed_storage()
        files = storage.from_(_STORAGE_BUCKET).list(user.id)
        return [
            f["name"]
            for f in (files or [])
            if isinstance(f, dict) and f.get("name", "").lower().endswith(".csv")
        ]
    except Exception:
        return []


def download_csv(file_name: str) -> Optional[bytes]:
    """
    Download a CSV file from the current user's storage folder.
    Returns None on any error.
    """
    user = get_user()
    if not user:
        return None
    try:
        storage = _authed_storage()
        path = f"{user.id}/{file_name}"
        return storage.from_(_STORAGE_BUCKET).download(path)
    except Exception:
        return None


def delete_csv(file_name: str) -> bool:
    """
    Delete a CSV file from the current user's storage folder.
    Returns True on success, False on any error.
    """
    user = get_user()
    if not user:
        return False
    try:
        storage = _authed_storage()
        path = f"{user.id}/{file_name}"
        storage.from_(_STORAGE_BUCKET).remove([path])
        return True
    except Exception:
        return False
