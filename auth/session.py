"""
Streamlit session helpers for Supabase Auth.

Session state keys used:
  _sb_session   — gotrue.Session object (access_token, refresh_token, user)
  _sb_profile   — dict from public.profiles (plan, limits, etc.)
"""

from __future__ import annotations

import os
import time
from typing import Optional

import streamlit as st

from .supabase_client import get_supabase

# ── Internal helpers ──────────────────────────────────────────────────────────

def _session():
    return st.session_state.get("_sb_session")


def _set_session(s) -> None:
    st.session_state["_sb_session"] = s


def _clear_session() -> None:
    st.session_state.pop("_sb_session", None)
    st.session_state.pop("_sb_profile", None)


# ── Public API ────────────────────────────────────────────────────────────────

def is_logged_in() -> bool:
    return _session() is not None


def get_user():
    """Return the gotrue User object or None."""
    s = _session()
    return s.user if s else None


def get_access_token() -> Optional[str]:
    s = _session()
    return s.access_token if s else None


def get_profile() -> Optional[dict]:
    return st.session_state.get("_sb_profile")


def fetch_and_cache_profile() -> Optional[dict]:
    """Load the profile from Supabase and cache it in session state."""
    _maybe_refresh_session()
    user = get_user()
    if not user:
        return None
    access_token = get_access_token()
    try:
        sb = get_supabase()
        # Set the user's JWT so RLS (auth.uid() = user_id) allows the read.
        # get_supabase() returns a fresh anon client each time, so we must
        # explicitly attach the session token before making PostgREST calls.
        if access_token:
            sb.postgrest.auth(access_token)
        result = (
            sb.table("profiles")
            .select("*")
            .eq("user_id", user.id)
            .maybe_single()
            .execute()
        )
        # result.data is None when no profile row exists yet (new / unsubscribed user)
        profile = result.data if result.data is not None else {}
        st.session_state["_sb_profile"] = profile
        return profile
    except Exception:
        return None


def invalidate_profile() -> None:
    """Force the profile to be re-fetched on the next access."""
    st.session_state.pop("_sb_profile", None)


def _maybe_refresh_session() -> None:
    """
    Refresh the Supabase JWT if it has expired or expires within 60 seconds.
    Without this, users silently hit a 401 error after ~1 hour and see the
    'Could not load your account profile' screen until they manually log out.
    """
    s = _session()
    if not s:
        return
    expires_at = getattr(s, "expires_at", None)
    if expires_at is None or time.time() < expires_at - 60:
        return  # Token still valid
    try:
        sb = get_supabase()
        sb.auth.set_session(s.access_token, s.refresh_token)
        refreshed = sb.auth.refresh_session()
        if refreshed and refreshed.session:
            _set_session(refreshed.session)
    except Exception:
        pass  # Let the caller surface the auth error naturally


# ── Auth actions ──────────────────────────────────────────────────────────────

def login(email: str, password: str) -> bool:
    """Sign in with email + password. Returns True on success."""
    try:
        sb = get_supabase()
        response = sb.auth.sign_in_with_password({"email": email, "password": password})
        _set_session(response.session)
        fetch_and_cache_profile()
        return True
    except Exception as exc:
        st.error(f"Login failed: {exc}")
        return False


def signup(email: str, password: str) -> tuple[bool, str]:
    """
    Create a new account.
    Returns (success, message).
    If Supabase email confirmation is enabled, session will be None until
    the user clicks the confirmation link.
    """
    try:
        sb = get_supabase()
        response = sb.auth.sign_up({"email": email, "password": password})
        if response.session:
            # Email confirmation disabled — user is logged in immediately
            _set_session(response.session)
            fetch_and_cache_profile()
            return True, "Account created."
        else:
            # Email confirmation required
            return True, "Check your email to confirm your account, then log in."
    except Exception as exc:
        return False, str(exc)


def reset_password(email: str) -> tuple[bool, str]:
    """Send a password reset email via Supabase. Returns (success, message)."""
    try:
        sb = get_supabase()
        app_url = st.secrets.get("APP_URL", "") or os.getenv("APP_URL", "")
        options = {"redirect_to": app_url} if app_url else {}
        sb.auth.reset_password_email(email, options)
        return True, "Password reset email sent. Check your inbox."
    except Exception as exc:
        return False, str(exc)


def exchange_recovery_code(code: str) -> bool:
    """Exchange a PKCE recovery code for a session. Returns True on success."""
    try:
        sb = get_supabase()
        response = sb.auth.exchange_code_for_session(code)
        if response.session:
            _set_session(response.session)
            return True
        return False
    except Exception:
        return False


def update_password(new_password: str) -> tuple[bool, str]:
    """Update the current user's password. Returns (success, message)."""
    try:
        sb = get_supabase()
        sb.auth.update_user({"password": new_password})
        return True, "Password updated successfully."
    except Exception as exc:
        return False, str(exc)


def logout() -> None:
    try:
        sb = get_supabase()
        sb.auth.sign_out()
    except Exception:
        pass
    _clear_session()
    st.rerun()


# ── Plan helpers ──────────────────────────────────────────────────────────────

ACTIVE_STATUSES = {"active", "trialing"}


def is_subscribed() -> bool:
    profile = get_profile()
    if not profile:
        return False
    return profile.get("subscription_status") in ACTIVE_STATUSES


def get_plan() -> Optional[str]:
    """Return 'lite', 'full', or None."""
    profile = get_profile()
    if not profile or not is_subscribed():
        return None
    return profile.get("subscription_plan")


def get_strategy_limit() -> Optional[int]:
    """
    Return the maximum number of strategies the user may track.
    None means unlimited.  0 or missing means no active subscription.
    """
    profile = get_profile()
    if not profile or not is_subscribed():
        return 0
    return profile.get("strategy_limit")  # None = unlimited (Full plan)


def within_strategy_limit(current_count: int) -> bool:
    """True if the user can track `current_count` strategies under their plan."""
    limit = get_strategy_limit()
    if limit is None:  # unlimited
        return True
    return current_count <= limit
