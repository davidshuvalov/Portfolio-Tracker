"""
Supabase client initialisation for the Streamlit frontend.

Uses the PUBLIC anon key — this is intentional and safe.
Row-Level Security on the `profiles` table ensures users can only
read their own row. All privileged writes go through the FastAPI backend
which holds the service-role key.
"""

from __future__ import annotations

import os

import streamlit as st
from supabase import Client, create_client


def get_supabase() -> Client:
    """Return a Supabase client using the anon (public) key."""
    url: str = st.secrets.get("SUPABASE_URL", "") or os.getenv("SUPABASE_URL", "")
    key: str = st.secrets.get("SUPABASE_ANON_KEY", "") or os.getenv("SUPABASE_ANON_KEY", "")

    if not url or not key:
        st.error(
            "Supabase credentials not configured. "
            "Add SUPABASE_URL and SUPABASE_ANON_KEY to `.streamlit/secrets.toml`."
        )
        st.stop()

    return create_client(url, key)
