"""
Auth form components (login / signup).
Rendered on the landing page for unauthenticated users.
"""

from __future__ import annotations

import streamlit as st

from auth.session import login, signup, reset_password


def render_auth_forms() -> None:
    """Render login and signup tabs."""
    tab_login, tab_signup = st.tabs(["Log In", "Create Account"])

    with tab_login:
        with st.form("pt_login_form", clear_on_submit=False):
            email = st.text_input("Email", placeholder="you@example.com")
            password = st.text_input("Password", type="password")
            submitted = st.form_submit_button("Log In", type="primary", use_container_width=True)

        if submitted:
            if not email or not password:
                st.error("Please enter your email and password.")
            elif login(email, password):
                st.rerun()

        with st.expander("Forgot your password?"):
            reset_email = st.text_input("Enter your account email", placeholder="you@example.com", key="reset_email")
            if st.button("Send Reset Link", use_container_width=True):
                if not reset_email:
                    st.error("Please enter your email address.")
                else:
                    ok, msg = reset_password(reset_email)
                    if ok:
                        st.success(msg)
                    else:
                        st.error(f"Reset failed: {msg}")

    with tab_signup:
        with st.form("pt_signup_form", clear_on_submit=True):
            email = st.text_input("Email", placeholder="you@example.com", key="su_email")
            password = st.text_input("Password (min 8 chars)", type="password", key="su_pw1")
            password2 = st.text_input("Confirm Password", type="password", key="su_pw2")
            submitted = st.form_submit_button(
                "Create Account", type="primary", use_container_width=True
            )

        if submitted:
            if not email or not password:
                st.error("Email and password are required.")
            elif len(password) < 8:
                st.error("Password must be at least 8 characters.")
            elif password != password2:
                st.error("Passwords do not match.")
            else:
                ok, msg = signup(email, password)
                if ok:
                    st.success(msg)
                    if "logged in" not in msg.lower() and "account created" in msg.lower():
                        st.rerun()
                    elif "check your email" in msg.lower():
                        pass  # Show message, no rerun needed
                    else:
                        st.rerun()
                else:
                    st.error(f"Sign-up failed: {msg}")
