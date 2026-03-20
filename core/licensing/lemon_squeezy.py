"""
Lemon Squeezy license validation for Portfolio Tracker.

Flow (per-machine):
  1. First run: POST /v1/licenses/activate  → saves instance_id to local cache
  2. Subsequent runs: serve from cache if < OFFLINE_GRACE_DAYS old
  3. Otherwise: POST /v1/licenses/validate  → refreshes cache on success
  4. If offline and cache is valid (any age): allow with warning

Lemon Squeezy's activation/validation endpoints are unauthenticated by design
(intended for client-side desktop apps) — no API key is embedded in the app.

License key format (Lemon Squeezy default): XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
"""

from __future__ import annotations

import base64
import getpass
import hashlib
import json
import re
import socket
from datetime import datetime, timedelta, timezone
from pathlib import Path

import requests
from cryptography.fernet import Fernet, InvalidToken

# ── Config ────────────────────────────────────────────────────────────────────

_LS_BASE = "https://api.lemonsqueezy.com/v1/licenses"
_TIMEOUT = 8           # seconds
_GRACE_DAYS = 7        # days before a cached validation is re-checked online
_CACHE_FILE = Path("~/.portfolio_tracker/license_cache.json").expanduser()

# Lemon Squeezy license keys look like a UUID
_LS_KEY_PATTERN = re.compile(
    r"^[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}$",
    re.IGNORECASE,
)


# ── Public helpers ─────────────────────────────────────────────────────────────

def is_ls_key(value: str) -> bool:
    """Return True if the string looks like a Lemon Squeezy license key."""
    return bool(_LS_KEY_PATTERN.match(value.strip()))


def validate(license_key: str) -> tuple[bool, str]:
    """
    Validate a Lemon Squeezy license key for this machine.

    Returns:
        (True,  "")            — valid
        (False, error_message) — invalid
    """
    key = license_key.strip().upper()
    cache = _load_cache()

    if cache.get("license_key", "").upper() == key:
        # ── Already activated on this machine ────────────────────────────────
        last_ok_str = cache.get("last_validated")
        if last_ok_str:
            try:
                last_ok = datetime.fromisoformat(last_ok_str)
                if last_ok.tzinfo is None:
                    last_ok = last_ok.replace(tzinfo=timezone.utc)
                age = datetime.now(timezone.utc) - last_ok
                if age < timedelta(days=_GRACE_DAYS):
                    return True, ""   # within grace period — no network call
            except Exception:
                pass

        # Try online validation to refresh the cache
        instance_id = cache.get("instance_id", "")
        ok, msg = _validate_online(key, instance_id)
        if ok:
            return True, ""
        # Offline fallback — still allow if we have a stored valid state
        if cache.get("valid"):
            return True, "Offline — using cached license (will re-verify when online)."
        return False, msg

    # ── Not yet activated on this machine — activate now ─────────────────────
    return _activate(key)


def deactivate(license_key: str) -> tuple[bool, str]:
    """
    Deactivate this machine's instance. Call when the user uninstalls.
    Silently succeeds if no cache exists.
    """
    key = license_key.strip().upper()
    cache = _load_cache()
    instance_id = cache.get("instance_id", "")

    if not instance_id:
        _clear_cache()
        return True, "No active instance to deactivate."

    try:
        r = requests.post(
            f"{_LS_BASE}/deactivate",
            json={"license_key": key, "instance_id": instance_id},
            headers={"Accept": "application/json"},
            timeout=_TIMEOUT,
        )
        _clear_cache()
        return True, "License deactivated."
    except Exception:
        _clear_cache()
        return True, "Cleared local license (could not reach server)."


# ── Internal ───────────────────────────────────────────────────────────────────

def _activate(key: str) -> tuple[bool, str]:
    """POST /v1/licenses/activate — first activation on this machine."""
    try:
        r = requests.post(
            f"{_LS_BASE}/activate",
            json={"license_key": key, "instance_name": _machine_name()},
            headers={"Accept": "application/json"},
            timeout=_TIMEOUT,
        )
        data = r.json()
    except requests.exceptions.Timeout:
        return False, "Could not reach the license server. Check your internet connection."
    except Exception as exc:
        return False, f"Network error: {exc}"

    if r.status_code == 200 and data.get("activated"):
        instance_id = data.get("instance", {}).get("id", "")
        _save_cache({
            "license_key": key,
            "instance_id": instance_id,
            "last_validated": datetime.now(timezone.utc).isoformat(),
            "valid": True,
        })
        return True, ""

    # If already activated on this machine Lemon Squeezy returns 400 with
    # an error string containing "already". Try validate instead.
    error = str(data.get("error", "")).lower()
    if r.status_code == 400 and "already" in error:
        return _validate_online(key, "")

    msg = data.get("error") or f"Activation failed (HTTP {r.status_code})"
    return False, str(msg)


def _validate_online(key: str, instance_id: str) -> tuple[bool, str]:
    """POST /v1/licenses/validate — check an activated key."""
    body: dict = {"license_key": key}
    if instance_id:
        body["instance_id"] = instance_id

    try:
        r = requests.post(
            f"{_LS_BASE}/validate",
            json=body,
            headers={"Accept": "application/json"},
            timeout=_TIMEOUT,
        )
        data = r.json()
    except requests.exceptions.Timeout:
        return False, "Could not reach the license server. Check your internet connection."
    except Exception as exc:
        return False, f"Network error: {exc}"

    if r.status_code == 200 and data.get("valid"):
        cache = _load_cache()
        cache.update({
            "license_key": key,
            "last_validated": datetime.now(timezone.utc).isoformat(),
            "valid": True,
        })
        if instance_id:
            cache["instance_id"] = instance_id
        _save_cache(cache)
        return True, ""

    msg = data.get("error") or f"Validation failed (HTTP {r.status_code})"
    return False, str(msg)


# ── Cache helpers ──────────────────────────────────────────────────────────────

def _get_fernet() -> Fernet:
    """Derive a stable, machine-specific Fernet key from hostname + username."""
    machine_id = f"portfolio-tracker:{socket.gethostname()}:{getpass.getuser()}"
    key_bytes = hashlib.sha256(machine_id.encode()).digest()  # 32 bytes
    return Fernet(base64.urlsafe_b64encode(key_bytes))


def _load_cache() -> dict:
    try:
        if _CACHE_FILE.exists():
            raw = _CACHE_FILE.read_bytes()
            try:
                # Try decrypted read first
                plaintext = _get_fernet().decrypt(raw)
                return json.loads(plaintext)
            except (InvalidToken, Exception):
                # Fallback: plaintext cache from before encryption was added
                return json.loads(raw.decode("utf-8"))
    except Exception:
        pass
    return {}


def _save_cache(data: dict) -> None:
    try:
        _CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        ciphertext = _get_fernet().encrypt(json.dumps(data).encode("utf-8"))
        _CACHE_FILE.write_bytes(ciphertext)
    except Exception:
        pass


def _clear_cache() -> None:
    try:
        if _CACHE_FILE.exists():
            _CACHE_FILE.unlink()
    except Exception:
        pass


def _machine_name() -> str:
    try:
        return socket.gethostname()
    except Exception:
        return "unknown"
