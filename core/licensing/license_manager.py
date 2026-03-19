"""
License validation — mirrors B_Licencing_Checks.bas.

Validation chain (Windows only):
  1. Customer ID must be in the hardcoded valid-customer list
     (from A_List_of_Valid_Licences.bas).
  2. MultiWalk program folder is read from the Windows registry
     (HKCU\\SOFTWARE\\MultiWalk\\MultiWalkProgramFolder).
  3. MultiWalkLicense64.dll (or 32-bit) is loaded via ctypes.
  4. MultiWalkIsLicensePro(folder, customer_id, "ShuvalovPortfolio") must
     return 0 (success).

On non-Windows platforms the DLL cannot be called; validate_full() returns
(False, "Windows required") so the app never unlocks on Mac/Linux in
production — but tests and DEV_MODE bypass this.

Return codes from MultiWalkIsLicensePro (mirrors VBA Select Case):
  0  — License valid
  1  — Invalid program folder
  3  — No license key file found
  4  — Multiple license keys found
"""

from __future__ import annotations

import ctypes
import sys
from pathlib import Path

# ── Valid customer IDs (from A_List_of_Valid_Licences.bas) ────────────────────

_VALID_CUSTOMER_IDS: frozenset[int] = frozenset([
    4209838,  # David Shuvalov
    1491687,  # Dave Fisher
    2227285,  # Shaun Lawman
    4628537,  # Christopher Kilgore
    767072,   # Al Biddinger
    1304080,  # Zamir & Dunn
    1577327,  # Timo Mohnani
    4576263,  # Bert Trouwers
    3352708,  # Simon Gale
    1438899,  # Clay Crandall
    4584703,  # Michal Filipkowski
    1623185,  # Robert Fleming
    982187,   # Sean Cooper
    3534333,  # chad ockham
    3281546,  # Greg Baker
    4132757,  # Don LaPel
    2706911,  # Philippe Bremard
    2791698,  # Ed Tulauskas
    874960,   # Sanjay Sardana
    3161289,  # Marc Jusseaume
    1993965,  # James Parker
    1224801,  # Gary McOmber
    1426160,  # Daniel Bangert
    4305976,  # David Aczel
    4056398,  # Niko Heir
    1447036,  # James Welborn
    828124,   # Mark Holland
    4304998,  # Nikita Gorbachenko
    2749694,  # Jayce Nugent
    2697537,  # Edwin Shih
    4565694,  # Victor Stokmans
    2546277,  # Herman Fuchs
    2210950,  # Rajendra Deshpande
    4653662,  # Rey Farne
    3923344,  # Ujae Kang
    2986779,  # Dave edwards
    1411262,  # Jonas Hellwig
    3144964,  # Justin Krick
    2069314,  # Tom Garesche
    4247492,  # Love Englund
    2453776,  # Livio Pietroboni
    4400422,  # Ender Araujo
    2809194,  # Dan Omalley
    3273971,  # Covington Creek
    1966649,  # Ron Mullet
    588649,   # James Mazzolini
    4363638,  # Seuk Oh
    645309,   # John Dorsey
    4613075,  # Haro Hollertt
    3213888,  # Ryan Williams
    4518976,  # Jani Talikka
    4363735,  # Richard Moore
    1808839,  # Vernon Pratt
    3870813,  # Michal Kodousek
    2824178,  # Robert Roubey
    3551682,  # Venkatesh Yarraguntla
    3285215,  # Venkatesh Yarraguntla (2nd)
    4744449,  # Pete LaDuke
    3518605,  # Youssef Oumanar
    4230663,  # Andreas Savva
    2335131,  # Timothy Krull
    3488626,  # Eric Rosko
    4795843,  # Denis Smirnov
    3277186,  # Thomas Uselton
    4352636,  # Anguera Antonio
    4301470,  # Miguel Bermejo
    4760562,  # Rohan Patil
    2139616,  # OUrocketman
    4877957,  # Arin
    4408498,  # Arturo Patino
    4874084,  # Younes Zerhari
])

_APP_NAME = "ShuvalovPortfolio"

# DLL return codes
_RC_VALID = 0
_RC_INVALID_FOLDER = 1
_RC_NO_KEY_FILE = 3
_RC_MULTIPLE_KEYS = 4

_RC_MESSAGES = {
    _RC_INVALID_FOLDER: "Invalid MultiWalk program folder.",
    _RC_NO_KEY_FILE:    "No license key file found.",
    _RC_MULTIPLE_KEYS:  "Multiple license keys found.",
}


# ── Registry helper ───────────────────────────────────────────────────────────

def get_multiwalk_folder() -> str | None:
    """
    Read MultiWalkProgramFolder from the Windows registry.
    Returns None on non-Windows or if MultiWalk is not installed.
    Mirrors GetMultiWalkProgramFolder() in B_Licencing_Checks.bas.
    """
    if sys.platform != "win32":
        return None
    try:
        import winreg
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"SOFTWARE\MultiWalk")
        value, _ = winreg.QueryValueEx(key, "MultiWalkProgramFolder")
        winreg.CloseKey(key)
        return str(value).strip()
    except Exception:
        return None


# ── DLL caller ────────────────────────────────────────────────────────────────

def _call_dll(folder: str, customer_id: int) -> tuple[bool, str]:
    """
    Load MultiWalkLicense64.dll and call MultiWalkIsLicensePro.
    Returns (success, error_message).
    Mirrors LibraryLoaded() + IsLicenseValid() in B_Licencing_Checks.bas.
    """
    if sys.platform != "win32":
        return False, "License DLL requires Windows."

    dll_path = Path(folder) / "MultiWalkLicense64.dll"
    if not dll_path.exists():
        # Fall back to 32-bit
        dll_path = Path(folder) / "MultiWalkLicense32.dll"
    if not dll_path.exists():
        return False, f"MultiWalk license DLL not found in: {folder}"

    try:
        dll = ctypes.WinDLL(str(dll_path))  # type: ignore[attr-defined]
        fn = dll["_MultiWalkIsLicensePro"]
        fn.restype = ctypes.c_int
        fn.argtypes = [
            ctypes.c_char_p,    # program_folder (ANSI, matches VBA ByVal String)
            ctypes.c_long,      # ts_customer_id
            ctypes.c_char_p,    # app_name (ANSI, matches VBA ByVal String)
        ]
        rc = fn(
            folder.encode("mbcs"),
            ctypes.c_long(customer_id),
            _APP_NAME.encode("mbcs"),
        )
    except Exception as e:
        return False, f"DLL call failed: {e}"

    if rc == _RC_VALID:
        return True, ""
    return False, _RC_MESSAGES.get(rc, f"Unexpected DLL return code: {rc}")


# ── Public API ────────────────────────────────────────────────────────────────

def is_known_customer(customer_id: int) -> bool:
    """Return True if customer_id is in the valid-customer list."""
    return customer_id in _VALID_CUSTOMER_IDS


def validate_lemon_squeezy(license_key: str) -> tuple[bool, str]:
    """
    Validate a Lemon Squeezy license key.
    This is the primary path for new customers who purchased via the Lemon Squeezy store.
    """
    from core.licensing.lemon_squeezy import validate as _ls_validate
    return _ls_validate(license_key)


def validate_full(customer_id: int, multiwalk_folder: str = "") -> tuple[bool, str]:
    """
    Full license validation — mirrors IsLicenseValid() in B_Licencing_Checks.bas.

    Steps:
      1. Check customer_id against the known-customer list.
      2. Resolve MultiWalk program folder: use multiwalk_folder if provided,
         otherwise read from the Windows registry.
      3. Call MultiWalkIsLicensePro via ctypes.

    Returns:
        (True, "")                — license valid
        (False, error_message)   — license invalid with reason
    """
    if not customer_id:
        return False, "No TradeStation Customer ID provided."

    if not is_known_customer(customer_id):
        return False, (
            f"Customer ID {customer_id} is not in the licensed-customer list. "
            "Please contact david@portfoliotracker.com for access."
        )

    folder = multiwalk_folder.strip() if multiwalk_folder else get_multiwalk_folder()
    if not folder:
        return False, (
            "MultiWalk program folder not found in the Windows registry. "
            "Enter it manually below."
        )

    return _call_dll(folder, customer_id)
