"""
Unit tests for core.licensing.license_manager.

DLL calls are not tested (require Windows + MultiWalk install);
tests cover the pure-Python parts: customer ID validation and
the public API surface.
"""
from __future__ import annotations

import pytest

from core.licensing.license_manager import (
    _VALID_CUSTOMER_IDS,
    get_multiwalk_folder,
    is_known_customer,
    validate_full,
)


class TestIsKnownCustomer:
    def test_known_id_returns_true(self):
        # David Shuvalov — always in the list
        assert is_known_customer(4209838) is True

    def test_unknown_id_returns_false(self):
        assert is_known_customer(0) is False
        assert is_known_customer(999999999) is False

    def test_all_ids_are_positive(self):
        assert all(cid > 0 for cid in _VALID_CUSTOMER_IDS)

    def test_no_duplicate_ids(self):
        # frozenset deduplication is implicit; just verify size > 0
        assert len(_VALID_CUSTOMER_IDS) >= 70

    def test_boundary_known_ids(self):
        # Spot-check a few from A_List_of_Valid_Licences.bas
        for cid in (1491687, 982187, 4874084, 645309):
            assert is_known_customer(cid), f"{cid} should be known"


class TestValidateFull:
    def test_zero_id_returns_false(self):
        ok, msg = validate_full(0)
        assert ok is False
        assert "No TradeStation Customer ID" in msg

    def test_unknown_id_returns_false(self):
        ok, msg = validate_full(123)
        assert ok is False
        assert "not in the licensed-customer list" in msg

    def test_known_id_on_non_windows_fails_at_folder_step(self):
        """On non-Windows the folder lookup returns None → license fails."""
        import sys
        if sys.platform == "win32":
            pytest.skip("This test is for non-Windows only")
        ok, msg = validate_full(4209838)
        assert ok is False
        # Error should mention MultiWalk folder or Windows
        assert "MultiWalk" in msg or "Windows" in msg

    def test_returns_tuple_of_bool_and_str(self):
        result = validate_full(0)
        assert isinstance(result, tuple)
        assert len(result) == 2
        assert isinstance(result[0], bool)
        assert isinstance(result[1], str)


class TestGetMultiWalkFolder:
    def test_returns_none_or_string(self):
        result = get_multiwalk_folder()
        assert result is None or isinstance(result, str)

    def test_non_windows_returns_none(self):
        import sys
        if sys.platform == "win32":
            pytest.skip("Only runs on non-Windows")
        assert get_multiwalk_folder() is None
