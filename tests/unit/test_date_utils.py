"""
Tests for core/ingestion/date_utils.py

Mirrors the exact behaviour of VBA I_MISC.bas date functions.
All edge cases documented in the architecture spec are covered.
"""

import pytest
from datetime import date, timedelta

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from core.ingestion.date_utils import (
    detect_date_format,
    parse_csv_date,
    resolve_oos_dates,
    cutoff_index,
    is_non_trading_day,
    _easter_sunday,
    _cme_holidays,
)
import pandas as pd


# ── parse_csv_date ────────────────────────────────────────────────────────────

class TestParseCsvDate:
    def test_dmy_standard(self):
        assert parse_csv_date("15/06/2023", "DMY") == date(2023, 6, 15)

    def test_mdy_standard(self):
        assert parse_csv_date("06/15/2023", "MDY") == date(2023, 6, 15)

    def test_dmy_with_dashes(self):
        assert parse_csv_date("15-06-2023", "DMY") == date(2023, 6, 15)

    def test_mdy_with_dashes(self):
        assert parse_csv_date("06-15-2023", "MDY") == date(2023, 6, 15)

    def test_empty_string_returns_none(self):
        assert parse_csv_date("", "DMY") is None

    def test_none_returns_none(self):
        assert parse_csv_date(None, "DMY") is None

    def test_invalid_date_returns_none(self):
        assert parse_csv_date("not-a-date", "DMY") is None

    def test_invalid_month_returns_none(self):
        assert parse_csv_date("15/13/2023", "DMY") is None

    def test_two_digit_year_not_accepted(self):
        # 2-digit year parses to a low year number — caller should validate
        result = parse_csv_date("15/06/23", "DMY")
        # Either None or year=23 — both acceptable, just not year=2023
        assert result is None or result.year < 100

    def test_whitespace_stripped(self):
        assert parse_csv_date("  15/06/2023  ", "DMY") == date(2023, 6, 15)

    def test_dmy_eu_format(self):
        """Confirm DMY correctly swaps day and month."""
        assert parse_csv_date("01/12/2022", "DMY") == date(2022, 12, 1)
        assert parse_csv_date("12/01/2022", "MDY") == date(2022, 12, 1)


# ── detect_date_format ────────────────────────────────────────────────────────

class TestDetectDateFormat:
    """
    Per-file date format detection based on unambiguous field values.

    Logic:
      • field[0] > 12  → day is first   → DMY
      • field[1] > 12  → month is first → MDY
      • all fields ≤ 12 → ambiguous     → use fallback
      • conflicting evidence             → ValueError
    """

    def test_detects_dmy_from_large_day(self):
        # "25/01/2024" → field[0]=25 > 12 → DMY
        fmt, src = detect_date_format(["25/01/2024", "20/03/2024"])
        assert fmt == "DMY"
        assert src == "detected"

    def test_detects_mdy_from_large_day_in_position_1(self):
        # "01/25/2024" → field[1]=25 > 12 → MDY
        fmt, src = detect_date_format(["01/25/2024", "03/20/2024"])
        assert fmt == "MDY"
        assert src == "detected"

    def test_fallback_when_all_ambiguous(self):
        # "01/05/2024" → both fields ≤ 12, can't distinguish
        fmt, src = detect_date_format(["01/05/2024", "03/06/2024"], fallback="DMY")
        assert fmt == "DMY"
        assert src == "fallback"

    def test_fallback_mdy_when_all_ambiguous(self):
        fmt, src = detect_date_format(["01/05/2024", "03/06/2024"], fallback="MDY")
        assert fmt == "MDY"
        assert src == "fallback"

    def test_empty_list_returns_fallback(self):
        fmt, src = detect_date_format([], fallback="DMY")
        assert fmt == "DMY"
        assert src == "fallback"

    def test_raises_on_contradictory_evidence(self):
        # "25/01/2024" → DMY; "01/25/2024" → MDY — same file can't be both
        with pytest.raises(ValueError, match="Contradictory"):
            detect_date_format(["25/01/2024", "01/25/2024"])

    def test_dash_delimiters_accepted(self):
        fmt, src = detect_date_format(["25-01-2024"])
        assert fmt == "DMY"
        assert src == "detected"

    def test_ignores_non_date_strings(self):
        # Non-date values (headers, names) should be skipped gracefully
        fmt, src = detect_date_format(["Strategy Name", "", "nan", "25/01/2024"])
        assert fmt == "DMY"
        assert src == "detected"

    def test_single_unambiguous_date_sufficient(self):
        # One date with day=31 is enough to confirm DMY
        fmt, src = detect_date_format(["31/12/2023"])
        assert fmt == "DMY"
        assert src == "detected"

    def test_mixed_ambiguous_and_unambiguous_detects_correctly(self):
        # Ambiguous dates alongside one unambiguous → detected
        fmt, src = detect_date_format(["01/06/2024", "05/06/2024", "25/06/2024"])
        assert fmt == "DMY"
        assert src == "detected"


# ── resolve_oos_dates ─────────────────────────────────────────────────────────

class TestResolveOosDates:
    """
    Three cases from VBA ResolveOOSDates:
        Case 1: cutoff < oos_begin  → oos_end = oos_begin
        Case 2: oos_begin <= cutoff < oos_end  → oos_end = cutoff
        Case 3: cutoff >= oos_end  → unchanged
    """

    OOS_BEGIN = date(2020, 1, 1)
    OOS_END = date(2023, 12, 31)

    def test_no_cutoff_returns_unchanged(self):
        begin, end = resolve_oos_dates(
            self.OOS_BEGIN, self.OOS_END, use_cutoff=False, cutoff_date=None
        )
        assert begin == self.OOS_BEGIN
        assert end == self.OOS_END

    def test_case1_cutoff_before_oos_begin(self):
        """Cutoff before OOS start: clamp end to begin."""
        cutoff = date(2019, 6, 1)
        begin, end = resolve_oos_dates(
            self.OOS_BEGIN, self.OOS_END, use_cutoff=True, cutoff_date=cutoff
        )
        assert begin == self.OOS_BEGIN
        assert end == self.OOS_BEGIN   # clamped to begin

    def test_case2_cutoff_within_oos_period(self):
        """Cutoff within OOS: cap end at cutoff."""
        cutoff = date(2022, 6, 15)
        begin, end = resolve_oos_dates(
            self.OOS_BEGIN, self.OOS_END, use_cutoff=True, cutoff_date=cutoff
        )
        assert begin == self.OOS_BEGIN
        assert end == cutoff

    def test_case3_cutoff_after_oos_end(self):
        """Cutoff after OOS end: unchanged."""
        cutoff = date(2025, 1, 1)
        begin, end = resolve_oos_dates(
            self.OOS_BEGIN, self.OOS_END, use_cutoff=True, cutoff_date=cutoff
        )
        assert begin == self.OOS_BEGIN
        assert end == self.OOS_END

    def test_open_oos_end_capped_at_cutoff(self):
        """Open OOS (no end date): cap at cutoff."""
        cutoff = date(2022, 6, 15)
        begin, end = resolve_oos_dates(
            self.OOS_BEGIN, None, use_cutoff=True, cutoff_date=cutoff
        )
        assert end == cutoff

    def test_none_oos_begin_returns_unchanged(self):
        """No OOS begin: nothing to resolve."""
        begin, end = resolve_oos_dates(
            None, self.OOS_END, use_cutoff=True, cutoff_date=date(2022, 1, 1)
        )
        assert begin is None

    def test_cutoff_exactly_on_oos_begin(self):
        """Cutoff == OOS begin: case 3 (>= begin, not < begin)."""
        cutoff = self.OOS_BEGIN
        begin, end = resolve_oos_dates(
            self.OOS_BEGIN, self.OOS_END, use_cutoff=True, cutoff_date=cutoff
        )
        # cutoff == begin, and begin <= cutoff < end → case 2: cap at cutoff
        assert end == cutoff


# ── cutoff_index ─────────────────────────────────────────────────────────────

class TestCutoffIndex:
    def _make_index(self, *dates):
        return pd.DatetimeIndex([pd.Timestamp(d) for d in dates])

    def test_no_cutoff_returns_last_index(self):
        idx = self._make_index(date(2022, 1, 3), date(2022, 1, 4), date(2022, 1, 5))
        assert cutoff_index(idx, use_cutoff=False, cutoff_date=None) == 2

    def test_cutoff_exactly_on_date(self):
        idx = self._make_index(date(2022, 1, 3), date(2022, 1, 4), date(2022, 1, 5))
        assert cutoff_index(idx, use_cutoff=True, cutoff_date=date(2022, 1, 4)) == 1

    def test_cutoff_between_dates(self):
        idx = self._make_index(date(2022, 1, 3), date(2022, 1, 5), date(2022, 1, 7))
        # Cutoff Jan 6 → last date <= Jan 6 is Jan 5 (index 1)
        assert cutoff_index(idx, use_cutoff=True, cutoff_date=date(2022, 1, 6)) == 1

    def test_cutoff_before_all_dates_returns_minus_one(self):
        idx = self._make_index(date(2022, 1, 3), date(2022, 1, 4))
        assert cutoff_index(idx, use_cutoff=True, cutoff_date=date(2021, 12, 31)) == -1

    def test_cutoff_after_all_dates_returns_last(self):
        idx = self._make_index(date(2022, 1, 3), date(2022, 1, 4))
        assert cutoff_index(idx, use_cutoff=True, cutoff_date=date(2025, 1, 1)) == 1


# ── CME Holiday Calendar ──────────────────────────────────────────────────────

class TestCmeHolidays:
    def test_weekends_are_non_trading(self):
        assert is_non_trading_day(date(2024, 1, 6))   # Saturday
        assert is_non_trading_day(date(2024, 1, 7))   # Sunday

    def test_weekdays_are_trading(self):
        assert not is_non_trading_day(date(2024, 1, 8))  # Monday

    def test_new_years_day_2024(self):
        # Jan 1 2024 is a Monday
        assert is_non_trading_day(date(2024, 1, 1))

    def test_new_years_observed_saturday(self):
        # Jan 1 2022 is a Saturday → observed on Dec 31 2021 (Friday)
        assert is_non_trading_day(date(2021, 12, 31))

    def test_new_years_observed_sunday(self):
        # Jan 1 2023 is a Sunday → observed on Jan 2 2023 (Monday)
        assert is_non_trading_day(date(2023, 1, 2))

    def test_christmas_2024(self):
        # Dec 25 2024 is a Wednesday
        assert is_non_trading_day(date(2024, 12, 25))

    def test_thanksgiving_2024(self):
        # 4th Thursday in November 2024 = Nov 28
        assert is_non_trading_day(date(2024, 11, 28))

    def test_good_friday_2024(self):
        # Easter 2024 = March 31 → Good Friday = March 29
        assert _easter_sunday(2024) == date(2024, 3, 31)
        assert is_non_trading_day(date(2024, 3, 29))

    def test_memorial_day_2024(self):
        # Last Monday in May 2024 = May 27
        assert is_non_trading_day(date(2024, 5, 27))

    def test_labor_day_2024(self):
        # 1st Monday in September 2024 = Sep 2
        assert is_non_trading_day(date(2024, 9, 2))

    def test_independence_day_2024(self):
        # July 4 2024 is a Thursday
        assert is_non_trading_day(date(2024, 7, 4))

    def test_mlk_day_2024(self):
        # 3rd Monday in January 2024 = Jan 15
        assert is_non_trading_day(date(2024, 1, 15))

    def test_presidents_day_2024(self):
        # 3rd Monday in February 2024 = Feb 19
        assert is_non_trading_day(date(2024, 2, 19))

    def test_regular_trading_day(self):
        # A regular Wednesday with no holiday
        assert not is_non_trading_day(date(2024, 3, 6))
