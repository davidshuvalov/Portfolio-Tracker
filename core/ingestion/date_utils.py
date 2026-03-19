"""
Date utilities — mirrors I_MISC.bas date functions exactly.

Key functions:
    parse_csv_date()       — handle DMY / MDY CSV date strings
    detect_date_format()   — per-file format detection from a sample of date strings
    resolve_oos_dates()    — apply cutoff date to OOS period (mirrors ResolveOOSDates)
    cutoff_index()         — last row index <= cutoff date (mirrors EndRowByCutoffSimple)
    is_non_trading_day()   — CME holiday calendar (mirrors IsNonTradingDay)
"""

from __future__ import annotations
from datetime import date, timedelta
from typing import Optional

import pandas as pd


# ── CSV Date Parsing ──────────────────────────────────────────────────────────

def parse_csv_date(date_str: str, date_format: str) -> date | None:
    """
    Parse a date string from a MultiWalk CSV file.

    date_format mirrors the Excel named range DateFormat:
        "DMY"  — day/month/year  (EU, UK, AU)
        "MDY"  — month/day/year  (US)

    Returns None if the string cannot be parsed.
    """
    if not date_str or not isinstance(date_str, str):
        return None
    date_str = date_str.strip()
    if not date_str:
        return None

    parts = date_str.replace("-", "/").split("/")
    if len(parts) != 3:
        return None

    try:
        if date_format == "MDY":
            # US format: month/day/year
            m, d, y = int(parts[0]), int(parts[1]), int(parts[2])
        else:
            # DMY format (EU/UK/AU): day/month/year
            # Mirrors VBA: CDate(dateArray(1) & "/" & dateArray(0) & "/" & dateArray(2))
            d, m, y = int(parts[0]), int(parts[1]), int(parts[2])

        return date(y, m, d)
    except (ValueError, TypeError):
        return None


# ── Per-File Date Format Detection ───────────────────────────────────────────

def detect_date_format(
    date_strings: list[str],
    fallback: str = "DMY",
) -> tuple[str, str]:
    """
    Infer the date format (DMY or MDY) from a sample of raw date strings.

    Strategy — scan each string for unambiguous signals:
      • If the FIRST field is > 12 it cannot be a month → format must be DMY
        (day is first).
      • If the SECOND field is > 12 it cannot be a month → format must be MDY
        (month is first, day is second).
      • If neither ever exceeds 12, all dates are ambiguous (e.g. "01/05/2024")
        — fall back to the caller-supplied default without error.

    Returns:
        (format, source) where:
          format — "DMY" or "MDY"
          source — "detected"  (at least one unambiguous date found)
                   "fallback"  (all dates ambiguous; using fallback)

    Raises:
        ValueError — contradictory evidence (some dates indicate DMY, others MDY).
                     This means the file contains mixed date formats, which is a
                     data quality problem that must be surfaced to the user.
    """
    dmy_evidence: list[str] = []   # date strings that prove DMY
    mdy_evidence: list[str] = []   # date strings that prove MDY

    for raw in date_strings:
        if not raw or not isinstance(raw, str):
            continue
        s = raw.strip().replace("-", "/")
        parts = s.split("/")
        if len(parts) != 3:
            continue
        try:
            p0, p1 = int(parts[0]), int(parts[1])
        except ValueError:
            continue

        if p0 > 12:
            dmy_evidence.append(raw.strip())   # e.g. "25/01/2024" → day=25 → DMY
        if p1 > 12:
            mdy_evidence.append(raw.strip())   # e.g. "01/25/2024" → day=25 → MDY

    if dmy_evidence and mdy_evidence:
        raise ValueError(
            f"Contradictory date formats in file: "
            f"'{dmy_evidence[0]}' indicates DMY but "
            f"'{mdy_evidence[0]}' indicates MDY. "
            f"Check that all dates in this file use the same format."
        )

    if dmy_evidence:
        return "DMY", "detected"
    if mdy_evidence:
        return "MDY", "detected"
    return fallback, "fallback"


def _sample_date_strings(raw: pd.DataFrame, col: int, max_rows: int = 40) -> list[str]:
    """Extract up to max_rows non-empty string values from a DataFrame column."""
    if col >= raw.shape[1]:
        return []
    col_data = raw.iloc[:max_rows, col].dropna().astype(str)
    return [v.strip() for v in col_data if v.strip() and v.strip().lower() not in ("nan", "none", "")]


# ── OOS Date Resolution ───────────────────────────────────────────────────────

def resolve_oos_dates(
    oos_begin: date | None,
    oos_end: date | None,
    use_cutoff: bool,
    cutoff_date: date | None,
) -> tuple[date | None, date | None]:
    """
    Apply cutoff date to an OOS period.
    Mirrors VBA ResolveOOSDates exactly — three cases:

        Case 1: cutoff < oos_begin  → oos_end = oos_begin  (no OOS before cutoff)
        Case 2: oos_begin <= cutoff < oos_end  → oos_end = cutoff
        Case 3: cutoff >= oos_end  → unchanged

    Returns (oos_begin, oos_end) with cutoff applied.
    """
    if not use_cutoff or cutoff_date is None or oos_begin is None:
        return oos_begin, oos_end

    if cutoff_date < oos_begin:
        # Case 1: cutoff before OOS start — clamp end to begin
        return oos_begin, oos_begin

    if oos_end is not None:
        if cutoff_date < oos_end:
            # Case 2: cutoff falls within OOS period — cap end at cutoff
            return oos_begin, cutoff_date
        else:
            # Case 3: cutoff after OOS end — keep original
            return oos_begin, oos_end
    else:
        # OOS end is open (ongoing) — cap at cutoff
        return oos_begin, cutoff_date


# ── Cutoff Index ──────────────────────────────────────────────────────────────

def cutoff_index(
    dates: pd.DatetimeIndex | pd.Series,
    use_cutoff: bool,
    cutoff_date: date | None,
) -> int:
    """
    Return the integer position of the last date <= cutoff_date.
    Mirrors VBA EndRowByCutoffSimple.

    Returns:
        len(dates) - 1  if no cutoff (use all data)
        -1              if cutoff is before all dates (no data)
        position        last index where date <= cutoff_date
    """
    if not use_cutoff or cutoff_date is None:
        return len(dates) - 1

    cutoff_ts = pd.Timestamp(cutoff_date)
    if isinstance(dates, pd.Series):
        dates = pd.DatetimeIndex(dates)

    mask = dates <= cutoff_ts
    if not mask.any():
        return -1

    return int(mask.nonzero()[0][-1])


# ── CME Holiday Calendar ──────────────────────────────────────────────────────

def is_non_trading_day(d: date) -> bool:
    """
    Return True if the date is a weekend or CME holiday.
    Mirrors VBA IsNonTradingDay exactly.

    CME holidays:
        New Year's Day (observed)
        MLK Jr. Day (3rd Monday in January)
        Presidents' Day (3rd Monday in February)
        Good Friday (Easter - 2 days)
        Memorial Day (last Monday in May)
        Independence Day (observed)
        Labor Day (1st Monday in September)
        Thanksgiving (4th Thursday in November)
        Christmas (observed)
    """
    # Weekend
    if d.weekday() >= 5:  # 5=Saturday, 6=Sunday
        return True

    # Check own year and next year (observed New Year's can fall on Dec 31 of prior year)
    return d in _cme_holidays(d.year) or d in _cme_holidays(d.year + 1)


def _cme_holidays(year: int) -> frozenset[date]:
    """Compute the set of CME holidays for a given year."""
    holidays: set[date] = set()

    # New Year's Day (observed)
    ny = date(year, 1, 1)
    holidays.add(_observe(ny))

    # MLK Jr. Day — 3rd Monday in January
    holidays.add(_nth_weekday(year, 1, 0, 3))  # 0=Monday, 3rd

    # Presidents' Day — 3rd Monday in February
    holidays.add(_nth_weekday(year, 2, 0, 3))

    # Good Friday — Easter Sunday - 2 days
    easter = _easter_sunday(year)
    holidays.add(easter - timedelta(days=2))

    # Memorial Day — last Monday in May
    holidays.add(_last_weekday(year, 5, 0))  # 0=Monday

    # Independence Day (observed)
    july4 = date(year, 7, 4)
    holidays.add(_observe(july4))

    # Labor Day — 1st Monday in September
    holidays.add(_nth_weekday(year, 9, 0, 1))

    # Thanksgiving — 4th Thursday in November
    holidays.add(_nth_weekday(year, 11, 3, 4))  # 3=Thursday, 4th

    # Christmas (observed)
    xmas = date(year, 12, 25)
    holidays.add(_observe(xmas))

    return frozenset(holidays)


def _observe(d: date) -> date:
    """
    Return the observed date for a holiday falling on a weekend.
    Saturday → Friday, Sunday → Monday.
    """
    if d.weekday() == 5:  # Saturday
        return d - timedelta(days=1)
    if d.weekday() == 6:  # Sunday
        return d + timedelta(days=1)
    return d


def _nth_weekday(year: int, month: int, weekday: int, n: int) -> date:
    """Return the nth occurrence of weekday (0=Mon) in the given month/year."""
    first = date(year, month, 1)
    # Days until first occurrence of target weekday
    delta = (weekday - first.weekday()) % 7
    first_occurrence = first + timedelta(days=delta)
    return first_occurrence + timedelta(weeks=n - 1)


def _last_weekday(year: int, month: int, weekday: int) -> date:
    """Return the last occurrence of weekday (0=Mon) in the given month/year."""
    if month == 12:
        last = date(year + 1, 1, 1) - timedelta(days=1)
    else:
        last = date(year, month + 1, 1) - timedelta(days=1)
    delta = (last.weekday() - weekday) % 7
    return last - timedelta(days=delta)


def _easter_sunday(year: int) -> date:
    """Compute Easter Sunday using the Anonymous Gregorian algorithm."""
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day = ((h + l - 7 * m + 114) % 31) + 1
    return date(year, month, day)


# ── Trading Day Utilities ─────────────────────────────────────────────────────

def trading_days_between(start: date, end: date) -> int:
    """Count CME trading days between two dates (inclusive)."""
    count = 0
    d = start
    while d <= end:
        if not is_non_trading_day(d):
            count += 1
        d += timedelta(days=1)
    return count


def next_trading_day(d: date) -> date:
    """Return the next CME trading day after d."""
    d = d + timedelta(days=1)
    while is_non_trading_day(d):
        d += timedelta(days=1)
    return d
