"""
Futures contract registry — maps standard symbols to their micro/mini equivalents.

Used by the optimizer to determine the effective minimum tradeable fraction for
a given symbol.  When a micro or mini contract is available, the optimizer can
keep a strategy that would otherwise be excluded for being too small to trade
with a full (or E-mini) contract.

Usage example::

    from core.data.contract_registry import CONTRACT_REGISTRY, effective_min_fraction

    # NQ has MNQ (0.1×), so smallest tradeable unit = 0.01 NQ (= 0.1 MNQ)
    frac = effective_min_fraction("NQ", base_fraction=0.1)  # → 0.01

    # ZN has no micro, so min fraction stays at 0.1
    frac = effective_min_fraction("ZN", base_fraction=0.1)  # → 0.1

Micro/mini ratios express contract size relative to the standard contract:
    MNQ = 0.1 × NQ  →  micro_ratio = 0.1
    MES = 0.1 × ES  →  micro_ratio = 0.1
    XC  = 0.2 × ZC  →  mini_ratio  = 0.2
    MBT = 0.02 × BTC →  micro_ratio = 0.02  (0.1 BTC / 5 BTC)
"""

from __future__ import annotations

from dataclasses import dataclass


# ── Data model ─────────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class ContractFamily:
    """
    A futures instrument and its smaller-denomination variants.

    ``micro_ratio`` and ``mini_ratio`` are expressed as fractions of the
    standard contract (e.g., MNQ = 0.1 × NQ → micro_ratio = 0.1).
    """
    symbol: str           # Standard (or E-mini) symbol used in the portfolio
    name: str             # Human-readable description
    sector: str           # Sector group for margin caps
    exchange: str         # Primary exchange

    micro_symbol: str | None = None   # CME "Micro" variant symbol
    micro_ratio: float | None = None  # micro / standard contract size

    mini_symbol: str | None = None    # "Mini" variant symbol (where applicable)
    mini_ratio: float | None = None   # mini / standard contract size

    @property
    def smallest_unit(self) -> float:
        """
        Smallest tradeable fraction expressed as a multiple of the standard
        contract.  Returns 1.0 when no micro/mini is available (no granularity
        gain).
        """
        ratios = [r for r in (self.micro_ratio, self.mini_ratio) if r]
        return min(ratios) if ratios else 1.0

    @property
    def smallest_symbol(self) -> str:
        """Symbol of the smallest available contract variant."""
        candidates = []
        if self.micro_ratio is not None and self.micro_symbol:
            candidates.append((self.micro_ratio, self.micro_symbol))
        if self.mini_ratio is not None and self.mini_symbol:
            candidates.append((self.mini_ratio, self.mini_symbol))
        return min(candidates)[1] if candidates else self.symbol

    def has_smaller_contract(self) -> bool:
        return bool(self.micro_symbol or self.mini_symbol)


# ── Registry ───────────────────────────────────────────────────────────────────

CONTRACT_REGISTRY: dict[str, ContractFamily] = {

    # ── Equity Indices ─────────────────────────────────────────────────────────
    # CME Globex — E-mini series are the "standard" for systematic traders;
    # Micro E-mini = 1/10th the notional value.
    "ES":  ContractFamily("ES",  "E-mini S&P 500",              "Index",       "CME",   micro_symbol="MES",  micro_ratio=0.1),
    "NQ":  ContractFamily("NQ",  "E-mini Nasdaq-100",           "Index",       "CME",   micro_symbol="MNQ",  micro_ratio=0.1),
    "RTY": ContractFamily("RTY", "E-mini Russell 2000",         "Index",       "CME",   micro_symbol="M2K",  micro_ratio=0.1),
    "YM":  ContractFamily("YM",  "E-mini Dow Jones Industrial", "Index",       "CBOT",  micro_symbol="MYM",  micro_ratio=0.1),
    "EMD": ContractFamily("EMD", "E-mini S&P MidCap 400",       "Index",       "CME"),
    "NIY": ContractFamily("NIY", "Nikkei 225 (USD)",            "Index",       "CME"),
    "NK":  ContractFamily("NK",  "Nikkei 225 (JPY)",            "Index",       "OSE"),
    "FESX":ContractFamily("FESX","Euro Stoxx 50",               "Index",       "EUREX"),
    "FDAX":ContractFamily("FDAX","DAX",                         "Index",       "EUREX", micro_symbol="FDXS",  micro_ratio=0.04, mini_symbol="FDXM",  mini_ratio=0.2),   # FDXM=Mini(1/5), FDXS=Micro(1/25)
    "FSMI":ContractFamily("FSMI","SMI",                         "Index",       "EUREX"),

    # ── Energy ─────────────────────────────────────────────────────────────────
    # NYMEX / ICE
    "CL":  ContractFamily("CL",  "Crude Oil WTI",               "Energy",      "NYMEX", micro_symbol="MCL",  micro_ratio=0.1),
    "NG":  ContractFamily("NG",  "Natural Gas",                  "Energy",      "NYMEX"),
    "RB":  ContractFamily("RB",  "RBOB Gasoline",                "Energy",      "NYMEX"),
    "HO":  ContractFamily("HO",  "NY Harbor ULSD (Heating Oil)", "Energy",      "NYMEX"),
    "BRN": ContractFamily("BRN", "Brent Crude Oil",              "Energy",      "ICE"),
    "QM":  ContractFamily("QM",  "E-mini Crude Oil",             "Energy",      "NYMEX"),   # QM = 0.5 × CL already
    "QG":  ContractFamily("QG",  "E-mini Natural Gas",           "Energy",      "NYMEX"),

    # ── Metals ─────────────────────────────────────────────────────────────────
    # COMEX / NYMEX
    "GC":  ContractFamily("GC",  "Gold (100 troy oz)",           "Metals",      "COMEX", micro_symbol="MGC",  micro_ratio=0.1),
    "SI":  ContractFamily("SI",  "Silver (5,000 troy oz)",       "Metals",      "COMEX", micro_symbol="SIL",  micro_ratio=0.2),   # SIL = 1,000 oz
    "HG":  ContractFamily("HG",  "Copper (25,000 lbs)",          "Metals",      "COMEX", micro_symbol="MHG",  micro_ratio=0.1),
    "PL":  ContractFamily("PL",  "Platinum (50 troy oz)",        "Metals",      "NYMEX"),
    "PA":  ContractFamily("PA",  "Palladium (100 troy oz)",      "Metals",      "NYMEX"),

    # ── Agricultural ───────────────────────────────────────────────────────────
    # CBOT — XC/XK/XW are the "Mini" (1,000-bushel) contracts = 0.2 × standard
    # MZC/MZS/MZW are the new Micro (500-bushel) contracts = 0.1 × standard (launched Feb 2025)
    "ZC":  ContractFamily("ZC",  "Corn (5,000 bu)",              "Agriculture", "CBOT",  micro_symbol="MZC", micro_ratio=0.1, mini_symbol="XC",  mini_ratio=0.2),
    "ZS":  ContractFamily("ZS",  "Soybeans (5,000 bu)",          "Agriculture", "CBOT",  micro_symbol="MZS", micro_ratio=0.1, mini_symbol="XK",  mini_ratio=0.2),
    "ZW":  ContractFamily("ZW",  "Wheat SRW (5,000 bu)",         "Agriculture", "CBOT",  micro_symbol="MZW", micro_ratio=0.1, mini_symbol="XW",  mini_ratio=0.2),
    "ZM":  ContractFamily("ZM",  "Soybean Meal (100 short tons)","Agriculture", "CBOT"),
    "ZL":  ContractFamily("ZL",  "Soybean Oil (60,000 lbs)",     "Agriculture", "CBOT"),
    "LE":  ContractFamily("LE",  "Live Cattle (40,000 lbs)",     "Agriculture", "CME"),
    "HE":  ContractFamily("HE",  "Lean Hogs (40,000 lbs)",       "Agriculture", "CME"),
    "GF":  ContractFamily("GF",  "Feeder Cattle (50,000 lbs)",   "Agriculture", "CME"),
    "CC":  ContractFamily("CC",  "Cocoa (10 metric tons)",       "Agriculture", "ICE"),
    "KC":  ContractFamily("KC",  "Coffee C (37,500 lbs)",        "Agriculture", "ICE"),
    "CT":  ContractFamily("CT",  "Cotton #2 (50,000 lbs)",       "Agriculture", "ICE"),
    "SB":  ContractFamily("SB",  "Sugar #11 (112,000 lbs)",      "Agriculture", "ICE"),
    "OJ":  ContractFamily("OJ",  "Orange Juice (15,000 lbs)",    "Agriculture", "ICE"),
    "LB":  ContractFamily("LB",  "Random Length Lumber",         "Agriculture", "CME"),
    "RS":  ContractFamily("RS",  "Canola",                       "Agriculture", "ICE"),

    # ── Interest Rates / Fixed Income ──────────────────────────────────────────
    # CBOT Treasuries
    "ZB":  ContractFamily("ZB",  "30-Year U.S. T-Bond",          "Rates",       "CBOT"),
    "UB":  ContractFamily("UB",  "Ultra T-Bond",                 "Rates",       "CBOT",  micro_symbol="MWN",  micro_ratio=0.1),
    "ZN":  ContractFamily("ZN",  "10-Year U.S. T-Note",          "Rates",       "CBOT"),
    "TN":  ContractFamily("TN",  "Ultra 10-Year T-Note",         "Rates",       "CBOT",  micro_symbol="MTN",  micro_ratio=0.1),
    "ZF":  ContractFamily("ZF",  "5-Year U.S. T-Note",           "Rates",       "CBOT"),
    "ZT":  ContractFamily("ZT",  "2-Year U.S. T-Note",           "Rates",       "CBOT"),
    "SR3": ContractFamily("SR3", "SOFR (3-Month)",               "Rates",       "CME"),
    "GE":  ContractFamily("GE",  "Eurodollar (3-Month)",         "Rates",       "CME"),   # Largely replaced by SR3
    # Eurex Rates
    "FGBL":ContractFamily("FGBL","Euro Bund (10Y)",              "Rates",       "EUREX"),
    "FGBM":ContractFamily("FGBM","Euro Bobl (5Y)",               "Rates",       "EUREX"),
    "FGBS":ContractFamily("FGBS","Euro Schatz (2Y)",             "Rates",       "EUREX"),
    "FGBX":ContractFamily("FGBX","Euro Buxl (30Y)",              "Rates",       "EUREX"),

    # ── Currencies ─────────────────────────────────────────────────────────────
    # CME — Micro FX = 1/10th the standard contract
    "6E":  ContractFamily("6E",  "Euro FX (125,000 EUR)",        "Currencies",  "CME",   micro_symbol="M6E",  micro_ratio=0.1),
    "6J":  ContractFamily("6J",  "Japanese Yen (12.5M JPY)",     "Currencies",  "CME",   micro_symbol="M6J",  micro_ratio=0.1),
    "6B":  ContractFamily("6B",  "British Pound (62,500 GBP)",   "Currencies",  "CME",   micro_symbol="M6B",  micro_ratio=0.1),
    "6A":  ContractFamily("6A",  "Australian Dollar (100K AUD)", "Currencies",  "CME",   micro_symbol="M6A",  micro_ratio=0.1),
    "6C":  ContractFamily("6C",  "Canadian Dollar (100K CAD)",   "Currencies",  "CME",   micro_symbol="MCD",  micro_ratio=0.1),   # MCD=CAD/USD; M6C is the inverse USD/CAD product
    "6S":  ContractFamily("6S",  "Swiss Franc (125,000 CHF)",    "Currencies",  "CME",   micro_symbol="MSF",  micro_ratio=0.1),   # MSF=CHF/USD; M6S is the inverse USD/CHF product
    "6N":  ContractFamily("6N",  "New Zealand Dollar (100K NZD)","Currencies",  "CME",   micro_symbol="M6N",  micro_ratio=0.1),
    "6R":  ContractFamily("6R",  "Russian Ruble",                "Currencies",  "CME"),
    "DX":  ContractFamily("DX",  "U.S. Dollar Index",            "Currencies",  "ICE"),

    # Kansas City Hard Red Winter Wheat — mini MKC (1,000 bu = 0.2×), no micro
    "KE":  ContractFamily("KE",  "Hard Red Winter Wheat (5,000 bu)", "Agriculture", "CBOT",  mini_symbol="MKC",  mini_ratio=0.2),

    # ── Volatility ─────────────────────────────────────────────────────────────
    "VX":  ContractFamily("VX",  "CBOE VIX Futures",             "Volatility",  "CBOE"),

    # ── Crypto ─────────────────────────────────────────────────────────────────
    # BTC futures = 5 BTC/contract; MBT (Micro) = 0.10 BTC → ratio = 0.10/5 = 0.02
    # ETH futures = 50 ETH/contract; MET (Micro) = 0.10 ETH → ratio = 0.10/50 = 0.002
    "BTC": ContractFamily("BTC", "Bitcoin (5 BTC)",              "Crypto",      "CME",   micro_symbol="MBT",  micro_ratio=0.02),
    "ETH": ContractFamily("ETH", "Ether (50 ETH)",               "Crypto",      "CME",   micro_symbol="MET",  micro_ratio=0.002),
}


# ── Reverse lookups ────────────────────────────────────────────────────────────

# Maps micro/mini symbol → parent standard symbol
_VARIANT_TO_STANDARD: dict[str, str] = {}
for _fam in CONTRACT_REGISTRY.values():
    if _fam.micro_symbol:
        _VARIANT_TO_STANDARD[_fam.micro_symbol] = _fam.symbol
    if _fam.mini_symbol:
        _VARIANT_TO_STANDARD[_fam.mini_symbol] = _fam.symbol

# Maps TradeStation legacy symbol → standard CME symbol
# These are the two-letter codes TS uses vs. the CME "6X" / "Z_" convention.
_TS_ALIASES: dict[str, str] = {
    # Interest Rates
    "TU": "ZT",   # 2-Year T-Note
    "FV": "ZF",   # 5-Year T-Note
    "TY": "ZN",   # 10-Year T-Note
    "US": "ZB",   # 30-Year T-Bond
    # Currencies
    "EC":  "6E",  # Euro FX
    "AD":  "6A",  # Australian Dollar
    "JY":  "6J",  # Japanese Yen
    "BP":  "6B",  # British Pound
    "CD":  "6C",  # Canadian Dollar
    "SF":  "6S",  # Swiss Franc
    "NE1": "6N",  # New Zealand Dollar
    # Agriculture (single-letter CBOT codes)
    "C":  "ZC",   # Corn
    "S":  "ZS",   # Soybeans
    "W":  "ZW",   # Chicago SRW Wheat
    "KW": "KE",   # Hard Red Winter Wheat
}


# ── Public helpers ─────────────────────────────────────────────────────────────

def get_family(symbol: str) -> ContractFamily | None:
    """
    Return the :class:`ContractFamily` for *symbol*.

    Accepts standard symbols (``"NQ"``), micro/mini symbols (``"MNQ"``),
    and TradeStation legacy aliases (``"EC"`` → ``6E`` family).
    Returns ``None`` when the symbol is unknown.
    """
    if symbol in CONTRACT_REGISTRY:
        return CONTRACT_REGISTRY[symbol]
    parent = _VARIANT_TO_STANDARD.get(symbol)
    if parent:
        return CONTRACT_REGISTRY[parent]
    canonical = _TS_ALIASES.get(symbol)
    if canonical:
        return CONTRACT_REGISTRY.get(canonical)
    return None


def effective_min_fraction(symbol: str, base_fraction: float = 0.1) -> float:
    """
    Smallest tradeable increment for *symbol* expressed as a multiple of the
    standard contract, given that ``base_fraction`` is the minimum number of
    micro/mini contracts you're willing to hold.

    If no micro/mini is available the standard ``base_fraction`` is returned
    unchanged.

    Examples::

        # NQ with MNQ (0.1×) — minimum 0.1 MNQ = 0.01 NQ
        effective_min_fraction("NQ", 0.1)   # → 0.01

        # ZC with XC (0.2×) — minimum 0.1 XC = 0.02 ZC
        effective_min_fraction("ZC", 0.1)   # → 0.02

        # ZN has no micro — minimum stays 0.1 standard ZN
        effective_min_fraction("ZN", 0.1)   # → 0.1
    """
    family = CONTRACT_REGISTRY.get(symbol)
    if family is None or not family.has_smaller_contract():
        return base_fraction
    return base_fraction * family.smallest_unit
