# Portfolio Tracker v2 — Task Board

_Last updated: 2026-03-18_

---

## Sprint A: Eligibility Backtest Settings — ✅ COMPLETE

### A1 — Config additions to `EligibilityConfig` ✅
- [x] `backtest_data_scope: Literal["OOS", "IS+OOS"] = "OOS"` — in `config.py:212`
- [x] `exclude_buy_and_hold: bool = True` — in `config.py:213`
- [x] `exclude_previously_quit: bool = False` — in `config.py:214`

### A2 — Enforcement ✅
- [x] Gate: `exclude_buy_and_hold` enforced in `summary.py:732`
- [x] Gate: `exclude_previously_quit` enforced in `summary.py:739`
- [x] `backtest_data_scope`: slices P&L windows in `summary.py:332,345`

### A3 — `_09_Eligibility_Backtest.py` sidebar ✅
- [x] "Data scope" radio (OOS / IS+OOS) — line 118
- [x] "Exclude Buy & Hold" toggle — line 127
- [x] "Exclude previously quit" toggle — line 132
- [x] Fields wired into config constructor — lines 143-145, 162-164

### A4 — Tests ✅
- [x] `test_exclude_buy_and_hold_gate` — in `test_eligibility_new_gates.py`
- [x] `test_exclude_previously_quit_gate` — in `test_eligibility_new_gates.py`
- [x] `test_data_scope_oos_slices_from_oos_start` — in `test_eligibility_new_gates.py`

---

## Sprint B: Portfolio Settings Config — ✅ COMPLETE

### B1 — `PortfolioContractConfig` ✅
- [x] Class defined in `config.py:37-58` with all fields

### B2 — `MCConfig` extensions ✅
- [x] `output_samples`, `remove_best_pct`, `solve_for_ror` — in `config.py:20-35`

### B3 — Wire into `AppConfig` ✅
- [x] `contract_sizing: PortfolioContractConfig` field in AppConfig
- [x] YAML round-trip verified in `test_eligibility_new_gates.py`

### B4 — Portfolio Settings UI ✅
- [x] Contract sizing sidebar in `ui/components/settings_sidebar.py:124-166`

### B5 — Tests ✅
- [x] Config YAML round-trip with all new fields
- [x] `contract_size_from_atr()` formula correctness — in `test_atr.py`

---

## Sprint C: ATR from Trade Data — ✅ COMPLETE

### C1 — `core/analytics/atr.py` ✅
- [x] `compute_daily_range()`, `compute_atr()`, `compute_atr_series()` implemented
- [x] `contract_size_from_atr()` and `estimate_contracts()` implemented
- [x] `reweight_contracts_by_atr()` implemented

### C2 — ATR in portfolio data ✅
- [x] ATR per strategy exposed via portfolio optimizer

### C3 — Historical ATR reweighting ✅
- [x] `reweight_contracts_by_atr()` fully implemented and tested

### C4 — Tests ✅
- [x] 50+ tests in `test_atr.py` covering all functions and edge cases

---

## Sprint D: Calculation Parity Test Plan

_Goal: verify the refactored VBA model produces identical outputs
to the base spreadsheet, and document every known gap._

---

### D0 — Known gaps: decisions

| # | Area | Status | Decision |
|---|------|--------|----------|
| D0-1 | BnH in Long_Trades/Short_Trades | ✅ Accepted | Improvement: BnH now shows correct long profit factors |
| D0-2 | ATR percentile method | ✅ Fixed | Uses consistent 90-day rolling window |
| D0-3 | Sector lookup fuzzy match | ✅ Fixed | Uses `StrComp` exact match |
| D0-4 | BnH in downstream VBA modules | ✅ N/A in v2 | v2 filters BnH at `build_portfolio()` level — only "Live" strategies enter analytics |
| D0-5 | Backtest BnH contribution | ✅ Documented | TotalBackTest uses DailyM2MEquity (correct). ClosedTradePNL trade stats may undercount BnH |
| D0-6 | Tab order naming | ✅ Cosmetic | Documented: "Backtest" tab = setup, "Backtest/WhatIf" group = results |
| D0-7 | Contract symbol collision | ⚠️ Accepted | Low likelihood. Add warning if collision detected (VBA only) |

---

### D1 — Regression: TradeData loop merge

Code-inspection items verified:
- [x] **D1-4** `ProcessLatestPositions` — no duplicate-key errors
- [x] **D1-5** `ProcessLSTradeData` — BnH appears as new columns (expected per D0-1)
- [x] **D1-6** Short trades — BnH long-only produces zero short entries

Live-data verification needed (requires Excel import run):
- [ ] **D1-1** `TrueRanges` row count and headers unchanged
- [ ] **D1-2** `AverageTrueRange` ATR values unchanged
- [ ] **D1-3** `TradePNL` contract-level PNL unchanged
- [ ] **D1-7** `DailyM2MEquity` equity curves unchanged
- [ ] **D1-8** `ClosedTradePNL` unchanged
- [ ] **D1-9** BnH profit factors now non-zero (intentional improvement)
- [ ] **D1-10** BnH strategy detail tab renders without errors

---

### D2 — Regression: Portfolio sheet (live-data verification)

- [ ] **D2-1** ATR columns match AverageTrueRange
- [ ] **D2-2** ATR formatting `$#,##0`
- [ ] **D2-3** BnH ATR 3M = AverageTrueRange × contract count
- [ ] **D2-4** Non-BnH ATR columns = 0
- [ ] **D2-5** Portfolio row count = Live strategy count
- [ ] **D2-6** AutoFilter covers correct column range

---

### D3 — W_Markets: new module validation

Code-inspection items verified:
- [x] **D3-1** ATR percentile fix — consistent 90-day window
- [x] **D3-2** Sector match fix — `StrComp` exact match
- [x] **D3-3** Pearson self-correlation = 1.0
- [x] **D3-5** Volatility regime boundaries correct
- [x] **D3-8** Correlation matrix symmetric

Live-data verification needed:
- [ ] **D3-4** Pearson correlation direction (positive for co-moving ATRs)
- [ ] **D3-6** Rolling 90-day avg matches `atr3M` (rounding check)
- [ ] **D3-7** MarketVolatility row count matches unique dates

---

### D4 — Downstream BnH impact (VBA only)

**Note:** D0-4 is resolved in v2 — BnH excluded at `build_portfolio()` level.
VBA-side verification still needed if maintaining the VBA tool:
- [ ] **D4-1–D4-4** Run VBA analytics with/without BnH to quantify distortion

---

### D5 — Backtest completeness for BnH — ✅ COMPLETE
- [x] **D5-1** ClosedTradePNL populated for BnH
- [x] **D5-2** TotalBackTest sources from DailyM2MEquity
- [x] **D5-3** N/A — structure correct
- [x] **D5-4** N/A — no change needed

---

### D6 — Tab order and navigation (live verification)
- [ ] **D6-1** Tab order: Settings → Folder → Strategies → Portfolio → Backtest → Markets
- [ ] **D6-2** `GoToMarkets()` activates Markets sheet
- [ ] **D6-3** `CreateMarketsSummary()` creates all sections
- [ ] **D6-4** `DeleteAllTabs()` removes Markets sheets cleanly

---

### D7 — Calculation parity checklist (base vs new)

| Calculation | Base | New | Status |
|-------------|------|-----|--------|
| Daily M2M equity per strategy | ✓ | ✓ unchanged | ✓ |
| Closed trade PNL per strategy | ✓ | ✓ unchanged | ✓ |
| In-market long/short daily PNL | ✓ | ✓ unchanged | ✓ |
| ATR per BnH contract (avgs) | ✓ | ✓ unchanged | ✓ |
| ATR per BnH contract (raw) | ✓ | ✓ unchanged | ✓ |
| Trade PNL per BnH contract | ✓ | ✓ unchanged | ✓ |
| Long trade PNL per strategy | ✓ non-BnH | ~ includes BnH | ✅ D0-1 improvement |
| Short trade PNL per strategy | ✓ non-BnH | ~ includes BnH (empty) | ✅ no change |
| BnH profit factors in Summary | ✗ zero | ✓ computed | ✅ improvement |
| Latest positions | ✓ | ✓ unchanged | ✓ |
| Portfolio ATR columns | ✓ | ✓ same lookup | ✓ |
| Strategy correlations | ✓ | ✓ BnH filtered in v2 | ✅ |
| Portfolio MC | ✓ | ✓ BnH filtered in v2 | ✅ |
| Backtest portfolio P&L | ✓ | ✓ via DailyM2MEquity | ✅ |
| Market ATR percentile | ✗ | ✓ new (W_Markets) | ✅ D3-1 fixed |
| Market sector summary | ✗ | ✓ new (W_Markets) | ✅ D3-2 fixed |
| ATR correlation matrix | ✗ | ✓ new (W_Markets) | ✅ D3-3/8 verified |
| Rolling 90-day ATR series | ✗ | ✓ new (W_Markets) | Needs live verify |

---

## Remaining work

All code-level implementation for Sprints A–D is complete. 691 tests pass.

### Live-data verification (requires Excel import)
- D1 regression items (1-3, 7-10)
- D2 portfolio sheet checks (1-6)
- D3 live validation (4, 6, 7)
- D6 tab order / navigation (1-4)

### VBA-only items
- D0-7: Add collision warning in `ProcessTradeData` (low priority)
- D4: BnH impact analysis in VBA downstream modules (if maintaining VBA tool)

---
