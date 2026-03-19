# Portfolio Tracker — Systematic Code Review Audit

_Date: 2026-03-19_
_Scope: Full codebase (VBA v1 + Python v2)_

---

## Executive Summary

This audit reviews the Portfolio Tracker across **five dimensions**: core analytics calculations, data ingestion/portfolio logic, UI/navigation, VBA modules, and test coverage. The codebase is well-structured with 691+ passing tests and clear separation of concerns, but there are several issues that could produce **incorrect financial calculations** in specific scenarios.

**Overall Assessment: 7/10** — Solid architecture, good test coverage (~75%), but critical calculation bugs in walkforward metrics and backtest look-ahead bias need immediate attention.

---

## 1. CRITICAL FINDINGS (Must Fix)

### ~~C1 — Dead Variable in Rule Backtest~~ ✅ FIXED
**File:** `v2/core/analytics/eligibility/rule_backtest.py:166`
**Issue:** `last_eval_month` was computed but unused. On closer review, the inner guards (`future_end >= n_months`) correctly prevent look-ahead — this was NOT a correctness bug, just dead code.
**Fix:** Removed dead variable, added clarifying comment.

### ~~C2 — OOS Annual SD Uses Wrong Sharpe Ratio~~ ✅ FIXED
**File:** `v2/core/ingestion/walkforward_reader.py:389`
**Issue:** `annual_sd_isoos` divided by `sharpe_is` instead of `sharpe_isoos`.
**Fix:** Changed to `sharpe_isoos`.

### ~~C3 — Monte Carlo RNG Is Non-Deterministic~~ ✅ FIXED
**File:** `v2/core/analytics/monte_carlo.py`, `v2/core/config.py`
**Issue:** No seed parameter exposed for MC reproducibility.
**Fix:** Added `seed: int | None` to `MCConfig`; `run_monte_carlo()` calls `np.random.seed()` when set.

### ~~C4 — Drawdown Clamped at 100%~~ ✅ DOCUMENTED
**File:** `v2/core/analytics/monte_carlo.py:72`
**Issue:** Drawdown fraction clamped to 1.0 when equity goes negative.
**Resolution:** This is a deliberate design choice — negative equity is captured by the ruin flag; the fraction stays in [0,1] for percentile stats. Added clarifying comment.

### ~~C5 — Incubation/Quitting Use Row Index Instead of Calendar Days~~ ✅ FIXED
**File:** `v2/core/portfolio/summary.py:609-615, 672-677`
**Issue:** `days_elapsed = i + 1` counted rows (trading days) but compared against calendar-day thresholds (`incubation_days = months * 30.5`). Weekends/holidays caused ~30% undercounting.
**Fix:** Changed to `(ts - first_date).days + 1` in both `_calc_incubation()` and `_calc_quitting_status()`.

---

## 2. HIGH-SEVERITY FINDINGS (Should Fix)

### H1 — Population Std Instead of Sample Std
**Files:** `diversification.py:49`, `leave_one_out.py:169-170`
**Issue:** `series.std()` defaults to `ddof=0` in pandas, underestimating volatility. Financial convention uses `ddof=1`.
**Impact:** Annualized volatility is systematically low. Sharpe ratios appear better than they are.
**Fix:** Use `.std(ddof=1)` consistently.

### H2 — RTD Hardcoded to 10.0 for Low Drawdown
**File:** `v2/core/portfolio/summary.py:282-286`
**Issue:** When max drawdown < $10, RTD is set to 10.0 regardless of actual profit. A strategy with $5 DD and $500 profit gets the same RTD as one with $9 DD and $90 profit.
**Impact:** Optimizer ranking becomes arbitrary for low-drawdown strategies.
**Fix:** Compute actual RTD or use a percentage-based threshold.

### H3 — LOO Silently Ignores Missing Strategies
**File:** `v2/core/analytics/leave_one_out.py:220-222`
**Issue:** If a strategy name isn't found in portfolio columns, LOO returns unchanged base PnL with no warning.
**Impact:** User believes a strategy was removed from the analysis, but it wasn't.
**Fix:** Raise warning or error.

### H4 — Aggregator Crashes on Incomplete Data
**File:** `v2/core/portfolio/aggregator.py:81-88`
**Issue:** No check that strategy columns exist in `daily_m2m` or `closed_trade_pnl` before indexing. Raises `KeyError` on incomplete data.
**Impact:** Portfolio build crashes with unhelpful error message.
**Fix:** Validate column existence and report missing strategies.

### H5 — Eligibility NaN Handling Asymmetry
**File:** `v2/core/portfolio/summary.py:813 vs 834`
**Issue:** Qualifiers treat NaN as "passing" (`| col_data.isna()`), but disqualifiers treat NaN as "passing" via different logic (`.fillna(False)` then negate). Semantics differ subtly.
**Impact:** Edge case where NaN in a metric produces unexpected eligibility outcomes.
**Fix:** Standardize NaN handling with explicit documentation.

### H6 — Margin Position Conflict Not Detected
**File:** `v2/core/analytics/margin.py:76-83`
**Issue:** If a strategy has both long and short values != 0, the first `if` wins (reports LONG). No detection of conflicting position signals.
**Impact:** Margin estimates could be understated if positions are actually hedged.
**Fix:** Add conflict detection and warning.

---

## 3. VBA-SPECIFIC FINDINGS

### V1 — Silent Zero on Missing Dictionary Keys
**Files:** `D_Import_Data.bas:292`, `N_BackTest.bas:427`, `M_Margin_Tracking.bas:314`
**Issue:** When dictionary lookups miss, data defaults to 0 with no warning. Financial data should never silently become zero.
**Impact:** Portfolio equity systematically understated on gap days. No audit trail.

### V2 — Uninitialized `startingEquity` in Backtest
**File:** `N_BackTest.bas:354-363`
**Issue:** `startingEquity` is declared but never set. Default is 0. Drawdown percentage calculations divide by `peakProfit + 0.000001`.
**Impact:** All VBA drawdown percentages are meaningless.

### V3 — Array Indexing Without Bounds Checking
**File:** `D_Import_Data.bas:436, 463-465`
**Issue:** CSV column access assumes exact column positions (e.g., column 2, 3, 4, 6, 7, 8). No header validation.
**Impact:** If CSV structure changes, wrong columns are used silently.

### V4 — Column Mapping Off-by-One
**File:** `N_BackTest.bas:390-391`
**Issue:** Dictionary maps strategy names to Excel column numbers, but uses these as array indices. Array is 1-based but column offset may differ.
**Impact:** Wrong strategy data used in backtest calculations.

### V5 — O(n²) Backtest Loop
**File:** `N_BackTest.bas:420-476`
**Issue:** Triple-nested loop: 5000 days × 50 strategies × 50 lookups = 12.5M iterations.
**Fix:** Use dictionary lookups instead of linear search.

### V6 — 244 Instances of `On Error Resume Next`
**Impact:** Errors are silently swallowed across the entire VBA codebase. Root cause analysis becomes nearly impossible.

### V7 — 60+ Hardcoded Sheet Name References
**Impact:** Renaming any sheet breaks the entire system with no clear error message.

---

## 4. NAVIGATION & UI FINDINGS

### N1 — Dead-End Analytics Pages
**Files:** `_04_Monte_Carlo.py:30`, `_09_Eligibility_Backtest.py:39`, others
**Issue:** Analytics pages are visible on the home page even when portfolio isn't built. Users navigate to pages that immediately `st.stop()` with "Go to Import" message.
**Fix:** Hide analytics links until `portfolio is not None`.

### N2 — Unreachable Migrate Page
**File:** `app.py:371`
**Issue:** `00_Migrate.py` is registered in navigation but never linked from any page.

### N3 — Session State Not Validated on Page Load
**Files:** `_Strategy_Detail.py:73`, `03_Portfolio.py:36`, `_16_Market_Analysis.py:47`
**Issue:** Multiple pages access session state keys that may not exist. Falls back silently or crashes.
**Fix:** Add strict initialization checks with helpful messages.

### N4 — No Confirmation on Destructive Portfolio Actions
**File:** `_15_Portfolio_Optimizer.py:835-857`
**Issue:** "Apply Suggested Portfolio" replaces current portfolio without showing diff or asking confirmation.

### N5 — Date Format Not Validated Against CSV Data
**File:** `00_Inputs.py:124-128`
**Issue:** User selects DMY/MDY format, but no validation that imported CSVs actually use that format. Wrong dates are parsed silently.
**Fix:** Auto-detect format from a sample of rows and warn on mismatch.

### N6 — Market Analysis Silently Drops Missing B&H Strategies
**File:** `_16_Market_Analysis.py:85-88`
**Issue:** If some Buy & Hold strategies are missing from daily M2M, they're silently dropped from analysis without notification.

---

## 5. TEST COVERAGE GAPS

### Coverage: ~75%

| Area | Coverage | Risk |
|------|----------|------|
| Position detection & margin | Excellent | Low |
| ATR calculation & sizing | Very Strong | Low |
| Eligibility rules (70+ rules) | Very Strong | Low |
| Portfolio optimizer workflow | Very Strong | Low |
| Correlations (3 modes) | Strong | Low |
| Monte Carlo simulation | Strong | Low |
| CSV/XLSB import | Strong | Low |
| Folder scanning | Excellent | Low |
| **Walkforward reader** | **None** | **HIGH** |
| **Leave-one-out** | **None** | **HIGH** |
| **Settings I/O** | **None** | **Medium** |
| Integration (end-to-end) | Weak | Medium |
| Export content validation | Weak | Medium |

### Most Important Missing Tests
1. **`test_walkforward_reader.py`** — metric extraction, date parsing, edge cases
2. **`test_leave_one_out.py`** — core LOO algorithm with multi-strategy scenarios
3. **`test_settings_io.py`** — YAML save/load with nested configs
4. **End-to-end portfolio scenario** — real data through full pipeline with constraint verification
5. **Degenerate data tests** — NaN correlation matrices, zero-variance series, negative equity

---

## 6. CROSS-CUTTING CONCERNS

### Inconsistent Tolerances
| Module | Tolerance | Context |
|--------|-----------|---------|
| monte_carlo.py | 1e-9 | Peak drawdown |
| correlations.py | 1e-12 | Std dev zero check |
| rule_backtest.py | 1e-4 | vs_base division |
| walkforward_reader.py | 1e-9 | Division guards |

**Fix:** Define `EPSILON` constant(s) and use globally.

### Inconsistent Error Returns
| Behavior | Where |
|----------|-------|
| Return 0.0 | monte_carlo (Sharpe), diversification (avg_dd) |
| Return NaN | closed_trade_mc |
| Return empty DF | multiple |
| Raise ValueError | atr.py |
| Silently ignore | leave_one_out |

**Fix:** Standardize: NaN for undefined metrics, ValueError for invalid inputs, empty for no data.

### Annualization Inconsistencies
- diversification.py: population std (ddof=0)
- leave_one_out.py: population std (ddof=0)
- monte_carlo.py: no annualization on drawdown
- walkforward_reader.py: extrapolation-based annualization

---

## 7. RECOMMENDATIONS (Prioritized)

### Immediate (This Sprint)
1. Fix look-ahead bias in `rule_backtest.py` (C1)
2. Fix OOS annual SD using wrong Sharpe (C2)
3. Fix incubation day counting (C5)
4. Add seed parameter to Monte Carlo (C3)

### Next Sprint
5. Standardize `ddof=1` across all std calculations (H1)
6. Fix RTD clamping logic (H2)
7. Add column existence validation in aggregator (H4)
8. Write `test_walkforward_reader.py` and `test_leave_one_out.py`

### Backlog
9. Standardize tolerances and error returns
10. Add NaN handling documentation for eligibility
11. Fix VBA `startingEquity` initialization (V2)
12. Improve UI navigation guards (N1, N3)
13. Add date format auto-detection (N5)

---

_Review performed by systematic audit of all 131 Python source files, 25 VBA modules, and 28 test files._
