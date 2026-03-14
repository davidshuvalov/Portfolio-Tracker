# Portfolio Tracker v2 — Task Board

_Last updated: 2026-03-14_

---

## Status Legend
- `[x]` Complete and verified
- `[ ]` Not yet started
- `[~]` Partial / in progress

---

## Pre-Phase — Golden Dataset Capture

> Run v1.24 on anonymised MultiWalk folders; export all metrics as JSON
> regression fixtures before writing any Python.

- [ ] Run v1.24 and export Summary, Portfolio, MC, Correlations, Eligibility
      results as JSON fixtures
- [ ] Store under `tests/fixtures/golden/`
- [ ] Wire fixtures into integration test suite as regression baseline

_Note: Development proceeded without golden fixtures. Phase 2 integration
tests validate against v1.24 output manually; full fixture-based regression
is deferred to Phase 7._

---

## Phase 1 — Data Foundation ✅

- [x] `core/ingestion/folder_scanner.py` — scan MultiWalk folders, validate CSVs
- [x] `core/ingestion/csv_importer.py` — parse EquityData + TradeData CSVs
- [x] `core/ingestion/date_utils.py` — IS/OOS period logic, cutoff dates, trading calendar
- [x] `core/ingestion/xlsb_importer.py` — one-time v1.24 Strategies tab importer
- [x] `core/config.py` + `core/data_types.py`
- [x] `ui/pages/00_Migrate.py` — one-time xlsb import wizard (shown only on first run)
- [x] `ui/pages/01_Import.py` — folder picker, import progress, data preview table
- [x] Unit tests for all ingestion code

---

## Phase 2 — Strategies Config + Portfolio ✅

- [x] `core/portfolio/strategies.py` — load/save strategies config (YAML backend)
- [x] `core/portfolio/aggregator.py` — aggregate strategies into portfolio
- [x] `core/portfolio/summary.py` — 80+ metrics per strategy
- [x] `ui/pages/02_Strategies.py` — editable `st.data_editor` table for strategy config
- [x] `ui/pages/03_Portfolio.py` — portfolio summary with filter/sort
- [x] Integration test: metrics cross-checked against v1.24

---

## Phase 3 — Monte Carlo ✅

- [x] `core/analytics/monte_carlo.py` — Numba JIT inner loop + iterative solver
- [x] `ui/pages/04_Monte_Carlo.py`
- [x] Regression test vs v1.24 (±0.5% tolerance; RNG difference documented)

---

## Phase 4 — Correlations + Diversification ✅

- [x] `core/analytics/correlations.py` — 3 modes (normal, negative, drawdown)
- [x] `core/analytics/diversification.py` — randomised + greedy combinators
- [x] `ui/pages/05_Correlations.py`
- [x] `ui/pages/06_Diversification.py`

---

## Phase 5 — Eligibility Backtest ✅

- [x] `core/analytics/eligibility/rules.py` — all rule types as pure functions
- [x] `core/analytics/eligibility/rule_backtest.py` — walk-forward rule statistics
- [x] `core/analytics/eligibility/portfolio_backtest.py` — portfolio construction backtest
- [x] `ui/pages/09_Eligibility_Backtest.py` — Tab 1 (rule stats) + Tab 2 (portfolio backtest)
- [x] Unit tests: each rule type with synthetic data, look-ahead bias validation

---

## Phase 6 — Advanced Analytics ✅

_Delivered: Leave One Out, Backtest, Margin Tracking, Position Check_

- [x] `core/analytics/leave_one_out.py`
- [x] `core/analytics/backtest.py`
- [x] `core/analytics/margin.py`
      — `PositionStatus` enum, `detect_positions`, `get_strategy_position_table`,
      `net_position_by_symbol`, `compute_daily_margin`, `margin_by_symbol`,
      `margin_by_sector`, `margin_summary_stats`
- [x] `ui/pages/07_Leave_One_Out.py`
- [x] `ui/pages/08_Backtest.py`
- [x] `ui/pages/10_Margin_Tracking.py`
      — editable symbol→margin config, total/symbol/sector area charts,
      activity days, monthly peak heatmap
- [x] `ui/pages/11_Position_Check.py`
      — colour-coded Long/Short/Flat table, net position by symbol + bar chart,
      data freshness alerts
- [x] Unit tests for margin analytics (27 tests)

_Suite: 274 tests, all passing._

---

## Phase 7 — Licensing + Packaging

- [ ] `core/licensing/license_manager.py` — MultiWalk DLL call via ctypes + subscription API check
- [ ] Subscription API server (minimal FastAPI, deployed to Railway/Render)
- [ ] Remove `DEV_MODE` bypass; gate all pages behind license check
- [ ] `core/reporting/excel_export.py` — optional .xlsx export
- [ ] Settings export/import (replaces v1.24 Q module)
- [ ] PyInstaller packaging + code signing
- [ ] Auto-update check on startup
- [ ] End-to-end regression test suite (golden fixtures)

---

## Deferred / Nice-to-Have

- [ ] Golden fixture capture (see Pre-Phase above) — needed for Phase 7 regression suite
- [ ] `core/portfolio/summary.py` edge cases: single-strategy portfolio, all-flat periods
- [ ] Streaming progress bar for large imports (1000+ strategies)
- [ ] Dark-mode compatible chart colours
- [ ] Export any page to PDF via headless browser

---

## Page Inventory (12 pages)

| # | File | Status | VBA mirror |
|---|------|--------|------------|
| 00 | `00_Migrate.py` | ✅ | xlsb import wizard |
| 01 | `01_Import.py` | ✅ | D_Import_Data |
| 02 | `02_Strategies.py` | ✅ | G_Create_Strategy_Tab |
| 03 | `03_Portfolio.py` | ✅ | Portfolio sheet |
| 04 | `04_Monte_Carlo.py` | ✅ | K_MonteCarlo |
| 05 | `05_Correlations.py` | ✅ | L_Correlations |
| 06 | `06_Diversification.py` | ✅ | T_Diversificator |
| 07 | `07_Leave_One_Out.py` | ✅ | S_LeaveOneOut |
| 08 | `08_Backtest.py` | ✅ | N_BackTest |
| 09 | `09_Eligibility_Backtest.py` | ✅ | U_BackTest_Eligibility |
| 10 | `10_Margin_Tracking.py` | ✅ | M_Margin_Tracking |
| 11 | `11_Position_Check.py` | ✅ | V_PositionCheck |

---

## Analytics Module Inventory

| Module | Tests | Status |
|--------|-------|--------|
| `core/ingestion/folder_scanner.py` | ✅ | done |
| `core/ingestion/csv_importer.py` | ✅ | done |
| `core/ingestion/date_utils.py` | ✅ | done |
| `core/ingestion/xlsb_importer.py` | ✅ | done |
| `core/ingestion/walkforward_reader.py` | ✅ | done |
| `core/portfolio/strategies.py` | ✅ | done |
| `core/portfolio/aggregator.py` | ✅ | done |
| `core/portfolio/summary.py` | ✅ | done |
| `core/analytics/monte_carlo.py` | ✅ | done |
| `core/analytics/correlations.py` | ✅ | done |
| `core/analytics/diversification.py` | ✅ | done |
| `core/analytics/leave_one_out.py` | ✅ | done |
| `core/analytics/backtest.py` | ✅ | done |
| `core/analytics/margin.py` | ✅ | done |
| `core/analytics/eligibility/rules.py` | ✅ | done |
| `core/analytics/eligibility/rule_backtest.py` | ✅ | done |
| `core/analytics/eligibility/portfolio_backtest.py` | ✅ | done |
