# Portfolio Tracker v2 — Task Board

_Last updated: 2026-03-16_

---

## Sprint A: Eligibility Backtest Settings

### Existing (no change needed)
- `EligibilityConfig.status_include: list[str]` — multi-status filter ✓
- `EligibilityConfig.oos_dd_vs_is_cap: float` — OOS DD threshold ✓
- `EligibilityConfig.enable_sector_analysis / enable_symbol_analysis` ✓

### A1 — Config additions to `EligibilityConfig`
- [ ] `backtest_data_scope: Literal["OOS", "IS+OOS"] = "OOS"` — IS data included in P&L lookbacks
- [ ] `exclude_buy_and_hold: bool = True` — auto-exclude strategies with `buy_and_hold_status`
- [ ] `exclude_previously_quit: bool = False` — exclude any strategy that ever reached Quit status

### A2 — `apply_eligibility_rules()` enforcement
- [ ] Gate: if `exclude_buy_and_hold`, skip strategies where `strategy.is_buy_and_hold`
- [ ] Gate: if `exclude_previously_quit`, skip strategies where quitting_status == "Quit"
- [ ] `backtest_data_scope`: when `OOS`, slice P&L series from `oos_start` date; when `IS+OOS` use full series

### A3 — `_09_Eligibility_Backtest.py` sidebar
- [ ] Add "Data scope" radio: OOS / IS+OOS
- [ ] Add "Exclude Buy & Hold" toggle
- [ ] Add "Exclude previously quit" toggle
- [ ] Pass new fields into `EligibilityConfig` constructor

### A4 — Tests
- [ ] `test_exclude_buy_and_hold_gate`
- [ ] `test_exclude_previously_quit_gate`
- [ ] `test_data_scope_oos_slices_from_oos_start`

---

## Sprint B: Portfolio Settings Config

### B1 — New `PortfolioContractConfig` class
Fields:
```
starting_equity: float = 705_000
use_percentage: bool = True
cease_type: Literal["Percentage","Dollar"] = "Percentage"
cease_trading_threshold: float = 0.25
reweight_on_atr: bool = True
reweight_index_contracts_only: bool = True
contract_margin_multiple: float = 0.50
contract_ratio_margin_atr: float = 0.50
contract_size_pct_equity: float = 0.01
atr_window: Literal["ATR Last 3 Months","ATR Last 6 Months","ATR Last 12 Months"] = "ATR Last 3 Months"
```

### B2 — Extend `MCConfig`
- [ ] `output_samples: int = 50`
- [ ] `remove_best_pct: float = 0.02`
- [ ] `solve_for_ror: bool = False`

### B3 — Wire into `AppConfig`
- [ ] Add `contract_sizing: PortfolioContractConfig` field
- [ ] YAML round-trip verified

### B4 — Portfolio Settings UI
- [ ] New section in `12_Settings.py`: Portfolio Settings + Contract Sizing
- [ ] Starting equity (shown when `solve_for_ror = False`)
- [ ] Cease type + threshold
- [ ] ATR window, margin multiple, ratio, % equity inputs
- [ ] Re-weight toggles
- [ ] Save persists to config

### B5 — Tests
- [ ] Config YAML round-trip with all new fields
- [ ] `contract_size_from_atr()` formula correctness

---

## Sprint C: ATR from Trade Data

### Key insight
No OHLC needed. `daily_range = MFE + MAE` per trade (both positive dollar values at 1 contract).
Rolling mean over N trading days = dollar ATR.

### C1 — `core/analytics/atr.py` (new file)
```
ATR_WINDOWS = {"ATR Last 3 Months": 63, "ATR Last 6 Months": 126, "ATR Last 12 Months": 252}

compute_daily_range(trades_df) -> pd.DataFrame
  # trades_df cols: strategy, date, pnl, mae, mfe
  # daily_range = mfe + mae per trade; group by (strategy, date) → sum
  # returns DataFrame(index=DatetimeIndex, columns=strategy_names)

compute_atr(trades_df, window="ATR Last 3 Months") -> pd.Series
  # rolling(window_days).mean() of daily_range per strategy
  # returns latest ATR per strategy

contract_size_from_atr(equity, contract_size_pct, atr_dollars, margin, ratio) -> int
  # dollar_risk = atr * ratio + margin * (1 - ratio)
  # return floor(equity * contract_size_pct / dollar_risk)
```

### C2 — ATR in portfolio data
- [ ] ATR per strategy exposed in `03_Portfolio.py` as "ATR (current $)" column
- [ ] Computed lazily when trade data is available

### C3 — Historical ATR reweighting for backtest
- [ ] `reweight_contracts_by_atr(daily_contracts, atr_series, current_atr) -> pd.DataFrame`
  - At each date t: `adj = floor(base × current_atr / atr_t)`
  - Only when `config.contract_sizing.reweight_on_atr = True`
- [ ] Wired into `_08_Backtest.py` + `_09_Eligibility_Backtest.py`

### C4 — Tests
- [ ] `test_daily_range_sums_mfe_mae`
- [ ] `test_atr_rolling_window_length`
- [ ] `test_contract_size_from_atr_formula`
- [ ] `test_contract_size_falls_back_to_margin_when_atr_zero`
- [ ] `test_reweight_scales_inversely_with_historical_atr`

---

## Build order
1. Sprint A — pure config + logic, no new data deps
2. Sprint B — config + UI only, no computation
3. Sprint C — new ATR module, wire into B's contract sizing

---

## Design notes
- Buy & Hold: imported as data columns (benchmark use), excluded from eligibility/contract sizing via `is_buy_and_hold` flag
- ATR is per-strategy from trade MFE+MAE; can aggregate by symbol later
- `starting_equity` is the manual value when `solve_for_ror = False`; MC solver overrides it when `True`
- `cease_trading_threshold` triggers "stop adding new positions" in backtest when portfolio drawdown exceeds it

---

## Sprint D: Calculation Parity Test Plan
_Added: 2026-03-16. Goal: verify the refactored VBA model produces identical outputs
to the base spreadsheet, and document every known gap._

---

### D0 — Known gaps: calculations the new model does DIFFERENTLY

These are confirmed behavioural differences introduced by recent changes.
Each item needs a decision: accept / fix / document.

| # | Area | What changed | Impact | Decision needed |
|---|------|-------------|--------|----------------|
| D0-1 | **Long_Trades / Short_Trades** | BnH strategies now flow through `ProcessLSTradeData` → appear in Long_Trades & Short_Trades. Previously only non-BnH strategies were there. | **Two downstream consumers do read these sheets:** (1) `F_Summary_Tab_Setup.CalculateTradeProfitFactors` reads them to compute long/short gross profit, gross loss, and profit factor in the Summary tab. (2) `G_Create_Strategy_Tab.GetLongTradeValues/GetShortTradeValues` reads them for strategy detail charts. Before the merge, BnH strategies showed **0** for all L/S profit metrics. After the merge, BnH shows **correct** long profit factors (BnH is long-only so short metrics remain 0). This is an improvement in correctness but is a **behavioural change vs. base**. | Verify BnH profit factors in Summary after import — they should now be non-zero for BnH long trades and zero for short. Document the change. |
| D0-2 | **ATR percentile method (W_Markets)** | `atrPct` is computed as: _count(raw exit-trade ATR values ≤ atr3M) / total_. This compares a 90-day rolling average against the raw trade ATR distribution — mixing time-frames. | Percentile reads "high" whenever atr3M > median single-trade ATR. May over-state volatility regime for strategies with few large trades. | Fix: compare atr3M against rolling 90-day historical series, not raw values. |
| D0-3 | **W_Markets sector lookup uses fuzzy InStr match** | Strategy symbol is matched to BnH contract name via `InStr`. E.g. "ES" matches "ESET". | Sector and strategy-count columns may be wrong for short ticker symbols. | Fix: add word-boundary matching (`" " & contracts(i) & " "` or exact match first). |
| D0-4 | **BnH strategies in Portfolio-level P&L sheets** | BnH always flowed into `PortfolioDailyM2M` and `TotalPortfolioM2M` (unchanged). But downstream modules — Correlations, LeaveOneOut, Diversification, Monte Carlo — now process BnH without filtering. | Correlation between BnH and active strategies is market-driven, not system-driven. Including BnH distorts diversification scores and MC risk estimates. | Requires explicit BnH exclusion flags in L, S, T, K modules OR a separate analysis path. |
| D0-5 | **Backtest (N_BackTest) incomplete for BnH** | `ClosedTradePNL` is built from closed trade P&L entries only. BnH strategies that never fully close a position will have zero or partial entries. `TotalBackTest` aggregates this — BnH contribution is missing. | Portfolio backtest P&L is understated if BnH strategies are in Portfolio. | Investigate: does `ClosedTradePNL` receive BnH exits? If not, `TotalBackTest` is materially wrong. |
| D0-6 | **Tab order group "Strategies / Backtest"** | The "Backtest" sheet tab sits in the Strategies group (it was moved there during tab reorder). Backtest analysis sheets (TotalBackTest etc.) sit in the new Backtest/WhatIf group. | Slight naming confusion — "Backtest" tab = setup/config, but "Backtest/WhatIf" group = results. | Cosmetic; document the naming convention. |
| D0-7 | **Contract symbol collision in ProcessTradeData** | If two BnH strategies map to the same contract symbol via `ExtractContractName` (e.g., two different ES strategies), and both have an exit trade on the same date, the second one is silently dropped. The `If Not tradeDataDict(FileNameOnly).Exists(dateStr)` guard prevents overwrite but gives no warning. | Low likelihood in practice (rare to have two BnH strategies on the same contract with same exit date), but causes silent data loss when it occurs. | Accept for now. Add a warning log/MsgBox if collision detected. |

---

### D1 — Regression: TradeData loop merge

The two separate TradeData loops (non-BnH → ProcessLSTradeData; BnH → ProcessTradeData)
were merged into a single loop that calls both for every strategy.
These tests verify the merge produces identical outputs for pre-existing behaviour.

- [ ] **D1-1** After import: `TrueRanges` row count and column headers unchanged vs pre-merge baseline.
- [ ] **D1-2** After import: `AverageTrueRange` ATR values for each BnH contract unchanged (1M/3M/6M/12M/24M/60M/All).
- [ ] **D1-3** After import: `TradePNL` contract-level PNL values unchanged.
- [ ] **D1-4** After import: `LatestPositionData` — BnH strategies still present with correct positions (ProcessLatestPositions was already called for BnH in both loops; verify no duplicates or missing rows).
- [ ] **D1-5** After import: `Long_Trades` — non-BnH strategies have same long PNL entries as pre-merge. BnH strategies appear as new columns (expected new behaviour per D0-1).
- [ ] **D1-6** After import: `Short_Trades` — non-BnH strategies have same short PNL entries. BnH strategies appear with empty/zero short entries (BnH is long-only; verify no phantom short trades).
- [ ] **D1-7** After import: `DailyM2MEquity` — all strategies (BnH and non-BnH) have same daily equity curve as before.
- [ ] **D1-8** After import: `ClosedTradePNL` — unchanged (not touched by the TradeData loop).
- [ ] **D1-9** After import: Summary sheet `COL_PROFIT_LONG_FACTOR` / `COL_PROFIT_SHORT_FACTOR` for a BnH strategy are **non-zero** (long profit factor should reflect real trades). In the base, these were 0. After the merge they should reflect actual trade P&L. This is an intentional improvement (D0-1).
- [ ] **D1-10** Strategy detail tab for a BnH strategy: Long trade chart now shows data (previously empty). Short trade chart remains empty (BnH is long-only). Verify no errors when generating BnH detail tab.

**How to test:** Run import on a known dataset, capture the sheet values before and after the merge commit (`fc1a53f` → `4f6b72f`). Compare using a checksum or row/column counts.

---

### D2 — Regression: Portfolio sheet calculations

- [ ] **D2-1** ATR columns (1M/3M/6M/12M/24M/60M/All) in Portfolio sheet match the values in AverageTrueRange for each strategy symbol.
- [ ] **D2-2** ATR formatting: all ATR cells show `$#,##0` format (spot-check 3 rows).
- [ ] **D2-3** For a known BnH strategy: ATR 3M in Portfolio = AverageTrueRange row for that contract × contract count.
- [ ] **D2-4** For a known non-BnH strategy: ATR columns = 0 (since no BnH data for that contract). Verify this matches the base behaviour.
- [ ] **D2-5** Portfolio row count = count of Live strategies in Summary (no BnH rows added/removed incorrectly).
- [ ] **D2-6** `COL_PORT_ATR_ALL_DATA` AutoFilter covers the correct column range.

---

### D3 — W_Markets: new module validation

These are NEW calculations with no base-spreadsheet equivalent. They need
internal correctness verification rather than regression comparison.

- [ ] **D3-1 ATR percentile fix (D0-2)**: Recompute `atrPct` using rolling 90-day ATR history.
  - Build a 90-day rolling average series for each contract from `TrueRanges`.
  - `atrPct` = count(historical 90d avg <= current 90d avg) / total * 100.
  - Manual cross-check: for a contract where current ATR = 90th percentile of history, `atrPct` should return ~90.
- [ ] **D3-2 Sector match fix (D0-3)**: Test ticker "ES" does not match "ESET" or "ESU".
  - Use exact match first, then `" " + contracts(i)` prefix/suffix check.
- [ ] **D3-3 Pearson correlation smoke test**: For a contract correlated with itself, `PearsonATR(data, rows, col, col)` = 1.0.
- [ ] **D3-4 Pearson correlation direction**: For two contracts whose ATRs move together (both high in volatile markets), correlation should be positive.
- [ ] **D3-5 Volatility regime boundaries**: With `atrPct=0`, regime = "Low". With `atrPct=33`, regime = "Normal". With `atrPct=66`, regime = "Normal". With `atrPct=67`, regime = "High".
- [ ] **D3-6 Rolling 90-day window in MarketVolatility**: For the most recent date in TrueRanges, the rolling avg should equal the `atr3M` from AverageTrueRange (within rounding — both compute the same 90-day window).
- [ ] **D3-7 MarketVolatility output row count**: Matches the number of unique exit-trade dates across all BnH contracts (no duplicates, no missing dates).
- [ ] **D3-8 MarketCorrelations symmetry**: Correlation matrix is symmetric — `corr(i,j) == corr(j,i)`.

---

### D4 — Downstream impact: BnH in aggregate sheets (D0-4)

These tests quantify how much BnH inclusion changes the downstream numbers.
Until D0-4 is resolved (add BnH exclusion flags), document the magnitude of distortion.

- [ ] **D4-1 Correlations**: Run `RunCorrelationAnalysis` with and without BnH strategies in Portfolio.
  - Record: does BnH presence change any non-BnH pair's correlation by > 5 percentage points?
  - Expected finding: BnH strategies will show high correlation with long-biased active strategies.
- [ ] **D4-2 LeaveOneOut**: Run both MC and backtest modes. Note whether BnH strategies rank at the top (largest positive impact when removed = most market-exposure in portfolio).
- [ ] **D4-3 Monte Carlo**: Run portfolio MC with and without BnH in `TotalPortfolioM2M`. Record difference in median return, median max drawdown, 5th-percentile drawdown. If > 10% difference, D0-4 is critical.
- [ ] **D4-4 Diversification**: Run T_Diversificator. Verify BnH is not recommended as a "diversifying addition" when it's already in the portfolio.

---

### D5 — Backtest completeness for BnH (D0-5)

- [ ] **D5-1** Open `ClosedTradePNL` after import. For each BnH strategy in the Portfolio: is there a column present? Does it have non-zero values for dates when the BnH strategy closed a position?
- [ ] **D5-2** Run `CreateBackTestSummary()`. Compare portfolio-level cumulative P&L in `TotalBackTest` against `TotalPortfolioM2M` for the same date range. If they differ by more than the BnH contribution, there is a structural bug.
- [ ] **D5-3** If BnH exit trades DO populate ClosedTradePNL, verify the dates and amounts match TradePNL for the same contracts.
- [ ] **D5-4** If BnH exit trades do NOT populate ClosedTradePNL, document the gap clearly in a code comment and decide: add BnH to ClosedTradePNL, or exclude BnH from TotalBackTest via status filter.

---

### D6 — Tab order and navigation

- [ ] **D6-1** Run `OrderVisibleTabsBasedOnList()`. Verify tabs appear in the new group order: Settings → Folder → Strategies → Portfolio → Backtest → Markets.
- [ ] **D6-2** `GoToMarkets()` activates the Markets sheet (or shows a helpful message if not yet created).
- [ ] **D6-3** Markets sheet created by `CreateMarketsSummary()` has all three sections: Market ATR Summary, Sector Summary, and the two sub-sheets.
- [ ] **D6-4** `DeleteAllTabs()` removes Markets, MarketCorrelations, MarketVolatility without error.

---

### D7 — Calculation parity checklist (base vs new)

A side-by-side reference of which calculations are in base spreadsheet
vs the new model. Tick each as ✓ (present + matching), ✗ (missing), or ~ (differs).

| Calculation | Base spreadsheet | New model | Status |
|-------------|-----------------|-----------|--------|
| Daily M2M equity per strategy | ✓ DailyM2MEquity | ✓ unchanged | ✓ |
| Closed trade PNL per strategy | ✓ ClosedTradePNL | ✓ unchanged | ✓ |
| In-market long/short daily PNL | ✓ InMarketLong/Short | ✓ unchanged | ✓ |
| ATR per BnH contract (period avgs) | ✓ AverageTrueRange | ✓ unchanged | ✓ |
| ATR per BnH contract (raw trades) | ✓ TrueRanges | ✓ unchanged | ✓ |
| Trade PNL per BnH contract | ✓ TradePNL | ✓ unchanged | ✓ |
| Long trade PNL per strategy | ✓ Long_Trades (non-BnH only) | ~ Long_Trades now includes BnH → Summary profit factors now non-zero for BnH | D0-1 ✅ improvement |
| Short trade PNL per strategy | ✓ Short_Trades (non-BnH only) | ~ Short_Trades includes BnH (empty for long-only BnH) → short profit factors stay 0 | D0-1 ✅ no change |
| BnH long/short profit factors in Summary | ✗ zero in base | ✓ now computed from real trades | D0-1 ✅ improvement |
| Latest positions | ✓ LatestPositionData | ✓ unchanged | ✓ |
| Portfolio ATR columns | ✓ from AverageTrueRange | ✓ same lookup | ✓ |
| Strategy correlations | ✓ PortfolioDailyM2M | ~ BnH included, no filter | D0-4 |
| Portfolio MC | ✓ TotalPortfolioM2M | ~ BnH included, no filter | D0-4 |
| Backtest portfolio P&L | ✓ ClosedTradePNL | ✗ BnH missing (likely) | D0-5 |
| Market ATR percentile | ✗ not in base | ✓ new in W_Markets | Verify D3-1 |
| Market sector summary | ✗ not in base | ✓ new in W_Markets | Verify D3-2 |
| ATR correlation matrix | ✗ not in base | ✓ new in W_Markets | Verify D3-3/4/8 |
| Rolling 90-day ATR time series | ✗ not in base | ✓ new in W_Markets | Verify D3-6/7 |

---

### D8 — Build & test sequence

Run in this order to catch issues early:

1. Fresh import on known test dataset
2. D1 regression checks (TrueRanges, AverageTrueRange, Long/Short_Trades)
3. D2 Portfolio sheet checks
4. D5 Backtest completeness check (quick — just inspect ClosedTradePNL)
5. Fix D0-2 (ATR percentile), D0-3 (sector match) — these are code bugs
6. Run `CreateMarketsSummary()`, then D3 validation checks
7. D4 impact analysis (BnH in aggregates) — informs D0-4 priority
8. D6 navigation / tab order check

---
