# Portfolio Tracker v2 — Task Board

_Last updated: 2026-03-15_

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
