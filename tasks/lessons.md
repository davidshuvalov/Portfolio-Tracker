# Lessons Learned

_Updated: 2026-03-16_

---

## L1 — Never assume a sheet has no downstream consumers without a full codebase search

**Context**: Sprint D test plan initially classified `Long_Trades` and `Short_Trades`
as having "no downstream readers" based on a partial audit, rating D0-1 as zero-impact.

**What actually happened**: A second audit found that two modules **do** read those sheets:
- `F_Summary_Tab_Setup.CalculateTradeProfitFactors` — reads Long_Trades/Short_Trades to compute
  long gross profit, short gross profit, and profit factors for every strategy in the Summary tab.
- `G_Create_Strategy_Tab.GetLongTradeValues / GetShortTradeValues` — reads them to populate
  trade charts in individual strategy detail tabs.

The first auditor searched for `Long_Trades` references but may have missed indirect references
(e.g., sheet names stored as string literals inside helper functions in separate modules).

**Rule**: Before classifying a sheet change as "no downstream impact", **grep the entire codebase**
for the sheet name as a string literal, including inside helper functions and utility modules —
not just in the main orchestration module.

```
grep -r "Long_Trades\|Short_Trades" /path/to/codebase --include="*.bas"
```

---

## L2 — BnH strategies showing zero profit factors in Summary was a silent bug

**Context**: Before the TradeData loop merge, BnH strategies never went through
`ProcessLSTradeData`, so they had no entries in Long_Trades/Short_Trades. This caused
`CalculateTradeProfitFactors` to find no column for them and return 0 for all L/S metrics.

The old behaviour (zeros) looked like valid data but was actually incomplete. The merge
**fixed** this silently — BnH strategies now show real long profit factors.

**Rule**: When merging code paths, check whether the old path had any silent gaps that
the new path implicitly fills. Distinguish "this is a regression" from "this is a fix".

---

## L3 — Silent data-drop in ProcessTradeData for same-contract BnH collision

**Context**: `ProcessTradeData` uses `If Not tradeDataDict(contract).Exists(dateStr) Then`
to prevent overwriting. If two BnH strategies share the same contract symbol AND exit on
the same date, the second strategy's ATR/PNL for that date is silently dropped.

**No error, no warning, no log entry.**

**Rule**: Whenever a `Dict.Exists()` guard silently skips data, add at minimum a `Debug.Print`
or accumulate a warning message. Silent data loss in financial calculations is unacceptable.
Treat first-one-wins collisions as bugs unless explicitly documented as intended behaviour.

---
