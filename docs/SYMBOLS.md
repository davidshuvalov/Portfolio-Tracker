# Futures Symbol Reference

Covers every symbol used in Portfolio Tracker / MultiCharts / TradeStation.
Symbols are shown in TradeStation (TS) format (prefix `@`) and canonical CME/exchange format.

---

## TradeStation ↔ Standard Symbol Map

| TS Symbol | CME/Std Symbol | Name | Exchange | Sector |
|-----------|---------------|------|----------|--------|
| @AD | 6A | Australian Dollar (100K AUD) | CME | Currencies |
| @BP | 6B | British Pound (62,500 GBP) | CME | Currencies |
| @CD | 6C | Canadian Dollar (100K CAD) | CME | Currencies |
| @EC | 6E | Euro FX (125,000 EUR) | CME | Currencies |
| @JY | 6J | Japanese Yen (12.5M JPY) | CME | Currencies |
| @NE1 | 6N | New Zealand Dollar (100K NZD) | CME | Currencies |
| @SF | 6S | Swiss Franc (125,000 CHF) | CME | Currencies |
| @MP1 | 6M | Mexican Peso (500K MXN) | CME | Currencies |
| @DX | DX | U.S. Dollar Index | ICE | Currencies |
| @ES | ES | E-mini S&P 500 | CME | Index |
| @NQ | NQ | E-mini Nasdaq-100 | CME | Index |
| @RTY | RTY | E-mini Russell 2000 | CME | Index |
| @YM | YM | E-mini Dow Jones | CBOT | Index |
| @EMD | EMD | E-mini S&P MidCap 400 | CME | Index |
| @NK | NK | Nikkei 225 (JPY-denominated) | OSE | Index |
| @FDXM | FDXM | Mini-DAX (€5 × index, 1/5 FDAX) | EUREX | Index |
| @FESX | FESX | Euro Stoxx 50 | EUREX | Index |
| @FGBL | FGBL | Euro Bund 10Y | EUREX | Rates |
| @FGBM | FGBM | Euro Bobl 5Y | EUREX | Rates |
| @FGBS | FGBS | Euro Schatz 2Y | EUREX | Rates |
| @FGBX | FGBX | Euro Buxl 30Y | EUREX | Rates |
| @TU | ZT | 2-Year U.S. T-Note | CBOT | Rates |
| @FV | ZF | 5-Year U.S. T-Note | CBOT | Rates |
| @TY | ZN | 10-Year U.S. T-Note | CBOT | Rates |
| @US | ZB | 30-Year U.S. T-Bond | CBOT | Rates |
| @CL | CL | Crude Oil WTI | NYMEX | Energy |
| @NG | NG | Natural Gas | NYMEX | Energy |
| @HO | HO | NY Harbor ULSD (Heating Oil) | NYMEX | Energy |
| @RB | RB | RBOB Gasoline | NYMEX | Energy |
| @BRN | BRN | Brent Crude Oil | ICE | Energy |
| @GC | GC | Gold (100 troy oz) | COMEX | Metals |
| @SI | SI | Silver (5,000 troy oz) | COMEX | Metals |
| @HG | HG | Copper (25,000 lbs) | COMEX | Metals |
| @PL | PL | Platinum (50 troy oz) | NYMEX | Metals |
| @PA | PA | Palladium (100 troy oz) | NYMEX | Metals |
| @C | ZC | Corn (5,000 bu) | CBOT | Agriculture |
| @S | ZS | Soybeans (5,000 bu) | CBOT | Agriculture |
| @W | ZW | Chicago SRW Wheat (5,000 bu) | CBOT | Agriculture |
| @KW | KE | Hard Red Winter Wheat (5,000 bu) | CBOT | Agriculture |
| @O | ZO | Oats (5,000 bu) | CBOT | Agriculture |
| @BO | ZL | Soybean Oil (60,000 lbs) | CBOT | Agriculture |
| @SM | ZM | Soybean Meal (100 short tons) | CBOT | Agriculture |
| @RR | ZR | Rough Rice (2,000 cwt) | CBOT | Agriculture |
| @DA | DC | Class III Milk (200K lbs) | CME | Agriculture |
| @LC | LE | Live Cattle (40,000 lbs) | CME | Agriculture |
| @LH | HE | Lean Hogs (40,000 lbs) | CME | Agriculture |
| @FC | GF | Feeder Cattle (50,000 lbs) | CME | Agriculture |
| @LB | LB | Random Length Lumber (110K bd ft) | CME | Agriculture |
| @KC | KC | Coffee C (37,500 lbs) | ICE | Softs |
| @CC | CC | Cocoa (10 metric tons) | ICE | Softs |
| @CT | CT | Cotton #2 (50,000 lbs) | ICE | Softs |
| @SB | SB | Sugar #11 (112,000 lbs) | ICE | Softs |
| @OJ | OJ | Orange Juice (15,000 lbs) | ICE | Softs |
| @BTC | BTC | Bitcoin (5 BTC) | CME | Crypto |
| @MBT | MBT | Micro Bitcoin (0.1 BTC) | CME | Crypto |
| @VX | VX | CBOE VIX Futures | CBOE | Volatility |

---

## Mini / Micro Contracts

### Equity Indices

| Standard (TS) | Standard | Mini | Mini Ratio | Micro | Micro Ratio |
|---------------|----------|------|-----------|-------|------------|
| @ES | ES | — | — | MES | 0.1× |
| @NQ | NQ | — | — | MNQ | 0.1× |
| @RTY | RTY | — | — | M2K | 0.1× |
| @YM | YM | — | — | MYM | 0.1× |
| @FDXM | FDXM (Mini-DAX) | — | — | FDXS | 0.2× FDXM / 0.04× FDAX |

> The ES, NQ, RTY, and YM are themselves the "Mini" versions of older floor contracts.
> The "Micro" series (MES, MNQ, etc.) are 1/10th of these.

### Currencies

| Standard (TS) | Standard | Mini | Mini Ratio | Micro | Micro Ratio |
|---------------|----------|------|-----------|-------|------------|
| @EC | 6E | E7 | 0.5× | M6E | 0.1× |
| @JY | 6J | J7 | 0.5× | M6J | 0.1× |
| @BP | 6B | — | — | M6B | 0.1× |
| @AD | 6A | — | — | M6A | 0.1× |
| @CD | 6C | — | — | MCD | 0.1× |
| @SF | 6S | — | — | MSF | 0.1× |
| @NE1 | 6N | — | — | M6N | 0.1× |

> **MCD** (CAD/USD direction, 10,000 CAD) is the correct micro for @CD/6C.
> **MSF** (CHF/USD direction, 12,500 CHF) is the correct micro for @SF/6S.
> M6C and M6S exist on CME but are **inverse-quoted** (USD/CAD, USD/CHF) — different products.
> **MJY** also exists (10,000 USD, USD/JPY inverse) — low volume, avoid.

### Energy

| Standard (TS) | Standard | Mini | Mini Ratio | Micro | Micro Ratio |
|---------------|----------|------|-----------|-------|------------|
| @CL | CL | QM | 0.5× | MCL | 0.1× |
| @NG | NG | QG | 0.5× | — | — |

### Metals

| Standard (TS) | Standard | Mini | Mini Ratio | Micro | Micro Ratio |
|---------------|----------|------|-----------|-------|------------|
| @GC | GC | — | — | MGC | 0.1× |
| @SI | SI | — | — | SIL | **0.2×** (1,000 oz — CME calls it "Micro" but it is 1/5, not 1/10) |
| @HG | HG | — | — | MHG | 0.1× |

### Agriculture

| Standard (TS) | Standard | Mini | Mini Ratio | Micro | Micro Ratio |
|---------------|----------|------|-----------|-------|------------|
| @C | ZC | XC | 0.2× | MZC | 0.1× (launched Feb 2025) |
| @S | ZS | XK | 0.2× | MZS | 0.1× (launched Feb 2025) |
| @W | ZW | XW | 0.2× | MZW | 0.1× (launched Feb 2025) |
| @KW | KE | MKC | 0.2× | — | — (no micro; low liquidity mini) |

### Interest Rates

| Standard (TS) | Standard | Micro | Micro Ratio | Notes |
|---------------|----------|-------|------------|-------|
| @US | ZB | — | — | No micro |
| — | UB | MWN | 0.1× | Ultra T-Bond micro; cash-settled |
| @TY | ZN | — | — | No micro |
| — | TN | MTN | 0.1× | Ultra 10-Year micro; cash-settled |
| @FV | ZF | — | — | No micro |
| @TU | ZT | — | — | No micro |

> The CME Micro Treasury **Yield** contracts (2YY, 5YY, 10Y, 30Y) trade in **yield**, not price —
> they are not true scaled-down versions of the bond futures and behave differently.

### Crypto

| Standard | Micro | Micro Ratio | Notes |
|----------|-------|------------|-------|
| BTC (5 BTC) | MBT | **0.02×** (0.1 BTC ÷ 5 BTC) | |
| ETH (50 ETH) | MET | **0.002×** (0.1 ETH ÷ 50 ETH) | Fee structure makes ETH very hard to trade profitably |

---

## What to Trade

### Trade freely

| Symbol | Notes |
|--------|-------|
| MES, MNQ, M2K, MYM | Most liquid micro equity index futures |
| MCL | Liquid micro crude |
| MGC, MHG, SIL | Liquid micro metals |
| M6E, M6A, M6B, M6J, MCD, MSF, M6N | Liquid micro FX |
| MBT | Liquid micro Bitcoin |
| E7, J7 | Mini EUR and JPY — liquid |
| XC, XK, XW | Mini ag (corn, soybean, wheat) — liquid |
| MZC, MZS, MZW | New micro ag (Feb 2025) — monitor liquidity |

### Use with caution

| Symbol | Reason |
|--------|--------|
| QM | E-mini crude — large tick size relative to profit target; MCL preferred |
| QG | E-mini natural gas — same tick size problem; limited benefit over standard NG |
| MJY | Micro JPY (inverse quote) — low volume; use M6J instead |
| FDXS | Micro-DAX on Eurex — newer, check liquidity before trading |
| MWN, MTN | Micro Ultra T-Bond / 10-Year — check liquidity; cash-settled |
| BTC / ETH | Large per-point value; use MBT; avoid MET |

### Avoid (liquidity / volume issues)

| Symbol | Reason |
|--------|--------|
| MKC | Mini HRW Wheat — not liquid |
| MMC | Micro MidCap — not liquid |
| MIR, MNH | Micro FX (Indian Rupee, offshore CNH) — minimal volume |
| MHO, MRB | Micro heating oil / gasoline — not liquid |
| MSC | Micro S&P MidCap — not liquid |
| PAM, PLM | Micro palladium / platinum — not liquid |
| 5YY, 30Y | Micro Treasury Yield — barely trades; note: yield-based, not price-based |
| 2YY, 10Y | Micro Treasury Yield — some volume but not price-based contracts |

---

## Continuous Contracts: Negative Price Warning

Some back-adjusted continuous contracts can produce **negative OHLC prices** in historical data.
Strategies must handle this explicitly (e.g., percent-based stops, log-price calculations, or
positive-only data checks).

**Known negative-price symbols:**
`@CL, @CT, @HO, @OJ, @RB, @S, @SM, @QM`

**Known clean (positive-only) continuous symbols:**
`@AD, @BO, @BP, @BTC, @C, @CC, @CD, @DX, @EC, @EMD, @ES, @FC, @FV, @GC, @HG, @JY, @KC, @KW, @LC, @LH, @MBT, @MP1, @NE1, @NG, @NK, @NQ, @O, @PL, @RR, @SB, @SF, @SI, @TU, @TY, @US, @VX, @W, @YM, @MES, @E7, @J7, @QN`

---

## Sector Groupings

### Indexes
`ES, NQ, RTY, YM, EMD` — micros: `MES, MNQ, M2K, MYM`

### European Indexes (Eurex)
`FESX, FDAX` — mini: `FDXM` — micro: `FDXS`

### Currencies
`6E, 6J, 6B, 6A, 6C, 6S, 6N, DX, 6M` — minis: `E7, J7` — micros: `M6E, M6J, M6B, M6A, MCD, MSF, M6N`

### Rates — CBOT Treasuries
`ZT, ZF, ZN, ZB, UB, TN` — micros: `MWN (UB), MTN (TN)`

### Rates — Eurex
`FGBS, FGBM, FGBL, FGBX`

### Energy
`CL, NG, HO, RB, BRN` — mini: `QM, QG` — micro: `MCL`

### Metals
`GC, SI, HG, PL, PA` — micros: `MGC, SIL (0.2×), MHG`

### Agriculture — Grains
`ZC, ZS, ZW, KE, ZO, ZR` — minis: `XC, XK, XW, MKC` — micros: `MZC, MZS, MZW`

### Agriculture — Oilseeds
`ZL, ZM`

### Agriculture — Livestock / Dairy
`LE, HE, GF, DC`

### Agriculture — Softs
`KC, CC, CT, SB, OJ, LB`

### Crypto
`BTC, ETH` — micros: `MBT (0.02×), MET (0.002×)`

### Volatility
`VX`

---

## TS Symbol Sets (as configured)

```
Full contracts (all exchanges incl. ICE):
@AD @BO @BP @C @CC @CD @CL @CT @DX @EC @EMD @ES @FC @FV @GC @HG @HO
@JY @KC @KW @LC @LH @MBT @MP1 @NG @NK @NQ @O @OJ @PL @RB @RR @RTY
@S @SB @SF @SI @SM @TU @TY @US @W

Mini/micro set:
@MES @MNQ @M2K @MYM @E7 @J7 @QM @QN @MGC @SIL @MHG @YW @YK

Kevin's list (adds VX, NE1, BTC):
@AD @BO @BP @BTC @C @CC @CD @CL @CT @DX @EC @EMD @ES @FC @FV @GC @HG
@HO @JY @KC @KW @LC @LH @MBT @MP1 @NE1 @NG @NK @NQ @O @OJ @PL @RB
@RR @S @SB @SF @SI @SM @TU @TY @US @VX @W @YM

Extended (Eurex + ICE):
@AD @BO @BP @BRN @BTC @C @CC @CD @CL @CT @DX @EC @EMD @ES @FC @FDXM
@FESX @FGBL @FGBM @FGBS @FGBX @FV @GC @HG @HO @JY @KC @KW @LC @LH
@MBT @MP1 @NE1 @NG @NK @NQ @O @OJ @PL @RB @RR @RTY @S @SB @SF @SI
@SM @TU @TY @US @VX @W @YM
```

> Note: `@BRN=105XC`, `@DA=108XC`, `@LB=108XC`, `@FGBX=108XC`, `@PA=105NC` etc. are TS exchange
> qualifiers for ICE/non-CME data feeds. Strip the `=XXX` suffix to get the base symbol.
