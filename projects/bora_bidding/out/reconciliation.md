# Bora bid model – reconciliation to the old engine

Model `out/Bora_Bid_Model.xlsm` (built 15-Oct-2026 08:00), recalculated in LibreOffice by `reconcile.py` on a private copy. Reconciliation set: key-assumption set "Overbuild built" (EU); CN = the same set with the input switch `In_CN_Turbine` = 1 (private copy only). Master check: EU 0.
Saved values: `reference/Bora_scenarios_meeting.xlsx` (SPEC §3), reproduced by `data/old_engine/old_rtb.py` and `old_htm.py` (≤ 1e-8). Each §8 fix is applied cumulatively in those shadows, then each user decision (SPEC §0) as its own step.
€ m to 2 dp, negatives in brackets; IRR in % (step changes in pp). Tolerance 0.1 % on values, 0.1 pp on IRR; `*` = above it (see Findings). Steps with no effect on a table are listed under it.

## 1. Headlines that tie without a bridge

| Item | BORA 1 saved | BORA 1 model | BORA 2 saved | BORA 2 model | Cells (B1, B2) |
|---|---|---:|---:|---:|---:|
| Points, EU turbine | 85.00 | 85.00 | 67.50 | 85.00 | `Dashboard!E25`, `Dashboard!F25` |
| Points, CN turbine | 70.00 | 70.00 | 52.50 | 70.00 | `Dashboard!E25`, `Dashboard!F25` |
| Installed wind MW | 151 | 151 | 140 | 151 | `Dashboard!E67`, `Dashboard!F67` |
| BESS MW | 51 | 51 | 40 | 51 | `Dashboard!E69`, `Dashboard!F69` |
| RtB price €m (gross premium; no price effect) | 15.00 | 15.00 | 15.00 | 15.00 | `Dashboard!E89`, `Dashboard!F89` |

BORA 2 now takes BORA 1's commitments (user decision, SPEC §0 item 21; step U5 below): 85 / 70 points, 151 MW wind and 51 MW BESS, as the shadow with that decision gives. The RtB price is 150 k€/MVA × 100 MVA = 15.00 with no price effect (user decision, step U1 below); saved effective prices were 3.93 / 10.08.

## 2. RtB exit bridge (set "Overbuild built")

### BORA 1

| # | Step | Price effect | RtB price | Distributable | FEP RtB profit | FEP RtB IRR |
|---|---|---:|---:|---:|---:|---:|
| 0 | Saved (meeting file; shadow reproduces it) | (11.07) | 3.93 | (3.55) | (1.22) | (3.91) % ¹ |
| 4 | PB posted award + 15 days; held to 31-Mar-2033 | +0.08 | +0.08 | +0.08 | +0.06 | +0.19 |
| 4b | PB carry in the price effect after tax (SPEC §5.6) | (0.03) | (0.03) | (0.03) | (0.02) | (0.06) |
| 5 | One FX (365) for devex (old 370) | 0.00 | 0.00 | (0.03) | (0.02) | (0.05) |
| 6 | Waterfall capped at cash; no developer floor | 0.00 | 0.00 | +3.53 | +0.27 | +0.87 |
| 7 | One participation fee per bid (old: both bids) | 0.00 | 0.00 | 0.00 | +0.01 | +0.03 |
| 8 | Guarantee fee on exact days (old: full months) | 0.00 | 0.00 | 0.00 | +0.03 | +0.09 |
| 10d | PPA for its tenor from COD (buyer margin too) | +0.25 | +0.25 | 0.00 | +0.25 | +0.81 |
| 11 | Quarterly timeline: price effect, flows at quarter ends | (0.29) | (0.29) | 0.00 | (0.29) | (1.32) |
| U1 | User decision: RtB price 150 k€/MVA fixed, no price effect | +11.05 | +11.05 | +7.49 | +7.47 | +25.52 |
|  | **Shadow, all fixes and decisions** | 0.00 | 15.00 | 7.49 | 6.55 | 22.15 % |
|  | Residual: workbook − shadow | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
|  | **Workbook, set "Overbuild built"** | 0.00 | 15.00 | 7.49 | 6.55 | 22.15 % |

Steps with no effect here: 1, 2, 3, 9; U2–U4 do not reach the RtB route (the price is fixed).
Developer at the exit (DSA + 30 %): saved 0.14, shadow 3.45, workbook 3.45.
¹ Excel's XIRR failed in the meeting file (2.98e-09); −3.91 % is the true root (LibreOffice, shadow). The model shows `n/a` plus a check when XIRR fails (fix 6).

### BORA 2

| # | Step | Price effect | RtB price | Distributable | FEP RtB profit | FEP RtB IRR |
|---|---|---:|---:|---:|---:|---:|
| 0 | Saved (meeting file; shadow reproduces it) | (4.92) | 10.08 | 2.84 | 3.27 | 10.15 % |
| 4 | PB posted award + 15 days; held to 31-Mar-2033 | +0.10 | +0.10 | +0.10 | +0.07 | +0.22 |
| 4b | PB carry in the price effect after tax (SPEC §5.6) | (0.03) | (0.03) | (0.03) | (0.02) | (0.07) |
| 5 | One FX (365) for devex (old 370) | 0.00 | 0.00 | (0.03) | (0.02) | (0.07) |
| 7 | One participation fee per bid (old: both bids) | 0.00 | 0.00 | 0.00 | +0.01 | +0.03 |
| 8 | Guarantee fee on exact days (old: full months) | 0.00 | 0.00 | 0.00 | +0.02 | +0.10 |
| 10d | PPA for its tenor from COD (buyer margin too) | +0.12 | +0.12 | +0.12 | +0.08 | +0.25 |
| 11 | Quarterly timeline: price effect, flows at quarter ends | (0.11) | (0.11) | (0.11) | (0.08) | +1.13 |
| U1 | User decision: RtB price 150 k€/MVA fixed, no price effect | +4.84 | +4.84 | +4.84 | +3.39 | +11.16 |
| U5 | User decision: BORA 2 takes BORA 1's commitments (151 MW wind, 51 MW BESS, community 3) | 0.00 | 0.00 | (0.23) | (0.18) | (0.76) |
|  | **Shadow, all fixes and decisions** | 0.00 | 15.00 | 7.49 | 6.55 | 22.15 % |
|  | Residual: workbook − shadow | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
|  | **Workbook, set "Overbuild built"** | 0.00 | 15.00 | 7.49 | 6.55 | 22.15 % |

Steps with no effect here: 1, 2, 3, 6, 9; U2–U4 do not reach the RtB route (the price is fixed).
Developer at the exit (DSA + 30 %): saved 1.97, shadow 3.45, workbook 3.45.

Cells (B1 / B2): RtB price `Dashboard!E89` / `Dashboard!F89`, distributable `Dashboard!E96` / `Dashboard!F96`, profit `Dashboard!E26` / `Dashboard!F26`, IRR `Dashboard!E27` / `Dashboard!F27`, developer `Dashboard!E101` / `Dashboard!F101`.

## 3. Build, sale at COD and hold to maturity (set "Overbuild built"; CN = same set, `In_CN_Turbine` = 1)

### BORA 1

| # | Step | All-in cost EU | All-in cost CN | COD IRR EU | COD IRR CN | HTM IRR EU | HTM IRR CN |
|---|---|---:|---:|---:|---:|---:|---:|
| 0 | Saved (meeting file; shadow reproduces it) | 273.86 | 196.32 | 31.67 % | 42.05 % | 10.89 % | 14.36 % |
| 4 | PB posted award + 15 days; held to 31-Mar-2033 | 0.00 | 0.00 | 0.00 | 0.00 | (0.03) | (0.04) |
| 5 | One FX (365) for devex (old 370) | +0.03 | +0.03 | (0.03) | (0.04) | 0.00 | 0.00 |
| 7 | One participation fee per bid (old: both bids) | (0.01) | (0.01) | +0.01 | +0.02 | 0.00 | 0.00 |
| 8 | Guarantee fee on exact days (old: full months) | (0.03) | (0.03) | +0.05 | +0.07 | 0.00 | 0.00 |
| 9 | Claims per class on actual days; portfolio floor | 0.00 | 0.00 | (0.02) | (0.02) | 0.00 | 0.00 |
| 10a | HTM stand-alone: developer buy-out, no deemed exit | 0.00 | 0.00 | 0.00 | 0.00 | (0.05) | (0.08) |
| 10b | No DSRA, no cash sweep | 0.00 | 0.00 | 0.00 | 0.00 | +0.48 | +0.99 |
| 10c | No loss cap, no decommissioning | 0.00 | 0.00 | 0.00 | 0.00 | +0.01 | +0.01 |
| 10d | PPA for its tenor from COD (buyer margin too) | 0.00 | 0.00 | 0.00 | 0.00 | (0.06) | (0.12) |
| U1 | User decision: RtB price 150 k€/MVA fixed, no price effect | 0.00 | 0.00 | 0.00 | 0.00 | (0.23) | (0.38) |
| U2 | User decision: EU turbine 600 k€/MW (was 950) | (58.15) | 0.00 | +21.97 | 0.00 | +3.13 | 0.00 |
| U3 | User decision: PPA 90 €/MWh for the whole life (was 91.6 for 10 y, then merchant) | 0.00 | 0.00 | 0.00 | 0.00 | (2.50) | (2.31) |
| U4 | User decision: BESS capex 100 k€/MWh = 400 k€/MW at 4 h (was 200 k€/MW) | +10.69 | +10.61 | (4.04) | (4.26) | (0.70) | (0.74) |
|  | **Shadow, all fixes and decisions** | 226.39 | 206.92 | 49.62 % | 37.82 % | 10.95 % | 11.70 % |
| 11c | All-in cost now includes the pre-COD land lease (Calc commitment overview) | +0.93 | +0.93 | 0.00 | 0.00 | 0.00 | 0.00 |
| 11a | New structure: amounts (quarterly debt, IDC, fees, days) | (0.91)* | (0.70)* | +0.17* | +0.14* | +0.06 | +0.04 |
| 11b | New structure: XIRR on quarter ends (old: annual / mid-year) | 0.00 | 0.00 | +1.57* | +1.71* | (0.06) | (0.07) |
|  | **Workbook, set "Overbuild built"** | 226.41 | 207.16 | 51.35 % | 39.67 % | 10.95 % | 11.66 % |

Steps with no effect here: 1, 2, 3, 4b, 6.
At-risk claim to COD (fix 9): portfolio floor 3.12, per-line floor 3.25 (+0.13 €m claim, +0.02 pp COD IRR EU): the model uses the portfolio floor (SPEC §5.4).

### BORA 2

| # | Step | All-in cost EU | All-in cost CN | COD IRR EU | COD IRR CN | HTM IRR EU | HTM IRR CN |
|---|---|---:|---:|---:|---:|---:|---:|
| 0 | Saved (meeting file; shadow reproduces it) | 255.49 | 183.50 | 31.90 % | 41.93 % | 10.57 % | 13.81 % |
| 4 | PB posted award + 15 days; held to 31-Mar-2033 | 0.00 | 0.00 | 0.00 | 0.00 | (0.03) | (0.04) |
| 5 | One FX (365) for devex (old 370) | +0.03 | +0.03 | (0.03) | (0.04) | 0.00 | 0.00 |
| 7 | One participation fee per bid (old: both bids) | (0.01) | (0.01) | +0.02 | +0.02 | 0.00 | 0.00 |
| 8 | Guarantee fee on exact days (old: full months) | (0.03) | (0.03) | +0.05 | +0.07 | 0.00 | 0.00 |
| 9 | Claims per class on actual days; portfolio floor | 0.00 | 0.00 | (0.02) | (0.02) | 0.00 | 0.00 |
| 10a | HTM stand-alone: developer buy-out, no deemed exit | 0.00 | 0.00 | 0.00 | 0.00 | +0.21 | +0.33 |
| 10b | No DSRA, no cash sweep | 0.00 | 0.00 | 0.00 | 0.00 | +0.48 | +0.96 |
| 10c | No loss cap, no decommissioning | 0.00 | 0.00 | 0.00 | 0.00 | +0.01 | +0.01 |
| 10d | PPA for its tenor from COD (buyer margin too) | 0.00 | 0.00 | 0.00 | 0.00 | (0.07) | (0.12) |
| U1 | User decision: RtB price 150 k€/MVA fixed, no price effect | 0.00 | 0.00 | 0.00 | 0.00 | (0.10) | (0.17) |
| U2 | User decision: EU turbine 600 k€/MW (was 950) | (53.91) | 0.00 | +21.25 | 0.00 | +3.11 | 0.00 |
| U3 | User decision: PPA 90 €/MWh for the whole life (was 91.6 for 10 y, then merchant) | 0.00 | 0.00 | 0.00 | 0.00 | (2.45) | (2.28) |
| U4 | User decision: BESS capex 100 k€/MWh = 400 k€/MW at 4 h (was 200 k€/MW) | +8.38 | +8.32 | (3.30) | (3.49) | (0.59) | (0.62) |
| U5 | User decision: BORA 2 takes BORA 1's commitments (151 MW wind, 51 MW BESS, community 3) | +16.43 | +15.11 | (0.25) | (0.65) | (0.20) | (0.19) |
|  | **Shadow, all fixes and decisions** | 226.39 | 206.92 | 49.62 % | 37.82 % | 10.95 % | 11.70 % |
| 11c | All-in cost now includes the pre-COD land lease (Calc commitment overview) | +0.93 | +0.93 | 0.00 | 0.00 | 0.00 | 0.00 |
| 11a | New structure: amounts (quarterly debt, IDC, fees, days) | (0.91)* | (0.70)* | +0.17* | +0.14* | +0.06 | +0.04 |
| 11b | New structure: XIRR on quarter ends (old: annual / mid-year) | 0.00 | 0.00 | +1.57* | +1.71* | (0.06) | (0.07) |
|  | **Workbook, set "Overbuild built"** | 226.41 | 207.16 | 51.35 % | 39.67 % | 10.95 % | 11.66 % |

Steps with no effect here: 1, 2, 3, 4b, 6.
At-risk claim to COD (fix 9): portfolio floor 3.12, per-line floor 3.25 (+0.13 €m claim, +0.02 pp COD IRR EU): the model uses the portfolio floor (SPEC §5.4).

Cells (B1 / B2): all-in cost `Dashboard!E108` / `Dashboard!F108`, COD IRR `Dashboard!E28` / `Dashboard!F28`, HTM IRR `Dashboard!E30` / `Dashboard!F30`.

New-structure detail, EU set "Overbuild built", € m (shadow with all fixes and decisions vs workbook; lifetime sums):

| Line | B1 shadow | B1 model | B1 Δ | B2 Δ | Driver |
|---|---|---:|---:|---:|---|
| IDC capitalised | 6.72 | 5.64 | (1.08) | (1.08) | interest on the quarter-opening balance; draws at quarter end (old: annual, half-year draws) |
| Commitment fee | 1.11 | 1.27 | +0.16 | +0.16 | on undrawn at the quarter start (old: net of capitalised IDC) |
| Upfront fee | 2.20 | 2.20 | 0.00 | 0.00 | same basis |
| Other cost-to-COD lines | 216.37 | 217.30 | +0.93 | +0.93 | pre-COD land lease (11c) in the model; other cost lines tie |
| COD price effect | (3.57) | (3.67) | (0.10) | (0.10) | community Δ per quarter, discounted from quarter ends |
| Debt repaid at COD | 132.47 | 131.40 | (1.08) | (1.08) | lower IDC |
| FEP contributions to COD | (94.81) | (95.02) | (0.21) | (0.21) | commitment fee, pre-COD lease |
| FEP proceeds at COD | 188.39 | 189.13 | +0.74 | +0.74 | lower debt, more distributable (70 %) |
| Generation, GWh | 10,289.7 | 10,289.8 | +0.06 | +0.06 | actual days (leap years), degradation by operating year |
| Revenue | 1,079.42 | 1,079.40 | (0.02) | (0.02) | same, PPA by quarter from COD |
| EBITDA | 760.47 | 760.33 | (0.14) | (0.14) | same |
| Interest paid | (40.74) | (36.38) | +4.36 | +4.36 | quarterly annuity from COD (old: annual from 2032) |
| Corporate income tax | (44.62) | (45.08) | (0.46) | (0.46) | lower interest, quarterly loss use |
| FEP HTM cash flow | 444.64 | 449.32 | +4.68 | +4.68 | sum of the above |

## 4. Set "Overbuild built" → Base (overbuild not built), for the IC

Base = the Base set (EU) and Base with `In_CN_Turbine` = 1 (CN). The middle column removes the build only (`In_Tolerance_Cap` = 100 % suppresses the sanction, private copy); the next adds the pro-rata PB sanction 6 months after COD.

### BORA 1

| Item | Set "Overbuild built" | Not built | PB sanction | Base | Δ total | Cell |
|---|---|---:|---:|---:|---:|---:|
| Points EU | 85.00 | 0.00 | 0.00 | 85.00 | 0.00 | `Dashboard!E25` |
| Points CN | 70.00 | 0.00 | 0.00 | 70.00 | 0.00 | `Dashboard!E25` |
| Distributable | 7.49 | 0.00 | 0.00 | 7.49 | 0.00 | `Dashboard!E96` |
| FEP RtB profit | 6.55 | 0.00 | 0.00 | 6.55 | 0.00 | `Dashboard!E26` |
| FEP RtB IRR | 22.15 % | 0.00 | 0.00 | 22.15 % | 0.00 | `Dashboard!E27` |
| Total commitment EU | 233.41 | (53.68) | 0.00 | 179.73 | (53.68) | `Dashboard!E32` |
| All-in cost EU | 226.41 | (53.68) | 0.00 | 172.74 | (53.68) | `Dashboard!E108` |
| All-in cost CN | 207.16 | (47.76) | 0.00 | 159.40 | (47.76) | `Dashboard!E108` |
| COD IRR EU | 51.35 % | (16.90) | (0.53) | 33.92 % | (17.43) | `Dashboard!E28` |
| COD IRR CN | 39.67 % | (15.40) | (0.58) | 23.68 % | (15.99) | `Dashboard!E28` |
| HTM IRR EU | 10.95 % | (1.96) | (0.11) | 8.88 % | (2.07) | `Dashboard!E30` |
| HTM IRR CN | 11.66 % | (2.04) | (0.11) | 9.51 % | (2.15) | `Dashboard!E30` |

### BORA 2

| Item | Set "Overbuild built" | Not built | PB sanction | Base | Δ total | Cell |
|---|---|---:|---:|---:|---:|---:|
| Points EU | 85.00 | 0.00 | 0.00 | 85.00 | 0.00 | `Dashboard!F25` |
| Points CN | 70.00 | 0.00 | 0.00 | 70.00 | 0.00 | `Dashboard!F25` |
| Distributable | 7.49 | 0.00 | 0.00 | 7.49 | 0.00 | `Dashboard!F96` |
| FEP RtB profit | 6.55 | 0.00 | 0.00 | 6.55 | 0.00 | `Dashboard!F26` |
| FEP RtB IRR | 22.15 % | 0.00 | 0.00 | 22.15 % | 0.00 | `Dashboard!F27` |
| Total commitment EU | 233.41 | (53.68) | 0.00 | 179.73 | (53.68) | `Dashboard!F32` |
| All-in cost EU | 226.41 | (53.68) | 0.00 | 172.74 | (53.68) | `Dashboard!F108` |
| All-in cost CN | 207.16 | (47.76) | 0.00 | 159.40 | (47.76) | `Dashboard!F108` |
| COD IRR EU | 51.35 % | (16.90) | (0.53) | 33.92 % | (17.43) | `Dashboard!F28` |
| COD IRR CN | 39.67 % | (15.40) | (0.58) | 23.68 % | (15.99) | `Dashboard!F28` |
| HTM IRR EU | 10.95 % | (1.96) | (0.11) | 8.88 % | (2.07) | `Dashboard!F30` |
| HTM IRR CN | 11.66 % | (2.04) | (0.11) | 9.51 % | (2.15) | `Dashboard!F30` |

## 5. Findings

- No unexplained residual on the RtB headlines, points, wind/BESS MW or RtB price: the shadow with every fix and decision ties to the workbook, both bids.
- All-in cost to COD after the lease (11c), new structure (11a): EU BORA 1 (0.91); BORA 2 (0.91); CN BORA 1 (0.70); BORA 2 (0.70) €m. IDC accrues on the quarter-opening balance with draws at quarter end (old: annual, half-year draws), partly offset by the commitment fee on the start-of-quarter undrawn amount. Convention, not a bug.
- COD IRR, date basis (11b): EU BORA 1 +1.57 pp; BORA 2 +1.57 pp; CN BORA 1 +1.71 pp; BORA 2 +1.71 pp. The old engine dated each year's contributions mid-year; the model dates them at quarter ends. Amounts (11a): EU BORA 1 +0.17 pp; BORA 2 +0.17 pp; CN BORA 1 +0.14 pp; BORA 2 +0.14 pp (lower IDC, less debt repaid at COD). Attributed.
- HTM IRR, date basis (11b): EU within tolerance; CN within tolerance; quarterly XIRR vs the old annual IRR from 2026. The workbook flows re-run on the old basis tie to the shadow (11a: EU within tolerance; CN within tolerance). Attributed.
- The CN runs and the not-built runs overwrite Live cells in the private copy, so only the stored-vs-live checks fail there (master check ≠ 0); every economic check is 0.
