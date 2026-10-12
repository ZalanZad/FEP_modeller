# Bora bid model – reconciliation to the old engine

Model `out/Bora_Bid_Model.xlsm` (built 12-Oct-2026 20:33), recalculated in LibreOffice by `reconcile.py` on a private copy. Reconciliation set: key-assumption set "Overbuild built" (EU); CN = the same set with the input switch `In_CN_Turbine` = 1 (private copy only). Master check: EU 0.
Saved values: `reference/Bora_scenarios_meeting.xlsx` (SPEC §3), reproduced by `data/old_engine/old_rtb.py` and `old_htm.py` (≤ 1e-8). Each §8 fix is applied cumulatively in those shadows, then each user decision (SPEC §0) as its own step.
€ m to 2 dp, negatives in brackets; IRR in % (step changes in pp). Tolerance 0.1 % on values, 0.1 pp on IRR; `*` = above it (see Findings). Steps with no effect on a table are listed under it.

## 1. Headlines that tie without a bridge

| Item | BORA 1 saved | BORA 1 model | BORA 2 saved | BORA 2 model | Cells (B1, B2) |
|---|---|---:|---:|---:|---:|
| Points, EU turbine | 85.00 | 85.00 | 67.50 | 68.25 | `Dashboard!E25`, `Dashboard!F25` |
| Points, CN turbine | 70.00 | 70.00 | 52.50 | 53.25 | `Dashboard!E25`, `Dashboard!F25` |
| Installed wind MW | 151 | 151 | 140 | 140 | `Dashboard!E43`, `Dashboard!F43` |
| BESS MW | 51 | 51 | 40 | 40 | `Dashboard!E45`, `Dashboard!F45` |
| RtB price €m (gross premium; no price effect) | 15.00 | 15.00 | 15.00 | 15.00 | `Dashboard!E61`, `Dashboard!F61` |

BORA 2 points +0.75 (EU and CN): user decision, community 2.01 €/MWh scores ÖSZ-2 (the shadow's tender tables give 68.25 / 53.25). The RtB price is 150 k€/MVA × 100 MVA = 15.00 with no price effect (user decision, step U1 below); saved effective prices were 3.93 / 10.08.

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

Steps with no effect here: 1, 2, 3, 9; U2–U5 do not reach the RtB route (the price is fixed).
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
|  | **Shadow, all fixes and decisions** | 0.00 | 15.00 | 7.73 | 6.73 | 22.91 % |
|  | Residual: workbook − shadow | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
|  | **Workbook, set "Overbuild built"** | 0.00 | 15.00 | 7.73 | 6.73 | 22.91 % |

Steps with no effect here: 1, 2, 3, 6, 9; U2–U5 do not reach the RtB route (the price is fixed).
Developer at the exit (DSA + 30 %): saved 1.97, shadow 3.43, workbook 3.43.

Cells (B1 / B2): RtB price `Dashboard!E61` / `Dashboard!F61`, distributable `Dashboard!E71` / `Dashboard!F71`, profit `Dashboard!E26` / `Dashboard!F26`, IRR `Dashboard!E27` / `Dashboard!F27`, developer `Dashboard!E74` / `Dashboard!F74`.

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
| U4 | User decision: BESS capex 240 k€/MW (was 200) | +2.14 | +2.12 | (0.81) | (0.86) | (0.14) | (0.15) |
|  | **Shadow, all fixes and decisions** | 217.84 | 198.44 | 52.85 % | 41.23 % | 11.51 % | 12.28 % |
| 11c | All-in cost now includes the pre-COD land lease (Calc commitment overview) | +0.93 | +0.93 | 0.00 | 0.00 | 0.00 | 0.00 |
| 11a | New structure: amounts (quarterly debt, IDC, fees, days) | (0.88)* | (0.67)* | +0.16* | +0.13* | +0.05 | +0.03 |
| 11b | New structure: XIRR on quarter ends (old: annual / mid-year) | 0.00 | 0.00 | +1.44* | +1.65* | (0.07) | (0.08) |
|  | **Workbook, set "Overbuild built"** | 217.90 | 198.70 | 54.45 % | 43.00 % | 11.49 % | 12.24 % |

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
| U4 | User decision: BESS capex 240 k€/MW (was 200) | +1.68 | +1.66 | (0.66) | (0.70) | (0.12) | (0.13) |
| U5 | User decision: BORA 2 community 2.01 €/MWh (was 2.0) | 0.00 | 0.00 | (0.01) | (0.01) | 0.00 | 0.00 |
|  | **Shadow, all fixes and decisions** | 203.26 | 185.16 | 52.50 % | 41.25 % | 11.61 % | 12.38 % |
| 11c | All-in cost now includes the pre-COD land lease (Calc commitment overview) | +0.87 | +0.87 | 0.00 | 0.00 | 0.00 | 0.00 |
| 11a | New structure: amounts (quarterly debt, IDC, fees, days) | (0.82)* | (0.62)* | +0.18* | +0.15* | +0.05 | +0.03 |
| 11b | New structure: XIRR on quarter ends (old: annual / mid-year) | 0.00 | 0.00 | +1.16* | +1.42* | (0.07) | (0.08) |
|  | **Workbook, set "Overbuild built"** | 203.31 | 185.41 | 53.84 % | 42.82 % | 11.59 % | 12.33 % |

Steps with no effect here: 1, 2, 3, 4b, 6.
At-risk claim to COD (fix 9): portfolio floor 3.12, per-line floor 3.25 (+0.13 €m claim, +0.02 pp COD IRR EU): the model uses the portfolio floor (SPEC §5.4).

Cells (B1 / B2): all-in cost `Dashboard!E89` / `Dashboard!F89`, COD IRR `Dashboard!E28` / `Dashboard!F28`, HTM IRR `Dashboard!E30` / `Dashboard!F30`.

New-structure detail, EU set "Overbuild built", € m (shadow with all fixes and decisions vs workbook; lifetime sums):

| Line | B1 shadow | B1 model | B1 Δ | B2 Δ | Driver |
|---|---|---:|---:|---:|---|
| IDC capitalised | 6.45 | 5.42 | (1.03) | (0.96) | interest on the quarter-opening balance; draws at quarter end (old: annual, half-year draws) |
| Commitment fee | 1.07 | 1.22 | +0.16 | +0.15 | on undrawn at the quarter start (old: net of capitalised IDC) |
| Upfront fee | 2.12 | 2.12 | 0.00 | 0.00 | same basis |
| Other cost-to-COD lines | 208.21 | 209.14 | +0.93 | +0.87 | pre-COD land lease (11c) in the model; other cost lines tie |
| COD price effect | (3.57) | (3.67) | (0.10) | 0.00 | community Δ per quarter, discounted from quarter ends |
| Debt repaid at COD | 127.32 | 126.28 | (1.03) | (0.96) | lower IDC |
| FEP contributions to COD | (91.42) | (91.62) | (0.20) | (0.19) | commitment fee, pre-COD lease |
| FEP proceeds at COD | 190.98 | 191.69 | +0.71 | +0.73 | lower debt, more distributable (70 %) |
| Generation, GWh | 10,289.7 | 10,289.8 | +0.06 | +0.05 | actual days (leap years), degradation by operating year |
| Revenue | 1,079.42 | 1,079.40 | (0.02) | (0.01) | same, PPA by quarter from COD |
| EBITDA | 760.47 | 760.33 | (0.14) | (0.13) | same |
| Interest paid | (39.16) | (34.96) | +4.19 | +3.91 | quarterly annuity from COD (old: annual from 2032) |
| Corporate income tax | (45.53) | (45.97) | (0.44) | (0.41) | lower interest, quarterly loss use |
| FEP HTM cash flow | 453.86 | 458.35 | +4.49 | +4.19 | sum of the above |

## 4. Set "Overbuild built" → Base (overbuild not built), for the IC

Base = the Base set (EU) and Base with `In_CN_Turbine` = 1 (CN). The middle column removes the build only (`In_Tolerance_Cap` = 100 % suppresses the sanction, private copy); the next adds the pro-rata PB sanction 6 months after COD.

### BORA 1

| Item | Set "Overbuild built" | Not built | PB sanction | Base | Δ total | Cell |
|---|---|---:|---:|---:|---:|---:|
| Points EU | 85.00 | 0.00 | 0.00 | 85.00 | 0.00 | `Dashboard!E25` |
| Points CN | 70.00 | 0.00 | 0.00 | 70.00 | 0.00 | `Dashboard!E25` |
| Distributable | 7.49 | 0.00 | 0.00 | 7.49 | 0.00 | `Dashboard!E71` |
| FEP RtB profit | 6.55 | 0.00 | 0.00 | 6.55 | 0.00 | `Dashboard!E26` |
| FEP RtB IRR | 22.15 % | 0.00 | 0.00 | 22.15 % | 0.00 | `Dashboard!E27` |
| Total commitment EU | 224.89 | (53.68) | 0.00 | 171.21 | (53.68) | `Dashboard!E32` |
| All-in cost EU | 217.90 | (53.68) | 0.00 | 164.22 | (53.68) | `Dashboard!E89` |
| All-in cost CN | 198.70 | (47.76) | 0.00 | 150.94 | (47.76) | `Dashboard!E89` |
| COD IRR EU | 54.45 % | (16.08) | (0.52) | 37.85 % | (16.60) | `Dashboard!E28` |
| COD IRR CN | 43.00 % | (14.59) | (0.56) | 27.85 % | (15.15) | `Dashboard!E28` |
| HTM IRR EU | 11.49 % | (1.88) | (0.12) | 9.49 % | (2.00) | `Dashboard!E30` |
| HTM IRR CN | 12.24 % | (1.96) | (0.12) | 10.16 % | (2.08) | `Dashboard!E30` |

### BORA 2

| Item | Set "Overbuild built" | Not built | PB sanction | Base | Δ total | Cell |
|---|---|---:|---:|---:|---:|---:|
| Points EU | 68.25 | 0.00 | 0.00 | 68.25 | 0.00 | `Dashboard!F25` |
| Points CN | 53.25 | 0.00 | 0.00 | 53.25 | 0.00 | `Dashboard!F25` |
| Distributable | 7.73 | 0.00 | 0.00 | 7.73 | 0.00 | `Dashboard!F71` |
| FEP RtB profit | 6.73 | 0.00 | 0.00 | 6.73 | 0.00 | `Dashboard!F26` |
| FEP RtB IRR | 22.91 % | 0.00 | 0.00 | 22.91 % | 0.00 | `Dashboard!F27` |
| Total commitment EU | 209.39 | (42.10) | 0.00 | 167.29 | (42.10) | `Dashboard!F32` |
| All-in cost EU | 203.31 | (42.10) | 0.00 | 161.20 | (42.10) | `Dashboard!F89` |
| All-in cost CN | 185.41 | (37.46) | 0.00 | 147.95 | (37.46) | `Dashboard!F89` |
| COD IRR EU | 53.84 % | (12.99) | (0.40) | 40.46 % | (13.38) | `Dashboard!F28` |
| COD IRR CN | 42.82 % | (11.78) | (0.43) | 30.60 % | (12.22) | `Dashboard!F28` |
| HTM IRR EU | 11.59 % | (1.54) | (0.10) | 9.95 % | (1.64) | `Dashboard!F30` |
| HTM IRR CN | 12.33 % | (1.61) | (0.10) | 10.62 % | (1.71) | `Dashboard!F30` |

## 5. Findings

- No unexplained residual on the RtB headlines, points, wind/BESS MW or RtB price: the shadow with every fix and decision ties to the workbook, both bids.
- All-in cost to COD after the lease (11c), new structure (11a): EU BORA 1 (0.88); BORA 2 (0.82); CN BORA 1 (0.67); BORA 2 (0.62) €m. IDC accrues on the quarter-opening balance with draws at quarter end (old: annual, half-year draws), partly offset by the commitment fee on the start-of-quarter undrawn amount. Convention, not a bug.
- COD IRR, date basis (11b): EU BORA 1 +1.44 pp; BORA 2 +1.16 pp; CN BORA 1 +1.65 pp; BORA 2 +1.42 pp. The old engine dated each year's contributions mid-year; the model dates them at quarter ends. Amounts (11a): EU BORA 1 +0.16 pp; BORA 2 +0.18 pp; CN BORA 1 +0.13 pp; BORA 2 +0.15 pp (lower IDC, less debt repaid at COD). Attributed.
- HTM IRR, date basis (11b): EU within tolerance; CN within tolerance; quarterly XIRR vs the old annual IRR from 2026. The workbook flows re-run on the old basis tie to the shadow (11a: EU within tolerance; CN within tolerance). Attributed.
- The CN runs and the not-built runs overwrite Live cells in the private copy, so only the stored-vs-live checks fail there (master check ≠ 0); every economic check is 0.
