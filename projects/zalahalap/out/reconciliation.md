# Zalahaláp – reconciliation (SPEC v2 §9)

Original: `reference/zalahalap/Zalahalap_model_10112025_v3.xlsb` (Negotiated case). New: `out/Zalahalap_Model.xlsx` (SPEC v2: PV only | PV + 2 h | PV + 4 h, unlevered, 5-year PV PPA). Written by `reconcile.py`.

Three independent calculations: the v3 shadow `data/old_engine/old_zhlp.py`, the v1-structure shadow in `reconcile.py` (quarterly, SPEC v1, with a switch per §13 fix) and the v2 shadow `shadow.py` (annual, coded from SPEC v2). The workbook is recalculated in LibreOffice.

## 1. The v3 shadow ties to the original

`old_zhlp.py` reproduces 31 of 31 targets of `data/old_engine/old_targets.csv` to 1e-6 (HTM targets from the copy with Aurora_Curve!C8 = 0, SPEC v1 §3.2).

## 2. Bridge: v3 → PV + 4 h column, Base, merchant BESS basis

Rows 1–2 are values; each further row switches one change on (cumulative) and shows its effect. Rows 2–10 run on the v1 quarterly shadow (unlevered: gearing 0 throughout), rows 11–14 on the frozen v2.0 annual shadow (`data/shadow_v20.py`), rows 15–16 on the v2.1 shadow; the hold IRR is levered from row 16 on; the last row is the workbook's PV + 4 h column (Base). The implied premium is shown before the 0 floor; the workbook floors it at 0.

| Step | SPEC | FEP hold IRR | Project IRR | RtB profit (k€) | RtB XIRR | Implied RtB premium before the 0 floor (k€/MWp) |
|---|---|---:|---:|---:|---:|---:|
| v3 original (targets; the v3 shadow ties) | v1 §3.2 | 5.50 % | 6.37 % | 2,941.4 | 224.51 % | – |
| v1 structure (quarterly), v3 rules: set 2 of the v1 workbook | v1 §11 | 5.60 % | 5.91 % | 2,921.6 | 169.51 % | -176.0 |
| + §13 fixes 2, 3, 5–8 (reimbursement timing, guarantee phantom, clipping, Aurora 4 h block, augmentation, act/365 return) | v1 §13.2–13.8 | +5.00 pp | +5.16 pp | +29.1 | -14.28 pp | +479.4 |
| + §13.14 share step (60 % at ≤ 90) | v1 §13.14 | +0.00 pp | +0.00 pp | +0.0 | +0.00 pp | +0.0 |
| + §13.16–13.17 Aurora 4 h direct and transfer prices, DC clipping 0 | v1 §13.16–17 | -2.84 pp | -2.90 pp | +0.0 | +0.00 pp | -280.4 |
| + 63 MWp, 1,265 MWh/MWp, 49.9 MVA | v2 §3.1 | +0.32 pp | +0.34 pp | +108.0 | +5.69 pp | +27.9 |
| + v2 dates: RtB 31-Dec-26, exit 31-Mar-27, NTP 30-Jun-27 | v2 §3.2 | +0.11 pp | +0.11 pp | +15.2 | -59.00 pp | +10.5 |
| + Early works four lines; E48/E49 to capex at NTP | v1 §13.15 | +0.00 pp | +0.01 pp | -13.6 | +10.06 pp | +0.2 |
| + SPV price 549 (no return), DSA 945, no PSH land step | v2 §6 | -0.14 pp | -0.28 pp | -685.7 | -49.40 pp | +0.0 |
| + Tax defaults: BESS 10 years, losses lapse NTP + 5 years | v2 §3.2 | +0.15 pp | +0.15 pp | +0.0 | +0.00 pp | +12.8 |
| + Annual timeline, year-end dates (the RtB route keeps exact dates), unlevered, v2 JV split | v2 §4, §6 | -0.18 pp | -0.20 pp | -9.5 | -6.95 pp | -18.5 |
| + BESS opex by duration: 18.70 k€/MW + 6.07 k€/MWh (Aurora; 4 h = 42.96) | fix round M2 | -1.49 pp | -1.51 pp | +0.0 | +0.00 pp | -125.2 |
| + HIPA 2 % and innovation contribution 0.3 % on revenue | v2 §3.2 | -0.49 pp | -0.50 pp | +0.0 | +0.00 pp | -39.3 |
| + PPA 5 years at 45 €/MWh on all PV output = v2.0 base (fe5a4fe) | v2 §5.2 | -0.29 pp | -0.30 pp | +0.0 | +0.00 pp | -27.4 |
| + Timeline rebase: spend to 31-Dec-26 in one column, all capex in 2027, COD 01-Mar-28 (unlevered) | v2.1 §12.1 | -0.06 pp | -0.06 pp | +0.0 | +0.00 pp | -14.0 |
| + Levered case: 60 % gearing, 5 %, 20-year annuity (hold IRR levered; base) | v2.1 §12.2 | +0.60 pp | +0.00 pp | +0.0 | +0.00 pp | +0.0 |
| **PV + 4 h, Base: the workbook (implied premium floored at 0)** | | 6.29 % | 5.91 % | 2,365.0 | 55.63 % | 0.0 |

Notes on the steps:

- Row 2 vs row 1: the v1 quarterly structure (period-end flows, all project cash in the project IRR, FEP taxed on the 12 % return, PSH land paid back); documented in SPEC v1 §13 and the v1 reconciliation.
- §13.6 (Aurora's 4 h block total instead of v3's 0.591-scaled stack) lifts the hold IRR; §13.16–17 then price the PV at Aurora's direct and transfer prices (v3: 65 % of baseload on all output) and cut it back.
- The v2 dates move the exit five months later and the RtB XIRR falls accordingly; the SPV price 549 (to FEP before the split, no return) and the DSA loan lower the RtB profit through the split.
- Annual timeline step: year-end dates on the grid, construction by calendar month share, the pre-COD cash dated (RtB route) and mapped by year, the buyer reimbursements inside the JV split, guarantee fees paid at release, deposits returned at COD on the hold route, no debt, one SPV tax.
- BESS opex step (fix round M2): v3's 30.8 k€/MW is Aurora's 2 h figure (Control!D96 note). Aurora's own table (Inputs rows 37–39, 2028 entry) gives 2 h 30.83 and 4 h 42.96 €/kW, i.e. 18.70 k€/MW + 6.07 k€/MWh; the 4 h battery now carries 42.96 instead of 30.8.
- Timeline rebase (v2.1 §12.1): spend to 31-Dec-2026 in one column at its year end, all construction capex in 2027 (was spread NTP 30-Jun-27 → COD 31-Mar-28), COD 01-Mar-2028 (10 months of production in 2028), tolling from the later of 01-Jan-28 and COD.
- Levered row: 60 % of the capex to COD drawn in 2027, IDC on mid-year draws capitalised, 20 level annuities from 31-Dec-2028, interest deductible; the hold IRR is the levered FEP IRR.
- PPA step: 45 €/MWh nominal on all PV output for 5 years (Aurora's direct price is 98–108 €/MWh nominal in 2028–29 for the 4 h block), no balancing in those years.

## 3. The v2 shadow ties to the workbook

Workbook recalculated in LibreOffice with `Dash_Case` = Base, Low, High; every Dashboard result below compared with `shadow.py` (EUR k, rates, multiples).

| Sensitivity | Master check | Results compared | Largest absolute difference | Largest relative difference |
|---|---:|---:|---:|---:|
| Base | 0 | 174 | 5.82e-11 | 7.51e-15 |
| Low | 0 | 171 | 5.09e-11 | 3.32e-13 |
| High | 0 | 174 | 5.82e-11 | 1.06e-14 |

Base results (workbook):

| Result | PV only | PV + 2h BESS | PV + 4h BESS |
|---|---:|---:|---:|
| PV generation (GWh) | 77.7 | 79.3 | 79.3 |
| Revenue | 3,644.3 | 7,282.0 | 9,729.2 |
| EBITDA | 2,094.6 | 3,978.6 | 5,712.2 |
| Revenue | 198,957.5 | 310,275.7 | 401,771.5 |
| EBITDA | 104,799.1 | 131,038.5 | 188,087.2 |
| Lifetime corporate income tax, hold route (SPV) | 5,837.1 | 5,646.0 | 8,227.7 |
| Total uses to COD (= equity, unlevered) | 36,595.0 | 47,595.0 | 58,595.0 |
| Sale at RtB: Implied premium at the buyer target IRR (floored at 0) | 0.0 | 0.0 | 0.0 |
| Sale at RtB: Buyer IRR at our price, unlevered after tax | 4.57 % | 4.27 % | 5.34 % |
| Sale at RtB: Buyer IRR at our price, unlevered before tax | 4.94 % | 4.65 % | 5.79 % |
| Sale at RtB: FEP profit, pre-tax | 2,392.6 | 2,392.6 | 2,392.6 |
| Sale at RtB: FEP profit, post-tax | 2,365.0 | 2,365.0 | 2,365.0 |
| Sale at RtB: MOIC, pre-tax | 1.320x | 1.320x | 1.320x |
| Sale at RtB: IRR, pre-tax | 56.26 % | 56.26 % | 56.26 % |
| Sale at RtB: IRR, post-tax | 55.63 % | 55.63 % | 55.63 % |
| Sale at RtB: NPV at the discount rate, post-tax (at acquisition) | 1,835.1 | 1,835.1 | 1,835.1 |
| Sale at RtB: PSH profit | 1,397.6 | 1,397.6 | 1,397.6 |
| Sale at COD: Implied COD price at the buyer target IRR (floored at 0) | 517.8 | 665.6 | 906.6 |
| Sale at COD: Enterprise value used | 32,621.7 | 41,931.9 | 57,117.9 |
| Sale at COD: Buyer IRR at our price, unlevered after tax | 8.00 % | 8.00 % | 8.00 % |
| Sale at COD: Buyer IRR at our price, unlevered before tax | 8.48 % | 8.55 % | 8.60 % |
| Sale at COD: FEP contributions to COD | 42,318.8 | 53,318.8 | 64,318.8 |
| Sale at COD: FEP proceeds at COD | 32,621.7 | 41,931.9 | 57,117.9 |
| Sale at COD: FEP profit, pre-tax | -9,697.0 | -11,386.9 | -7,200.8 |
| Sale at COD: FEP profit, post-tax | -9,697.0 | -11,386.9 | -7,200.8 |
| Sale at COD: MOIC, pre-tax | 0.774x | 0.789x | 0.889x |
| Sale at COD: IRR, pre-tax | -20.80 % | -19.73 % | -10.41 % |
| Sale at COD: IRR, post-tax | -20.80 % | -19.73 % | -10.41 % |
| Sale at COD: NPV at the discount rate, post-tax (at acquisition) | -10,672.9 | -12,701.2 | -10,096.3 |
| Sale at COD: PSH profit | -150.0 | -150.0 | -150.0 |
| Hold: FEP equity invested (incl. the 2048 repower) | 44,401.9 | 68,577.2 | 96,614.6 |
| Hold: FEP distributions | 99,497.5 | 121,740.3 | 175,881.6 |
| Hold: FEP profit, pre-tax | 60,932.8 | 58,809.1 | 87,494.8 |
| Hold: FEP profit, post-tax | 55,095.6 | 53,163.1 | 79,267.0 |
| Hold: IRR, after tax | 4.97 % | 4.64 % | 5.69 % |
| Hold: IRR, before tax | 5.34 % | 5.03 % | 6.14 % |
| Hold: NPV at the discount rate, after tax (at acquisition) | -11,991.1 | -14,019.3 | -11,414.5 |
| Hold: Memo: project IRR, unlevered after tax (no buy-out) | 5.22 % | 4.88 % | 5.91 % |
| Hold: Memo: project IRR, unlevered before tax | 5.60 % | 5.28 % | 6.37 % |
| Debt drawn in the capex year (gearing × total uses) | 21,957.0 | 28,557.0 | 35,157.0 |
| Interest during construction, capitalised | 548.9 | 713.9 | 878.9 |
| Debt at COD | 22,505.9 | 29,270.9 | 36,035.9 |
| Annual debt service (level annuity) | 1,805.9 | 2,348.8 | 2,891.6 |
| Minimum DSCR (information) | 1.0 | 1.4 | 1.6 |
| Average DSCR (information) | 1.5 | 1.5 | 1.7 |
| Lifetime corporate income tax, levered | 4,772.6 | 4,062.3 | 6,266.0 |
| Sale at RtB: Buyer IRR at our price, levered after tax (raises debt) | 4.43 % | 3.97 % | 5.66 % |
| Sale at COD: Buyer IRR at our price, levered after tax (takes debt) | 10.31 % | 12.22 % | 11.56 % |
| Sale at COD: Levered: FEP profit, post-tax (debt repaid from price) | -10,246.0 | -12,100.8 | -8,079.8 |
| Sale at COD: Levered: MOIC, pre-tax | 0.510x | 0.522x | 0.728x |
| Sale at COD: Levered: IRR, pre-tax | -43.50 % | -43.26 % | -24.28 % |
| Sale at COD: Levered: IRR, post-tax | -43.50 % | -43.26 % | -24.28 % |
| Sale at COD: Levered: PSH profit | -150.0 | -150.0 | -150.0 |
| Hold: Levered: FEP equity invested (uses less debt) | 22,444.9 | 40,020.2 | 61,457.6 |
| Hold: Levered: FEP profit, post-tax | 41,998.5 | 36,328.4 | 58,553.4 |
| Hold: Levered: IRR, after tax | 4.99 % | 4.52 % | 6.29 % |
| Hold: Levered: IRR, before tax | 5.36 % | 4.88 % | 6.74 % |
| Hold: Levered: MOIC, after tax | 2.871x | 1.908x | 1.953x |

## 4. Each §13 change inside the v3 structure (v3 shadow, cumulative; appendix)

Values in € m (profit, NPV, lifetime totals).

| Step | FEP profit – RtB | FEP IRR – RtB | FEP MOIC – RtB | FEP IRR – HTM | FEP NPV – HTM | Project IRR | Buyer IRR | Lifetime revenue | Lifetime EBITDA |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| v3 original (shadow = targets) | 2.941 | 224.51 % | 1.427x | 5.50 % | -13.009 | 6.37 % | 5.57 % | 349.819 | 177.872 |
| + Fix 3: no bank-guarantee phantom | +0.023 | +3.80 pp | +0.006x | +0.01 pp | +0.075 | +0.00 pp | +0.00 pp | +0.000 | +0.000 |
| + Fix 2: pre-RtB costs reimbursed as spent | +0.000 | -43.38 pp | +0.000x | +0.00 pp | +0.000 | +0.00 pp | +0.00 pp | +0.000 | +0.000 |
| + Fix 8: simple act/365 return, 31-Dec capitalisation | -0.002 | -0.12 pp | -0.000x | -0.00 pp | -0.002 | +0.00 pp | +0.00 pp | +0.000 | +0.000 |
| + User: share step (60 % at 90) | +0.000 | +0.00 pp | +0.000x | +0.00 pp | +0.000 | +0.00 pp | +0.00 pp | +0.000 | +0.000 |
| + User: early works, four lines (E48/E49 out of the RtB period) | -0.012 | +14.33 pp | +0.049x | +0.00 pp | +0.045 | +0.00 pp | +0.00 pp | +0.000 | +0.000 |
| + Fix 5: PV export clipped (2 %) | +0.000 | +0.00 pp | +0.000x | -0.22 pp | -1.037 | -0.23 pp | -0.22 pp | -4.798 | -4.415 |
| + Fix 6: Aurora 4 h block total for the BESS | +0.000 | +0.00 pp | +0.000x | +6.09 pp | +33.830 | +6.62 pp | +6.14 pp | +145.138 | +145.138 |
| + Fix 7: augmentation indexed | +0.000 | +0.00 pp | +0.000x | -0.28 pp | -2.294 | -0.27 pp | -0.28 pp | +0.000 | +0.000 |
