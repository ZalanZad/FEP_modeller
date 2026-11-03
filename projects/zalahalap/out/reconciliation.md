# Zalahaláp – reconciliation (SPEC v2 §9)

Original: `reference/zalahalap/Zalahalap_model_10112025_v3.xlsb` (Negotiated case). New: `out/Zalahalap_Model.xlsx` (SPEC v2: PV only | PV + 2 h | PV + 4 h, unlevered, 5-year PV PPA). Written by `reconcile.py`.

Three independent calculations: the v3 shadow `data/old_engine/old_zhlp.py`, the v1-structure shadow in `reconcile.py` (quarterly, SPEC v1, with a switch per §13 fix) and the v2 shadow `shadow.py` (annual, coded from SPEC v2). The workbook is recalculated in LibreOffice.

## 1. The v3 shadow ties to the original

`old_zhlp.py` reproduces 31 of 31 targets of `data/old_engine/old_targets.csv` to 1e-6 (HTM targets from the copy with Aurora_Curve!C8 = 0, SPEC v1 §3.2).

## 2. Bridge: v3 → PV + 4 h column, Base, merchant BESS basis

Rows 1–2 are values; each further row switches one change on (cumulative) and shows its effect. Rows 2–10 run on the v1 quarterly shadow (unlevered: gearing 0 throughout), rows 11–13 on the v2 annual shadow; the last row is the workbook's PV + 4 h column (Base). The implied premium is shown before the 0 floor; the workbook floors it at 0.

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
| + HIPA 2 % and innovation contribution 0.3 % on revenue | v2 §3.2 | -0.45 pp | -0.46 pp | +0.0 | +0.00 pp | -39.3 |
| + PPA 5 years at 45 €/MWh on all PV output = base | v2 §5.2 | -0.30 pp | -0.31 pp | +0.0 | +0.00 pp | -27.4 |
| **PV + 4 h, Base: the workbook (implied premium floored at 0)** | | 7.27 % | 7.52 % | 2,365.0 | 55.63 % | 0.0 |

Notes on the steps:

- Row 2 vs row 1: the v1 quarterly structure (period-end flows, all project cash in the project IRR, FEP taxed on the 12 % return, PSH land paid back); documented in SPEC v1 §13 and the v1 reconciliation.
- §13.6 (Aurora's 4 h block total instead of v3's 0.591-scaled stack) lifts the hold IRR; §13.16–17 then price the PV at Aurora's direct and transfer prices (v3: 65 % of baseload on all output) and cut it back.
- The v2 dates move the exit five months later and the RtB XIRR falls accordingly; the SPV price 549 (to FEP before the split, no return) and the DSA loan lower the RtB profit through the split.
- Annual timeline step: year-end dates on the grid, construction by calendar month share, the pre-COD cash dated (RtB route) and mapped by year, the buyer reimbursements inside the JV split, guarantee fees paid at release, deposits returned at COD on the hold route, no debt, one SPV tax.
- PPA step: 45 €/MWh nominal on all PV output for 5 years (Aurora's direct price is 98–108 €/MWh nominal in 2028–29 for the 4 h block), no balancing in those years.

## 3. The v2 shadow ties to the workbook

Workbook recalculated in LibreOffice with `Dash_Case` = Base, Low, High; every Dashboard result below compared with `shadow.py` (EUR k, rates, multiples).

| Sensitivity | Master check | Results compared | Largest absolute difference | Largest relative difference |
|---|---:|---:|---:|---:|
| Base | 0 | 117 | 5.82e-11 | 9.93e-15 |
| Low | 0 | 117 | 5.82e-11 | 3.42e-14 |
| High | 0 | 117 | 5.82e-11 | 4.55e-15 |

Base results (workbook):

| Result | PV only | PV + 2h BESS | PV + 4h BESS |
|---|---:|---:|---:|
| PV generation (GWh) | 79.3 | 79.3 | 79.3 |
| Revenue | 3,718.6 | 7,282.0 | 9,729.2 |
| EBITDA | 2,167.3 | 3,980.8 | 6,371.7 |
| Revenue | 203,149.5 | 310,420.9 | 401,994.6 |
| EBITDA | 108,393.3 | 131,005.5 | 220,473.1 |
| Lifetime corporate income tax, hold route (SPV) | 6,160.6 | 5,643.0 | 11,142.4 |
| Total uses to COD (= equity, unlevered) | 36,595.0 | 47,595.0 | 58,595.0 |
| Sale at RtB: Implied premium at the buyer target IRR (floored at 0) | 0.0 | 0.0 | 0.0 |
| Sale at RtB: Buyer IRR at our price, unlevered after tax | 4.85 % | 4.32 % | 6.90 % |
| Sale at RtB: Buyer IRR at our price, unlevered before tax | 5.23 % | 4.71 % | 7.45 % |
| Sale at RtB: FEP profit, pre-tax | 2,392.6 | 2,392.6 | 2,392.6 |
| Sale at RtB: FEP profit, post-tax | 2,365.0 | 2,365.0 | 2,365.0 |
| Sale at RtB: MOIC, pre-tax | 1.320x | 1.320x | 1.320x |
| Sale at RtB: IRR, pre-tax | 56.26 % | 56.26 % | 56.26 % |
| Sale at RtB: IRR, post-tax | 55.63 % | 55.63 % | 55.63 % |
| Sale at RtB: NPV at the discount rate, post-tax (to the acquisition date) | 1,835.1 | 1,835.1 | 1,835.1 |
| Sale at RtB: PSH profit | 1,397.6 | 1,397.6 | 1,397.6 |
| Sale at COD: Implied COD price at the buyer target IRR (floored at 0) | 530.8 | 659.7 | 1,041.5 |
| Sale at COD: Enterprise value used | 33,439.4 | 41,562.5 | 65,614.1 |
| Sale at COD: Buyer IRR at our price, unlevered after tax | 8.00 % | 8.00 % | 8.00 % |
| Sale at COD: Buyer IRR at our price, unlevered before tax | 8.49 % | 8.54 % | 8.64 % |
| Sale at COD: FEP contributions to COD | 42,321.2 | 53,321.2 | 64,321.2 |
| Sale at COD: FEP proceeds at COD | 33,439.4 | 41,562.5 | 65,344.8 |
| Sale at COD: FEP profit, pre-tax | -8,881.8 | -11,758.7 | 1,023.7 |
| Sale at COD: FEP profit, post-tax | -8,881.8 | -11,758.7 | 967.9 |
| Sale at COD: MOIC, pre-tax | 0.713x | 0.692x | 1.022x |
| Sale at COD: IRR, pre-tax | -24.66 % | -27.26 % | 1.92 % |
| Sale at COD: IRR, post-tax | -24.66 % | -27.26 % | 1.82 % |
| Sale at COD: NPV at the discount rate, post-tax (to the acquisition date) | -9,475.9 | -12,208.2 | -2,637.3 |
| Sale at COD: PSH profit | -150.0 | -150.0 | 119.2 |
| Hold: FEP equity invested (incl. the 2048 repower) | 42,281.8 | 64,518.8 | 89,752.4 |
| Hold: FEP distributions | 100,645.7 | 117,649.5 | 198,488.1 |
| Hold: FEP profit, pre-tax | 64,524.5 | 58,773.7 | 119,878.2 |
| Hold: FEP profit, post-tax | 58,363.9 | 53,130.7 | 108,735.7 |
| Hold: IRR, after tax | 5.24 % | 4.68 % | 7.27 % |
| Hold: IRR, before tax | 5.63 % | 5.08 % | 7.83 % |
| Hold: NPV at the discount rate, after tax (to the acquisition date) | -10,794.1 | -13,526.4 | -3,699.2 |
| Hold: Memo: project IRR, unlevered after tax (FEP's before the buy-out) | 5.50 % | 4.93 % | 7.52 % |
| Hold: Memo: project IRR, unlevered before tax | 5.90 % | 5.33 % | 8.09 % |

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
