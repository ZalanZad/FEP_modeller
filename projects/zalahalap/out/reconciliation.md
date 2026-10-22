# Zalahaláp – reconciliation of Scenario A to the original model

Original: `reference/zalahalap/Zalahalap_model_10112025_v3.xlsb` (Negotiated case), targets in `data/old_engine/old_targets.csv` ([R] = from the copy with Aurora_Curve!C8 = 0, SPEC §3.2). New: `out/Zalahalap_Model.xlsm`, Calc_1 = case A. Written by `reconcile.py`.

Three independent calculations: the v3 shadow `data/old_engine/old_zhlp.py` (ties to every target), the new-structure shadow in `reconcile.py` (quarterly, coded from the SPEC; stored in `data/recon_targets.csv` and tested on the Checks sheet), and the workbook recalculated in LibreOffice with the solvers re-run.

Model checks: set 2 = 0, base = 0. Tolerance 0.1 % on values, 0.1 pp on rates.

## 1. The original targets

| Target | Unit | v3 cell | v3 target | v3 shadow | Shadow ties | New model, set 2 | Ties | Why it differs |
|---|---|---|---:|---:|---|---:|---|---|
| rtb_xirr_incl | % | Summary!D6 | 224.51 % | 224.51 % | yes | 170.61 % | no | quarterly grid (§13.12): the acquisition flows move from 30-Nov-25 to 31-Dec-25 and the sale from 31-Oct-26 to 31-Dec-26, the pre-RtB costs to quarter ends; the profit ties |
| rtb_xirr_excl | % | Calculation!F144 | 181.83 % | 181.83 % | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| rtb_moic_incl_dep | x | Summary!D7 | 1.427x | 1.427x | yes | 1.426x | yes |  |
| rtb_moic_excl_dep | x | Summary!D25 | 1.576x | 1.576x | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| calc_moic_incl_dep | x | Calculation!G179 | 1.795x | 1.795x | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| calc_moic_excl_dep | x | Calculation!G180 | 2.529x | 2.529x | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| rtb_net_profit | EUR m | Summary!D8 | 2.941 | 2.941 | yes | 2.941 | yes |  |
| rtb_capital | EUR m | Summary!D9 | 6.882 | 6.882 | yes | 6.901 | no | guarantee fees run to the end of the exit quarter (31-Dec-26; v3: the sale month) on quarter-end balances; deposits equal (1,776) |
| rtb_proceeds | EUR m | Summary!D23 | 9.824 | 9.824 | yes | 9.842 | no | mirror of capital and profit |
| developer_profit | EUR k | Calculation!F226 | 2,429.6 | 2,429.6 | yes | 2,561.1 | no | v3 deducts the PSH land apport (150) from the price but never credits it to PSH (Calculation F225); the new model pays it back to PSH |
| rtb_breakeven | EUR k/MW | RTB_Model!D78 | 9.6 | 9.6 | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| htm_xirr | % | Summary!H6 [R] | 5.50 % | 5.50 % | yes | 5.60 % | yes |  |
| htm_moic | x | Summary!H7 [R] | 1.959x | 1.959x | yes | 1.977x | no | as the HTM IRR |
| htm_npv | EUR m | Summary!H8 [R] | -13.009 | -13.009 | yes | -12.415 | no | as the HTM IRR; discounted to 31-Dec-25 in both |
| htm_capital | EUR m | Summary!H9 [R] | 85.419 | 85.419 | yes | 84.846 | no | as the HTM IRR |
| htm_payback | years | Summary!H10 [R] | 16.8 | 16.8 | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| project_irr | % | Summary!D31 [R] | 6.37 % | 6.37 % | yes | 5.91 % | no | §13.4: includes all project cash (deposits, guarantees, pre-RtB costs, early works when paid); no working capital; tax base incl. development costs |
| buyer_irr | % | Summary!D33 [R] | 5.57 % | 5.57 % | yes | 5.22 % | no | buyer pays the price, early works, cash deposits and pre-RtB costs at the exit (v3: price and early works only), bears guarantee fees to COD and gets the deposits back at COD |
| buyer_moic | x | Summary!D34 [R] | 1.986x | 1.986x | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| buyer_npv | EUR m | Summary!D36 [R] | -12.524 | -12.524 | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| ffy_gwh | GWh | Summary!H21 | 75.6 | 75.6 | yes | 75.6 | yes |  |
| ffy_price | EUR/MWh | Summary!H22 | 72.6 | 72.6 | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| ffy_revenue | EUR m | Summary!H24 | 8.849 | 8.849 | yes | 8.849 | yes |  |
| life_pv_revenue | EUR m | Summary!H25 [R] | 233.190 | 233.190 | yes | 232.619 | no | operations end 31-Mar-2063 = COD + 35 years (v3: 30-Apr-2063, one month more) |
| life_goo | EUR m | Summary!H26 | 6.708 | 6.708 | yes | 6.692 | no | as PV revenue |
| life_revenue | EUR m | Summary!H27 [R] | 349.819 | 349.819 | yes | 348.985 | no | as PV revenue |
| life_ebitda | EUR m | Summary!H30 [R] | 177.872 | 177.872 | yes | 177.480 | no | as PV revenue (opex rules identical) |
| life_margin | % | Summary!H31 [R] | 50.85 % | 50.85 % | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| life_capex | EUR m | Summary!H33 | 57.170 | 57.170 | yes | 57.170 | yes |  |
| life_opex | EUR m | Summary!H53 [R] | 171.947 | 171.947 | yes | – | – | no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |
| life_bess_net | EUR m | Summary!H59 | 109.921 | 109.921 | yes | 109.674 | no | same v3 stack (fix 6 off); one month shorter life |

## 2. Each §13 change inside the v3 structure (v3 shadow, cumulative)

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
| **v3 structure, all switchable fixes** | 2.951 | 199.13 % | 1.482x | 11.10 % | 17.609 | 12.48 % | 11.21 % | 490.159 | 318.595 |

Values in € m (profit, NPV, lifetime totals). At the price of 90 the share step changes nothing (both rules give 60 %, user 22-Oct-2026).

## 3. New structure: reconciliation set → base set (workbook, cumulative)

Values in € k (profit, NPV, lifetime totals). Each step removes the set-2 override of the inputs named (the set then inherits the base value) and re-runs Solver.bas and GoalSeek.bas.

| Step | SPEC | FEP profit – RtB | FEP IRR – RtB | FEP MOIC – RtB | FEP IRR – HTM | FEP NPV – HTM | Project IRR | Buyer IRR at 90 | Lifetime revenue | Lifetime EBITDA |
|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| Workbook, set 2 (Reconciliation, v3 rules) | §10 | 2,941.4 | 170.61 % | 1.426x | 5.60 % | -12,415.3 | 5.91 % | 5.22 % | 348,984.6 | 177,480.5 |
| + Fix 2: pre-RtB costs reimbursed as spent | §13.2 | +0.0 | -17.83 pp | +0.000x | +0.00 pp | +0.0 | +0.00 pp | +0.00 pp | +0.0 | +0.0 |
| + Fix 3: no bank-guarantee phantom | §13.3 | +31.6 | +3.63 pp | +0.008x | +0.01 pp | +71.2 | +0.02 pp | +0.01 pp | +0.0 | +0.0 |
| + Fix 5: PV export clipped (2 %) | §13.5 | +0.0 | +0.00 pp | +0.000x | -0.22 pp | -1,044.7 | -0.22 pp | -0.21 pp | -4,786.2 | -4,403.8 |
| + Fix 6: Aurora 4 h block total for the BESS | §13.6 | +0.0 | +0.00 pp | +0.000x | +6.15 pp | +34,066.7 | +6.34 pp | +5.96 pp | +144,764.1 | +144,764.1 |
| + Fix 7: augmentation indexed, Aurora repower year | §13.7 | +0.0 | +0.00 pp | +0.000x | -0.28 pp | -2,233.3 | -0.27 pp | -0.28 pp | +0.0 | +0.0 |
| + Fix 8: dated act/365 return, 31-Dec capitalisation | §13.8 | -3.9 | -0.21 pp | -0.001x | -0.00 pp | -3.6 | +0.00 pp | +0.00 pp | +0.0 | +0.0 |
| + User: share step (60 % at 90) | §13.14, §0.4 | +0.0 | +0.00 pp | +0.000x | +0.00 pp | +0.0 | +0.00 pp | +0.00 pp | +0.0 | +0.0 |
| + User: early works, four lines; E48/E49 at NTP | §13.15, §0.3, §0.5 | -8.8 | +10.24 pp | +0.050x | +0.01 pp | +17.6 | +0.01 pp | +0.00 pp | +0.0 | +0.0 |
| + Tax: HIPA and innovation contribution | §13.10 | +0.0 | +0.00 pp | +0.000x | -0.46 pp | -2,686.1 | -0.47 pp | -0.44 pp | +0.0 | -11,246.1 |
| + Tax defaults: BESS 10 years, losses lapse after 5 years | §15.9 | +0.0 | +0.00 pp | +0.000x | +0.14 pp | +730.8 | +0.15 pp | +0.14 pp | +0.0 | +0.0 |
| + JV sheet 221025: SPV 549, DSA loan, no PSH land step | §7, §15.10 | -667.2 | -80.91 pp | -0.175x | -0.16 pp | -651.0 | -0.32 pp | +0.01 pp | +0.0 | +0.0 |
| + Debt: 55 % gearing (Bana terms) | §13.9, §15.1 | +0.0 | +0.00 pp | +0.000x | +2.43 pp | +3,832.4 | +0.16 pp | +0.15 pp | +0.0 | +0.0 |
| **Workbook, base set** | | 2,293.2 | 85.53 % | 1.308x | 13.24 % | 19,684.6 | 11.30 % | 10.54 % | 488,962.5 | 306,594.6 |

Last step vs the base set: largest difference 0.00e+00 (the steps restore every set-2 override).

## 4. Independent shadow vs the workbook (set 2, Calc_1)

| Headline | Shadow | Workbook | Difference |
|---|---:|---:|---:|
| FEP profit – RtB | 2,941.4 | 2,941.4 | -2.27e-12 |
| FEP IRR – RtB | 170.61 % | 170.61 % | -1.33e-15 |
| FEP MOIC – RtB | 1.426x | 1.426x | -2.22e-16 |
| FEP IRR – HTM | 5.60 % | 5.60 % | -7.63e-17 |
| FEP NPV – HTM | -12,415.3 | -12,415.3 | 5.46e-12 |
| Project IRR | 5.91 % | 5.91 % | -4.86e-17 |
| Buyer IRR at 90 | 5.22 % | 5.22 % | -9.02e-17 |
| Lifetime revenue | 348,984.6 | 348,984.6 | 0.00e+00 |
| Lifetime EBITDA | 177,480.5 | 177,480.5 | 0.00e+00 |

## 5. Structural changes (§13 items without a switch)

- §13.1: Aurora_Curve!C8 = 'no ' breaks the v3 HTM chain; the [R] targets come from the patched copy.
- §13.4: the project IRR includes all project cash (v3 HTM_Model!D79 left out deposits, guarantees and pre-RtB costs).
- §13.9: v3 has no debt; set 2 runs at 0 % gearing, the base at 55 % (Bana terms, SPEC §15.1).
- §13.10: HIPA and the innovation contribution are off in set 2; tax on the sale gain per deal form (share deal, exempt). The tax base also carries IDC, pre-RtB costs, the DSA and the SPV price (v3: construction capex only).
- §13.11: one FX (400) for the HUF deposits and land payments.
- §13.12: one quarterly timeline from 01-Oct-2025 (v3: monthly RtB route, quarterly HTM); flows at period ends; the exit quarter belongs to the seller; the 12 % return on a payment starts in the next quarter (v3: in the month paid).
- §13.13: no typed values on calc sheets, no data tables, one break-even (the implied price at the buyer target IRR), `n/a` plus a check instead of IFERROR(…,0).
- Not in §13, by design of the new structure: no working capital (v3: 30 days receivable and payable), construction capex spread linearly over NTP → COD (v3: buyer-model monthly curve), case A COD = 31-Mar-2028 (v3 01-Apr-2028 = operations from the same day), degradation on a 365-day year.

## 6. v3 bugs not reproduced (SPEC §13)

Summary!D28 → RTB_Model!D75 instead of D78; Sensitivities!E39:E42 use D73/D74; Calculation!G225 omits G221; Dyn_Inp!G24/H26:Y26 row offsets; Control!D26 overrides the construction-start default; stale notes. Also found: the developer's proceeds (Calculation F225) omit the 150 land apport deducted at F155.
