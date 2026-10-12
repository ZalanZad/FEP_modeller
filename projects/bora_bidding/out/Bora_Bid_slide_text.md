# Project Bora – IC deck text (6 slides)

Production notes (not deck text): every figure is read from `out/Bora_Bid_Model_recalc.xlsx` (Dashboard, base case; master check 0). Each line ends with its cells (the `Out_` name where one exists) and display format; the full list is in `Bora_Bid_slide_text_figures.csv`. Formats: € m 1 dp with negatives in brackets; signed Δ € m `+1.4` / `(0.5)`; % 1 dp with a true minus; pp 1 dp, signed; points as integers, or 2 dp where fractional; k€/MW 0 dp with a thousands separator; MOIC 2 dp. "B1 / B2" = BORA 1 / BORA 2. [key row] = bold on `E6F4EC`.  
  <sub>Cells: `Master_Check` (0 dp)</sub>

## Slide 1 – Cover

- Title: Project Bora – BORA 1 and BORA 2 capacity bids
- Subtitle: Approval of the tender commitments, 2 × 100 MVA, Szombathely OVIT
- Date: 12-Oct-2026  
  <sub>Cells: `Dashboard!E6` (date)</sub>

## Slide 2 – Decision and timetable

- Section: Executive summary
- Action title: Submit both bids at 85 / 68.25 points: FEP earns €6.5m / €6.7m at the RtB exit, a 22.2% / 22.9% IRR  
  <sub>Cells: `Out_Points_1`, `Out_Points_2` (points (2 dp if fractional)); `Out_RtB_Profit_1`, `Out_RtB_Profit_2` (€ m 1 dp); `Out_RtB_IRR_1`, `Out_RtB_IRR_2` (% 1 dp)</sub>
- Subtitle: Against the zero bid, the commitments cost €0.8m / €0.7m of RtB profit and 2.4 / 1.9 pp of HTM IRR, and add €23.4m / €19.5m to the all-in commitment.  
  <sub>Cells: `Dashboard!I26`, `Dashboard!J26`, `Dashboard!I32`, `Dashboard!J32` (€ m 1 dp (magnitude)); `Dashboard!I30`, `Dashboard!J30` (pp 1 dp (magnitude))</sub>

Table (returns vs the zero bid):

| Item | Unit | Zero bid | BORA 1 | BORA 2 | Cells (format) |
|---|---|---:|---:|---:|---|
| Points (EU turbine) | points | 0 | 85 | 68.25 | `Dashboard!H25`, `Out_Points_1`, `Out_Points_2` (points (2 dp if fractional)) |
| FEP profit – RtB route | € m | 7.4 | 6.5 | 6.7 | `Dashboard!H26`, `Out_RtB_Profit_1`, `Out_RtB_Profit_2` (€ m 1 dp) |
| FEP IRR – RtB route | % | 25.8% | 22.2% | 22.9% | `Dashboard!H27`, `Out_RtB_IRR_1`, `Out_RtB_IRR_2` (% 1 dp) |
| FEP IRR – sale at COD | % | 49.6% | 37.8% | 40.5% | `Dashboard!H28`, `Out_COD_IRR_1`, `Out_COD_IRR_2` (% 1 dp) |
| FEP IRR – hold to maturity | % | 11.9% | 9.5% | 9.9% | `Dashboard!H30`, `Out_HTM_IRR_1`, `Out_HTM_IRR_2` (% 1 dp) |
| All-in commitment if built (securities, development, construction, financing) [key row] | € m | 147.8 | 171.2 | 167.3 | `Dashboard!H32`, `Out_Commitment_Total_1`, `Out_Commitment_Total_2` (€ m 1 dp) |

Card – Approval requested and timetable:
- Submit both bids by 30-Oct-2026: municipal 50 k€/MW; community 3.00 / 2.01 €/MWh; overbuild 51% / 40% (not built); BESS 51% / 40% of MVA, 4 h; PB at the 1.0x minimum; partial award, shared connection, EU turbines and transformers.  
  <sub>Cells: `In_Bid_Date` (date); `Dashboard!E12`, `Dashboard!E16` (0 dp); `Dashboard!E13`, `Dashboard!F13` (2 dp); `Dashboard!E14`, `Dashboard!F14`, `Dashboard!E15`, `Dashboard!F15` (% 0 dp); `Dashboard!E17` (x 1 dp)</sub>
- Day one, at submission: participation fee €0.008m per bid (cash) and bid guarantee €1.4m / €1.2m.  
  <sub>Cells: `Dashboard!E50` (€ m 3 dp); `Dashboard!E51`, `Dashboard!F51` (€ m 1 dp)</sub>
- After award (13-Dec-2026): PB guarantee €7.0m / €6.1m by 28-Dec-2026, which replaces the bid guarantee; TSO cash security €1.2m per bid by 30-Dec-2026.  
  <sub>Cells: `In_Award_Date`, `Dashboard!I52`, `Dashboard!I53` (date); `Dashboard!E52`, `Dashboard!F52`, `Dashboard!E53` (€ m 1 dp)</sub>
- Before submission: legal confirmation that the two bids sit within the 100 MVA group cap per node.

Footnote: Post-tax, FEP share, EU turbine, base case (overbuild not built, BESS built, PB lost pro rata to the shortfall); RtB exit 31-Dec-2029. Zero bid = every commitment at its zero level. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!E60`, `Dashboard!E6` (date)</sub>

## Slide 3 – Why these commitments

- Section: Commitments
- Action title: Municipal buys 20 points for 0.4 pp of HTM IRR; BESS (1.2 pp) and community (0.6 pp) carry the return cost on BORA 1  
  <sub>Cells: `Dashboard!I12` (points (2 dp if fractional)); `Dashboard!O12`, `Dashboard!O15`, `Dashboard!O13` (pp 1 dp (magnitude))</sub>
- Subtitle: Each commitment is measured alone against the zero bid; the PB stays at the 1.0x minimum because 2.0x adds 5 points for one more PB of guarantee.  
  <sub>Cells: `Dashboard!E17` (x 1 dp)</sub>

Table (values B1 / B2; Δ vs the zero bid):

| Commitment | Offered | Points | Δ RtB profit € m | Δ COD IRR pp | Δ HTM IRR pp | Δ deposits and guarantees € m | Logic | Cells (format) |
|---|---|---:|---:|---:|---:|---:|---|---|
| Municipal one-off (ÖSZ-1) | 50 / 50 k€/MW | 20 / 20 | 0.0 / 0.0 | −1.2 / −1.2 | −0.4 / −0.4 | 0.0 / 0.0 | Paid at COD, after the RtB exit: no FEP cash on the RtB route. | `Dashboard!E12`, `Dashboard!F12` (0 dp); `Dashboard!I12`, `Dashboard!J12` (points (2 dp if fractional)); `Dashboard!K12`, `Dashboard!L12`, `Dashboard!Q12`, `Dashboard!R12` (Δ € m 1 dp); `Dashboard!M12`, `Dashboard!N12`, `Dashboard!O12`, `Dashboard!P12` (pp 1 dp) |
| Community contribution (ÖSZ-2) | 3.00 / 2.01 €/MWh | 15 / 0.75 | 0.0 / 0.0 | −2.6 / −1.7 | −0.6 / −0.4 | 0.0 / 0.0 | Operating cost after COD. At the band floor, BORA 2 earns 0.75 points. | `Dashboard!E13`, `Dashboard!F13` (2 dp); `Dashboard!I13`, `Dashboard!J13` (points (2 dp if fractional)); `Dashboard!K13`, `Dashboard!L13`, `Dashboard!Q13`, `Dashboard!R13` (Δ € m 1 dp); `Dashboard!M13`, `Dashboard!N13`, `Dashboard!O13`, `Dashboard!P13` (pp 1 dp) |
| Overbuild (ÖSZ-3) | 51% / 40% | 10 / 9 | (0.5) / (0.4) | −1.1 / −0.9 | −0.2 / −0.2 | +1.4 / +1.1 | Not built: the cost is the larger PB, lost pro rata to the shortfall. | `Dashboard!E14`, `Dashboard!F14` (% 0 dp); `Dashboard!I14`, `Dashboard!J14` (points (2 dp if fractional)); `Dashboard!K14`, `Dashboard!L14`, `Dashboard!Q14`, `Dashboard!R14` (Δ € m 1 dp); `Dashboard!M14`, `Dashboard!N14`, `Dashboard!O14`, `Dashboard!P14` (pp 1 dp) |
| BESS power (ÖSZ-4), 4 h | 51% / 40% of MVA | 15 / 13.50 | (0.4) / (0.3) | −6.5 / −5.1 | −1.2 / −1.0 | +2.8 / +2.2 | Sized with the overbuild and built; valued at cost, with no BESS revenue. | `Dashboard!E16` (0 dp); `Dashboard!E15`, `Dashboard!F15` (% 0 dp); `Dashboard!I15`, `Dashboard!J15` (points (2 dp if fractional)); `Dashboard!K15`, `Dashboard!L15`, `Dashboard!Q15`, `Dashboard!R15` (Δ € m 1 dp); `Dashboard!M15`, `Dashboard!N15`, `Dashboard!O15`, `Dashboard!P15` (pp 1 dp) |
| PB multiple (ÖSZ-6) | 1.0x / 1.0x | 0 / 0 | 0.0 | 0.0 | 0.0 | 0.0 | Not raised: 2.0x adds 5 points for one more PB of €7.0m / €6.1m of guarantee. | `Dashboard!E17`, `Dashboard!F17` (x 1 dp); `Dashboard!I17`, `Dashboard!J17` (points (2 dp if fractional)); `Dashboard!K17:L17`, `Dashboard!M17:N17`, `Dashboard!O17:P17`, `Dashboard!Q17:R17` (all zero, 1 dp); `Dashboard!E52`, `Dashboard!F52` (€ m 1 dp) |
| Partial award (ÖSZ-5), shared connection (ÖSZ-8), EU turbines / transformers (ÖSZ-9) | yes | 5 / 5 / 12 / 3 each bid | 0.0 | 0.0 | 0.0 | 0.0 | No cash or return impact. EU sourcing is in the base cost (turbine 600 k€/MW). | `Dashboard!I18`, `Dashboard!I20`, `Dashboard!I21`, `Dashboard!I22` (points (2 dp if fractional)); `Dashboard!K18:L22`, `Dashboard!M18:N22`, `Dashboard!O18:P22`, `Dashboard!Q18:R22` (all zero, 1 dp); `Dashboard!V18` (0 dp) |
| Full bid vs zero bid [key row] | | 85 / 68.25 | (0.8) / (0.7) | −11.7 / −9.1 | −2.4 / −1.9 | – | The rows do not add to the total because the commitments interact. | `Out_Points_1`, `Out_Points_2` (points (2 dp if fractional)); `Dashboard!I26`, `Dashboard!J26` (Δ € m 1 dp); `Dashboard!I28`, `Dashboard!J28`, `Dashboard!I30`, `Dashboard!J30` (pp 1 dp) |

Footnote: Δ = zero bid plus that commitment alone, minus the zero bid. Zero levels: municipal 0, community 0, overbuild 0%, BESS 0%, PB 1.0x, yes/no items no. Δ deposits and guarantees = extra security posted, not a PB loss. ÖSZ-6: 5 points at 2.0x (Kiírás X.4 bf, Annex 4). Brownfield (ÖSZ-7) is not offered. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!E6` (date)</sub>

## Slide 4 – Securities and cash to the RtB exit

- Section: Economics – RtB exit
- Action title: At the RtB exit FEP recovers its cash and earns €6.5m / €6.7m, a 1.25x / 1.26x MOIC  
  <sub>Cells: `Out_RtB_Profit_1`, `Out_RtB_Profit_2` (€ m 1 dp); `Out_RtB_MOIC_1`, `Out_RtB_MOIC_2` (x 2 dp)</sub>
- Subtitle: Of the €15.0m exit premium per bid, FEP nets €8.4m / €8.4m and A.RES €3.4m / €3.4m; at a 100 k€/MVA premium FEP's RtB profit falls to €3.0m / €3.2m.  
  <sub>Cells: `Out_RtB_Price_1`, `Out_Exit_FEP_Net_1`, `Out_Exit_FEP_Net_2`, `Out_Exit_ARES_Total_1`, `Out_Exit_ARES_Total_2` (€ m 1 dp); `Inputs!L74` (0 dp, In_RtB_Premium in the "Low premium" set); `Dashboard!L97`, `Dashboard!L98` (`RtB_Profit_1_Stored` / `RtB_Profit_2_Stored`, "Low premium"; € m 1 dp)</sub>

Table (RtB exit waterfall, devex-file order, paid amounts; Total = both bids, Dashboard column H of the same row):

| RtB exit waterfall (paid) | Unit | BORA 1 | BORA 2 | Total | Cells (format) |
|---|---|---:|---:|---:|---|
| Memo: buyer reimbursement (deposits, fees, land) | € m | 21.8 | 21.8 | 43.7 | `Out_Exit_Reimbursement_1`, `Out_Exit_Reimbursement_2`, `Dashboard!H61` (€ m 1 dp) |
| Exit premium (RtB price) | € m | 15.0 | 15.0 | 30.0 | `Out_RtB_Price_1`, `Out_RtB_Price_2`, `Dashboard!H62` (€ m 1 dp) |
| Less sale costs | € m | (0.5) | (0.5) | (1.0) | `Out_WF_Sale_Costs_1`, `Out_WF_Sale_Costs_2`, `Dashboard!H63` (€ m 1 dp) |
| Less FEP claim 12% on deposits and connection fees | € m | (1.6) | (1.6) | (3.1) | `In_Cash_Rate` (% 0 dp); `Out_WF_Return_12_1`, `Out_WF_Return_12_2`, `Dashboard!H64` (€ m 1 dp) |
| Less FEP claim 20% / 2× on devex and land | € m | (3.1) | (3.1) | (6.2) | `In_AtRisk_Rate` (% 0 dp); `In_Min_Multiple` (0 dp ×); `Out_WF_Return_AtRisk_1`, `Out_WF_Return_AtRisk_2`, `Dashboard!H65` (€ m 1 dp) |
| Less FEP claim 6% on guarantees | € m | (1.1) | (1.0) | (2.1) | `In_Guar_Credit` (% 0 dp); `Out_WF_Return_6_1`, `Out_WF_Return_6_2`, `Dashboard!H66` (€ m 1 dp) |
| Less FEP interest claims, total | € m | (5.8) | (5.7) | (11.5) | `Out_Exit_Claims_1`, `Out_Exit_Claims_2`, `Dashboard!H67` (€ m 1 dp) |
| Less DSA | € m | (1.2) | (1.1) | (2.3) | `Out_WF_DSA_1`, `Out_WF_DSA_2`, `Dashboard!H68` (€ m 1 dp) |
| Distributable | € m | 7.5 | 7.7 | 15.2 | `Out_Distributable_1`, `Out_Distributable_2`, `Dashboard!H69` (€ m 1 dp) |
| FEP 70% | € m | 5.2 | 5.4 | 10.7 | `In_FEP_Share` (% 0 dp); `Out_Exit_FEP_70_1`, `Out_Exit_FEP_70_2`, `Dashboard!H70` (€ m 1 dp) |
| FEP-borne devex principal | € m | (2.7) | (2.7) | (5.4) | `Out_Exit_FEP_Devex_1`, `Out_Exit_FEP_Devex_2`, `Dashboard!H71` (€ m 1 dp) |
| FEP net = claims + 70% − devex [key row] | € m | 8.4 | 8.4 | 16.7 | `In_FEP_Share` (% 0 dp); `Out_Exit_FEP_Net_1`, `Out_Exit_FEP_Net_2`, `Dashboard!H72` (€ m 1 dp) |
| A.RES 30% | € m | 2.2 | 2.3 | 4.6 | `Out_Exit_ARES_30_1`, `Out_Exit_ARES_30_2`, `Dashboard!H73` (€ m 1 dp) |
| A.RES total = 30% + DSA [key row] | € m | 3.4 | 3.4 | 6.9 | `Out_Exit_ARES_Total_1`, `Out_Exit_ARES_Total_2`, `Dashboard!H74` (€ m 1 dp) |
| Guarantee fees, signing and DD, tax | € m | (1.8) | (1.6) | (3.4) | `Out_RtB_Route_Costs_1`, `Out_RtB_Route_Costs_2`, `Dashboard!H75` (€ m 1 dp) |
| FEP profit – RtB route [key row] | € m | 6.5 | 6.7 | 13.3 | `Out_RtB_Profit_Buildup_1`, `Out_RtB_Profit_Buildup_2`, `Dashboard!H76` (€ m 1 dp) |

Card – Securities to the exit:
- Bid guarantee €1.4m / €1.2m, 30-Oct-2026 to 28-Dec-2026.  
  <sub>Cells: `Dashboard!E51`, `Dashboard!F51` (€ m 1 dp); `Dashboard!I51`, `Dashboard!J51` (date)</sub>
- PB guarantee €7.0m / €6.1m from 28-Dec-2026; the buyer replaces it at the exit.  
  <sub>Cells: `Dashboard!E52`, `Dashboard!F52` (€ m 1 dp); `Dashboard!I52` (date)</sub>
- Cash: TSO security €1.2m (30-Dec-2026) and connection fees €1.4m + €18.8m (08-Jan-2029). The buyer reimburses deposits, fees and land at the exit (memo line).  
  <sub>Cells: `Dashboard!E53`, `Dashboard!E54`, `Dashboard!E55` (€ m 1 dp); `Dashboard!I53`, `Dashboard!I54` (date)</sub>

Footnote: RtB price 150 k€/MVA × 100 MVA, per the Bora scenarios meeting file (Bora!N5), confirmed 12-Oct-2026. Paid amounts at the RtB exit on 31-Dec-2029, in devex-file order. The participation fee is not recovered. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `In_RtB_Premium`, `In_MVA` (0 dp); `Dash_Exit_Date`, `Dashboard!E6` (date)</sub>

## Slide 5 – COD and HTM value and the commitment breakdown

- Section: Economics – COD and HTM
- Action title: All-in cost of 1,642 / 1,612 k€ per built MW against a 2,450 k€/MW COD price; HTM NPV of €13.4m / €17.2m at 8.0%  
  <sub>Cells: `Out_Cost_To_COD_MW_1`, `Out_Cost_To_COD_MW_2`, `In_COD_Price` (k€/MW 0 dp); `Out_HTM_NPV_1`, `Out_HTM_NPV_2` (€ m 1 dp); `Dashboard!V22` (% 1 dp)</sub>
- Subtitle: The commitments reduce HTM NPV by €18.1m / €14.2m against the zero bid; construction is €150.9m / €148.3m of the all-in commitment.  
  <sub>Cells: `Dashboard!I31`, `Dashboard!J31` (€ m 1 dp (magnitude)); `Out_Commitment_Construction_1`, `Out_Commitment_Construction_2` (€ m 1 dp)</sub>

Table:

| Item | Unit | Zero bid | BORA 1 | BORA 2 | Δ B1 | Δ B2 | Cells (format) |
|---|---|---:|---:|---:|---:|---:|---|
| FEP NPV – sale at COD | € m | 44.6 | 30.8 | 33.6 | (13.9) | (11.0) | `Dashboard!H29`, `Out_COD_NPV_1`, `Out_COD_NPV_2` (€ m 1 dp); `Dashboard!I29`, `Dashboard!J29` (Δ € m 1 dp) |
| FEP NPV – hold to maturity [key row] | € m | 31.4 | 13.4 | 17.2 | (18.1) | (14.2) | `Dashboard!H31`, `Out_HTM_NPV_1`, `Out_HTM_NPV_2` (€ m 1 dp); `Dashboard!I31`, `Dashboard!J31` (Δ € m 1 dp) |
| All-in cost to COD per built MW | k€/MW | – | 1,642 | 1,612 | – | – | `Out_Cost_To_COD_MW_1`, `Out_Cost_To_COD_MW_2` (k€/MW 0 dp) |
| COD sale price per built MW | k€/MW | – | 2,450 | 2,450 | – | – | `In_COD_Price` (k€/MW 0 dp) |
| Securities: cash posted and guarantees | € m | – | 8.2 | 7.3 | – | – | `Out_Commitment_Securities_1`, `Out_Commitment_Securities_2` (€ m 1 dp) |
| Devex, land, signing and DD | € m | – | 3.5 | 3.5 | – | – | `Out_Commitment_Devex_1`, `Out_Commitment_Devex_2` (€ m 1 dp) |
| Construction capex incl. connection fees, municipal and pre-COD lease | € m | – | 150.9 | 148.3 | – | – | `Out_Commitment_Construction_1`, `Out_Commitment_Construction_2` (€ m 1 dp) |
| Financing: IDC, debt fees, guarantee fees | € m | – | 8.5 | 8.2 | – | – | `Out_Commitment_Financing_1`, `Out_Commitment_Financing_2` (€ m 1 dp) |
| Memo, not in the commitment: PB lost pro rata, overbuild not built | € m | – | 1.4 | 1.1 | – | – | `Out_Commitment_Penalty_1`, `Out_Commitment_Penalty_2` (€ m 1 dp) |

Card:
- Sale at COD: FEP proceeds of €124.4m / €127.2m on a €245.0m price, a 1.78x / 1.85x MOIC.  
  <sub>Cells: `Out_COD_Proceeds_1`, `Out_COD_Proceeds_2`, `Out_COD_Price_1` (€ m 1 dp); `Out_COD_MOIC_1`, `Out_COD_MOIC_2` (x 2 dp)</sub>
- Hold to maturity: 4.89x / 5.13x MOIC; minimum DSCR 1.33x / 1.37x.  
  <sub>Cells: `Out_HTM_MOIC_1`, `Out_HTM_MOIC_2`, `Out_Min_DSCR_1`, `Out_Min_DSCR_2` (x 2 dp)</sub>

Footnote: NPV at 8.0%. COD 30-Sep-2031, 100 MW built; PPA 90 €/MWh; gearing 60.0%. HTM includes the buy-out of the developer at the RtB exit date. The four breakdown lines sum to the all-in commitment on slide 2. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!V22`, `Dashboard!V20` (% 1 dp); `Dashboard!V23`, `Dashboard!E6` (date); `Out_Wind_Built_MW_1`, `Dashboard!V19` (0 dp)</sub>

## Slide 6 – Risks and mitigants

- Section: Risks and mitigants
- Action title: The main risk is the unbuilt overbuild: the PB is lost pro rata (€1.4m / €1.1m) with 33.8 / 28.6 MVA of connection right; a full call is unresolved  
  <sub>Cells: `Out_Commitment_Penalty_1`, `Out_Commitment_Penalty_2` (€ m 1 dp); `Out_MVA_At_Risk_1`, `Out_MVA_At_Risk_2` (1 dp)</sub>
- Subtitle: The base COD and HTM returns include the pro-rata PB loss; the lost MVA, any RtB price deduction and a possible group exclusion are not in the returns.

Table:

| Risk | Exposure (B1 / B2) | Mitigant | Cells (format) |
|---|---|---|---|
| Overbuild not built: PB lost pro rata to the shortfall above 10% (Kiírás XII.12; GYIK Q28/30/44) – base | €1.4m / €1.1m, the PB on the extra MW. | PB at the 1.0x minimum. Part of the loss may be deducted from the RtB price. | `Out_Commitment_Penalty_1`, `Out_Commitment_Penalty_2` (€ m 1 dp); `Dashboard!E17` (x 1 dp) |
| Full PB call if the points-commitment rule (Kiírás XIII.17; GYIK Q72) also applies to the overbuild – unresolved | €7.0m / €6.1m: COD IRR 35.8% / 38.7%, HTM IRR 9.1% / 9.5%. | Downside case, not the base. PB at the minimum limits it. | `Out_Sanction_Full_1`, `Out_Sanction_Full_2` (€ m 1 dp); `Dashboard!J99`, `Dashboard!J100`, `Dashboard!J101`, `Dashboard!J102` (% 1 dp) |
| 3-year group exclusion (Kiírás VII.2 n; GYIK Q71) | Possible if the PB is called; unclear whether a pro-rata loss counts as a call. A call after COD would likely hit the group then holding the SPV. | Allocate it in the RtB SPA. | |
| Connection right lost pro rata (GYIK Q28/30) | 33.8 / 28.6 MVA, in the pro-rata reading only (GYIK Q72: a full call keeps the right). Not reflected in generation or price. | The buyer can build the overbuild (card). | `Out_MVA_At_Risk_1`, `Out_MVA_At_Risk_2` (1 dp) |
| 100 MVA group cap per node (Kiírás V.7; GYIK Q52) | Both bids are invalid if FEP, A.RES and the SPVs count as one company group. | Legal confirmation before submission. | |
| ÖSZ-4 scored on installed wind (GYIK Q20); BESS valued at cost only | Points 80.50 / 63.75, returns unchanged. BESS costs −6.5 / −5.1 pp of COD IRR with no BESS revenue. | Ranking risk only. Any BESS revenue is upside. | `Dashboard!M95`, `Dashboard!M96` (points (2 dp if fractional)); `Dashboard!M15`, `Dashboard!N15` (pp 1 dp) |

Card – Why the base does not build the overbuild:
- The grid connection stays at 100 MVA. Building the overbuild lifts the all-in cost to COD from €164.2m / €161.2m to €217.9m / €203.3m. In that case COD IRR is 54.4% / 53.8% and HTM IRR is 11.5% / 11.6%. The build decision passes to the buyer at RtB.  
  <sub>Cells: `In_MVA` (0 dp); `Out_Cost_To_COD_1`, `Out_Cost_To_COD_2`, `Dashboard!I105`, `Dashboard!I106` (€ m 1 dp); `Dashboard!I99`, `Dashboard!I100`, `Dashboard!I101`, `Dashboard!I102` (% 1 dp)</sub>

Footnote: Shortfall tested 6 months after COD. Pro rata = shortfall share × the wind PB (GYIK Q28/30). The full-call and overbuild-built figures are alternative cases. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!E6` (date)</sub>
