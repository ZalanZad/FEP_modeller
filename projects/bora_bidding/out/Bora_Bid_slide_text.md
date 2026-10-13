# Project Bora – IC deck text (6 slides)

Production notes (not deck text): one strategy for both bids (SPEC §0 item 21). Per-bid figures come from the BORA 1 cells, which equal BORA 2 (master check 0). "Both bids" = the Dashboard Total column where it exists, else BORA 1 + BORA 2 cells (both listed). Each line ends with its cells and display format; the full list is in `Bora_Bid_slide_text_figures.csv`. Formats: € m 1 dp with negatives in brackets; signed Δ € m; % 1 dp with a true minus; pp 1 dp, signed; points as integers, or 2 dp where fractional; k€/MW 0 dp with a thousands separator; MOIC 2 dp. [key row] = bold on `E6F4EC`.  
  <sub>Cells: `Master_Check` (0 dp)</sub>

## Slide 1 – Cover

- Title: Project Bora – BORA 1 and BORA 2 capacity bids
- Subtitle: Approval of the tender commitments, 2 × 100 MVA, Szombathely OVIT
- Date: 12-Oct-2026  
  <sub>Cells: `Dashboard!E6` (date)</sub>

## Slide 2 – Decision and timetable

- Section: Executive summary
- Action title: Submit both bids with the same commitments at 85 points each: FEP earns €13.1m at the RtB exit, a 22.2% IRR  
  <sub>Cells: `Out_Points_1` (points (2 dp if fractional)); `Dashboard!H76` (€ m 1 dp); `Out_RtB_IRR_1` (% 1 dp)</sub>
- Subtitle: Against the zero bid, the commitments cost €0.8m of RtB profit and 2.4 pp of HTM IRR per bid, and add €23.4m per bid to the all-in commitment.  
  <sub>Cells: `Dashboard!I26`, `Dashboard!I32` (€ m 1 dp (magnitude)); `Dashboard!I30` (pp 1 dp (magnitude))</sub>

Table (returns vs the zero bid):

| Item | Unit | Zero bid, per bid | Per bid | Both bids | Cells (format) |
|---|---|---:|---:|---:|---|
| Points (EU turbine) | points | 0 | 85 | – | `Dashboard!H25`, `Out_Points_1` (points (2 dp if fractional)) |
| FEP profit – RtB route | € m | 7.4 | 6.5 | 13.1 | `Dashboard!H26`, `Out_RtB_Profit_1`, `Dashboard!H76` (€ m 1 dp) |
| FEP IRR – RtB route | % | 25.8% | 22.2% | – | `Dashboard!H27`, `Out_RtB_IRR_1` (% 1 dp) |
| FEP IRR – sale at COD | % | 49.6% | 37.8% | – | `Dashboard!H28`, `Out_COD_IRR_1` (% 1 dp) |
| FEP IRR – hold to maturity | % | 11.9% | 9.5% | – | `Dashboard!H30`, `Out_HTM_IRR_1` (% 1 dp) |
| All-in commitment if built (securities, development, construction, financing) [key row] | € m | 147.8 | 171.2 | 342.4 | `Dashboard!H32`, `Out_Commitment_Total_1`, `Out_Commitment_Total_1 + Out_Commitment_Total_2` (€ m 1 dp) |

Card – Approval requested and timetable:
- Submit BORA 1 and BORA 2 by 30-Oct-2026, each with: municipal 50 k€/MW; community 3.00 €/MWh; overbuild 51% (not built); BESS 51% of MVA, 4 h; PB at the 1.0x minimum; partial award, shared connection, EU turbines and transformers.  
  <sub>Cells: `In_Bid_Date` (date); `Dashboard!E12`, `Dashboard!E16` (0 dp); `Dashboard!E13` (2 dp); `Dashboard!E14`, `Dashboard!E15` (% 0 dp); `Dashboard!E17` (x 1 dp)</sub>
- Day one, at submission (both bids): participation fees €0.016m in cash and bid guarantees of €2.8m.  
  <sub>Cells: `Dashboard!E50 + Dashboard!F50` (€ m 3 dp); `Dashboard!E51 + Dashboard!F51` (€ m 1 dp)</sub>
- After award (13-Dec-2026), both bids: PB guarantees of €14.0m by 28-Dec-2026, replacing the bid guarantees; TSO cash security of €2.5m by 30-Dec-2026.  
  <sub>Cells: `In_Award_Date`, `Dashboard!I52`, `Dashboard!I53` (date); `Dashboard!E52 + Dashboard!F52`, `Dashboard!E53 + Dashboard!F53` (€ m 1 dp)</sub>

Footnote: Post-tax, FEP share, EU turbine, base case (overbuild not built, BESS built, PB lost pro rata to the shortfall); RtB exit 31-Dec-2029. Zero bid = every commitment at its zero level. The two bids are identical, so IRRs are the same per bid and for both. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!E60`, `Dashboard!E6` (date)</sub>

## Slide 3 – Why these commitments

- Section: Commitments
- Action title: Municipal buys 20 points for 0.4 pp of HTM IRR; BESS (1.2 pp) and community (0.6 pp) carry the return cost  
  <sub>Cells: `Dashboard!I12` (points (2 dp if fractional)); `Dashboard!O12`, `Dashboard!O15`, `Dashboard!O13` (pp 1 dp (magnitude))</sub>
- Subtitle: Each commitment is measured alone against the zero bid, per bid; the PB stays at the minimum because 2.0x adds 5 points for one more PB of guarantee.

Table (per bid; Δ vs the zero bid):

| Commitment | Offered | Points | Δ RtB profit € m | Δ COD IRR pp | Δ HTM IRR pp | Δ deposits and guarantees € m | Logic | Cells (format) |
|---|---|---:|---:|---:|---:|---:|---|---|
| Municipal one-off (ÖSZ-1) | 50 k€/MW | 20 | 0.0 | −1.2 | −0.4 | 0.0 | Paid at COD, after the RtB exit: no FEP cash on the RtB route. | `Dashboard!E12` (0 dp); `Dashboard!I12` (points (2 dp if fractional)); `Dashboard!K12`, `Dashboard!Q12` (Δ € m 1 dp); `Dashboard!M12`, `Dashboard!O12` (pp 1 dp) |
| Community contribution (ÖSZ-2) | 3.00 €/MWh | 15 | 0.0 | −2.6 | −0.6 | 0.0 | Operating cost after COD. | `Dashboard!E13` (2 dp); `Dashboard!I13` (points (2 dp if fractional)); `Dashboard!K13`, `Dashboard!Q13` (Δ € m 1 dp); `Dashboard!M13`, `Dashboard!O13` (pp 1 dp) |
| Overbuild (ÖSZ-3) | 51% | 10 | (0.5) | −1.1 | −0.2 | +1.4 | Not built: the cost is the larger PB, lost pro rata to the shortfall. | `Dashboard!E14` (% 0 dp); `Dashboard!I14` (points (2 dp if fractional)); `Dashboard!K14`, `Dashboard!Q14` (Δ € m 1 dp); `Dashboard!M14`, `Dashboard!O14` (pp 1 dp) |
| BESS power (ÖSZ-4), 4 h | 51% of MVA | 15 | (0.4) | −6.5 | −1.2 | +2.8 | Sized with the overbuild and built; valued at cost, with no BESS revenue. | `Dashboard!E16` (0 dp); `Dashboard!E15` (% 0 dp); `Dashboard!I15` (points (2 dp if fractional)); `Dashboard!K15`, `Dashboard!Q15` (Δ € m 1 dp); `Dashboard!M15`, `Dashboard!O15` (pp 1 dp) |
| PB multiple (ÖSZ-6) | 1.0x | 0 | 0.0 | 0.0 | 0.0 | 0.0 | Not raised: 2.0x adds 5 points for one more PB of €7.0m of guarantee. | `Dashboard!E17` (x 1 dp); `Dashboard!I17` (points (2 dp if fractional)); `Dashboard!K17`, `Dashboard!M17`, `Dashboard!O17`, `Dashboard!Q17` (all zero, 1 dp); `Dashboard!E52` (€ m 1 dp) |
| Partial award (ÖSZ-5), shared connection (ÖSZ-8), EU turbines / transformers (ÖSZ-9) | yes | 5 / 5 / 12 / 3 | 0.0 | 0.0 | 0.0 | 0.0 | No cash or return impact. EU sourcing is in the base cost (turbine 600 k€/MW). | `Dashboard!I18`, `Dashboard!I20`, `Dashboard!I21`, `Dashboard!I22` (points (2 dp if fractional)); `Dashboard!K18:K22`, `Dashboard!M18:M22`, `Dashboard!O18:O22`, `Dashboard!Q18:Q22` (all zero, 1 dp); `Dashboard!V18` (0 dp) |
| Full bid vs zero bid [key row] | | 85 | (0.8) | −11.7 | −2.4 | – | The rows do not add to the total because the commitments interact. | `Out_Points_1` (points (2 dp if fractional)); `Dashboard!I26` (Δ € m 1 dp); `Dashboard!I28`, `Dashboard!I30` (pp 1 dp) |

Footnote: Δ = zero bid plus that commitment alone, minus the zero bid. Zero levels: municipal 0, community 0, overbuild 0%, BESS 0%, PB 1.0x, yes/no items no. Δ deposits and guarantees = extra security posted, not a PB loss. ÖSZ-6: 5 points at 2.0x (Kiírás X.4 bf, Annex 4). Brownfield (ÖSZ-7) is not offered. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!E6` (date)</sub>

## Slide 4 – RtB exit

- Section: Economics – RtB exit
- Action title: At the RtB exit FEP earns €13.1m on both bids, a 1.25x MOIC; its interest claims are paid ahead of the 70/30 split  
  <sub>Cells: `Dashboard!H76` (€ m 1 dp); `Out_RtB_MOIC_1` (x 2 dp)</sub>
- Subtitle: FEP keeps 55 k€ and A.RES 23 k€ per installed wind MW; at a 100 k€/MVA RtB price, FEP's profit falls to €6.1m for both bids.  
  <sub>Cells: `Out_Exit_FEP_Net_MW_1`, `Out_Exit_ARES_Total_MW_1` (k€/MW 0 dp); `Inputs!L74` (0 dp); `Dashboard!L97 + Dashboard!L98` (€ m 1 dp)</sub>

Table (devex-file waterfall at the exit, 31-Dec-2029):  
  <sub>Cells: `Dashboard!E60` (date)</sub>

| Item | Unit | Per bid | Both bids | Cells (format) |
|---|---|---:|---:|---|
| RtB price | € m | 15.0 | 30.0 | `Out_RtB_Price_1`, `Dashboard!H62` (€ m 1 dp) |
| Less sale costs | € m | (0.5) | (1.0) | `Out_WF_Sale_Costs_1`, `Dashboard!H63` (€ m 1 dp) |
| Less FEP interest claims (12% deposits, 20% / 2x devex, 6% guarantees) | € m | (5.8) | (11.6) | `Out_Exit_Claims_1`, `Dashboard!H67` (€ m 1 dp) |
| Less DSA | € m | (1.2) | (2.4) | `Out_WF_DSA_1`, `Dashboard!H68` (€ m 1 dp) |
| Distributable | € m | 7.5 | 15.0 | `Out_Distributable_1`, `Dashboard!H69` (€ m 1 dp) |
| FEP net: interest claims + 70%, less FEP-borne devex | € m | 8.4 | 16.7 | `Out_Exit_FEP_Net_1`, `Dashboard!H72` (€ m 1 dp) |
| A.RES: 30% + DSA | € m | 3.4 | 6.9 | `Out_Exit_ARES_Total_1`, `Dashboard!H74` (€ m 1 dp) |
| Less guarantee fees, signing and DD, tax | € m | (1.8) | (3.6) | `Out_RtB_Route_Costs_1`, `Dashboard!H75` (€ m 1 dp) |
| FEP profit – RtB route [key row] | € m | 6.5 | 13.1 | `Out_RtB_Profit_Buildup_1`, `Dashboard!H76` (€ m 1 dp) |

Card – Cash to the exit:
- FEP funds the deposits, connection fees and land; the buyer reimburses €43.7m for both bids at the exit.  
  <sub>Cells: `Dashboard!H61` (€ m 1 dp)</sub>
- The largest outlay is the connection fees, €2.8m direct and €37.6m indirect, paid on 08-Jan-2029.  
  <sub>Cells: `Dashboard!E54 + Dashboard!F54`, `Dashboard!E55 + Dashboard!F55` (€ m 1 dp); `Dashboard!I54` (date)</sub>
- The buyer replaces the PB guarantee at the exit.

Footnote: RtB price 150 k€/MVA × 100 MVA per bid, per the Bora scenarios meeting file (Bora!N5), confirmed 12-Oct-2026. Per-MW figures use 151 MW of installed wind per bid. The participation fee is not recovered. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `In_RtB_Premium`, `In_MVA`, `Out_Wind_MW_1` (0 dp); `Dashboard!E6` (date)</sub>

## Slide 5 – COD and HTM value and the commitment breakdown

- Section: Economics – COD and HTM
- Action title: All-in cost of 1,642 k€ per built MW against a 2,450 k€/MW COD price; HTM NPV of €26.7m for both bids at 8.0%  
  <sub>Cells: `Out_Cost_To_COD_MW_1`, `In_COD_Price` (k€/MW 0 dp); `Out_HTM_NPV_1 + Out_HTM_NPV_2` (€ m 1 dp); `Dashboard!V22` (% 1 dp)</sub>
- Subtitle: Against the zero bid, the commitments reduce FEP's NPV by €13.9m on a sale at COD and €18.1m held to maturity, per bid.  
  <sub>Cells: `Dashboard!I29`, `Dashboard!I31` (€ m 1 dp (magnitude))</sub>

Table:

| Item | Unit | Zero bid, per bid | Per bid | Δ per bid | Both bids | Cells (format) |
|---|---|---:|---:|---:|---:|---|
| FEP NPV – sale at COD | € m | 44.6 | 30.8 | (13.9) | 61.5 | `Dashboard!H29`, `Out_COD_NPV_1`, `Out_COD_NPV_1 + Out_COD_NPV_2` (€ m 1 dp); `Dashboard!I29` (Δ € m 1 dp) |
| FEP NPV – hold to maturity [key row] | € m | 31.4 | 13.4 | (18.1) | 26.7 | `Dashboard!H31`, `Out_HTM_NPV_1`, `Out_HTM_NPV_1 + Out_HTM_NPV_2` (€ m 1 dp); `Dashboard!I31` (Δ € m 1 dp) |
| All-in cost to COD per built MW | k€/MW | – | 1,642 | – | – | `Out_Cost_To_COD_MW_1` (k€/MW 0 dp) |
| COD sale price per built MW | k€/MW | – | 2,450 | – | – | `In_COD_Price` (k€/MW 0 dp) |
| Securities: cash posted and guarantees | € m | – | 8.2 | – | 16.5 | `Out_Commitment_Securities_1`, `Out_Commitment_Securities_1 + Out_Commitment_Securities_2` (€ m 1 dp) |
| Devex, land, signing and DD | € m | – | 3.5 | – | 7.0 | `Out_Commitment_Devex_1`, `Out_Commitment_Devex_1 + Out_Commitment_Devex_2` (€ m 1 dp) |
| Construction capex incl. connection fees, municipal and pre-COD lease | € m | – | 150.9 | – | 301.8 | `Out_Commitment_Construction_1`, `Out_Commitment_Construction_1 + Out_Commitment_Construction_2` (€ m 1 dp) |
| Financing: IDC, debt fees, guarantee fees | € m | – | 8.5 | – | 17.1 | `Out_Commitment_Financing_1`, `Out_Commitment_Financing_1 + Out_Commitment_Financing_2` (€ m 1 dp) |

Card (per bid):
- Sale at COD: FEP proceeds of €124.4m on a €245.0m price, a 1.78x MOIC.  
  <sub>Cells: `Out_COD_Proceeds_1`, `Out_COD_Price_1` (€ m 1 dp); `Out_COD_MOIC_1` (x 2 dp)</sub>
- Hold to maturity: 4.89x MOIC; minimum DSCR 1.33x.  
  <sub>Cells: `Out_HTM_MOIC_1`, `Out_Min_DSCR_1` (x 2 dp)</sub>

Footnote: NPV at 8.0%. COD 30-Sep-2031, 100 MW built per bid; PPA 90 €/MWh; gearing 60.0%. HTM includes the buy-out of the developer at the RtB exit date. The four breakdown lines sum to the all-in commitment on slide 2. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!V22`, `Dashboard!V20` (% 1 dp); `Dashboard!V23`, `Dashboard!E6` (date); `Out_Wind_Built_MW_1`, `Dashboard!V19` (0 dp)</sub>

## Slide 6 – Risks and mitigants

- Section: Risks and mitigants
- Action title: The main risk is the unbuilt overbuild: €2.8m of PB lost pro rata and 67.5 MVA of connection right across both bids; a full call is unresolved  
  <sub>Cells: `Out_Commitment_Penalty_1 + Out_Commitment_Penalty_2` (€ m 1 dp); `Out_MVA_At_Risk_1 + Out_MVA_At_Risk_2` (1 dp)</sub>
- Subtitle: The base COD and HTM returns include the pro-rata PB loss; the lost MVA, any RtB price deduction and a possible group exclusion are not in the returns.

Table (per bid):

| Risk | Exposure | Mitigant | Cells (format) |
|---|---|---|---|
| Overbuild not built: PB lost pro rata to the shortfall above 10% (Kiírás XII.12; GYIK Q28/30/44) – base | €1.4m, the PB on the extra MW. | PB at the 1.0x minimum. Part of the loss may be deducted from the RtB price. | `Out_Commitment_Penalty_1` (€ m 1 dp); `Dashboard!E17` (x 1 dp) |
| Full PB call if the points-commitment rule (Kiírás XIII.17; GYIK Q72) also applies to the overbuild – unresolved | The whole PB: COD IRR 35.8%, HTM IRR 9.1%. | Downside case, not the base. PB at the minimum limits it. | `Dashboard!J99`, `Dashboard!J101` (% 1 dp) |
| 3-year group exclusion (Kiírás VII.2 n; GYIK Q71) | Possible if the PB is called; unclear whether a pro-rata loss counts as a call. A call after COD would likely hit the group then holding the SPV. | Allocate it in the RtB SPA. | |
| Connection right lost pro rata (GYIK Q28/30) | 33.8 MVA, in the pro-rata reading only (GYIK Q72: a full call keeps the right). Not reflected in generation or price. | The buyer can build the overbuild (card). | `Out_MVA_At_Risk_1` (1 dp) |
| ÖSZ-4 scored on installed wind (GYIK Q20) | Points fall to 80.50; returns unchanged. | Ranking risk only; the commitment is the same. | `Dashboard!M95` (points (2 dp if fractional)) |
| BESS valued at cost only | Its return cost (slide 3) excludes any BESS revenue or sale value. | Conservative: any BESS revenue is upside. | |

Card – Why the base does not build the overbuild:
- The grid connection stays at 100 MVA. Building the overbuild lifts the all-in cost to COD from €164.2m to €217.9m per bid. In that case COD IRR is 54.4% and HTM IRR is 11.5%. The build decision passes to the buyer at RtB.  
  <sub>Cells: `In_MVA` (0 dp); `Out_Cost_To_COD_1`, `Dashboard!I105` (€ m 1 dp); `Dashboard!I99`, `Dashboard!I101` (% 1 dp)</sub>

Footnote: Shortfall tested 6 months after COD. Pro rata = shortfall share × the wind PB (GYIK Q28/30). The full-call and overbuild-built figures are alternative cases. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!E6` (date)</sub>
