# Project Bora – IC deck text (7 slides)

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
  <sub>Cells: `Out_Points_1` (points (2 dp if fractional)); `Dashboard!H103` (€ m 1 dp); `Out_RtB_IRR_1` (% 1 dp)</sub>
- Subtitle: Against the zero bid, the commitments cost €0.8m of RtB profit and 3.0 pp of HTM IRR per bid, and add €31.9m per bid to the all-in commitment.  
  <sub>Cells: `Dashboard!I26`, `Dashboard!I32` (€ m 1 dp (magnitude)); `Dashboard!I30` (pp 1 dp (magnitude))</sub>

Table (returns vs the zero bid):

| Item | Unit | Zero bid, per bid | Per bid | Both bids | Cells (format) |
|---|---|---:|---:|---:|---|
| Points (EU turbine) | points | 0 | 85 | – | `Dashboard!H25`, `Out_Points_1` (points (2 dp if fractional)) |
| FEP profit – RtB route | € m | 7.4 | 6.5 | 13.1 | `Dashboard!H26`, `Out_RtB_Profit_1`, `Dashboard!H103` (€ m 1 dp) |
| FEP IRR – RtB route | % | 25.8% | 22.2% | – | `Dashboard!H27`, `Out_RtB_IRR_1` (% 1 dp) |
| FEP IRR – sale at COD | % | 49.6% | 33.9% | – | `Dashboard!H28`, `Out_COD_IRR_1` (% 1 dp) |
| FEP IRR – hold to maturity | % | 11.9% | 8.9% | – | `Dashboard!H30`, `Out_HTM_IRR_1` (% 1 dp) |
| All-in commitment if built (securities, development, construction, financing) [key row] | € m | 147.8 | 179.7 | 359.5 | `Dashboard!H32`, `Out_Commitment_Total_1`, `Out_Commitment_Total_1 + Out_Commitment_Total_2` (€ m 1 dp) |

Card – Approval requested and timetable:
- Submit BORA 1 and BORA 2 by 30-Oct-2026, each with: municipal 50 k€/MW; community 3.00 €/MWh; overbuild 51% (not built); BESS 51% of MVA, 4 h; PB at the 1.0x minimum; partial award, shared connection, EU turbines and transformers.  
  <sub>Cells: `In_Bid_Date` (date); `Dashboard!E12`, `Dashboard!E16` (0 dp); `Dashboard!E13` (2 dp); `Dashboard!E14`, `Dashboard!E15` (% 0 dp); `Dashboard!E17` (x 1 dp)</sub>
- Day one, at submission (both bids): participation fees €0.016m in cash and bid guarantees of €2.8m.  
  <sub>Cells: `Dashboard!E74 + Dashboard!F74` (€ m 3 dp); `Dashboard!E75 + Dashboard!F75` (€ m 1 dp)</sub>
- After award (13-Dec-2026), both bids: PB guarantees of €14.0m by 28-Dec-2026, replacing the bid guarantees; TSO cash security of €2.5m by 30-Dec-2026.  
  <sub>Cells: `In_Award_Date`, `Dashboard!I76`, `Dashboard!I77` (date); `Dashboard!E76 + Dashboard!F76`, `Dashboard!E77 + Dashboard!F77` (€ m 1 dp)</sub>

Footnote: Post-tax, FEP share, EU turbine, base case (overbuild not built, BESS built, PB lost pro rata to the shortfall); RtB exit 31-Dec-2029. Zero bid = every commitment at its zero level. The two bids are identical, so IRRs are the same per bid and for both. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!E87`, `Dashboard!E6` (date)</sub>

## Slide 3 – Why these commitments

- Section: Commitments
- Action title: Municipal buys 20 points for 0.4 pp of HTM IRR; BESS (1.9 pp) and community (0.6 pp) carry the return cost  
  <sub>Cells: `Dashboard!I12` (points (2 dp if fractional)); `Dashboard!O12`, `Dashboard!O15`, `Dashboard!O13` (pp 1 dp (magnitude))</sub>
- Subtitle: Each commitment is measured alone against the zero bid, per bid; the PB stays at the minimum because 2.0x adds 5 points for one more PB of guarantee.

Table (per bid; Δ vs the zero bid):

| Commitment | Offered | Points | Δ RtB profit € m | Δ COD IRR pp | Δ HTM IRR pp | Δ deposits and guarantees € m | Logic | Cells (format) |
|---|---|---:|---:|---:|---:|---:|---|---|
| Municipal one-off (ÖSZ-1) | 50 k€/MW | 20 | 0.0 | −1.2 | −0.4 | 0.0 | Paid at COD, after the RtB exit: no FEP cash on the RtB route. | `Dashboard!E12` (0 dp); `Dashboard!I12` (points (2 dp if fractional)); `Dashboard!K12`, `Dashboard!Q12` (Δ € m 1 dp); `Dashboard!M12`, `Dashboard!O12` (pp 1 dp) |
| Community contribution (ÖSZ-2) | 3.00 €/MWh | 15 | 0.0 | −2.6 | −0.6 | 0.0 | Operating cost after COD. | `Dashboard!E13` (2 dp); `Dashboard!I13` (points (2 dp if fractional)); `Dashboard!K13`, `Dashboard!Q13` (Δ € m 1 dp); `Dashboard!M13`, `Dashboard!O13` (pp 1 dp) |
| Overbuild (ÖSZ-3) | 51% | 10 | (0.5) | −1.1 | −0.2 | +1.4 | Not built: the cost is the larger PB, lost pro rata to the shortfall. | `Dashboard!E14` (% 0 dp); `Dashboard!I14` (points (2 dp if fractional)); `Dashboard!K14`, `Dashboard!Q14` (Δ € m 1 dp); `Dashboard!M14`, `Dashboard!O14` (pp 1 dp) |
| BESS power (ÖSZ-4), 4 h | 51% of MVA | 15 | (0.4) | −10.2 | −1.9 | +2.8 | Sized with the overbuild and built; valued at cost, with no BESS revenue. | `Dashboard!E16` (0 dp); `Dashboard!E15` (% 0 dp); `Dashboard!I15` (points (2 dp if fractional)); `Dashboard!K15`, `Dashboard!Q15` (Δ € m 1 dp); `Dashboard!M15`, `Dashboard!O15` (pp 1 dp) |
| PB multiple (ÖSZ-6) | 1.0x | 0 | 0.0 | 0.0 | 0.0 | 0.0 | Not raised: 2.0x adds 5 points for one more PB of €7.0m of guarantee. | `Dashboard!E17` (x 1 dp); `Dashboard!I17` (points (2 dp if fractional)); `Dashboard!K17`, `Dashboard!M17`, `Dashboard!O17`, `Dashboard!Q17` (all zero, 1 dp); `Dashboard!E76` (€ m 1 dp) |
| Partial award (ÖSZ-5), shared connection (ÖSZ-8), EU turbines / transformers (ÖSZ-9) | yes | 5 / 5 / 12 / 3 | 0.0 | 0.0 | 0.0 | 0.0 | No cash or return impact. EU sourcing is in the base cost (turbine 600 k€/MW). | `Dashboard!I18`, `Dashboard!I20`, `Dashboard!I21`, `Dashboard!I22` (points (2 dp if fractional)); `Dashboard!K18:K22`, `Dashboard!M18:M22`, `Dashboard!O18:O22`, `Dashboard!Q18:Q22` (all zero, 1 dp); `Dashboard!V18` (0 dp) |
| Full bid vs zero bid [key row] | | 85 | (0.8) | −15.6 | −3.0 | – | The rows do not add to the total because the commitments interact. | `Out_Points_1` (points (2 dp if fractional)); `Dashboard!I26` (Δ € m 1 dp); `Dashboard!I28`, `Dashboard!I30` (pp 1 dp) |

Footnote: Δ = zero bid plus that commitment alone, minus the zero bid. Zero levels: municipal 0, community 0, overbuild 0%, BESS 0%, PB 1.0x, yes/no items no. Δ deposits and guarantees = extra security posted, not a PB loss. ÖSZ-6: 5 points at 2.0x (Kiírás X.4 bf, Annex 4). Brownfield (ÖSZ-7) is not offered. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!E6` (date)</sub>

## Slide 4 – RtB exit

- Section: Economics – RtB exit
- Action title: At the RtB exit FEP earns €13.1m on both bids, a 1.25x MOIC; its interest claims are paid ahead of the 70/30 split  
  <sub>Cells: `Dashboard!H103` (€ m 1 dp); `Out_RtB_MOIC_1` (x 2 dp)</sub>
- Subtitle: FEP keeps 55 k€ and A.RES 23 k€ per installed wind MW; at a 100 k€/MVA RtB price, FEP's profit falls to €6.1m for both bids.  
  <sub>Cells: `Out_Exit_FEP_Net_MW_1`, `Out_Exit_ARES_Total_MW_1` (k€/MW 0 dp); `Inputs!L74` (0 dp); `Dashboard!L123 + Dashboard!L124` (€ m 1 dp)</sub>

Table (devex-file waterfall at the exit, 31-Dec-2029):  
  <sub>Cells: `Dashboard!E87` (date)</sub>

| Item | Unit | Per bid | Both bids | Cells (format) |
|---|---|---:|---:|---|
| RtB price | € m | 15.0 | 30.0 | `Out_RtB_Price_1`, `Dashboard!H89` (€ m 1 dp) |
| Less sale costs | € m | (0.5) | (1.0) | `Out_WF_Sale_Costs_1`, `Dashboard!H90` (€ m 1 dp) |
| Less FEP interest claims (12% deposits, 20% / 2x devex, 6% guarantees) | € m | (5.8) | (11.6) | `Out_Exit_Claims_1`, `Dashboard!H94` (€ m 1 dp) |
| Less DSA | € m | (1.2) | (2.4) | `Out_WF_DSA_1`, `Dashboard!H95` (€ m 1 dp) |
| Distributable | € m | 7.5 | 15.0 | `Out_Distributable_1`, `Dashboard!H96` (€ m 1 dp) |
| FEP net: interest claims + 70%, less FEP-borne devex | € m | 8.4 | 16.7 | `Out_Exit_FEP_Net_1`, `Dashboard!H99` (€ m 1 dp) |
| A.RES: 30% + DSA | € m | 3.4 | 6.9 | `Out_Exit_ARES_Total_1`, `Dashboard!H101` (€ m 1 dp) |
| Less guarantee fees, signing and DD, tax | € m | (1.8) | (3.6) | `Out_RtB_Route_Costs_1`, `Dashboard!H102` (€ m 1 dp) |
| FEP profit – RtB route [key row] | € m | 6.5 | 13.1 | `Out_RtB_Profit_Buildup_1`, `Dashboard!H103` (€ m 1 dp) |

Card – Cash to the exit:
- FEP funds the deposits, connection fees and land; the buyer reimburses €43.7m for both bids at the exit.  
  <sub>Cells: `Dashboard!H88` (€ m 1 dp)</sub>
- The largest outlay is the connection fees, €2.8m direct and €37.6m indirect, paid on 08-Jan-2029.  
  <sub>Cells: `Dashboard!E78 + Dashboard!F78`, `Dashboard!E79 + Dashboard!F79` (€ m 1 dp); `Dashboard!I78` (date)</sub>
- The buyer replaces the PB guarantee at the exit.

Footnote: RtB price 150 k€/MVA × 100 MVA per bid, per the Bora scenarios meeting file (Bora!N5), confirmed 12-Oct-2026. Per-MW figures use 151 MW of installed wind per bid. The participation fee is not recovered. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `In_RtB_Premium`, `In_MVA`, `Out_Wind_MW_1` (0 dp); `Dashboard!E6` (date)</sub>

## Slide 5 – Funding and capex

- Section: Funding and capex
- Action title: Both bids need €146.8m of cash equity to COD and up to €14.0m of guarantees; all-in cost is 1,727 k€ per built MW  
  <sub>Cells: `Out_Funding_Equity_To_COD`, `Out_Funding_Peak_Guarantee` (€ m 1 dp); `Out_Capex_Total_MW` (k€/MW 0 dp)</sub>
- Subtitle: €49.9m of the cash is spent by the RtB exit on 31-Dec-2029; on the build route, construction equity of €77.6m follows from NTP to COD.  
  <sub>Cells: Funding block "Cum. cash €m" on the last line dated on or before `Dash_Exit_Date` (`Dashboard!I43`); `Dash_Exit_Date` (date); "NTP → COD: construction equity" line, "Cash €m" (`Dashboard!F44`) (€ m 1 dp)</sub>

Table left – Funding need to COD, both bids (one row per line of the Dashboard block "Funding need to COD – both bids", in model order: date, milestone label, cash, guarantee, cumulative cash; € m 1 dp, 2 dp below €0.05m; empty cells shown as –). Total row [key row]: "Total cash equity to COD; peak guarantee" = `Out_Funding_Equity_To_COD`, `Out_Funding_Peak_Guarantee`.  
  <sub>Cells: `Dashboard!E35:I47` (dates, € m); `Dashboard!F48`, `Dashboard!H48`</sub>

Table right – Capex to COD, both bids, 200 MW (one row per `Capex_Items` line: item, € m 1 dp, k€/MW 0 dp). Total row [key row]: "All-in cost to COD" = `Out_Capex_Total`, `Out_Capex_Total_MW`.  
  <sub>Cells: `Capex_Items` (`Dashboard!E51:E63`), `Dashboard!F51:F63`; `Out_Wind_Built_MW_1 + Out_Wind_Built_MW_2` (0 dp)</sub>

Footnote: Both bids together, base case, build route. Cash = FEP equity; guarantees = bank guarantees outstanding (peak shown), not cash. IDC is debt-funded. k€/MW on 200 MW built. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Out_Wind_Built_MW_1 + Out_Wind_Built_MW_2` (0 dp); `Dashboard!E6` (date)</sub>

## Slide 6 – COD and HTM value

- Section: Economics – COD and HTM
- Action title: Both bids clear the 8.0% discount rate: HTM IRR 8.9% and NPV €16.3m for both bids; a sale at COD returns 33.9%  
  <sub>Cells: `In_Discount_Rate`, `Out_HTM_IRR_1`, `Out_COD_IRR_1` (% 1 dp); `Out_HTM_NPV_1 + Out_HTM_NPV_2` (€ m 1 dp)</sub>
- Subtitle: Against the zero bid, the commitments reduce FEP's NPV by €18.1m on a sale at COD and €23.3m held to maturity, per bid.  
  <sub>Cells: `Dashboard!I29`, `Dashboard!I31` (€ m 1 dp (magnitude))</sub>

Table:

| Item | Unit | Zero bid, per bid | Per bid | Δ per bid | Both bids | Cells (format) |
|---|---|---:|---:|---:|---:|---|
| FEP IRR – sale at COD | % / pp | 49.6% | 33.9% | −15.6 | – | `Dashboard!H28`, `Out_COD_IRR_1` (% 1 dp); `Dashboard!I28` (pp 1 dp) |
| FEP NPV – sale at COD | € m | 44.6 | 26.5 | (18.1) | 53.1 | `Dashboard!H29`, `Out_COD_NPV_1`, `Out_COD_NPV_1 + Out_COD_NPV_2` (€ m 1 dp); `Dashboard!I29` (Δ € m 1 dp) |
| FEP IRR – hold to maturity | % / pp | 11.9% | 8.9% | −3.0 | – | `Dashboard!H30`, `Out_HTM_IRR_1` (% 1 dp); `Dashboard!I30` (pp 1 dp) |
| FEP NPV – hold to maturity [key row] | € m | 31.4 | 8.1 | (23.3) | 16.3 | `Dashboard!H31`, `Out_HTM_NPV_1`, `Out_HTM_NPV_1 + Out_HTM_NPV_2` (€ m 1 dp); `Dashboard!I31` (Δ € m 1 dp) |

Card – Per bid:
- Sale at COD: FEP proceeds of €121.9m on a €245.0m price, a 1.66x MOIC.  
  <sub>Cells: `Out_COD_Proceeds_1`, `Out_COD_Price_1` (€ m 1 dp); `Out_COD_MOIC_1` (x 2 dp)</sub>
- Hold to maturity: 4.60x MOIC; minimum DSCR 1.26x.  
  <sub>Cells: `Out_HTM_MOIC_1`, `Out_Min_DSCR_1` (x 2 dp)</sub>

Footnote: NPV at 8.0%. COD 30-Sep-2031, 100 MW built per bid; PPA 90 €/MWh; gearing 60.0%. HTM includes the buy-out of the developer at the RtB exit date. Funding and capex to COD are on slide 5. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `In_Discount_Rate`, `In_Gearing` (% 1 dp); `In_COD`, `Dashboard!E6` (date); `Out_Wind_Built_MW_1`, `In_PPA_Price` (0 dp)</sub>

## Slide 7 – Risks and mitigants

- Section: Risks and mitigants
- Action title: The main risk is the unbuilt overbuild: €2.8m of PB lost pro rata and 67.5 MVA of connection right across both bids; a full call is unresolved  
  <sub>Cells: `Out_Commitment_Penalty_1 + Out_Commitment_Penalty_2` (€ m 1 dp); `Out_MVA_At_Risk_1 + Out_MVA_At_Risk_2` (1 dp)</sub>
- Subtitle: The base COD and HTM returns include the pro-rata PB loss; the lost MVA, any RtB price deduction and a possible group exclusion are not in the returns.

Table (per bid):

| Risk | Exposure | Mitigant | Cells (format) |
|---|---|---|---|
| Overbuild not built: PB lost pro rata to the shortfall above 10% (Kiírás XII.12; GYIK Q28/30/44) – base | €1.4m, the PB on the extra MW. | PB at the 1.0x minimum. Part of the loss may be deducted from the RtB price. | `Out_Commitment_Penalty_1` (€ m 1 dp); `Dashboard!E17` (x 1 dp) |
| Full PB call if the points-commitment rule (Kiírás XIII.17; GYIK Q72) also applies to the overbuild – unresolved | The whole PB: COD IRR 31.8%, HTM IRR 8.5%. | Downside case, not the base. PB at the minimum limits it. | `Dashboard!J125`, `Dashboard!J127` (% 1 dp) |
| 3-year group exclusion (Kiírás VII.2 n; GYIK Q71) | Possible if the PB is called; unclear whether a pro-rata loss counts as a call. A call after COD would likely hit the group then holding the SPV. | Allocate it in the RtB SPA. | |
| Connection right lost pro rata (GYIK Q28/30) | 33.8 MVA, in the pro-rata reading only (GYIK Q72: a full call keeps the right). Not reflected in generation or price. | The buyer can build the overbuild (card). | `Out_MVA_At_Risk_1` (1 dp) |
| ÖSZ-4 scored on installed wind (GYIK Q20) | Points fall to 80.50; returns unchanged. | Ranking risk only; the commitment is the same. | `Dashboard!M121` (points (2 dp if fractional)) |
| BESS valued at cost only | Its return cost (slide 3) excludes any BESS revenue or sale value. | Conservative: any BESS revenue is upside. | |

Card – Why the base does not build the overbuild:
- The grid connection stays at 100 MVA. Building the overbuild lifts the all-in cost to COD from €172.7m to €226.4m per bid. In that case COD IRR is 51.4% and HTM IRR is 10.9%. The build decision passes to the buyer at RtB.  
  <sub>Cells: `In_MVA` (0 dp); `Out_Cost_To_COD_1`, `Dashboard!I131` (€ m 1 dp); `Dashboard!I125`, `Dashboard!I127` (% 1 dp)</sub>

Footnote: Shortfall tested 6 months after COD. Pro rata = shortfall share × the wind PB (GYIK Q28/30). The full-call and overbuild-built figures are alternative cases. Source: Bora_Bid_Model_recalc.xlsx, 12-Oct-2026.  
  <sub>Cells: `Dashboard!E6` (date)</sub>
