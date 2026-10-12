# Bora – Bid Strategy Model (BORA 1 / BORA 2): SPEC v1

Owner: orchestrator. Builder: `model-builder`. Deck/email: `deck-builder` + `ic-writer`. Audit: `model-auditor`.

## 1. Decision
The IC is asked to approve the tender commitments for two 100 MVA capacity bids at the Szombathely OVIT node,
**BORA 1** and **BORA 2** (2 × 100 MVA, each within the 100 MVA per group per node cap). The IC also needs to see what
these commitments do to (a) the **present value of the RtB price** and FEP's RtB-exit proceeds, and (b) the **HTM IRR**,
with the sale at COD shown as well.

## 2. Sources (primary first)
1. `reference/Bora_scenarios_meeting.xlsx` (new version):
   - `Bora` sheet: columns F (BORA 1) and G (BORA 2) hold the commitments (rows 6–17), points (26–36),
     key assumptions (L5:N21), the RtB proceeds waterfall (48–61), and cost to COD (73–89).
   - `Devex_FEP` and `BORA_Devex`: the devex lines.
   - `Calc`, `Tender`, `Deposits`, `MOIC`, `HTM`: the engines.
2. `reference/BORA_Devex_20261002_BG.xlsx` for the method behind fees, deposits, interest and the waterfall
   (see `docs/domain/bora_devex_and_overbuild.md`).
3. `docs/domain/tender_2026_rules.md`: scoring bands, securities, sanctions.
4. `reference/BORA_tulepites.xlsx`: the overbuild penalty.

The old engine is the source for **numbers and logic**, not for **structure**. Rebuild it lean.

## 3. Saved values to reconcile (Bora sheet, new file)
| Item | BORA 1 | BORA 2 |
|---|---|---|
| Points, EU turbine | 85 | 67.5 |
| Installed wind / BESS MW | 151 / 51 | 140 / 40 |
| Gross RtB premium €m | 15.0 | 15.0 |
| Effective RtB price €m | 3.93 | 10.08 |
| Distributable proceeds €m | −3.55 | 2.84 |
| FEP IRR at COD sale, EU | 31.7% | 31.9% |
| HTM IRR, EU / CN | 10.9% / 14.4% | 10.6% / 13.8% |
| All-in cost to COD €m | 273.9 | 255.5 |

Our model must reproduce these within 0.1% (value) or 0.1 pp (IRR), **or** explain each difference as a deliberate fix
from §8 and quantify it. The BORA 1 RtB FEP IRR in the old file is an XIRR failure (2.98e-09); our model shows `n/a` with a check note instead.

## 4. Workbook architecture (`out/Bora_Bid_Model.xlsm`)
| Sheet | Content |
|---|---|
| `Dashboard` | The one-page view (see §5). Holds the **deal levers** as input cells (blue on yellow): the BORA 1 / BORA 2 commitments and the key-assumption set selector. This is the only exception to "inputs only on Inputs". |
| `Inputs` | All other scalar inputs, with Live + scenario columns (key-assumption sets, see §6). Each row cites its source. |
| `Curves` | Annual curves (capture price, baseload, CPI, EURIBOR if needed) and the tender **band tables** (ÖSZ-1…9 thresholds → inner points), as lookup tables, not nested IFs. |
| `Calc_1`, `Calc_2` | Identical sheets, one per bid. Cell `A1` = bid number. Every bid-specific input is pulled with `INDEX(Dashboard!<row>, bid)`. Quarterly timeline from 01-Oct-2026 to the end of operating life. |
| `Checks` | Every check, rolled up into `Master_Check`. |

Keep it to these sheets. No data tables, no Run_Map, no pasted snapshots.

### Calc block order (each Calc sheet)
1. Timeline + flags: bid/award, RtB, RtB exit (input date), financial close, construction, COD, operations, end of life.
2. Tender: commitment values → inner points (band lookup) → weighted points → total; validity tests
   (MVA ≤ 100, wind ≥ MVA, BESS MW ≤ wind MW, BESS ≥ 4 h).
3. Capacity: requested MVA; committed installed wind MW = MVA × (1 + overbuild %); committed BESS MW / MWh;
   **built** wind / BESS MW (switch: overbuild built yes/no, BESS built yes/no).
4. Securities and fees (devex method): participation fee, bid security, performance security (k from ÖSZ-6),
   DSO connection-request security / TSO deposit, indirect connection fee (per awarded **MVA**), direct connection fee
   (lump sum per connection point, shared per the shared-connection commitment), land, devex lines.
   Each one is **projected** in time (post / pay / release). Form (cash vs guarantee) and carrying cost (12% cash, 6% guarantee) are inputs.
5. Overbuild sanction: shortfall vs committed MW, wind and BESS tested separately, 10% tolerance → PB call share.
   A regime switch picks pro-rata or full call (FAQ Q72). The call is dated after the COD compliance test.
6. RtB exit (devex waterfall): premium = RtB k€/MVA × **awarded MVA**, then the price effect of the commitments borne by the buyer.
   The price effect is the PV of the municipal one-off, the community €/MWh stream, the extra BESS capex, PB carry and the overbuild
   penalty exposure. The overbuild energy credit is behind a switch. This gives the effective price, discounted to today as the **RtB price PV**.
   The waterfall, step by step, each capped at the cash available:
   1. buyer reimbursement of securities, connection fee and land;
   2. sale costs;
   3. recoverable loans (principal, then 12% interest);
   4. land principal;
   5. at-risk return MAX(20%, 2×);
   6. 6% guarantee interest;
   7. developer success fee;
   8. distributable proceeds, split FEP 70% / developer 30%.
   **The exit date is an input.** Interest accrues to the exit or freeze date, act/365, compounded each 31-Dec.
   FEP cash flows give FEP profit, IRR and MOIC.
7. Build case to COD: turbine (EU/CN), other wind capex, substation, BESS capex, grid fee, municipal one-off,
   securities, guarantee fees, IDC and fees. This gives the all-in cost to COD.
8. Operations (HTM): generation (P50, availability, the grid-bottleneck loss on the overbuilt MW only, degradation),
   PPA plus merchant capture revenue, opex incl. the community €/MWh, EBITDA, tax, debt (gearing, rate, tenor), equity cash flows.
   Outputs: **HTM IRR**, MOIC and NPV at the discount rate.
9. Sale at COD: price k€/MW × built MW, giving FEP IRR and MOIC.

### Impact analysis (macro)
`RunCommitmentImpact`, in `vba/BoraImpact.bas`:
- For each bid and each commitment row, set that commitment to its zero level, recalculate, and record
  Δ points, Δ RtB price PV, Δ FEP RtB proceeds, Δ HTM IRR and Δ COD-sale IRR. Then restore.
- The results go into a purple stored block on the Dashboard.
- Check: the stored base case equals the live values.
- `RunAllScenarios` stores the headline outputs for each key-assumption set.

## 5. Dashboard layout (top to bottom; two value columns, BORA 1 | BORA 2)
1. Header: title, date, key-assumption set selector, master check status.
2. **Key assumptions** panel on the right: live values of the RtB premium (k€/MVA), COD sale price, turbine price, PPA price/tenor,
   gearing, rate, discount rate, COD date, RtB exit date, FX, interest freeze, the overbuild-built and BESS-built switches,
   and the sanction regime.
3. **Tender commitments** (inputs) with range, points per item, and the **impact columns** from the macro: Δ HTM IRR,
   Δ RtB price PV, plus the upfront cash/guarantee needed and points per €m.
4. Points total (EU / CN turbine) and the validity tests.
5. **Securities and fees** (devex method), projected by date: amount, form, and carrying cost to the RtB exit.
6. **RtB exit:** the waterfall as above, with the exit date shown. FEP / developer split, FEP profit, IRR, MOIC.
7. **Sale at COD:** cost to COD, sale price, FEP IRR and MOIC.
8. **Hold to maturity:** HTM IRR, MOIC, NPV.

## 6. Key-assumption sets (Inputs scenario columns)
1. Base (EU turbine, overbuild not built, BESS built, pro-rata sanction)
2. CN turbine
3. Overbuild built
4. Full PB call on an unmet overbuild
5. Interest freeze off
6. Low premium (RtB 100 k€/MVA)
7–8. Spare

All sets use the values from the new Bora sheet unless a row says otherwise.

## 7. Deliverables
- `out/Bora_Bid_Model.xlsm` (plus a `.xlsx` recalculated copy)
- `out/Bora_Bid_IC.pptx`: IC paper, about 8–10 slides (storyline in §9)
- `out/Bora_Bid_email_EN.md` and `out/Bora_Bid_email_HU.md`: cover email to the IC / partners
- `out/reconciliation.md`: our model vs the saved values in §3
- `build.py`: rebuilds everything with `python3 -m fepm build projects/bora_bidding`

## 8. Deliberate fixes vs the old engine (state each one in the model notes)
1. The premium is on the awarded MVA (100 per bid), not on installed MW.
2. The direct connection fee is a lump sum per connection point, not per MW.
3. The bid security and the PB are separate guarantees, unless both are cash.
4. The PB is posted at award + 15 days and held to ≥ 31-Mar-2033.
5. One FX input.
6. A failed XIRR shows `n/a` plus a check, never an epsilon.

## 9. IC storyline (deck + email)
1. Recommendation: submit BORA 1 and BORA 2 with the commitments shown, giving X / Y points.
2. Why these points, the logic per commitment:
   - **Municipal contribution**: 20 points, small HTM IRR impact, and no financial commitment from FEP before the RtB exit.
   - **Overbuild**: many points. We do not plan to build it, so the downside is limited to the PB sanction share, which is capped
     because the PB is kept at the minimum. Show both sanction regimes.
   - **BESS power**: pairs with the overbuild, low capex per point, many points.
   - **Performance security above minimum**: rejected. It ties up a lot of cash or guarantee capacity up front for at most 5 points.
   - **Other commitments** (partial award, shared connection, EU equipment, community contribution where taken): accepted because they carry
     no material financial burden.
3. Securities and cash needs to the RtB exit.
4. RtB exit economics and the waterfall.
5. COD exit and HTM.
6. Risks: the FAQ conflict on full PB calls, the 3-year exclusion, the connection right lost pro rata, and the 100 MVA cap
   interpretation for two bids.
7. Approval requested and next steps.
