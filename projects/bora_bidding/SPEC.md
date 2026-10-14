# Bora – Bid Strategy Model (BORA 1 / BORA 2): SPEC v2

Owner: orchestrator. Builder: `model-builder`. Text: `ic-writer`. Deck: `deck-builder`. Audit: `model-auditor`.
v2 (12-Oct-2026) replaces v1. It adds the logic recovered from the old engine (`data/old_engine/rtb_logic.md`,
`htm_logic.md`; Python reproductions `old_rtb.py`, `old_htm.py` match the saved values to machine precision) and the user's decisions.

**Build style (user, non-negotiable).** Lean and streamlined, like Project Rock / Alpha:
- no AI-style notes or text blocks;
- no formula over 2 nested functions or ~100 characters;
- no link-to-a-link chains or copy-only rows (always point at the source);
- source notes ≤ 60 characters, in the form "Kiírás XII.2" or "Old engine Inputs E201";
- only the sheets listed in §4.

`fepm audit` lint (incl. `g-chain`, `g-long`, `g-note`, `c-sheet`) must be clean. Fewer rows beat clever rows.

## 0. v2.1 changes (user review, 12-Oct-2026) – these override the sections below
1. **RtB price = 150 k€/MVA × 100 MVA for both bids.** The RtB price effect (and the price reference) is removed.
   The RtB route stays: waterfall, FEP profit, IRR, MOIC. The COD price effect stays (community Δ vs reference, sanction exposure).
2. **Impact block (macro).** For each commitment, measure it against the **zero bid** (all levers at their zero level):
   zero bid + that commitment alone − zero bid. Show Δ COD IRR, Δ HTM IRR and Δ upfront securities per bid, and a total row
   (full bid − zero bid). Drop the Δ points, Δ RtB PV, Δ RtB proceeds and points-per-€m columns.
3. **Zero bid case.** The macro stores the zero bid's RtB, COD and HTM results (FEP RtB profit and IRR, COD IRR and FEP NPV,
   HTM IRR and NPV, points). They are shown next to BORA 1 | BORA 2 with the difference.
4. **Turbines.** EU 600 k€/MW (was 950), CN 500. The Dashboard shows the EU case only. A single Inputs switch `In_CN_Turbine`
   (0 EU / 1 CN) moves turbine price, substation, gearing, COD price and the ÖSZ-9 points. Set 2 "CN turbine" goes.
5. **PPA 90 €/MWh for the whole operating life**, escalating 1 %/yr as before. There is no merchant tail, so the Aurora curves go.
6. **Dashboard clean-up.** The "Source" column goes. The securities section gets a summary: cash posted, guarantees,
   carrying cost to the exit, and total. The cost-to-COD section gets a summary: development, construction, financing, all-in, per MW.

### v2.2 changes (user review, 12-Oct-2026) – override v2.1 where they differ
7. **One zero-bid column.** The zero bid is the same project for both bids, so it gets one column, computed on Calc_1.
8. **Returns at a glance, at the top of the Dashboard**, right under the header. Columns: Zero bid | BORA 1 | BORA 2 | Δ B1 | Δ B2.
   Rows: points; FEP RtB profit and IRR; COD IRR and NPV; HTM IRR and NPV; total commitment (item 10).
   It replaces the separate zero-bid block. Each figure appears once on the Dashboard: no repeats further down.
9. **Drop the PB-at-2.0× block**, with its input and Calc rows.
10. **Commitment overview** (one logical block per bid, replacing the separate securities and cost-to-COD summaries):
    - securities (cash posted + guarantees outstanding);
    - devex, land and signing/DD;
    - construction capex incl. connection fees and municipal;
    - financing (IDC + fees);
    - penalty if the overbuild is not built (selected regime, with the full call as a memo);
    - total.
    Line detail stays below it, without duplicate rows.
11. **Deposits as one input block.** Every security and fee input (rate or amount, basis, posting date, release, form) sits together
    in one "Deposits and fees" block on Inputs, in the order of the timeline. No new mechanics.
13. **Exit split after the waterfall (devex-file view)**, in the RtB section:
    - FEP: claims plus 70 %;
    - A.RES: DSA plus 30 %;
    - total.
    Show each in € m and in k€ per installed (committed) wind MW, the devex file's basis.
14. **The bid at the very top** of the Dashboard, right under the header: the commitments per bid with their points and impacts.
    "Returns at a glance" follows it.
15. **RtB impact per commitment.** Add Δ FEP RtB profit (vs the zero bid) to the impact block, so that the extra deposits of more BESS or
    overbuild show on the RtB route as well. In the RtB section, show FEP's result as a build-up:
    - receipts at the exit;
    - less cash posted and devex;
    - less signing/DD;
    - less guarantee fees (the 6 % carry on the bonds);
    - less tax;
    - = profit.
    The interest cost is visible there.
16. **Remove the RtB price PV row.** The price is fixed, so discounting it adds nothing.
17. **Per-MW label.** "per installed wind MW (151 / 140)", not "committed".
18. **Total commitment excludes the overbuild penalty** (a call on the PB already counted). The penalty stays visible on its own line.
    Part of it may in practice be deducted from the RtB price: say so in the text, but do not model it.
19. **Deck and email: streamlined.** Use the FEP IC paper design (`reference/Bora_IC_RtB_strategies.pptx` via the template).
    About 7 slides, only what the decision needs:
    1. cover;
    2. recommendation and key numbers (bid, points, returns vs the zero bid, total commitment);
    3. why these commitments: one table per commitment with points, Δ RtB profit, Δ COD IRR, Δ HTM IRR, Δ securities,
       plus one line of logic each (municipal: many points, no cash before COD; overbuild: points, not built, penalty only;
       BESS: pairs with the overbuild; PB above the minimum: no, it ties up guarantee for few points; others: no burden);
    4. securities and cash to the RtB exit, with the RtB build-up and the FEP / A.RES split;
    5. COD and HTM returns vs the zero bid, and the commitment overview;
    6. risks:
       - FAQ Q72 full call vs pro rata, with both penalty amounts and a note that part may be priced into the RtB sale;
       - the 3-year group exclusion;
       - the connection right lost pro rata;
       - the 100 MVA cap for two bids;
       - ÖSZ-4 on installed wind (GYIK Q20);
       - BESS capex only;
    7. approval requested and next steps (bid deadline 30-Oct-2026).
    The email is the same storyline in ≈ 250–350 words, EN and HU.
20. **100 MVA group cap: not raised.** The user removed it from the IC material (12-Oct-2026). Do not add it back as a risk
    in the deck, the emails or the reviews; the model's capacity test stays as it is.
21. **One strategy for both bids (user, 13-Oct-2026).** BORA 2 takes BORA 1's commitments (85 points). The model keeps two
    identical bid columns; only BORA 2's lever values change. The deck and the emails present one strategy: figures per bid
    (identical) and for both bids together (2 × 100 MVA), plus the zero bid for comparison. Nothing is shown twice.
22. **Funding need and capex breakdown (user, 14-Oct-2026; house standard, MODELLING_STANDARDS §9).**
    For both bids together (200 MW), no per-bid repeat.
    - Funding need to COD, one line per milestone with its date:
      - bid submission: participation fee (cash), bid guarantee;
      - PB posting: PB guarantee;
      - connection request: TSO cash;
      - devex and land by year;
      - HCSSZ: connection fees;
      - guarantee fees;
      - NTP → COD: construction equity (capex less debt), municipal at COD, IDC and fees.
      Each line is split into cash and guarantee. Then the cumulative cash equity to COD and the peak guarantee.
      This answers "how much equity now, and to COD".
    - Capex breakdown, each item in € m and k€ per built MW:
      - turbines, installation, foundations, BoP, owner's costs, land/easements, contingency;
      - substation, BESS, connection fees, municipal;
      - development (devex, land, signing/DD, securities);
      - financing (IDC, fees, guarantee fees).
      It sums to the all-in cost to COD.
    - Placement: compact Dashboard blocks after the commitment overview; one deck slide "Funding and capex".
12. **Inputs (user):** BESS capex 240 k€/MW; BORA 2 community contribution 2.01 €/MWh (scores ÖSZ-2 points, 68.25 total).

## 1. Decision
The IC is asked to approve the tender commitments for two 100 MVA capacity bids at the Szombathely OVIT node,
**BORA 1** and **BORA 2** (2 × 100 MVA, each within the 100 MVA per group per node cap). For each bid the IC needs to see:
- what the commitments do to the **RtB price PV** and FEP's RtB proceeds;
- what they do to the **HTM IRR**;
- the sale at COD.

## 2. Sources and user decisions
Sources (primary first):
1. `reference/Bora_scenarios_meeting.xlsx` (meeting file): commitments (Bora F/G 6–17), key assumptions (Bora L5:N21), and every business input of the
   old engine. Use the input tables in `data/old_engine/rtb_logic.md` and `htm_logic.md` §16; they carry the cell refs to cite.
2. `reference/BORA_Devex_20261002_BG.xlsx`: the **method** for securities, interest (act/365, 31-Dec compounding, freeze) and the waterfall.
3. `docs/domain/tender_2026_rules.md`: scoring bands, securities, sanctions, dates.
4. `reference/BORA_tulepites.xlsx`: overbuild penalty mechanics.

User decisions (12-Oct-2026):
- **ÖSZ-4 basis = net MVA**, as in the meeting file. This gives 85 / 67.5 points. The GYIK Q20 reading (installed wind, 80.5 / 63.0) is a disclosed risk and
  a key-assumption set (§6), not the base.
- Lean build as above.

Orchestrator decisions (where the old engine was inconsistent; each is in §8 and the reconciliation):
- **Devex lines.** Use the meeting file's devex lines (`Devex_FEP` ← pasted BORA_Devex 25-Sep), converted at the single FX.
  Each bid carries the full devex, as in the meeting file (open point: devex allocation if both bids win).
- **HTM is stand-alone.**
  - The old "deemed exit" (FEP profit forgone) is dropped.
  - FEP buys out the developer at the RtB exit date at its RtB-waterfall entitlement (DSA + 30 % share). That is a real cash flow.
- **Lean debt.** Gearing-sized annuity with capitalised IDC, upfront fee and commitment fee. No DSRA and no cash sweep.
- **Tax.** No loss-cap and no decommissioning.
- **PPA.** Runs for its tenor from COD.

## 3. Saved values to reconcile (meeting file)
| Item | BORA 1 | BORA 2 |
|---|---|---|
| Points, EU / CN turbine | 85 / 70 | 67.5 / 52.5 |
| Installed wind / BESS MW | 151 / 51 | 140 / 40 |
| Gross RtB premium €m | 15.0 | 15.0 |
| Price effect vs current bid €m | −11.07 | −4.92 |
| Effective RtB price €m | 3.93 | 10.08 |
| Distributable proceeds €m | −3.55 | 2.84 |
| FEP RtB profit €m / IRR | −1.22 / −3.9 % (Excel showed 2.98e-09, an XIRR failure) | 3.27 / 10.2 % |
| FEP IRR at COD sale, EU / CN | 31.7 % / 42.0 % | 31.9 % / 41.9 % |
| HTM IRR, EU / CN | 10.9 % / 14.4 % | 10.6 % / 13.8 % |
| All-in cost to COD €m, EU / CN | 273.9 / 196.3 | 255.5 / 183.5 |

The saved values assume the **overbuild is built** (151 / 140 MW of turbines). Reconcile against key-assumption set 3 "Overbuild built".

Reconciliation method (`out/reconciliation.md`, produced by `reconcile.py`):
- For each headline, run a step-by-step bridge in Python with the old-engine shadows (`old_rtb.py`, `old_htm.py`, extend their toggles as needed):
  old value → each deliberate fix in §8, one at a time → "new structure" (quarterly timeline and lean simplifications) → the new workbook value.
  Each step is quantified. The last step is the residual between the shadow with every fix applied and the recalculated workbook.
- Ties: within 0.1 % on values and 0.1 pp on IRR at each fix step. A residual above that is a finding to explain or fix.
- Then show set 3 → base (overbuild not built), per bid.

## 4. Workbook (`out/Bora_Bid_Model.xlsm`, layout "dashboard")
| Sheet | Content |
|---|---|
| `Dashboard` | One-page view (§5). Holds the deal levers (blue on yellow): the commitments per bid, the RtB exit date per bid and the key-assumption set selector. |
| `Inputs` | Scalar inputs with Live + 8 scenario columns (§6). Short source note per row. |
| `Curves` | Aurora wind-capture curves (Low/Central/High, real 2025), the tender band tables (lower bound, inclusive flag, inner points), and the devex schedule (§7). |
| `Calc_1`, `Calc_2` | Identical; `A1` = bid. Bid levers via `calc.lever("Dash_X")`. |
| `Checks` | Every check → `Master_Check`. |

Timeline:
- **Quarterly** from 01-Jul-2026: the signing fee on 31-Jul-2026 falls in the first period. It ends at 30-Sep-2061 (end of the 30-year life), 141 periods.
- Flows sit at the period end. XIRR and XNPV use the period-end dates. XNPV discounts to the first period end.

## 5. Calc blocks (each Calc sheet, top to bottom)
Units: k€, MW, MWh, dates. Securities are computed in mHUF and converted once with `In_FX` (HUF/EUR).

**5.1 Timing and flags.** Window/event flags (standard patterns) for:
- bid date, award, PB posting (award + `In_PB_Days`, 15);
- connection request (TSO security) and HCSSZ (connection fees);
- RtB exit (lever, per bid), NTP, construction (NTP → COD), COD;
- operations (COD → COD + life);
- the PB release for the hold/COD routes (`In_PB_Release` 31-Mar-2033);
- the sanction date (COD + `In_Sanction_Lag` months);
- the interest **accrual flag** (0 inside the freeze window `In_Freeze_Start` → `In_Freeze_End`; 1 otherwise). The freeze switch sets the end equal to the start.
- a Q4 flag for 31-Dec capitalisation;
- year index for curves and escalation;
- operating-year counter (for PPA tenor, degradation, depreciation life).

**5.2 Tender.** Per item: commitment x (from the lever) → inner points by band lookup on Curves (lower bound + inclusive flag, the old
`SUMPRODUCT((lb<x)+(lb=x)*incl)` → `INDEX(points, band+1)`) × eligibility → weighted points.
- Items and weights: ÖSZ-1 20, -2 15, -3 10, -4 15, -5 5, -6 5, -7 10, -8 5, -9 15. ÖSZ-9 = 80 × EU turbine + 20 × EU transformer.
- Totals: EU = ROUND(Σ, 2); CN = ROUND(Σ − ÖSZ-9, 2).
- ÖSZ-4 x = BESS MW ÷ (net MVA if `In_OSZ4_Basis` = 0, else committed wind MW) × 100.
- Validity tests (1 = OK): MVA ≤ 100 group cap; committed wind ≥ MVA; BESS MW ≤ committed wind MW; duration ≥ 4 h;
  ÖSZ-1 and ÖSZ-2 within the offerable range or at 0.

**5.3 Capacity.**
- Net MVA (`In_MVA` 100).
- Committed wind = MVA × (1 + overbuild).
- Committed BESS MW = BESS % × MVA; MWh = MW × hours.
- Built wind = committed if `In_Overbuild_Built` = 1, else MVA. Built BESS = committed × `In_BESS_Built`.
- Grid-bottleneck loss share = (built/MVA − 1)/(built/MVA) × `In_Bottleneck_Loss` (5 %). It is 0 when nothing is overbuilt.

**5.4 Securities and fees (devex-file method).** One row per line, amount in E (k€), date in F, projected on the timeline:
- **Participation fee:** 3 mHUF, one per bid, at the bid date. Cash. Not recovered.
- **Bid security:** 0.5 % × (committed wind × 406 + MWh × 200) mHUF. Guarantee. Posted at the bid date, **released at PB posting** (fix 3).
- **Performance security (PB):** 2.5 % × ÖSZ-6 multiple × the same base, **gross**. Guarantee. Posted at award + 15 days (fix 4). Released at the RtB exit
  (the buyer replaces it) on the RtB route, and at `In_PB_Release` on the hold route.
- **TSO connection-request security:** 4.5 mHUF/MVA × MVA, cash, at the connection request. Credited against the indirect fee.
- **Direct connection fee:** 508.35 mHUF lump sum per connection point (fix 2), cash, at HCSSZ.
- **Indirect connection fee:** 73.071 mHUF/MVA × MVA − TSO security, cash, at HCSSZ.
- **Devex and land:** from the Curves schedule by date (SUMIFS on the period). The contingency line = `In_Contingency` × (devex + land + participation fee),
  dated 31-Dec-2028, as in the old engine.

Each line on the timeline gives:
- cash out (cash lines);
- guarantee outstanding (window);
- **guarantee fee** = outstanding × `In_Guarantee_Fee` (6 %) × days/365 (cost to FEP to the exit or release).

**FEP claims (devex-file method)**, one compact block per claim class:
- **cash 12 %:** TSO security + direct fee + indirect net;
- **guarantee 6 %:** bid security + PB face value;
- **at-risk 20 % with a 2× floor:** devex + contingency + land, plus the participation fee at 12 % shown separately (as the old engine; tiny).

Per class:
- new principal in the period;
- date-weighted new principal (Σ amount × pay date; use SUMPRODUCT on the Curves table for devex);
- opening balance (principal + interest capitalised at prior 31-Decs);
- return accrued in the period = rate/365 × [opening × accrual days + (accrual end + 1) × new principal − date-weighted principal], where accrual days and
  accrual end come from the accrual flag. This reproduces act/365 stubs counting both ends.
- capitalised at the Q4 flag (`SUMIF` of the year's accruals);
- closing balance.

The claim to a date is the accrued return to that date. At-risk claim = MAX(accrued 20 %, (`In_Min_Multiple` − 1) × principal): a portfolio floor, equal to the
per-line floor while it binds on every line (it does at the base dates; state the difference at COD in the reconciliation).

**5.5 Overbuild and BESS sanction** (tested at the sanction date):
- Shortfall share s = MAX(0, 1 − built/committed), wind and BESS separately.
- **Pro rata** (GYIK 28/30, reading A): IF s > `In_Tolerance_Cap` (10 %), s × the PB component of that technology.
- **Full call** (Kiírás XIII.17 / GYIK 72): the whole PB if either s > 0.
- Show both amounts. `In_Sanction_Regime` (1 pro rata, 2 full) picks the one used.
- Memo: connection capacity lost pro rata (MVA) = s_wind × MVA if s_wind > 10 %. 3-year group exclusion flag = 1 if any call.

**5.6 RtB exit.**
- **Gross premium** = `In_RtB_Premium` (k€/MVA) × MVA.
- **Price effect** = pass-through × Σ items. Each item is a timeline row of the buyer's Δ after-tax flows vs the **price reference** (`In_Ref_*`: the meeting file's
  "current bid" 20 k€/MW, 2 €/MWh, 30 % overbuild, 30 % BESS, 1.5× PB), discounted to the exit date at `In_Buyer_Rate` 8 %
  (factor = 1/(1+r)^(MAX(0, end − exit)/365)):
  - municipal one-off: Δk€/MW × MVA at COD, plus the depreciation shield (20 years × CIT);
  - community: Δ€/MWh × generation × (1 − CIT);
  - BESS capex: (BESS MW − ref BESS MW) × capex/MW × construction share, plus the shield;
  - PB carry: (PB − PB at the ref multiple and ref MWh) × guarantee fee × (1 − CIT), from exit to `In_PB_Release` (fix 4);
  - overbuild:
    - if built, the bottleneck loss vs ref (generation × (loss − ref loss)/(1 − loss) × net margin × (1 − CIT));
    - if not built, the sanction at the sanction date. The reference assumes no sanction.
  - overbuild energy credit (`In_Overbuild_Credit`, base 0): the extra MWh of the built overbuild × net margin × (1 − CIT).
- **Effective price** = gross premium + price effect.
- **RtB price PV** = effective price discounted from the exit date to `In_Valuation_Date` (12-Oct-2026) at `In_Discount_Rate`.
- **Waterfall** at the exit, each step = MIN(cash available, amount due) with "cash available after" subtotals:
  - Cash available = effective price + buyer reimbursement (cash securities and connection fees posted + land).
  - Paid in this order:
    1. sale costs;
    2. recoverable principal;
    3. 12 % return (incl. the participation fee's 12 %);
    4. land principal;
    5. at-risk return;
    6. 6 % guarantee return;
    7. DSA = MIN(cap, 2.5 k€/MW/yr × YEARFRAC(signing, exit, 0)) × committed wind MW;
    8. distributable = remainder, FEP 70 % / developer 30 % (fix 6: no negative distributable, no developer floor).
- **FEP RtB cash flows:**
  - outflows: signing + DD at 31-Jul-2026; cash securities, connection fees, devex, land; guarantee fees;
  - at the exit: the waterfall receipts;
  - tax at the exit = −CIT × MAX(0, pre-tax profit − FEP's distributable share).
  - Outputs: FEP profit, IRR (`IFERROR(XIRR,"n/a")` + check), MOIC, NPV. The developer total = DSA paid + 30 %.

**5.7 Build case to COD.**
- **Capex on built MW:**
  - turbine (EU 950 / CN 500 k€/MW via `In_Turbine`);
  - other wind 355 k€/MW (installation 80, foundations 90, BoP 130, owner's 40, land 15), with 5 % contingency on these and the substation;
  - substation 12,000 k€ (CN 10,800);
  - BESS 200 k€/MW built.
  All phased by construction share.
- Connection fees (5.4) and the municipal one-off (k€/MW × MVA at COD).
- **Debt:** `In_Gearing` × (capex + connection fees + municipal), drawn with capex; IDC at `In_Debt_Rate` capitalised;
  upfront fee `In_Upfront_Fee` at first draw; commitment fee on undrawn × `In_Commitment_Fee`.
- **All-in cost to COD** = devex incl. land + signing/DD + capex + connection fees + municipal + cash securities + guarantee fees to COD + IDC and fees.

**5.8 Operations (HTM).** Each item is one row; opex lines merge only where the old engine treats them alike.
- **Generation** = built wind × `In_P50_Hours` × (1 − curtailment) × (1 − bottleneck loss) × degradation × operating share of the period.
- **Revenue:** PPA price (`In_PPA_Price` +1 %/yr) on 100 % during the first `In_PPA_Tenor` years from COD; then Aurora Central capture (Curves, × inflation index).
- **Opex:**
  - variable 5 €/MWh (balancing 4 + grid 1);
  - fixed 22 k€/MW/yr (O&M 12, BoP 3, insurance 2.5, AM 1.5, property tax 3);
  - land lease 3 k€/MW/yr from NTP;
  - admin 150 k€/yr;
  - contingency 5 % on these;
  - MEKH 0.1 % and HIPA 2 % of revenue;
  - community €/MWh (lever), not escalated;
  - escalation 2.5 %/yr from the COD year.
- **EBITDA.**
- **Debt service:** quarterly annuity over `In_Debt_Tenor` years from COD (PMT on the closing balance at COD).
- **Tax:**
  - depreciation (capex + connection fees + municipal + capitalised IDC + TSO security) / 20 years from COD;
  - taxable = EBITDA − depreciation − interest − fees + devex;
  - loss carry-forward corkscrew, no cap;
  - CIT 9 %.
- **Sanction cash flow** at the sanction date.
- **HTM FEP equity flows:**
  - development phase: signing/DD, devex, land, cash securities, guarantee fees to `In_PB_Release`;
  - the developer buy-out at the exit date (= DSA + 30 % from 5.6);
  - capex equity, operations after debt service and tax.
- **Outputs:** HTM IRR (XIRR), MOIC, NPV at `In_Discount_Rate`, min DSCR (memo).

**5.9 Sale at COD.**
- **Price** = `In_COD_Price` (EU 2,450 / CN 2,100 k€/MW) × built MW × (1 − bottleneck loss) + the COD price effect (community Δ vs ref and sanction exposure, PV at COD).
- Less debt outstanding and sale costs.
- **JV waterfall at COD:** FEP funding returned, then the claims to COD (5.4 classes accrued to COD), then DSA (capped 10 k€/MW × committed MW), then 70/30.
- **FEP flows:** all contributions to COD + proceeds at COD. Outputs: FEP IRR, MOIC, NPV, profit.

## 6. Key-assumption sets (Inputs scenario columns; blank = base)
1. Base: EU turbine, overbuild not built, BESS built, pro-rata sanction, freeze on, premium 150 k€/MVA.
2. CN turbine: turbine 500, substation 10,800, gearing 50 %, COD price 2,100. Points shown as CN.
3. Overbuild built (= meeting-file basis; the reconciliation set).
4. Full PB call on unmet commitments.
5. Interest freeze off.
6. Low premium 100 k€/MVA.
7. ÖSZ-4 on installed wind (GYIK Q20).
8. Spare.

## 7. Inputs (values = meeting file unless noted; note = short source)
- **Dates:** bid 30-Oct-26; award 13-Dec-26; PB days 15; TSO request 30-Dec-26; HCSSZ 08-Jan-29; NTP 30-Sep-29; COD 30-Sep-31; life 30 y; signing 26-Oct-26;
  freeze 30-Jun-29 → 01-Jan-30; PB release 31-Mar-33; sanction lag 6 m; valuation 12-Oct-26.
- **Money:** FX 365; CIT 9 %; signing 250; DD 150; sale costs 500; FEP share 70 %; DSA 2.5 k€/MW/yr, cap 10.
- **Tender:** wind benchmark 406 mHUF/MW; BESS benchmark 200 mHUF/MWh; bid security 0.5 %; PB 2.5 %; fee 3 mHUF.
- **Grid:** TSO security 4.5 mHUF/MVA; direct fee 508.345804 mHUF; indirect fee 73.071306 mHUF/MVA; group cap 100 MVA; tolerance 10 %.
- **Rates:** cash 12 %; guarantee credit 6 %; at-risk 20 %; floor 2.0×; guarantee fee 6 %; contingency on devex 10 %.
- **Exit:** premium 150 k€/MVA; pass-through 100 %; buyer rate 8 %; discount rate 8 %; COD price EU 2,450 / CN 2,100.
- **Price reference** (current bid): 20 / 2 / 30 % / 30 % / 1.5.
- **Switches:** overbuild built 0; BESS built 1; sanction regime 1; ÖSZ-4 basis 0; overbuild credit 0; bottleneck loss 5 %.
- **Capex, operations and debt:** htm_logic §16.
- **Curves:** Aurora capture Low/Central/High (curves.csv), inflation 2 % from 2025, selector "Central".
- **Devex schedule:** 31 lines + land from `devex_lines.csv` (source `engine_pasted_20260925`; exclude the "Other costs" line). Columns: label, HUF m, EUR k, date, land flag.

## 8. Deliberate fixes and changes vs the old engine (state each in `reconciliation.md` with its quantified effect)
1. Premium on awarded MVA (unchanged at 100 MVA; stated for clarity).
2. Direct connection fee = lump sum per connection point (0 effect at 100 MVA).
3. Bid security and PB are separate guarantees: the bid is released at PB posting and the PB is gross.
4. PB posted at award + 15 days and held to ≥ 31-Mar-2033 (hold route and RtB price-effect carry).
5. One FX (365) for every HUF line, devex included (the old engine embedded 370).
6. A failed XIRR shows `n/a` plus a check, never an epsilon. Waterfall steps are capped at the cash available (no negative distributable, no developer floor).
7. One participation fee per bid (the old engine charged both bids' fees to each).
8. Guarantee fee on exact days (the old engine used full months).
9. Interest per class on the quarterly timeline (act/365, 31-Dec capitalisation). Leap-year days count at /365 and the at-risk floor is a portfolio floor.
10. HTM stand-alone with the developer buy-out (no deemed exit). No DSRA, no sweep, no loss cap, no decommissioning. PPA from COD.
11. Quarterly timeline and XIRR on period-end dates (old: monthly RtB, annual HTM IRR, mid-period COD XIRR).
12. Overbuild not built in the base, with the sanction and the price-effect item that follow from it (new mechanics; set 3 removes it).

## 9. Dashboard (top to bottom; columns: label | unit | BORA 1 | BORA 2 | extra columns)
1. **Header:** title, date, key-assumption set selector (right panel), master check status.
2. **Key assumptions panel** (right; live links to Inputs): RtB premium, COD price, turbine price, PPA price/tenor, gearing, debt rate, discount rate, COD, FX,
   freeze on/off, overbuild built, BESS built, sanction regime, ÖSZ-4 basis.
3. **Tender commitments (levers)**, one row each:
   - Rows: municipal k€/MW, community €/MWh, overbuild %, BESS % of MVA, BESS h, PB multiple, partial award, brownfield, shared connection, EU turbine,
     EU transformer.
   - Columns: offerable range; points per bid; then from the macro per bid: Δ points, Δ HTM IRR, Δ RtB price PV, upfront cash/guarantee, points per €m.
4. **Points** total EU / CN and the validity tests (OK/CHECK).
5. **Securities and fees** per bid:
   - amount, form, post / release date and carry to the exit, per line;
   - total cash posted and peak guarantee;
   - sanction pro rata vs full call;
   - MVA at risk;
   - 3-year exclusion flag.
6. **RtB exit:** exit date (lever, per bid), gross premium, price effect by item, effective price, RtB price PV, the waterfall steps, FEP / developer,
   FEP profit, IRR, MOIC.
7. **Sale at COD:** all-in cost to COD (by line), sale price, FEP IRR, MOIC, profit.
8. **Hold to maturity:** HTM IRR, MOIC, NPV.
9. **Stored results:** headline outputs per key-assumption set (macro, purple). Outputs that can be "n/a" stay out of the stored block.

## 10. Macros (`vba/BoraImpact.bas`, named ranges only, `Application.Calculate` after each change)
- **`RunCommitmentImpact`:** store the base outputs first. Then, for each bid and each commitment, set the lever to its zero level, recalculate,
  store Δ points, Δ RtB price PV, Δ FEP RtB proceeds, Δ HTM IRR, Δ COD IRR and Δ upfront securities, then restore.
  - Zero levels: municipal 0, community 0, overbuild 0, BESS 0, PB 1.0, yes/no items 0. BESS hours stay.
  - Results go to `Impact_Stored`.
  - Check: the stored base equals the live values.
- **`Scenarios.RunAllScenarios`** stores the headline outputs per set.
- `package()` runs both: impact first, then scenarios.

## 11. Checks (all → `Master_Check`)
- Lever ranges valid.
- Waterfall: paid ≤ available, and Σ paid + distributable = available.
- FEP + developer + sale costs = effective price − FEP returns.
- Securities: posted − released − credited = outstanding at the end.
- Debt fully repaid by the end of the tenor.
- Sources = uses to COD.
- Stored impact base = live; stored scenario = live.
- XIRR returned a number for every route of the live case, or the `n/a` is flagged.

## 12. Deliverables
- `out/Bora_Bid_Model.xlsm` and `_recalc.xlsx`.
- `out/Bora_Bid_IC.pptx` (8–10 slides).
- `out/Bora_Bid_email_EN.md`, `out/Bora_Bid_email_HU.md`.
- `out/reconciliation.md`.
- `build.py`, `reconcile.py`.

## 13. IC storyline (deck + email)
1. Recommendation: submit BORA 1 and BORA 2 with the commitments shown (85 / 67.5 points, EU turbine).
2. Logic per commitment, with the macro impacts:
   - **Municipal:** 20 points, small Δ HTM IRR, no FEP cash before the RtB exit.
   - **Overbuild:** many points. We do not build it, so the downside is the PB sanction, which stays small because the PB is at the minimum. Show pro rata vs full call.
   - **BESS power:** pairs with the overbuild; low capex per point; many points.
   - **PB above the minimum:** rejected. Large upfront guarantee for ≤ 5 points.
   - **Other commitments:** no material financial burden.
3. Securities and cash to the RtB exit.
4. RtB economics and the waterfall.
5. COD exit and HTM.
6. Risks:
   - the FAQ Q72 full-call conflict;
   - the 3-year group exclusion;
   - the connection right lost pro rata;
   - the 100 MVA cap reading for two bids;
   - **ÖSZ-4 scored on installed wind (GYIK Q20): 80.5 / 63.0 points**.
7. Approval and next steps.
