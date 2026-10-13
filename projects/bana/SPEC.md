# Bana – 50 MW onshore wind: acquisition, build and exit model (SPEC v1)

Owner: orchestrator. Builder: `model-builder`. Deck: `deck-builder`. Audit: `model-auditor`.
Written 13-Oct-2026 from the user brief: "the same Excel as for Bora, same logic but even simpler; one project;
the bid stays at the top but is not the main point; no zero-bid comparison; no RtB exit, so the sale at COD and
hold to maturity (HTM) cases; check debt and tax for HTM and COD on the Rock / Alpha pattern, Hungarian rules;
PPA 90 €/MWh for the whole life; then update the deck, streamlined."

**Build style (as Bora, non-negotiable).** Lean, formula-driven, no AI-style notes, no formula over 2 nested
functions or ~100 characters, no link chains, source notes ≤ 60 characters, only the sheets in §4.
`fepm audit` lint clean and `Master_Check` = 0.

## 1. Decision
FEP can acquire the Bana wind project (GEI; 50 MW at the Ács Katódgyár 132 kV node) through a transfer of rights,
bid it into the 30-Oct-2026 MEKH capacity tender and, if awarded, close, finish development, build and either
**sell at COD (30-Sep-2030)** or **hold to maturity**. The IC is asked to approve the purchase price (€4.5m) and
give its view on buying and developing the project. The permits name Goldwind (Chinese) turbines; a European
turbine adds 12 tender points but needs a permit amendment. Both turbine routes are shown side by side.

## 2. Sources
1. `reference/Bana_IC_model_v3.xlsm` (old engine, 09-Oct-2026): Inputs block M (project), Inputs rows 11–274,
   `Bana` sheet (saved returns), `Tender` (bid). Cited as "Bana v3 Inputs E162".
2. `reference/Bana_IC_20260109_vDSz.pptx` (IC deck, 09-Oct-2026): bid (ÖSZ-1 35 k€/MW, ÖSZ-2 2.5 €/MWh, ÖSZ-6 at
   the minimum), executive-summary narrative and the Dentons DD risk table (qualitative, kept as written).
3. `projects/bora_bidding` (SPEC v2.2, build.py): structure, securities method, user decisions that apply to every
   project (EU turbine 600 k€/MW, PPA 90 €/MWh for the whole life with 1 %/yr escalation).
4. Project Rock (`reference/Project_Rock_-_Financial_Model_2.1.xlsm`, Assumptions rows 533–653) and Project Alpha:
   debt sizing on CFADS / DSCR with a leverage cap, sculpted repayment, DSRA, tax-loss rules, COPY/PASTE tax.
5. `docs/domain/tender_2026_rules.md`: scoring bands, securities, PB validity.

## 3. Simplifications vs Bora (user: "even simpler")
- One project, one bid. No RtB exit, no JV / developer (FEP 100 %), so no claims, no waterfall, no DSA.
- No zero-bid case and no commitment-impact macro.
- No overbuild and no BESS (50 MW installed on 50 MVA, Bana v3 Inputs E297/E282): ÖSZ-3, ÖSZ-4 and ÖSZ-7 are
  not offered and not modelled; no sanction block.
- No COD price effect: the COD price is per installed MW (old Bana engine had no adjustment at the base bid).
- Devex: one budget (500 k€ incl. contingency, Bana v3 Inputs E292) spread evenly from closing to NTP.

## 4. Workbook (`out/Bana_Model.xlsm`, layout "dashboard")
| Sheet | Content |
|---|---|
| `Dashboard` | The bid (levers, blue on yellow), returns at a glance, funding to COD, securities, sale at COD, HTM debt and tax, stored results by key-assumption set; key-assumption panel on the right. |
| `Inputs` | Scalar inputs, Live + 8 key-assumption sets. Short source per row. |
| `Curves` | Tender band tables (ÖSZ-1, -2, -6) and yes/no points (as Bora). |
| `Calc_1`, `Calc_2` | Identical. Calc_1 = **European turbine**, Calc_2 = **Chinese turbine**. Column levers via `calc.lever`. |
| `Checks` | Every check → `Master_Check`. |

Timeline: quarterly, 01-Oct-2026 → 30-Sep-2060 (COD + 30 years), 136 periods. Flows at period ends; XIRR/XNPV on
period-end dates, NPV discounted to the first period end at `In_Discount_Rate` (8 %).

## 5. Dashboard levers (one column per turbine route)
| Lever | European | Chinese | Range |
|---|---|---|---|
| ÖSZ-1 municipal one-off, k€/MW | 35 | 35 | 0 or 10–50 |
| ÖSZ-2 community contribution, €/MWh | 2.5 | 2.5 | 0 or 2–3 |
| ÖSZ-6 performance security multiple | 1.0 | 1.0 | ≥ 1.0 |
| ÖSZ-5 partial award accepted | 0 | 0 | 0 / 1 |
| ÖSZ-8 shared connection | 0 | 0 | 0 / 1 |
| ÖSZ-9 European turbines ≥ 80 % (also sets the turbine route) | 1 | 0 | 0 / 1 |
| ÖSZ-9 European transformers ≥ 80 % | 0 | 0 | 0 / 1 |
Points: 29 European / 17 Chinese (Bana deck p4). The ÖSZ-9 turbine lever drives the turbine price, substation,
COD price and gearing cap of its column, so the commitment and the cost cannot diverge.

## 6. Calc blocks (each Calc sheet)
Units k€, MW, MWh. HUF items in mHUF, converted once at `In_FX` 365.

**6.1 Timing.** Construction share NTP → COD (as Bora), COD event, operations window, lease window (NTP → end),
debt tenor window, tax-depreciation share, devex window (closing → NTP), operating year, escalation indices.

**6.2 Tender points** (Bora band lookup): ÖSZ-1, -2, -6 band tables; ÖSZ-5, -8, -9a, -9b yes/no. Total rounded to
2 dp. Validity: levers inside the offerable range.

**6.3 Development, securities and fees** (Bora method, no claims):
- Purchase price 4,500 k€ paid at closing (`In_Closing` 31-Dec-2026: closing after the award, deck p2);
  DD 150 k€ at the bid date. Devex 500 k€ evenly closing → NTP.
- Participation fee 3 mHUF (cash, bid date). Bid security 0.5 % × base (guarantee, bid date → PB posting).
  PB 2.5 % × ÖSZ-6 × base (guarantee, award + 15 days → `In_PB_Release` 31-Mar-2033 on the hold route; the
  buyer takes it over at the sale). Base = installed MW × 406 mHUF/MW.
- TSO security 4.5 mHUF/MVA (cash, 30-Dec-2026), credited to the indirect fee.
- Direct connection fee 460 mHUF per connection point; indirect fee 54.830209 mHUF/MVA × 50 − TSO security;
  both cash at HCSSZ (08-Jan-2029). Bana v3 Inputs E288/E289 (MEKH Annex 1).
- Guarantee fee 6 % × outstanding guarantees × days/365.

**6.4 Build** (Bana v3 Inputs J; EU turbine price per Bora user decision):
- Turbine EU 600 / CN 500 k€/MW; other wind 355 k€/MW (installation 80, foundations 90, BoP 130, owner's 40,
  land 15); substation 12,000 k€ EU / 10,800 k€ CN; 5 % contingency on all of these. Phased by construction share.
- Municipal one-off = ÖSZ-1 × MVA, paid at COD. Connection fees from 6.3.

**6.5 Senior debt (Rock / Alpha pattern, lean).** One term loan, fixed all-in 5 %.
- Draws = debt share k × each period's construction cost (capex + connection fees + municipal); IDC capitalised.
- Debt capacity = PV at the debt rate (to COD) of CFADS ÷ target DSCR over the tenor (18 years from COD,
  Rock), CFADS = EBITDA − tax (tax from the PASTE row).
- k = MIN(gearing cap, capacity ÷ Σ cost × IDC growth to COD). Gearing cap 80 % EU (Rock leverage cap) /
  70 % CN (the old engine's 10 pp haircut for Chinese turbines). Target DSCR 1.20x (Rock, contracted).
- Repayment sculpted: debt service = CFADS ÷ DSCR × MIN(1, debt at COD ÷ capacity); principal = service −
  interest. The loan amortises to zero at the tenor end by construction (check).
- Fees: upfront 1.75 % of the facility at first draw, commitment 0.9 % p.a. on undrawn (Bana v3 E219/E222),
  equity-funded. DSRA = 6 months of debt service, equity-funded at COD, released at the tenor end.
- COD sale: the buyer takes the SPV with its debt; FEP's price is EV − debt outstanding − sale costs.

**6.6 Operations.** Generation = 50 MW × 2,460 h × (1 − 2 % curtailment) × degradation 0.3 %/yr.
Revenue = PPA 90 €/MWh × 1.01^(year − COD year) on 100 % of output for the whole life (user). Opex as Bora:
variable 5 €/MWh, fixed 22 k€/MW + admin 150 k€, lease 3 k€/MW from NTP, contingency 5 %, escalation 2.5 %;
MEKH 0.1 %, HIPA 2 % and innovation contribution 0.3 % of revenue; community contribution (ÖSZ-2 €/MWh).

**6.7 Tax (Hungary).**
- CIT 9 %. Depreciation base: capex, connection fees (incl. TSO security), municipal one-off, capitalised IDC,
  purchase price, DD and devex; straight line over 20 years from COD (Bana v3 E213, to confirm).
- Period costs expensed: participation fee, guarantee fees, debt fees, pre-COD lease, interest after COD.
- Tax losses carried forward, used up to 50 % of the year's positive tax base (Tao. 17. §; Rock row 653).
  The 5-year expiry is not modelled; a check confirms the loss balance is nil 5 years after COD.
- Interest limitation (30 % EBITDA, 3 bn HUF safe harbour) not binding at this size; Robin Hood tax not applied
  (≤ 50 MW, Bana v3 E197). Both are notes, not rows.
- Circularity (tax ↔ interest ↔ debt size) broken with COPY/PASTE (`Tax_1`, `Tax_2`, Solver.bas).
- Sale at COD: gain exempt under the participation exemption (`In_Sale_Gain_Taxed` 0; 1 taxes it at CIT).

**6.8 Returns.**
- Sale at COD: FEP flows = development + equity share of construction + fees to COD, and the net proceeds at COD.
  FEP profit, IRR, MOIC, NPV; peak equity.
- HTM: FEP equity flows (100 %), IRR, MOIC, NPV; project IRR (pre-financing, post-tax); min and average DSCR.

## 7. Key-assumption sets (blank = base)
1 Base · 2 COD price −10 % · 3 Turbine price +10 % · 4 Yield −10 % · 5 PPA price −10 % · 6 Purchase price +10 % ·
7 Debt rate +1 pp · 8 Tenor 10 years (old engine; the gearing cap binds at 18 years, so this set tests the DSCR sizing). `RunAllScenarios` (with the tax solver) stores per set and route:
points, peak equity, COD IRR and MOIC, HTM IRR and MOIC, debt at COD, min DSCR.

## 8. Checks
Levers in range; debt repaid by the tenor end; sources = uses to COD; sale proceeds = price − debt − costs;
tax losses used within 5 years; solver deltas = 0; stored = live; XIRR returned numbers.

## 9. Deck (`out/Bana_IC.pptx`, FEP template, ~7 slides, every figure bound)
1. Cover. 2. Executive summary (narrative from the deck, condensed, plus key numbers). 3. Risks and mitigants
(Dentons DD table, unchanged). 4. Technical overview and tender bid. 5. Funding to COD (uses and sources,
securities). 6. Returns: sale at COD vs hold, with debt and tax. 7. Sensitivities (stored sets).

## 10. Deliverables
`out/Bana_Model.xlsm` (+ `_recalc.xlsx`), `out/Bana_IC.pptx`, `out/deck_bindings.csv`, `out/reconciliation.md`
(old engine vs new, with the reasons), `build.py`, `deck.py`.
