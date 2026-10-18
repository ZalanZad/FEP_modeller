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
   project (PPA 90 €/MWh for the whole life with 1 %/yr escalation).
4. Project Rock (`reference/Project_Rock_-_Financial_Model_2.1.xlsm`, Assumptions rows 533–653) and Project Alpha:
   debt sizing on CFADS / DSCR with a leverage cap, sculpted repayment, DSRA, tax-loss rules, COPY/PASTE tax.
5. `docs/domain/tender_2026_rules.md`: scoring bands, securities, PB validity.

## 3. Simplifications vs Bora (user: "even simpler")
- One project, one bid. No RtB exit, no JV / developer (FEP 100 %), so no claims, no waterfall, no DSA.
- No zero-bid case and no commitment-impact macro.
- **BESS (user, 13-Oct-2026): 20 % of MVA (10 MW, 4 h, 40 MWh), built**, as in Bora: ÖSZ-4 scored on net MVA,
  capex 240 k€/MW only (no BESS revenue or opex, not in the COD price), MWh × 200 mHUF in the security base,
  shortfall tested separately from wind. Brownfield (ÖSZ-7) is not available.
- **Overbuild (user, 13-Oct-2026): bid 20 % (ÖSZ-3, 60 MW committed), build 10 % (55 MW).** The 8.3 % shortfall is
  inside the 10 % tolerance, so the pro-rata reading (GYIK 28/30) calls nothing; the full-call reading (Kiírás XIII.17,
  GYIK 72) loses the whole PB 6 months after COD. Bora's sanction block and grid-bottleneck loss (5 % of the
  overbuilt MW's energy, Bora!N18) apply.
- No COD price effect: the COD price is per installed MW (old Bana engine had no adjustment at the base bid).
- Devex: one budget (500 k€ incl. contingency, Bana v3 Inputs E292) spread evenly from closing to NTP.

## 4. Workbook (`out/Bana_Model.xlsm`, layout "dashboard")
| Sheet | Content |
|---|---|
| `Dashboard` | The bid (levers, blue on yellow), returns at a glance, funding to COD, securities, sale at COD, HTM debt and tax, stored results by key-assumption set; key-assumption panel on the right. |
| `Inputs` | Scalar inputs, Live + 15 key-assumption sets (§20). Short source per row. |
| `Curves` | Tender band tables (ÖSZ-1, -2, -6) and yes/no points (as Bora). |
| `Calc_1`, `Calc_2` | Identical. Calc_1 = **European turbine**, Calc_2 = **Chinese turbine**. Column levers via `calc.lever`. |
| `Checks` | Every check → `Master_Check`. |

Timeline: quarterly, 01-Oct-2026 → 30-Sep-2060 (COD + 30 years), 136 periods. Flows at period ends; XIRR/XNPV on
period-end dates, NPV discounted to the first period end at `In_Discount_Rate` (8 %).

## 5. Dashboard levers (one column per turbine route)
| Lever | European | Chinese | Range |
|---|---|---|---|
| ÖSZ-1 municipal one-off, k€/MW (user, 13-Oct-2026) | 50 | 50 | 0 or 10–50 |
| ÖSZ-2 community contribution, €/MWh (user, 13-Oct-2026) | 3 | 3 | 0 or 2–3 |
| ÖSZ-3 overbuild committed (built: `In_Overbuild_Built` 10 %) | 20 % | 20 % | ≥ 0 % |
| ÖSZ-4 BESS power, share of MVA; duration 4 h | 20 % | 20 % | ≤ wind MW, ≥ 4 h |
| ÖSZ-6 performance security multiple | 1.0 | 1.0 | ≥ 1.0 |
| ÖSZ-5 partial award accepted | 0 | 0 | 0 / 1 |
| ÖSZ-8 shared connection | 0 | 0 | 0 / 1 |
| ÖSZ-9 European turbines ≥ 80 % (also sets the turbine route) | 1 | 0 | 0 / 1 |
| ÖSZ-9 European transformers ≥ 80 % | 0 | 0 | 0 / 1 |
Points: 59.5 European / 47.5 Chinese (20 + 15 + 5 + 7.5 + 12 European turbine). The ÖSZ-9 turbine lever drives the turbine price, substation,
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
  PB 2.5 % × ÖSZ-6 × base (guarantee, decision received 31-Jan-2027 (GYIK Q68) + 15 days → `In_PB_Release` 31-Mar-2033 on the hold route; the
  buyer takes it over at the sale). Base = committed wind MW × 406 mHUF/MW + BESS MWh × 200 mHUF/MWh.
- TSO security 4.5 mHUF/MVA (cash, decision + 15 days, Kiírás XIII.6), credited to the indirect fee.
- Direct connection fee 460 mHUF per connection point; indirect fee 54.830209 mHUF/MVA × 50 − TSO security;
  both cash at HCSSZ (08-Jan-2029). Bana v3 Inputs E288/E289 (MEKH Annex 1).
- Guarantee fee 6 % × outstanding guarantees × days/365.

**6.4 Build** (Bana v3 Inputs J; EU turbine price per Bora user decision):
- Turbine EU 950 (user, 13-Oct-2026) / CN 500 k€/MW; other wind 355 k€/MW (installation 80, foundations 90, BoP 130, owner's 40,
  land 15); substation 12,000 k€ EU / 10,800 k€ CN; 5 % contingency on all of these. Phased by construction share.
- Municipal one-off = ÖSZ-1 × MVA, paid at COD. Connection fees from 6.3.

**6.5 Senior debt (Rock / Alpha pattern, lean).** One term loan, fixed all-in 5 %.
- Draws = debt share k × each period's construction cost (capex + connection fees + municipal); IDC capitalised.
- Debt capacity = PV at the debt rate (to COD) of CFADS ÷ target DSCR over the tenor (18 years from COD,
  Rock), CFADS = EBITDA − guarantee fees − tax after COD (tax from the PASTE row).
- k = MIN(gearing cap, capacity ÷ Σ cost × IDC growth to COD). Gearing cap 55 % on both routes (user, 14-Oct-2026). Target DSCR 1.20x (Rock, contracted).
- Repayment sculpted: debt service = CFADS ÷ DSCR × MIN(1, debt at COD ÷ capacity); principal = service −
  interest. The loan amortises to zero at the tenor end by construction (check).
- Fees: upfront 1.75 % of the facility (debt at COD incl. capitalised interest) at first draw, commitment 0.9 % p.a. on undrawn (Bana v3 E219/E222),
  equity-funded. DSRA = 6 months of debt service, equity-funded at COD, released at the tenor end.
- COD sale: the buyer takes the SPV with its debt; FEP's price is EV − debt outstanding − sale costs.

**6.6 Operations.** Generation = 55 MW built × 2,600 h (user, 13-Oct-2026) × (1 − bottleneck loss) × (1 − 2 % curtailment) × degradation 0.3 %/yr.
Revenue = PPA 90 €/MWh × 1.01^(year − COD year) on 100 % of output for the whole life (user). Opex as Bora:
variable 5 €/MWh, fixed 22 k€/MW + admin 150 k€, lease 3 k€/MW from NTP, contingency 5 %, escalation 2.5 %;
MEKH 0.1 %, HIPA 2 % and innovation contribution 0.3 % of revenue; community contribution (ÖSZ-2 €/MWh).

**6.7 Tax (Hungary).**
- CIT 15 % (§12: 9 % plus a possible Pillar Two top-up, proxy; set 11 = 9 %). Depreciation base: capex, connection fees (incl. TSO security), municipal one-off, capitalised IDC,
  purchase price, DD and devex; straight line over 20 years from COD (Bana v3 E213, to confirm).
- Period costs expensed: participation fee, guarantee fees, debt fees, pre-COD lease, interest after COD.
- Tax losses carried forward, used up to 50 % of the year's positive tax base (Tao. 17. §; Rock row 653).
  Losses still unused 5 years after the NTP lapse (most losses arise from NTP; conservative for 2029–2030 losses).
- Interest limitation (30 % EBITDA, 3 bn HUF safe harbour) not binding at this size (note). Robin Hood tax 31 % is a set (§12), off in the base per Bana v3 E197
  (≤ 50 MW); with 55 MW built the exemption is to confirm.
- Circularity (tax ↔ interest ↔ debt size) broken with COPY/PASTE (`Tax_1`, `Tax_2`, Solver.bas).
- Sale at COD: gain exempt under the participation exemption (`In_Sale_Gain_Taxed` 0; 1 taxes it at CIT).

**6.8 Returns.**
- Sale at COD: FEP flows = development + equity share of construction + fees to COD, and the net proceeds at COD.
  FEP profit, IRR, MOIC, NPV; peak equity.
- HTM: FEP equity flows (100 %), IRR, MOIC, NPV; project IRR (before financing, with the actual tax); min DSCR.
- COD sale: the buyer takes the SPV with its debt and funds the DSRA.

## 7. Key-assumption sets (blank = base)
See §12 (11 sets).

## 8. Checks
Levers in range; debt within the DSCR capacity and repaid by the tenor end; DSCR ≥ target; sources = uses to COD;
solver deltas = 0; stored = live; XIRR returned numbers. The checks result is stored for every key-assumption set.

## 9. Deck (`out/Bana_IC.pptx`, FEP template, 9 slides (§17), every figure bound)
1. Cover. 2. Executive summary (narrative from the deck, condensed, plus key numbers). 3. Risks and mitigants
(Dentons DD table, unchanged). 4. Technical overview and tender bid. 5. Funding to COD (uses and sources,
securities). 6. Returns: sale at COD vs hold, with debt and tax. 7. Sensitivities (stored sets).

## 10. Deliverables
`out/Bana_Model.xlsm` (+ `_recalc.xlsx`), `out/Bana_IC.pptx`, `out/deck_bindings.csv`, `out/reconciliation.md`
(old engine vs new, with the reasons), `build.py`, `deck.py`.

## 11. User corrections (13-Oct-2026, Bana_Model_2.xlsm)
EU turbine 950 k€/MW (set 3: base × (1 + `Sens_Step` 10 %)); net P50 2,600 h (set 4: base × (1 − `Sens_Step`)); debt-share cap 55 % EU / 50 % CN;
set 5 = PPA +10 % (99 €/MWh). Deck slide 2: narrative 16 pt plus the COD-price sentence (bound to `In_COD_Price`).

## 12. User changes (14-Oct-2026)
- **Debt:** one debt-share cap of 55 % for both turbine routes (`In_Gearing`), still limited by the DSCR capacity.
- **Tax:** CIT 15 % (Pillar Two minimum) replaces 9 %; HIPA and the innovation contribution stay. Robin Hood tax
  (31 % of positive pre-tax profit, not deductible, `In_RH_On`) is a sensitivity only.
- **Buyer at COD (share deal):** unlevered IRR (pays the EV, receives EBITDA, guarantee fees, any PB call and the SPV's
  tax after COD) and levered IRR (pays EV − debt, funds the DSRA, services the debt).
- **Key-assumption sets (engine widened to 12):** 1 Base · 2 COD price −10 % · 3 COD price +10 % · 4 Turbine +10 % ·
  5 Yield −10 % · 6 PPA +10 % · 7 Full PB call · 8 Rate +1 pp · 9 Tenor 10 years · 10 Robin Hood 31 % · 11 CIT 9 %.
  The ±10 % sets are formulas on the base (`Sens_Step`). Stored per set: FEP COD and HTM IRR/MOIC, buyer IRRs, debt,
  equity, min DSCR, checks.
- **Deck:** slide 7 "Sale price and buyer return" (buyer and FEP IRR at the three price cases); buyer IRR on slides 2 and 6.

## 13. Robin Hood tax (user, 14-Oct-2026)
Sensitivity only, never the base (`In_RH_On` = 0; set 10). Note for the IC: the energy suppliers' income tax (Thtv.,
2008/LXVII) exempts producers under 50 MW only when they sell under KÁT or METÁR, so a PPA-only plant may be liable;
31 % from 2026 (41 % to 2025). To confirm with the tax adviser.

## 14. User changes (14-Oct-2026, later)
- Buyer IRRs start at the COD column (`cod_col`): Excel's XIRR returned ~0 on the full row (leading zeros).
  New engine option `Calc.xirr(..., start=)` and lint rule `x-xirr0`.
- Set 12 "Debt 6x EBITDA": debt = 6 × first full-year EBITDA (`In_Debt_Mult`), cap raised to 70 %, still within the
  DSCR capacity. Buyer IRR added to the sensitivity table.

## 15. More sets (user, 15-Oct-2026; engine limit raised to 16)
13 "BESS 100 MWh": BESS commitment 50 % of MVA (25 MW × 4 h), built (`In_BESS_Set`). 14/15 "Not built, pro rata /
full call": bid unchanged, neither the overbuild nor the BESS built (as Bora); PB call per regime; connection right lost
pro rata is a memo (`mva_risk`), not in generation. 16 "BESS not built": wind as base, BESS not built, pro rata.

## 16. BESS capex (user, 15-Oct-2026)
100 k€ per MWh built (was 240 k€/MW), capex only: `In_BESS_Capex` × MWh × `In_BESS_Built`.

## 17. Deck: "Promise or build" slide (user, 15-Oct-2026)
Slide 8 compares the base build with sets 14-16 (PB call, MVA lost, equity, sale and hold IRR); 9 slides in total.

## 18. Promise all, deliver nothing (user, 15-Oct-2026; replaces sets 14-16 of §15)
Set 14: the bid is unchanged (points kept), but nothing is delivered (`In_Deliver` = 0): no overbuild, no BESS, no municipal
one-off or community contribution, Chinese turbines (and the Chinese COD price) on the European route. Any missed
commitment triggers the full PB call (Kiírás XIII.17, GYIK 72), so the whole PB is called. Slide 8 compares it with the base.

## 19. Debt on the bank indications (user, 15-Oct-2026; replaces the 55 % cap and the 6x EBITDA set)
Raiffeisen (base): debt €0.9m per built wind MW (±10 %), 15-year annuity, minimum DSCR 1.20x on PPA revenue
(1.40x on merchant, not used: PPA for the whole life), cash sweep up to full (`In_Sweep`, 0 in the base; set 15 = 100 %).
Debt = MIN(per-MW quantum, annuity capacity at the DSCR, 70 % LTC). UniCredit (set 12): 1.30x, 12-year annuity, same
quantum. Repayment is a level annuity (scheduled) plus the sweep; DSCR on scheduled debt service.

## 20. Two bank base cases (user, 15-Oct-2026; replaces "Raiffeisen base, UniCredit set 12" of §19)
Set 1 "Base: Raiffeisen" (live model): €0.9m per built wind MW, 15-year annuity, 1.20x. Set 2 "Base: UniCredit": 12-year
annuity, 1.30x, margin above 2.3 %. UniCredit gave no quantum, only "lower LTC than Raiffeisen", so its €/MW is the
Raiffeisen quantum scaled by the annuity capacity ratio (PV factor 12 y ÷ 1.30 over PV factor 15 y ÷ 1.20 at 5 % =
0.79) → €0.71m per MW. Same 5 % all-in rate. The other sets change one input from the Raiffeisen base (sets 3-15).
Deck: slides 2, 5 and 6 show both banks; slide 8 drops the municipality legal/reputational risk (the commitments are
made to MEKH in the bid; the sanction is the PB call).

## 21. 55 % gearing base; bank terms as two sets (user, 15-Oct-2026; replaces §19-§20)
Base: debt = 55 % of construction cost on both routes (`In_Gearing`), 15-year level annuity at 5 %, the 1.20x DSCR on PPA
revenue as a sizing floor (not binding), PPA 90 €/MWh for the whole life. `In_Debt_Per_MW` = 0 in the base (off).
Sets 15 "Raiffeisen terms" (€0.9m/MW, 15 y, 1.20x) and 16 "UniCredit terms" (€0.71m/MW = Raiffeisen scaled by the
annuity capacity ratio, 12 y, 1.30x) lift the LTC cap to 70 % so the bank amount binds. Slide 6 "Bank feedback" reports
what the banks said and the model on each basis; 10 slides, 16 sets.

## 22. Upside case (user, 18-Oct-2026)
Sets 17 "Upside: PPA 100" and 18 "Upside: PPA 105" (key-assumption limit raised to 18 in the engine): net P50 2,850 h
(user; no Bana EYA in the data room – only the old IC paper's 2,460 h – so to be confirmed), zero bid (`In_Zero_Bid` = 1:
no overbuild and no BESS committed or built, ÖSZ-1/2 at the minimum 10 k€/MW and 2 €/MWh that Kiírás fn 1 says may
become mandatory, 0 points on both), PPA 100 / 105 €/MWh for the whole life, 60 % gearing. Turbine route and PB multiple
unchanged. Slide 10 "Upside case" compares them with the base; the sale at COD is the same in both, as the price is per MW.
