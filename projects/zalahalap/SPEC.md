# Zalahaláp (ZHLP) – 63 MWp PV + 50 MW / 200 MWh BESS: investment model SPEC v1 (build-ready)

Owner: orchestrator (Fable). Builder: `model-builder` (Opus). Audit: `model-auditor`.
Status: **v1, build-ready with frozen defaults** (22-Oct-2026). v0 was the draft architecture; v1 records the user's
three answers (§0) and freezes every still-open item as an explicit, flagged default (§15) so that Phases 1–2 can be
built now. Items marked **[confirm]** are those defaults; they sit on `Inputs` in blue with the note
"DEFAULT – confirm" and are listed in §16.

## 0. User answers (22-Oct-2026, verbatim: "1 should jump from 60-50 at 90k i have sent you the aurora use that 3. use from the input sheet")
1. **JV share is a step, not a slide:** Futureal 60 % when the RtB price is at or below `In_Share_Threshold` =
   90 k€/MW, 50 % above it (`In_FEP_Share_Low` 60 %, `In_FEP_Share_High` 50 %). Boundary: confirmed in item 4
   below (60 % at exactly 90, as the original slide gives); the original slide stays as the switch `In_Share_Rule`.
2. **Aurora: only the file the user sent** (`Aurora_Q2_26_HUN_Flexible_Energy_Data_AG.xlsm`). Scenario B uses the
   Aurora-implied PV capture; no flat 65 %, no synthetic Low/High (the ±15 % and Bora-ratio proxies are dropped);
   price risk is carried by the ±% price sets (§10). Scenario A keeps the original model's own curves and its ±15 %
   switch, because it reproduces that model.
3. **Devex/early works from the input sheet:** `Stat_inp!E50` = SUM(E44:E47) = 12.87 k€/MW, not Calculation's
   24.75. See §3.1 (what E48/E49 are) and §13 item 15 (effect).

**User answers, second round (22-Oct-2026, verbatim: "2 its 60 , 1 for now go into construction but the earlyworks
needs to be thought over again"):**
4. **Share boundary (§16 old Q1):** at exactly 90 k€/MW the share is **60 %**: Futureal 60 % when the RtB price is
   ≤ `In_Share_Threshold` (90), 50 % when it is above. At the original price of 90 the step rule now gives the same
   60 % as the original slide, so §13 item 14 no longer moves the Scenario A targets at 90 (`In_Share_Rule` stays as
   the switch).
5. **Early works (§16 old Q2):** keep the v1 default for now – E48/E49 (725 k€) go to construction capex at NTP
   (`In_Milestones_Late_To_Capex` = 1). The early-works treatment as a whole will be revisited, so the lines, their
   reimbursement and the 12 % return sit in a few isolated input rows and switches (`In_Milestones_Scope`,
   `In_Milestones_Late_To_Capex`, `In_EW_Reimbursed`, `In_EW_Interest`, `In_EW_To_PSH`) with the note
   "provisional – early works to be revisited (user 22-Oct-26)".

**Fix round 1 (22-Oct-2026: Fable's financial review + the audit; user: "build now, we will adjust").** Built as
defaults; every item pending the user is a flagged placeholder on `Inputs`:
6. **Grid charging (M1; confirmed by the user, 22-Oct-2026: "batery cannot chage from the grid"):** export-only
   connection, `In_Grid_Import_MW` = 0 (confirmed). Fix round 2 replaces the crude whole-stack scaling (0.595 in 2027)
   with a **component** PV-only factor per Aurora block, from `data/aurora_extract.py`: [(DA s+14 + ID s+16
   discharge cashflow) × MIN(1, −s+13 × RTE ÷ (s+5 + s+7)) + PV transfer s+20 + aFRR-up capacity s+22] ÷ s+27 (DC 4h:
   0.779 in 2027, 0.941 on average 2027–60); aFRR energy, grid-charge costs and volume network charges (RHD, grid-drawn
   energy only) drop out. aFRR-down + FCR capacity ((s+23 + s+24) ÷ s+27, 0.167 in 2027) is added back times
   `In_AFRR_Down_Share` (0, placeholder: the share deliverable from PV state of charge). Both pieces are blue Curves
   rows; the Calc sheets combine them. Hourly limits live inside Aurora's dispatch.
7. **Two PV prices (M2):** PV exported directly at Aurora's direct price (s+34 + s+35)/s+31; PV sent to the BESS at
   the transfer price s+36/s+32 (the counterpart of the stack's −s+20). Offers and fees act on the direct price only.
   (Fix round 2: the unused blended and DA÷grid curves are removed from Curves.) Curves carry short codes ("DC 4h direct", "DC 4h transfer",
   "DC 4h" for the BESS block); the full Aurora names are on `Curves`.
8. **Unlevered tax:** project, buyer-unlevered and implied-price views use an unlevered tax (no interest or fees, own
   loss carry-forward); levered views keep the actual SPV tax. (Bana had no such convention.)
9. **COD sale DSRA:** the seller is credited the DSRA balance (cash available = price − debt + DSRA); the levered
   buyer pays for it once.
10. **Shareholder items:** the DSA loan and the FEP SPV price leave the SPV tax base (D6). FEP pays CIT on the 12 %
    return as interest income even under the participation exemption; the gain is taxed when `In_Sale_Form` = 1
    (asset deal) or `In_Sale_Gain_Taxed` = 1 (D7).
11. **COD share (D1):** `In_COD_Share_Basis` 0 = the RtB-price rule (default), 1 = fixed `In_COD_Share` **[pending
    user, Q4]**; the hold buy-out is at PSH's RtB entitlement.
12. **Offers (D2):** fixed BESS fees are not scaled above the offered duration (MET 1.8 h, E.ON 2 h); a flag shows
    when the offer differs from the built configuration. Splitting a 4 h battery into a tolled slice and a merchant
    slice is not physical with one 50 MW inverter, so the spare energy earns nothing under tolling (conservative).
13. **PSH (D3):** PSH profit and its share of proceeds are shown; its IRR is a memo (150 k€ in-kind apport).
14. **Sets (D5):** "RtB price = implied", "BESS capex 200 k€/MWh" and "No BESS repower" (BESS retired at the
    repower year, no augmentation) replace "PV price +10 %", "Opex +10 %" and "Debt rate +1 pp" (§10).
15. **Other:** one BESS availability haircut (`In_BESS_Haircut`; Aurora's 98 % is inside the blocks, the two
    cancelling availability inputs are removed); implied prices floored at 0 with a "below zero" flag; clipping 0 % for the DC-coupled cases B and C
    (`In_Clip_Loss_DC`), v3 + fix 5 (2 %) kept in A; min DSCR "n/a" without debt; the six asset-deal goal seeks are
    dropped until the deal form is answered (`In_Sale_Form` now only drives the FEP gain tax; the asset-deal buyer
    step-up is not modelled); tax losses lapse on one date (NTP + 5 years), a flagged simplification.

**Deliverable: one workbook, `out/Zalahalap_Model.xlsm` (+ `_recalc.xlsx`).** No deck, no IC email. The deck/ic-writer
steps of `CLAUDE.md` do not apply to this project; the model-builder → model-auditor loop and `MODELLING_STANDARDS.md` do.

**Build style (as Bora/Bana, non-negotiable).** Lean and formula-driven; no formula over 2 nested functions or
~100 characters; no link chains; source notes ≤ 60 characters ("ZHLP v3 Stat_inp E45", "Teaser p1", "Aurora Q2-26 …");
only the sheets in §4; `fepm audit` lint clean (incl. `g-chain`, `g-long`, `g-note`, `c-sheet`, `x-xirr0`);
`Master_Check` = 0. Fewer rows beat clever rows.

---

## 1. Decision
FEP (in a JV with Pannon Solar Holding, PSH) develops Zalahaláp. The model must show, on one framework, what FEP, the
JV partner and a buyer earn under three investment strategies and three revenue views, so that the IC can choose
between **selling at RtB**, **selling at COD** and **holding to maturity**, and can value the project at a **target
buyer IRR** (implied RtB / COD price).

Project frame (teaser, Apr-2026, `reference/zalahalap/Project_Zalahalap_Teaser_ES_300426.pdf`):
63 MWp DC / 49.9 MVA AC at the future Zalahaláp substation (2018-vintage rights, **zero grid connection fee**, 12 km 35 kV
cable to be built by the project); ~60 ha in Nyirád; 1,265 kWh/kWp → 79.7 GWh/yr; optional 50 MW / 200 MWh (4 h)
co-located BESS via a technology-change request; RtB targeted Q4-2026 (teaser timeline: RTB Dec-2026, COD Mar-2028);
remuneration full merchant (arbitrage + partial ancillary services with BESS); the teaser envisages a full share sale
at RtB. Jurisdiction: Hungary (Hungarian tax rules, as in Bana).

## 2. Benchmark: what the wind models do (and we copy)
Evidence: `projects/bana/build.py` (origin/claude/zealous-planck-fhcy42) and `projects/bora_bidding/build.py` +
`SPEC.md` v2.2 (origin/claude/dreamy-hypatia-morwsf); engine `fepm/xl/builder.py` on those branches.

| Convention | Where | Reuse for ZHLP |
|---|---|---|
| Layout `dashboard`: `Dashboard | Inputs | Curves | Calc_1..n | Checks`; no Cover/Summary | `Model(layout="dashboard", calc_sheets=[...], bids=[...])` | Yes. The Dashboard is the "Summary" the brief asks for. |
| One **quarterly** timeline, flows at period ends, XIRR/XNPV on period-end dates, NPV to the first period end | Bana: 136 quarters from 01-Oct-2026; Bora: 141 from 01-Jul-2026 | Yes (§5). The engine only supports annual or quarterly (`builder.Model.__init__`); the old ZHLP engine's monthly/annual split is not copied (MODELLING_STANDARDS §3: never a second timeline). **[confirm]** |
| Identical Calc sheets with `A1` = sheet number; sheet-specific values via `calc.lever("Dash_X")` = `INDEX(Dash_X,1,$A$1)` | Bana Calc_1 EU / Calc_2 CN turbine | Yes: one Calc sheet per **revenue case** (§6). |
| Inputs: Live column + up to 18 key-assumption sets, blank = base; `Sens_Step` constant for ±10 % sets (`{4: "=H{r}*(1+Sens_Step)"}`) | Bana `inputs()` | Yes (§10). |
| Curves: named blocks picked by label with `INDEX/MATCH` (`Calc.curve`) | Bora Aurora Low/Central/High | Yes: PV capture, baseload, BESS revenue stacks, inflation. |
| Flags: window share / event / cumulative (`Calc.window`, `.event`), operating year `ROUNDUP(YEARFRAC(COD,end),0)` | Bana rows `constr`, `ops`, `opyr` | Yes. |
| Debt: draw = k × construction cost per period; IDC capitalised; **k = MIN(gearing cap, DSCR annuity capacity ÷ cost, €/MW quantum ÷ cost)**; level annuity (PMT) from COD; optional sweep; upfront + commitment fee; DSRA = n months of service funded at COD, released at tenor end; min DSCR row | Bana `Debt sizing on CFADS`, `Senior debt` | Yes, as is. |
| Tax (Hungary): CIT on EBITDA − depreciation − interest − fees; straight-line depreciation base = capex + capitalised IDC + development; loss corkscrew with 50 % use cap and lapse; HIPA + innovation levy on revenue; Robin Hood as a set; tax circularity via `Calc.paste/copy` + Solver.bas | Bana `Tax` | Yes, plus the asset-vs-share sale treatment (§8). |
| Sale at COD: price − debt − sale costs (− tax on gain if `In_Sale_Gain_Taxed`); **buyer unlevered and levered IRR** on post-COD rows starting at the COD column (`Calc.xirr(..., start="cod_col")`) | Bana `Sale at COD` | Yes; extended to RtB and to the implied price (§9). |
| RtB exit: JV claims (cash 12 % / guarantee 6 % / at-risk 20 % with 2× floor, act/365, 31-Dec capitalisation), exit **waterfall** helper `waterfall(c, prefix, steps)` = `MIN(due, cash left)` per step + distributable, FEP / developer split, developer buy-out at the exit on the HTM route | Bora `FEP claims`, `RtB exit`, `waterfall()` | Pattern yes; the **steps and rates come from the ZHLP profit-distribution model** (§7), not from Bora. |
| Checks: levers in range, COD on a quarter end, debt within capacity and repaid, DSCR ≥ target, sources = uses, XIRR returned a number, solver deltas, stored = live; `Model_Checks` (all but stored = live) stored per set | Bana `checks()` | Yes, plus energy-balance checks (§11). |
| Dashboard: levers at the top, "Returns at a glance", funding to COD, sale at COD, HTM debt and tax, **stored results by key-assumption set** (`d.scenario_results(outs=[...])`), right-hand key-assumption panel + set selector | Bana `dashboard()` | Yes (§12). |
| Macros: `Scenarios.RunAllScenarios` (restores the set on error), `Solver.bas`, `GoalSeek.bas` (secant, named ranges `<Topic>_Input/_Achieved/_Target`, check row); `package(xlsx, xlsm, run=[...], include=[...])` | `vba/`, `fepm/xl/macros.py` | Yes. `Model.goal_seek` exists in the engine but no built project has used it yet; §9 says where it is needed. |
| Reconciliation: `reconcile.py` → `out/reconciliation.md`, bridge old value → each deliberate fix → new value, ties within 0.1 % / 0.1 pp | Bora, Bana | Yes, against the original ZHLP model (§3). |

**Engine prerequisite.** The branch this project starts from carries the *old* engine (standard layout only, 8 sets,
no `lever`, no `xirr(start=)`, fewer lint rules). The upgraded engine on `origin/claude/zealous-planck-fhcy42` is a
superset of the one on `dreamy-hypatia` (diff: `MAX_SETS = 18`, `Model_Checks`, `xirr(start=)`; the dreamy branch has
`Dashboard.section(values=)` and `Dashboard.across()` which planck lacks). Task 0 in §14 ports the planck engine
(`fepm/`, `vba/`, `tests/`) and the two dreamy-only Dashboard helpers onto this branch, with the tests green.

## 3. Sources and reconciliation targets
| # | File (`reference/zalahalap/`) | Role | Notes |
|---|---|---|---|
| 1 | `Project_Zalahalap_Teaser_ES_300426.pdf` | Project frame, dates, capacities | §1. Timeline on p6 is internally inconsistent (closing Aug-26 vs RTB Dec-26): **[confirm]** dates. |
| 2 | `Zalahalap_model_10112025_v3.xlsb` | **Original model = Scenario A**; source of numbers and of the reconciliation targets | Convert to xlsx in the scratchpad with `soffice --headless --convert-to xlsx` to read formulas. Inventory in §3.1. |
| 3 | `Zalahalap_buyer_IRR_mini_model_13-Jan-25.xlsx` | Buyer-return / valuation method | §9. |
| 4 | `ZHLP_Profit_Distribution_Model_Internal_ES_171025.xlsx` | JV waterfall (sheet `PURCHASE PRICE BASED 221025` is the latest) | §7. |
| 5 | `Aurora_Q2_26_HUN_Flexible_Energy_Data_AG.xlsm` | Aurora Q2-2026 Hungary flexible-energy dataset: prices and BESS revenue stacks for Scenario B | §6.2. |
| 6 | `Offer_EON_indicative_Futureal_PV.pdf`, `Offer_E2_Futureal_Zalahalap_PVBESS_20261020.pdf`, `Offer_MVM_offtake.pdf` | Offtake offers = Scenario C1 / C2 / C3 | §6.3; all indicative. |
| – | Bank term sheet(s) | Financing sets | **Not received** (Bana's Raiffeisen/UniCredit indications are wind-specific; use only as flagged placeholders). |

### 3.1 Input inventory from the original model (Scenario A) – `Zalahalap_model_10112025_v3` (cells of the xlsx conversion)
Sheets and flow: `Stat_inp` (deal inputs; column E = `IF(ISBLANK(INDEX(G:Q,$E$17)),G,INDEX(G:Q,$E$17))`, the Bora-style
case pick; **E17 = 4 = the "Negotiated case" column J**), `Dyn_Inp` (deposit profiles, EURIBOR, RTB/sale flags),
`Buyer_Inp`, `Control` (switches, column D) → `Aurora_Raw` (Aurora Q2-26 HUN flexible-energy databook, real 2025,
Central only, 2026–2060) → `Aurora_Curve` (annual 2025–2070) → `Calculation` (**monthly** H:Z, 30-Nov-25 → 31-May-27,
RtB route; Excel **data tables** at rows 187, 236, 242 and 254–258 per the Phase 0 extraction – LibreOffice does not
compute them and none is a target) → `RTB_Model` → `Summary`; `HTM_Model` (**quarterly** I:FL, 01-Oct-25 → 30-Sep-65)
→ `Summary`; `Sens_Engine` → `Sensitivities`. Phase 0 extracts: `data/old_engine/{old_inputs.json, old_curves.csv,
old_targets.csv, dump_old.py}` (commit 65a6fb2).

| Item | Value | Cell |
|---|---|---|
| PV capacity / grid export | 61 MWp / 50 MVA (export only, grid intake 0) | Control!D15, D16, D90 |
| BESS | 50 MW × 4 h = 200 MWh, switch on; capex 110 k€/MWh = 22,000 k€; augmentation 110 k€/MWh in 2048 (22.0 m€, not indexed) | Control!D91, D93, D89, D94, D95, D100, D101 |
| Yield | 1,250 MWh/MWp; availability 99.5 %; degradation 0.3 %/yr; **no P50/P90, no clipping/curtailment** (76 GWh vs 50 MVA) | Control!D54:D57 |
| BESS operation | no RTE/availability/cycle inputs; volumes, retention and revenue from the Aurora 4 h solar-DC block; "no-grid discharge scaling" 0.591 (uses the 2 h rows 62–63); network-charge retention 25 %; PV energy valued flag 1 | Aurora_Curve T:AC; Control!D98, D97, D99 |
| Dates | acquisition 30-Nov-25; RtB 31-Jul-26; sale 31-Oct-26; construction 01-Apr-27 → 31-Mar-28; COD 01-Apr-28; life 35 y (end 30-Apr-63) | Control!D19, D21, D24, D26, D28, D29, D30, D31 |
| Capex | PV equipment 330 + BoP 20 = 350 k€/MWp; MV cable 3,700; HV substation 9,500; land 620; BESS 22,000; total 57,170 k€; pre-RtB devex 3,182; early works 1,510 (milestones 24.75 k€/MW) | Control!D82, D83, D52, D104:D106, D115; HTM_Model!D147; Calculation!F60, F30 |
| Opex (real 2025) | O&M 14.96 k€/MW; balancing 5 €/MWh; land lease 96 k€/yr from COD; MV cable O&M 5 k€/MW; route-to-market 0 %; BESS fixed 30.8 k€/MW; indexed with the Aurora euro index | Control!D59:D62, D77, D96; HTM_Model row 20 |
| Revenue | Aurora Central baseload real 2025 × **flat 65 % capture**; GoO 1.75 €/MWh real; post-2060 inflation 2 % | Aurora_Curve!C, D4; Control!D7, D58, D13 |
| RtB deal | premium 90 k€/MW; **Futureal share price-dependent** in the Negotiated case: `IF(price<90, 60 %, IF(price<=100, 50 % + (100 − price)/100, 50 %))` = 60 % at 90, 50 % at ≥ 100 (case 3 column I = 60 % flat, case 1 column G = 50 %); connection rights 150; transaction costs 200; PSH land apport 150; preferred return 12 % p.a. monthly; BG margin 1.75 % over 1M EURIBOR; deposits 3,192.56 | Stat_inp!E24, J36 (K36 for Calculation!G34), E33, E37, E34, E69, E110, E97; Dyn_Inp row 50 |
| Early works / milestones ("LDA & Early works", Stat_inp!C43) | six lines E44:E49: Land 1st 2.46, Land 2nd 2.46, ICS execution design 1st 0.41, critical equipment (HV/ICS/substation + 132 kV transformer 1st) 7.54, **E48 "ICS Execution Design 2nd + 132 kV transformer 2nd" 10.66 (650 k€)**, **E49 "Land 3rd payment" 1.23 (30 mHUF = 75 k€)**; E48/E49 are timed at month 8 = the RtB month (Stat_inp!E57:E58). `Stat_inp!E50 = SUM(E44:E47)` = 12.87 (the range was never extended; G50:I50 sum the same four lines, while the Negotiated case's own total **J50 = SUM(J44:J49) = 24.75**). Calculation!F30 = 24.75 feeds the early works (Calculation!F60 = 1,510 k€ = 24.75 × 61), which the buyer reimburses and which sit inside HTM capex. **Verdict: E48/E49 are genuine project costs, not redundant lines.** Per the user (§0.3) the model's early-works scope = the four lines (12.87, 785 k€); E48/E49 are then **moved to construction capex at NTP** (`In_Milestones_Late_To_Capex` = 1) so that nothing is lost – they are no longer reimbursed at the RtB sale and no longer accrue the 12 % return. Switch `In_Milestones_Scope` (4 = user, 6 = original) for the reconciliation. | Stat_inp!C44:C49, E44:E50, J50; Calculation!F24:F30, F60 |
| Debt | **none** (no gearing, DSCR, DSRA, IDC, fees) | – |
| Tax | CIT 9 %; 30-year straight line excl. land; loss cap 50 %; no HIPA, no tax on the RtB gain | Control!D66, D65, D102 |
| Valuation | NPV rate 8 %; FX 400 HUF/EUR; buyer pays the RtB price (90); buyer target IRR only as a sensitivity axis 7–11 % | Control!D9, D70, D71; Stat_inp!E23; Sensitivities!D71:D75 |

Aurora inside the model: Q2-26 HUN flexible-energy databook, Central only, annual, real 2025 €. Baseload 2026/27/28 =
120.53 / 119.34 / 110.00 (Aurora_Curve!C9:C11) – identical to Bora's `Curves` Central baseload (Bora_IC_model_v4
Curves!E22:E23: 119.34, 110.00) and to the new dataset (§6.2). Selected PV capture nominal 79.83 / 80.70 / 75.89
(P9:P11). BESS 4 h degraded 2027: gross discharge 162.09 €/kW (after the 0.591 scaling), ancillary 14.70, network
−6.40, cost of PV stored −77.00, net 97.16 €/kW nominal (Aurora_Curve!V10, W10, Y10, AA10, AC10). The Low/High
switch **is wired** (Aurora_Curve!G8:G53 = `INDEX(D:F,$G$2)`, G2 = Control!D6 = 2 Central), but the Low/High curves are
**Central × (1 − 15 %) / (1 + 15 %)** (Aurora_Curve!D9 = C9×(1+$D$2) = 102.45, F9 = 138.61), i.e. scalar stresses, not
Aurora outputs (the databook is Central only).

### 3.2 Reconciliation targets (original model)
**Caveat:** the cached HTM chain is broken (Aurora_Curve!C8 = the text "no "), so Summary's HTM cells cache as
#VALUE!/#N/A/0. The targets marked [R] come from a scratch copy with C8 = 0 recalculated in LibreOffice (which
reproduces every other cached sheet exactly). **[confirm]** that the user accepts [R] as the Scenario A targets.

| Output | Value | Cell |
|---|---|---|
| RtB: Futureal XIRR incl. pre-RtB reimbursement / excl. | 224.5 % / 181.8 % | Summary!D6; Calculation!F144 |
| RtB: MOIC incl. deposits / excl. deposits | 1.427x / 1.576x (Calculation basis 1.795x / 2.529x) | Summary!D7, D25; Calculation!G179, G180 |
| RtB: Futureal net profit; capital; proceeds; developer profit | 2,941 k€; 6,882; 9,824; 2,430 | Summary!D8, D9, D23; Calculation!F226 |
| RtB: break-even price (1.0x) | 9.63 k€/MW | RTB_Model!D78 (Summary!D28 points at D75 – bug) |
| HTM: Futureal XIRR; MOIC; NPV @ 8 %; capital; payback | 5.50 % [R]; 1.959x [R]; −13.01 m€ [R]; 85.42 m€ [R]; 16.8 y [R] | Summary!H6:H10 |
| Project IRR (unlevered, post-CIT) | 6.37 % [R] | Summary!D31 (HTM_Model!D79) |
| Buyer IRR at 90 k€/MW; buyer MOIC; buyer NPV @ 8 % | 5.57 % [R]; 1.986x [R]; −12.52 m€ [R] | Summary!D33, D34, D36 |
| First full year 2029: 75.61 GWh; 72.56 €/MWh nominal; revenue 8.849 m€ | | Summary!H21:H24 |
| Lifetime: PV revenue 233.19 [R]; GoO 6.71 (cached); total 349.82 [R]; EBITDA 177.87 (50.8 %) [R]; capex 57.17 (cached); opex 171.95 [R]; BESS net 109.92 m€ (cached) | | Summary!H25:H34, H53, H59 |

Tolerance 0.1 % / 0.1 pp on each line, or a quantified fix (§13).

## 4. Workbook (`out/Zalahalap_Model.xlsm`, layout "dashboard")
| Sheet | Content |
|---|---|
| `Dashboard` | One page: levers (blue on yellow) = revenue case per Calc sheet, offtake offer selected, BESS on/off, BESS MW/MWh, strategy price inputs (RtB price, COD price), buyer target IRRs; "Returns at a glance" by strategy × revenue case; funding to COD; debt and tax; buyer view and implied prices; JV split; stored results by key-assumption set; right-hand key-assumption panel + set selector. |
| `Inputs` | Scalar inputs, Live + up to 18 sets. Short source per row. Placeholders flagged in the note. |
| `Curves` | Annual curves by name: Aurora baseload and solar capture (Low/Central/High, real base year), BESS revenue stack per MW-year by duration (Aurora), inflation index; the original model's curves (Scenario A); offtake offer price schedules (Scenario C); the revenue-case parameter table (§6.4); optional PV monthly shape. |
| `Calc_1`, `Calc_2`, `Calc_3` | Identical sheets. `A1` = 1/2/3. **Calc_1 = Scenario A (reconstruction), Calc_2 = Scenario B (Aurora merchant, base case), Calc_3 = Scenario C (selected offtake offer).** Every case-specific value comes through `calc.lever("Dash_Revenue_Case")` → the revenue-case table (§6.4). **[confirm]** this 3-sheet form vs. one Calc sheet + a case selector (then the A/B/C comparison would only exist in the stored results). |
| `Checks` | Every check → `Master_Check`. |

No other sheets. No data tables, no pasted snapshots outside the purple stored blocks.

## 5. Timeline
- **Quarterly**, start **01-Oct-2025** (the original model's HTM grid starts there, HTM_Model!I3 = Control!D32, and the
  acquisition of 30-Nov-2025 falls in the first quarter), end = COD + 35 years (Control!D30) → 152 periods to
  30-Jun-2063 **[confirm]**. The original RtB route is monthly (Calculation H:Z); the quarterly rebuild is a deliberate
  structural change, quantified in the reconciliation as Bora did (§13).
- Period-end flows; XIRR on period-end dates; NPV at `In_Discount_Rate` to the first period end.
- Key dates as inputs: signing/JV start, RtB, RtB exit (sale closing), FC/NTP, COD, BESS COD (may differ from PV COD
  **[confirm]**), end of life; `In_COD_Delay` months shifts NTP→COD for the COD-delay set.
- Construction share NTP → COD (window flag); operating year counter for degradation, PPA tenor, depreciation and
  BESS augmentation.

## 6. Revenue framework: one engine, three cases
All three cases run through the same rows. A case differs only in what the revenue-case table (§6.4) and the
selected curves feed in. Nothing is copied.

### 6.1 Scenario A – reconstruction of the original model
Takes every operating, price, capex, opex, debt and tax assumption of `Zalahalap_model_10112025_v3` (§3.1) **as its
Negotiated case (Stat_inp!E17 = 4) selects them**, including its own Aurora curves (Central, and its ±15 % Low/High
switch as it was, on `Curves` labelled "ZHLP v3 Aurora_Curve") and, in the reconciliation set only, the original's
sliding share rule (`In_Share_Rule` = 0: 60 % below 90, linear to 50 % at 100) and the six-line early works
(`In_Milestones_Scope` = 6). The base rule for every case is the user's step (§0.1). It is the reconciliation
benchmark, not a decision case.
Deliberate fixes to its methodology are listed in §13 and switched by `In_Fix_*` flags so that the reconciliation can
step through them (fix off = old method).

### 6.2 Scenario B – Aurora merchant (base case)
**What the dataset is** (`Aurora_Q2_26_HUN_Flexible_Energy_Data_AG.xlsm`, release 6-May-2026, Intro!E9:E15): Hungary,
EUR **real 2025**, calendar-year averages, **annual** 2026–2060, **Central scenario only** (no Low/High anywhere in the
file), battery revenue blocks pre-computed by Aurora's dispatch (method not stated; markets: day-ahead, intraday, aFRR
capacity + energy, FCR; no mFRR). It is the same databook the original model pasted into `Aurora_Raw` (baseload
2027/28 = 119.34 / 110.00 in both, and in Bora's `Curves`). It contains **no solar capture price series**
(Central prices rows 8–52: gas, CO2, demand, baseload row 20, intraday row 21 = baseload, spreads, FCR row 47,
aFRR rows 49–52). Hourly data are not in the file and may not be used in third-party tools without consent
(Disclaimer!E6), so there is no Python dispatch option on this data.

**PV price (one row, selectable by curve name; user decision §0.2: this file and nothing else):**
- `Capture_Aurora_Implied` (**Scenario B base**): year-by-year PV capture price = Aurora's own co-located solar block
  RES day-ahead cash flow ÷ RES day-ahead volume (Degraded batteries, block rows s+34 / s+29 of the selected block;
  ≈ 0.68 × baseload in 2027), real 2025 €/MWh, derived on `Curves` with the source rows cited. It is Aurora's solar
  value in this databook, not a published capture series (none exists in the file).
- `ZHLP_v3_capture`: the original model's curve (Central baseload × 65 %, with its ±15 % switch) – Scenario A only.
- No Low/High curves for B: price risk is the ±% price sets (§10). Real → nominal with the Aurora euro index
  (Common Assumptions!E14:AN14; 2027 = 1.0404); flat real after 2060. GoO as a separate €/MWh input (1.75 real, as the
  original) **[confirm]**.

**BESS revenue (Scenario B):** Aurora's **Degraded batteries** block for the configuration chosen by name on the
Dashboard (`In_BESS_Block`: "2027 solar DC 1.0cy 4h", "2027 solar DC 1.0cy 2h", "2027 solar DC 1.5cy 2h",
"2027 standalone 1.0cy 2h", or the 2030-entry blocks), using **one total only**: "Storage total cashflow"
(block row s+27, EUR/kW storage/yr = k€/MW/yr), which already nets day-ahead/intraday/aFRR-energy charging costs,
the opportunity cost of PV energy stored (s+20) and volume-related network charges (s+26), and adds aFRR/FCR capacity
(s+22:s+24). Components are shown on `Curves` for information only; never added to the total. Degradation and
repowering are inside the block (retention s+3; repower in 2048 for the 4 h 1.0cy block), so **augmentation capex
is timed to Aurora's repower year** (`In_Augment_Year` default = block repower year, cost `In_Augment_Cost`
k€/MWh **[placeholder: original 110 k€/MWh]**). Capacity-market revenue (NPV calculation row 22) is excluded.
Multiplied by built BESS MW × availability (`In_BESS_Avail`, Aurora Inputs!C18 = 98 %) × a user haircut
(`In_BESS_Haircut`, base 0).
First five years, Storage total, k€/MW (Degraded, rows 166 (4 h 1.0cy) / 123 (2 h 1.0cy) / 80 (2 h 1.5cy)):
4 h: 187.0, 163.1, 126.2, 110.6, 109.8 · 2 h 1.0cy: 140.4, 118.2, 89.0, 78.6, 75.4 · 2 h 1.5cy: 146.2, 120.9, 89.6,
74.6, 71.7.

**Co-location.** Aurora's co-located blocks are DC-coupled, fixed at 1 MW PV : 1 MW storage : 1 MW grid export, with
charging from own PV, clipping and curtailment inside the dispatch. ZHLP at 63 MWp / 50 MW BESS / 49.9 MVA is close to
that ratio, so the block is used per MW of BESS as is (the original model did the same, plus an unexplained 0.591
"no-grid discharge scaling" – dropped, §13). PV generation in the model is **our** 1,265 kWh/kWp yield, not Aurora's
1,863 MWh/MW RES (Aurora_Raw!E35); the only co-location rows kept in Excel are: PV export cap at the grid limit
(`In_Grid_Cap` 49.9 MVA × hours × a clipping-shape factor **[placeholder]**, or the yield report's clipping loss),
and the energy balance check. Whether ZHLP is AC- or DC-coupled, and whether grid charging is permitted, is §16 Q7.

**What does NOT re-optimise in Excel (stated on the Dashboard panel):** Aurora's storage revenue is a fixed annual
series per MW for a given duration, cycle cap, RTE (0.86), availability (0.98), DC co-location at 1:1:1 and entry
year. Changing MW scales it linearly; changing hours or cycles switches blocks; changing RTE, degradation, grid cap,
PV:BESS ratio or charging rules does **not** re-run Aurora's dispatch – those inputs only act through the explicit
haircut/availability multipliers and the augmentation timing. A bespoke dispatch would need hourly prices (not
supplied) and a Python pre-processor writing annual curves to `Curves` (rerun on every such change).

### 6.3 Scenario C – offtake offers (C1 E.ON, C2 E2/MET, C3 MVM), each selectable
Sources: `reference/zalahalap/Offer_EON_indicative_Futureal_PV.pdf` (E.ON Energiamegoldások, 09-Oct-2026, valid to
12-Dec-2026, indicative, Hungarian), `Offer_E2_Futureal_Zalahalap_PVBESS_20261020.pdf` (E2 Hungary / MET, 20-Oct-2026),
`Offer_MVM_offtake.pdf` (MVM ONEnergy e-mail, 22-Sep-2026, one page, scanned). All three are **indicative**; every
term below is "as offered", to be confirmed in the term sheets. The offers are short (3–4 years), so each case C is
**offer for its tenor, then the Aurora merchant tail (Scenario B rows)**; the switch-over is the `opyr` ≤ tenor flag.
A PV offer and a BESS offer are selected independently (`In_Offer_PV`, `In_Offer_BESS` on the Dashboard), because the
PV and BESS contracts are separate instruments in all three offers.

**PV offers (pay-as-produced, physical offtake, buyer in the balance group).**
| Case | Product | Price rule (EUR/MWh) | Tenor / start | Volume | Other terms |
|---|---|---|---|---|---|
| C1 E.ON PPA | Spot-indexed PaP on HUPX DAM 15-min | payable = HUPX DAM − handling fee; handling fee = MAX(7.6 % × DAM, 5) (p4 §5.1; the PDF prints "7,6%") | max 3 years; COD 2027 Q4–2028 Q1 (p1) | 100 % of output; economic curtailment option at ≤ 0 prices / DSO limits | balance-group fee 250 €/month/POD; GoO purchase option at 90 % of the HUPX GO auction price; payment 30 days |
| C2a E2/MET fixed | Fixed-price PaP | **43.96** for 01-Jan-2028 → 31-Dec-2030 (est. 235,368 MWh); **40.76** for 01-Jan-2028 → 31-Dec-2031 (313,824 MWh) (p3) | 3 or 4 years from 01-Jan-2028 | 100 % (MET takes balancing, forecasting, scheduling) | no indexation stated; GoO not mentioned |
| C2b E2/MET index | HUPX-indexed PaP | payable = DAM − fee; fee = MIN(10, MAX(5 % × DAM, 5.75)); when DAM − fee ≤ 0 the reference volume is 0 (no penalty); downregulation compensation +10 €/MWh on curtailed energy (blended ancillary/imbalance) (p4) | as C2a | 100 % | – |
| C3 MVM | PaP at a fixed indicative price for "PV+BESS co-location near Nyirád" | **45** for 2028–2032 (the e-mail's 2029–2034 series 96.88 → 82.82 is for a wind+BESS site, not ZHLP) | 2028–2032 | generated volume "taking the BESS optimisation effect into account" | scheduling and balancing borne by MVM; final price depends on the distance to the grid connection point |

Modelling of the indexed offers: PV capture revenue = generation × (Aurora **solar capture** price − handling fee),
with the handling-fee rule applied on the annual capture price (approximation: the fee's MAX/MIN kinks act on
15-min prices; stated on the Dashboard). The 5-year E.ON "economic curtailment" is captured by Aurora's curtailment
assumption, not modelled separately. Fixed-price offers: generation × price, no shape cost (PaP).

**BESS offers (separate contracts; the PV offer's merchant tail applies after tenor).**
| Case | Product | Economics | Tenor / start | Config in the offer |
|---|---|---|---|---|
| C1 E.ON option A | Floor + tiered profit share, **2-hour BESS**; floor renegotiated every 2 years | 2027 bands (€/kW/yr, fee on the band): ≤ 47 → 0 % (guaranteed), 47–52 → 100 % to E.ON, 52–80 → 10 %, 80–109 → 15 %, > 109 → 18 %; 2028 bands: 38 / 43 / 67 / 91 with the same fees (p5). Example: 120 → owner 105.87. Years without a floor: option B. | ≥ 2 years, max 3 (p1) | 50 MW / 100 MWh, RTE 87 %, 2 cycles/day, DoD 100 % (p3; one line says "10 MWh", a typo) |
| C1 E.ON option B | Full profit share | owner 90 % / E.ON 10 % of Net Optimization Profit = market revenues (DAM, IDM, balancing, aFRR/mFRR capacity + activation, imbalance) − charging energy − regulatory charges (RHD, system fees, taxes) (p4–5) | as above | as above |
| C2c E2/MET tolling | Full tolling, 2 cycles/day, "green" BESS | **49.7 k€/MW/yr (3 y)**, **46.4 k€/MW/yr (4 y)** (p5) | from 01-Jan-2028 | 50 MW / **90 MWh** |
| C2d E2/MET floor + PSA | Floor + 50/50 profit share above it | floor **45.6 k€/MW/yr (3 y)**, **41.9 k€/MW/yr (4 y)**; upside split 50 % Futureal / 50 % MET (p6; the reference for "upside" is not stated) | from 01-Jan-2028 | 50 MW / 90 MWh |
| C3 MVM | no BESS instrument | the 45 €/MWh already "takes the BESS optimisation into account" (unclear) | – | – |

Modelling: tolling and floors are revenue per MW-year × contracted MW × availability (the owner carries no energy
cost: charging is the toller's, so the Aurora stack must **not** be added on top during the tenor). Profit-share
cases: Aurora net arbitrage/ancillary stack per MW-year (Scenario B basis) run through the band table on `Curves`
(lower bound + fee per band, the Bora band-lookup pattern), owner share = stack − fees. Option B = stack × 90 %.
Grid charges on charging (RHD) are an input `In_BESS_Grid_Charge` and are deducted before the share only where the
offer says so (E.ON yes; MET tolling n/a).

`In_Offer_PV` ∈ {C1, C2a, C2b, C3} and `In_Offer_BESS` ∈ {C1A, C1B, C2c, C2d, none} feed Calc_3 via the revenue-case
table; one key-assumption set per sensible pairing (§10) stores all results. Open questions on the offers are in §16.

### 6.4 Revenue-case table (on `Curves`)
One row per case-specific parameter, one column per case (A, B, C1…Cn); Calc picks `INDEX(row, case)`. Keep it to
what truly differs by case: price curve name (PV), price curve name (BESS stack), PPA/offer share, offer price curve,
merchant tail switch, and — for A only, so that the reconstruction ties — yield, degradation, availability, opex and
capex overrides where the original model differs from the current project data. Everything else is a normal `Inputs`
row shared by all cases.

## 7. JV split at the RtB sale (profit distribution)
Source: `ZHLP_Profit_Distribution_Model_Internal_ES_171025.xlsx`, sheet **`PURCHASE PRICE BASED 221025`** (B3 =
22-Oct-2025, the latest; the two 071025 sheets are superseded). It is a **one-date settlement split, not a tiered
waterfall**: no time axis, no hurdle/pref/tiers/carry; interest is a flat 12 % of principal with no dates. All EUR.
Steps, as the sheet computes them (cached values):
1. Consideration = price per MWp × MWp: G8 = F8 × E8 = 120,000 × 63 = **7,560,000**.
2. less development (DSA) loan repaid: G12 = −G58 = −945,000 (G55 = 63 × 15,000 €/MWp).
3. less interest: G16 = −(G30 + G31) = −186,978.52, where G30 = 12 % × 945,000 = 113,400 and
   G31 = 12 % × 562,500 + 162,093,750 × 3/400 × 0.005 = 67,500 + 6,078.52 (the constant is unexplained; looks like a
   HUF guarantee-fee term at FX 400).
4. less "FEP SPV acquisition price": G19 = −549,000 (typed; basis not stated).
5. plus net cash G22 = 0 → **net purchase price G24 = 5,879,021.48**.
6. Split of the net price: FUTUREAL 55 % (F39 = 1 − 0.45) = 3,233,461.82; PSH 45 % = 2,645,559.67.
7. Loan receivables to FUTUREAL: G32 = 945,000 + 562,500 + 186,978.52 = 1,694,478.52 (the non-devex loan of 562,500 =
   "4,500,000 × 50/400" sits **on top of** the consideration in the buyer's total G34 = G8 + G19 + G29 + G22 = 7,573,500).
8. Totals: FUTUREAL G73 = lot + interest + 549,000 = **3,969,440.33 (52.51 %)**; PSH G79 = lot + DSA 945,000 =
   **3,590,559.67 (47.49 %)**; they sum to the 7,560,000 consideration.
Changes 071025 → 221025: the 549,000 line added (deducted from the price, credited to FUTUREAL only); split 60/40 →
55/45; result FUTUREAL −74,351, PSH +74,351.

**Inconsistencies to resolve before building (§16 Q9–Q11):** the RtB price here is 120 k€/MWp on 63 MWp vs 90 k€/MW
on 61 MW in the original model (Stat_inp!E24); the original's Negotiated-case share rule (Stat_inp!J36: 60 % below
90 k€/MW, sliding to 50 % at 100 and above) would give Futureal **50 %** at 120 k€/MW, whereas this sheet splits
55/45 after deductions and then adds the loans, interest and 549 k€ back to Futureal (52.5 % of the gross
consideration) – two different mechanisms, not one; the "consideration based" variant
(FUTUREAL alone bears the deductions) has no 221025 version; the 549,000 and the 562,500 lines are unlabelled;
rows 70/73 are mislabelled ("50 % PSH shares", "TOTAL PSH" in the FUTUREAL block).

**Model implementation (Bora pattern, dated):**
- Inputs: `In_JV_Price_Basis` (k€/MWp), the **step share rule** `In_FEP_Share_Low` 60 % / `In_FEP_Share_High` 50 % /
  `In_Share_Threshold` 90 k€/MW (§0.1; `In_Share_Rule` = 0 restores the original's sliding rule for the
  reconciliation set), `In_DSA_Per_MWp` 15 k€, `In_Loan_Rate` 12 %
  (act/365 from each draw to the exit, 31-Dec capitalisation as Bora – **a deliberate fix vs the flat 12 %**, switchable),
  `In_SPV_Price` 549 k€, `In_NonDevex_Loan` 562.5 k€ + its basis once explained.
- Calc rows (per Calc sheet): FEP loans drawn (devex schedule → principal corkscrew per class, Bora claims block
  reduced to two classes: DSA/devex loan and non-devex loan), accrued interest; at the exit the `waterfall()` helper:
  cash available = price + buyer-reimbursed loan principal; steps = sale costs → loan principal → interest → SPV price
  item → distributable, split by the step share rule (§0.1) / (1 − share); PSH receives its lot + the DSA.
- **COD sale:** same claims accrued to COD, same split of the net proceeds after debt and sale costs **[confirm: the
  JV document covers only the RtB sale]**.
- **HTM:** base = FEP buys PSH out at the RtB date at PSH's RtB entitlement (the original model's "PSH buy-out"
  Control!D79 = 2,134.55 k€ is the same idea); alternative `In_PSH_Stays` = 1: PSH keeps 45 % of equity flows
  **[confirm]**.
- PSH returns: PSH flows = land apport (Stat_inp!E34) and any PSH-funded devex → receipts; PSH IRR/MOIC per strategy.

## 8. Strategies, financing and tax – what is computed where
All three strategies are computed simultaneously on each Calc sheet (no strategy selector), as in Bora.

| Strategy | FEP / JV cash flows | Buyer view | Tax treatment (Hungary) |
|---|---|---|---|
| **RtB sale** (at `In_RtB_Exit`) | Contributions to the exit (devex, land, securities, JV loans), exit receipts via the waterfall | Buyer pays the RtB price, builds and operates: unlevered and levered IRR, pre- and post-tax, from the RtB column | Share sale of the SPV: participation exemption (`In_Sale_Gain_Taxed` 0) vs. taxed gain; asset/rights sale: CIT on the gain, buyer capitalises the price **[confirm deal form]** |
| **COD sale** (at `In_COD`) | Contributions to COD (equity share of construction, fees, DSRA if sold with debt) + net proceeds at COD via the waterfall | Buyer takes the SPV with its debt (levered) or pays EV (unlevered); pre- and post-tax, from the COD column | Share deal: SPV keeps its tax base and debt; asset deal: step-up depreciation for the buyer, gain taxed in the SPV **[confirm]** |
| **Hold to maturity** | FEP equity flows over life (and PSH's), project flows before financing | – | Operations taxed in the SPV: CIT, HIPA, innovation levy; Robin Hood as a set |

Financing dimension: the **unlevered** view is the project cash flow before financing (`proj_cf`, Bana) and the
**levered** view is the equity flow after debt service and DSRA; both always exist. "Alternative terms" are sets
(§10), not switches. Tax dimension: pre-tax rows = the same flows without the tax row; post-tax = with it. So the
returns block per strategy holds: FEP IRR/MOIC/NPV (post-tax, levered), project IRR (pre- and post-tax, unlevered),
buyer IRR (unlevered pre/post-tax, levered post-tax), PSH IRR/MOIC.

Debt (Bana method): k = MIN(`In_Gearing`, DSCR annuity capacity ÷ cost, `In_Debt_Per_MW` quantum ÷ cost) on PV + BESS
construction cost; CFADS = EBITDA − tax; sizing DSCR on the selected case's revenue (banks size merchant/BESS revenue
at a haircut: `In_DSCR_Merchant` vs `In_DSCR_Contracted` **[placeholder until the term sheet]**); level annuity from
COD; IDC capitalised; upfront and commitment fees; DSRA; cash sweep set. Tax circularity via `Tax` COPY/PASTE.

## 9. Valuation: buyer IRR and implied price
- **Buyer IRR at our price** (Bana rows `buy_u`, `buy_l`): post-exit flows from the exit column; unlevered pays EV,
  levered pays EV − debt and funds the DSRA.
- **Implied price at a target buyer IRR.** Where the post-exit flows do not depend on the price (share deal: SPV tax
  base unchanged; debt sized on CFADS), the implied price is **closed-form**: `XNPV(target, post-exit flows)` for the
  unlevered buyer and `XNPV(target, post-exit equity flows) + debt` for the levered buyer. No macro, live for every
  set. Where the price feeds back (asset deal: the buyer's depreciation base; RtB levered buyer with debt sized on cost
  including the price), use `Model.goal_seek` → `GoalSeek.bas` writes the price into a purple `Implied_*_Input` cell
  with the tolerance check. **[confirm]** this split; the user asked for goal seek/solver as in the wind models, but the
  fepm wind models never needed it (the closed form held).
- Targets: `In_Buyer_IRR_RtB`, `In_Buyer_IRR_COD` (unlevered, post-tax). **[placeholder 8 %** = the original model's
  NPV rate (Control!D9) and the midpoint of its 7–11 % sensitivity axis (Sensitivities!D71:D75)]; a ±1 pp set each.
- **What the existing buyer tools do** (so we keep their definitions): `Zalahalap_buyer_IRR_mini_model_13-Jan-25`
  is unlevered, post-CIT, quarterly from 01-Oct-2025 (Model!I3), 61 MWp × 1,250 MWh/MWp, 45 €/MWh flat real + GoO 1.75,
  capex 515 k€/MW, opex 14.96 k€/MW, purchase 90 k€/MW typed (Stat_Inp!H26); results Project IRR 9.384 %
  (Stat_Inp!F3), buyer IRR 7.957 % (F4). It has **no goal seek, no data table, no debt, no BESS**, and a material
  sign error (cash tax added, Model row 32/35: corrected IRRs ≈ 8.34 % / 7.00 %). The original model's buyer block
  (HTM_Model rows 81–91) gives 5.57 % at 90 k€/MW. Both confirm the closed-form approach above: the price enters only
  the first flow.
- Seller view at the implied price: the implied price feeds back into the FEP/PSH returns as a set ("Price = implied at
  target IRR") so that the Dashboard shows both "at our price" and "at the buyer's price".

## 10. Key-assumption sets (Inputs columns; blank = base; max 18)
1 Base (B on Calc_2 is the decision case; A on Calc_1 and the selected C on Calc_3 run alongside; PV offer C2a fixed +
BESS tolling C2c as the default C pairing) · 2 Reconciliation (Scenario A inputs, every §13 switch off, gearing 0) ·
3 PV price −10 % · 4 PV yield −10 % (P90 when the yield report arrives) · 5 BESS revenue −20 % · 6 Capex +10 % ·
7 Buyer IRR +1 pp · 8 RtB and COD price −10 % · 9 RtB price = implied · 10 Gearing −10 pp · 11 COD +12 months ·
12 BESS 2 h (block "DC 2h", 100 MWh) · 13 No BESS · 14 BESS capex 200 k€/MWh · 15 No BESS repower · 16 PV and BESS
prices −20 % (the merchant downside) · 17 C pairing: E.ON indexed PPA + E.ON option B · 18 C pairing: MVM 45 €/MWh +
MET floor/PSA (fix round 1: PV +10 %, opex +10 % and debt rate +1 pp dropped for sets 9, 14, 15). **[confirm the list;
18 is the engine cap — further pairings are run by changing the `In_Offer_*` levers and re-running the macro.]**

## 11. Checks (→ `Master_Check`)
Bana list (levers in range, COD on a quarter end, debt within capacity and repaid, DSCR ≥ target, sources = uses to COD,
XIRR numbers, solver deltas, stored = live) plus: PV energy balance (generation = export + BESS charging + clipped),
BESS discharge ≤ charge × RTE, export ≤ grid cap × hours, waterfall paid ≤ available and Σ paid + distributable =
available, FEP + PSH + costs = price, implied-price tolerance, revenue-case table complete for the selected case,
Scenario A ties to the stored reconciliation targets within tolerance (one row per headline, tolerance input).

## 12. Dashboard (top to bottom; value columns = Calc_1 A | Calc_2 B | Calc_3 C)
1. Header, set selector and master check; key-assumption panel on the right (prices, yield, BESS size, gearing, rate,
   DSCR, CIT, discount rate, buyer targets, FX, dates).
2. **Levers:** revenue case per sheet (text, validated), offer selected, BESS built (0/1), BESS MW, hours, RtB price,
   COD price, buyer target IRRs.
3. **Returns at a glance:** per strategy (RtB / COD / HTM) × case: FEP profit, IRR, MOIC, NPV; PSH IRR; project IRR
   pre/post-tax; buyer IRR unlevered/levered; implied price at target IRR. Each figure once.
4. Capacity and energy: MWp, MWac, BESS MW/MWh, first-year yield, export, clipped, BESS throughput, cycles.
5. Funding to COD: uses by line, debt, equity, peak equity (Bana block).
6. Debt and tax: capacity, k, DSRA, min DSCR, lifetime tax, losses lapsed.
7. JV waterfall at the RtB exit and at COD: each step due / paid, FEP / PSH receipts.
8. Stored results by key-assumption set (purple).

## 13. Deliberate fixes vs the original model (each a step in `out/reconciliation.md`; switchable where marked S)
1. Aurora_Curve!C8 text "no " breaks the HTM chain: targets taken from the patched recalc (§3.2). Not a model switch.
2. Pre-RtB cost reimbursement booked at the sale month (RTB_Model!H69, Oct-26) before the costs are incurred (H68,
   Oct-26 → Mar-27), which drives the 224.5 % vs 181.8 % XIRR gap: reimbursement at the later of sale and spend. (S)
3. Bank-guarantee phantom: Calculation!H118:H120 point one row high (Dyn_Inp!G13:G15), booking 1,213.88 at month 0;
   Control!D75 = `INDEX(Calculation!H122:Z122,D49)` = 2,023.12 inherits it (correct 809.25). (S)
4. Project IRR (HTM_Model!D79) excludes pre-RtB and early-works cash despite its label: ours includes all project cash.
5. PV export not capped at 50 MVA (76 GWh/yr DC against 50 MVA, no clipping line): clipping/grid cap row added. (S)
6. BESS "no-grid discharge scaling" 0.591 (Control!D98) taken from the 2 h rows while the 4 h block is selected, and the
   25 % network-charge retention (D97): dropped in favour of Aurora's block total; configuration chosen by name. (S)
7. Augmentation not indexed and dated 2048 by hand (Control!D101, HTM_Model!S138): indexed and tied to the Aurora
   repower year. (S)
8. Flat 12 % preferred return with no dates in the JV sheet and monthly 12 % in Calculation: act/365 dated accrual. (S)
9. No debt in the original: Scenario A reconciles unlevered (gearing set to 0 in the reconciliation set); the base
   case is levered.
10. Tax: HIPA + innovation levy added (Bana), tax on the RtB/COD gain per deal form (§8); Scenario A reconciliation
    switches them off.
11. FX 400 (Stat_inp!E23) vs Aurora 397.77: one `In_FX`.
12. Quarterly timeline for the RtB route (original monthly) and XIRR on period-end dates; NPV discounted to the first
    period end (original: 31-Dec-25, HTM_Model!I4).
13. House rules: no typed values on calc sheets (Calculation!G267 = 29.648 is the one found; Control!D75/D76,
    Stat_inp column E and Aurora_Curve!G8:G53 are formulas), no Excel data tables (Calculation rows 187, 236, 242,
    254–258: not recalculated by LibreOffice, replaced by stored macro results), no stale sensitivity caches
    (Sensitivities/Sens_Engine differ from recalc in 81 cells), one break-even price (the original shows
    9.63 / 29.65 / 107.7), `n/a` + check instead of IFERROR(…,0).
14. **JV share rule (user §0.1, §0.4):** step 60 % (price ≤ 90 k€/MW) → 50 % (above) instead of the original's slide
    (60 % at or below 90, linear to 50 % at 100). At the original's price of exactly 90 both give 60 %, so the step
    does not move the Scenario A targets; the rules differ only between 90 and 100. The reconciliation set runs the
    slide (`In_Share_Rule` = 0). (S)
15. **Early works scope (user §0.3):** four milestone lines (12.87 k€/MW, 785 k€) instead of six (24.75, 1,510 k€).
    E48 (650 k€) and E49 (75 k€) go to construction capex at NTP, so total project cost is unchanged but 725 k€ is no
    longer reimbursed at the RtB sale, no longer earns the 12 % return and no longer sits in "capital" (RtB MOIC
    and capital targets move; the reconciliation set uses `In_Milestones_Scope` = 6). (S)
16. **Grid charging (fix rounds 1–2, M1):** no grid charging (user-confirmed); the Aurora stack kept by component
    (DA/ID discharge at the PV-covered share, PV transfer, aFRR-up) instead of v3's 0.591 discharge scaling.
    (S: `In_Grid_Import_MW`, `In_AFRR_Down_Share`)
17. **PV price (fix round 1, M2):** direct export and PV-to-BESS energy priced separately (v3: one flat 65 % capture
    on all output). (S: curve names)
18. **Tax views (fix round 1):** unlevered views on an unlevered tax; FEP taxed on the 12 % return; DSA and SPV price
    out of the SPV tax base.
Known original-model bugs that are **not** reproduced (noted in the reconciliation): Summary!D28 → RTB_Model!D75
instead of D78; Sensitivities!E39:E42 use D73/D74; Calculation!G225 omits G221; Dyn_Inp!G24/H26:Y26 row offsets;
Control!D26 formula overrides the construction-start default; stale notes (Control!E8, E52, E82, E97, E115; HTM!E46…).

## 14. Implementation tasks for `model-builder` (Opus)
Phase 0 – engine (unblocked):
0.1 Port the engine from `origin/claude/zealous-planck-fhcy42` (`fepm/`, `vba/`, `tests/`, `.claude/agents`) onto this
    branch with `git checkout origin/claude/zealous-planck-fhcy42 -- fepm vba tests`; add `Dashboard.section(values=)`
    and `Dashboard.across()` from `origin/claude/dreamy-hypatia-morwsf:fepm/xl/builder.py`; `python3 -m pytest -q` green.
    Do not touch `projects/bora_bidding` or `projects/bana`.
0.2 Scratch-convert `Zalahalap_model_10112025_v3.xlsb` to xlsx (scratchpad) and write `data/old_engine/` extracts
    (inputs JSON, curves CSV, summary targets CSV) with a dump script, as Bora did.
Phase 1 – skeleton (unblocked):
1.1 `build.py`: Model with `layout="dashboard"`, 3 Calc sheets, quarterly timeline, Inputs sections (dates, capacity,
    PV ops, BESS ops, capex, opex, debt, tax, exit, JV, valuation), revenue-case table on Curves, flags, PV generation
    and co-location rows, BESS rows, EBITDA, debt (Bana), tax (Bana + sale treatment), COD sale, HTM, buyer view,
    closed-form implied prices, checks, Dashboard, `RunAllScenarios`. Placeholders flagged.
1.2 Scenario A: load the original model's inputs and curves; `reconcile.py` with the old-engine Python shadow.
Phase 2 – buildable now from the documents, final values pending the §16 answers:
2.1 JV split rows (§7, as the 221025 sheet, with the dated-interest switch). 2.2 Aurora curves on `Curves` (§6.2:
    baseload, implied capture, the degraded storage blocks by name, euro index) with a loader script
    `data/aurora_extract.py` writing `data/aurora_curves.csv` with source rows. 2.3 Offer cases C1–C3 (§6.3) incl. the
    E.ON band table. 2.4 Bank terms (placeholders until Q6). 2.5 Goal seek wiring only where §9 says the closed form
    fails (asset-deal switch).
Phase 3 – audit: `python3 -m fepm audit projects/zalahalap`, reconciliation report, commit.

## 15. Frozen defaults (build now; each is an `Inputs` row in blue, note "DEFAULT – confirm", listed in §16)
| # | Item | Default used in the build | Set / switch |
|---|---|---|---|
| 1 | Bank terms (the original has no debt) | gearing cap 55 % of PV+BESS construction cost; 15-year level annuity from COD at 5.0 % all-in; sizing DSCR 1.20x on contracted revenue, **1.35x on merchant and BESS revenue** (`In_DSCR_Merchant`); upfront 1.75 %, commitment 0.9 %; DSRA 6 months; no sweep. Values are the Bana wind indications, not ZHLP terms. | set 10 (gearing −10 pp) |
| 2 | Deal form | `In_Sale_Form` = share sale at RtB and at COD; gain exempt (`In_Sale_Gain_Taxed` 0); buyer reimburses the JV loans and the four-line early works (as RTB_Model row 69) | `In_Sale_Form` drives the FEP gain tax only (fix round 1) |
| 3 | Buyer target IRR | 8.0 % unlevered post-tax at RtB and at COD (original Control!D9; midpoint of Sensitivities!D71:D75); implied prices closed-form (§9), floored at 0 | set 7 |
| 4 | Dates | Scenario A: original (acq. 30-Nov-25, RtB 31-Jul-26, sale 31-Oct-26, NTP 01-Apr-27, COD 01-Apr-28). B and C: RtB 31-Dec-26, sale 31-Mar-27, NTP 30-Jun-27, **COD 31-Mar-2028 for PV and BESS** (teaser); MET/E.ON contracts from 01-Jan-28 / COD. Life 35 years. | set 13 (COD +12 m) |
| 5 | BESS | 50 MW × **4 h = 200 MWh** (teaser, original), DC co-located, Aurora block "2027 solar DC 1.0cy 4h", availability 98 %, RTE 0.86 (Aurora Inputs!C17:C18, information only), haircut 0 %, capex 110 k€/MWh (original Control!D94; Aurora's reference ≈ 267 k€/MWh is flagged), augmentation 110 k€/MWh indexed at the block's repower year (2048), **no grid charging (user-confirmed): component PV-only factor** (fix rounds 1–2) | sets 12 (2 h), 13 (no BESS), 14 (capex 200), 15 (no repower) |
| 6 | PV price | Aurora direct-export and PV-to-BESS transfer prices (fix round 1), Central only | sets 3, 16 |
| 7 | PV yield | 1,265 kWh/kWp (teaser; original 1,250 in Scenario A), availability 99.5 %, degradation 0.3 %/yr, clipping 2 % in case A (v3 + fix 5), 0 % in the DC-coupled cases B and C (placeholder until the yield report) | set 4 (−10 %) |
| 8 | Capex / opex | original model values (§3.1): PV 350 k€/MWp, MV cable 3,700, substation 9,500, land 620, opex lines as Control!D59:D62, D96; devex to RtB 3,182 k€; early works 12.87 k€/MW (four lines) + E48/E49 in capex at NTP | sets 7, 8 |
| 9 | Tax | CIT 9 % (original; Bana's 15 % Pillar Two proxy as a set), HIPA 2 % + innovation 0.3 % on revenue, PV and cable/substation depreciation 30 y, BESS 10 y **[placeholder]**, loss cap 50 % / 5-year lapse on one date (simplification), Robin Hood off; HIPA, innovation and the lapse flagged | – |
| 10 | JV | 221025 sheet mechanics (§7) with the step share rule (§0.1); price basis **90 k€/MW on MWp** (`In_JV_Price_Basis`, the original's figure; the JV sheet's 120 is a set); 549 k€ SPV price and 562.5 k€ non-devex loan as typed in the sheet; COD sale uses the same split; HTM = FEP buys PSH out at the RtB entitlement (`In_PSH_Stays` 0) | sets 10, 18 |
| 11 | Offers | terms as read in §6.3; E.ON fee 7.6 %; option A bands 2027/2028 then option B; BESS fixed fees not scaled above the offered duration, flag when the offer ≠ built (fix round 1); MVM 45 €/MWh on 100 % of PV output, no BESS instrument | sets 1, 17, 18 |

## 16. Still open (answers refine the defaults above; nothing blocks the build)
(Q1 answered 22-Oct-2026: 60 % at exactly 90, §0 item 4. The other numbers are kept, Inputs notes cite them.)
2. Early works: the treatment as a whole is to be revisited (user, 22-Oct-2026). For now E48 (650 k€) and E49 (75 k€)
   go to construction capex at NTP; the lines, reimbursement and 12 % return are isolated switches on `Inputs`.
3. JV: the 549 k€ "FEP SPV acquisition price" and the 562.5 k€ loan (= 4,500,000 × 50/400) – what are they; price
   basis 120 k€/MWp × 63 (JV sheet) vs 90 k€/MW × 61 (model); does the JV sheet's 55/45 survive the step rule?
4. JV at COD sale and HTM: same split at COD; buy-out at the RtB entitlement vs PSH staying in?
5. Dates: one set for B/C (RtB, closing, NTP, COD; BESS COD = PV COD?).
6. BESS: 4 h/200 MWh vs 2 h (both offers), AC/DC coupling, grid charging permission, RTE/availability/cycles,
   capex and augmentation quotes.
7. Deal form (share vs asset) at RtB and COD; what the buyer reimburses.
8. Buyer target IRRs; closed-form implied prices with goal seek only for the asset deal?
9. Bank terms for merchant PV+BESS (quantum, DSCR by revenue type, tenor, margin, DSRA, sweep).
10. Tax: CIT 9 % vs 15 %; Robin Hood; BESS depreciation life; HIPA.
11. Scenario A targets: accept the LibreOffice-recalculated HTM values and the 224.5 % RtB XIRR?
12. Offers: E.ON fee 7.6 % and GoO treatment; E.ON BESS bands after 2028 and A vs B; MET fixed-PPA indexation/GoO,
    downregulation scope, PSA 50/50 reference, 90 MWh; MVM 45 €/MWh scope and BESS; which pairings as sets.
13. Yield report (P50/P90, clipping) and cable/substation quotes.
14. ~~Grid import~~ answered 22-Oct-2026: no grid charging. Still open: the share of aFRR-down/FCR capacity the
    battery can deliver from PV state of charge (`In_AFRR_Down_Share`, 0 now; worth up to ~17 % of the 2027 stack).
15. Offers vs the built BESS: re-quote MET/E.ON for 50 MW / 200 MWh, or build the offered duration?
