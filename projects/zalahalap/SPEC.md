# Zalahaláp (ZHLP) – 63 MWp PV + 50 MW / 200 MWh BESS: investment model SPEC v0 (draft architecture)

Owner: orchestrator (Fable). Builder: `model-builder` (Opus). Audit: `model-auditor`.
Status: **draft for user confirmation**. Written 22-Oct-2026 from the user brief, the repository benchmark review
(Bora and Bana fepm models, `docs/MODELLING_STANDARDS.md`) and the four project files in `reference/zalahalap/`.
Items marked **[confirm]** need a user decision; items marked **[placeholder]** are illustrative values the user must
replace (they sit on `Inputs` in blue, with the note "PLACEHOLDER – confirm").

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
| Early works / milestones | six milestone lines E44:E49 = 2.46, 2.46, 0.41, 7.54, 10.66, 1.23 k€/MW; Calculation!F30 = SUM(F24:F29) = **24.75 k€/MW** (all six); `Stat_inp!E50 = SUM(E44:E47)` = 12.87 is a partial subtotal of the first four lines and is not referenced | Stat_inp!E44:E50; Calculation!F24:F30 |
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
Negotiated case (Stat_inp!E17 = 4) selects them**, including its own Aurora curves (Central, and the ±15 % Low/High
stresses, on `Curves` labelled with their vintage) and the price-dependent Futureal share (a `Curves` band table:
60 % below 90, linear to 50 % at 100). It is the reconciliation benchmark, not a decision case.
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

**PV price (one row, selectable by curve name):**
- `Capture_Flat`: Aurora Central baseload × a flat capture rate (`In_PV_Capture_Rate`, 65 % = the original model's
  Aurora_Curve!D4) – the Scenario A method.
- `Capture_Aurora_Implied`: year-by-year capture rate implied by Aurora's own co-located solar block (RES day-ahead
  cash flow ÷ RES day-ahead volume, Degraded batteries block rows s+34 / s+29; ≈ 0.68 × baseload in 2027). Derived
  on `Curves` with the source rows cited; it is Aurora's solar value, not a published capture series.
- `Capture_Aurora_Solar`: an empty labelled curve for the official Aurora HU solar capture series **once requested**
  (§16 Q3). Base = `Capture_Aurora_Implied` **[confirm]**; Low/High sets scale baseload by the Bora `Curves` Low/High
  ÷ Central ratios (Aurora Q2-26 HU, Bora_IC_model_v4 Curves!D22:F55), the only Aurora Low/High available.
- Real → nominal with the Aurora euro index (Common Assumptions!E14:AN14; 2027 = 1.0404); flat real after 2060.
  GoO as a separate €/MWh input (original: 1.75 real).

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
- Inputs: `In_JV_Price_Basis` (k€/MWp), `In_FEP_Share` 55 % (or the original's price-dependent band table when
  `In_Share_Rule` = 1, so Scenario A reconciles), `In_DSA_Per_MWp` 15 k€, `In_Loan_Rate` 12 %
  (act/365 from each draw to the exit, 31-Dec capitalisation as Bora – **a deliberate fix vs the flat 12 %**, switchable),
  `In_SPV_Price` 549 k€, `In_NonDevex_Loan` 562.5 k€ + its basis once explained.
- Calc rows (per Calc sheet): FEP loans drawn (devex schedule → principal corkscrew per class, Bora claims block
  reduced to two classes: DSA/devex loan and non-devex loan), accrued interest; at the exit the `waterfall()` helper:
  cash available = price + buyer-reimbursed loan principal; steps = sale costs → loan principal → interest → SPV price
  item → distributable, split `In_FEP_Share` / (1 − share); PSH receives its lot + the DSA.
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
3 PV price −10 % · 4 PV price +10 % · 5 PV yield −10 % (P90 when the yield report arrives) · 6 BESS revenue −20 % ·
7 Capex +10 % · 8 Opex +10 % · 9 Buyer IRR +1 pp · 10 RtB and COD price −10 % · 11 Gearing −10 pp · 12 Debt rate +1 pp ·
13 COD +12 months · 14 BESS 2 h (block "2027 solar DC 1.0cy 2h", 100 MWh) · 15 No BESS · 16 Aurora Low (Bora ratio) ·
17 C pairing: E.ON indexed PPA + E.ON option B · 18 C pairing: MVM 45 €/MWh + MET floor/PSA. **[confirm the list;
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

## 15. Open inputs / blockers (after all documents received, 22-Oct-2026)
| # | Missing item | Depends on it | Can proceed without it? |
|---|---|---|---|
| 1 | Bank term sheet (quantum, DSCR by revenue type, tenor, margin, sweep, DSRA, fees). The original model has no debt at all. | Debt sizing, levered and buyer-levered returns | Yes with flagged placeholders (Bana's 55 %, 15 y, 5 %, 1.20x are wind indications). |
| 2 | Deal form per strategy (share vs asset sale at RtB and COD) and which costs the buyer reimburses | Tax per strategy, "cash available" | Yes: `In_Sale_Form`, default share sale; reimbursements as in the JV sheet. |
| 3 | Buyer target IRRs | Implied prices | Yes: 8 % placeholder. |
| 4 | One set of dates (teaser RTB Dec-26 / COD Mar-28; original model RtB 31-Jul-26, sale 31-Oct-26, COD 01-Apr-28; E.ON COD 27Q4–28Q1; MET start 01-Jan-28) | Timeline, offer start dates | Yes: original-model dates for A, teaser dates for B/C, flagged. |
| 5 | BESS configuration (teaser up to 50 MW/200 MWh; original 50/200 4 h; E.ON RFQ 50/100 2 h; MET 50/90) and technical data (RTE, availability, cycles, AC/DC coupling, grid charging, grid-cap sharing) | BESS block choice, capex, offers | Yes: hours as a lever (4 h base, 2 h set), Aurora's 0.86 / 0.98 as placeholders. |
| 6 | Aurora solar capture series (not in the dataset) and Low/High for Hungary | Scenario B PV price | Yes: implied capture from Aurora's solar block; Bora ratios for Low/High. |
| 7 | PV yield report (P50/P90, clipping, DC/AC) | Generation | Teaser 1,265 kWh/kWp (original 1,250); P90 placeholder. |
| 8 | Capex/opex quotes (PV, BESS, 35 kV cable, substation); devex to RtB; land | Capex/opex | Original model values (BESS 110 k€/MWh vs Aurora's ≈267 k€/MWh reference, Aurora_Raw!F79 – flag). |
| 9 | Tax confirmations (CIT 9 % vs Pillar Two 15 % proxy; Robin Hood; depreciation lives PV 30 y / BESS; HIPA) | Tax | Yes: Bana settings as placeholders. |
| 10 | JV terms: the 549 k€ SPV price, the 562.5 k€ non-devex loan, the price basis (120 vs 90 k€/MW), the share (55/45 vs 60/40), COD-sale and HTM treatment of PSH | Waterfall | Use the 221025 sheet, flagged. |
| 11 | Offer clarifications (§16 Q12–Q16) | Scenario C detail | Yes: terms as read. |

## 16. Questions for the user (numbered; answers unblock the items above)
1. Dates: which RtB, sale-closing, NTP and COD dates are the base (teaser vs original model vs offers), and is the BESS
   COD the same as the PV COD?
2. BESS base configuration: 4 h / 200 MWh (teaser, original) or 2 h / 90–100 MWh (both offers)? AC- or DC-coupled?
   Is grid charging allowed under the connection agreement (original model: export only)? RTE, availability, cycle
   cap, warranty/augmentation plan, BESS capex quote (original 110 k€/MWh).
3. Aurora: can we obtain the Q2-26 HU **solar capture price** series and Low/High cases (the flexible-energy file has
   neither)? Until then: implied capture from Aurora's solar block, or the original's flat 65 %?
4. Deal form at RtB and at COD (share vs asset sale), and what the buyer reimburses (loans, deposits, early works).
5. Buyer target IRRs (RtB and COD; unlevered post-tax?) – 8 % placeholder; and should implied prices be closed-form
   XNPV where exact, with GoalSeek.bas only for the asset-deal case?
6. Bank term sheet or indications for a merchant PV+BESS (quantum, DSCR on merchant vs contracted revenue, tenor,
   margin, DSRA, sweep); otherwise which placeholders?
7. Tax: CIT 9 % or the 15 % Pillar Two proxy used in Bana; Robin Hood applicability; depreciation lives for PV and
   BESS; HIPA 2 %.
8. Scenario A targets: accept the LibreOffice-recalculated HTM values (the original caches are #VALUE! because of
   Aurora_Curve!C8) and the 224.5 % RtB XIRR as the reconciliation benchmark?
9. JV: is the 221025 purchase-price-based sheet the agreed mechanism (55/45 after deductions, 15 k€/MWp DSA loan,
   12 % flat interest)? What are the 549,000 "FEP SPV acquisition price" and the 562,500 non-devex loan
   (= 4,500,000 × 50/400)?
10. JV share rule: the original model's Negotiated case makes Futureal's share price-dependent (60 % below 90 k€/MW,
    sliding to 50 % at ≥ 100; Stat_inp!J36), which gives 50 % at the JV sheet's 120 k€/MWp, not 55 %. Which rule is
    current – the sliding scale, the fixed 55/45, or 55/45 as the negotiated point of the scale – and is the price
    basis 120 k€/MWp × 63 or 90 k€/MW × 61?
11. JV at COD and HTM: does the same split apply to a COD sale, and does FEP buy PSH out for the hold case (at the RtB
    entitlement) or does PSH stay at 45 %?
12. E.ON PPA: confirm the handling fee is MAX(7.6 % × DAM, 5 €/MWh) (printed "7,6%"); is a 3-year tenor the maximum;
    GoO: do we keep them or sell at 90 % of the auction price?
13. E.ON BESS: option A bands beyond 2028, and which option (A floor/tiered or B 90/10) to carry as the base C1 BESS
    case; is the 100 MWh in the RFQ (not 10 MWh) correct?
14. MET: fixed PPA 43.96 (3 y) / 40.76 (4 y) – any indexation or GoO terms? Index PPA: confirm the downregulation
    compensation applies to curtailed MWh only. BESS floor + PSA: 50/50 share of what reference (net optimisation
    profit above the floor)? Is the 90 MWh their assumption or ours?
15. MVM: is 45 €/MWh for 2028–2032 pay-as-produced on 100 % of PV output, and is there any BESS instrument? (The
    e-mail's 2029–2034 series is for a wind site.)
16. Which offer pairings should be stored as sets (max 18 sets in total; proposal in §10)?
17. Yield report (P50/P90, clipping at 49.9 MVA, degradation) and capex/opex quotes for the 35 kV cable and substation
    bay (original: 3,700 and 9,500 k€).
