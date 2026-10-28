# Zalahaláp – SPEC v2: rebuild in the HU-wind-bids format

Owner: Fable (architecture). Builder: `model-builder` (Opus). Status: design for discussion, 28-Oct-2026.
Supersedes the layout parts of `SPEC.md` (§4, §5, §10, §12); the financial logic of `SPEC.md` §6–§9, §11, §13 stays
unless this document says otherwise. User request (28-Oct-2026): "this is the format I need … the offers and the
Aurora curve should be seen in a sheet, very straightforward". Format reference: `reference/HU_Wind_bids_271026.xlsx`.
Assumptions and the offer term sheet: `ASSUMPTIONS.md` (same folder).

---

## 1. The format, dissected (`HU_Wind_bids_271026.xlsx`)

Two sheets, no macros (.xlsx), no Cover, no Curves, no Calc, no Checks sheet. Font Aptos Narrow 10, gridlines off,
zoom 85 %. Defined names for every scalar input (`FX`, `CIT_Rate`, `Life`, `First_Year`, …), for the Dashboard levers
(`Dash_Case`, `Dash_Hours`, …) and for the band tables (`Osz1_LB`, `Osz1_Pts`, …).

### 1.1 Dashboard (B2:AS358) – everything on one page, bids side by side

| What | Where | Convention |
|---|---|---|
| Title | C2 | bold, plain text ("HU wind projects"); no banner formula, no version block |
| Master check | C3 label, D3 `flag`, **H3 `=Master_Check`** | number format `"CHECK" 0;"CHECK" -0;"OK"`, conditional font green `00B050` when 0, red `FF0000` otherwise (same on H356) |
| Columns | A 2 · B markers (`x` on every section bar) · **C 56 label** · **D 10 unit** (italic grey `808080`) · **E, F, G 13 = one column per bid** (Bora, Iváncsa, Dunaújváros) · **H 15 = Total** (`=SUM(E:G)`) or Live · I 2 spacer · **J → AS 10–13 = annual timeline** (36 years) |
| Section bar | B..H, fill `00563C`, white bold; the bid names repeated in E:G of every bar, "Total" in H (e.g. B5:H5, B30:H30, B246:H246) | one bar per topic; the bar is also the column header of the block below it |
| Sub-header | C..H bold `3C3C3B` on light green `E6F4EC` (C105, C113, C120, C127, C138, C201, C219, C231) | strategy / sub-block titles inside a section |
| Inputs **on the Dashboard** | E6:G20 ("Key inputs: the bids"), E24:G28 (case table), H23 (case selector) | blue `0000FF` on `FFF2CC`, thin border all sides, right-aligned; the levers that differ per bid live here, not on Inputs |
| Case selector | B22:H22 bar "Case: …" with **E22 Base, F22 Low, G22 High, H22 Live**; H23 blue, list validation `"Base,Low,High"` (name `Dash_Case`); H24 `=INDEX(E24:G24,MATCH(Dash_Case,$E$22:$G$22,0))` | the only scenario mechanism: three live columns, no macro, no stored results |
| Calculations | all on the Dashboard, by bid column: points (E31 `=INDEX(Osz1_Pts,1+SUMPRODUCT((Osz1_LB<E11)+(Osz1_LB=E11)*Osz1_Incl))*Osz1_W`), securities in HUF (rows 47–55) then EUR k (58–66), cash to RtB by nature (69–74), capex to COD (77–93), devex (96–102) | one formula per bid cell; same formula across E:G; H totals |
| Returns | rows 104–133, section "Returns RTB COD HTM", sub-headers per strategy: RtB exit case A (106–112), RtB exit case B (114–119), COD exit (121–126), Hold (128–133). Rows: exposure, proceeds, **profit pre-tax, profit post-tax, MOIC, IRR**; hold: capex, revenue, opex, EBITDA, **Project IRR pre-tax, NPV** | one figure per row, nothing repeated |
| Annual summary | rows 135–160: "Year" row 136 (`=J247`), then per bid a 6-row block (spending, revenue, opex, EBITDA, net CF, cumulative `=SUM($J143:J143)`) linking to the detail blocks | the reader's per-year view |
| Detail | rows 162–345 **grouped (outline level 1, collapsed)**: "Detailed calculations … – expand the group" | claims (164–186), waterfall (188–199), **dated FEP flows (201–213)**, case B variants (215–244), annual timeline (246–248), per-bid annual blocks (250–279, 281–310, 312–341), tax on the COD gain (343–345) |
| Dated flows | rows 202–210: amounts in E:G, **dates in J:L** (sub-header J201 "Date: Bora", K201 "Iváncsa", L201 "Dunaújváros"); `E243 =XIRR(E232:E238,J232:J238)`, `E242 =SUMIF(E232:E238,">0")/-SUMIF(E232:E238,"<0")` | the RtB route is a short list of dated flows, not a timeline |
| Claims (12 %/20 %) | accrual factors as closed forms: E169 `=(1+Rate_Cash*E168/Days_In_Year)*(1+Rate_Cash)^(YEAR(E165)-YEAR(Inputs!E$27)-1)*E166` (days to the first 31-Dec, full years, stub) | act/365, 31-Dec capitalisation without a timeline |
| Waterfall | E192 `=-MIN(E191,Sale_Costs)`, E193 `=-MIN(SUM(E191:E192),E190)`, … E197 `=SUM(E191:E196)` distributable, E198 FEP share | `MIN(cash left, due)` per step |
| Annual timeline | J247 `=First_Year+COLUMN()-COLUMN($J$1)`, J248 `=DATE(J247+1,1,1)-1` | calendar years, year-end dates |
| Per-bid annual block (30 rows) | J251 years from COD year; J252 construction months by year end, J253 share `=MAX(0,MIN(J252,$E252)-MAX(J252-Months_In_Year,0))/$E252`; J254/J255 operating months/share; J256 COD-year flag; J257 outlays `=SUMPRODUCT((YEAR($J$202:$J$209)=J$247)*$E$202:$E$209)`; J259 capex `=-$E$87*J253`; J262 degradation; J263 generation; J264 price; J265 revenue; J266 opex index; 267–273 opex lines; 274 total; 275 EBITDA; 276 unlevered pre-tax CF; J277 DF to COD `=(1+Discount_Rate)^((Inputs!$E$33-J$248)/Days_In_Year)`; J278 EBITDA × DF (**H278 = COD price**); J279 FEP CF on the COD route | identical blocks; only the bid column letter changes (`$E$44` → `$F$44`, `Inputs!$E$33` → `$F$33`); row constants sit in the bid's own column (E252, F283, G314); H = `SUM(J:AS)` |
| Checks | rows 347–356 on the Dashboard: `=N(ABS(…)>Tolerance)` per bid, H sum, H356 `=SUM(H348:H355)` = Master_Check | 8 checks; `***END***` at C358 |
| Number formats | money `#,##0;\(#,##0\);"–"`, 1 dp `#,##0.0;…`, % `0.0%;\(0.0%\);"–"`, multiple `0.00"x"`, date `dd\-mmm\-yy`, flag `0;\-0;"–"`, factor `0.000`, accrual `0.0000`, counts `0` | as `MODELLING_STANDARDS.md` §5 |
| Emphasis | totals bold with top border (E41, E55, E66, E87, E91, E239); subtotal with bottom border (E71); memo rows italic (C244) | |
| Freeze / grouping | freeze `E86` (artefact: 85 frozen rows; Inputs freezes `E5`), rows 163–358 outline level 1, summaryBelow off | |
| Density | 358 rows, every number visible, no charts, no notes column on the Dashboard; sources only on Inputs | |

### 1.2 Inputs (B2:P112)

C 54 label · D 11 unit · **E, F, G 13 per bid** (common inputs "all projects; value in column E") · **H 50 Source** (italic
grey, "Bora Inputs E11; Bid_Cash E45"). Section bar B4:H4 "Inputs" with the bid names; sub-headers (E6F4EC) by topic:
tender and grid; per project fees, devex, dates; RtB exit terms; capex; generation; opex; valuation, timeline and
constants (`Thousand`, `Days_In_Year`, `Months_In_Year`, `Tolerance` are inputs here, E79:E82). Band tables at the end
(rows 85–110, grouped) as flat rows `Osz1_LB`, `Osz1_Incl`, `Osz1_Pts`, weights. Freeze E5. A few per-project
formulas sit on input fill (F18 `=$E18/SUM($E$18:$E$23)*F$24`, black font) – we keep inputs blue only.

### 1.3 What "looks and works like this" means for Zalahaláp

1. One Dashboard: levers that differ **per case at the top in blue**, a three-column live sensitivity case (Base/Low/High),
   the results per strategy **side by side in one column per case**, a per-year summary, the detail grouped below.
2. Inputs sheet with a Source column; common values in column E, per-case values across the case columns.
3. **Annual** timeline, calendar years, year-end dates; the RtB route as dated flows with closed-form accruals.
4. No macros, no stored-scenario table (the live Base/Low/High replaces the 18 key-assumption sets).
5. Identical calculation blocks per case, same rows, only the case column letter differs.

---

## 2. Restructure design

### 2.1 Sheets (5; the format's two plus the sheet the user asked for, the grouped detail and the house check sheet)

| Sheet | Role | Format precedent |
|---|---|---|
| `Dashboard` | levers per case, sensitivity case, capacity/energy, revenue, cash to RtB, capex to COD, returns by strategy × case, JV waterfalls, debt and tax, annual summary per case, dated RtB blocks (grouped) | wind Dashboard rows 2–245 |
| `Inputs` | every other input, one row each, common in E, per-case across E:I, Source in J | wind Inputs |
| `Revenue` | **new, the user's ask**: Aurora curves used, the three offers' terms side by side (PDF page cited), and what each case earns per year (PV €/MWh, BESS k€/MW) | – (wind: `Case_PPA` only) |
| `Calc` | five identical annual blocks, one per case (the wind format's rows 250–341 moved off the Dashboard: 5 × ~130 rows would bury it) | wind rows 246–345 |
| `Checks` | every check row → `Master_Check` (CLAUDE.md rule 6 is non-negotiable); the master check is mirrored at Dashboard H3 as in the format | wind rows 347–356 |

Gone: `Curves` (its content splits into `Revenue` (Aurora, offers, bands) and `Inputs` (early-works and deposit tables)),
`Calc_1..3` (→ `Calc` blocks), the 18 key-assumption sets, `Scenarios.bas`, `Solver.bas` (no circularity, §2.6),
`GoalSeek.bas` (implied prices stay closed-form). Output: `out/Zalahalap_Model.xlsx` (+ `_recalc.xlsx`).

### 2.2 Column grid (same letters on Dashboard, Revenue and Calc so the years line up)

A 2 · B 3 (`x` markers) · C 56 · D 10 · **E..I 13 = the five cases** · J 2 spacer · **K..AW 10 = 2025 … 2063** (39 years:
acquisition Nov-2025 to COD Mar-2028 + 35 years). On `Inputs` J = Source (50). Freeze `E5` everywhere. Section bars
B..I; H is **not** a total column (one project) – it is the fourth case column. Where the format puts "Total" in H, we put
the lifetime total in the block's own column of `Calc` (H = `SUM(K:AW)` as the wind H column) – see §2.5.

### 2.3 The five cases (Dashboard columns) – one consistent configuration each

| Col | Case | PV / grid / yield | BESS | PV product | BESS product | Purpose |
|---|---|---|---|---|---|---|
| E | **v3 original** | 61 MWp / 50 MVA / 1,250 | 4 h, v3 curves | none | none (v3 stack) | reconciliation column; grouped (hidden) in the IC view |
| F | **Aurora merchant** | 63 / 49.9 / 1,265 | **4 h**, DC 4h block | none | none (Aurora) | base case; flip hours to 2 for the like-for-like with the offers |
| G | **E.ON** | 63 / 49.9 / 1,265 | **2 h** (as offered: 50 MW / 100 MWh), DC 2h block | E.ON DAM-indexed PPA, 3 y from COD | option A (2028 bands) then option B; lever "E.ON B" | offer C1 |
| H | **MET** | 63 / 49.9 / 1,265 | **2 h** (offered 90 MWh; 100 MWh built, fee not scaled), DC 2h block | MET fixed 43.96 (3 y) / 40.76 (4 y) or MET indexed | MET tolling 49.7/46.4 or floor + PSA 45.6/41.9 | offer C2 |
| I | **MVM** | 63 / 49.9 / 1,265 | 4 h, DC 4h block (no BESS instrument) | MVM fixed 45, 2028–2032 | none (Aurora merchant) | offer C3 |

The three strategies (sale at RtB, sale at COD, hold) are **row blocks** under "Returns", each with the five case columns
– exactly the wind layout (strategies as sub-headers, bids as columns). Offer sub-options are text levers per column
(wind precedent: "Turbine route (EU or CN)" E17:G17).

### 2.4 Dashboard – section order (top to bottom)

1. **C2** "Project Zalahaláp – 63 MWp PV + 50 MW BESS: sale at RtB, sale at COD or hold"; **C3/H3** master check.
2. **Key inputs: the cases** (bar with the five case names in E:I; blue levers, names `Dash_*`): case name (text); PV
   capacity MWp; grid export MVA; specific yield MWh/MWp; BESS built (flag); BESS power MW; **BESS duration h**; Aurora
   block (formula: `="DC "&hours&"h"`, v3: "v3"); PV product (text, validated: none / E.ON index / MET fixed / MET index /
   MVM fixed); BESS product (none / E.ON A / E.ON B / MET tolling / MET floor); offer tenor years (3 or 4; MVM 5); RtB
   price k€/MWp (0 = implied); COD price (0 = implied); gearing cap (v3: 0). 14 rows.
3. **Sensitivity case (Base, Low, High)** – bar E "Base", F "Low", G "High", H "Live"; H selector validated; rows:
   PV price change %, BESS revenue haircut %, yield change %, capex change %, BESS capex k€/MWh, buyer target IRR,
   COD delay months, debt rate. Live `=INDEX(E:G,MATCH(Dash_Case,…))`. 8 rows. These replace sets 3–8, 10–11, 14, 16;
   sets 9, 12, 13, 17, 18 are now column levers; set 2 is column E + the v3 switches on Inputs; set 15 is an Inputs flag.
4. **Capacity and energy, first full year**: generation, sent to the BESS, exported directly, BESS discharge, cycles,
   PV price received (blended, €/MWh), BESS revenue per MW (links to `Calc`).
5. **Revenue and EBITDA**: first full year and lifetime: PV revenue, BESS revenue, GoO, total, opex, EBITDA, margin.
6. **Cash to RtB by nature** (format rows 68–74): SPV price, transaction costs, DSA loan, early works (4 lines), pre-RtB
   development, deposits and guarantees, guarantee fees; split sunk / reimbursed by the buyer / returned; FEP exposure.
7. **Capex to COD**: PV, MV cable, substation, land, BESS (MW × h × k€/MWh), early works at NTP, capitalised interest and
   fees, DSRA; total uses; debt; equity; per MWp.
8. **Returns** (sub-headers): *Sale at RtB*: price used, **implied price at the target IRR**, all-in price incl.
   reimbursements, buyer IRR unlevered after tax / before tax / levered, FEP equity, proceeds, profit pre- and post-tax,
   MOIC, IRR, NPV, PSH profit. *Sale at COD*: EV used, implied EV, buyer IRRs, FEP contributions, proceeds, profit, MOIC,
   IRR, NPV, PSH. *Hold*: FEP equity, distributions, profit, IRR levered after/before tax, MOIC, NPV, project IRR after/
   before tax, PSH buy-out. (Current Dashboard rows 28–86, same content, five columns.)
9. **JV waterfall at the RtB exit** and **at COD** (paid amounts, as current rows 130–153).
10. **Debt and tax**: debt at COD, k, DSCR capacity used, min DSCR, DSRA, lifetime tax levered/unlevered, losses lapsed.
11. **Annual summary per case** (format rows 135–160): year row, then per case 6 rows: investment, revenue, opex, EBITDA,
    net cash flow (unlevered after tax), cumulative. Links to `Calc`.
12. **Detail – expand the group** (outline level 1): per case the **dated RtB-route blocks**: dated lines (amount E:I,
    date K:O, sub-header "Date: v3 | Aurora | E.ON | MET | MVM"), accrual factors (format rows 165–177 pattern), claims,
    guarantee fees, waterfall at the exit, FEP dated flows → XIRR/MOIC, PSH dated flows, tax on the return.
13. `***END***`.

### 2.5 `Calc` – five identical annual blocks (one per case; the block's case column is a constant in E of its bar)

Rows per block (≈ 130; one formula per row across K:AW; H = lifetime `SUM`; E/F row constants pulled from the case's
Dashboard column with `INDEX(Dash_<lever>,1,$E$<bar>)`):

- **Timing**: years from COD year; construction months / share (format J252:J253); operating months / share (J254:J255);
  COD-year flag; exit-year flag; operating years elapsed; repower-year flag; offer windows (PV, BESS) from `Revenue`.
- **Energy**: generation (MWp × yield × availability × degradation × operating share), sent to the BESS
  (`MIN(generation, −charge_from_PV per MW × BESS MW × share)`), exported directly, discharge (PV-coverable), cycles.
- **Revenue**: PV direct export × PV price received; PV stored × stored-energy price; GoO; BESS MW × BESS revenue per
  MW; total; contracted share (DSCR weight). Prices per year come **from `Revenue` rows by case** (§3).
- **Opex**: O&M + cable O&M; balancing **× (1 − PV offer in force)** (the offers include balancing – current model double
  counts); lease; BESS opex; offer fixed fees; HIPA + innovation; total; EBITDA.
- **Capex**: construction (straight line NTP → COD by construction share), early works at NTP, augmentation at the
  repower year (indexed), total.
- **Pre-COD outlays on the annual grid**: `SUMPRODUCT((YEAR(dates)=year)*amounts)` from the dated block (format J257);
  deposits returned at COD on the hold/COD routes.
- **Debt** (closed form, no macro): cost basis; IDC per € of debt (one row: `rate × cumulative construction share at the
  year start`, summed); k = `MIN(gearing, DSCR capacity ÷ cost × (1 + IDC factor), €/MW quantum)`; capacity =
  `XNPV(rate, CFADS_unlevered ÷ DSCR over the tenor)`; draws = k × capex; IDC; balance; stub interest in the COD year;
  level annuity over the tenor from the next year-end; interest; repayment; upfront and commitment fees; DSRA funded at
  COD, released at tenor end; DSCR row (actual CFADS ÷ service).
- **Tax**: depreciation PV/grid/development 30 y, BESS 10 y, augmentation; tax base levered and unlevered; loss
  corkscrew (50 % cap, lapse) × 2; CIT; Robin Hood (flag); HIPA sits in opex.
- **Cash flows**: project before tax; project after unlevered tax; equity levered after actual tax; buyer at RtB
  (unlevered after/before tax, levered) from the exit year; DF to the exit year end → **implied RtB EV = XNPV** (floored
  at 0, flag); buyer at COD; implied COD EV; FEP COD-route flows (contributions, waterfall receipts); hold flows (FEP
  share, PSH buy-out at the exit); PSH flows.

Returns (XIRR with K$4:AW$4 year-end dates; MOIC by SUMIF; XNPV at `Discount_Rate`) are computed in the block's E/F
constants and linked to the Dashboard – the format's E119/E126/E132 pattern.

### 2.6 Logic kept, logic changed, and what wins where the format conflicts with the standards

| Topic | Decision | Why |
|---|---|---|
| Timeline | **Annual** (format) instead of quarterly | the format is annual; Aurora is annual; the RtB route keeps exact dates through the dated block (better than quarterly: the quarterly grid moved the v3 RtB XIRR from 224 % to 170 %) |
| RtB route | dated flows + closed-form accruals (format rows 165–213) | exact dates for the 12 % return, guarantee fees, reimbursements; XIRR on the list |
| Buyer at RtB / COD | annual flows, price at the exit (COD) year end, DF to that year end | one convention for the IRR and the implied price; the quarterly → annual effect goes into the reconciliation bridge |
| Debt circularity | sized on **CFADS after unlevered tax**, IDC factor closed-form | no COPY/PASTE, no Solver.bas; actual DSCR ≥ sizing DSCR by construction |
| Implied prices | closed-form XNPV (unchanged) | no goal seek |
| Scenarios | Base/Low/High live case + per-column levers; **no stored results, no macros** | format; MODELLING_STANDARDS §7 only requires a macro where a value must be stored – nothing must |
| Inputs on the Dashboard | the per-case levers (blue, named `Dash_*`) | format wins over §6.1 "no hardcodes outside Inputs/Curves"; the wind model does the same |
| Sub-header colour | `E6F4EC` light green (format) over `F2F2F2` grey | format wins |
| Banner / Cover | plain bold title (format); no Cover, no `=Cover!$C$3` banner | format wins |
| Checks | separate `Checks` sheet (CLAUDE.md rule 6) + master check at Dashboard H3 (format) | both |
| Literals, one formula per row, blue inputs, sources, short formulas | standards unchanged | |
| Freeze panes | `E5` (format's Inputs), not the Dashboard's `E86` artefact | |
| Reconciliation | column E = v3 case with the v3 switches on Inputs (Fix 2/3/5/6/7/8, share slide, 6 early-works lines, SPV price 150, DSA 0, HIPA 0, BESS dep. 30 y, losses 40 y, PSH land step, EW to PSH); targets block on `Checks` ties within 0.1 %; `reconcile.py` gets an annual shadow and a new bridge step "annual timeline" | the v3 targets are monthly/quarterly; the step is quantified as the quarterly one was |

### 2.7 Offer logic – the straightforward rule (replaces SPEC §6.3 "Modelling")

The project earns three things: PV exported directly, PV stored in the BESS, and the BESS uplift. Each product says
where it settles. Per case and year (`Revenue` §3.3):

| Product | PV exported directly, €/MWh | PV stored in the BESS, €/MWh | BESS owner revenue, k€/MW/yr | Source |
|---|---|---|---|---|
| Aurora merchant | direct-export price of the block × index | transfer price of the block × index | storage total × PV-only factor × index × (1 − haircut) (+ aFRR-down share) = "uplift kept" | Aurora |
| E.ON PPA + A/B (settles at the **POD**: all fed-in energy at DAM − fee) | DAM proxy − MAX(7.6 % × DAM, 5) | transfer price (the owner's own energy; its value is inside the uplift) | A: `MAX(floor, uplift − band fee)`; B: `0.9 × uplift`; **minus the handling fee on discharged MWh** (`MAX(7.6 % × discharge revenue, 5 €/MWh × discharge)`); band years 2027/2028, then B (flag to keep the 2028 bands) | E.ON p4–5 |
| MET fixed / indexed + tolling (settles at the **PV meter**: pay-as-produced on PV output; MET dispatches the BESS) | fixed price by tenor, or DAM proxy − MIN(10, MAX(5 % × DAM, 5.75)) | **the same PV price** (all PV output is paid as produced) | tolling fee by tenor (49.7 / 46.4) | MET p3–5 |
| MET + floor + PSA | as above | as above | `floor + 50 % × MAX(0, uplift − floor)` (45.6 / 41.9) | MET p6 |
| MVM fixed (reading: PV meter; BESS merchant) | 45 (2028–2032) | 45 | uplift kept (merchant) | MVM e-mail |

After the tenor every product falls back to the Aurora merchant rows of the case's block. DAM proxy = the block's
direct-export price (annual; the fee kinks act on 15-min prices – stated on `Revenue`). "Uplift" = Aurora's storage
total (net of the PV energy's export value, s+20), kept at the PV-only factor – the current model's base; the gross
reading (before s+20) is a term-sheet switch `Offer_Uplift_Gross` (0) to be confirmed with E.ON/MET (ASSUMPTIONS §2.4).

Changes vs the current model that this rule implies: (1) balancing cost not charged under an offer; (2) E.ON handling fee
on discharged MWh; (3) MET fixed/indexed price on stored PV energy (PV-meter settlement) instead of the transfer price;
(4) BESS hours per column (2 h for E.ON and MET) instead of a 4 h battery earning a 1.8–2 h fee; (5) MET 4-year options
and the tenor lever; (6) E.ON option A/B and MET tolling/floor as column levers instead of key-assumption sets.

---

## 3. `Revenue` sheet – layout (the one-screen view of the Aurora curves, the offers and what each case earns)

Columns as §2.2 (E:I case/offer columns, H/J as stated, K:AW years). Three sections, ~95 rows. Blue = input with
source; black = formula; every number traceable to a CSV row or a PDF page.

### 3.1 Aurora curves used (bar "Aurora Q2-2026 Hungary, Central, real 2025 €; data/aurora_curves.csv, source rows cited")
Row 4 year header (K4 2025 … AW4 2063; Aurora 2026–2060, held flat real after 2060, indexed with +2 % as today).
One row per curve, D unit, E = CSV series/row ("Degraded batteries!166"), blue values across K:AW:

| # | Row | Unit | Blocks |
|---|---|---|---|
| 1 | Baseload (information) | EUR/MWh | Central prices!20 |
| 2–3 | PV direct-export price, DC 4h / DC 2h | EUR/MWh | (s+34 + s+35) ÷ s+31 |
| 4–5 | PV transfer price (energy sent to the BESS), DC 4h / DC 2h | EUR/MWh | s+36 ÷ s+32 |
| 6–7 | BESS storage total, DC 4h / DC 2h | k€/MW/yr | s+27 |
| 8–9 | PV-only factor (no grid charging), DC 4h / DC 2h | factor | aurora_extract.py |
| 10–11 | aFRR-down + FCR share of the total, DC 4h / DC 2h | factor | (s+23 + s+24) ÷ s+27 |
| 12–13 | BESS charge from PV per MW, DC 4h / DC 2h | MWh/MW | s+13 |
| 14–15 | BESS discharge per MW, DC 4h / DC 2h | MWh/MW | s+11 |
| 16 | Euro index (2025 = 1) | index | Common Assumptions!14 |
| 17–20 | v3 curves (grouped): v3 capture, v3 BESS stack, v3 discharge, v3 charge from PV | | old_curves.csv |
| 21 | Repower year by block (DC 4h 2048, DC 2h 2048, v3 2048) | year | s+1 |

Dropped from the workbook (kept in the CSV): the 1.5-cycle, standalone and 2030-entry blocks (the user confirmed DC
co-location and no grid charging; nobody will select them), the Low/High ±15 % v3 curves, the component rows s+14…s+26
(information only), the unused blended/DA÷grid curves.

### 3.2 Offer term sheet (bar "Offtake offers as stated in the PDFs – indicative; blue = as offered; H = source page")
Columns E = E.ON, F = MET (E2 Hungary), G = MVM ONEnergy, H = source, I = "our reading" flag where needed. Rows (text
or number; dates as dates): offeror / document / date; validity; PV product; fixed price 3 y; fixed price 4 y; fee % of
DAM; fee floor; fee cap; contract start; tenor (max / options); volume; balancing and scheduling; balance-group fee;
GoO; indexation; curtailment / downregulation; **settlement point (our reading: POD / PV meter)**; BESS product; BESS
configuration offered (MW / MWh / RTE / cycles); tolling fee 3 y / 4 y; floor 3 y / 4 y; share above the floor; option
B share; profit definition; offeror's yield assumption. Then the **E.ON option A band table** (2 rows × 5 bands: lower
bounds 2027 `0/47/52/80/109`, 2028 `0/38/43/67/91`; fee in band `0/100/10/15/18 %`; marginal-rate row as today) and
the switches `Offer_Uplift_Gross` (0), `EON_Bands_After_2028` (0). Full content: `ASSUMPTIONS.md` §2.
Named ranges per term (`Offer_Fixed_3y`, `Offer_Fee_Pct`, …, one range E:G each) so `Calc` picks by product column.

### 3.3 What each case earns per year, nominal (bar "Resulting prices by case – offer in its tenor, then the Aurora merchant tail")
Year header repeated. For each of the five cases (sub-header with the case name from the Dashboard), the rows:

| Row | Formula (one per row, K:AW) | E / F constants |
|---|---|---|
| Merchant PV price, direct export | `=INDEX(direct price of the block, year) × index × (1 + PV price change)` | E = block |
| Offer in force (PV) | `=(year ≥ start year) × (year ≤ end year)` | E start, F end (from product, tenor, COD) |
| Handling fee | `=MIN(cap, MAX(pct × merchant, floor))` | E pct, F floor/cap from the term sheet |
| Offer PV price | `=fixed_flag × fixed price + (1 − fixed_flag) × (merchant − fee)` | E fixed flag, F fixed price by tenor |
| **PV price received, direct export** | `=merchant + window × (offer − merchant)` | H = lifetime generation-weighted average |
| **PV price on stored energy** | `=transfer × index + window × settle_at_PV_meter × (PV price received − transfer × index)` | E settlement flag |
| BESS uplift kept (merchant) | `=storage total × (PV-only + share × down-cap) × index × (1 − haircut)` | E block |
| E.ON band fee / handling fee on discharge | two rows, E.ON columns only (0 elsewhere) | |
| Offer BESS revenue to the owner | `=tolling + floor-PSA + E.ON A + E.ON B` by product flags (one row, ≤ 2 nested) | E product |
| **BESS revenue received** | `=uplift + window × (owner − uplift)` | H = lifetime average |

≈ 10 rows × 5 cases. `Calc` links only to the two bold PV rows and the bold BESS row of its case. The Dashboard's
"Revenue and EBITDA" block shows the first-full-year and lifetime figures.

---

## 4. Opus task list (`model-builder`; audit by `model-auditor` after every step)

0. **Engine**: (a) `Model(periodicity="annual")` with calendar years and year-end dates (exists); (b) a `Block`
   helper that writes the same row set n times on one sheet with the case column as a parameter (`$E$` → `$F$` …) and
   returns row keys per block; (c) a `DatedBlock` helper (amounts E:I, dates K:O, accrual factor rows, XIRR); (d)
   Dashboard lever rows per column (`Dash_*` names, validation lists); (e) sub-header style `E6F4EC`; (f) `package()`
   to `.xlsx` when no macros; (g) lint: allow blue inputs on the Dashboard lever rows only. Tests green.
1. **Inputs** (§2.4 item 2 for what moves to the Dashboard): sections dates (per case), PV, BESS, capex, development /
   early works / deposits (the dated tables with HUF, FX, dates, reimbursed and returned flags), opex, financing, tax,
   JV and exit, valuation and constants, **v3 reconstruction switches (per case)**. Source column J, ≤ 60 characters.
2. **Revenue** per §3; `data/aurora_extract.py` unchanged; the offer terms from `ASSUMPTIONS.md` §2 with page refs.
3. **Calc** per §2.5, five blocks; closed-form debt; unlevered tax for sizing; loss corkscrews; buyer and implied
   price rows; COD and hold routes; PSH rows.
4. **Dashboard** per §2.4; dated RtB blocks in the group; Base/Low/High.
5. **Checks**: levers valid (product names, block exists, hours ∈ {2, 4}, tenor ∈ {3, 4, 5}), energy balance, discharge ≤
   charge × RTE, waterfalls close, debt repaid, DSRA released, losses ≥ 0, XIRR numeric, implied price ≥ 0 flag, **v3
   targets within tolerance**, Master_Check = 0.
6. **Reconciliation**: `reconcile.py` annual shadow; bridge rows: v3 → fixes → user decisions → **annual timeline** →
   tax → JV → debt → base; `out/reconciliation.md`.
7. `python3 -m fepm audit projects/zalahalap`; commit "Rebuild Zalahalap in the HU-wind-bids format".
