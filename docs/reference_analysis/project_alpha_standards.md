# Project Alpha – IRR model: modelling standards (reverse-engineered)

Source: `ref/Project_Alpha_-_IRR_model_20260806_WIP.xlsb` (Excel binary, macro-enabled), read via the
LibreOffice-converted `conv/...xlsx` (openpyxl) plus direct BIFF12 decoding of the xlsb for the formulas, named
ranges and calc settings that the conversion broke. Author per comments: Alexander Gibbs (FEP); SharePoint path
`.../FEP/02_ONGOING_DEALS/Baltics/Saules Energy Dec-25/...`. Status: WIP (see Notes sheet).

Conversion caveats: the IRR rows (`XIRR`, BIFF12 function id 0x1AD) and every formula that uses a defined name
show as `#NAME?` in the xlsx. They were decoded from the xlsb (see §4, §5). Defined names lost their references
in the xlsx and were recovered from `workbook.bin`.

---

## 1. Sheets: list, order, purpose

| # | Sheet | Size | Purpose |
|---|-------|------|---------|
| 1 | `Notes` | B2:B8 | WIP to-do list ("Funding only considers equity – add debt and create funding macro", "Add BS", "Relink PV premium"...). No cover page, version log or index. |
| 2 | `Stat_Inp` | A1:AI303 | **Static inputs + dashboard.** Top: macro checks (rows 2-4), output summary (rows 6-19), scenario selector (row 21-22). Then all scalar inputs by theme (rows 24-229) with a scenario matrix. Bottom: sensitivities (tornado, 2-D data tables, rows 230-249) and an IRR bridge (rows 294-303). |
| 3 | `Dyn_Inp` | A1:AT73 | **Dynamic (time-series) inputs.** Annual curves by calendar year (inflation, PV price, BESS revenue stack, base rate) and monthly construction CAPEX profiles. Each block has a selector row ("... in use") plus a library of alternative curves. |
| 4 | `Model` | A1:FN565 | **The single calculation sheet**, quarterly. Timing → flags → inflation → returns/valuation (placed at the top) → cash flow statement → income statement → generation → devex → CAPEX → premium → revenue → OPEX → funding/S&U → SHL → equity → financing (debt) → working capital → taxes. Ends with `A565 = "End of sheet"`. |
| 5 | `Charts` | A1:FN177 | Chart feeds and presentation calcs: same timing header as Model, pasted per-case cash flows, annualisation via SUMIFS, €m conversions, CAPEX benchmark comparison; holds 6 line/bar charts. |

There is **no separate checks sheet, no cover sheet and no separate outputs sheet**: outputs sit at the top of
`Stat_Inp` (next to the inputs, so the user sees the effect of an input change on one screen), and returns are
calculated near the top of `Model` (rows 36-115) before the detailed operations blocks.

Charts (8): `chart1` tornado (horizontal clustered bar on `Stat_Inp!F236:G241`), `chartEx1` waterfall (IRR bridge,
names `_xlchart.v1.0/1` → `Stat_Inp!B299:B303 / D299:D303`), line charts of equity cash flows, cumulative €m,
revenue curves (€k/MW p.a.), stacked column CAPEX comparison (FEP vs Aurora).

## 2. Layout conventions

### Model / Charts (time-series sheets) – identical skeleton
| Column | Width | Content |
|---|---|---|
| A | 3.6 | Marker column. `x` with orange fill `FFC000` flags rows of the Funding/SHL/Equity block (A286-A415, rows still being built), A69/A70 marked too. |
| B | 3.6 | Section marker `x` on every section header row (B11 Flags, B29 Inflation, B36 Valuation, B116 Cash flow statement, B129 Income statement, B151, B174, B180, B216, B256, B286, B371, B397, B417, B530, B537) – used for navigation (Ctrl+arrow). |
| C | 37.0 | Row label. Labels are often **linked** to the input label (`C184 = Stat_Inp!B88`, `C225 = C221`) so renaming propagates. |
| D | 9.1 | Unit (`€k`, `%`, `MWh`, `Flag`, `Count`, `Index`, `Month`, `x`, `Check`). |
| E | 12.8 | Row constant – the input pulled from `Stat_Inp` (`E154 = Stat_Inp!F67`), green font, centred, thin box. Also used for single-cell results (`E53` Project IRR, `E54` MOIC, `E479` debt quantum). |
| F, G (H) | default | Secondary constants (end date of a flag window `F13 = Stat_Inp!F41`, sensitivity % `F231`, "At acquisition / At COD" split `F206/G206`) or memo text (`F258 '€k / MW p.a.'`, `G258 'Inflated'`), memo KPIs (`H272 = I272/(45*15)` avg cost per MW p.a.). |
| I | 12.8 | **Total / "Sum" column**: `=SUM(J:FN)` on every flow row (`I46 = SUM(J46:FN46)`), light grey fill (theme0 −5% ≈ `F2F2F2`). Also holds the period-0 seed for corkscrews/counters (`I14 = 0` grey `D9D9D9`, `I33 = 1` index base). Balances have no total. |
| J:FN | 10.1 | Timeline: **160 quarters**, `J3 = 1-Oct-2025` → `FN4 = 30-Sep-2065`. Nothing to the right of FN (FP, JZ, PM, SG set hidden). |

Header block rows 2-9 (navy bar on row 2 "Timing", with `I2 = "Sum"`):
`3 Period start` (`J3` hard date, `K3 = J4+1`), `4 Period end` (`=EOMONTH(J3,2)`), `5 Days in period`, `6 Days in year`,
`7 Fraction of year` (`=J5/J6`, used for every annual→quarterly conversion), `8 Quarter` (`=MONTH(J4)/3`), `9 Year`.
Rows 11-44 are flags/counters. Charts repeats rows 2-9 exactly.

Freeze panes: Model `E331` (looks accidental — freeze is normally on J10), Charts `K105`. Gridlines off on
Stat_Inp/Dyn_Inp/Model/Charts; zoom 85% on Stat_Inp and Model. No row/column outline grouping on Model (only 3
hidden rows grouped in Charts). Default row height 15.

### Stat_Inp
`B` label (33.4), `C` unit (italic, 9.1), `D` notes (red italic), `E` memo calc (e.g. `E94 = F94 * $F$76` €k/MW → €k),
**`F` = "Live" value used by the model** (17.6), `G` notes/comments, **`H:Q` = scenario slots 1-10** (header row 24
`1, =H24+1 ...`; names in row 25: "Target case (2hrs, Aurora)", "Vendor", "2.6 hrs (Capalo)", "2.0 hrs (Capalo)",
..., "Downside", "Test", "Spare"). Inputs are grouped in titled blocks (bold title, then optional asset sub-title PV/BESS):
Timing, Generation/Dispatch, DEVEX, CAPEX, Other initial costs, Revenue, OPEX, Natural Resource Tax, Equity financing,
Debt financing (Interest, Fees, Reserve account, Sizing & cash sweep), NWC, Tax, Valuation. Derived input lines
(end dates, totals) sit inline in column F (`F41 = MAX(EDATE(F39,F40)-1,0)`, `F46 = EDATE(F44,F45*12)-1`).

### Dyn_Inp
`B` label, `C` unit, `D` selector link / source note (`D19 'Fixed PV uncurtailed'`, `D33 'Reverts to Aurora after 2031'`),
`F:AT` = calendar years 2025-2065 (`F4 = 2025`, `G4 = F4+1`). Monthly CAPEX profiles reuse the grid as construction
month 1..n (`F53 = 1`), with a check-sum in `D` (`D57 = SUM(F57:AC57)` must be 100%). Library rows include
"Spare" placeholders so new curves can be added without breaking the INDEX range.

## 3. Formatting conventions

Font: **Aptos Narrow 11** everywhere (theme minor font; Aptos Display for headings theme). No size variation.

| Element | Format |
|---|---|
| Section header (sheet-level) | Fill `002060` (navy), **bold white** (theme0) text, fill runs across the full row incl. timeline (C2:FN2). E.g. Model `C2 'Timing'`, `C11 'Flags'`, Stat_Inp `B6 'Outputs'`, `B21 'Scenarios'`, `B230 'Sensitivities'`. |
| Sub-header (asset / block) | Fill theme `dk2` (`0E2841`) +90% tint ≈ very light navy (`E3E8F0`-ish), bold, e.g. Model `C12 'PV'`, Charts `C22 'Annual timeline'`, Stat_Inp `H243`. Input-sheet block titles (`B37 'Timing'`) are bold only. |
| Hard-coded input | Fill **`FFFFCC`** (light yellow), black text, **thin border** box, centred. 171 on Stat_Inp, 490 on Dyn_Inp, 323 on Model (macro paste rows `J433:FN433`, `J472:FN472`). Text inputs (selectors) styled the same. |
| Recently changed / to-review input | Fill `FFFF00` (bright yellow), e.g. Stat_Inp `J28`, `H114`, Dyn_Inp `G30`. Red fill `FF0000` on a label to flag a questionable input (`B115`). Grey `808080` fill = parked/old value (`Stat_Inp!R118`). |
| Formula | Black (theme1) font, no fill. |
| Cross-sheet link (input pulled into calc) | **Green font `00B050`** – Model E/F/G columns (`E13 = Stat_Inp!F39`), Stat_Inp output links (`F7 = Model!E53`), Dyn_Inp selector links (`D6 = Stat_Inp!F26`). 601 green links on Model. Exception: Charts `J12 = Model!J78` also green; links inside the timeline that pull from Dyn_Inp (`J221 INDEX(Dyn_Inp!...)`) stay black. |
| Derived/adjusted input curves | Dark red font `C00000` for formulas inside an input sheet (Dyn_Inp `H22 = H28 * $D22`, `D19` notes). |
| Notes / warnings | Italic red `FF0000` / `C00000` (`Stat_Inp!D208 'Note - must be 0 in debt sizing case'`, `Model!E499 'Needs S&U & funding waterfall'`). Cell comments hold sources ("Energy Pro 08-Jul + 5% contingency", "Implies Aurora availability of 98%"). |
| Totals column / computed inputs | Fill theme0 −5% (≈ `F2F2F2`): Model column I, Stat_Inp computed lines (`F35`, `F41`, `F97`). |
| Key result / subtotal rows | Thin top (and sometimes bottom) border across the row (`C50:FN50` FCF, `C69:FN69` total equity CF boxed). |
| Units column | Italic (Stat_Inp column C). |

Number formats (custom, with negatives in brackets and zero as a dash):
- Money: `#,##0;\(#,##0\);" -"` (43k cells on Model), 1-2 dp variants `#,##0.0;\(#,##0.0\);" -"`.
- Percent (timeline): `0.0%;" ("0.0%\);" -%"`, `0%;...`, `0.00%;...`; scalars `0.0%`.
- Multiples: `#,##0.00\x;\(#,##0.00"x)";" -x"` (MOIC, DSCR).
- Dates: `d\-mmm\-yy` (e.g. 1-Oct-25).
- Check cells: `"CHECK";" CHECK";" OK"` (any non-zero shows CHECK) + conditional format (`Stat_Inp C2:C4 cellIs = 0`).
- Dyn_Inp curves: `#,##0.0` and `0.0%`.

## 4. Formula conventions

**Universal row structure:** constant in `$E` (anchored column), timeline header via row-anchored refs (`J$4`, `J$7`),
row-relative logic in J copied unchanged to FN. Consistency test (openpyxl Translator, J vs every column K:FN):
**326 of 327 timeline formula rows are perfectly consistent**; the single deviation (row 258) is cosmetic
(`J$7` in J vs `K7` in K). First-period exceptions are avoided by using column I as period 0
(`K3 = J4+1` is the only one on the header; corkscrews read `I` as the prior closing balance).

Typical patterns (cell examples from Model):
- **Fractional period flag** (share of the quarter inside a date window):
  `J13 = MAX(0,(MIN($F13,J$4)-MAX($E13,J$3)+1)/J$5)+((($E13-$F13+1)/J$5)*($E13>J$3)*($F13<J$4))`
  with `$E` = start date and `$F` = end date linked from Stat_Inp. Reused for operations, terms 1-3, cash-sweep window, debt period.
- **Event flag**: `J43 = N(AND(J$3<=$E43,J$4>=$E43))`; **first/last period**: `J16 = N(AND(I15=0,J15>0))`, `J17 = N(AND(J15>0,K15=0))`.
- **Cumulative/ownership flag**: `J44 = SUM($J43:J43)`; **counter**: `J18 = IF(J15, J$9-YEAR($E15)+1, 0)`.
- **Indexation**: annual rate looked up by year `J30 = SUMIFS(Dyn_Inp!$F$6:$AT$6, Dyn_Inp!$F$4:$AT$4, J$9)`, step-up only in the
  chosen quarter (`J31`), index `J33 = IF($E33 = J$9, 1, I33 * (1 + J$32))` (base year = 1).
- **Annual curve → quarter**: `J221 = INDEX(Dyn_Inp!$F$16:$AT$16, MATCH(J$9, Dyn_Inp!$F$4:$AT$4, 0))`; real × index → nominal (`J225 = J221 * J$33`).
- **Quantity × price × time × flag × index** for every cost: `J272 = -1 * $E272 * J$7 * J$24 * J$163 * J$34`.
- **Sign convention**: costs, capex, taxes are negative in the P&L/cash-flow blocks via explicit `-1 *` at the link point
  (`J47 = -1 * J563`, `J49 = -1 * J201`); detail blocks (CAPEX, interest) are positive.
- **Subtotals** with `SUM(range)` or `SUM(a, b)` – never `a+b+c` chains (`J140 = SUM(J138:J139)`, `J146 = SUM(J140,J143:J145)`).
- **Corkscrews** opening / movements / closing: `J464 = I469`, `J469 = SUM(J464:J468)`; same for SHL (374-378), accrued
  interest (392-395), share capital (402-406), DSRF (504-507), depreciation (540-543), tax losses (558-561).
- **Cash waterfall** in priority order (rows 289-316): CFADS → fees → interest → principal → DSRF → SHL interest/principal → dividends.
  Payments use `-1 * MAX(MIN(available, due), 0)`.
- **Scenario-selector lookup** (Dyn_Inp): `F6 = INDEX(F$8:F$13, MATCH($D6,$B$8:$B$13,0))` where `$D6` = text selected in Stat_Inp.
- **Profiles in construction months**: `J182 = SUMIFS(Dyn_Inp!$F$55:$AC$55, Dyn_Inp!$F$53:$AC$53, ">"&I14, Dyn_Inp!$F$53:$AC$53, "<="&J14)`.
- **Division guarded** with `IFERROR(x, 0)` (`J141`, `J176`); booleans coerced with `N()` or `(cond > 0)`.
- Formula typography: spaces around operators and after commas (`=$E191 * $E$163 * J$189 * (1 + $F191)`), `= ` sometimes with a leading space.

**Hardcode avoidance**: all drivers enter via `$E`/`$F`/`$G` links (green). Remaining literals are structural
(1000 for €/MWh×MWh→€k, 966×; calendar 2/3/12/31) plus two smells: epsilon nudges `- (J43 * 0.0001)` /
`0.000001` in the acquisition quarter (483 cells, to make XIRR see an initial negative flow even when the price is 0)
and memo KPIs with hard-coded MW×years `G248 = I248 / (45 * 15)` (12 cells).

**Named ranges**: only for the macro (no names in formulas otherwise):
`M1_PF_DebtService_COPY = Model!J432:FN432`, `_PASTE = J433:FN433`, `_Delta = E434`;
`M1_PF_PaySchedule_COPY = J471:FN471`, `_PASTE = J472:FN472`, `_Delta = E472`;
`M1_DSRF_COPY = E496`, `_PASTE = E497`, `_Delta = F497`. Prefix `M1_` = Module1 macro. Plus `_xlchart.v1.*` for the waterfall.

**Circularity**: iterative calculation is **off** (BrtCalcProp fIter = 0; 100 iter / 0.0001 defaults). Circular items
(debt sculpting on CFADS/DSCR, amortisation profile, DSRF sizing) are broken by **copy/paste rows**: a "(copy)" row is
calculated, a "(paste)" row holds values (yellow input fill), a Delta cell `ROUND(copy − paste, 3)` measures convergence,
and the macro loops until all deltas are 0.

## 5. Checks and aggregation
- Stat_Inp rows 2-4 (top of the input/dashboard sheet): `C2 = M1_PF_DebtService_Delta` "Debt capacity check",
  `C3 = M1_PF_PaySchedule_Delta` "Amortisation check", `C4 = M1_DSRF_Delta` "DSRF check"; format `"CHECK";" CHECK";" OK"` + CF.
- Delta cells: `Model!E434 = F434 * I434` (only active when debt-sizing switch on), `E472 = ROUND(I473,3)` with `J473 = ABS(J471-J472)`.
- Sources & Uses check row `Model!J344 = J334 - J342`, total in I344 (not linked to the summary).
- Charts reconciliation rows `J104 = J103 - Model!J86`, `J120 = J119 - Model!J109` (economic profit re-derived and compared).
- Input check-sums `Dyn_Inp!D57 = SUM(F57:AC57)` (profile must equal 100%).
There is **no master check / error-count cell**; checks are partly aggregated (I column sums) but not rolled up.

## 6. Macros (Module1, decoded with olevba; sheet/ThisWorkbook modules empty)
```vb
Sub M1_PF_Debt_()
Application.ScreenUpdating = False
Do Until [M1_PF_DebtService_Delta] = 0 And [M1_PF_PaySchedule_Delta] = 0 And [M1_DSRF_Delta] = 0
'And [M1_Funding_Requirements_Delta] = 0          (commented out: future S&U/funding loop)
[M1_PF_DebtService_PASTE].Value = [M1_PF_DebtService_COPY].Value
[M1_PF_PaySchedule_PASTE].Value = [M1_PF_PaySchedule_COPY].Value
[M1_DSRF_PASTE].Value = [M1_DSRF_COPY].Value
Calculate
Loop
Application.ScreenUpdating = True
End Sub
```
Single copy-paste convergence loop for project-finance debt sizing; no scenario runner, no output-storage macro,
no iteration cap (risk of infinite loop). `Stat_Inp!B25 'Macro'` labels the row above the scenario columns.

## 7. Scenario / sensitivity handling
- **Scenario matrix on Stat_Inp**: each input row has slots H:Q (10 cases, names in row 25). Live value
  `F26 = IF(ISBLANK(INDEX(H26:R26, B$22)), H26, INDEX(H26:R26, B$22))` → **blank = inherit case 1 (column H)**, so only
  deviations are typed. Selector `Stat_Inp!B22` (= 3 in the file).
- **Curve selection** by name: Stat_Inp F (text, e.g. "Aurora Q1-26 Latvia (Central)") → Dyn_Inp `$D` → INDEX/MATCH into a
  library of named curves; derived curves (upside/+1%/−1%/average) are formulas on the library rows.
- **Flat sensitivity % inputs** applied multiplicatively in the model: `Stat_Inp!F100` CAPEX, `F111` PV revenue, `F116` BESS
  revenue, `F154` OPEX, `F69` yield (`Model!J191 ... * (1 + $F191)`, `J283 = SUM(...) * (1 + $E283)`).
- **Switches**: `F59` BESS refresh on/off, `F161` debt sizing case on/off, `F208` cash sweep %.
- **Results storage is manual**: per-scenario outputs pasted as values next to the live links (`Stat_Inp!H7:K19`), tornado
  deltas pasted (`F233:G241`, base 19.1% vs live `F234 = F12`), IRR bridge pasted (`C299:C303`, deltas in D), per-case
  equity cash flows pasted into `Charts!J15:FN18` (case name `C15 = INDEX(Stat_Inp!$H$25:$Q$25, B15)`).
- **2-D Excel data tables** (`TABLE(H221,H223)`, documented by the label `H242 'H221, H223'`): hold period (1.5/2/3/5 yrs)
  × buyer cost of equity (8-12%) on economic profit (`G244 = F19/1000`) and IRR (`Q244 = $F$16`).

## 8. Returns outputs and summary design
Calculated in Model rows 36-115 (before the operating detail):
- **Project (unlevered)**: FCF/CFADS row 50; `E53` Project IRR; `E54` MOIC `= -SUMIF(>0)/SUMIF(<0)`.
- **IRR trick**: per-column `J53 = IFERROR(J$41 * XIRR(J50:$FN50, J$4:$FN$4, 0), 0)` – XIRR of the flows from this
  column onward, multiplied by a one-period start flag (row 41 "Earliest flag", row 43 acquisition flag, row 90 sale
  flag); `E53 = SUM(J53:FN53)` picks the one non-zero result. This makes the IRR start date dynamic without OFFSET.
- **Equity**: funding (SHL + share capital) vs distributions (SHL interest/repayment, dividends, capital reductions) →
  row 69 total equity CF; `E72` equity IRR, `E73` MOIC.
- **Developer carry** on positive flows after cumulative payback (row 76) → FEP CF after carry (row 78) → `E81` FEP IRR
  (HTM), `E82` MOIC, `I78` net profit.
- **Economic profit** = CF + cost of capital × min(cumulative position, 0) × year fraction (rows 85-86, `I86`).
- **Selldown route**: sale flag (row 90), exit price `J93 = J90 * IF($F93, $F93, XNPV($E93, J69:$FN69, J4:$FN4) - J69)`
  (fixed price or buyer cost-of-equity NPV), CGT, carry → `G100` gross IRR, `G105` net IRR, `H103` MOIC, `I103` net
  profit, `I109` economic profit; **buyer IRR** `G113` from `J112 = J69 - J94`.
- **Debt metrics**: DSCR per period and annual (`I527/I528 = MINIFS(..., ">0")`).
- **Summary** (`Stat_Inp` rows 6-19): Project IRR/MOIC, FEP acquisition cost PV/BESS, FEP levered IRR/MOIC/net profit/
  econ profit for HTM and Sale; live column F (green links) beside pasted scenario columns H:K; units in C (`%`, `x`, `€k`).
  NPV appears only as XNPV-based exit pricing and an NPV of revenue delta on Charts (`E177 = NPV(0.2, ...)`).

## 9. Evidence count (formulas vs hard-codes)
| Sheet | Formulas | Numeric constants | Notes |
|---|---|---|---|
| Model | 52,905 | 331 | 322 = the two macro PASTE rows (433, 472 incl. I); rest = `J3` start date, `J31 = 1`, period-0 seeds (`I14`, `I15`, `I33`, `I34` …), `G114` pasted buyer IRR. |
| Charts | 3,222 | 228 | Pasted case cash flows (rows 15-18), benchmark CAPEX (Aurora F154:F158), pasted case comparisons (G83:L94). |
| Stat_Inp | 140 (+~200 array/selector formulas in F) | 214 | All yellow inputs; pasted scenario outputs and tornado. Some inputs are formulas with literals (`H96 = 1632 / F76`, `J94 = 262.07 * 1.05`, `H130 = (18432 + 36000)/(H129*1000)`). |
| Dyn_Inp | 689 | 493 | Curves (yellow). |
Literals inside Model/Charts formulas: 1000 (1,260 uses, unit conversion), calendar constants, epsilon 0.0001/0.000001
(483), MW×years `45*15` (12), `45/120` (1), NPV 20% (1). Timeline row consistency: 326/327 Model rows and 65/68 Charts rows identical across J:FN.

## 10. What the model is about
1. FEP (the investor; deal folder "Baltics/Saules Energy Dec-25") acquiring a Latvian
   ready-to-build **45 MW / 2-hour BESS** (template also carries a PV leg, currently 0 MWp) on 31-May-2026.
2. Price = development premium €45k/MW (€40k at acquisition, €5k at COD) + transaction costs €100k; construction 4 months after a 2-month development stage; BESS CAPEX ~€130k/MWh + BOP + grid.
3. Revenues from merchant BESS stacks (Aurora / Capalo / vendor curves, €k/MW p.a.), less RTM commission, distribution fees, O&M, insurance, Latvian natural-resource tax.
4. Funding: 75% shareholder loan (6.96%) / 25% share capital; optional project-finance debt (DSCR-sculpted, 3 margin terms, DSRF, cash sweep) currently off (LTV 0).
5. Two exit routes compared: hold-to-maturity (15-yr) vs selldown after 3.5 years to a buyer at a 10% cost of equity; outputs FEP levered IRR/MOIC/net & economic profit and buyer IRR.
