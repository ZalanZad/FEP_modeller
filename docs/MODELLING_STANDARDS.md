# FEP Modelling Standards

The house standard for every Excel model the engine builds. It is distilled from Project Rock (PF model)
and Project Alpha (IRR model), with their weak spots removed. The detailed evidence is in `docs/reference_analysis/`.

Guiding principle: **a model a human modeller would be happy to inherit**. That means lean, readable and
fully formula-driven, with nothing in it that the decision does not need.

---

## 1. Sheet stack

| # | Sheet | Purpose |
|---|---|---|
| 1 | `Cover` | Project name, version, date, author, colour legend, and the master check status. |
| 2 | `Summary` | Headline outputs for the live case, the stored scenario results table and the checks panel. Only links, no calculations. |
| 3 | `Inputs` | Scalar inputs, one row each, with up to 8 scenario columns and a `Live` column. |
| 4 | `Curves` | Time-series inputs (prices, indices, capex profiles, production), annual or by period. Curves are picked by name. |
| 5 | `Calc` | The single calculation sheet on a uniform timeline. Use one sheet per asset only when the SPEC asks for it. |
| 6 | `Checks` | Every check row, rolled up into one `Master_Check` cell (0 = OK). |

Every sheet has a row-2 banner `=Cover!$C$3&" – <purpose>"` and ends with an `***END***` row.
Gridlines are off and unused columns are hidden. A sheet is added only when the SPEC needs it.

## 2. Column layout (Calc and Curves)

| Col | Width | Content |
|---|---|---|
| A–B | 3 | Section markers (`x` on each header row for navigation) |
| C | 40 | Label in plain English, sentence case |
| D | 10 | Unit, in grey italic: `EUR k`, `MW`, `MWh`, `%`, `x`, `date`, `flag` |
| E–F | 12 | Row constants (start/end dates, rates, pulled from Inputs as green links) or row results (IRR, min DSCR) |
| G–H | 2 | Spacer |
| I | 12 | Lifetime total `=SUM(Jr:<last>r)`, or the opening seed for a corkscrew (light grey fill) |
| J → | 11 | Timeline |

Panes are frozen at `J10`, so the labels and timeline headers stay visible.

## 3. Timeline header (Calc rows 3–9)

| Row | Item | Formula pattern |
|---|---|---|
| 3 | Period start | `J3 = I4+1`, with I4 seeded from `Inputs` (model start − 1) |
| 4 | Period end | `=EOMONTH(J3, Months_Per_Period-1)` |
| 5 | Days in period | `=J4-J3+1` |
| 6 | Year fraction | `=J5/Days_In_Year` |
| 7 | Year | `=YEAR(J4)` |
| 8 | Period counter | `=I8+1` |
| 9 | Spare | |

Periodicity is an input (quarterly by default, annual for simple bid models). Never use a second timeline.

## 4. Inputs sheet and scenario selection

- Column layout: B label, C unit, D source/note, F `Live`, H:O scenario 1–8, P spare.
- Formula in `Live`: `=IF(ISBLANK(INDEX($H r:$O r,1,Scenario_Selected)),$H r,INDEX($H r:$O r,1,Scenario_Selected))`.
  A blank scenario cell inherits the base case (column H), so a scenario shows only what differs from base.
- `Scenario_Selected` is a single input cell at the top, validated against the list of scenario names.
- Calc pulls only from `Live`. Never pull from a scenario column directly.
- Curves are selected by a text name on Inputs via `INDEX/MATCH` into named blocks on `Curves`.
  Leave two spare curve rows in each block.

## 5. Formatting

Font: **Aptos Narrow 10** throughout. Number formats right-aligned.

| Element | Font | Fill |
|---|---|---|
| Hard input | blue `0000FF` | light yellow `FFF2CC`, thin border |
| Link from another sheet | green `00B050` | none |
| Calculation | black `000000` | none |
| Macro-written value (PASTE rows, stored scenario results) | purple `7030A0` | light grey `EDEDED` |
| Section bar | white bold | FEP dark green `00563C` (full row) |
| Sub-header | bold `3C3C3B` | `E6F4EC` |
| Totals / key results | bold | top border |
| Check cells | `"OK"` green / `"CHECK"` red via conditional formatting | none |

Number formats:
- Money: `#,##0;(#,##0);"–"`; money with 1 dp `#,##0.0;(#,##0.0);"–"`
- Percent: `0.0%;(0.0%);"–"`; multiple: `0.00"x"`; date: `dd-mmm-yy`; flag: `0;-0;"–"`
- Units: calculations in **EUR k** unless the SPEC says otherwise. Conversions use the named constants `Thousand` and `Million`.

## 6. Formula rules

1. **No hardcodes outside Inputs/Curves.** Literals allowed in formulas: `0`, `1`, `-1`, and booleans. Everything
   else is a named constant (`Thousand`, `Million`, `Days_In_Year`, `Months_In_Year`, `Months_Per_Period`,
   `Tolerance`), defined once on `Inputs`.
2. **One formula per row**, copied unchanged across the whole timeline. Column E/F/I formulas may differ from the timeline.
3. **Short formulas.** At most ~2 nested functions; split longer logic into named intermediate rows.
4. **Direction:** refer upward, leftward or to earlier sheets. Use `$E`/`$F` for row constants and `J$` for header rows.
5. **Flags are multiplied, not wrapped in IF:**
   - Window share: `=MAX(0,(MIN($F r,J$4)-MAX($E r,J$3)+1)/J$5)`
   - Event: `=N(AND(J$3<=$E r,J$4>=$E r))`
   - Cumulative: `=MIN(1,SUM($J r:J r))`
6. **Corkscrews** for every balance: opening `=I closing` (the previous column), movements, then closing `=SUM(opening:last movement)`.
7. **Waterfalls** pay in order with `=MIN(cash available, amount due)`, then show "cash available after X" as a subtotal.
8. **Indexation:** index row `=IF(J$7=Base_Year,1,I r*(1+rate))`; nominal = real × index.
9. **Signs:** inflows positive, outflows negative, consistently from Calc to Summary. Labels never carry the sign.
10. **Returns:**
    - IRR: `=XIRR(cash flow row, J$4:<last>$4)` on a clean row that starts at the investment date. Never nudge with
      epsilon constants. If a flow may lack a sign change, wrap the result in `IFERROR(...,"n/a")` and flag it in Checks.
    - MOIC: `=SUMIF(row,">0")/-SUMIF(row,"<0")`. NPV: `=XNPV(rate, row, dates)` with the rate taken from Inputs.
11. **No** OFFSET, INDIRECT, data tables, iterative calculation, external links or merged cells.

## 7. Values that must be stored → macros

If a value cannot be a live formula (circularity, goal seek, a scenario snapshot), **a macro writes it**.

- **Circularity** (interest during construction, debt sculpting, DSRA, tax). Use a COPY / PASTE / Delta row
  triplet with the same names as Rock (`<Topic>_Copy`, `<Topic>_Paste`, `<Topic>_Delta`):
  - The macro loops paste → `Calculate` until every Delta is 0, **capped at 100 iterations**.
  - Each Delta feeds into Checks.
- **Goal seek** (for example a bid price that hits a target IRR): a macro solves it and writes the result into a
  purple input cell. A check row verifies `ABS(achieved − target) < Tolerance`.
- **Scenario runner**: `RunAllScenarios` loops `Scenario_Selected` over 1–8, recalculates (and runs the solvers),
  and writes the headline outputs into the stored results table on `Summary`. Afterwards it restores the original
  scenario. A check verifies that the live case equals its stored column.
- Macro source of truth: `vba/*.bas`. Macros refer to **named ranges only**, never to A1 addresses.

## 8. Checks

- Every check is a row on `Checks` with one test per period, summed in column F. Value 0 = OK.
- Standard checks: sources = uses, cash balance ≥ 0, balance-sheet balance (if a BS exists), every macro Delta,
  goal-seek tolerance, stored scenario = live, and debt fully repaid.
- `Master_Check = SUM(ABS(...))` of all check totals. Show it on `Cover` and `Summary` as OK/CHECK.
- Never delete or loosen a check to get to green.

## 9. Summary

- Panel 1: live case headline outputs (investment, revenue, EBITDA, Project IRR, Equity IRR, MOIC, NPV, profit,
  and the bid metrics the SPEC defines).
- Panel 2: stored scenario results (macro output, purple): one column per scenario, the same rows as Panel 1.
- Panel 3: checks status.
- Use EUR m for presentation (`=Calc!I r/Thousand`). The deck binds to these cells.

## 10. What we deliberately do NOT copy from the references

- Epsilon nudges in cash-flow rows to force a sign change for XIRR (Rock, Alpha).
- Literals such as `/1000`, `/4`, `*50%` inside formulas (Rock).
- Pasted scenario results with no check, data tables, and charts built on pasted values (Alpha).
- Macros without an iteration cap (Alpha, Rock), and broken names or external links (Rock).
- Links styled as inputs, and ad-hoc red/green font colours.
