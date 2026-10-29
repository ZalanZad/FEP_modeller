# Learnings (kept by Fable, the orchestrator)

Standing memory of the repository: things learned while working with the user that are useful beyond one project.
Fable adds an entry whenever a round teaches something new and useful (a user preference, a modelling convention, a
tool limitation, a trick that worked). Short entries, dated, newest at the top. Project-specific decisions stay in the
project's SPEC; rules that bind every model go to `CLAUDE.md` or `docs/MODELLING_STANDARDS.md`.

## 2026-10-29 (flow-through test)
- **Run a one-at-a-time shock test before delivery** (`projects/hu_wind_2/flow_test.py` is the template): set each
  key input in LibreOffice, recalc, read the headline outputs, count error cells with
  `sheet.queryFormulaCells(4)`, check the Master check, restore. It finds what static audits miss: formulas that are
  right for the base dates but break when a date crosses a year boundary, and timelines too short for a COD delay.
- **Date-driven accruals must survive any ordering of dates:** use MAX(0, …) on day counts and exponents; a guard
  check is a fallback, not a fix.
- **Give the timeline headroom** beyond COD + life (≥ 4 years) so a COD delay is a sensitivity, not a check failure.
- **Insert columns inside a range through UNO** (before the last column) to extend every $J:$AS reference
  automatically, then copy the last column's formulas into the new ones.

## 2026-10-29
- **JV waterfall, devex:** only the *return* on FEP's devex funding (MAX(20 % accrued, 1× principal)) is paid before
  the 70/30 split; the principal is covered by the RtB premium through FEP's share. Do not add a principal row
  (user, 29-Oct-2026 – reverted an audit "fix").
- **Hold-case tax and debt mechanics = the Bora model (HTM sheet):** depreciable base = construction capex +
  connection fees + municipal (+ capitalised IDC), not land or the RtB premium; straight line 20 y from COD pro rata
  to the operating share; debt sized on the same base, drawn pro rata with construction, IDC capitalised, 1.75 %
  upfront fee equity-paid, annuity from the first full operating year; DSCR on post-tax CFADS (EBITDA + CIT).
  The Bora model also has loss carry-forward (50 % cap), DSRA, sweep, commitment fee, Robin Hood tax and a
  refinancing case – switched off or case-specific, so the lean model leaves them out and says so in the SPEC.
- **COD price = PV of unlevered post-tax cash flow** at the buyer's unlevered return (asset-pricing convention);
  never price off a levered equity IRR.
- **Signing fee + DD = 135 k€ per project** (user, 29-Oct-2026); never label them "exit fees".
- **Reviewer findings are hypotheses:** a "missing principal" finding was wrong because the deal logic (premium covers
  it) was not in the dump. Put deal-structure notes in the SPEC where the reviewer reads them, and confirm structural
  "fixes" with the user before building them.

## 2026-10-28 (audit round)
- **Audit pairs pay off:** run `model-auditor` (numbers, lint, scenarios) and `model-reviewer` (business sense, labels,
  structure) in parallel on a finished file; the reviewer found a waterfall omission (loan principal not repaid before
  the split) that the number tie-out could not, because the Python shadow copied the same logic.
- **Checks must be falsifiable.** A check that compares a residual with itself is always 0; tie the dated cash-flow
  list to the profit stack instead, and count errors in the summary block (`SUMPRODUCT(--ISERROR(range))`).
- **Prefer plain `XIRR` on the whole timeline row** (zeros are fine) over `XIRR(FILTER(...))`: FILTER is Excel-365
  only, shows #NAME? in LibreOffice and older Excel, and hides errors from the Master check.
- **Shareholder-loan waterfall order:** sale costs → principal of every FEP loan line (securities, fees, land, devex)
  → interest/returns → DSA → split. Each line's principal and return are separate rows, so the 2× floor on devex is
  principal + MAX(accrued, 1× principal).
- **Outflows negative everywhere**, debt service included; DSCR = EBITDA ÷ debt service. Do not bake parameter
  values ("15 %", "60 %") into labels; they go stale.
- **The user's dates are common to all projects** (their decision), so a portfolio XIRR in the Total column is valid;
  say so in the row note.

## 2026-10-28
- **The user's hand-edited copy is the base.** When the user uploads an edited workbook, copy it to `reference/` with
  the date, diff it against our last output to learn the corrections, record them in the SPEC, and build the next
  change on the user's copy. Never rebuild from our earlier output.
- **Edit user workbooks through LibreOffice UNO** (`Rows.insertByIndex`, formulas set per row range), not openpyxl
  inserts: references, names and formats survive; then compare every pre-existing cell before/after.
- **LibreOffice 24.2 cannot evaluate FILTER-based array IRRs** (`_xlfn._xlws.FILTER`): they show #NAME? in the recalc
  and must be confirmed in Excel; verify them with XIRR on the same rows in Python instead. LibreOffice also drops
  Excel dynamic-array metadata and lower-cases `_xlfn._xlws.FILTER(`; `build_v3.py` restores both.
- **Shareholder-loan tax logic (HU wind, JV projects):** FEP's "claims" (12 % on cash lines, 20 % with a 2× floor on
  land/devex) are interest on shareholder loans to the JV; taxable at FEP every year when capitalised (31-Dec), the
  rest at the sale. Share gains stay exempt (participation exemption). A project without a JV partner has no loans,
  no claims and no such tax.
- **Hold-case tax and debt pattern:** CIT 15 % on EBITDA − 20-year straight-line tax depreciation, no loss
  carry-forward; debt 60 % of the all-in cost at COD, 5 %, 15-year annuity from the first full operating year,
  interest deductible in the levered CIT; outputs: CIT over the life, project IRR post-tax, equity IRR, min DSCR,
  "debt repaid" check.
- **Streamlined means:** one Dashboard with key inputs on top, a case selector (Base/Low/High), summary before detail,
  detail blocks in the same column layout (E/F/G = projects, J:AS = years), labels that say what the number is in a
  few words, nothing hidden, no feature the decision does not need. "EBITDA", not "NOI".
- **Tender arithmetic that keeps coming back:** a bid of 50 MW is not a 50 % overbuild; the max-points ÖSZ-3 band
  is > 50 % (use 51 %); build tolerance 10 % below the bid without penalty; bid security 0.5 % credited into the
  performance bond; TSO security credited against the indirect fee.
