# Learnings (kept by Fable, the orchestrator)

Standing memory of the repository: things learned while working with the user that are useful beyond one project.
Fable adds an entry whenever a round teaches something new and useful (a user preference, a modelling convention, a
tool limitation, a trick that worked). Short entries, dated, newest at the top. Project-specific decisions stay in the
project's SPEC; rules that bind every model go to `CLAUDE.md` or `docs/MODELLING_STANDARDS.md`.

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
