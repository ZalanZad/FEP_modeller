# FEP_modeller (`fepm`)

Builds formula-driven Excel investment models to `docs/MODELLING_STANDARDS.md` and IC decks to
`docs/IC_DECK_STANDARDS.md`. Every deck number is read from the recalculated workbook.

## Setup
Python 3.11+, `pip install -r requirements.txt`, LibreOffice 24.2+ with the Python UNO bridge (`python3-uno`).

## Use
```
python3 -m fepm build projects/_demo     # runs projects/_demo/build.py → projects/_demo/out/
python3 -m fepm audit projects/_demo     # LibreOffice recalc + lint + headline outputs + master check
python3 -m fepm.audit.lint model.xlsx [recalculated.xlsx]
python3 -m pytest -q
```

## Layout
- `fepm/xl/styles.py` named styles and number formats; `fepm/xl/builder.py` the `Model` API
  (Inputs / Curves / Calc / Checks / Summary). Calc rows are formula templates (`{key}`, `{key@prev}`,
  `{key.E}`, `{key.all}`, `{key.todate}`, header keys `start end days yf year period`) rendered once per
  column, so one formula per row holds by construction. See the module docstring.
  Two layouts: `layout="standard"` (Cover | Summary | Inputs | Curves | Calc | Checks) and
  `layout="dashboard"` (Dashboard | Inputs | Curves | Calc_1..Calc_n | Checks). `calc_sheets=["Calc_1", "Calc_2"]`
  writes one set of row templates to identical sheets with the bid number in `A1`
  (`m.calc.lever("Dash_X")` = `INDEX(Dash_X,1,$A$1)`, `m.calc.ref(key, "E", sheet="Calc_2")`); checks with row
  tokens get one row per calc sheet. `m.dashboard`: `section`, `subheader`, `input` (`Dash_<key>`), `link`
  (`out=` → `Out_<key>_<bid>`), `status`, `stored_row` / `stored_block` (`<name>_Stored`), `scenario_results`,
  `panel_section`, `panel_link`, `selector` (`Scenario_Name`).
- Lint rules beyond the standards: `g-chain` (a link to a link, or a row that only copies another row),
  `g-long` (over 100 characters or 2 nested function levels), `g-note` (source notes over 60 characters),
  `c-sheet` (Calc_n formula text differs from Calc_1), `a-bid` (Calc_n!A1 is not n).
- `fepm/recalc.py` headless LibreOffice over UNO; `fepm/audit/lint.py` static + post-recalc lint.
- `fepm/ppt/deck.py` deck builder on `templates/fep_ic_template.pptx` (made by `make_template()` from the
  reference deck: slides 2–8 dropped, cover slide kept with its images and neutral text).
  `Deck.bind(sheet, cell, fmt)` / `Deck.stored(output, scenario, fmt)`; `save()` writes `deck_bindings.csv`.
- `vba/` macro source: `Solver.bas` (COPY→PASTE, 100-iteration cap), `GoalSeek.bas` (secant on
  `<Topic>_Input/_Achieved/_Target`), `Scenarios.bas` (`RunAllScenarios`). Named ranges only.

## Recalculation
`fepm.recalc.recalc()` starts `soffice --headless` on a private profile, loads the workbook over UNO,
runs `calculateAll()` and stores it with the "Calc MS Excel 2007 XML" filter: formulas are kept, cached values
are LibreOffice's (XIRR/XNPV included, tested against Python to 1e-9). Without `uno` it falls back to
`soffice --convert-to xlsx` (recalculates, but cannot set inputs or run macros).

## Macros
`fepm.xl.macros.package(xlsx, xlsm, run=...)` embeds every `vba/*.bas` with LibreOffice and runs the macros in
order headlessly (`run` = `("Scenarios", "RunAllScenarios")` by default, or a list such as
`["BoraImpact.RunCommitmentImpact", "Scenarios.RunAllScenarios"]`):
the modules go into the document's Basic library in VBA-compatibility mode (with `ModuleInfo`, project name set),
and LibreOffice's experimental VBA exporter (`/org.openoffice.Office.Calc/Filter/Import/VBA/UseExport`) writes a
real `xl/vbaProject.bin` (PROJECT, VBA/dir, VBA/_VBA_PROJECT, one stream per module); `olevba` decodes the
source and LibreOffice reopens the file and runs the macro. Caveats: not yet opened in desktop Excel here,
there are no ThisWorkbook/Sheet document modules, and LibreOffice silently stops a macro on a runtime error
(the Checks sheet is the safety net). If Excel rejects the project, build the `.xlsx` and run
`tools/inject_vba.ps1 -Workbook <model>.xlsx -Run` on Windows (needs "Trust access to the VBA project object model").
