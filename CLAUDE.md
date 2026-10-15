# FEP_modeller

Engine that builds FEP investment models in Excel and Investment Committee (IC) papers in PowerPoint.
Working language: English (code, workbooks, decks, docs). Chat replies to the user may be in Hungarian.

## Commands
- Build a project's model and deck: `python3 -m fepm build projects/<name>`
- Recalculate in LibreOffice and run the audit: `python3 -m fepm audit projects/<name>`
- Tests: `python3 -m pytest -q`

## Modelling rules (non-negotiable; full guide in `docs/MODELLING_STANDARDS.md`)
1. Every number on a calculation sheet is a formula. Hardcoded values exist only on input sheets,
   in blue font on an input fill, with a unit and a source note.
2. If a value must be stored as a number (goal seek, circularity breaker, scenario snapshot),
   a VBA macro writes it, and the macro lives in `vba/*.bas`. No manual paste-values.
3. One formula per row, copied unchanged across the whole timeline. No row with a different formula
   in some columns.
4. Flow: Inputs → Timing → Calculations → Outputs → Checks. Formulas only refer to the left or upward,
   or to earlier sheets. No circular references with iterative calculation on.
5. Keep it lean and readable, the way a human modeller would build it: only the sheets the SPEC lists, no helper
   clutter, no formulas over 2 nested functions or ~100 characters, no link-to-a-link chains (point at the source),
   short source notes instead of explanations, and no feature the decision does not need.
6. Each model has a Checks sheet that rolls up into one master check, which must equal 0.
7. Every figure in a deck is read from the recalculated workbook and carries its cell reference.
   Never retype a number into a slide.

## Standing assumptions (user decisions; apply unless a SPEC says otherwise)
- BESS capex: 100 k€/MWh on the BESS MWh (400 k€/MW at 4 h). User, 15-Oct-2026.

## Workflow
1. **Plan (main session, Fable):** write `projects/<name>/SPEC.md` before anything is built. It covers the decision,
   scenarios, model architecture (sheets, rows, timeline), inputs with sources, outputs and the deck storyline.
2. **Build (Opus):** `model-builder` builds the workbook, `deck-builder` the deck and `ic-writer` the emails and deck text, all strictly to SPEC.md (writing rules: `docs/WRITING_STANDARDS.md`).
3. **Check (always):** `model-auditor` runs the lint, the LibreOffice recalc and an independent Python
   recomputation of the headline outputs, then ties the deck out to the workbook. A failed check is a bug to fix,
   not one to hide.
4. Commit on the session branch with a short imperative message, then push.

## Layout
- `fepm/xl` workbook builder (styles, timeline, row writer), `fepm/ppt` deck builder,
  `fepm/audit` lint + recalc + tie-out.
- `vba/` macro source (.bas). `templates/` the IC deck template and the macro-enabled workbook template.
- `reference/` example models and decks that define the house standard (Project Rock, Project Alpha, Bora v4),
  plus the source documents (tender call + FAQ, Bora devex, overbuild). Read-only.
- `docs/domain/` the rules extracted from the source documents, each with citations. Read these instead of the raw PDFs/xlsx.
- `projects/<name>/` SPEC.md, inputs, build script, outputs (`out/`).
