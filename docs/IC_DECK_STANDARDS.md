# FEP IC Deck Standards

Distilled from `reference/Bora_IC_RtB_strategies.pptx`. The full design dump is in `docs/reference_analysis/ic_deck_standards.md`,
with sample slides `slide1/6/8.png`.

## Storyline (IC paper)
1. Cover
2. **Executive summary and recommendation**: the decision asked, the recommended option, 3–5 key numbers, and the approval request.
3. Project overview
4. Context (for example tender rules and scoring)
5. Options / scenarios: what differs between them
6. Economics per option: returns table + one chart
7. Bridge or sensitivity: what drives the difference
8. **Risks and mitigants**
9. Summary comparison and **next steps / approval requested**
10. Appendix: assumptions table, with every input stated alongside its source

Every content slide shows the section navigator on the left, with the current section highlighted.

## Slide pattern
- Action title: the conclusion, not the topic.
- Bold subtitle of one sentence (18 pt) giving the "so what". Its numbers are generated from the model.
- Main evidence: a table (13 pt) and/or one editable chart. Optional grey side card for messages or caveats.
- Footnote (11 pt, grey `919191`): assumptions, then `Source: <model file>, <date>.`
- Page number bottom right. No other footer.

## Design system
- Slide size 21.99 × 12.37 in (taken from the template). Font: Aptos everywhere; titles use Aptos Black 33 pt (from the layout).
- Colours: text `3C3C3B`, FEP green `00A651`, dark green `00563C`, lime `8CC63F`, grey `919191`, label fill `F2F2F2`,
  key-row fill `E6F4EC`.
- Tables: green header with white bold text, grey label column, grey unit column, numbers right-aligned. Key rows are bold on `E6F4EC`.
- Charts: editable python-pptx charts with no value axis and no gridlines, data labels on, legend at the bottom,
  series in dark green / FEP green / lime / grey.

## Number formats
- EUR m to 1 dp with negatives in brackets: `(5.3)`. Signed deltas: `+3.3`.
- Percent to 1 dp with a true minus sign: `−7.5%`. Multiple: `1.31x`. EUR k/MW with a thousands separator: `2,450`.
- `–` for not applicable.

## Engine rules
- The template is `templates/fep_ic_template.pptx` (the reference deck with its slides stripped). Pick layouts by master + name.
  Clone the page-number placeholder from the layout XML. The cover images come from the template's cover slide.
- Every number is bound to `Sheet!Cell` on the recalculated workbook. The deck build writes a binding manifest
  (`out/deck_bindings.csv`: slide, shape, text, cell, value) for the auditor tie-out.
- Language: English (en-GB). Keep Hungarian proper nouns and tender codes as they are.
