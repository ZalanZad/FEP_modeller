# IC deck standards – reverse-engineered from `Bora_IC_RtB_strategies.pptx`

Source: `scratchpad/ref/Bora_IC_RtB_strategies.pptx` (8 slides, 11 Oct 2026).
Tools: `scratchpad/tools/dump_pptx.py` (structure dump), `strip_test.py` (template reuse test),
`find_vals.py` / `print_block.py` (Excel cross-check). Full dump: `scratchpad/render/dump.txt`.
Rendered PNGs here: `slide1.png` (cover), `slide6.png` (table + stacked bridge chart), `slide8.png` (summary).

Note: rendering used LibreOffice, which lacks Aptos / Aptos Black, so the titles look lighter in the PNGs than in PowerPoint.

---

## 1. Slide-by-slide

| # | Title (placeholder) | Subtitle (bold 18 pt text box) | Content | Key numbers / messages |
|---|---|---|---|---|
| 1 | Cover: "Project Bora / RtB exit and tender bid strategies / 11 October 2026" | – | Full-bleed photo (wind farm), white arc swoosh PNG on the right, FEP logo (white) bottom-left; one text box | Project name 82.5 pt bold, topic 44 pt bold, date 32 pt; all white Aptos |
| 2 | Bora overview | "200 MVA onshore wind with battery storage – technical data and deal terms" | Two 2-column key/value tables side by side (Item/Assumption; Transaction structure/Terms) with section sub-header rows | 200 MVA grid, 261 MW wind (+30.5% overbuild), 90 MW × 4 h BESS, RtB 30-Jun-29 / sale 31-Mar-30, NTP 30-Sep-29 / COD 30-Sep-31; JV FEP 70% / A.RES 30%; DSA 2.5 k€/MW p.a. cap 10 k€/MW; signing fee €0.25m; devex €3.14m; RtB premium 150 k€/MW × 200 MVA = €30.0m; COD price 2,450 / 2,100 k€/MW |
| 3 | Tender scoring – what can be offered and what it scores | Offerable range, band thresholds, weighted points | Large matrix table 17×14 (criterion / offerable / max / up to 11 band columns, alternating "threshold" and "points" rows) | ÖSZ-1…ÖSZ-9b criteria, 100-point total |
| 4 | Tender strategies | Commitment and points per criterion for the five bids | Table 13×13: 2-row header (strategy name merged over Offer/Pts pair) | Scores EU / CN: S0 0/0, S1 36/24, S2 51.5/36.5, S3 69.5/54.5, S4 90/75; installed MW / clipping per strategy |
| 5 | Cost per tender point at the RtB exit | RtB price given up per point added to the current bid | Table 11×6 (two section-header rows) + grey "callout card" on the right ("Not priced in the RtB price", bullets) | k€ per point ranges, e.g. municipal 278–417 RtB k€/pt; EU turbine binds buyer to €117.5m more capex |
| 6 | RtB price bridge – vs the current bid | Gross premium €30.0m ± PV of promises | Table 10×7 (metric × 5 strategies, unit column) + stacked column chart | Effective RtB price €58.9m / 30.0 / 29.9 / 24.3 / 5.8; price effect +28.9 / 0 / (0.1) / (5.7) / (24.2) |
| 7 | RtB returns per strategy | FEP returns at the RtB sale (31-Mar-30) with A.RES share and DSA | Table 16×7 (KPI rows highlighted, sensitivity section) + clustered column chart (IRR by premium 150/200/250 k€/MW) | FEP net profit €37.6m / 15.2 / 14.7 / 10.0 / (5.3); IRR 50.2% / 20.9% / 20.0% / 13.7% / −7.5%; MOIC 1.82x … 0.90x; NPV@8%; peak exposure; break-even premium 150.0…301.4 k€/MW |
| 8 | Summary – RtB exit against building | FEP IRR / MOIC per strategy: RtB vs sale at COD vs hold to maturity | Route × strategy table (cells "IRR / MOIC"), "Key messages" bullets (bold lead-in + text), grey card with a small reference table + caveat | RtB 20.9% / 1.31x (S1); COD EU 31.4% / 1.54x; COD CN 42.1% / 1.80x; HTM EU 10.0% / 6.13x; HTM CN 13.1% / 7.79x; node-cap flag; "no BESS revenue" caveat |

**Storyline** (left-hand navigation table on every content slide lists exactly these 7 sections, active one highlighted):
Overview (project + deal terms) → Tender scoring rules → Strategies (options) → Cost per point (trade-off) → Price bridge (value) → Returns (FEP IRR/MOIC/NPV + sensitivity) → Summary (routes comparison, key messages, caveats).

Observations for a generic IC template:
- There is **no** separate exec summary / recommendation slide up front, **no** risks, **no** next steps / approval-request slide, no appendix, no agenda slide (the left nav is the agenda). The "Summary" at the end is the decision slide. A full IC paper engine should add: Exec summary & recommendation (front), Risks & mitigants, Next steps / IC approval request, Appendix (assumptions, integrity checks) – styled with the same components.
- Every content slide = Title + bold subtitle (one-sentence "so-what") + main table (+ chart or side card) + grey footnote ending with "Source: Bora IC model, 11 Oct 2026."
- Slide-number placeholder text in the file is stale ("4", "7") but it is a real `slidenum` field, so PowerPoint shows 2…8 correctly.
- No speaker notes, no hidden slides.

## 2. Design system

**Slide size**: 21.99 × 12.37 in (≈ 55.9 × 31.4 cm) – a 16:9 "large" canvas (not the default 13.33 × 7.5). Font sizes are correspondingly large (body text 13 pt reads small).

**Masters / layouts** (two masters):
- Master 0 (`slideMaster1`, theme1 "Office" colours – black tx1): layouts `Vertical Title and Text` (only slide-number ph; used by the **cover**) and `Base text` (**used by all content slides 2–8**).
- Master 1 (`slideMaster2`, theme2 colour scheme **"FEP"**; has a thin horizontal rule freeform at y = 1.55 in from x = 4.92 in): layouts `Custom Layout`, `Base text`, `1_Base text` (two-column with big swoosh image), `Image + text 1`, `1_Image + text 1`, `text + tables`, `text + tables + image`, `Large image`, `Gallery` (6 pictures + captions), `Vertical Title and Text`. Unused by this deck, but available.
- Every non-cover layout carries the FEP logo (`Picture 5`, image1.png, at x 18.99 in, y 0, 2.2 × 1.51 in → top right) and an empty `TextBox 12`.
- Content layout placeholders (`Base text`): title idx 0 at (4.92, 0.31) 3.76 × 0.87 in (slides widen it to 13.8 in); body idx 10 at (4.92, 3.74) 11.0 × 3.02 in (unused – deck uses free text boxes/tables); slide number idx 4 at (15.53, 11.47) 4.95 × 0.66 in, right-aligned 12 pt.

**Grid**: left nav column 0 → 4.12 in; content column x = 4.92 in, width 15.0 in (right edge 19.92 in). Title y 0.31, subtitle y 1.20 (h 0.45), body starts y 1.75–1.95, footnote at y ≈ 10.55–10.85 (14.4 in wide), page number y 11.47.

**Fonts**: Aptos everywhere (theme major "Aptos Display", minor "Aptos"); title placeholder style = **Aptos Black 33 pt, not bold, colour tx1**. Body sizes:
| Element | Size | Weight | Colour |
|---|---|---|---|
| Cover project name / topic / date | 82.5 / 44 / 32 pt | bold / bold / regular | FFFFFF |
| Slide title | 33 pt Aptos Black (layout-inherited) | – | tx1 |
| Subtitle | 18 pt | bold | 3C3C3B |
| Left nav items | 18 pt | regular; **bold + white** when active | 3C3C3B |
| Table text (all tables) | 13 pt | header bold; KPI rows bold | 3C3C3B; units 919191 |
| Card header ("Key messages", card titles) | 16 pt | bold | 00563C |
| Card / bullet text | 13 pt | bold lead-in phrase, then regular | 3C3C3B |
| Footnote + source | 11 pt | regular | 919191 |
| Chart title / labels / legend | 13 / 12 pt | title bold | 3C3C3B |
| Page number | 12 pt | – | tx1 tint 82% |

**Colour palette** (theme2 "FEP" + direct RGB in slides):
| Role | Hex |
|---|---|
| Primary text (dk1) | `3C3C3B` |
| FEP dark green (accent1) – card headers, chart main series, nav gradient end | `00563C` |
| FEP green (accent2) – table header fill, section bars, nav gradient start | `00A651` |
| Sage (accent3) | `799385` |
| Light green (accent4) | `82CA9C` |
| Mid grey (accent5) – units, footnotes | `919191` |
| Light grey (accent6) | `CACACA` |
| Lime (chart) | `8CC63F` |
| Row-label / card fill | `F2F2F2` |
| KPI highlight / section-header row fill | `E6F4EC` |
| Chart light greens | `B7DEC6`, `CFE8D8` |
| Chart greys | `A6A6A6`, `BFBFBF` |
| White | `FFFFFF` |

**Left navigation**: a 7×1 table at (0, 0), 4.12 × 5.95 in, 18 pt, left margin 180000 EMU, 0.5 pt (6350 EMU) top/bottom rules in 3C3C3B, no fill; active row = linear gradient fill `00A651` → `00563C`, bold white text.

**Logo**: layout picture top-right (content slides); white logo bottom-left on cover (image5.png). Cover also uses image3.jpeg (1920×1080 photo) and image4.png (arc swoosh).

**Footers / confidentiality**: none – no "Strictly confidential" mark, no date/footer placeholders; only the page number bottom right and the per-slide "Source:" line.

**Tables** (all built as direct-formatted tables; tableStyleId = Medium Style 2 – Accent 1 but every cell overrides fill/fonts; white 1 pt cell borders come from the style):
- Header row: fill `00A651`, white bold 13 pt; first columns left-aligned, numeric columns right-aligned; header row height ≈ 0.59 in (536792 EMU), body rows ≈ 0.34 in (311685 EMU).
- Cell margins: L/R 73152 EMU (0.08 in), T/B 27432 EMU (0.03 in), anchor middle.
- Label column and unit column: fill `F2F2F2`; unit text `919191`. Data cells: no fill (white).
- Section header rows inside a table: fill `E6F4EC`, bold dark-green text (slides 2, 5); or a full green `00A651` bar with white bold text (slide 7 "RtB premium sensitivity").
- KPI / "answer" rows (Effective RtB price, FEP net profit, IRR, MOIC, RtB sale route): bold, data cells filled `E6F4EC`.
- Indented sub-rows use leading spaces ("   to FEP (70%)").
- Column widths (returns table): label 4.0 in, unit 0.9 in, 5 × 2.02 in data columns.

**Number formats** (pre-formatted strings in tables):
- Money: €m, 1 decimal, thousands separator, negatives in parentheses `(5.3)`; deltas with explicit sign `+3.3`; zero `0.0`; not applicable `–` (en dash).
- Percent: 1 decimal `20.9%`; negative IRR uses a true minus sign `−7.5%` (U+2212), not parentheses.
- Multiples: 2 decimals + `x` → `1.31x`.
- k€/MW, k€ per point: 0–1 decimal with thousands separator `2,450`, `3,557`, ranges with en dash `278–417`.
- Pairs: `IRR / MOIC` cells "20.9% / 1.31x", score pairs "36.0 / 24.0".
- Dates: `30-Jun-29`, `31-Mar-30` in tables; "11 October 2026" on cover; "11 Oct 2026" in source line.
- Units in a separate "Unit" column (€m, %, x, k€, k€/MW, pts).

**Charts** (native, editable charts with embedded workbook `Sheet1` – python-pptx `CategoryChartData` style):
- Slide 6: `COLUMN_STACKED`, gap width 55, overlap 100, series colours 00563C / 00A651 / 8CC63F / A6A6A6 / CFE8D8 / B7DEC6, data labels centred (format `#,##0.0` and `+#,##0.0;(#,##0.0)`), value axis hidden, no gridlines, legend bottom, title 13 pt bold inside chart, segments < €1m label-suppressed.
- Slide 7: `COLUMN_CLUSTERED`, gap 80, overlap −10, colours A6A6A6 / 8CC63F / 00563C (light → dark = low → high premium), labels outside end `0.0%`, value axis hidden, no gridlines, legend bottom.
- Chart sits below the table across the full 15 in width (h 2.6–4.4 in).

**Callout card**: rounded/plain rectangle fill `F2F2F2`, no line; header 16 pt bold `00563C`; bullets 13 pt with bold lead-in.

## 3. Numbers that come from the Excel model

All figures are sourced from `Bora_IC_model_v4.xlsx` (verified: values in the deck charts match cell values to full precision). The model stores money in **k€**; the deck shows **€m, 1 dp** (÷1000). Main source block: sheet **`Control Panel`**:
| Deck item | Cells (cols E–I = S0…S4, J = Δ S0 vs S1) | Model unit → deck format |
|---|---|---|
| Tender score EU / CN | C153 / C154 | pts → 0.0 |
| FEP net profit per route | rows 156–160 | k€ → €m 0.0 |
| FEP IRR per route | rows 162–166 | fraction → 0.0% |
| FEP MOIC per route | rows 168–172 | x → 0.00x |
| FEP NPV per route | rows 174–178 | k€ → €m |
| RtB bridge: gross premium, PV of promises, price effect, effective price, per MVA | rows 181–185 | k€ → €m; k€/MVA |
| Sale at COD price / per MW / price effect / buyer IRR / implied prices (EU / CN) | rows 192–210 | k€, k€/MW |
| HTM IRR / MOIC / NPV / min DSCR | rows 214–222 | – |
| Integrity status | E145–E150, rows 226–231 ("25 of 25") | text |
| Tender scoring bands (slide 3) | rows 235–300 | thresholds / points |
| Live single-route KPIs (net profit, IRR, MOIC, NPV, NPV/Equity, peak exposure) | E139–E144 | k€ / % / x |
| 25-run data table (scenario × route) | rows 305–330 | k€ / % |
Also sheet **`Tender`** rows 225–240 (per run: score, capex, debt, proceeds, peak exposure, LCOE, min DSCR) and the input rows on `Control Panel` (e.g. E130 discount 8%, E133/E134 COD price 2,450/2,100 k€/MW).

Generated (should be linked): every table number on slides 2–8, both chart series, the numbers embedded in subtitles, footnotes, key messages and card text (e.g. "€30.0m", "4.7%", "€117.5m", "190.6 k€/MW", "−17.6 k€/MW") – the narrative text is templated around model values. Static: criterion names, strategy names, deal-term wording, scoring-rule text.

Gaps vs WindHun: the WindHun workbook (`HUN_Wind_MOIC_HTM_IC.xlsm`) uses the 4-project × 3-case structure (Control_Panel), not the 5 tender strategies; the engine needs a mapping layer from WindHun output cells to these slide components.

## 4. Reuse as a python-pptx template

**Yes – copy the file, delete all slides, reuse the masters.** Tested (`tools/strip_test.py`): removing every `sldId` + dropping rels, then `add_slide(master0 'Base text')` renders correctly with logo, title, bullet body. Caveats:
1. **Pick layouts by (master, name)** – names are duplicated across masters (`Base text`, `Vertical Title and Text` exist in both). Content slides use master 0's `Base text`; the FEP colour theme is only on master 1 (master 0 has the Office theme, so `tx1` = black there – the deck hard-codes `3C3C3B` instead). Engine should set explicit RGB colours, as the deck does.
2. **Slide-number placeholder is not cloned by python-pptx** `add_slide` (it skips footer/date/sldNum placeholders). Copy the layout's `sldNum` `<p:sp>` into each new slide (with `<a:fld type="slidenum">`) or add a text box.
3. **Cover assets are on the slide, not the layout** (photo image3.jpeg, swoosh image4.png, white logo image5.png); deleting slide 1 drops them. Either keep slide 1 and edit its text box, or extract the three images and re-add them.
4. Title placeholder width in layout is only 3.76 in – widen to 13.8 in per slide (deck does this).
5. Body placeholder idx 10 is positioned mid-slide (y 3.74) and is unused; the deck draws tables/text boxes at absolute positions instead. The left nav and subtitle/footnote are free shapes, not layout elements – generate them in code.
6. Master 1 layouts (`text + tables`, `Image + text 1`, `Large image`, `Gallery`, etc.) offer extra placeholders (picture idx 10/11/22–26, body idx 12/13/15) usable for project photos/maps.
7. Table style GUID `{5C22544A-…}` is the python-pptx default; every cell must be formatted explicitly (fill, font, margins) – write a `style_table()` helper.
8. Charts: python-pptx `add_chart` + set series fills, `gap_width`, `overlap`, data-label `number_format`, hide value axis, legend bottom – all reproducible.
9. File size: stripped template ≈ 1.6 MB (layout swoosh image2.png 5436×2988 is the bulk).

## 5. Language

Content is entirely **English** (en-GB, `lang="en-GB"`), with Hungarian proper nouns and tender codes kept (ÖSZ-1…ÖSZ-9b, Szombathely OVIT). Technical tokens RtB, COD, NTP, DSA, IRR, MOIC, NPV, MVA, A.RES unchanged. The model itself has bilingual labels in places (Deposits sheet "Item (HU)/Item (EN)"), so an HU/EN switch is feasible, but the reference IC deck is English-only.
