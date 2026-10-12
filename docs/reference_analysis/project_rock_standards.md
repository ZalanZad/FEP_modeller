# Project Rock Financial Model 2.1: reverse-engineered modelling standards

Source: `ref/Project_Rock_-_Financial_Model_2.1.xlsm`. This is a Futureal/FEP project-finance model for a 45 MW Polish wind farm acquisition (Jarszewo), written in PLN. It was read with openpyxl twice, once for formulas and once for cached values, and the raw XML was inspected directly. VBA was extracted with `olevba` and none of it was executed. Helper scripts are in `scratchpad/tools/rock_*.py` and the VBA dump is in `analysis/vba_dump.txt`.

Theme colour scheme is "RPC Custom 1": dk1 000000, lt1 FFFFFF, dk2 44546A, lt2 E7E6E6, accent1 568AB5, accent2 F76721, accent3 A5A5A5, accent4 FFC000, accent5 00AF64, accent6 3E6882. The theme font is Arial for both headings and body.
openpyxl theme indices below follow the standard order: 0 = lt1 white, 1 = dk1 black, 2 = lt2, 3 = dk2 44546A, 4 = accent1 568AB5, 5 = accent2 F76721, 6 = accent3, 7 = accent4 FFC000, 8 = accent5 00AF64, 9 = accent6.

---

## 1. Sheets: order and purpose

| # | Sheet | Tab colour | Used range | Purpose |
|---|---|---|---|---|
| 1 | **Cover** | red FF0000 | B2:C10 | Title (B2 "Project Rock", Arial 16 bold) and the **formatting legend** (B7:C10): input, input from another sheet, and "pattern = leave empty". It has no formulas. |
| 2 | **Summary** | red FF0000 | B2:M61 | The one-page dashboard. It holds deal inputs (purchase price, tranche dates, FX, refi terms) as blue-on-yellow inputs, operating CF totals in PLNm and EURm, returns, and the **check panel** (C26:D40). It also has the "RECALCULATE MODEL" button. |
| 3 | **Assumptions** | theme accent1 | A2:Z1615 (content to row ~662) | **Scenario input sheet** for scalar inputs. It has 8 scenario columns (I:P) plus an active column G. Every scalar input lives here. It carries three macro buttons in G3:G5. |
| 4 | **Time-based** | theme accent1 | B2:QV300 | **Time-series input sheet**. It holds curves by year or quarter: capex schedules, availability, degradation, inflation, FX, power prices, PPA, GoO, base rates, balancing and the debt repayment profile. Each curve block ends in a "Selection" row that picks the active curve. |
| 5 | **Land costs** | theme accent1 | A2:HD76 | Register of land lease and easement contracts (rows 23–76, one row per contract, autofilter A22:Z76). It has its own quarterly timeline from AA18 (2016) and is aggregated by SUMIFS on type (rows 21–22). This is a legacy sheet: it has a Polish header ("Inflacja"), its own inflation table and hardcoded years. |
| 6 | **1** | accent2 tint 0.4 (light orange) | A1:GH1875 (content to row 1492) | The **single calculation sheet** ("asset sheet 1"). It covers flags, valuation/IRR, the 3 financial statements, construction, operations, financing, taxes and macro copy/paste rows. The name "1" is the asset number. The design anticipates "2", "3"… sheets (see `A1` Model_Number and the note in row 269: "when adding asset sheets, schedules need to …"). |
| 7 | **Names** | none | B2:H25 | Named constants: Days_Yr, Months_Yr, Thousand, Million, Verysmallnumber, rounding tolerances, and the funding-mechanism list. |

Flow: Cover → Summary (outputs + key deal inputs) → Assumptions (scalars, scenarios) → Time-based (curves) → Land costs (register) → 1 (all calculations) → Names (constants).
The workbook has no separate checks sheet. Checks are computed inside "1" (column F), and the Summary C26:D40 panel collects them.

Every sheet has a common banner in row 2: `B2 = TEXT(Cover!B2,"#")&" - <Sheet purpose>"`, for example "Project Rock - Asset Scenarios", "… - Time-based inputs" and "… - Names and Constants". The banner is Arial 12 bold, white font (theme 0), fill dk2 **44546A**, spanning the used width.
Each sheet ends with an `***END***` row with section-header styling (Time-based!B300, '1'!B1492).

Closed-off sheets: Cover, Summary, "1" and Names use `defaultColWidth="0"` and `zeroHeight="1"`, so every column and row beyond the used area is hidden. In "1", columns FT onward are hidden.

---

## 2. Layout conventions

### 2.1 Calculation sheet "1" (the timeline sheet)
| Column | Width | Role |
|---|---|---|
| A | 3.9 | Spacer. A1 holds `Model_Number` = 1 (asset id, blue input) |
| B | 3.9 | **Section / sub-section labels**, e.g. B13 "Flags", B82 "Cashflow statement", B843 "PF / Long-term Senior Debt" |
| C | 32.3 | **Line-item label** |
| D | 8.4 | **Unit**, italic grey (font theme0 tint −0.35 = A6A6A6): `LCY'000`, `%`, `flag`, `#`, `x`, `date`, `Ratio`, `Check` |
| E | 24.3 | **Row constants / imported scalars**: start dates, rates, switches, pulled from Assumptions with `INDEX(Assumptions!$G…,MATCH($A$1,Model_Number,0))`. Mid-blue link font 568AB5 with a thin box border. Also holds summary results (MIN DSCR, IRR, MOIC). |
| F | 20.1 | Second row constant (end date) or **check result** (a "Check" caption sits in the row above, e.g. F250, F660) or a secondary result (avg DSCR, XNPV) |
| G | 2.7 | Spacer |
| H | 12.4 | Occasional reconciliation deltas (H754:H756, H1207 "IDC Delta") |
| I | 12.4 | **Lifetime total** `=SUM(J r:FR r)` on almost every flow row. It is also the "period 0" / opening column (I5 = model start date, I16 = 0) that the first timeline formula references as the previous period. |
| J:FR | 12.4–12.6 | **Timeline**: 165 quarterly columns, 01-Jan-2023 to 2064. FS is a buffer and FT onward is hidden. |

Timeline header rows 2–11, frozen at **E6** (panes xSplit=4, ySplit=5):
- Rows 2–5 are the banner band (fill 44546A, white font). A2 = asset name.
- Row 4 `Period Start`: `J4 = I5+1`.
- Row 5 `Period End`: `J5 = IF(J4>=$F$17, EOMONTH(J4,$F$5-1), EOMONTH(J4,$E$5-1))`. E5 and F5 are the construction and operations periodicity (3 months each), so the periodicity can change at COD.
- Row 6 `Quarter`: `=ROUNDUP(MONTH(J5)/3,0)` with format `\Q#`.
- Row 7 `Year`: `=YEAR(J5)`.
- Row 8 `Days in period`: `=J5-J4+1`.
- Row 9 `Fraction of year`: `=J8/J10`.
- Row 10 `Days in year`: `=DATE(J7,12,31)-DATE(J7,1,1)+1`.
- Row 11 `Hours in year`: `=J10*Hours_Day`.
- Row 12 is a note: `"LCY" = "Local Currency"`.
- Rows 13–24 hold the **Flags** section (see §4.1).

Section hierarchy in "1":
- Super-sections are bold text in B with no fill: "Valuation" (B49), "Financial statements" (B80), "Construction" (B266), "Operations" (B286), "Financing" (B616), "Taxes" (B1199).
- Sections are full-width bars: fill accent1 **568AB5**, white bold font, label in B. Examples: B51 "SPV Equity IRR %", B82 Cashflow statement, B145 Income Statement, B219 Balance sheet, B288 Generation, B337 Revenue, B404 Operating costs, B618 Funding, B843 PF debt, B1019 Refinancing loan, B1145 Shareholder Loan, B1175 Equity, B1274 CIT, B1481 Taxes summary.
- Sub-sections are plain text in B, for example "Cash waterfall", "Sources & Uses", "Coverage ratios", "Reserve accounts / facilities".
- Sub-sub labels are bold in C, for example "*Equity first" C727 and "*Uses" C665.
- Corkscrews (account blocks) follow the order **Balance b/f → movements → Balance c/f**. b/f is `=Z910` (the previous column's c/f) and c/f is `=SUM(b/f:last movement)`. Examples: senior debt rows 904–910, SHL rows 1150–1155, equity rows 1182–1186, DSRA rows 970–974.

Row grouping in "1": 129 rows at outline level 1 (e.g. 1327–1455, the tax-loss vintage matrix). Summary rows are below the detail (`summaryBelow=True`).

### 2.2 Assumptions (scalar scenario sheet)
| Column | Role |
|---|---|
| C | Sub-section label (e.g. C531 "Scenario", C538 "Long term debt financing"). Section bars sit in C, fill 568AB5 (e.g. C12 "Timing", C69 "Generation", C124 "Capex", C199 "Revenue", C529 "Project Finance long-term debt financing", C629 "Tax & Accounting") |
| D | Input label (width 41.9) |
| E | Unit, italic (`date`, `months`, `%`, `x`, `PLNk`, `PLNk / MW`, `Switch`, `Selection`, `type`) |
| F, H | Spacers (3.7) |
| G | **Active / live value** (width 15.7) |
| I:P | **Scenario 1–8 input columns** (I = "Jarszewo", J–P = "Spare") |
| Q | Spare scenario, included in the INDEX range I:Q |

Header block:
- I4:P4 holds the scenario names, fill accent2 tint 0.4 (light orange).
- I5:P5 holds the scenario numbers (`I5=1, J5=I5+1…`), fill accent5 tint 0.8 (light green).
- I6:P6 holds the descriptions.
- G7 "Active Inputs". G8 `=I4` and G9 = 1 (the `Model_Number` named range), fill **FAA47A** (orange).
- **G10 = active scenario selector**, a blue input.

No freeze panes. Row groups at level 1: 81–95 (wind build-up) and 294–315 (variable/capacity O&M).

### 2.3 Time-based (curve input sheet)
- D holds the curve label, E the unit, F the meta value (e.g. "Real prices as of" date, contract type), and G the year or pricing-year offset. The timeline starts at **H**.
- Each block carries its own date header rows. Annual blocks have a Year row (e.g. D69 `Year` H=2018, `I=H69+1`). Quarterly blocks have Period start, Period end, YEAR and MONTH rows.
- Spare rows are reserved and labelled "Spare" so lists can grow.
- Each block ends in a **Selection row**, labelled bold `1 | Selection` in C:D. E links the chosen curve name from Assumptions (e.g. `E180 = Assumptions!$G$204`), G links a sensitivity (`=Assumptions!$G$205`), and the timeline cells pick the curve: `H180 = IFERROR(INDEX(H$162:H$179, MATCH($E180,$D$162:$D$179,0),) * (1+$G180), …)`. In C, the "1" equals the asset number that sheet "1" matches on (`MATCH($A$1,'Time-based'!$C$55…)`).
- Frozen at H3 (xSplit=7, ySplit=2). The saved scroll position is O108.

### 2.4 Summary
- Three panels side by side, each a column triple of label / unit / value.
- **C:D** is "Tests": label in C, OK/CHECK in D.
- **F:G:H** is the inputs and KPI panel: label F, unit G (italic), value H.
- **J:K:L** is the cash-flow panel: label J, PLNm K, EURm L.
- Panel headers use fill white darkened 15% (D9D9D9) with bold text. Bold totals use fill white darkened 5% (F2F2F2), e.g. J6 "OPERATING CASH FLOWS", J29, J32, J56.
- Columns A, E, I and M are spacers (≈3.7). There are no gridlines.

### 2.5 Column widths and spacers (all sheets)
- Spacer columns are 2.7–3.9 wide: column A on every sheet, and G in "1".
- Label columns are 32–52 wide. Timeline columns are 12.4–12.6.
- Gridlines are off on every sheet except Land costs. Zoom is 85% on Assumptions and 79% on Land costs.

---

## 3. Formatting conventions

### 3.1 Font and colour code (from the Cover legend B8:B10, verified in cells)
| Meaning | Font | Fill | Example |
|---|---|---|---|
| **Hard-coded input** | Arial 10, **0000FF** (bright blue) | **FFF2CC** (light yellow). Summary also uses theme accent4 FFC000 tint +0.8, which renders as the same FFF2CC. | Names!E7 365; Assumptions!I14 31-Dec-22; Summary!H5 2.8 |
| **Input brought from another sheet** | Arial 10, **568AB5** (accent1 mid-blue) | none | '1'!E15 `=INDEX(Assumptions!$G$19:$G$19,…)`; '1'!C29 `='Time-based'!D71` |
| **Calculation** | Arial 10, black (theme 1) | none | '1'!J15 |
| Cell to leave empty | — | pattern fill | Cover!B10 |
| Computed value on an input sheet (not a scenario selector) | black | **EDEDED** (light grey) | Assumptions!G25 `=G21`, G30 `=EDATE(G28,G29*Months_Yr)-1` |
| Active-model id cells | black | **FAA47A** orange | Assumptions!G8:G9 |
| Sheet banner (row 2) | Arial 12 bold, white | **44546A** (dk2) | every sheet B2 |
| Section bar | Arial 10 bold, white | **568AB5** (accent1) | '1'!B13 "Flags", Names!B6 "Constants" |
| Table header (Names) | Arial 11 bold | none, **medium bottom border** | Names!B4:E4 |
| Unit column | Arial 10 *italic*, grey (theme1 tint +0.5, or theme0 tint −0.35) | none | Names!C7 `Num#`; '1'!D15 `%` |
| Note text | Arial 10 italic | none | '1'!C12 |

Deviations found:
- Links pasted onto input cells keep the blue-on-yellow input style. For example Assumptions!I19 `=Summary!H11` and I598:I615 `=Summary!$H$…` look like inputs although they are links.
- Summary!H13 `=H9*50%` is styled as an input but is a formula with an embedded literal.
- Time-based uses red FF0000 font (166 formula cells) and green 00B050 (50) for ad-hoc highlighting.

### 3.2 Number formats (formula cells, by frequency in "1")
| Format | Use | Count in "1" |
|---|---|---|
| `#,##0;\(#,##0\);\-` | k-currency flows (LCY'000), counters | 97,689 |
| `#,##0%;\(#,##0%\);\-\%` | period-fraction flags | 9,907 |
| `#,##0.00%;\(#,##0.00%\);\-\%` | rates, IRR | 5,784 |
| `#,##0;\-#,##0;"-"` | legacy variant | 5,443 |
| `#,##0.00\x;\(#,##0.00\x\);\-\x` | DSCR, multiples | 4,636 |
| `#,##0.0;\(#,##0.0\);\-` | Summary PLNm / EURm | 105 on Summary |
| `#,##0.0%;\(#,##0.0%\);\-\%` | CF %, fractions of year | 2,148 |
| `dd\-mmm\-yy;@` | dates (Assumptions, timeline rows 4–5) | — |
| `\Q#;\Q#;\-` | quarter number | row 6 |
| `#,##0.0000000;…` | Verysmallnumber | Names!E18 |

The common pattern is **negatives in parentheses and zero shown as "-"** with a matching unit suffix (`\x`, `\%`).

### 3.3 Units and scale
- Calculations are in **LCY'000** (PLN thousands). The unit label is `LCY'000` on "1" and `PLNk` on Assumptions.
- Summary converts to PLNm with a hardcoded `/1000` (e.g. `K8 = SUM('1'!I148:I148)/1000`) and to EURm with `=K/$H$17` (the EURPLN input).
- Named scale constants `Thousand` and `Million` exist and are used in "1" (e.g. `AA380 = AA376*AA357/Thousand`), but not consistently.

### 3.4 Borders and conditional formatting
- **Borders** are sparse. Imported constants in E:F of "1" have a thin box border. The Assumptions G column has hair left/right borders, and G7:G10 is a hair-boxed panel.
- **Conditional formatting**:
  - Summary D28:D40: "CHECK" gives red fill FF7C80 with font C00000. "OK" gives green fill 92D050.
  - '1'!F251/F253: =0 gives green 35EB39.
  - '1'!F699 and H754:H756: >0 gives pink FFA7A7.
  - '1'!E963/E978/E1021/E1080/E1087/E1102 (DSRA/DSRF status text): "Active" gives green 92D050.
  - '1'!J15:FR15: colour scale on the construction flag.

---

## 4. Formula conventions

### 4.1 Timing flags ("1" rows 13–24)
- **Fractional period flag** (the core pattern, used for every dated window: construction, operations, PPA, debt tenor, margin terms, O&M terms, WTG commissioning):
  `J15 = MAX(0,(MIN($F15,J$5)-MAX($E15,J$4)+1)/J$8) + ((($E15-$F15+1)/J$8)*($E15>J$4)*($F15<J$5))`
  E is the start date and F the end date. The result is the fraction of the period's days inside the window, formatted as %.
- **Event flag (date inside period)**: `J22 = AND($E22>=J$4,$E22<=J$5)*1`. Variants are `N(AND(AA4<=$E1026, AA5>=$E1026))` and `AND($E846>=AA4,$E846<=AA5)*1`.
- **First / last operating period**: `J23 = AND($E17>=J$4,$E17<=J$5)*1` and `J24 = AND($F17>=J$4,$F17<=J$5)*1`.
- **Operating year counter**: `J18 = MAX(ROUNDUP(SUMPRODUCT($J17:J17,$J9:J9),0),1)*(J17>0)`.
- **Construction month counter**: `J16 = (J15>0)*MAX(MONTH(J5)-MONTH($E15), I16+(MONTH(J5)-MONTH(J4)+1))`.
- **Annual step-up flag** (indexation quarter): `J27 = (J6=$E27)*1`.
- **Indexation index**: `AA340 = IF(AA$7<=$F340, 1, Z340*(1+INDEX(AA$29:AA$42, MATCH($E340,$C$29:$C$42,0))))`. It is a compounding index row, with the base year in F.
- Flags are applied by multiplication rather than IF. Examples: `AA879 = $E879*AA870*AA$9*(AA$850>0)*AA$15` and `AA627 = -1*(J$17>0)*SUM(…)`.

### 4.2 Lookups
- **Scenario pick** in Assumptions (388 of 445 G cells), as an array formula:
  `G14 {=IF(ISBLANK(INDEX(I14:Q14,,MATCH(G$10,I$5:Q$5,0))), I14, INDEX(I14:Q14,,MATCH(G$10,I$5:Q$5,0)))}`
  This means **a blank scenario cell falls back to scenario 1 (column I)**. Only deltas need to be entered in scenarios 2–8.
- **Asset pick** in "1": `=INDEX(Assumptions!$G$19:$G$19, MATCH($A$1, Model_Number, 0))`, or the array form `{=INDEX(Assumptions!$G533:$H533,$A$1)}`. A1 is the asset number. Every scalar enters "1" this way, through columns E/F.
- **Curve pick by name**: `INDEX(curve block, MATCH(name, label column, 0))`, with named label ranges such as `PPA_Price`, `GWA_Prices`, `GoO_Prices` and `Inflation_Range`.
- **Annual → quarterly mapping**: `J29 {=IFERROR(INDEX('Time-based'!$H71:$BE71,,MATCH(J$7,'Time-based'!$H$69:$BE$69,0))*J$27,0)}`. Base-rate curves map by date with `INDEX(...,MATCH(AA$5,…))`.
- **Date-keyed aggregation**: `SUMIFS('1'!J374:FR374,'1'!J5:FR5,EOMONTH(G224,MOD(3-MONTH(G224),3)))`. DSCR uses annualised SUMIF by year: `AA1010 = IFERROR(SUMIF($J$7:$FR$7,AA$7,$J$1008:$FR$1008)/SUMIF($J$7:$FR$7,AA$7,$J$1006:$FR$1006),…)`.
- **SUMPRODUCT** appears for cumulative and weighted measures (operating year, cash yields `I74`, contracted revenue %).
- **CHOOSE** selects the funding mechanism: `J753 = IF($E$750=0,0,(CHOOSE($E$750,J731,J738,J745)+J720)*(1-$E$751))`.
- **OFFSET** is used exactly once (`AA965 = IF($E955="DSRA",OFFSET(AA961,,1),0)`, a forward-looking DSRA target). Otherwise it is avoided.

### 4.3 One formula per row
- In "1", 954 rows carry formulas in J:FR, and **941 rows (98.6%) have exactly one formula copied across all 165 columns** (checked with openpyxl's Translator, relative R1C1 equivalence).
- The 13 exceptions:
  - first-column variants: rows 70, 731, 746, 971, 998, 1114, 1300, 1466, 631
  - a 43-column annual block: row 1304
  - stray constants: J656 `;.`, J915 `1`, a cell in row 700
- Row totals sit in I (`=SUM(Jr:FRr)`). Ratios and min/max results sit in E/F, for example `E1010 = MIN(J1010:FR1010)` and `F1010 = AVERAGE(J1010:FR1010)`.
- Sign convention: cash outflows are negative in the cash-flow statement, and a `-1 *` or `-` is applied at the point where a positive calc row is linked in (e.g. `J102 = -1*SUM(J839,J949)`).

### 4.4 Named ranges (35)
- **Constants (Names sheet)**: `Days_Yr` E7 365, `Months_Yr` E8 12, `Months_Semian` E10 6, `Qtrs_Yr` E11 4, `Months_Qtr` E12 3, `Hours_Day` E13 24, `Hours_Year` E14 8760, `Thousand` E16, `Million` E17, `Verysmallnumber` E18 1e-7, `Rounding_Tolerance` E19 6, `BS_Rounding_Tolerance` E20 1, `FundingMechanism` E23:E25 (Equity first / Pro-rata / Debt first).
- **Selectors and lists**: `Model_Number` Assumptions!G9, `Refi_sizing_switch` Assumptions!I599:P599, `EURPLN` 'Time-based'!F111, `GWA_Prices` TB!D162:D179, `PPA_Price` TB!D196:D208, `PPA_GWA` TB!D213:D225, `GoO_Prices` TB!D239:D245, `Inflation_Range` TB!D71:D84.
- **Macro COPY / PASTE / Delta triplets** (all on "1", J:FR):
  - `M1_Funding_Requirements_COPY` J718, `_PASTE` J719, `_Delta` E720
  - `M1_PF_DebtService_COPY` J864, `_PASTE` J865, `_Delta` E866
  - `M1_PF_PaySchedule_COPY` J912, `_PASTE` J913, `_Delta` E913
  - `M1_PF_LockUp_COPY` J1012, `_PASTE` J1015, `_Delta` E1017
  - `M1_Refi_DebtService_COPY` J1040, `_PASTE` J1041, `_Delta` E1042
  - `M1_Refi_PaySchedule_COPY` J1064, `_PASTE` J1065, `_Delta` E1065
  - `M1_Taxes_COPY` J1483:FR1486, `_PASTE` J1487:FR1490, `_DELTA` E1483
  - `M1_IDC_Copy` F1207:F1208, `M1_IDC_Paste` E1207:E1208
  - The `M1_` prefix stands for model/asset 1.
- **Broken names**: `Master_Check` = `#REF!` (the master check name is dead) and `EURPLN2` → external workbook `[1]Założenia czasowe!$E$128`.

### 4.5 Hardcode avoidance
- Day, month and scale constants come from Names, e.g. `EDATE(G28,G29*Months_Yr)-1` and `/Thousand`. The rule is followed only partially; see §9.
- Lists are reserved with "Spare" rows and columns: Assumptions J:P, Time-based curve rows, capex schedule rows 13–14 "Spare".

### 4.6 Circularity handling
- **Iterative calculation is OFF.** `<calcPr calcId="191029" calcMode="manual"/>` has no `iterate` attribute, and calculation mode is **manual**.
- Every circularity is broken with a **copy-paste macro**:
  1. A live calc row is computed (COPY).
  2. A hardcoded PASTE row feeds the dependent calculations.
  3. A Delta row computes `=ABS(COPY-PASTE)`, summed and rounded in column E, e.g. `E720 = ROUNDDOWN(SUM($J$720:$FR$720),Rounding_Tolerance)`.
  4. A VBA loop pastes values and recalculates until every delta is 0.
- Circularities broken this way:
  - funding requirement including IDC and fees (rows 718–720)
  - debt sculpting / debt service capacity (rows 864–866)
  - sculpted amortisation profile (rows 912–914)
  - DSCR lock-up (rows 1012–1017)
  - refi debt service and profile (rows 1040–1042, 1064–1066)
  - taxes (rows 1483–1490)
  - IDC capitalised into tax additions (E/F1207:1208)
- PASTE cells are styled as inputs (blue on yellow). That is why "1" holds about 1,815 hardcoded numbers (see §9).
- Debt sizing by CFADS is not goal-seek. It uses the PV of the debt-service capacity: `AA893 = IF(AA846=1,1,Z893*(1+AA890))` gives a discount index, `AA895 = AA865/AA893` the discounted cash available, and the sum in `F895` feeds `E919` "PF debt sizing quantum - paste".

---

## 5. Checks
- **Local checks** sit in "1" column F under a "Check" caption, each summing a row of per-period tests:
  - B/S check I: `J251 = ROUNDDOWN(ABS(J249-J244),BS_Rounding_Tolerance)` with `F251 = SUM(J251:FR251)`.
  - B/S check II: rows 252–253, `J253 = IF(J252=0,0,1)`.
  - Cash waterfall vs cash-flow statement: `J661 = ROUND(J659-J141,BS_Rounding_Tolerance)`, F661.
  - Sources = Uses: `J699 = IF(ROUND(ABS(J687-J697),BS_Rounding_Tolerance)=0,0,1)`, F699.
  - Trapped cash: `F141 = ROUNDDOWN(MAX(J141:FR141),BS_Rounding_Tolerance)`.
  - Negative equity: `F256 = NOT(I256>=0)*1`.
  - Funding reconciliations: H754:H756.
  - SHL fully repaid: E1154.
  - Macro convergence deltas: E720, F866, F913, E1017, F1042, E1483.
  - DSCR lock-up and default counts: F1012, F1013.
- **Aggregation**: Summary!C26:D40 "Tests". Each line is `=IF(<cell>=0,"OK","CHECK")`, for example:
  - `D28 = IF(('1'!$F$251+'1'!$F$253)=0,"OK","CHECK")` (B/S)
  - D29 PPA input consistency
  - D30 trapped cash `'1'!$F$141`
  - D31 waterfall `'1'!$F$661`
  - D32 funding macro `'1'!$F$720`
  - D33 PF macro `'1'!$F$866+'1'!$F$913`
  - D34 lock-up `'1'!$F$1012`
  - D35 default `'1'!$F$1013`
  - D36 SHL repaid `'1'!$F$1154`
  - D37 IDC `'1'!H1207/H1208`
  - D38 negative equity
- The **overall check** is `D40 = IF(AND(D28="OK",…,D38="OK"),"OK","CHECK")`. It is bold, conditionally formatted red or green, and has the instruction text "Recalculate if you see [CHECK]" (C42).
- There is no check banner on each sheet. Assumptions!E2 has a "Checks:" caption in the banner with no value. The `Master_Check` name is broken.
- Tolerances are named: `BS_Rounding_Tolerance` = 1 decimal, `Rounding_Tolerance` = 6 decimals.

---

## 6. Macros (VBA)
Modules: Module1 (empty), Module2, Module3, ThisWorkbook and Sheet1–Sheet11 (all empty). There are no event handlers, so all VBA is button-driven. Buttons are drawing shapes with `macro="[0]!…"`:
- Assumptions G3 "M1 Equity Solver", G4 "M1 PF Debt Sizing", G5 "M1 Refi Debt Sizing".
- Summary B5 ">> RECALCULATE MODEL <<" and C44 "Refi Debt Sizing", both calling `Refi_Debt_Sizing`.

| Procedure | Module | What it does |
|---|---|---|
| `M1_Equity_Solver` | Module3 | Loops until the funding, lock-up and tax deltas are 0, capped at **5,000 iterations**. Each pass pastes COPY→PASTE for funding requirements, DSCR lock-up and taxes, then runs `Calculate`. It finally pastes the IDC values (`M1_IDC_Paste = M1_IDC_Copy`). |
| `M1_Equity_Button` | Module3 | Calls `M1_Equity_Solver` and shows a "Calculations completed." message box. |
| `M1_PF_Debt_` | Module2 | Loops (no iteration cap) until the PF debt-service, PF pay-schedule and funding deltas are 0. It pastes those 3 COPY→PASTE pairs and recalculates on each pass, then shows a message box. |
| `M1_Refi_Debt_` | Module2 | Same loop for the refi debt service and refi pay schedule. |
| `Refi_Debt_Sizing` | Module2 | **Master "recalculate" routine.** Sets `Refi_sizing_switch=1`, runs 10 alternating passes of `M1_Refi_Debt_` and `M1_Equity_Solver` (with `DoEvents`), sets the switch back to 0, runs `M1_Equity_Solver` once more and shows a message box. |

Key snippet:
```vba
Do Until [M1_Funding_Requirements_Delta] = 0 And [M1_PF_LockUp_Delta] = 0 And [M1_Taxes_Delta] = 0 Or i >= 5000
  [M1_Funding_Requirements_PASTE].Value = [M1_Funding_Requirements_COPY].Value
  [M1_PF_LockUp_PASTE].Value = [M1_PF_LockUp_COPY].Value
  [M1_Taxes_PASTE].Value = [M1_Taxes_COPY].Value
  i = i + 1
  Calculate
Loop
[M1_IDC_Paste].Value = [M1_IDC_Copy].Value
```
Conventions:
- Macros reference **named ranges only**, never addresses.
- Every loop is driven by a Delta cell that is a rounded sum of ABS differences.
- `ScreenUpdating` is turned off during the loop.
- There are no scenario-runner, data-table or goal-seek macros.

---

## 7. Scenario handling
- **Scenario selector**: `Assumptions!G10`, an integer 1–8 entered as a blue input. Column G is the live case. I:P hold up to 8 stored cases, and only I ("Jarszewo") is populated; J–P are "Spare". Blank cells fall back to scenario 1.
- **Asset selector**: `'1'!A1` = 1 matched against `Model_Number` (Assumptions!G9). This leaves room for several asset sheets sharing one Assumptions sheet.
- **Curve selectors**: text names held on Assumptions (data-validated lists), e.g.
  - I77 production case P50/P90 list `$D$72:$D$75`
  - I204 GWA price curve (`GWA_Prices`)
  - I212 PPA (`PPA_Price`)
  - I152:P173 inflation curve (`Inflation_Range`)
  - I480 funding mechanism (`FundingMechanism`)
  - I586 "DSRA,DSRF"
  - I381 "1,0"

  Time-based Selection rows resolve them.
- **Sensitivity hooks**: each Selection row has a `$G` sensitivity adder, e.g. `*(1+$G180)` on prices and `+$G55` on availability, plus `'1'!J324` "Generation sensitivity". Switches use `0/1` with unit `Switch`, e.g. Summary!H34 refi ON/OFF, Assumptions I533 debt sizing on/off, I575 cash sweep.
- **No Excel data tables, no scenario manager, no stored-output table.** To compare cases you change G10 and rerun the macro.
- The Summary also acts as a deal-input sheet. Assumptions links to it, e.g. I19 `=Summary!H11`, I598:I615 `=Summary!H34…H42`, and 'Time-based'!H72 `=Summary!$H$25`. This is a layered override: Summary key inputs → Assumptions scenario column → G live → "1".

---

## 8. Outputs
- **Returns** live in "1" rows 51–76 and are shown on the Summary:
  - Equity funding (rows 54–56) and distributions (59–63). Asset CF (row 65) = FCF before funding. Equity CF: `J66 = J63+J56-(J22*Verysmallnumber)`. The tiny epsilon forces a sign change for XIRR.
  - **IRR** per column: `J67 = IFERROR(XIRR(J65:$FR65,J5:$FR5)*J$22,0)`. It is evaluated only in the investment-date column, and `E67 = MAX(J67:FR67)` picks it up. Asset IRR is E67 and equity IRR E68.
  - **MOIC**: `E65 = -SUMIFS(J65:FR65,J65:FR65,">0")/SUMIFS(J65:FR65,J65:FR65,"<0")`, likewise E66.
  - **NPV**: cash flows are re-bucketed to the investment date (`J70`), then `F71 = XNPV(E71,J71:FR71,$J$70:$FR$70)` with rates hardcoded in E71/E72 (7%).
  - **Cash yields**: Y1–5, Y6–15 and life (rows 74–76) via SUMPRODUCT of an operating-year window flag.
  - **Contracted NPV share** at 7% and 10% (rows 401–402).
- **Debt metrics**:
  - DSCR per period is annualised by year: E1010 = min, F1010 = average. Refi DSCR is in row 1133.
  - Lock-up and default flags are in rows 1012–1013.
  - Transfer-pricing metrics (rows 261–264): Debt/EV and Debt/EBITDA.
  - LTV is shown on Summary!H29 `=-H8/H7`.
- **Financial statements**: cash-flow statement (rows 82–143, CFADS row 88), income statement (rows 145–217), balance sheet (rows 219–256).
- **Summary design** is three panels:
  - **Tests**: OK/CHECK panel.
  - **Deal and KPI** column F:H. Acquisition (price/MW → EV → bank debt → equity, tranche dates), energy production, initial debt, refinancing inputs, profitability (FCFF, IRR on assets, MOIC; FCFE, equity IRR, MOIC).
  - **Lifetime cash flows** J:L in PLNm and EURm. Operating CF, revenues, opex lines, EBITDA, CIT, CFADS; investment CF; financing CF with every debt and fee line; net cash flows.

  Values are links to the "1" column-I lifetime totals divided by 1000, plus links to the E/F result cells. The page has no charts.

---

## 9. Hardcodes vs formulas (evidence for the "everything is a formula" standard)

| Sheet | Formulas (incl. array) | Numeric constants | Text | Formulas with embedded literals ≠ 0/1 |
|---|---|---|---|---|
| Cover | 0 | 0 | 8 | 0 |
| Summary | 132 | 12 (all blue/yellow inputs) | 156 | 40 (mostly `/1000`, plus `H13 = H9*50%`) |
| Assumptions | 529 (388 array selectors) | 255 (246 blue on FFF2CC) | 967 | 6 (`*12`, `*1000`) |
| Time-based | 12,366 (449 array) | 3,956 (2,335 blue/yellow inputs, 1,234+348 header dates in banner style) | 312 | 828 (e.g. `R59 = Q59-0.2%` degradation, `P279 = 4.25%*P180` balancing) |
| Land costs | 11,194 | 504 (489 blue on FFFFCC) | 325 | 10,572 (`EOMONTH(…,3)`, `/1000` in every cell) |
| **1 (calc)** | **136,636 (2,528 array)** | **2,153** | 1,769 | 4,417 cells in 69 rows |
| Names | 1 | 13 (all inputs) | 40 | 0 |

The 2,153 numeric constants in calc sheet "1" break down as follows:
- **1,815 are macro PASTE values** in rows 719, 865, 913, 1015, 1041, 1065, 1141 and 1487–1490. These are deliberate circularity breakers, styled as inputs.
- 165 zeros in row 1474 "Gross sale proceeds", a placeholder exit.
- 157 stray numbers in row 1157 "SHL interest rate" (J:FR). These are junk values: the row's real driver is E1157.
- About 16 genuine row-constants in E/F: NPV rates E71/E72 = 7%, E401/E402 = 7%/10%, cash-yield windows E74:F75, contracted-% windows E393:F393, the IDC seed E1207:E1208, A1 Model_Number, and the I16/I17 opening zeros.

Embedded literals in "1" are mainly calendar arithmetic and rounding:
- calendar: `ROUNDUP(MONTH(J5)/3,0)`, `DATE(J7,12,31)`, `EOMONTH(J70,3)`, `$F$1217/4`
- scale: `/1000` (rows 512, 525, 565–566, 1289, 1294, 1313, 1327–1455)
- rounding precision: `ROUND(…,4)` and `ROUND(…,2)`
- one switch test: `IF($F$5=3,…)`

Net assessment: the calculation sheet is about **98.5% formulas**. Excluding the macro paste rows, only about 340 constants remain, nearly all of them placeholder or junk rows. Inputs are confined to Assumptions, Time-based, Land costs, Summary and Names, as the Cover legend requires ("these should be confined to the Scenarios and Time-Based tabs"). The standard is followed with leaks: literal `/1000`, `*12`, `/4`, `3`, `0.2%`, `4.25%` and `50%` inside formulas, while `Thousand`, `Months_Yr` and `Qtrs_Yr` names exist and go unused in those cells.

---

## 10. Distilled rules for an engine generating similar models
1. Use a fixed sheet stack: Cover (legend) → Summary (KPIs + checks + key deal inputs) → Inputs scalar (scenario columns) → Inputs time-series (curves + Selection rows) → one calculation sheet per asset → Names (constants).
2. Give every sheet a row-2 banner `="<Project> - <Purpose>"` (Arial 12 bold, white on 44546A) and a closing `***END***` row. Hide unused rows and columns.
3. Calc sheet columns: B section, C label, D unit (italic grey), E/F row constants or results (mid-blue links, thin box), I lifetime total `SUM(J:last)` and opening column, J onward the timeline. Freeze at E6.
4. Timeline header rows: Period Start, Period End (EOMONTH by periodicity), Quarter, Year, Days, Fraction of year, Days in year. All flags derive from these rows plus E/F dates through the fractional-overlap formula.
5. Write one formula per row across the whole timeline with no first-column special cases. Use the previous-column reference (`Z910`) for b/f, and lay corkscrews out b/f → movements → c/f.
6. Colour code: input = 0000FF on FFF2CC; cross-sheet link = 568AB5; calc = black; computed-on-input-sheet = black on EDEDED. Section bar = white bold on 568AB5; banner = white bold on 44546A.
7. Number formats: `#,##0;(#,##0);-` for k-currency, `0.00%` style with a `-%` zero, `0.00x` for multiples, `dd-mmm-yy` for dates. Show units explicitly in D (`LCY'000`).
8. Scenarios: active column G = `IF(ISBLANK(INDEX(I:Q,,MATCH(sel,ids,0))), I, INDEX(…))` with selector G10. Calc sheets pull scalars only through E/F INDEX/MATCH.
9. Circularity: no iterative calculation. Use COPY / PASTE / Delta row triplets with named ranges and a VBA loop that pastes values until the deltas are 0, plus an iteration cap. Expose each delta as a Summary check.
10. Checks: compute per-period tests in calc rows, put the sum in column F, and turn each into OK/CHECK on Summary with an AND-aggregated overall check and red/green conditional formatting. Keep a working Master_Check name (Rock's is broken).
11. Returns: XIRR evaluated in the investment-date column only, MAX'd into E. MOIC via SUMIFS of positives over negatives. XNPV with the discount rate in E. DSCR annualised with SUMIF by year, min in E and average in F.
12. Gaps to avoid: literal constants inside formulas (use the Names constants), links styled as inputs, stray constants in calc rows, broken names and external links, and legacy sheets carrying their own timelines and inflation tables.
