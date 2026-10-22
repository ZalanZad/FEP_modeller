# Bora v5 Summary – one sheet on the colleague's model

**Decision this serves:** the IC needs to see, on one sheet, what the full build of both Bora bids costs (total and per MW,
on both capacity bases), what it generates per year, what the returns are, and what the bid strategy is (overbuild 40 %,
BESS 40 %, changeable). The user chose to do this inside the colleague's workbook, keeping its calculation engines.

**Source (read-only):** `reference/HUN_Wind_MOIC_HTM_IC_Bora_v5_20261022.xlsm` (MOIC model v5, last run 22-Oct-2026 16:48).
**Output:** `projects/bora_v5_summary/out/Bora_v5_Summary.xlsm` built by `projects/bora_v5_summary/build.py`.

## 1. What stays, what goes

Engines stay untouched except for the cells listed in §4: `MOIC` (monthly RtB), `HTM` (annual build & operate), `Deposits`,
`Curves`, `Devex_FEP`, `BORA_Devex` (hidden data), `Inputs`. Do not insert or delete rows or columns on any kept sheet
(openpyxl does not re-point formulas). `Open_Items` and `Change_Log` stay hidden; add one Change_Log line for this change.

Delete (all duplicates of inputs or results): `Control_Panel`, `Dashboard` (298 rows, incl. the 01-Oct tornado snapshot and
4 charts), `Bora`, `Run_Map`. Every formula that refers to a deleted sheet must be re-pointed first (scan every sheet):
`Inputs` G:I links to `Control_Panel` become plain blue input values; `Inputs!E7`/`E6` selector per §5.
On `Cover` delete the headline block (B26:C37, duplicate of Dashboard) and rewrite the sheet list (B5:D14) and the update note.

Add a new `Dashboard` (replaces the old one) as the first sheet after `Cover`. One formula per figure, every figure once. Links only (green), no
calculations on Dashboard except the points lookups, the per-MW divisions, the "both bids" sums and the checks.

## 2. Capacity bases (the question the user asked)

- Connection capacity `MVA` = 100 per bid × 2 bids = 200 MVA (the tender right; connection fees, TSO security, municipal
  one-off, RtB price are per MVA).
- Installed wind `MW_inst` = MVA × (1 + overbuild) = 280 MW at 40 % (capex, O&M, land, PB base, DSA are per installed MW).
- BESS `MW_B` = BESS share × MVA = 80 MW; `MWh_B` = MW_B × 4 h = 320 MWh (PB base, BESS capex per MWh).
- Generation: installed MW × net hours × (1 − bottleneck loss × (MW_inst − MVA)/MW_inst). Bottleneck loss on the overbuilt
  MW = 5 % (Bora!N18, as in the fepm model).
Every per-MW figure on Dashboard is shown twice, in two columns: `k€/MW installed` and `k€/MVA`. Equivalent hours likewise.

## 3. Inputs: new block "U. Bid strategy and Dashboard levers" (Inputs rows 300+, same column layout C label | D unit | E value | F note)

Blue input values; source notes ≤ 60 chars.
Lever names (amended 22-Oct-26): the levers live on the Dashboard as `Dash_*` (`Dash_MVA_Per_Bid`, `Dash_Overbuild`, …, see §8.4); `In_Bids`, `In_Bottleneck_Factor` and `In_Min_Hours` keep the `In_` prefix.

| Name (defined name) | Value | Note |
|---|---|---|
| `In_MVA_Per_Bid` Connection capacity per bid, MVA | 100 | Kiírás Annex 1 |
| `In_Bids` Number of bids | = Inputs!E82 (link, not a new input) | |
| `In_Overbuild` Overbuild, % of MVA (ÖSZ-3) | 40 % | User 22-Oct-26 |
| `In_BESS_Share` BESS power, % of MVA (ÖSZ-4) | 40 % | User 22-Oct-26 |
| `In_BESS_Capex_MWh` BESS capex, k€/MWh | 100 | User 15-Oct-26 (CLAUDE.md) |
| `In_Bottleneck_Loss` Energy lost on overbuilt MW | 5 % | Bora!N18 |
| `In_PB_Mult` Performance security multiple (ÖSZ-6) | 1.0 | User 12-Oct-26 |
| `In_Partial`, `In_Brownfield`, `In_Shared`, `In_EU_Turbine`, `In_EU_Transformer` flags | 1, 0, 1, 1, 1 | bid |
| `In_OSZ4_Basis` ÖSZ-4 basis (0 = MVA, 1 = installed wind) | 0 | Meeting file basis |
| Points band tables ÖSZ-1, 2, 3/4, 6 (lower bounds, inner points, inclusive flag), weights, yes/no inner points | from `docs/domain/tender_2026_rules.md` "Points tables" | Kiírás Annex 3/4, GYIK Q20, Q34 |

Copy the band-table layout and the lookup formula from `projects/bora_bidding/build.py` (`Osz{k}_LB`, `Osz{k}_Pts`,
`Osz{k}_Incl`, `Osz{k}_W`; formula `INDEX(pts, 1+SUMPRODUCT((LB<x)+(LB=x)*Incl))*W`). ÖSZ-4 ratio = BESS share ÷
(1 + overbuild) when `In_OSZ4_Basis` = 1. Points are rounded to 2 dp.

## 4. Engine cells that change (the only edits to kept sheets; each becomes a formula or a new value)

| Cell | Was | Becomes | Why |
|---|---|---|---|
| Inputs!E16 gross wind capacity | case INDEX of 200 | `=In_MVA_Per_Bid*In_Bids*(1+In_Overbuild)` | installed wind = committed |
| Inputs!E20 BESS MW | case INDEX of 100 | `=In_BESS_Share*In_MVA_Per_Bid*In_Bids` | |
| Inputs!E168 BESS capex k€/MW | 400 | `=In_BESS_Capex_MWh*E21` | 100 k€/MWh × hours |
| Inputs!E90 PB uplift | 1.5 | `=In_PB_Mult` | PB = 2.5 % × base × multiple |
| Inputs!E78 supplementary security rate (and G78) | 5 % | 0 | none in this tender (GYIK Q70); rows 13–17 of Deposits go to 0 |
| Inputs!G233 / G234 municipal, community | 20 / 2 | 50 / 3 | bid (SPEC bora_bidding §0) |
| Inputs!G162 / H162 / I162 turbine k€/MW | 950 / 950 / 500 | 600 / 600 / 500 | user 12-Oct-26 |
| Inputs!G190:I190 merchant share; G193 PPA tenor | 1 / 0 / 1; 10 | 0 / 0 / 0; 30 | PPA 90 €/MWh for the whole life (user) |
| Inputs!G217 / H217 gearing | 55 % / 50 % | 60 % / 60 % (I217 CN stays 0 → set 50 %, as the fepm model's CN case) | user |
| Inputs!G11 FX | 370 | 365 | Bora model |
| Inputs!G268 DSRA months, G269 cash sweep | 6, 50 % | 0, 0 | no DSRA / sweep (bora_bidding reconciliation 10b) |
| Deposits!G11, G12, G18 capacity basis | `=Inputs!E18` (installed) | `=In_MVA_Per_Bid*In_Bids*Inputs!E17` | TSO security and connection fees are per MVA |
| HTM!H29:AQ29 municipal one-off | `Inputs!E233*Inputs!E18` | `...*In_MVA_Per_Bid*In_Bids` (one formula, copied) | ÖSZ-1 is per MVA |
| HTM!E36 net hours | `E179*(1-E180)*(1-E181)*E183` | `*(1-In_Bottleneck_Loss*(E35-MVA)/E35)` → split: add the loss factor as a new named cell on Inputs block U (`In_Bottleneck_Factor` = formula) so E36 stays ≤ 100 chars | overbuild curtailment |
| MOIC!G143 RTB premium proceeds | `E141*E18` | `=Inputs!E141*In_MVA_Per_Bid*In_Bids` | RtB price is per MVA |

Everything else in the engines stays as it is, including the deemed-exit switch (Inputs E228 = 1), the timing (RTB
31-Mar-29, sale 30-Sep-29, COD 30-Jun-31) and the Aurora curves. These are shown on Dashboard under "Key assumptions" so
the reader sees them; the user decides later whether to align them to the fepm model (there: exit 31-Dec-29, COD 30-Sep-31,
developer buy-out instead of a deemed exit).

Also remove the duplicate VBA: the source holds `Module1` and `Module2` with the same `RunAllScenarios`; both go (they write
to deleted sheets). Replace per §5.

## 5. Cases and the macro (house standard: MODELLING_STANDARDS §7)

Keep cases 1–3 only (Inputs columns G, H, I: "FEP base case / sell at RTB", "European turbine", "Chinese turbine").
Clear columns J:P of the case block (values only, they are spares) and hide them. `Inputs!E6` (case #) becomes
`=MATCH(Scenario_Name,Scenario_Names,0)`, where `Scenario_Name` is the selector cell on Dashboard (text, validated against
`Scenario_Names` = `Inputs!G7:I7`). `Dashboard` holds `Live_Results` (one column, the headline outputs of §6.5) and
`Stored_Results` (3 columns, purple on grey) written by `vba/Scenarios.bas` (`RunAllScenarios`, unchanged); the build runs
`package(xlsx, xlsm, run=["Scenarios.RunAllScenarios"])`. A check row verifies stored = live for the selected case.

## 6. Dashboard sheet layout (columns: B marker, C label, D unit, E value, F value 2, G value 3, H note). Sections in order:

1. **Bid strategy (inputs, blue):** MVA per bid, bids, overbuild %, BESS %, BESS hours (link to Inputs E21), PB multiple,
   municipal k€/MW, community €/MWh, the five yes/no flags, turbine route = `Scenario_Name` selector. Next to each scored
   item its points; total points (2 dp). One line "offerable range" per item as a note only where a range exists.
2. **Capacity and basis:** connection MVA (both bids), installed wind MW, BESS MW and MWh, overbuilt MW, bottleneck loss
   applied (%). Text note: "k€/MW = per installed wind MW; k€/MVA = per connection MVA".
3. **Build cost to COD, both bids** (from Inputs rows 162–172 and HTM rows 27–32, 78–79, 85): one line per item, columns
   `€ m | k€/MW installed | k€/MVA`: turbines, installation, foundations, BoP civil, BoP electrical, substation, owner's
   costs, land/easements, contingency, BESS, connection fee indirect, connection fee direct, municipal one-off, IDC,
   debt fees; **all-in cost to COD**. One check: the items sum to HTM total capex + IDC + fees.
4. **Generation and operations, first full operating year and lifetime:** net generation GWh/yr (full year),
   equivalent hours per installed MW and per MVA, net capacity factor on both bases, energy lost to the bottleneck (GWh/yr),
   achieved price €/MWh, revenue, EBITDA, EBITDA margin, opex k€/MW/yr; lifetime generation GWh and LCOE.
5. **Returns (headline outputs = `Live_Results`):** project IRR unlevered pre-tax and post-tax, project NPV, equity IRR
   levered, FEP equity IRR, FEP NPV, FEP MOIC, FEP peak exposure, min DSCR, debt at COD and gearing, payback year;
   RtB route: FEP net profit post-tax, IRR, MOIC, NPV, A.RES profit; buyer IRR unlevered at the RtB price and the implied
   max RtB premium. Each with its source cell in the note column.
6. **Securities and fees (from Deposits, both bids):** participation fee, bid security, performance security, TSO security,
   connection fee indirect and direct, total; of which cash / guarantee (Deposits E26:E27).
7. **Stored results by case** (`Stored_Results`, 3 columns) for the rows of section 5.
8. **Checks:** MOIC overall (MOIC!G270), HTM overall (HTM!G172), Deposits reconciliation (Deposits!E34), capex breakdown
   = engine, stored = live, points within offerable ranges, capacity tests (BESS ≤ wind, ≥ 4 h); `Master_Check` = number
   of failing rows, shown OK / CHECK; link it on Cover.
9. **Key assumptions (links):** turbine k€/MW, PPA €/MWh and tenor, merchant share, gearing, debt rate, discount rate,
   FX, RtB date, sale date, COD, operating life, deemed exit switch, DSRA months, cash sweep.

Formatting per MODELLING_STANDARDS §5 (Aptos Narrow 10, section bars, blue inputs, green links, purple stored).
No figure appears twice on the sheet (e.g. all-in cost once, in section 3; section 5 does not repeat it).

## 7. Build pipeline and verification

`build.py`: copy the source → openpyxl (`keep_vba=False`, keep links) → edits of §1, §3, §4, §5, §6 → save
`out/Bora_v5_Summary.xlsx` → `fepm.xl.macros.package(..., run=["Scenarios.RunAllScenarios"])` → `out/Bora_v5_Summary.xlsm`
→ LibreOffice recalc copy `out/Bora_v5_Summary_recalc.xlsx`.
Known losses from the openpyxl round trip, acceptable: the 4 Dashboard charts (sheet deleted anyway), one x14 conditional
format, the ▶ form button. List anything else that differs (data validations, number formats) in the build report.

Verification, in this order:
1. **Regression (engine untouched):** build once with the levers at the source values (overbuild 0 %, BESS 50 %, PB 1.5,
   supplementary 5 %, municipal 20, community 2, turbine 950/950/500, merchant 1/0/1, tenor 10, gearing 55/50/0, FX 370,
   DSRA 6, sweep 50 %) and compare the stored case results to the source Dashboard/Bora values: RtB FEP net profit post-tax
   15,682.56 k€, IRR 33.894 %, MOIC 1.3284; HTM EU equity IRR 9.0237 %, project IRR post-tax 8.5337 %, capex/MW 1,845.45;
   HTM CN equity IRR 12.918 %. Tolerance 1e-6 relative. Only then switch to the §3/§4 values.
2. Lint the Dashboard sheet and the new Inputs block (hardcodes, chains, length); Master_Check 0 after RefreshAll.
3. Independent Python recomputation: capex breakdown and per-MW figures on both bases, PB = 2.5 % × (280 × 406 + 320 × 200)
   mHUF × 1.0 / FX, first-year generation = 280 × 2,460 × 0.98 × (1 − 0.05 × 80/280), points = 82.50 (20 + 15 + 9 + 13.5
   + 5 + 0 + 0 + 5 + 12 + 3).
4. Compact dump `out/summary_dump.txt` (label | unit | values | note) for the Fable review.

## 8. Amendments after the first build (22-Oct-2026, Fable review)

1. **RtB price basis everywhere.** MOIC!H226:N226 (premium sensitivity) = premium × `In_MVA_Per_Bid*In_Bids`; HTM E227:E243
   and HTM G221/G222 (implied maximum premium) on the MVA basis, unit k€/MVA. MOIC G266 then passes.
2. **Deposits reconciliation.** The source totals (Inputs G87/G88) describe the old 200 MW / PB 1.5× / supplementary case,
   so that check is a regression test only (it lives in `regression.py`). On the sheet, Deposits E34 becomes an internal
   check: gross total H19 = sum of the item rows, and payable J19 = MOIC!G40 (total payable flows). Relabel C28:C34.
3. **Two cases, not three.** Cases 1 and 2 are now identical. Keep "EU turbine" (column G) and "CN turbine" (column H,
   moved from I; clear I). Update G7:H8 names and descriptions, `Scenario_Names` = G7:H7, `Stored_Results` 2 columns.
4. **Levers live on Dashboard.** Section 1 holds the bid levers as blue inputs (the `In_*` names point at these cells);
   Inputs block U keeps only the band tables, weights, offerable ranges and the derived `In_Bottleneck_Factor`.
   One ÖSZ-4 row: offer in E, points in F, the scored ratio in G (note: basis).
5. **No duplicate rows.** Drop "FEP equity IRR" (equals the 100 % equity IRR); keep one row "Equity IRR, levered post-tax
   (FEP, 100 %)". Relabel "Total gross, incl. supplementary security" → "Total gross".
6. **Round-trip repairs.** HTM G7, G10, G14 number format back to General (they show 3000.00 %).
7. Master_Check must be 0 after RunAllScenarios; `regression.py` must still pass.

## 9. Amendments, second round (user, 22-Oct-2026)

1. **EU turbine 950 k€/MW** (Inputs G162 = 950; CN H162 stays 500). The 600 of §4 was the fepm bid-model decision; this
   workbook keeps the colleague's 950.
2. **No Cover sheet.** Delete it. The master check status moves to the top of the Dashboard.
3. **Dashboard column H (notes) removed.** No source notes on the Dashboard; the cell references are enough.
4. **New first sheet `Cash_Flow`** (replaces the Cover): the 36-year annual grid of the HTM engine (2026–2061, HTM H5:AQ5),
   formulas only, one formula per row copied across, every row a link to its HTM row. Layout: C label | D unit |
   E Total | F per installed MW (k€/MW) | G per MVA (k€/MVA) | I.. timeline (year header row linked to HTM row 5).
   Rows, in order (sign convention as HTM: outflows negative):
   - Spending: development phase (HTM 24), construction capex excl. connection (27), connection fees (28), municipal
     one-off (29), IDC and debt fees paid by equity (HTM 117), **total spending** (sum, bold);
   - Operations: revenue (54), operating costs (69), **EBITDA** (70), tax (100), decommissioning and terminal value
     (107/108 rows as HTM uses them in 122);
   - Project: **unlevered post-tax cash flow** (110), cumulative (111);
   - Debt: drawdown (116), debt service (119), DSRA (120);
   - Equity: **equity cash flow** (123), cumulative (124).
   E = SUM across the timeline; F = E ÷ installed MW (Inputs E18) ÷ 1 (k€ per MW); G = E ÷ (Dash_MVA_Per_Bid × In_Bids).
   Units: EUR k on the timeline (as HTM), E in EUR m. No IRRs here (they are on the Dashboard). Freeze panes at I6.
   A check row at the bottom: total spending + EBITDA + tax + decommissioning = unlevered post-tax cash flow (0 = OK),
   feeding Master_Check.

## 10. Amendments, third round (user, 22-Oct-2026)

1. **CN case zeroes the EU commitments.** The ÖSZ-9 rows on the Dashboard are no longer levers: EU turbines flag and EU
   transformers flag = `N(Scenario_Name="EU turbine")` (black formulas, points follow: 82.50 EU, 67.50 CN). Keep the
   defined names `Dash_EU_Turbine` / `Dash_EU_Transformer` on these cells.
2. **Wind built = 10 % below the bid, no sanction.** New lever `Dash_Build_Shortfall` = 10 % (GYIK Q28/30/76: a shortfall
   up to 10 % carries no sanction). New named cell on Inputs block U `In_Wind_Built` = `Inputs!E18*(1-Dash_Build_Shortfall)`
   (252 MW at 40 %). Committed MW (Inputs E16/E18, 280 MW) stays the base of the PB (E92), the bid security and the points.
   Built MW drives: HTM E35 (installed capacity → `In_Wind_Built`), the capex (Inputs E173 becomes "per built MW":
   per-MW lines × (1 + contingency) + substation × (1 + contingency) ÷ In_Wind_Built + BESS capex × E20 ÷ In_Wind_Built;
   E174 likewise; HTM E27 = E173 × In_Wind_Built), the Dashboard capex lines (section 3) and both per-MW columns, the
   bottleneck factor ((built − MVA) ÷ built). BESS is built in full. Dashboard section 2 shows "Wind committed (bid)" 280
   and "Wind built (bid less 10 %, no sanction)" 252; the per-MW column header reads "k€/MW built".
   HTM G32 and G150 (capex per MW) divide by `In_Wind_Built`, label "Capex per MW built".
3. **Cash_Flow, fewer rows:** Total spending | Revenue | Operating costs | EBITDA (NOI) | Tax | Unlevered post-tax cash
   flow | Debt drawdown | Debt service | Equity cash flow | Cumulative equity cash flow. Nothing else; drop the tie-out check
   row (and its Dashboard check row).
4. **NOI on the Dashboard section 4:** label "EBITDA (NOI), first full year" and add one row "EBITDA (NOI), lifetime" (HTM G70).
5. **Wording:** replace "both bids" everywhere by "BORA 1 + BORA 2 (2 × 100 MVA)" in section headers; elsewhere drop it.

## 11. Amendments, fourth round (user, 22-Oct-2026): one capacity lever, buyer at 8 %, consistency

1. **One capacity lever.** `Dash_MVA` = connection capacity in MVA, default **100** (one bid). Everything scales with it:
   200 doubles the model. "Number of bids" is no longer a lever: `In_Bids = ROUNDUP(Dash_MVA / In_MVA_Per_Bid, 0)` with
   `In_MVA_Per_Bid` = 100 on Inputs block U (note "Bora bids are 100 MVA each"). Inputs E82 links to `In_Bids`.
   Remove every "BORA 1 + BORA 2" / "both bids" wording; section headers say "at the capacity in E7" or nothing.
   All € m, GWh, MW figures are for `Dash_MVA`; the per-MW and per-MVA columns stay.
2. **Buyer at 8 % – sale at COD.** New block at the bottom of HTM (after the last used row, same column layout; labels in
   C, units in D, scalar results in G, timeline H:AQ), "Sale at COD – buyer earns the unlevered discount rate":
   - Post-COD unlevered post-tax cash flow: row 110 × (year > COD year) (one formula, copied).
   - COD sale price (buyer NPV = 0 at Inputs E224): `=SUMPRODUCT(post-COD flow, HTM row 12 DF) / DF at the COD year`
     → price at 31-Dec of the COD year (the sale is assumed at the end of the COD year; FEP keeps the COD-year flows).
   - Price per MW built and per MVA (k€).
   - Debt repaid at the sale: closing debt balance at the COD year (row 84).
   - FEP cash flow – sale at COD: equity cash flow (row 123) × (year ≤ COD year) + (price − debt repaid) × (year = COD year).
   - FEP IRR, FEP profit (sum), FEP MOIC (inflows ÷ outflows) – sale at COD.
   The existing RtB buyer block (HTM rows 202–223) stays: buyer IRR at the 150 k€/MVA price and the implied premium at 8 %.
3. **Deemed exit in HTM.** HTM E21 (FEP profit forgone at RtB, an opportunity cost) = 0: a cash IRR carries no opportunity
   cost. HTM E22 (cash paid to A.RES at the RtB date, the buy-out of the developer) stays 1. Key assumptions row relabelled.
4. **Dashboard section 5 "Returns"** gains a sub-block "Buyer at 8 %": COD sale price (€ m, k€/MW built, k€/MVA),
   FEP IRR / profit / MOIC on a sale at COD; RtB: buyer IRR at 150 k€/MVA, implied premium at 8 %, FEP IRR / profit at 150.
   `Live_Results`/`Stored_Results` include the new FEP COD-sale IRR and profit and the COD price.
5. **Review.** After the build, the `model-reviewer` (Fable) reads `out/summary_dump.txt` as a senior modeller against the
   tender rules (`docs/domain/tender_2026_rules.md`) and market practice; its findings are verified by `model-auditor`
   and fixed in this build before delivery.
6. **Lump sums scale with capacity.** The colleague's lump sums were quoted for the 200 MVA Bora project (substation
   Inputs E167, admin E207, signing fee E24, DD E25, sale costs E26, the BORA_Devex line items via the devex scaling
   factor E294). New block U cells: `In_Source_MVA` = 200 ("Source basis of the lump sums: Bora 2 × 100 MVA") and
   `In_Scale = Dash_MVA / In_Source_MVA`. Each lump sum = source value × `In_Scale` (the case columns keep the source
   values; the Live column E applies the factor; E294 = `In_Scale`). At 200 MVA nothing changes (regression); at 100 MVA
   every € m figure is exactly half. Show "Lump sums scaled from the 200 MVA source" among the key assumptions.
