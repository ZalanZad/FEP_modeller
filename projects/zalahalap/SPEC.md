# Zalahaláp – 63 MWp PV, with or without a 50 MW BESS: investment model SPEC v2 (streamlined rebuild)

Owner: Fable (architecture, review). Builder: `model-builder` (Opus). Auditor: `model-auditor`. Date: 03-Nov-2026.
Supersedes `SPEC_v1.md` (kept for the record: the original-model dissection §3, the JV sheet dissection §7, the fixes
§13 and the reconciliation targets §3.2 are still the reference for what v3 did). Format reference:
`reference/HU_Wind_bids_271026.xlsx` ("the wind model"). Assumption values and the offer term sheet: `ASSUMPTIONS.md`.
Deliverable: **one workbook, `out/Zalahalap_Model.xlsx`**, no macros, no deck.

---

## 0. User decisions (verbatim and how they are read)

**03-Nov-2026:** "anfd make it like the wind model streamlined and easy to follow and that color coding need to make
the case with only pv with 2h battery and with 4h battery make it unevered for now with ppa 5 years for pv"

1. **Format = the wind model**, exactly: Dashboard-centred, annual, one column per case, the wind colour coding
   (green section bars `00563C`, sub-headers bold on light green `E6F4EC`, blue inputs on `FFF2CC`). The grey
   sub-header introduced on 22-Oct (`fepm/xl/styles.py` SUB_FILL = F2F2F2) is reverted to `E6F4EC`.
2. **Three cases, one column each:** PV only (63 MWp, no BESS) | PV + 2 h BESS (50 MW / 100 MWh) | PV + 4 h BESS
   (50 MW / 200 MWh). The revenue scenarios of v1 (original / Aurora / offer) are no longer columns.
3. **Unlevered for now:** no debt, no levered rows, no DSRA, no IDC. Equity funds everything. Financing is out of
   scope until a term sheet exists (one note on Inputs).
4. **PPA 5 years for PV:** a 5-year pay-as-produced PPA from COD at an input price, on all PV output; Aurora
   merchant after that. Default price 45 €/MWh (the MVM offer, the one 5-year offer; `ASSUMPTIONS.md` §2.1).

**Carried over from 22-Oct-2026 (SPEC_v1 §0, unchanged):** JV share step 60 % at price ≤ 90 k€/MWp, 50 % above ·
Aurora Q2-2026 Central is the only market source, no synthetic Low/High · early works = the four input-sheet lines
(785 k€), E48/E49 (725 k€) in construction capex, treatment "to be revisited" · the battery cannot charge from the
grid (PV-only factor by Aurora component) · "build now, we will adjust": every undecided value is a flagged
placeholder, nothing is invented.

**Decision on the v3 reconciliation column (Fable):** dropped from the workbook. The user's three cases do not
include the original configuration (61 MWp, v3 curves, 35-year merchant), and a grouped fourth column would
contradict "streamlined". The tie-out to the original model stays in `reconcile.py` (Python shadow of v3 → bridge
to the PV + 4 h column, §9). The workbook carries only what the decision needs.

---

## 1. Decision
Which configuration to develop (PV only, PV + 2 h, PV + 4 h) and which route to take (sell at RtB, sell at COD,
hold), with a 5-year PV PPA then merchant, unlevered, Base / Low / High. Outputs per case and route: FEP (Futureal)
profit pre- and post-tax, MOIC, IRR; buyer IRR at our price and the implied price at the buyer's target; PSH's share.

## 2. Format conventions (from the wind model; cell refs are the wind workbook's)
| Item | Convention |
|---|---|
| Sheets | wind: Dashboard + Inputs only. Ours: `Dashboard` · `Inputs` · `Revenue` · `Calc` · `Checks` (§3). Font Aptos Narrow 10, gridlines off, zoom 85 %. No Cover, no banner formula: **C2** plain bold title. |
| Columns | A 2 · B 3 (`x` on every section bar) · **C 56 label · D 10 unit** (italic grey `808080`) · **E, F, G 13 = the three cases** · H 15 = Total/Live where meaningful, else blank · I 2 · **J → AP 10 = years 2025 … 2063** (39 columns; wind: J → AS). Freeze `E5` on every sheet. |
| Master check | C3 label, **H3 `=Master_Check`**, format `"CHECK" 0;"CHECK" -0;"OK"`, conditional font green `00B050` when 0, red otherwise (wind H3, H356). |
| Section bar | B..H filled `00563C`, white bold, the **case names repeated in E:G** on every bar, "Total" or "Live" in H (wind B5:H5, B30:H30). |
| Sub-header | C..H bold `3C3C3B` on **`E6F4EC`** (wind C105, C113, C120, C127). |
| Inputs on the Dashboard | the per-case levers at the top: blue `0000FF` on `FFF2CC`, thin border, right-aligned, named `Dash_*` (wind E6:G20). Everything else on `Inputs`. |
| Case selector | bar "Sensitivity case" with **E Base, F Low, G High, H Live**; H selector blue, validation `"Base,Low,High"`, name `Dash_Case`; live row `=INDEX(E:G, MATCH(Dash_Case,$E$r:$G$r,0))` (wind E22:H24). The only scenario mechanism. |
| Calculations | one formula per case cell, the same formula across E:G; identical annual blocks per case with only the column letter changing (wind rows 250–341). |
| Returns | per strategy a sub-header and the rows exposure / proceeds / profit pre-tax / profit post-tax / MOIC / IRR (wind rows 104–133). |
| Annual summary | a "Year" row, then per case a 6-row block (investment, revenue, opex, EBITDA, net cash flow, cumulative) (wind rows 135–160). |
| Detail | grouped, outline level 1, collapsed, titled "Detailed calculations – expand the group" (wind rows 162–345); `summaryBelow` off. |
| Dated flows | the RtB route as a short list: amounts in E:G, **dates in J:L** (wind rows 201–213), `XIRR(amounts, dates)`, MOIC `=SUMIF(>0)/−SUMIF(<0)`; accrual factors closed-form act/365 with 31-Dec capitalisation (wind E169). |
| Annual timeline | J4 `=First_Year+COLUMN()-COLUMN($J$1)`, J5 `=DATE(J4+1,1,1)-1` (year-end dates) (wind J247:J248). |
| Checks | `=N(ABS(…)>Tolerance)` per case, H = sum; master check = sum of the check rows (wind rows 347–356). Ours on `Checks` (CLAUDE.md rule 6), mirrored at Dashboard H3. |
| Number formats | `MODELLING_STANDARDS.md` §5: money `#,##0;(#,##0);"–"`, 1 dp, `0.0%`, `0.00"x"`, `dd-mmm-yy`, flag `0;-0;"–"`, factor `0.000`, accrual `0.0000`. Totals bold with top border; memo rows italic. |
| Inputs sheet | C 54 label · D 11 unit · **E common value** (per-case values across E:G) · **H Source** ≤ 60 characters, italic grey (wind Inputs). Sub-headers by topic. Blue only on inputs. |
| Density | every number visible, no charts, no notes column on the Dashboard, `***END***` last row. |

Where the format and `MODELLING_STANDARDS.md` conflict: format wins on layout, annual grid, the levers on the
Dashboard and the sub-header colour; the standards win on the `Checks` sheet, blue inputs only, one formula per row,
no literals in formulas, ≤ 2 nested functions, ≤ ~100 characters.

## 3. Sheet map
### 3.1 `Dashboard` (top to bottom; E = PV only, F = PV + 2 h, G = PV + 4 h)
1. **C2** "Project Zalahaláp – 63 MWp PV, with or without a 50 MW BESS: sale at RtB, sale at COD or hold". **C3/H3** master check.
2. **Key inputs: the cases** (bar; blue levers, `Dash_*`): case name (text) · PV capacity MWp (63/63/63) · grid export
   MVA (49.9) · specific yield MWh/MWp (1.265) · **BESS built** flag (0/1/1) · BESS power MW (0/50/50) · **BESS duration h**
   (–/2/4) · BESS energy MWh (formula `=MW×h`, black) · Aurora block (formula `=IF(built,"DC "&h&"h","–")`, black) ·
   **PPA price €/MWh** (45/45/45, real 2025 or nominal – see §5.2) · PPA indexation %/yr (0) · PPA years (5) ·
   **BESS revenue basis** (text, validated "Merchant,Tolling"; Merchant) · RtB price k€/MWp (90; 0 = implied) ·
   COD price k€/MWp (0 = implied). 15 rows.
3. **Sensitivity case: Base, Low, High** (bar E/F/G/H Live; H selector `Dash_Case`): PV price change % (0 / −20 % / +20 %) ·
   BESS revenue change % (0 / −20 % / +20 %) · yield change % (0 / −2 % / +2 %) · capex change % (0 / +10 % / −5 %) ·
   COD delay months (0 / 12 / 0) · buyer target IRR (8 % / 9 % / 7 %). Live column = `Case_*` names. 6 rows.
4. **Capacity and energy, first full operating year**: generation GWh · sent to the BESS · exported directly ·
   BESS discharge · full cycles · PV price received €/MWh (blended, all output) · BESS revenue k€/MW. Links to `Calc`.
5. **Revenue and EBITDA**: first full year and lifetime: PV revenue · BESS revenue · GoO · total · opex · EBITDA · margin.
6. **Cash to RtB by nature** (wind rows 68–74): SPV acquisition price · transaction costs · DSA loan to PSH · early works
   (4 lines, one row) · pre-RtB development · deposits and guarantees · guarantee fees · total; of which reimbursed by
   the buyer / returned at the exit / sunk; FEP exposure at RtB. One project, so the rows are identical across the
   cases: show them in E only, F:G blank, the bar reading "all cases".
7. **Capex to COD** (per case): PV equipment + BoP · MV cable · substation/bay · land · BESS (MW × h × k€/MWh) ·
   early works at NTP (E48/E49) · total uses · per MWp. Equity = total (unlevered).
8. **Returns** (sub-headers): *Sale at RtB* – price used · **implied price at the buyer target** · all-in price incl.
   reimbursements · buyer IRR unlevered after tax / before tax at our price · FEP exposure · FEP proceeds · FEP profit
   pre-tax / post-tax · MOIC · IRR · NPV @ 8 % · PSH profit. *Sale at COD* – EV used · implied EV · buyer IRR after /
   before tax · FEP contributions · FEP proceeds · profit pre/post · MOIC · IRR · NPV · PSH profit. *Hold* – FEP
   equity (incl. 2048 repower) · distributions · profit pre/post · IRR after / before tax · MOIC · NPV · project IRR
   after / before tax (= FEP's here, memo) · PSH buy-out. RtB rows are identical across cases except the implied
   price and buyer IRR (which depend on the BESS) – keep per case.
9. **JV split at the RtB exit** (wind waterfall rows 188–199): cash available · sale costs · loan principal to FEP ·
   12 % return · SPV price item · distributable · FEP share % (step rule) · FEP lot · PSH lot · FEP total · PSH total.
   Same at COD (grouped).
10. **Tax**: lifetime CIT on the hold route · tax losses lapsed · tax on the 12 % return at FEP.
11. **Annual summary per case** (wind rows 135–160): Year row; per case investment · revenue · opex · EBITDA · net cash
    flow after tax · cumulative. Links to `Calc`.
12. **Detailed calculations – expand the group** (grouped): the **dated RtB block** (one, shared: lines amount E, date J:
    SPV price, transaction costs, deposits paid, DSA draws by milestone, early works by milestone, pre-RtB development
    by month-end → accrual factor to the exit `=(1+r×days₁/365)×(1+r)^fullyears×(1+r×stub/365)` as wind E169, claim at
    the exit; reimbursements and returns at the exit; FEP dated flows → XIRR, MOIC; PSH dated flows). 3 case columns
    where the exit price differs (implied price per case). ≈ 45 rows.
13. `***END***`.

### 3.2 `Inputs` (E common; per case E:G only where stated; H Source)
Dates (acquisition 30-Nov-25, RtB 31-Dec-26, exit 31-Mar-27, NTP 30-Jun-27, COD 31-Mar-28, life 35 y, Aurora first year
2026, First_Year 2025) · PV (availability 99.5 %, degradation 0.3 %, clipping 0 % DC-coupled, GoO €/MWh 1.75) ·
BESS (RTE 0.86 inside the blocks, haircut 0 %, aFRR-down/FCR share 0, capex k€/MWh 110, repower year by block from
`Revenue`, repower flag 1, opex k€/MW/yr 30.8) · PPA (start = COD, balancing €/MWh 5 merchant only, tolling fee k€/MW/yr
49.7 for 3 y then merchant, tolling start 01-Jan-28) · capex (PV 350 k€/MWp, cable 3,700, substation 9,500, land 620,
grid fee 0, E48/E49 725 at NTP) · opex (O&M 14.96 k€/MWp, cable O&M 5 k€/MWp, lease 96, HIPA 2 %, innovation 0.3 %,
opex indexation = euro index) · development and JV (SPV price 549, transaction costs 200, DSA 15 k€/MWp, early works 4
lines with dates and HUF, pre-RtB development 3,182 by month, deposits/guarantees table in HUF with dates, FX 400,
guarantee fee EURIBOR 1.87 % + 1.75 %, loan rate 12 %, share step 60 % / 50 % / threshold 90, PSH land apport 150,
reimbursement and return flags) · tax (CIT 9 %, Robin Hood flag 0, depreciation PV/grid 30 y, BESS 10 y, loss cap
50 %, lapse 5 y after NTP, sale form share, gain taxed 0) · valuation (discount rate 8 %, sale costs 0 %, buyer IRR per
case row = Live) · constants (Thousand, Days_In_Year, Months_In_Year, Tolerance). **Financing: one row "Debt – out of
scope for now (user 03-Nov-2026); the model is unlevered"**, no inputs. Every placeholder keeps its "placeholder –
pending user" source note; values exactly as `ASSUMPTIONS.md` §1.

### 3.3 `Revenue` (the one-screen view: Aurora curves, offers, what each case earns)
Same column grid; years J:AP; blue = input with source in H; one formula per row.
- **Bar 1 "Aurora Q2-2026 Hungary, Central, real 2025 €"** (source `data/aurora_curves.csv`, Aurora row cited in H):
  baseload (information) · PV direct-export price DC 4 h · DC 2 h · PV transfer price DC 4 h · DC 2 h · BESS storage
  total k€/MW/yr DC 4 h · DC 2 h · PV-only factor DC 4 h · DC 2 h · aFRR-down + FCR share DC 4 h · DC 2 h · charge from
  PV MWh/MW DC 4 h · DC 2 h · discharge MWh/MW DC 4 h · DC 2 h · euro index (2025 = 1) · repower year DC 4 h / DC 2 h
  (E/F). Aurora runs 2026–2060; 2061–2063 flat real; 2025 blank. 17 rows.
- **Bar 2 "Offtake offers as stated – indicative; E = E.ON, F = MET/E2, G = MVM; H = PDF page"**: offeror and date ·
  validity · PV product · fixed price 3 y · fixed price 4 y · fee % of DAM · fee floor €/MWh · fee cap · start ·
  tenor (max/options) · volume · balancing included (flag) · balance-group fee €/month · GoO treatment · downregulation
  · settlement point (reading) · BESS product · BESS configuration offered MW/MWh · tolling fee 3 y / 4 y · floor 3 y
  / 4 y · share above the floor · option B share · offeror's yield assumption. Text or blue numbers, from
  `ASSUMPTIONS.md` §2.1–2.2. ≈ 25 rows. **Reference only**: the Dashboard PPA price lever and the "Tolling" basis
  read from here through two named cells (`Offer_PPA_MVM` = 45, `Offer_Tolling_MET_3y` = 49.7); nothing else links.
- **Bar 3 "Resulting prices by case, nominal – PPA in its tenor, then the Aurora merchant tail"** (year header
  repeated; per case a sub-header with the case name, 6 rows): merchant PV price, direct export (block × index ×
  (1 + Case_PV_Price)) · merchant PV price, stored energy (transfer × index × (1 + Case_PV_Price)) · PPA in force
  flag · PPA price (price × (1 + indexation)^(year − COD year)) · **PV price received, direct export** = merchant +
  flag × (PPA − merchant) · **PV price received, stored energy** = merchant stored + flag × (PPA − merchant stored) ·
  BESS merchant uplift kept k€/MW = storage total × (PV-only factor + aFRR share × In_AFRR_Down_Share) × index × (1 −
  haircut) × (1 + Case_BESS) · tolling in force flag · **BESS revenue received** = uplift + flag × (tolling fee −
  uplift). PV-only case: BESS rows show "–" (block "–" → 0). `Calc` links only to the three bold rows of its case.
  ≈ 30 rows.

### 3.4 `Calc` – three identical annual blocks (one per case; ≈ 75 rows each)
Block constants in E/F of the block's own rows, pulled with `INDEX(Dash_<lever>,1,$E$<bar>)` where E<bar> = 1/2/3.
Rows (one formula across J:AP; H = lifetime `SUM(J:AP)` where it means something):
- **Timing**: years since COD year · construction share by year (wind J252:J253 with NTP, COD + Case_COD_Delay) ·
  operating share (COD → COD + life) · COD-year flag · repower-year flag · operating years elapsed.
- **Energy**: generation = MWp × yield × (1 + Case_Yield) × availability × (1 − degradation)^years × operating share ·
  sent to the BESS = MIN(generation, charge-from-PV per MW × MW × operating share) · exported directly = generation −
  sent · discharge = discharge per MW × MW × PV-covered share × operating share (PV-covered share = MIN(1, sent × RTE
  ÷ (discharge per MW × MW)), one row) · full cycles · BESS retired flag (repower flag 0 and year ≥ repower year → 0
  for all BESS rows).
- **Revenue** (k€): PV direct = exported × price received direct ÷ Thousand · PV stored = sent × price received stored
  ÷ Thousand · GoO = generation × GoO price × index ÷ Thousand · BESS = MW × BESS revenue received × (1 − retired) ·
  total.
- **Opex** (k€, real × index): O&M + cable O&M (per MWp) · balancing = generation × 5 × (1 − PPA flag) ÷ Thousand ·
  lease · BESS opex × (1 − retired) · HIPA + innovation on revenue · total · **EBITDA**.
- **Capex** (k€): construction = total uses × construction share × (1 + Case_Capex) · early works at NTP · augmentation
  = MWh × capex × index at the repower year × repower flag · total.
- **Tax** (unlevered = the only tax here): depreciation PV/grid/development 30 y from COD, BESS 10 y, augmentation
  10 y · tax base = EBITDA − depreciation · loss corkscrew (50 % cap, lapse 5 y after NTP, flagged simplification) ·
  CIT · Robin Hood × flag.
- **Cash flows** (k€): project before tax = EBITDA − capex − pre-COD outlays (from the dated block by year,
  `SUMPRODUCT((YEAR(dates)=year)×amounts)`) · project after tax · **buyer at RtB** (flows from the exit year: pays
  nothing here; the price enters only in E) · DF to the exit year end `=(1+target)^((exit − year end)/365)` · **implied
  RtB EV** (E) = `SUMPRODUCT(post-exit after-tax flows × DF)` floored at 0 with flag · buyer IRR at our price (E) =
  XIRR on [−all-in price at exit; flows] · the same for COD (flows from the COD year, DF to COD) · FEP hold flows =
  project after tax, with the PSH buy-out at the exit date · PSH hold flows (apport → buy-out) · FEP COD-route flows
  (contributions to COD, proceeds at COD via the JV split).
- **Returns** in E/F: XIRR on the annual flows with the year-end dates (J5:AP5) plus the dated pre-COD flows merged
  through the RtB dated block (the RtB route IRR is XIRR on the dated list only); MOIC; XNPV at `Discount_Rate` to the
  acquisition date.

### 3.5 `Checks` (→ `Master_Check`, every row `=N(ABS(x)>Tolerance)` or a flag, per case in E:G, H = sum)
Levers valid (hours ∈ {2,4} when built, basis text valid, PPA years ≤ life) · energy: sent + direct = generation;
discharge ≤ sent × RTE; direct export ≤ MVA × 8,760 × operating share (annual, information) · BESS rows 0 when not
built · revenue: PPA price received = PPA price inside the tenor · JV split sums to cash available (RtB and COD) ·
FEP + PSH totals = consideration + reimbursements · tax losses ≥ 0 · XIRR numeric (every IRR cell) · implied price ≥ 0
flag · lifetime project CF = Σ annual summary · `Master_Check` = sum; mirrored at Dashboard H3.

## 4. Timeline
Annual, calendar years 2025 → 2063 (acquisition Nov-2025; COD 31-Mar-2028 + 35 years → 2063), year-end dates.
Pre-COD cash (SPV price, deposits, DSA draws, early works, pre-RtB development, exit proceeds) is **dated** in the
RtB block and mapped to the annual grid by `SUMPRODUCT` on `YEAR(date)`. Construction capex straight line NTP → COD by
calendar share. COD delay (Low) shifts COD, the operating window and the PPA window together. First full operating
year = 2029.

## 5. Revenue logic (formulas in words; `Revenue` bar 3 holds the prices, `Calc` the volumes)
### 5.1 Energy per case
PV only: everything is exported directly (no "sent", no discharge); the direct-export price of the **DC 2 h block**
would overstate a PV-only plant's capture (the block's PV is backed by a battery). **Use Aurora's standalone-solar
capture if the databook has one; otherwise the 2 h block's direct price with a flagged note "PV-only capture proxy"**
– Opus checks `Aurora_Q2_26…xlsm` `Central prices` for a solar capture row first (SPEC_v1 §6.2 found none in the
flexible-energy file; the proxy is then the default, flagged).
PV + 2 h / 4 h: Aurora DC co-located block of that duration, per MW of BESS: sent to the BESS = MIN(generation, charge
from PV per MW × 50 MW); discharge at the PV-covered share (no grid charging, user-confirmed); the PV-only factor
of **that block** (2 h: 0.675 in 2027 / 0.786 average; 4 h: 0.779 / 0.941) – never the 4 h factor on the 2 h block.

### 5.2 PV price
- **PPA years (COD → COD + 5 y):** pay-as-produced **on all PV output, directly exported and stored alike** (MET-style
  PV-meter settlement; Fable finding 5 in `ASSUMPTIONS.md` §2.4). Price = `Dash_PPA_Price` × (1 + indexation)^(year −
  COD year). The lever is **nominal €/MWh as offered** (MVM 45 is a nominal fixed price), so no euro index on it;
  the note says so. Balancing cost **not** charged in PPA years (the offers include it). GoO still sold (1.75 real).
- **Merchant years:** direct export at the block's direct-export price × euro index; stored energy at the block's
  transfer price × euro index (the counterpart of the −s+20 inside the storage total: never price stored energy at
  the direct price). Balancing 5 €/MWh real on generation. Case_PV_Price (±20 %) applies to merchant years only.

### 5.3 BESS revenue (k€/MW/yr, nominal)
- **Merchant (default):** storage total of the block × (PV-only factor + aFRR-down/FCR share × `In_AFRR_Down_Share`)
  × euro index × (1 − haircut) × (1 + Case_BESS). The stack is Aurora's fixed dispatch: it does not re-optimise when
  prices or the PPA change (note on `Revenue`). A PPA on PV does not change the BESS uplift (the uplift is net of
  the PV energy's export value already).
- **Tolling (lever):** MET full tolling 49.7 k€/MW/yr from 01-Jan-2028 for 3 years (fixed fee; applies as offered to a
  90 MWh battery – on the 4 h column the fee is **not** scaled, flag "offer ≠ built configuration"), then merchant.
- Repower: at the block's repower year (2048) augmentation capex = MWh × capex × index; `In_Repower` = 0 retires the
  battery instead (all BESS rows 0 from that year).

## 6. Capex, opex, tax, JV, strategies
- **Capex to COD** per case = PV 350 k€/MWp × 63 + cable 3,700 + substation 9,500 + land 620 + BESS 50 × h × 110 +
  E48/E49 725; × (1 + Case_Capex). Unlevered: equity = uses. No IDC, no fees, no DSRA.
- **Cash to RtB** (dated, shared): as SPEC_v1 §7 and `ASSUMPTIONS.md` #53–54, #80–82, #85–88: SPV price 549 (JV sheet;
  the 150 "connection rights" of v3 is the reconciliation value), transaction costs 200, DSA 945 by milestone, early
  works 785 (four lines), pre-RtB development 3,182, deposits/guarantees in HUF at FX 400 with guarantee fees; which
  lines the buyer reimburses or returns at the exit per the flags. 12 % act/365 return on FEP's loans and reimbursable
  costs, 31-Dec capitalisation (closed-form factor), from each payment date to the exit.
- **Opex**: §3.4; HIPA and innovation on revenue (placeholders).
- **Tax** (SPV, unlevered): CIT 9 % on EBITDA − depreciation, loss cap 50 %, lapse 5 y after NTP (one date, flagged),
  Robin Hood off (flag). Sale: share deal, participation exemption → gain exempt; FEP pays 9 % on the 12 % return as
  interest income; `In_Gain_Taxed` = 1 taxes the gain at 9 %. Pre-tax rows = the same flows without CIT.
- **JV split at the exit** (RtB; same mechanics at COD on EV − sale costs): cash available = price × MWp + reimbursed
  principal → sale costs → FEP loan principal → 12 % return → SPV price item (549, to FEP) → distributable → **FEP share
  = 60 % if price per MWp ≤ 90 (`<=`), else 50 %**; PSH = rest + its DSA; `MIN(due, cash left)` at each step. Hold: FEP
  buys PSH out at the exit date at PSH's RtB entitlement (what PSH would have received at the RtB price used), label
  "PSH buy-out at RtB entitlement, pending JV confirmation".
- **Strategies**: RtB sale (dated list, FEP XIRR/MOIC; buyer unlevered after/before tax from the exit on the annual
  grid, all-in price = premium + reimbursements); COD sale (FEP funds everything to COD; EV at the COD year end,
  consistent with the annual grid and the DF; FEP proceeds via the split; buyer after/before tax
  from COD); hold (FEP keeps the project; PSH bought out). Implied prices closed-form: implied EV = Σ post-exit
  after-tax flows × DF at `Case_Buyer_IRR` (unlevered, after tax); the implied **premium** per MWp = (EV − reimbursements)
  ÷ MWp, floored at 0 with a flag; RtB/COD price levers 0 = use implied.

## 7. Base / Low / High (Dashboard §3.1 item 3)
Low: PV price −20 % (merchant years), BESS −20 %, yield −2 %, capex +10 %, COD +12 months, buyer target 9 %.
High: +20 %, +20 %, +2 %, −5 %, 0, 7 %. Six levers, nothing else. The 18 key-assumption sets, `RunAllScenarios`, the
stored-results table and `Model_Checks` per set are gone; the `Scenarios.bas`/`Solver.bas`/`GoalSeek.bas` modules are
not packaged (plain `.xlsx`).

## 8. Checks
§3.5. Lint (`fepm audit`) clean incl. `g-chain`, `g-long`, `g-note`, `c-sheet`, `x-xirr0`; LibreOffice recalc;
`Master_Check` = 0 in Base, Low and High (the auditor switches `Dash_Case` on a scratch copy) and with each lever
flipped (BESS basis Tolling, repower 0, RtB price 0, COD price 0).

## 9. Reconciliation (`reconcile.py`, `out/reconciliation.md`)
The v3 shadow (`data/old_engine/old_zhlp.py`) stays and ties to `old_targets.csv`. New bridge from v3's HTM/RtB/buyer
targets to the **PV + 4 h column, Base, merchant basis**: v3 → SPEC_v1 §13 fixes (2, 3, 5–8, 13–18) → 63 MWp / 1,265
→ early works 4 lines / E48–E49 to capex → SPV price 549, DSA 945 → annual timeline (year-end dates; the RtB route
keeps exact dates) → HIPA → PPA 5 y at 45 → base. Each step a quantified row (FEP hold IRR, project IRR, RtB profit,
RtB XIRR, implied RtB premium). An independent Python shadow of the new PV + 4 h block ties to the workbook at
machine precision. No v3 column in the workbook (§0).

## 10. Opus tasks (`model-builder`; `model-auditor` after each phase)
0. **Engine**: revert SUB_FILL to `E6F4EC` in `fepm/xl/styles.py` (and the §5 table of `MODELLING_STANDARDS.md`);
   confirm `Model(periodicity="annual")` gives calendar years with year-end dates; add a `Block` helper that writes the
   same row set n times on `Calc` with the case column as a parameter and returns row keys per block; a `DatedBlock`
   helper (amounts E:G, dates J:L, accrual factor, XIRR); Dashboard lever rows with `Dash_*` names and list
   validation; sub-header style; `.xlsx` output when no macros; lint exception for blue cells on the Dashboard lever
   rows only. `pytest` green; the Bana/Bora builds untouched.
1. **Inputs** §3.2 with the Source column; the dated tables (early works, deposits, pre-RtB development) as flat rows.
2. **Revenue** §3.3; `data/aurora_extract.py` writes the 2 h and 4 h series incl. the PV-only factor and aFRR share
   per block (exists) plus a standalone-solar capture row if the databook has one.
3. **Calc** §3.4, three blocks; §5 exactly: PPA on all PV output, balancing merchant only, transfer price on stored
   merchant energy, PV-only factor of the block, retire/repower, JV step `<=` 90.
4. **Dashboard** §3.1; the dated RtB block in the group; Base/Low/High.
5. **Checks** §3.5; `fepm audit`; LibreOffice recalc of Base/Low/High and the lever flips.
6. **Reconciliation** §9; `out/reconciliation.md`.
7. Delete the v1 build artefacts that no longer apply (`out/Zalahalap_Model.xlsm`, `_recalc.xlsx`); commit "Rebuild
   Zalahalap in the wind-model format: PV only / 2 h / 4 h, unlevered, 5-year PPA".
Report: sheet/row counts; the Base results per case and strategy; the Low/High spread; the bridge; anything infeasible.

## 11. Open items (placeholders stay flagged; answers refine the defaults)
1. PPA price basis: 45 €/MWh nominal flat (MVM) is the default; E.ON indexed or MET 43.96 can be typed into the lever.
   Confirm the MVM scope (PV meter vs POD; BESS inside the 45?).
2. PV-only capture: Aurora standalone-solar row if it exists, else the 2 h block's direct price as a flagged proxy.
3. BESS capex 110 k€/MWh (2026 quote for 2 h and 4 h); repower 2048.
4. aFRR-down/FCR share deliverable from PV state of charge (0; worth 12–17 % of the stack).
5. Dates (RtB 31-Dec-26 is two months away; COD Mar-28; 9-month build).
6. Yield 1,265 vs the offerors' 1,199–1,245; yield report.
7. Buyer target IRR 8 % (9 % in Low); RtB price basis premium 90 vs all-in.
8. Tax: Robin Hood, depreciation lives, participation exemption, HIPA rate.
9. JV: SPV price 549, DSA 945, the unmodelled 562.5 k€ loan, 120 vs 90 k€/MWp basis, COD-sale and hold treatment of PSH.
10. Early works treatment as a whole (user: to be revisited).
11. Financing: out of scope for now; a term sheet brings it back (SPEC_v1 §8 debt design, Bana method).
12. Sale costs 0 %, route-to-market fee 0 %, GoO 1.75 (Fable recommends 1.5 %, 2 %, 1.0).
13. **Fix round (review of a4d876f).** Clipping is a per-case Dashboard lever: PV only 2 % (v3, fix 5; 63 MWp on
    49.9 MVA with no battery to absorb it), BESS cases 0 % (DC-coupled) – placeholder. BESS fixed opex is split into
    18.70 k€/MW + 6.07 k€/MWh a year (real 2025), derived from Aurora's own table that v3's 30.8 cites (Aurora Inputs
    rows 37–39, 2028 entry: 2 h 30.83, 4 h 42.96 €/kW): the 2 h battery keeps 30.8, the 4 h battery carries 43.0
    (v3 applied the 2 h figure to 4 h). Aurora's 10 % co-location opex saving is not applied – confirm with an O&M
    quote. The PV-only price proxy (2 h block) is below v3's 65 % × baseload from 2031; the HU solar capture series is
    requested from Aurora. Tolling starts at the later of the offer start and COD + 1 day and runs its full tenor.
14. **v2.1 (§12), as built.** Timeline "to 2026" | 2027 … 2064 (to 2064, not 2063, so the Low case's COD +12 months
    keeps its 35-year life). Debt sized on total uses to COD incl. land (§12.2 lists PV, BESS, grid, E48/E49). IDC on
    2027 draws at mid-year (`Draw_Timing` 50 %), capitalised; interest from the COD year paid in the annuity (the
    first, 31-Dec-2028, is a full year). Repayment profile not a lever (annuity only). IDC is not tax-depreciated
    (small). The levered buyer IRR at COD is n/a when the price is below the debt at COD (Low: all cases). The COD
    sale is shown unlevered and levered (FEP repays the debt from the price before the split). Open: a term sheet
    (DSCR sizing, fees, DSRA); PV only has a minimum DSCR below 1 in Base.

## 12. v2.1 – timeline rebase and the levered case (user, 04-Nov-2026)
User, verbatim: "now it is the end of 2026 so make the model based on that the costs will arrive in 2027 and
production strats 2028 march" and "also make a levered case with 60% gearing 5% cost 20 years amortisation".

### 12.1 Timeline rebase (replaces §4)
- Model date 31-Dec-2026 = RtB. Everything paid before it (SPV price, DSA draws, early works, pre-RtB development,
  deposits, guarantee fees) is **spent to date**: it keeps its actual dates in the dated RtB block (XIRR needs them)
  and sits in one annual column "to 2026" on the grid. No cost is moved; only the first grid column changes.
- Grid columns: "to 2026" | 2027 | 2028 | … | 2063, year-end dates. First full operating year 2029 (unchanged).
- NTP = 01-Jan-2027. **All construction capex in 2027** (one `In_Capex_Year` lever: 2027; the straight-line NTP→COD
  spread goes). E48/E49 (725 k€) at NTP, i.e. Jan-2027.
- COD = 01-Mar-2028; production from March 2028 (10/12 of a year in 2028). PPA window Mar-2028 → Feb-2033. The
  COD lever and the Low case's +12 months keep working (COD Mar-2029, capex still 2027 – say so in the note).
- RtB sale closing stays a lever (default 31-Mar-2027) so the 12 % accrual runs from the actual dates to the exit.
- Checks: the first-column sum = the dated block's sum to 31-Dec-2026; capex sum = `In_Capex_Year` column.

### 12.2 Levered case (new financing block; unlevered rows stay)
- Levers on the Dashboard, per case: `Gearing` 60 % (0 = unlevered), `Debt rate` 5.0 % all-in, `Tenor` 20 years,
  `Repayment` = annuity (level debt service; a note says straight-line is the alternative). Pending a term sheet:
  flagged "user instruction 04-Nov-2026, not a bank term".
- Debt = Gearing × construction capex to COD (PV + BESS + grid + E48/E49; not development, not deposits), drawn pro
  rata with the capex (2027). IDC at the rate, capitalised into the debt at COD. No fees, no DSRA, no sweep (one note).
- Repayment: 20 annual annuities from the first year-end after COD (2028, prorated? No – first instalment 31-Dec-2028,
  full annuity; keep it simple and say so). Interest on the opening balance; interest deductible for CIT; the HIPA base
  unchanged.
- Rows per case on `Calc`: draw, IDC, opening balance, interest, principal, closing balance, debt service, CFADS,
  DSCR. `Dashboard`: min and average DSCR (information, no check), debt at COD, and the levered rows:
  - Hold: FEP levered IRR pre/post-tax, levered MOIC, levered profit (equity = capex − debt).
  - COD sale: buyer levered IRR at the implied price (buyer assumes the debt: equity = price − debt at COD).
  - RtB sale: buyer levered IRR at 90 k€/MWp (buyer raises the debt on the capex).
  The implied prices stay unlevered (8 % unlevered target), as §6; FEP's COD/RtB proceeds are unchanged by the debt
  except the COD-sale case where FEP repays the debt from the price before the JV split (price − debt at COD).
- Low/High: unchanged; gearing is not a sensitivity.
- Checks: debt closing balance = 0 after the last annuity; sources (equity + debt + IDC) = uses; CFADS ≥ debt service
  flagged as information (a DSCR < 1 is not an error in an unlevered-target model).

### 12.3 Opus tasks
1. Rebase the timeline per §12.1 (Inputs dates, the "to 2026" column, capex year lever); keep the dated RtB block.
2. Financing block per §12.2 on Calc and the Dashboard; the levered rows next to the unlevered ones.
3. Audit as §8/§10.5 plus: gearing 0 reproduces the unlevered figures exactly; the Base/Low/High and lever flips.
4. reconcile.py / shadow.py: add the debt rows to the shadow; refresh out/reconciliation.md (timeline step).
5. ASSUMPTIONS.md: dates and financing rows updated; SPEC §11 open items. Commit and push.
