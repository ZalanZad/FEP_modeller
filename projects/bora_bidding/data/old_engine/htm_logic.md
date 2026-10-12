# Old engine – build, sale at COD and hold to maturity: logic map

Source: `reference/Bora_scenarios_meeting.xlsx` (read-only). Sheets: `HTM` (annual engine), `Deposits` (securities and
connection-fee schedule), `Tender` (strategy → MW, grid-bottleneck loss, COD price effect), `Inputs` (case columns
G:P, column E = active case), `Curves`, `Calc` (data table), `Run_Map`. This file describes the workbook as it is;
`docs/reference_analysis/bora_models.md` describes an older version (200 MVA, other strategies) and is superseded here.

Shadow: `old_htm.py` (stdlib only) reproduces every HTM / COD row of this map. Inputs: `old_inputs_htm.json`,
`curves.csv`. Reference values: `lo_runs_htm.json` (LibreOffice live values per run + the values saved in the workbook).
Units k€ unless stated. Sign: + = inflow to the party, − = outflow (HTM!A2).

## 0. Result and how it was checked

| | S1 current bid | S2 = BORA 1 | S3 = BORA 2 |
|---|---|---|---|
| Installed wind / BESS (MW) | 130 / 30 | 151 / 51 | 140 / 40 |
| Grid-bottleneck loss (Tender J45) | 1.154 % | 1.689 % | 1.429 % |
| All-in cost to COD EU / CN (m€, Bora rows 87 / 104) | 236.42 / 169.49 | 273.86 / 196.32 | 255.49 / 183.50 |
| COD sale EU: FEP IRR / MOIC (Bora 63–64) | 31.308 % / 1.558 | 31.666 % / 1.531 | 31.901 % / 1.546 |
| COD sale CN: FEP IRR / MOIC (Bora 65–66) | 40.863 % / 1.801 | 42.046 % / 1.774 | 41.935 % / 1.787 |
| HTM EU: FEP IRR / MOIC (Bora 68–69) | 10.275 % / 6.35 | 10.887 % / 6.93 | 10.567 % / 6.61 |
| HTM CN: FEP IRR / MOIC (Bora 70–71) | 13.344 % / 8.02 | 14.363 % / 8.91 | 13.806 % / 8.40 |
| HTM EU: FEP NPV @ 8 % / min DSCR | 36,544 / 1.21 | 49,729 / 1.19 | 42,873 / 1.20 |

Match (`python3 old_htm.py`, runs 2–5, 7–10, 12–15, 23–25 metrics each): every value (profits, NPVs, MOIC, exposure,
DSCR, all 13 cost-to-COD lines, COD price) within **2e-15 relative** of both the saved Excel data table and LibreOffice;
HTM IRRs within 2e-14; COD XIRRs identical to LibreOffice and within **5.3e-9** of the saved Excel values (Excel's own
XIRR tolerance). All ~45 annual HTM rows and block-17 rows agree year by year with LibreOffice (tolerance 1e-7).

**LibreOffice caveat.** With the data table `Calc!G306:DH330` present, LibreOffice leaves some cells in a stale
data-table state after `Calc!E10` is changed (e.g. `HTM!G23`, `Inputs!E230`, `HTM` row 340 showed S4 COD-CN values in an
HTM run, while the IRRs were right). Clear `Calc!G306:DH330` in the private copy before recalculating; then every
`Calc` row-305 value equals the saved table. `lo_runs_htm.json` was produced that way.

## 1. Runs, routes and what differs between them

Run code = strategy × 5 + route (Calc!E10 → `Inputs!E6` = INDEX(Calc!E306:E330, E10), `Tender!E5` = INDEX(Calc!D306:D330, E10)).
Routes → cases: COD-EU 7, COD-CN 8, HTM-EU 2, HTM-CN 3. BORA 1 = runs 7, 8, 9, 10; BORA 2 = 12–15; S1 = 2–5.

Inputs that differ between the four build cases (checked cell by cell over Inputs!E6:E298 in LibreOffice):

| Input | HTM-EU (2) | HTM-CN (3) | COD-EU (7) | COD-CN (8) |
|---|---|---|---|---|
| Turbine supply Inputs E162 (k€/MW) | 950 | 500 | 950 | 500 |
| Substation + 132 kV bay Inputs E167 (k€) | 12,000 | 10,800 (= G167 × (1 − Tender E29 10 %)) | 12,000 | 10,800 |
| Gearing Inputs E217 | 60 % | 50 % | 60 % | 50 % |
| Deemed exit Inputs E228 | 1 | 1 | 0 | 0 |
| COD price Inputs E230 (k€/gross MW) | 0 | 0 | 2,450 × (1 − J45) | 2,100 × (1 − J45) |

Everything else (yield, availability, degradation, O&M, PPA, merchant curve, debt terms, tax) is identical for EU and CN.
The tender score differs (ÖSZ-9 EU equipment) but has no economic effect. Versus case 1 (RtB) the build cases also set
merchant share E190 = 0 (100 % PPA for 10 years) and gearing 60 / 50 % instead of 55 %.

Strategy inputs (Tender E:I ← Calc E:I ← Bora D:H): municipal one-off E233 (S1 20, S2 50, S3 50 k€/MW), community
E234 (2 / 3 / 2 €/MWh), overbuild (30 % / 51 % / 40 %), BESS ratio of the net MVA (30 % / 51 % / 40 %, basis 0, 4 h),
performance-security multiple E90 (1.5 / 1.0 / 1.0). Net MVA = Calc!E74 = 100.

## 2. Timing conventions

- Annual columns H:AQ = calendar years 2026–2061 (`HTM!H5 = YEAR(Inputs!E27)`, +1). Row 4 `t = year − 2026` is the
  discounting exponent; AR is used only in block 17 (sale on the COD date).
- Dates: investment `E27` 31-Jul-2026; COD input `E298` = Bora!N16 30-Sep-2031; `E55 = (YEAR(E298) − YEAR(E27))·12 + MONTH(E298) − MONTH(E27) − E152 − E153` = 35;
  RtB `E56 = EOMONTH(E27, E55)` 30-Jun-2029; RtB-sale lag `E28` = months from RtB to Bora!N17 = 6 → `E57` = 41,
  sale `E58` 31-Dec-2029; NTP `E154 = EOMONTH(E56, E152 = 3)` 30-Sep-2029; COD `E155 = EOMONTH(E154, E153 = 24)` 30-Sep-2031;
  first full year `E156` 2032; last year `E158 = YEAR(COD) + 30` = 2061.
- Row 7 construction share = MAX(0, MIN(31-Dec-y, COD) − MAX(1-Jan-y − 1, NTP)) / (COD − NTP) (sums to 1; 2029 12.6 %, 2030 50 %, 2031 37.4 %).
- Row 14 construction fraction of the calendar year = same numerator / days in year.
- Row 10 operating fraction = (12 − MONTH(COD))/12 in the COD year (0.25) + 1 for 2032…2060 + MONTH(COD)/12 in 2061 (0.75): exactly 30 years.
- Row 9 operating-year index = year − 2031 for operating years ≥ 2032, 0 in the COD year. Row 8 full-year flag. Row 11 decommissioning flag (2061).
- Discount factors rows 12/13 `1/(1+8 %)^t`, end-of-year, t = 0 in 2026 (no mid-year).

## 3. Strategy → physical project (Tender section 2)

- Installed wind `Tender!row40 = ROUND(MVA·(1 + overbuild), 6)` → `Inputs!E297` → `E29 = E18·E297/E281` (E18 = net MW 100).
- BESS `row41 = ROUND(ratio·(basis = 0 → MVA), 6)` → `E20`; MWh `E22 = E20·E21` (4 h).
- Grid-bottleneck loss `J45 = ob/(1 + ob)·Tender!E27` with ob = installed/MVA − 1 and E27 = Bora!N18 = 5 %: 5 % of the
  energy of the overbuilt MW is lost. Applied to generation (HTM E36) and to the COD price (`J46/J47 = price·(1 − J45)`).

## 4. HTM block 1 – development phase (rows 16–24), all × `E15 = Inputs!E226` (= 1)

| Row | Formula (column y) | Source |
|---|---|---|
| 16 devex incl. land | −Σ_{months m in y}(MOIC!row18 + MOIC!row20) | interface (Devex_FEP) |
| 17 signing fee + DD | −(E24 + E25) in YEAR(E27) | 250 + 150 |
| 18 cash securities | −Σ_{Deposits lines, YEAR(K)=y, R ≠ "Connection fee"} P·M | participation fee 16.4 + TSO bond 1,232.9 in 2026 |
| 19 cash securities returned at COD | in the COD year: Σ_{R ∈ Security, Refund} P·M − P11·M11 | 0 (bid / PB are guarantees) |
| 20 DSA in cash at COD | −MOIC!G25·E20 (Inputs E227 = 0) | 0 |
| 21 deemed exit – FEP profit forgone | −E21·MOIC!G182 in YEAR(E58); E21 = Inputs E228·Tender J48 | interface |
| 22 deemed exit – cash to A.RES | −E22·(MOIC!G154 − MOIC!G151) in YEAR(E58) | interface |
| 23 guarantee fees | −[Σ_{months in y} MOIC!row79 + E23·(−MOIC!G75)·days(MAX(31-Dec-(y−1), E58) → MIN(31-Dec-y, COD))/365 + Deposits!M17·(1 − P17)·days(K17 → COD)/365] | E23 = rate in the sale month (6 %) |
| 24 total | Σ rows 16–23 | |

MOIC row 79 (monthly) = (opening + issued guarantees)·6 %/12 for months 0…41; guarantees issued = Deposits O·(1 − P) in
the line's grid month L; all released in the sale month (−G75 = 6,993.6 k€ for S2). So guarantee fees run on the bid
bond from Oct-2026 and the (netted) performance bond from Dec-2026, at 6 % flat (Inputs E113, E114 = 1), to COD.

## 5. Deposits lines used (Deposits rows 7–18; mHUF → k€ at `M = J·1000/E11`, E11 = 365)

| Line | Gross H | Netting I | Date K | Cash P | Type R | S2 k€ (M) |
|---|---|---|---|---|---|---|
| 7 participation fee | E81·E82 = 3·2 | – | E95 30-Oct-26 | 1 | Fee | 16.4 |
| 8 bid bond | E76·E91, E91 = E16·(E29/E18)·E75 + E22·E89 (committed MW and MWh) | – | E95 | 0 | Security | 1,398.7 |
| 10 performance bond | E77·E90·E92, E92 = E29·E75 + E23·E89 | −H8·E107 (bid bond credited) | E96 + 17 d = 30-Dec-26 | 0 | Security | 5,594.9 |
| 11 TSO connection-request bond | E83·E18 = 4.5·100 | – | E97 30-Dec-26 | 1 | Security | 1,232.9 |
| 12 direct connection fee | E84·E18 = 5.0835 mHUF/MW·100 (Annex 1 lump sum ÷ 100 MVA) | – | E99 8-Jan-29 | 1 | Connection fee | 1,392.7 |
| 18 indirect connection fee | E85·E18 = 73.071·100 | −J11 (TSO bond credited) | E99 | 1 | Connection fee | 18,786.7 |
| 9, 13–17 refund / supplementary | 0 (E19 failed MW = 0; E78 = 0 per GYIK 70) | | | | | 0 |

`O = M·(E109 if grid month L > E57 else 1)` (instalments after the RtB sale dropped from MOIC; HTM rows 18/28 use M).

## 6. HTM block 2 – capex (rows 27–32) and the cost-to-COD lines

- `E173 = (Σ E162:E166 + E169 + E170)·(1 + E171)·E29/E18 + E167·(1 + E171)/E18 + E168·E20/E18` (k€ per net MW; turbine/BoP/owner's/land on installed MW, 5 % contingency, BESS 200 k€/MW without contingency).
- Row 27 = −E173·E18·row7; row 28 connection fee = −Σ_{YEAR(K)=y, R="Connection fee"} M (not × E15; 8-Jan-2029); row 29 municipal = −E233·E18 in the COD year (on net MW); row 30 = 27 + 28 + 29.
- Cost to COD (Calc CQ:DB, Bora rows 75–89 EU from runs 4/9/14…, rows 92–106 CN from runs 5/10/15…): CQ = −G16, CR = −G17,
  CS = E162·1.05·E29, CT = (E163:E166 + E169 + E170)·1.05·E29, CU = E167·1.05, CV = E168·E20, CW = −G28, CX = −G29, CY = −G30,
  CZ = −(G18 + G19), DA = −G23, DB = G78 − G85 (IDC + upfront + commitment fees); all-in = CQ + CR + CY + CZ + DA + DB.

## 7. Block 3 – generation (rows 35–39)

`E35 = E29` (installed MW, overbuild included); `E36 = E179·(1 − E180)·(1 − E181)·E183·(1 − Tender J45)` = 2,460·0.98·(1 − J45)
(S2: 2,370.1 h); degradation row 37 = (1 − 0.3 %)^MAX(idx − 1, 0) while operating; row 38 GWh = E35·E36·row37·row10/1000.
Availability is inside the 2,460 net P50 hours (E180 = 0); curtailment 2 % (E181).

## 8. Block 4 – prices (rows 42–49)

- Merchant row 42 = `Curves!K(y)·E184`, K = Aurora Q2-26 Central wind capture, real 2025 × 1.02^(y − 2025) (Curves E19, G19);
  after 2060: K2060·1.02^(y − 2060); 2026 not in the curve → falls back to the flat estimate (rows 43–44), never used.
- PPA row 47 = E191·(1 + E192)^(y − YEAR(COD)) = 91.6 €/MWh (Bora!N10) in 2031, +1 % p.a.; PPA share row 48 = 1 − E190 = 100 % when
  1 ≤ idx ≤ E193 (10), i.e. 2032–2041. The COD-year stub (Oct–Dec 2031) and 2042+ are merchant.
- Row 49 blended price (memo).

## 9. Block 5 – revenue, opex, EBITDA (rows 52–70); escalation index row 55 = 1.025^(y − YEAR(COD)) while operating

Revenue 52 = GWh·merchant·(1 − share), 53 = GWh·PPA·share. Opex (negative): balancing 4 €/MWh·GWh·esc (56), grid 1 €/MWh (57),
MEKH 0.1 % of revenue (58), community E234 €/MWh·GWh, not escalated (59), O&M 12, BoP 3, insurance 2.5, AM 1.5, property tax 3
k€/MW·E35·esc·opfrac (60, 61, 63–65), land lease 3 k€/MW·E35·1.025^(y − 2031)·(opfrac + constr. fraction) (62, from NTP),
admin 150 k€·esc·opfrac (66), contingency 5 % on rows 56–58 + 60–66 (67; not on community and HIPA), HIPA 2 % of revenue (68).
EBITDA 70 = 54 + 69.

## 10. Block 6 – senior debt (rows 75–89)

- Size `E75 = E217·(−G30)` (total capex incl. connection fee and municipal; 60 % EU, 50 % CN). Drawdown 77 = E75·row7.
- IDC 78 = (open + draw/2)·5 %·[constr. fraction/(constr. fraction + opfrac) in the COD year, else 1]; capitalised (E221 = 1) row 79.
- Interest 81: COD year (open + draw + IDC)·5 %·opfrac; later open·5 %.
- Annuity `E82 = PMT(5 %, 10, −(G77 + G79))`; principal 83 = MIN(MAX(E82 − interest, 0), open) for idx 1–10.
- Fees 85 = −1.75 %·E75 in the first drawdown year − 0.9 %·MAX(0, E75 − open − draw/2)·constr. fraction (commitment fee).
- Sweep 87 (idx 1–5) = MIN(MAX(EBITDA + tax + scheduled DS + DSRA flow, 0)·50 %, balance after scheduled principal); prepays, annuity amount unchanged → shorter final payment (BORA 1 EU: balance 6.1 m€ left for 2041 instead of a full annuity).
- DSRA target 88 = 6/12·MIN(E82, (open + draw + IDC − principal)·1.05) while operating and idx + 1 ≤ 10; flow 89 = −Δtarget (funded in the COD year by equity, S2 EU 10.4 m€, released 2040–41).
- Closing 84 = open + draw + IDC − principal − sweep. Check G158: repaid by end of tenor.

## 11. Block 7 – tax (rows 90–101)

- Depreciation profile 90: opfrac in the COD year, 1 for idx 1–19, 1 − 0.25 for idx 20 (sum 20). Base `(−G30 + G79 + Deposits!P11·M11)/20`
  = (capex incl. connection fee and municipal + capitalised IDC + TSO bond)/20.
- Taxable 93 = EBITDA − depreciation − interest(81) + fees(85) + devex(16). Not deductible: signing fee, guarantee fees, deemed exit.
- Losses: used 95 = MIN(opening, MAX(taxable, 0)·50 %) (E214); closing 96 = opening + MAX(−taxable, 0) − used; CIT 98 = −9 %·(MAX(taxable, 0) − used); Robin Hood 99 off.
- Unlevered proxy 101 = −MAX(EBITDA − depreciation, 0)·9 % (no interest shield, no loss relief).

## 12. Blocks 8–10 – cash flows and FEP share

- Unlevered 109 = 24 + 30 + 70 + decom (−30 k€/MW·E35 in 2061) + TV (0); post-tax 110 = 109 + 101.
- Equity 123 = 24 + 30 + draw(77) + (85 − 80) + EBITDA + (86 − 87) + DSRA(89) + tax(100) + decom/TV. Check G171 ties it to 109 + debt flows.
- FEP 129 = 123·E128·(1 − row383) + row384; `E128 = 1 if Inputs E228 = 1 else E33 (70 %)`; rows 383–384 community purchase right off (Tender E370 = 0).

## 13. Block 11 – HTM results (the numbers on the Bora sheet)

- **HTM FEP IRR = `HTM!G139 = IRR(H129:AQ129)`** – annual IRR (not XIRR) on the FEP equity cash flow (levered, post-tax, incl.
  development phase and the deemed exit), t0 = calendar 2026 (flows treated as one per year). Run_Map D24 → Calc I → Bora rows 68/70.
  (G137 = IRR of row 123 is identical while E128 = 1.) An XIRR on mid-period dates differs by only ≤ 0.002 pp.
- MOIC `G147 = Σ positive annual row-129 flows / −Σ negative` (within-year netting). NPV `G140 = Σ row129/(1.08)^t`, t = year − 2026
  (2026 flows undiscounted, i.e. valued at the 2026 point, end-of-year convention). NPV/Equity G142 = G140 / PV of negative years.
  Profit G144 = Σ row 129 (30-year nominal sum); pre-tax Run_Map D22 = G144 − Σ row121·E128. Peak G146, project NPV G136 (row 110 at 8 %), min DSCR G126 = MIN((EBITDA + tax)/−DS) in debt years.

## 14. Block 17 – sale at COD (rows 333–380) and the COD price effect

- Flags: pre-sale 333 = year ≤ YEAR(COD); sale 334 = year = YEAR(COD) and E230 > 0.
- 335 = row24·333; 336 = row30·333; 337 = row77·333; 338 = row117·333; 339 = −E203·E35·1.025^(y − 2031)·row14·333 (pre-COD land lease, no contingency).
- Price 340 = 334·(E230·E29 + Tender!E225), E230 = Tender J46/J47 = 2,450/2,100·(1 − J45) per installed MW.
  `Tender!E225 = E30·Σ_y [−(E234 − 2)·GWh_y·(1 − 9 %)]·1/1.08^(MAX(0, 31-Dec-y − COD)/365)` – PV at COD of the extra community
  €/MWh versus S1 (S2: −3,569 k€; S3 0). Only the community item is passed through at COD.
- Debt repaid 341 = −334·(open + draw + IDC) in the COD year; sale costs 342 = −E26 (500); gain tax 344 = 0 (E231 = 0).
- 17b JV waterfall (all in column G, × (E230 > 0)): net proceeds G365 = 340 + 341 + 342 + 344;
  funding returned G366 = −[−Σ rows 335–339 + G17 + G23 + (1 − E71)·(G16 + (MOIC!G20 − MOIC!G124)·E15)] (all contributions
  except signing fee/DD, guarantee fees and the at-risk devex principal and participation fee);
  Stage-1 returns G367 = −(Deposits!AF22 + AF23 + AF24 + MOIC!G280) (6 %/12 % per line to COD with the interest pause, + devex/land
  MAX(20 %, 2×) multiple to COD); DSA G368 = −G375, `G375 = E38·MIN(E37 10, E36 2.5·YEARFRAC(E70, COD, 30/360))·E29` (= 10 k€/MW·installed MW);
  distributable G369 = G365 + G366 + G367 + G368; developer G371 = MAX(G368, 30 %·G369); FEP G370 = G369 − G371;
  FEP proceeds G372 = −G366 − G367 + G370.
- FEP flows: contributions 359 = Σ rows 335–339 (FEP funds 100 %), AR359 = G372 dated COD; 360 = 334·G372.
- **COD FEP IRR `G352 = XIRR(H359:AR359, H362:AR362)`**, dates row 362 = (MAX(1-Jan-y, E27) + MAX(MIN(31-Dec-y, COD), 1-Jan-y))/2
  (mid-point of the active part of each year, x.5 values **truncated to whole days by Excel/LibreOffice**), sale on COD.
  MOIC G353 = G360/−G359; NPV G354 = Σ row346/1.08^t (annual, t = year − 2026, sale in 2031 at t = 5 – a different date basis
  from the XIRR); profit G351 = Σ row 346; peak G355 = −MIN cumulative row 359.

## 15. Interface: cells that feed HTM / COD from the RtB engine

| HTM cell | Reads | Meaning | Shadow |
|---|---|---|---|
| row 16 | MOIC!H18:BP18 (← Devex_FEP!H5:H38 by month Devex_FEP!I), MOIC!H20:BP20 (← Devex_FEP!K, Inputs E67), MOIC!H5:BP5 | devex and land by month | LibreOffice vectors in JSON |
| rows 18, 19, 28, 92 | Deposits!K7:K18, M7:M18, P7:P18, R7:R18, P11, M11 | cash securities, connection fee, TSO bond | recomputed (§5) |
| row 20 | MOIC!G25 | DSA at the RtB sale (× E227 = 0) | 0 |
| rows 21, 22 | MOIC!G182; MOIC!G154, MOIC!G151; Tender!J48 | FEP RtB profit; developer share + DSA | LibreOffice values per strategy |
| row 23, E23 | MOIC!H79:BP79, H7:BP7, H78:BP78, G75; Deposits!M17, P17, K17 | guarantee fees to the sale, balance at the sale | recomputed (MOIC rows 72–79 logic) |
| E36, row 340 | Tender!J45, J46/J47 (via Inputs E230), Tender!E225 | bottleneck loss, COD price, COD price effect | recomputed |
| G366 | MOIC!G20, MOIC!G124 | land principal; unrecoverable cash (participation fee) | land from JSON; G124 recomputed |
| G367 | Deposits!AF22:AF24, MOIC!G280 | Stage-1 returns to COD | AF recomputed; G280 = 3,216.7 from JSON |
| rows 383–387 | Tender!E370, E372 | community purchase right | off (asserted) |

Note the loop across sheets: `MOIC!G182` (row 21) depends on `Tender!E219` (RtB price effect), whose net margin (Tender row 206)
reads HTM rows 10, 45, 47, 48, 55 of the **active case**. Under cases 2/3 the RtB price effect is therefore valued with the
100 %-PPA revenue, not with case 1 – see bug B1.

## 16. Inputs (Inputs column E; value from the BORA 1 HTM-EU run; strategy- and case-specific rows see §1)

| Cell | Input | Value (HTM-EU, BORA 1 run) | Unit | Inputs F note (abridged) |
|---|---|---|---|---|
| Inputs!E11 | EUR / HUF | 365 | HUF/EUR | Single FX for the whole model (also used for mHUF reporting outputs). |
| Inputs!E12 | Hungary corporate income tax | 0.09 | % | Applied to positive cumulative profit at sale, less any gain exempted by the participation exemption (E14). |
| Inputs!E16 | Gross capacity entered in tender | 100 | MW | Bid-stage items (participation fee, bid bond) scale with this. |
| Inputs!E17 | Conversion rate (capacity awarded and taken to RTB) | 1 | % | Post-award items and exit proceeds scale with successful capacity. |
| Inputs!E21 | BESS duration | 4 | Hours | Storage hours at rated power – linked to Tender J16 (active tender strategy; 4 h in every strategy, Munka1 0.25C) |
| Inputs!E24 | Signing fee to current owners | 250 | k€ | Paid at month 0. |
| Inputs!E25 | Transaction / DD costs to FEP | 150 | k€ | Paid at month 0. |
| Inputs!E26 | Sale costs at exit | 500 | k€ | Deducted from proceeds in the waterfall. |
| Inputs!E27 | Investment date (month 0) | 2026-07-31 | Date | Start of the monthly grid; every month steps by EOMONTH. |
| Inputs!E28 | Lag from RTB to sale | 6 | Months | RTB + 6 months = RtB sale (BORA Devex proposal; was 3). RTB 30/06/2029 → sale 31/12/2029 |
| Inputs!E33 | FEP effective share of distributable proceeds | 0.7 | % |  |
| Inputs!E34 | Developer effective share | 0.3 | % |  |
| Inputs!E36 | Developer success fee p.a. | 2.5 | k€/MW | DSA on installed MW per BORA_Devex – user 09-Oct-26: accrues pro rata to the sale on the installed wind MW (E29, BORA_Devex V37 = 2 500 €/MW × V6); settled at exit fro… |
| Inputs!E37 | Success fee cap | 10 | k€/MW |  |
| Inputs!E38 | Success fee live | 1 | Switch | Fee only exists in the JV case. |
| Inputs!E61 | Recoverable loan interest (refundable securities) | 0.12 | % p.a. | Compounded monthly on the drawn balance. |
| Inputs!E64 | Interest pause start (no accrual after this date until the pause end, row 73) | 2029-06-30 | Date | Interest pause 30/06/2029 – 01/01/2030: no accrual in between (user 09-Oct-26) |
| Inputs!E65 | Interest on bank-guaranteed securities (credited to FEP) | 0.06 | % p.a. | Accrues monthly-compounded on the guarantee-backed securities outstanding; paid to FEP (and Developer pro rata) from sale proceeds, like loan interest. |
| Inputs!E70 | DSA / developer success fee accrual start (signing) | 2026-10-26 | Date | Signing 26/10/2026. Fee accrues pro rata (30/360) from here to the sale date at E36, capped at E37. |
| Inputs!E71 | At-risk loan principal (devex) repaid from sale proceeds (1) or carried by FEP (0) | 0 | Switch | 0 = FEP carries its devex (only the 2× return is paid from proceeds), as in the Devex file. |
| Inputs!E73 | Interest pause end (accrual resumes on this date; user 09-Oct-26) | 2030-01-01 | Date | Every interest / return accrual (securities 6 % / 12 %, participation fee, devex and land 20 % with the 2× minimum, shareholder-loan and guarantee interest) pauses fro… |
| Inputs!E75 | Benchmark CAPEX – wind (csatlakozási díj alapdíj) | 406 | mHUF/MW | Wind leg of the %-driven security base. Matches the FEP workbook's 406 mHUF/MW. |
| Inputs!E76 | Bid bond rate (ajánlati biztosíték) | 0.005 | % of base | 0.5% of the bid-stage base (E91), no uplift – matches BORA Devex I6 (606 mHUF). |
| Inputs!E77 | Performance bond rate (teljesítési biztosíték) | 0.025 | % of base |  |
| Inputs!E78 | Supplementary security total rate (kiegészítő biztosíték) | 0 | % of base | 0 – no supplementary security (kiegészítő biztosíték) in this procedure: GYIK 70 (user 09-Oct-26; removed 07-Oct-26, was 5 %). Deposits rows 13–17 stay at 0; the Bora … |
| Inputs!E81 | Participation fee (részvételi díj) | 3 | mHUF per bid | Final tender: HUF 3.0m per bid (draft was 2.0m). Source: MEKH kiírás 30-Aug-26 via palyazatok.org, https://palyazatok.org/szeleromu-palyazat/ |
| Inputs!E82 | Number of separate bids (fee is per bid) | 2 | # | IC paper: BORA = 6 SPVs (5 × 50 MW + 1 × 30 MW). Case 1 keeps 6 bids at 200 MW as a conservative fee count. |
| Inputs!E83 | Connection request security (igénybejelentési biztosíték) | 4.5 | mHUF/MW | Cash only – the TSO bond cannot be posted as a bank guarantee; credited against the indirect connection fee. |
| Inputs!E84 | Connection fee – direct (csatlakozási díj – közvetlen) | 5.08346 | mHUF/MW | BORA_Devex_20261002_BG row 19 (unchanged since 25-Sep): direct connection fee 508,345,804 HUF (k€ 1,374 at 370 HUF/EUR in the example); expressed per MW so it scales w… |
| Inputs!E85 | Connection fee estimate (csatlakozási díj) | 73.0713 | mHUF/MW | BORA_Devex_20261002_BG row 18: 73,071,306 HUF/MW (unchanged), booked net of the TSO bond. The example applies it to its 260 MW wind capacity (V6; k€ 48,915); the model… |
| Inputs!E89 | Benchmark CAPEX – BESS | 200 | mHUF/MWh | BESS leg of the %-driven security base. |
| Inputs!E95 | Bid submission (pályázat benyújtása) | 2026-10-30 | Date | Final tender: window 4 Sep – 30 Oct 2026 (draft was 1–18 Sep). Source: MEKH kiírás 30-Aug-26 via Világgazdaság/Portfolio, https://www.portfolio.hu/gazdasag/20260831/me… |
| Inputs!E96 | Award (eredményhirdetés) | 2026-12-13 | Date | Final tender: results by 13 Dec 2026 (draft was 2 Dec). Same source as above. |
| Inputs!E97 | Connection request (igénybejelentés) | 2026-12-30 | Date | Within 15 days of award (linked: award + 14 days). |
| Inputs!E99 | HCsSz signature (connection agreement) | 2029-01-08 | Date | Connection fee paid here. |
| Inputs!E102 | Performance bond payment deadline after award | 17 | Days |  |
| Inputs!E107 | Bid bond credited into performance bond (1 = yes) | 1 | Switch | Only the top-up is a new cash outflow; the bid bond is not released separately. |
| Inputs!E109 | Include instalments falling due after the sale in FEP cash flows (1 = yes) | 0 | Switch | Default 0: instalments after exit are the buyer's obligation. Set 1 to stress. |
| Inputs!E110 | Bid bond on failed capacity refunded at award (1 = yes) | 1 | Switch | Per tender rules the bid bond is refunded if the bid is unsuccessful. |
| Inputs!E113 | Assumed all-in guarantee rate – guarantee-backed securities | 0.06 | % p.a. | Applied rate for guarantee-backed lines. Overrides EURIBOR + margin when E114 = 1. |
| Inputs!E114 | Use assumed all-in rate (1) or EURIBOR + margin (0) | 1 | Switch | 1 = cost guarantees at E113 flat; 0 = cost at the selected EURIBOR curve (E112) + E111 margin. |
| Inputs!E117 | Részvételi díj | 1 | Switch | Fee – cash only; guarantee not applicable. |
| Inputs!E118 | Ajánlati biztosíték | 0 | Switch |  |
| Inputs!E119 | Teljesítési biztosíték | 0 | Switch |  |
| Inputs!E120 | Igénybejelentési biztosíték | 1 | Switch |  |
| Inputs!E121 | Csatlakozási díj – közvetlen | 1 | Switch |  |
| Inputs!E126 | Csatlakozási díj | 1 | Switch |  |
| Inputs!E152 | NTP / FID lag after RTB | 3 | Months | RTB 30/06/2029 (E56) → NTP 30/09/2029 after this lag. |
| Inputs!E153 | Construction duration (NTP → COD) | 24 | Months | Two build seasons; 132 kV connection works (Szombathely OVIT node) on the critical path. |
| Inputs!E157 | Operating life | 30 | Years | Measured from COD (partial first and last calendar years). 30 years = grid-connection life constraint; 35 years is the usual wind assumption. |
| Inputs!E162 | Turbine supply incl. transport (6.x MW class) | 950 | k€/MW | ESTIMATE – base = European turbine 950 k€/MW (user decision 02-Oct-26, ≥80% EU-equipment tender points). IC: all-in ~1,500 k€/MW (Nordex 6.8 MW). Replace with turbine … |
| Inputs!E163 | Turbine installation and erection | 80 | k€/MW | ESTIMATE. |
| Inputs!E164 | Foundations | 90 | k€/MW | ESTIMATE. |
| Inputs!E165 | Balance of plant – civil (roads, crane pads, drainage) | 70 | k€/MW | ESTIMATE. |
| Inputs!E166 | Balance of plant – electrical (MV cabling, internal grid) | 60 | k€/MW | ESTIMATE. |
| Inputs!E167 | On-site substation and 132 kV bay (lump sum, portfolio) | 12000 | k€ | ESTIMATE. 132 kV connection works and substation incl. generator transformer – route length to be confirmed. Chinese cases (3 / 8): non-EU transformer, lower by Tender… |
| Inputs!E168 | BESS capex – all-in (capex only: no BESS revenue or opex modelled) | 200 | k€/MW BESS | ESTIMATE. Applied to BESS connection capacity (E20), not wind MW. No contingency applied. The BESS earns nothing in HTM – revenue and opex are not modelled, so HTM ret… |
| Inputs!E169 | Owner's costs, engineering, PM, construction insurance | 40 | k€/MW | ESTIMATE. |
| Inputs!E170 | Land purchase / easements at NTP | 15 | k€/MW | ESTIMATE. Land is leased; this covers rights-of-way and buy-outs. |
| Inputs!E171 | Contingency on the lines above | 0.05 | % | ESTIMATE. |
| Inputs!E175 | Decommissioning cost net of scrap (end of life) | 30 | k€/MW | ESTIMATE. |
| Inputs!E176 | Terminal value at end of life | 0 | k€/MW | Conservative: nil. Enter a repowering / land value if wanted. |
| Inputs!E179 | Net equivalent hours P50 (after wake, availability and electrical losses) | 2460 | MWh/MW/yr | IC paper: NCF 28.07%, P50 2,460 h. Already net of losses. |
| Inputs!E180 | Additional availability / loss haircut | 0 | % | Zero because the hours above are already net. |
| Inputs!E181 | Curtailment / grid constraint | 0.02 | % | ESTIMATE. |
| Inputs!E182 | Annual degradation | 0.003 | % p.a. | ESTIMATE. Compounding from first operating year. |
| Inputs!E183 | Yield sensitivity factor (1.00 = P50) | 1 | x | Set to P90/P50 ratio for a P90 case (typically ~0.85–0.88). |
| Inputs!E184 | Merchant power price sensitivity factor (1.00 = Aurora curve as selected) | 1 | x | Scales the Aurora wind capture price used in HTM row 42 (tornado ±10%). Keep at 1.00. |
| Inputs!E186 | Price source (1 = Aurora curve on Curves, 2 = flat estimate below) | 1 | Switch | 1 = Aurora Q2 2026 Central (Curves E16) converted to nominal (Curves E19); 2 = flat estimate below (kept for reference / fallback). |
| Inputs!E187 | Estimate: baseload price in first operating year | 100 | EUR/MWh | ESTIMATE, nominal. Replace with Aurora. |
| Inputs!E188 | Estimate: baseload escalation | 0.02 | % p.a. | ESTIMATE. |
| Inputs!E189 | Estimate: wind capture rate vs baseload | 0.85 | % | ESTIMATE. Hungarian wind cannibalisation is modest at low penetration. |
| Inputs!E190 | Merchant share of output | 0 | % | Balance sold under the PPA / CfD below. |
| Inputs!E191 | PPA / CfD strike in first operating year | 91.6 | EUR/MWh | ESTIMATE. |
| Inputs!E192 | PPA / CfD strike escalation | 0.01 | % p.a. |  |
| Inputs!E193 | PPA / CfD tenor from first operating year | 10 | Years | Merchant thereafter. |
| Inputs!E194 | Balancing / BRP cost | 4 | EUR/MWh | ESTIMATE. IC flags balancing charges as material. |
| Inputs!E195 | Grid use-of-system and system-operator charges | 1 | EUR/MWh | ESTIMATE. |
| Inputs!E196 | MEKH supervisory fee | 0.001 | % of revenue |  |
| Inputs!E197 | Robin Hood tax (energy suppliers' income tax) applies (1 = yes) | 0 | Switch | FLAG – default 0: each SPV is ≤ 50 MW, so the 31% levy on generators > 50 MW should not bite. Switch on the Control Panel (section c). Confirm with tax adviser. |
| Inputs!E198 | Robin Hood tax rate | 0.31 | % |  |
| Inputs!E201 | O&M full-service contract (turbines) | 12 | k€/MW/yr | ESTIMATE. Long-term FSA assumed for the full 30-year operating life (user decision 02-Oct-26) – no post-FSA uplift. |
| Inputs!E202 | Balance of plant maintenance | 3 | k€/MW/yr | ESTIMATE. |
| Inputs!E203 | Land lease | 3 | k€/MW/yr | ESTIMATE. Replace with signed lease terms (38 turbines contracted). |
| Inputs!E204 | Operational insurance | 2.5 | k€/MW/yr | ESTIMATE. |
| Inputs!E205 | Asset and commercial management | 1.5 | k€/MW/yr | ESTIMATE. |
| Inputs!E206 | Property tax (építményadó) | 3 | k€/MW/yr | ESTIMATE. Municipal; varies by settlement. |
| Inputs!E207 | Administration, audit, compliance (portfolio) | 150 | k€/yr | ESTIMATE. |
| Inputs!E208 | Opex contingency | 0.05 | % | ESTIMATE. |
| Inputs!E209 | Opex escalation | 0.025 | % p.a. | ESTIMATE. EUR-denominated, single index (no double indexation). |
| Inputs!E210 | HIPA local business tax | 0.02 | % of revenue | Base approximated as revenue (conservative – actual base deducts materials and subcontractors). |
| Inputs!E213 | Tax depreciation period (straight-line from COD) | 20 | Years | Hungarian tax rate for power-generation assets to be confirmed. |
| Inputs!E214 | Loss carry-forward usable per year (cap on taxable income) | 0.5 | % | Hungarian rule: losses offset up to 50% of the year's pre-loss tax base. |
| Inputs!E217 | Gearing – debt as % of total capex incl. connection fee | 0.6 | % | ESTIMATE. Merchant-heavy Hungarian wind may gear lower. |
| Inputs!E218 | All-in interest rate (fixed) | 0.05 | % | ESTIMATE. ~2.5% swap + 250 bps margin. |
| Inputs!E219 | Upfront and arrangement fees | 0.0175 | % of debt | ESTIMATE. |
| Inputs!E220 | Tenor from the first full operating year (annuity repayment) | 10 | Years | ESTIMATE. 10-year annuity from the first full operating year. Proxy for a sculpted loan with an early cash sweep (annual model, near-flat generation). |
| Inputs!E221 | Interest during construction capitalised (1 = yes, 0 = paid by equity) | 1 | Switch |  |
| Inputs!E222 | Commitment fee on undrawn senior debt (construction) | 0.009 | % p.a. | ESTIMATE ~0.9% p.a. (≈35% of margin), charged quarterly on the undrawn facility NTP → COD; added 02-Oct-26 (Open_Items #14). Paid in cash from equity, part of HTM fina… |
| Inputs!E224 | Discount rate – unlevered project NPV | 0.08 | % | Linked to the single discount rate (Inputs row 253 ← Control Panel section c), per decision 02-Oct-26 (Open_Items #10). Previously 7% typed. |
| Inputs!E226 | Include Stage 1 development-phase costs (devex, fees, securities) in HTM | 1 | Switch | 1 = full-cycle FEP returns; 0 = project-only from RTB. |
| Inputs!E227 | Developer success fee paid in cash at COD when the JV is live (superseded by the deemed exit) | 0 | Switch | No sale proceeds in HTM, so the fee is settled from equity at COD. |
| Inputs!E228 | Charge HTM a deemed exit at RTB (1 = yes, 0 = historic cost) | 1 | Switch | 1 = HTM is charged the net sale proceeds it gives up by not selling at RTB, so its IRR measures the incremental value of building versus selling. 0 = HTM carries the p… |
| Inputs!E231 | Sale at COD – tax the FEP gain (1) or participation exemption as at RTB (0) | 0 | Switch | 0 = gain on the SPV share sale exempt (same treatment as the RTB exit, E14); 1 = gain taxed at the CIT rate (E12). |
| Inputs!E233 | One-off municipal contribution commitment (scoring: points above 10, max at 50) | 50 | k€/MW | BID OFFERING – FEP's intended commitment, not an estimate. Paid to host municipalities in the COD year and capitalised into capex (HTM row 29). Not borne by the sell-a… |
| Inputs!E234 | Annual community contribution on energy fed in (scoring: points above 2, max at 3) | 3 | €/MWh | BID OFFERING – FEP's intended commitment, not an estimate. Charged as an opex line on net generation in every operating year, not escalated (HTM row 59). Sits at the b… |
| Inputs!E253 | Discount rate for NPV / Equity (all NPV / Equity metrics) | 0.08 | % | Per review 27-Sep-26: 8%, kept as a variable. Drives MOIC G185:G188 and row 229-230, HTM G136:G139, buyer NPV / Equity (HTM G214) and gearing block 15 NPV columns. |
| Inputs!E256 | Részvételi díj | 0 | % | ESTIMATE pending BORA_Devex_20260921_FEP classification. Different from section J (reimbursed by the BUYER at exit). Sunk fee. |
| Inputs!E257 | Ajánlati biztosíték | 0 | % | Credited into the performance bond at award; forfeited if the project fails afterwards. |
| Inputs!E258 | Teljesítési biztosíték | 0 | % | Forfeited if COD / tender obligations are missed. |
| Inputs!E259 | Igénybejelentési biztosíték | 1 | % | REFUNDED if the project fails (user correction 29-Sep-26, BORA Devex row 8). Cash deposit → funded on the 12% recoverable loan. Credited against the indirect connectio… |
| Inputs!E260 | Csatlakozási díj – közvetlen | 1 | % | Connection fee – agreed at 12% ('connection fee not at risk', BORA_Devex_20260921_FEP). Buyer reimburses at exit (section J). |
| Inputs!E265 | Csatlakozási díj | 1 | % | Connection fee – agreed at 12% ('connection fee not at risk', BORA_Devex_20260921_FEP). Buyer reimburses at exit (section J). |
| Inputs!E268 | DSRA – months of next year's scheduled debt service | 6 | months | MARKET STANDARD: 6-month DSRA funded at COD by equity, released as debt amortises. 0 = off. |
| Inputs!E269 | Cash sweep – share of cash after scheduled debt service | 0.5 | % | Merchant wind without offtake: lenders typically sweep 50–100% of excess cash in early years. Applied as prepayment (shortens tenor). 0 = off. |
| Inputs!E270 | Cash sweep – applies in operating years 1 to | 5 | years | Sweep window from the first full operating year. |
| Inputs!E298 | COD override (timeline moved forward) | 2031-09-30 | Date | Blank = RTB from milestones (E49:E54). If set, E55 = months so that RTB + E152 + E153 lands on this COD. |

Other inputs: Tender!E27 bottleneck loss on overbuilt MW 5 % (Bora!N18); Tender!E29 CN substation −10 %; Tender!E30 price-effect
pass-through 100 % (Calc E89); Tender!E31 PB held 6 months after COD (used only in the RtB price effect, not in HTM);
Calc!E133/E134 = Bora!N6/N7 COD price 2,450 / 2,100 k€ per installed MW; Calc!E74 net MVA 100; Curves (see curves.csv).

## 17. Materiality (shadow toggles, `python3 old_htm.py --toggles`; Δ FEP IRR in pp, Δ FEP NPV in k€)

| Mechanic switched off / changed | BORA 1 HTM-EU | BORA 2 HTM-EU | BORA 1 HTM-CN | BORA 1 COD-EU |
|---|---|---|---|---|
| Deemed exit charge off (FEP share kept 100 %) | −0.05 / −664 | +0.34 / +4,252 | −0.08 / −664 | – |
| Deemed exit off, old rule (FEP share → 70 %) | −0.05 / −15,383 | +0.34 / −9,885 | −0.08 / −28,887 | – |
| Deemed exit valued on the RtB-route MOIC result | +0.02 / +197 | +0.01 / +94 | +0.02 / +197 | – |
| DSRA off | +0.34 / +3,629 | +0.31 / +3,372 | +0.29 / +1,726 | 0 |
| Cash sweep off | +0.12 / +986 | +0.12 / +978 | +0.61 / +2,324 | 0 |
| DSRA and sweep off | +0.49 / +4,797 | +0.46 / +4,531 | +1.02 / +4,595 | 0 |
| Loss cap off (100 % usable) | +0.003 / +29 | +0.003 / +28 | +0.004 / +19 | 0 |
| Commitment fee off | +0.07 / +940 | +0.07 / +877 | +0.07 / +554 | +0.71 / +726 |
| Upfront fee off | +0.16 / +2,003 | +0.15 / +1,867 | +0.15 / +1,182 | +1.90 / +1,578 |
| IDC paid by equity instead of capitalised | −0.14 / −1,188 | −0.12 / −1,108 | −0.14 / −599 | −0.87 / −152 |
| Degradation off | +0.39 / +8,535 | +0.39 / +7,996 | +0.36 / +8,553 | −0.03 / −46 |
| Opex contingency off | +0.14 / +2,600 | +0.14 / +2,418 | +0.16 / +2,631 | 0 |
| HIPA off | +0.31 / +5,725 | +0.30 / +5,322 | +0.35 / +5,792 | 0 |
| Decommissioning off | +0.01 / +306 | +0.01 / +284 | +0.01 / +306 | 0 |
| Development phase excluded (E226 = 0) | +0.51 / +5,837 | +0.49 / +5,622 | +0.86 / +5,814 | +10.68 / +5,675 |
| PB guarantee held 6 months after COD | −0.01 / −137 | −0.01 / −119 | −0.02 / −137 | 0 |
| No bid-bond netting (both guarantees outstanding) | −0.03 / −328 | −0.02 / −285 | −0.04 / −328 | −0.32 / −247 |
| FX 400 instead of 365 HUF/EUR | +0.11 / +1,502 | +0.11 / +1,483 | +0.18 / +1,538 | +1.36 / +1,118 |
| Direct fee as Annex 1 lump sum | 0 (identical at 100 MVA) | 0 | 0 | 0 |
| Overbuild not built (100 MW) | −1.06 / −26,850 | −0.94 / −21,672 | −1.62 / −42,020 | −11.16 / −19,446 |
| … plus proportional PB call | −1.15 / −27,815 | −1.00 / −22,429 | −1.75 / −42,985 | −11.82 / −20,122 |
| BESS not built | +0.47 / +6,494 | +0.37 / +5,093 | +0.71 / +6,705 | +4.02 / +5,277 |
| … plus proportional PB call | +0.32 / +4,592 | +0.25 / +3,601 | +0.49 / +4,803 | +3.02 / +3,945 |
| Community €/MWh = 0 | +0.41 / +7,210 | +0.26 / +4,468 | +0.47 / +7,317 | +3.27 / +5,101 |
| CIT 0 % | +0.55 / +11,482 | +0.54 / +10,863 | +0.78 / +14,988 | −0.11 / −168 |
| Annual IRR vs XIRR on mid-period dates (date basis) | −0.002 | −0.002 | +0.001 | – |

Reading: the loss cap, decommissioning, the PB holding period, bid-bond netting, the direct-fee basis, the HTM date basis and
the deemed-exit basis are immaterial (< 0.05 pp) and a lean rebuild can drop or simplify them. DSRA + sweep (0.5–1.0 pp),
fees/IDC (0.2–0.3 pp HTM, 2–3 pp COD), degradation, HIPA and the development phase matter. The deemed exit is small in IRR for
BORA 1 (because its RtB profit is negative) but 0.34 pp for BORA 2 and changes NPV by up to 4 m€ – it is a methodology choice,
not a detail. Whether overbuild/BESS are actually built is the largest driver of all.

## 18. Bugs, inconsistencies and oddities

- **B1 Deemed exit ≠ RtB route.** Row 21/22 read MOIC!G182/G154/G151 recalculated under case 2/3. The RtB price effect
  (Tender E219 → MOIC G143) then uses PPA revenue: BORA 1 charge −1,046.7 vs the RtB route's −1,220.3 k€ profit (G143 4,175.7
  vs 3,927.8); BORA 2 3,355.8 vs 3,273.2; S1 unaffected. Impact ≤ 0.02 pp.
- **B2 Negative RtB profit becomes a credit.** BORA 1 RtB profit is negative, so the "opportunity cost" adds +1,046.7 k€ (and
  the A.RES buy-out is only −210.8 k€) to the HTM flows in 2029. The HTM IRR is then not "incremental to selling at RtB" in a meaningful sense.
- **B3 FEP share flips to 70 % when the deemed exit is off** (E128): 70 % of all equity flows including the development costs
  that FEP funds 100 % – not the JV deal used in the RtB and COD waterfalls (−15 to −29 m€ NPV).
- **B4 COD-year stub sold merchant.** PPA share applies for idx 1–10 (2032–2041); Oct–Dec 2031 (~89 GWh for BORA 1) is sold at
  Aurora (115.7 €/MWh) instead of the 91.6 PPA, ≈ +2.2 m€ revenue; the "10-year PPA" runs from 1-Jan-2032.
- **B5 PB released exactly at COD in HTM** (row 23 ends at COD) while Tender E31 (6 months after COD) is used in the RtB price
  effect and the call releases the PB only after the post-COD compliance decision. Immaterial (−0.01 pp).
- **B6 Bid bond netted into the PB although both are bank guarantees** (Deposits I10, Inputs E107 = 1; the call nets only cash).
  Economically equal to "bid guarantee released, PB in full" because the bid line stays outstanding; immaterial.
- **B7 Connection fee basis.** Direct fee stored per MW (E84 = 508.3 mHUF ÷ Calc E74) × net MW: right only while net MW = 100; municipal
  one-off on net MW (E18) while the basis is open in the call; DSA (10 k€/MW cap) on installed MW.
- **B8 Date bases differ:** HTM IRR annual from 2026 (t0 = 2026 regardless of month), HTM NPV end-of-year from 2026, COD IRR XIRR
  on mid-period dates (truncated), COD NPV annual with the sale at t = 5, RtB NPV daily from 31-Jul-2026. Effects small, but the NPVs are not on one date.
- **B9 MOIC nets inflows and outflows within a year** (G147, annual); the COD MOIC is proceeds ÷ contributions; RtB MOIC is monthly.
- **B10 Commitment fee on E75 − open − draw/2**, where open includes capitalised IDC → undrawn understated (tiny).
- **B11 Opex contingency on land lease and MEKH fee but not on the pre-COD lease in block 17** (row 339 omits the 5 %); HIPA on revenue (conservative).
- **B12 BESS earns nothing** (capex only, Inputs F168) – "BESS not built" raises returns.
- **B13 Aurora wind capture above baseload** in every year (Curves L22:L55, 100.6–104.3 %), inflated at 2 %; 2061 extrapolated.
- **B14 Data table and LibreOffice:** stale cells unless Calc!G306:DH330 is cleared (§0). `HTM!G257` gearing-block check still reads FAIL (block 15, not used by the Bora outputs).
- **B15 Unlevered tax proxy (row 101) ignores loss relief**, so the project IRR (G135) is slightly understated; not on the Bora sheet.

## 19. What a lean rebuild can drop

Keep: timeline flags, generation (P50 net hours, curtailment, degradation, bottleneck loss on the overbuilt MW), PPA + merchant
curve, the opex lines (could be merged into fixed €/MW + variable €/MWh + % of revenue), annuity debt sized on gearing with
capitalised IDC and fees, CIT with 20-year depreciation, the development phase and securities/guarantee fees, the COD sale with
the JV waterfall. Drop or simplify: loss-carry-forward cap (0.003 pp), decommissioning (0.01 pp), PB holding period and bid-bond
netting (≤ 0.03 pp), the deemed-exit basis switch, the community purchase right (off), supplementary securities (0 by GYIK 70),
Robin Hood switch (off), block 13–16 sensitivities. Decide explicitly: deemed exit (methodology, B2/B3), DSRA + sweep
(0.5–1 pp), and whether overbuild and BESS are built (largest driver).
