# Bora wind + BESS – analysis of the two existing Excel workbooks

Files analysed (read-only, formulas and cached values via openpyxl; no macros, neither file contains VBA):

| File | Sheets | Role |
|---|---|---|
| `Bora_IC_model_v4.xlsx` ("IC model") | 15 (5 hidden) | Full IC model, dated 11-Oct-2026 (Cover B3/B45) |
| `Bora_scenarios_meeting.xlsx` ("meeting file") | 12 (11 hidden, only `Bora` visible) | Meeting front-end on top of the **same engine** |

**Key finding on the relationship:** apart from naming, the meeting file is the IC model. `Control Panel` was renamed `Calc`, and its tender-strategy inputs (E22:I34) and 12 business inputs now link to a new one-page `Bora` sheet. `Tender, Inputs, HTM, MOIC, Dashboard, Curves, Deposits, Run_Map, Devex_FEP, BORA_Devex` match cell for cell (0 formula differences once the sheet rename is allowed for), and every cached result is the same. The meeting file drops `Cover`, `Open_Items`, `Change_Log` and the IC "Bora" project-view sheet. So all numbers below hold for both files. Cell refs are IC-model refs; in the meeting file read `Calc!` for `'Control Panel'!`.

Units: k€ unless stated, mHUF = HUF million, FX 365 HUF/EUR (Inputs E11). MW = net grid MVA unless "gross / installed".

---

## 1. Structure, sheets and flow

### 1.1 IC model – sheet map

| Sheet | Type | Purpose |
|---|---|---|
| Cover | Output / doc | Index, conventions, headline numbers of the active case (C27:C37) and update notes (B39:B45) |
| **Control Panel** (A1:CE330) | **Input hub + results** | a) route and strategy selectors (E8, E9, data-table driver E10); b) 5 tender strategies S0–S4 (E22:I34) with derived MW and points; c) every business input (rows 72–135, columns E base / F European / G Chinese); d) results; e) integrity per run; f) **Excel data table** `{=TABLE(,E10)}` G306:CC330 = 5 strategies × 5 routes = 25 runs; g) tender band tables (rows 233–300) |
| Tender (A1:AQ376) | Calc | MEKH 2026 capacity-tender scoring (Annex 3 weights, Annex 4 band tables), MW derived per strategy, promise price-effect valuation (section 4), 5×5 returns matrix (section 5), live tie-out (section 6), 19 checks (section 7), group caps (section 8), community purchase right (section 9) |
| Inputs (A1:P298) | Inputs | All static inputs in 10 case columns G:P; column E = active case via `INDEX($G:$P,$E$6)` with fallback to G when blank. Many G cells link to Control Panel / Tender |
| Curves | Inputs (data) | 3 EURIBOR monthly forward curves (rows 9–11); Aurora Q2-2026 HU baseload and wind-capture Low / Central / High, real 2025, 2027–2060 (rows 22–55), inflated at 2 % (E19) |
| Deposits | Calc | Schedule of MEKH securities and grid fees (12 lines, rows 7–18): basis, rate, gross/payable mHUF, date, k€, cash vs guarantee, recovery %; reconciliation to the BORA_Devex source |
| MOIC (A1:BP286) | Calc, monthly | **Sell-at-RtB engine**, months 0–60 (H:BP, Jul-26 → Jul-31), 16 blocks: devex, DSA, securities, cash vs guarantee, recovery, shareholder loans (12 % / 20 %·2×), exit waterfall, FEP and developer cash flows, returns, RtB-premium sensitivity, checks |
| HTM (A1:AR398) | Calc, annual | **Hold-to-maturity engine** 2026–2061 (H:AQ), plus buyer view at RtB (block 14), gearing sensitivity (15), lender downside (16), **sale at COD** (17), community right (17c), buyer view at our COD price (18) |
| Dashboard | Output | Headline comparison, sources and uses, waterfall, sensitivities, HTM KPIs, turbine comparison (section 11), tornado (section 12 – **static snapshot**), 4 charts |
| Bora | Output | Project view: inputs (rows 4–18) and all 5 routes for the active strategy, read from the data table (rows 22–54) |
| Run_Map (hidden) | Mapping | Live result formulas by route (C = RtB/MOIC, D = HTM, F = COD sale) that feed the data table output row 305 |
| Devex_FEP (hidden) | Calc | 34 devex line items (k€) with dates, FEP-pays flag, buyer-reimbursed flag, 20 %/2× multiples |
| BORA_Devex (hidden) | Pasted source | Copy of the developer's BORA_Devex_20260925.xlsx (HUF, own formulas, FX 370) |
| Open_Items, Change_Log (hidden) | Docs | 55 open items and a 117-entry audit trail |

### 1.2 Flow

```
Control Panel (selectors, S0–S4, business inputs)
   ├─► Tender (strategy → MW, BESS, points, price effect)  ──► Inputs G20/G21/G90/G233/G234/G297, M230/N230
   └─► Inputs (case columns; E = active case)
            ├─► Deposits (securities, grid fees)   ─┐
            ├─► Devex_FEP ◄─ BORA_Devex             ├─► MOIC (monthly, RtB sale)  ─┐
            ├─► Curves (EURIBOR, Aurora)            ─┴─► HTM (annual: HTM, COD sale, buyer views)
            └──────────────────────────────────────────────────────────────────────┤
Run_Map (picks MOIC / HTM results by route) ─► Control Panel row 305 ─► data table G306:CC330 (25 runs)
   ─► Control Panel sections d/e, Tender section 5, Bora, Dashboard sections 1 & 11, Cover
```

Two selectors: `Inputs!E6` (case/route: 1 RtB, 2 HTM-EU, 3 HTM-CN, 7 COD-EU, 8 COD-CN; cases 4/5/6 are legacy stresses, 9/10 spare) and `Tender!E5` (strategy 0 = S1, 1 = S2, 2 = S3, 3 = S4, 4 = S0). While Excel fills the data table, both are driven from `Control Panel!E10` (run code 1–25).

### 1.3 Meeting file

- `Bora` (visible, A1:W55): **inputs** D5:H17 (the 5 strategies, blue font) and N5:N16 (RtB premium 150, COD prices 2,450 / 2,100 k€/MW, turbine 950 / 500, PPA 90 €/MWh / 10 yrs, gearing 60 % / 50 %, interest 5 %, discount rate 8 %, COD 30-Sep-31); band and points reference M18:W31; **results** D37:H53 (RtB price, price effect, FEP profit / IRR / MOIC, ΔIRR vs S1, COD-sale and HTM IRR/MOIC for EU and CN); B55 model check.
- `Calc` = old Control Panel; E22:I34, E85, E91, E101:G101, E117, E119, F126:G126, E127, E130, E133, E134 now `=Bora!…`. All other business inputs (capex lines, opex, yield, tax…) stay as hard inputs on hidden `Calc`, so the meeting file has **inputs on two sheets**, one of them hidden.

---

## 2. Key inputs (IC model refs; active = Case 1 / S1)

### 2.1 Project and deal (Control Panel section c → Inputs)
| Input | Value | Cell |
|---|---|---|
| Grid node | Szombathely OVIT (SZHO B 132.00) | CP E73 → Inputs G280 |
| Net grid connection capacity | 200 MVA | CP E74 → Inputs G16/G281 |
| Gross installed wind (S1) | 261 MW (= 200 × 1.305) | Tender E40 → Inputs G297, E29 |
| BESS (S1) | 90 MW / 4 h / 360 MWh | Tender E41/E16 → Inputs G20, G21 |
| Signing fee to current owners | 250 k€ | CP E75 → Inputs G24 |
| DD / transaction costs; sale costs | 150 k€; 500 k€ | Inputs G25; G26 |
| JV 70/30 with A.RES (developer) | on, FEP 70 % | CP E76/E77 → Inputs G31/G32 |
| Developer success fee (DSA) | 2.5 k€/MW/yr on installed MW from 26-Oct-26, cap 10 k€/MW | Inputs G36, G37, G70 |
| Number of tender bids | 2 (participation fee 3 mHUF/bid) | CP E80; Inputs G81 |
| Devex incl. 10 % contingency | 3,137.85 k€ (FEP-borne 2,662 + land 425 reimbursed) | Inputs G292, Devex_FEP E39/H39/K39 |
| Investment date (month 0) | 31-Jul-2026 | Inputs G27 |

### 2.2 Tender securities and grid fees (Inputs F/G/H/I/J, Deposits)
| Item | Value | Cell |
|---|---|---|
| Benchmark capex – wind / BESS | 406 mHUF/MW; 200 mHUF/MWh | Inputs G75, G89 |
| Bid bond | 0.5 % of base (base 177,966 mHUF) = 890 mHUF, bank guarantee | Inputs G76, E91; Deposits row 8 |
| Performance bond | 2.5 % × uplift 1.5 (S1) = 6,674 mHUF gross, 5,784 payable, guarantee | Inputs G77, G90; Deposits row 10 |
| TSO connection-request security | 4.5 mHUF/MW, cash (900 mHUF) | Inputs G83; Deposits row 11 |
| Indirect connection fee | 73,071,306 HUF/MVA → 14,614 mHUF (37.6 m€), cash, 08-Jan-2029 | CP E81; Deposits row 18 |
| Direct connection investment | 508,345,804 HUF lump sum | CP E82; Deposits row 12 |
| Supplementary security | 0 % (none in this procedure, GYIK 70) | Inputs G78 |
| Guarantee fee; interest on guarantees / on cash loan | 6 % p.a.; 6 % / 12 % p.a. to FEP | Inputs G113, G65, G61 |
| At-risk loan return | 20 % p.a., min. 2× (flat 2× switch on) | Inputs G62, G63, G66 |
| Interest pause | 30-Jun-2029 → 01-Jan-2030 | Inputs G64, G73 |
| Totals | gross 23,592 mHUF; payable 59,732 k€ (cash 41,448 / guarantee 18,284) | Deposits E22:E27 |

### 2.3 Timing
| Item | Value | Cell |
|---|---|---|
| Bid submission / award | 30-Oct-2026 / 13-Dec-2026 | Inputs G95 / G96 |
| COD (input – drives the timeline back-solve) | 30-Sep-2031 | CP E91 → Inputs G298 |
| RtB (derived) / RtB sale (RtB + 9 m) | 30-Jun-2029 / 31-Mar-2030 (grid month 44) | Inputs E56, E58; CP E88 |
| NTP (RtB + 3 m), construction | 30-Sep-2029, 24 months | CP E93, E92 |
| Valid COD window | 30-Sep-2030 … 31-Dec-2031 (tender deadline 30-Sep-2032) | CP E95:E96; Inputs G237, G238 |
| Operating life | 30 years (to 2061) | CP E94 |

### 2.4 Capex (k€/MW; turbine and BoP lines on gross MW)
| Line | Base / EU | Chinese | Cell |
|---|---|---|---|
| Turbine supply incl. transport | 950 | 500 | CP E101:G101 → Inputs G162:I162 |
| Installation 80, foundations 90, BoP civil 70, BoP electrical 60, owner's costs 40, land 15 | same | same | CP E102:E107, Inputs G170 |
| Substation + 132 kV bay (lump sum) | 12,000 k€ | 10,800 (−10 % non-EU transformer) | Inputs G167, I167 (Tender E29) |
| BESS all-in | 200 k€/MW BESS (no revenue, no opex) | same | CP E106 → Inputs G168 |
| Contingency | 5 % (not on BESS) | | Inputs G171 |
| Total capex excl. / incl. grid fee | 1,941 / 2,136 k€ per net MW (HTM: 431.2 m€ = 2,156 k€/MW incl. municipal contribution) | CN: 306.6 m€, 1,533 k€/MW | Inputs E173/E174, HTM G30/G32 |
| Decommissioning; terminal value | 30 k€/MW; 0 | | Inputs G175, G176 |

### 2.5 Generation, prices, opex, tax, financing
| Input | Value | Cell |
|---|---|---|
| Net P50 hours | 2,460 h (IC paper NCF 28.07 %) | CP E109 |
| Curtailment; degradation; extra haircut | 2 %; 0.3 % p.a.; 0 % | CP E111, E112, E110 |
| Overbuild clipping | L20 = 2 % at 20 % overbuild, scaled quadratically: S1 4.65 %, S2/S3 8 %, S4 14.58 % | CP E113, Tender E27/E45 |
| Net hours after all losses (S1) | 2,298.7 h → 600 GWh/yr | HTM E36, N38 |
| Merchant price | Aurora Q2-26 HU **wind capture**, Central, real 2025 + 2 % inflation; 114.5 €/MWh nominal in 2032; life average 146 €/MWh nominal | CP E114, Inputs G186/G199, Curves E19, HTM N42, G153 |
| PPA | 90 €/MWh, +1 % p.a., 10 yrs, **100 % of output in cases 2/3/7/8**; base (case 1 HTM) is 100 % merchant | CP E116:G116, E117:E119 |
| Balancing / grid charges / MEKH fee | 4 €/MWh / 1 €/MWh / 0.1 % of revenue | Inputs G194:G196 |
| Opex (k€/MW/yr, 2031 prices, gross MW) | O&M 12, BoP 3, land lease 3 (from NTP), insurance 2.5, asset mgmt 1.5, property tax 3; admin 150 k€/yr; contingency 5 %; escalation 2.5 % | Inputs G201:G209 |
| HIPA local business tax | 2 % of revenue | Inputs G210 |
| Community contribution (tender) | 2 €/MWh (S1) | Tender J13 → Inputs G234 |
| Municipal one-off (tender) | 20 k€/MW net (S1), paid in the COD year | Tender J12 → Inputs G233 |
| CIT; Robin Hood | 9 %, 20-yr SL depreciation, 50 % loss cap; Robin Hood 31 % **off** | Inputs G12, G213, G214, G197 |
| Participation exemption on share sale | on | Inputs G14 |
| Gearing (of capex incl. grid fee) | base 55 %, EU 60 %, CN 50 % | CP E126:G126 |
| Debt | 5.0 % fixed, 1.75 % upfront, 0.9 % commitment fee, 10-yr annuity from first full year, IDC capitalised, 6-month DSRA, 50 % cash sweep yrs 1–5 | CP E127:E129, Inputs G219:G222, G268:G270 |
| Discount rate (NPV, hurdle, economic profit, buyer) | single 8 % | CP E130 → Inputs G253 (→ E13, E224, E225) |
| Lender tests | min DSCR 1.5×, target equity IRR 11 %, downside P90 0.9 × Aurora Low, min DSCR 1.05× | Inputs G251:G252, G272:G274 |

### 2.6 Exit / price inputs
| Input | Value | Cell |
|---|---|---|
| RtB exit premium | **150 k€/MW × 200 MVA (net basis) = 30.0 m€** | CP E85, E86 (0 = net MVA) → Inputs G141, G149/E150 |
| RtB premium sensitivity grid | 50, 75, 100, 125, 150, 175, 200 k€/MW | Inputs G142:G148 |
| Buyer reimburses cash securities, connection fee, land | 100 % per line (participation fee 0 %) | Inputs G129:G138 |
| Pass-through of the promises' PV into the price | 100 % | CP E89 → Tender E30 |
| Sale-at-COD price per gross MW (before clipping) | EU 2,450 / CN 2,100 k€/MW | CP E133, E134 |
| Buyer target IRRs for implied COD price | 7 / 8 / 9 % | CP E135:G135 |
| Community purchase right (25 % of SPV at cost) | off | Tender E360:E362 |

---

## 3. Scenarios / strategies and headline results

### 3.1 The two dimensions

**Routes (exit strategy; Control Panel C14:C18)**
| Route | Case | Turbine | Exit | Finance / revenue |
|---|---|---|---|---|
| RtB sale | 1 | strategy's own (row 22) | sale of project rights 9 months after RtB, premium per MVA ± price effect | no construction |
| Sale at COD – EU | 7 | European 950 k€/MW | 100 % of SPV sold at COD for 2,450 k€/gross MW × (1 − clip) | 60 % debt, 100 % PPA |
| Sale at COD – CN | 8 | Chinese 500 k€/MW | as above at 2,100 k€/MW | 50 % debt, 100 % PPA |
| Hold to maturity – EU | 2 | European | build and operate 30 years, FEP 100 % (A.RES bought out at a deemed RtB exit) | 60 % debt, 10-yr PPA, then merchant |
| Hold to maturity – CN | 3 | Chinese | as above | 50 % debt |

**Tender strategies S0–S4 (ÖSZ commitments; Control Panel E22:I34, meeting Bora D5:H17)**

| Commitment | S0 Zero bid | S1 Current bid (price reference) | S2 Free points | S3 Cost-efficient (~70) | S4 Max score |
|---|---|---|---|---|---|
| ÖSZ-1 municipal one-off (k€/MW) | 0 | 20 | 20 | 50 | 50 |
| ÖSZ-2 community (€/MWh) | 0 | 2 | 2 | 2 | 3 |
| ÖSZ-3 overbuild | 0 % | 30.5 % | 40 % | 40 % | 54 % |
| ÖSZ-4 BESS ratio / basis | 0 | 34.48 % of wind | 45 % of MVA | 45 % of MVA | 51 % of wind |
| ÖSZ-6 performance security (× min.) | 1.0 | 1.5 | 1.6 | 2.0 | 2.1 |
| ÖSZ-5 partial award / ÖSZ-8 shared connection | no / no | no / no | yes / yes | yes / yes | yes / yes |
| ÖSZ-9a EU turbine / 9b EU transformer | 0 / 0 | 1 / 0 | 1 / 1 | 1 / 1 | 1 / 1 |
| Installed wind / BESS (MW) | 200 / 0 | 261 / 90 | 280 / 90 | 280 / 90 | 308 / 157 |
| Clipping loss | 0 % | 4.65 % | 8 % | 8 % | 14.58 % |
| **Score – EU routes / CN routes** (CP rows 52/53) | 0 / 0 | 36 / 24 | 51.5 / 36.5 | 69.5 / 54.5 | 90 / 75 |

Scoring: Annex 3 weights (Tender E153:E161: 20/15/10/15/5/5/10/5/15 = 100) × Annex 4 inner points from lookup band tables (Tender rows 91–150). ÖSZ-7 brownfield is not available at Bora's node. Validity needs installed wind ≥ net MVA (all valid). Totals are rounded to 2 dp (Tender rows 84–86, check path row 88).

### 3.2 Headline results – full 5 × 5 matrix (Control Panel data table rows 306–330; also rows 156–222, meeting Bora rows 37–53)

FEP view, post-tax. Δ columns on CP J:M.

**FEP IRR**
| Route | S0 | S1 | S2 | S3 | S4 |
|---|---|---|---|---|---|
| RtB sale (CP I326/306/311/316/321) | 50.2 % | **20.9 %** | 20.0 % | 13.7 % | −7.5 % |
| COD sale EU (I327/307…) | 38.2 % | 31.4 % | 29.0 % | 27.4 % | 16.3 % |
| COD sale CN | 47.4 % | 42.1 % | 40.8 % | 39.0 % | 29.6 % |
| HTM EU | 9.7 % | 10.0 % | 9.7 % | 9.7 % | 9.0 % |
| HTM CN | 12.4 % | 13.1 % | 12.9 % | 12.9 % | 12.3 % |

**FEP net profit post-tax (k€; column H)**
| Route | S0 | S1 | S2 | S3 | S4 |
|---|---|---|---|---|---|
| RtB sale | 37,649 | **15,174** | 14,664 | 10,039 | −5,349 |
| COD sale EU | 106,896 | 103,212 | 98,468 | 93,022 | 57,533 |
| COD sale CN | 129,417 | 135,332 | 135,127 | 129,714 | 102,798 |
| HTM EU (undiscounted 30-yr nominal sum) | 941,142 | 1,137,959 | 1,155,103 | 1,152,984 | 1,103,513 |
| HTM CN | 1,063,494 | 1,299,225 | 1,328,203 | 1,326,543 | 1,296,069 |

**FEP MOIC** – RtB 1.82 / **1.31** / 1.30 / 1.20 / 0.90; COD-EU 1.74 / 1.54 / 1.48 / 1.45 / 1.25; COD-CN 2.02 / 1.80 / 1.75 / 1.70 / 1.50; HTM-EU 5.65 / 6.13 / 5.95 / 5.99 / 5.64; HTM-CN 6.95 / 7.79 / 7.62 / 7.68 / 7.40 (column L).

**FEP NPV @ 8 % (k€, column J)** – RtB 24,465 / **7,253** / 6,826 / 3,234 / −8,589; COD-EU 62,268 / 57,525 / 53,744 / 49,989 / 24,495; COD-CN 78,435 / 80,447 / 79,851 / 76,096 / 56,526; HTM-EU 45,836 / 62,081 / 55,912 / 55,966 / 34,049; HTM-CN 105,572 / 139,279 / 138,606 / 138,567 / 124,487.

**NPV/Equity (column K)** – RtB S1 0.178; COD-EU S1 0.401; COD-CN S1 0.631; HTM-EU S1 0.377; HTM-CN S1 0.973.
**Peak exposure (column M)** – RtB ~46–52 m€; COD 127–232 m€; HTM 179–238 m€.
**Minimum DSCR** (CP rows 217/222) – HTM-EU 1.27 / 1.16 / 1.13 / 1.12 / 1.00; HTM-CN 2.12 / 1.92 / 1.87 / 1.84 / 1.64.

**RtB price and price effect (CP rows 181–186)**
| | S0 | S1 | S2 | S3 | S4 |
|---|---|---|---|---|---|
| Gross RtB premium (k€) | 30,000 | 30,000 | 30,000 | 30,000 | 30,000 |
| Price effect vs S1 (PV of extra promise costs to the buyer) | +28,930 | 0 | −136 | −5,687 | −24,212 |
| Effective RtB price (k€) | 58,930 | 30,000 | 29,864 | 24,313 | 5,788 |
| per MVA (k€) | 294.7 | 150.0 | 149.3 | 121.6 | 28.9 |

**Sale-at-COD price and buyer view (CP rows 192–208)** – EU: EV 609.7 m€ in S1 (2,336 k€/MW after clipping); buyer unlevered IRR at our price 7.24 %; implied price at 7/8/9 % buyer IRR = 2,400 / 2,146 / 1,932 k€/MW. CN: EV 522.6 m€ (2,002 k€/MW), buyer IRR 8.55 %; implied 2,377 / 2,124 / 1,912 k€/MW.

### 3.3 Live (active) case – Case 1 RtB × S1 (Cover C28:C37, MOIC G182:G194)
FEP net profit 15,174 k€ (5,538 mHUF); MOIC 1.31×; XIRR 20.9 %; NPV 7,253 k€; NPV/Equity 0.18; peak exposure 48.5 m€; economic profit 9,943 k€. Integrity: MOIC G270 "ALL OK", HTM G172 "ALL OK", Tender E343 "OK", 25 of 25 runs OK (CP E231).

What the RtB profit is made of (MOIC rows 143–154, Dashboard rows 28–35): proceeds 71.9 m€ (premium 30.0 + reimbursement of securities, grid fee and land 41.9). From that FEP takes 12 % interest on its cash loan (4.37 m€), the 2× return on devex/land (3.09 m€) and 6 % on guarantees (3.22 m€). After the DSA (2.24 m€) and sale costs, the distributable amount is 16.58 m€, split 70/30 (FEP 11.61 m€). Most of FEP's profit is therefore "interest" on its own shareholder loans; the RtB premium itself only drives the 70 % share.

RtB premium sensitivity (MOIC rows 225–243, Dashboard rows 62–69): 50 → profit 1.2 m€, IRR 1.7 %; 100 → 8.2 m€, 11.6 %; 150 → 15.2 m€, 20.9 %; 200 → 22.2 m€. Each 25 k€/MW ≈ 3.5 m€ of FEP profit. NPV breaks even at about 112 k€/MW.

HTM view of the same active case (case-1 parameters: 55 % gearing, 100 % merchant): project IRR 10.9 % post-tax, equity IRR 12.0 %, MOIC 6.4×, LCOE 102 €/MWh, min DSCR 1.64, downside DSCR 0.89 ("BELOW", HTM G325). Buyer at the RtB price earns 11.1 % unlevered / 12.3 % levered (HTM G215/G216). **Implied maximum RtB premium (buyer NPV = 0 at 8 %) is 944 k€/MW unlevered and 1,012 k€/MW levered (HTM G221/G222), against the 150 used.**

---

## 4. Calculation logic (main chains)

### 4.1 Strategy → physical project (Tender section 2)
- Gross wind `Tender!E40 = ROUND($E37*(1+E$14),6)` (net MVA × (1 + overbuild)).
- BESS `E41 = ROUND(E$15*CHOOSE(E$25+1,$E37,E40),6)`; MWh `E42 = E41*E$16`.
- Clipping `E45 = $E$28*(MAX(0,E43)/$E$27)^2`.
- Active column J = `INDEX($E:$I,1,$E$5+1)`; this feeds Inputs G20/G21/G90/G233/G234/G297 and M230/N230.
- Points: band number `E53 = SUMPRODUCT(($D$93:$D$101<E51)+($D$93:$D$101=E51)*$E$93:$E$101)`, inner points `INDEX($F$92:$F$101,E53+1)*eligible`, weighted `×weight/100`; total `E84 = ROUND(E83+E82,2)`.

### 4.2 Promise price effect (Tender section 4)
Δ commitments vs S1 (rows 185–194) → annual Δ cash flows to the buyer (rows 205–208: municipal one-off less depreciation shield, community €/MWh × generation × (1 − 9 %), BESS capex phased with depreciation shield, Δ performance bond × 6 % guarantee fee) → PV at the RtB sale date at 8 % (`F205 = SUMPRODUCT($H205:$AQ205,$H$197:$AQ$197)`) → adjustment `E213 = $E$30*$E$212` → `MOIC!G143 = Inputs!E141*Inputs!E150 + Tender!E213`. For the COD sale only the post-COD community contribution is passed through (E218:E219 → HTM row 340).

### 4.3 RtB route (MOIC, monthly)
Flags (rows 6–10) → devex `H18 = SUMIFS(Devex_FEP!$H$5:$H$38,Devex_FEP!$I$5:$I$38,H$4)` → securities per line `IF(H$4=Deposits!$L$n,Deposits!$O$n,0)` (rows 28–39) → cash vs guarantee split (blocks 4/5), guarantee fee `H79 = (H73+H74)*H78/12` → shareholder loans compounding `IF(Inputs!$E$69=1,(1+r)^(1/12)-1,r/12)*(open+draw)*flag` (rows 111, 126) → waterfall G143:G154 → FEP flows rows 158–175 → tax at sale `-flag*9%*MAX(cum − exempt share gain,0)` (row 174) → `G187 = XIRR(H175:BP175,H5:BP5)`, `G186 = (inflows + tax)/outflows`, `G192 = SUMPRODUCT(flow/(1+8%)^(days/365))`.

### 4.4 HTM route (annual) – revenue → EBITDA → CF → returns
- Generation `N38 = N35*N36*N37*N10/1000` (gross MW × net hours × degradation × operating fraction).
- Price `N45` = Aurora capture (Curves K, × factor) or flat estimate; PPA `N47 = $E$47*(1+esc)^(year−COD year)`, share `N48`.
- Revenue `N52 = N38*N45*(1-N48)`, `N53 = N38*N47*N48`.
- Opex rows 56–68 (per MWh, per MW × escalation index N55, % of revenue, 5 % contingency, HIPA 2 %) → EBITDA `N70 = N54+N69`.
- Debt: size `E75 = gearing × capex`, drawn with the construction share, IDC capitalised, annuity `E82 = PMT(5%,10,−(draw+IDC))`, sweep row 87, DSRA rows 88–89.
- Tax: depreciation `N92 = (capex + IDC + TSO bond)/20 × profile`; taxable `N93 = N70−N92−N81+N85+N16`; losses capped at 50 %; CIT 9 %.
- Unlevered CF `N110 = dev phase + capex + EBITDA + decom + TV + unlevered tax proxy`.
- Equity CF `N123 = Σ rows 114–122`; FEP `N129 = N123*E128*(1−community)+buy-in`.
- Results: `G135 = IRR(H110:AQ110)`, `G139 = IRR(H129:AQ129)`, `G140` NPV @ 8 %, `G147` MOIC, `G156` LCOE, `G126` min DSCR.
- The development phase rows 16–23 bring in the Stage-1 cash from MOIC/Deposits, plus a **deemed exit charge**: row 21 = −RtB profit forgone (−15.2 m€) and row 22 = −cash to buy out A.RES (−7.2 m€), both in 2030.

### 4.5 Sale-at-COD route (HTM block 17)
Pre-COD flows rows 335–339 → price `M340 = flag*(Inputs!E230*Inputs!E29 + Tender!E219)` → repay debt (row 341) → JV waterfall 17b (same claims as the RtB waterfall) → FEP flows row 359/360 → `G352 = XIRR(H359:AR359,H362:AR362)`. Buyer view at our price: block 18, XIRR / XNPV at 7/8/9 %.

### 4.6 Data table
`Control Panel!G306 = {=TABLE(,E10)}`; column F holds the run code 1–25; the output row 305 = `INDEX(Run_Map!$C$n:$F$n,1,Tender!$E$293)`. All the strategy × route comparisons shown in the IC model and the meeting file are cached data-table values. Tender section 6 (E305) and Bora row 54 flag a stale table.

---

## 5. Model quality vs "everything a formula, inputs separated, consistent timeline"

### 5.1 Strengths
- Calculation grids are clean. My R1C1 consistency scan of MOIC H:BP, HTM H:AQ and the Tender timeline found every row consistent; only the anchor first column differs (opening balances = 0, start dates).
- Formulas contain almost no literals: only check tolerances, `8760` in the capacity factor, `DATE()` inputs on Inputs, and `EDATE(…,24/36)` in Deposits K16/K17.
- No #REF!, #DIV/0!, #VALUE! or #N/A in any cached value in either file. The only "#REF" text is a string in a Change_Log note.
- Extensive integrity checks: MOIC rows 254–270, HTM rows 163–173, Tender rows 324–343, Control Panel link checks in column P.
- Clear units column D, a documented colour convention, and a Change_Log / Open_Items trail.

### 5.2 Errors and things that look wrong (ordered by importance for a bidding model)
1. **There is no probability of winning or competitive-bid logic.** The strategies are ranked only on economics given an award. S0 (0 points, almost certain to lose) shows the best RtB IRR (50 %) and profit (37.6 m€), because the model credits the buyer's saving on promises as +28.9 m€ of price (Tender E213, CP E183). No score-to-award probability, cut-off score, competitor field or expected value is modelled.
2. **The price effect is asymmetric.** The buyer is charged the PV of the promise costs (BESS capex, municipal and community contributions, bond fee) relative to S1, but the RtB premium stays on 200 MVA net (CP E86 = 0). The extra energy from the overbuild (S2/S3 at 280 MW, S4 at 308 MW) is never credited to the RtB price. Overbuild therefore only adds cost (securities on gross MW, DSA on installed MW) on the RtB route.
3. **RtB price and COD price disagree.** The RtB premium of 150 k€/MW leaves the buyer an 11.1 % unlevered IRR (implied maximum premium 944 k€/MW at 8 %, HTM G221). The sale at COD is priced at a 7.2 % buyer IRR (CP I195). The two exits are not valued at a consistent buyer hurdle, which inflates COD-sale returns relative to the RtB sale. The RtB premium is a hard input (IC paper range 150–250), not derived.
4. **Merchant price assumption.** Aurora "wind capture" is above baseload in every year: capture rate 100.6–104.3 %, average 102.5 % (Curves L22:L55, E72). This is unusual for wind, so check it with Aurora: the columns may be swapped or mislabelled. Inflated at 2 %, it averages 146 €/MWh nominal over life (HTM G153), and there is no merchant haircut (Open_Items #54). The HTM routes' 1.1–1.3 bn€ "net profit" and 6–8× MOIC follow from this. Curves still carries the amber "awaiting paste" format and Cover D9 says "Aurora awaiting paste", although the data is in.
5. **Run_Map F34 and F35 are mislinked.** "FEP net profit pre-tax / post-tax – Case B" for the sale-at-COD route point to `HTM!G357` (senior debt repaid at COD) and `HTM!G356` (equity value at COD). As a result, CP R307/S307 and the IC `Bora` G34:H35 show "Case B profit" of 272,538 / 336,670 k€ (EU) and 161,497 / 360,610 k€ (CN), which are debt and equity values, not profits.
6. **A hidden FAIL check.** `HTM!G257` "Check: scenario at live gearing ties to base FEP equity IRR (G135)" reads **FAIL** (F252 11.963 % vs G139 11.966 %, because the gearing block approximates sweep/DSRA). It is not counted in HTM G172 "ALL OK". Its label also says G135 but it compares G139.
7. **Lender downside fails.** HTM G324 min DSCR is 0.89× on Aurora Low + P90, and G325 reads "BELOW – debt would be sized down". Debt is still sized at a fixed 55–60 % gearing, not by DSCR. The HTM-EU S4 base min DSCR is 1.00×. Open_Items #1 still quotes the downside DSCR as 1.36×.
8. **The deemed exit charge sits inside HTM cash flows.** HTM rows 21–22 (−22.4 m€ in 2030) put an opportunity cost into the project and equity IRR/MOIC/NPV of the HTM routes. This is a methodological choice still open (Open_Items #29): HTM returns are "incremental to selling at RtB", not stand-alone.
9. **Node cap flag.** Tender I353 / CP E65: "FLAG – eligibility: 200 MVA > 100 MVA node cap (V.7 / VII.36)". This is shown as a flag only, not an integrity check (CP E67). If the 100 MVA per-node group cap applies, the 200 MVA bid at Szombathely would breach it. **This is critical for the bid.**
10. **The tornado is a stale static snapshot.** Dashboard section 12 (E247:H258, E286:H295) is typed values from 01-Oct-26: P50 2,706 h vs live 2,460; EU base IRR 10.78 % vs live 9.99 %; COD Jun-31 vs live Sep-31. The 4 charts plot these values, including the meeting file's chart "European turbine – FEP equity IRR range by driver (base 10.8%)".
11. **Deposits reconciliation hides a FAIL.** `Deposits!E48` = "FAIL" (gross difference 2,413 mHUF), but E34 / MOIC G254 show "n/a – inputs differ from BORA_Devex source" because 3 drivers differ (E47). The check is suppressed, not passed.
12. **FX is inconsistent.** The model uses 365 HUF/EUR (Inputs G11). The pasted `BORA_Devex` sheet, which drives the devex amounts in Devex_FEP column L, uses 370 (BORA_Devex AF6), and the Inputs G11 comment also says 370.
13. **Chinese and European turbines differ only in capex, gearing and score.** Yield, availability, degradation, O&M and warranty are identical (Dashboard E217:F221: same 17,250 GWh), so the CN routes always win on returns. They lose 12–15 tender points (ÖSZ-9), but that has no economic effect without an award probability.
14. **There is no BESS revenue or opex.** BESS is pure capex (18 m€ in S1, 31 m€ in S4) plus security base; this is a user decision (Inputs F168). For bidding this undervalues the BESS commitment.
15. **HTM discounting convention.** End-of-year factors indexed from 2026 (`H12 = 1/(1+r)^H4`, no mid-year, flows in 2026 undiscounted), while MOIC discounts daily from month 0 and the COD sale uses XIRR. The RtB, COD and HTM NPVs are therefore not on an identical date basis.
16. **"Net profit" mixes horizons.** The undiscounted 30-year nominal sum for HTM sits beside 3–5 year exits (CP rows 156–160, Bora rows 22–23); the RtB IRR is XIRR on monthly flows while the HTM IRR is annual IRR.

### 5.3 Hardcodes and inputs inside calculation sheets
- **Tender** holds 256 typed inputs. Some are genuine scheme parameters (band tables rows 91–150, weights E153:E172, caps E349/E350, community right E360:E362). Others are business assumptions outside the Control Panel: E27 reference overbuild 20 %, E29 CN substation −10 %, E31 6-month security holding, E48 RtB route available, and the route codes in rows 226–228. The "one source" claim on the Control Panel is therefore only partly true.
- **Inputs** G-column values not on the Control Panel include the substation 12,000 k€ (G167), land 15 k€/MW (G170), capex contingency (G171), decommissioning (G175), balancing/grid/MEKH (G194:G196), insurance/AM/property tax/admin (G204:G207), HIPA (G210), tax depreciation (G213), fees (G219), lender structuring (G268:G274), deposit switches and recovery % (rows 117–138, 256–265), and the securities unit rates (rows 75–90).
- **Control Panel F101 = 950** duplicates E101 as a typed value rather than blank, against its own rule "blank = same as Base". In the meeting file it is linked.
- **Dashboard J48:J57** hold typed BORA Devex snapshot values (documented in comments), and section 12 holds the tornado values (see 5.2 #10).
- **BORA_Devex** is a pasted developer file whose formulas embed many literals (e.g. `I15 = 38000000*0.2*6`, `I16 = 72000000/280*AF7*1.15`, `I48 = 38000000*0.4*6`) and its own FX of 370. It feeds the model devex through Devex_FEP column L.
- **Curves** Aurora and EURIBOR data is typed by design (input data). Curves E69 uses the target COD (Inputs E101), not the model COD, for the "first operating year" memo.
- MOIC rows 224–243 reuse the month columns H:N as 7 premium scenarios on the same sheet. This is labelled, but a structural inconsistency.

### 5.4 Stale or inconsistent labels and documentation
- Version naming: Cover "IC model v4", Inputs A1 / HTM A1 / Dashboard B1 "HUN Onshore Wind – MOIC model v5", Cover B44 "Bora IC model v2".
- Inputs F72 says "8.9 %" while the value is 8 %. Inputs G21 comment says "BESS 100 MW / 200 MWh (2h)" while the live value is 90 MW / 4 h. Inputs F82 says "6 SPVs … keeps 6 bids" while the value is 2.
- Open_Items cell references and values are partly stale: #4 15-yr tenor / 60 % gearing, #6 "Inputs E177", #1 DSCR 1.36×, #30–#33 superseded by DSRA/sweep now modelled.
- Strategy column order differs between sheets: Control Panel E:I = S0, S1, S2, S3, S4; Tender E:I = S1, S2, S3, S4, S0; selector codes 0 = S1 … 4 = S0. This invites mislinks, although the CP column P link checks currently pass.
- The data table needs Excel. A LibreOffice headless recalculation reproduced the MOIC/HTM headline cells but left other cells in a data-table run state (e.g. HTM G147 7.40 vs 6.41; HTM G221 839 vs 944). Do not rely on LibreOffice for these files.

---

## 6. Bidding, auction, tender, price, RtB sale, buyer/seller, valuation – what exists

- **Tender = MEKH/MAVIR capacity tender 2026** (kiírás 30-Aug-2026, GYIK rounds up to 5). Bid window 4 Sep – 30 Oct 2026; award by 13-Dec-2026; building permit by 01-Jan-2030; earliest connection 30-Sep-2030; COD deadline 30-Sep-2032 (Inputs G95:G104, G236:G241).
- **The bid is not a price auction in this model.** It is a non-price scoring of ÖSZ-1 … ÖSZ-9 commitments (municipal contribution, community €/MWh, overbuild, BESS ratio and duration, partial award, performance security, brownfield, shared connection, EU equipment). All rules sit in Tender sections 1–3 and Control Panel section g (offerable ranges and points per band, max 100). Minimums: ÖSZ-1 10–50 k€/MW, ÖSZ-2 2–3 €/MWh (below = 0 points but valid, GYIK 17); ÖSZ-4 ≥ 5 % of installed wind, ≥ 4 h, BESS MW ≤ wind MW; ÖSZ-6 100–200 % of the minimum security.
- **Bid costs and securities**: participation fee 3 mHUF/bid; bid bond 0.5 % of the benchmark-capex base; performance bond 2.5 % × uplift (the uplift is the ÖSZ-6 commitment); TSO bond 4.5 mHUF/MW; connection fees (Deposits sheet). The performance bond is "at risk if COD/permit deadlines missed" (Inputs F236, F238).
- **Eligibility**: company-group caps of 100 MVA per node and 200 MVA in total (Tender section 8, **currently flagged**); 25 % community purchase right at cost within 2 years of award (Tender section 9, off).
- **Seller (FEP + A.RES JV) economics at RtB**: premium × MVA ± price effect + buyer reimbursement of cash securities, connection fee and land, less claims (loans + interest, 2× devex, DSA, sale costs), then 70/30 (MOIC block 9; Dashboard sections 3–4; tie-out to BORA Devex 01-Oct-26 in Dashboard rows 46–58).
- **Buyer valuation**: HTM block 14 (buyer buys at the RtB price, builds and operates: IRR and NPV, implied max premium 944 / 1,012 k€/MW, premium grid 14b/14c); block 18 / CP rows 190–208 (buyer IRR at our COD price, implied COD price at 7/8/9 %).
- **Promise price effect**: Tender section 4 (described in §4.2), 100 % pass-through input.
- **Not present**: win probability or score-threshold modelling, competitor scores, expected-value / decision tree across award/no-award, bid-bond forfeiture or failure scenarios for the awarded project (only a "conversion rate" Inputs G17 = 100 % and a refund-on-failure split), optimisation of commitments, and a CfD / METÁR premium (the only contracted revenue is the 90 €/MWh PPA; no CfD/METÁR logic exists).

---

## 7. Formatting conventions

IC model (Cover B19:B23; observed counts):
- **Inputs**: blue font `#0000FF` on cream fill `#FFF9E6` (119 cells on the Control Panel, 256 on Tender, 113 on Inputs).
- **Key assumptions to confirm**: yellow fill `#FFFF00` (Inputs, about 130 cells).
- **Awaiting data**: amber fill `#FFC000` (Curves Aurora block, 410 cells, still amber although filled).
- **Links from other sheets**: green font `#008000`.
- **Formulas**: black / dark-grey `#262626` or theme font.
- Check and link-status cells: green fill `#E2EFDA` with dark-green font `#375623`. Headers and totals: grey `#F2F2F2`. Key outputs: `#FFF2CC`. HTM block 17: `#D9E1F2`. Table headers: navy `#1F3864` with white font.
- **Number formats**: `#,##0;(#,##0);-` (HTM k€), `#,##0.0;(#,##0.0);-` (MOIC k€), `0.0%` / `0.00%`, `0.00x` multiples, `d-mmm-yy` / `dd/mm/yyyy` dates, `+#,##0;-#,##0;0` for Δ columns.
- Units in column D on every sheet; sign convention + = inflow to the party.
- Comments (author Zádeczki Zalán) document sources on Inputs, Dashboard J48:J57 and Devex_FEP.
- Calculation is automatic with data tables (`fullCalcOnLoad=1`). The sheet note recommends "Automatic except for data tables", but that setting is not saved.

Meeting file `Bora` sheet: inputs in blue font without fill, formulas in black; custom formats `"Yes";"Yes";"No"` for 1/0 switches, `"max "0`, `0.0#`, `0.00"x"`, `0.0%`. Results in €m (/1000).

---

## 8. Implications for the new bidding model and IC deck
1. Reuse the Tender scoring engine (band tables, weights, GYIK eligibility). It is well structured and checked.
2. Add an **award-probability layer** (score vs expected cut-off or competitor distribution), and evaluate S0–S4 on expected value: P(win) × value if awarded − bid costs and securities at risk.
3. Value commitments symmetrically: credit overbuild energy and BESS value, not only their costs. Decide the RtB price basis (MVA vs installed MW).
4. Use **one buyer hurdle** for both the RtB and COD exits. Derive the RtB price from buyer DCF (HTM block 14 already exists) instead of the hard 150 k€/MW, and show the gap.
5. Resolve the node-cap flag (100 MVA per node vs 200 MVA requested) before any bid recommendation.
6. Fix Run_Map F34/F35, refresh or remove the static tornado, investigate HTM G257 FAIL and Deposits E48 FAIL, align FX (365 vs 370), and verify the Aurora capture-above-baseload data.
7. Keep one visible input sheet. The meeting file's split between `Bora` N5:N16 and hidden `Calc` section c is a trap.
