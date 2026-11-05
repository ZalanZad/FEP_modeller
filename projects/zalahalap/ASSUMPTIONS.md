# Zalahaláp – assumption register and offer / Aurora term sheet

For discussion with the user, 28-Oct-2026. Values read from the built model (`out/Zalahalap_Model_recalc.xlsx`,
`Inputs` Live column, base set), not from the SPEC. Status: **U** confirmed by the user · **v3** from the original model
`Zalahalap_model_10112025_v3` · **T/O** from the teaser / the offers / the JV sheet · **A** Aurora Q2-2026 · **P**
placeholder (Bana wind terms or a default). "Rec." = Fable's recommendation where the value is questionable.

## 1. Assumption register

### 1.1 Timeline
| # | Assumption | Value | Unit | Source | Status | Rec. |
|---|---|---|---|---|---|---|
| 1 | Acquisition / JV start, deposits paid | 30-Nov-2025 | date | v3 Stat_inp E39 | v3 | ok |
| 2 | Ready to build = model date | 31-Dec-2026 | date | Teaser p6 "RTB Dec-2026"; user 04-Nov-2026 (v2.1) | U | spend before it is "spent to date": dated in the RtB block, one grid column "to 2026" |
| 3 | RtB sale closing (exit) | 31-Mar-2027 | date | placeholder (teaser transaction timeline shows "Closing Aug-26", inconsistent with RTB Dec-26) | P | confirm; drives the 12 % accrual and the FEP IRR |
| 4 | NTP; construction capex | 01-Jan-2027; all construction capex in 2027 (`Capex_Year`) | date | user 04-Nov-2026: "costs will arrive in 2027" | U | the NTP→COD spread is gone; E48/E49 at NTP (Jan-2027) |
| 5 | COD, PV and BESS together | 01-Mar-2028 (production from March: 10/12 of 2028); PPA Mar-2028 → Feb-2036 (8 years, v2.2) | date | user 04-Nov-2026: "production starts 2028 march" | U | Low: COD +12 months (Mar-2029), capex still in 2027 |
| 6 | Operating life | 35 | years | v3 Control D30 | v3 | long for PV (30 typical); BESS retires/repowers 2048 anyway |
| 7 | Pre-RtB development cost window | Oct-2026 → Mar-2027 | dates | v3 Control D116:D117 | v3 | ok |
| 8 | v3 case dates (column E only) | acq 30-Nov-25, RtB 31-Jul-26, exit 31-Oct-26, NTP 31-Mar-27, COD 31-Mar-28 | | v3 Control D19:D31 | v3 | reconciliation only |
| 9 | Timeline | annual, year-end dates: "to 2026" \| 2027 … 2064 (v2.1 §12.1) | | format | design | dated RtB flows keep exact dates |

### 1.2 PV
| # | Assumption | Value | Unit | Source | Status | Rec. |
|---|---|---|---|---|---|---|
| 10 | PV capacity | 63 (v3: 61) | MWp | Teaser p1 | T | ok |
| 11 | Grid export capacity | 49.9 (v3: 50) | MVA | Teaser p1 | T | ok |
| 12 | Specific yield | 1,265 (v3: 1,250) | kWh/kWp | Teaser p1 (79.7 GWh) | T | the offerors assume less: E.ON P50 73,141 MWh on 61 MWp = 1,199 (p3); MET 235,368 MWh / 3 y on 63 MWp = 1,245 (p3). **Use 1,245 until the yield report; keep 1,265 as High** |
| 13 | PV availability | 99.5 % | % | v3 Control D56 | v3 | ok |
| 14 | Degradation | 0.3 %/yr | % | v3 Control D57 | v3 | ok (0.4 % would be prudent for bifacial) |
| 15 | Clipping / curtailment, per case (Dashboard lever) | PV only 2 % (v3 fix 5: 63 MWp on 49.9 MVA, no battery); BESS cases 0 % (DC-coupled BESS absorbs the clipped energy) | % | v3 fix 5; Aurora s+29 = 0; placeholder | P | fix round 03-Nov-26; depends on DC coupling being confirmed (§1.3 #20); yield report |
| 16 | Guarantee of origin price | 1.75 real 2025 | €/MWh | v3 Control D58 | v3, P | likely high for 2028+; E.ON offers to buy at 90 % of the HUPX GO auction price (p4) – use that under E.ON; otherwise 1.0 as Base |
| 17 | PV sent to the BESS | `MIN(generation, Aurora charge-from-PV per MW × 50 MW)` ≈ 50.6 GWh of 79.3 (4 h) / 24.7 GWh (2 h) | MWh | Aurora s+13 | A | consequence of the 1:1:1 block; fine |

### 1.3 BESS
| # | Assumption | Value | Unit | Source | Status | Rec. |
|---|---|---|---|---|---|---|
| 18 | BESS built, power | yes, 50 | MW | Teaser p1 | T | ok |
| 19 | Duration / energy | 4 h / 200 MWh in all cases | h | Teaser "up to 200 MWh"; v3 Control D93 | T, P | **offers are for 2 h (E.ON 100 MWh) and 1.8 h (MET 90 MWh): run the offer columns at 2 h, Aurora at 4 h, and re-quote MET/E.ON for 200 MWh** (SPEC_v2 §2.3) |
| 20 | Coupling / grid charging | DC-coupled, no grid import | – | user 22-Oct-2026 | U | confirm DC coupling with the EPC concept (drives #15 and the Aurora block choice) |
| 21 | Aurora block | "2027 solar DC 1.0cy 4h" (Degraded batteries, 1 MW PV : 1 MW BESS : 1 MW grid) | – | Aurora | A | per MW of BESS as is; ZHLP is 1.26 : 1 : 1 |
| 22 | PV-only factor (no grid charging) | 0.779 (2027), 0.833 (2028), 0.941 average | factor | aurora_extract.py | U/A | ok |
| 23 | aFRR-down / FCR capacity deliverable from PV state of charge | 0 % of the 0.17 (2027) / 0.12 (2029) share | % | placeholder | P | worth 12–17 % of the stack; ask the BESS integrator / E.ON; 50 % as High |
| 24 | RTE / availability | 0.86 / 98 % (inside the blocks) | | Aurora Inputs C17:C18 | A | E.ON assumes 87 % |
| 25 | Haircut on merchant BESS revenue | 0 % (Low: 20 %) | % | – | P | ok |
| 26 | BESS capex | 110 | k€/MWh | v3 Control D94 | v3, P | an old figure; get a 2026 quote for 2 h and 4 h DC-coupled; Low 200 |
| 27 | Augmentation | 110 k€/MWh, indexed, at the block's repower year 2048; "repowered" = yes | | v3 Control D100; Aurora s+1 | v3, A | ok; "no repower" is an Inputs flag |
| 28 | BESS fixed opex | 18.70 k€/MW + 6.07 k€/MWh a year (2 h: 30.8 k€/MW; 4 h: 43.0 k€/MW) | k€/yr real 2025 | Aurora Inputs rows 37–39, 2028 entry (2 h 30.83, 4 h 42.96 €/kW); v3 Control D96 = the 2 h figure | A | fix round 03-Nov-26: v3 applied the 2 h opex to the 4 h battery; Aurora's 10 % co-location saving not applied – O&M quote |
| 29 | Grid charge on charging (offers that deduct it) | 0 | k€/MW | – | P | 0 is right without grid import |
| 30 | BESS tax depreciation | 10 | years | placeholder | P | tax adviser |

### 1.4 Revenue – Aurora (Central only, real 2025 €, calendar years 2026–2060, flat real after; nominal with Aurora's euro index, +2 % after 2060)
| # | Assumption | Value | Unit | Source | Status | Rec. |
|---|---|---|---|---|---|---|
| 31 | Market source | Aurora Q2-2026 HUN flexible-energy databook, Central; no Low/High (price risk via ±% cases) | | user | U | ok |
| 32 | Baseload (info) | 119.3 / 110.0 / 103.1 / 94.9 (2027–2030) | €/MWh real | Central prices!20 | A | |
| 33 | PV direct-export price, DC 4h block | 107.5 / 97.4 / 90.6 / 82.5 (2027–2030) | €/MWh real | (s+34 + s+35) ÷ s+31 | A | high capture because the battery takes the midday energy |
| 34 | PV direct-export price, DC 2h block | 83.4 / 74.8 / 70.4 / 62.2 | €/MWh real | same, s = 96 | A | |
| 35 | PV transfer price (energy to the BESS) | 36.6 / 36.1 / 36.0 / 29.8 (4 h); 33.0 / 33.1 / 33.5 / 31.0 (2 h) | €/MWh real | s+36 ÷ s+32 | A | |
| 36 | Blended PV value (63 MWp, 50 MW BESS), 2029 nominal | 60.4 (4 h) / 63.8 (2 h) | €/MWh | model | A | ≈ 0.54 × baseload – the honest capture |
| 37 | BESS storage total (uplift, net of the PV energy's export value) | 187.0 / 163.1 / 126.2 / 110.6 (4 h); 140.4 / 118.2 / 89.0 / 78.6 (2 h) | k€/MW/yr real | s+27 | A | × PV-only factor → 2029 nominal 120 (4 h) / 71 (2 h) |
| 38 | PV price change / BESS haircut | 0 / 0 (Low −10 % / 20 %) | % | – | P | ok |
| 39 | Route-to-market fee (merchant) | 0 % | % | v3 Control D77 | v3 | a merchant PV+BESS needs an aggregator: 1.5–3 % of revenue, or the E.ON/MET fee structure – set 2 % in Base |

### 1.5 Offers (detail in §2)
| # | Assumption | Value | Source | Status | Rec. |
|---|---|---|---|---|---|
| 40a | PV PPA in the model | **per revenue structure (v2.4): PPA uplift 55 / 80 / 85 €/MWh (PV only / 2 h / 4 h); Fixed + share 45 (all cases); Merchant 80 (all cases)**; nominal flat, 8 years from COD (Mar-2028 → Feb-2036), then Aurora merchant; pay-as-produced on all PV output, no balancing in the PPA years | Revenue "Commercial structures" table, picked by the Dashboard lever | U | user 06-Nov-2026 (SPEC §13); 80/85 are above every offer received (43.96–45 fixed) – need a counterparty |
| 40b | Revenue structure (per case lever) | **PPA uplift** (PV PPA carries the battery: BESS basis None) · **Fixed + share** (default) · **Merchant** (BESS on Aurora central); the per-case PPA price, PPA tenor, BESS basis and fixed terms on the Dashboard are black links into the table | Dashboard row "Revenue structure"; Revenue table | U | SPEC §13.1; Tolling stays selectable by typing it into the table's BESS basis row |
| 40c | BESS basis "None" | no charging (all PV exported, PV priced on the PV-only block), no BESS revenue; the battery is still built, paid for, operated and augmented in 2048 | Inputs "BESS revenue basis options" | U | user 06-Nov-2026; the PPA-uplift 4 h case therefore loses money (hold IRR −18.6 %, Low: no IRR exists, shown n/a) |
| 40d | Fixed + share terms | BESS fixed 70 €/MWh (2 h) / 55 €/MWh (4 h) on discharge up to 1 cycle/day (cap = cycles × MWh × 365 × operating share × in service); above the cap 70 % of Aurora's realised €/MWh (merchant uplift ÷ Aurora discharge per MW); Aurora merchant after the fixed tenor | Revenue table (blue) | U | user 06-Nov-2026; with PV-only charging the battery cycles ≈ 0.6/day, so the above-cap tier is 0 in Base |
| 40e | Fixed tenor (BESS) | 8 years from COD, equal to the PPA tenor | Revenue table "Fixed tenor" | P | the user did not give a BESS tenor (SPEC §13.1: 8 unless the user says otherwise) – confirm |
| 40f | IRR start guess / n/a | Calc column H: MIN(10 %, MOIC − 1); buyer at RtB MIN(target, MOIC − 1); buyer at COD the target; an IRR with no root shows n/a (Checks: allowed only when MOIC < 1) | Calc returns rows | model | Excel's XIRR returns ~0 when Newton fails; the target start picks the 8 % root where the PPA-uplift 4 h buyer flows have three (−18 %, 6.8 %, 8 %) |
| 40 | Case C pairing today | MET fixed 43.96 (3 y) + MET tolling 49.7 on a **4 h** battery | Inputs F67:F68 | P | replace by one column per offeror, 2 h for E.ON/MET (SPEC_v2 §2.3) |
| 41 | Offer start / tenor | E.ON: COD + 3 y; MET: 01-Jan-2028, 3 y (4 y option not selectable); MVM: 01-Jan-2028, 5 y | Curves rows 155–156, 168–169 | O | add the 4 y lever |
| 42 | Balancing cost under an offer | 5 €/MWh still charged | Calc opex_bal | bug | **double count** – the fees include balancing; charge it only outside the tenor |
| 43 | E.ON handling fee on discharged MWh | not charged | Calc | omission | charge 7.6 % (min 5 €/MWh) on discharge revenue – ≈ 0.4–0.6 m€/yr at 50 MW |
| 44 | Profit-share base (E.ON A/B, MET PSA) | Aurora uplift net of the PV energy's export value (s+20) | Calc stack | reading | keep; gross reading as a switch; confirm with E.ON/MET |
| 45 | MET pay-as-produced settlement | PV exported directly at 43.96; energy sent to the BESS at the transfer price (≈ 38 nominal) | Calc pv_rev | reading | PaP on **all PV output** at the PV meter (MET dispatches the tolled BESS): +≈ 0.3 m€/yr under fixed, −0.3 under indexed |
| 46 | MVM scope | 45 €/MWh on PV output only; BESS merchant on top | Curves row 151 | reading | **ask MVM**: PV meter or POD, and whether the price includes the BESS optimisation |

### 1.6 Capex (EUR k, 2025 prices unless noted)
| # | Assumption | Value | Source | Status | Rec. |
|---|---|---|---|---|---|
| 47 | PV equipment + BoP | 350 k€/MWp (330 + 20) | v3 Control D82:D83 | v3 | low end for a 2027 build incl. 12 km export route works; EPC indication needed; Low +10 % |
| 48 | MV cable to the Zalahaláp substation (12 km, 35 kV) | 3,700 | v3 Control D104 | v3, P | quote (§16 Q13) |
| 49 | HV substation / bay | 9,500 | v3 Control D105 | v3, P | quote; the teaser says the connection fee is zero (2018 rights) – confirm no MAVIR contribution |
| 50 | Land to COD (not depreciated) | 620 | v3 Control D106 | v3 | ok |
| 51 | Grid connection fee | 0 | Teaser p1 | T | ok |
| 52 | BESS | 50 × 4 × 110 = 22,000 | #19, #26 | P | 11,000 at 2 h |
| 53 | Pre-RtB development costs (reimbursed by the buyer) | 3,182 | v3 Control D115, D118 | v3 | reimbursement on top of the 90 k€/MWp premium – a commercial assumption to confirm |
| 54 | Early works | four lines 785 k€ (land 1st/2nd 60+60 mHUF, ICS design 25 k€, critical equipment 460 k€) paid in development, reimbursed, earning 12 %; E48/E49 (650 + 75 k€) to construction capex at NTP | v3 Stat_inp E44:E49; user 22-Oct | U (provisional) | to be revisited as a whole |
| 55 | Capex change | 0 % (Low +10 %) | – | P | ok |
| 56 | Construction spread | straight line NTP → COD | – | design | ok |

### 1.7 Opex (real 2025, indexed with Aurora's euro index)
| # | Assumption | Value | Source | Status | Rec. |
|---|---|---|---|---|---|
| 57 | O&M, all-in | **10 k€/MWp/yr** (PV, BoP, MV cable, insurance, management), real 2025 × euro index | user 05-Nov-2026 (v3 Control D59: 14.96) | U | fix round v2.3; confirm with an O&M quote |
| 58 | MV cable O&M | 0 – inside the all-in 10 k€/MWp/yr (row kept) | user 05-Nov-2026 (v3 Control D62: 5) | U | – |
| 59 | Balancing | 5 €/MWh (merchant only after the fix #42) | v3 Control D60 | v3 | ok |
| 60 | Land lease | 96 k€/yr from COD | v3 Control D61 | v3 | ok |
| 61 | HIPA + innovation contribution | 2.0 % + 0.3 % of revenue | Bana | P | confirm Nyirád's rate; base is revenue less materials/subcontracting |
| 62 | Opex change | 0 % | – | P | – |

### 1.8 Financing (levered case, v2.1 §12.2: user instruction 04-Nov-2026, not a bank term)
| # | Assumption | Value | Status | Rec. |
|---|---|---|---|---|
| 63 | Gearing | 60 % of the capex to COD (total uses incl. land and E48/E49; not development or deposits), per case lever; 0 = unlevered | U | user 04-Nov-2026; banks size merchant PV+BESS on DSCR – term sheet |
| 64 | Sizing DSCR | none: debt = gearing × capex; min / average DSCR shown as information (PV only min 0.98 in Base) | U, P | a bank would cap the debt by DSCR |
| 65 | Tenor / rate | 20 level annuities from 31-Dec-2028 (first one full) / 5.0 % all-in; straight-line not built | U | rate = swap + margin – ask the bank |
| 66 | Fees / DSRA / sweep / IDC | none / none / none; IDC on mid-year 2027 draws capitalised; interest deductible (own loss carry-forward) | P | add with a term sheet |
| 67 | Debt per MWp cap | none (gearing only) | P | – |

### 1.9 Tax (Hungary)
| # | Assumption | Value | Source | Status | Rec. |
|---|---|---|---|---|---|
| 68 | CIT | 9 % | v3 Control D66 | v3 | ok; Pillar Two 15 % only if the group is in scope (> 750 m€ revenue) – confirm |
| 69 | Robin Hood tax (energy suppliers' income tax, 31 %) | off | placeholder | P | **a licensed producer > 50 MW may be in scope; tax adviser – potentially the largest tax item** |
| 70 | Tax depreciation | PV, grid, development 30 y; BESS 10 y | v3 Control D65 | v3, P | 30 y is an accounting life; Hungarian tax rates for machinery are faster (14.5 %/yr class) – tax adviser; faster depreciation lifts the levered FEP return |
| 71 | Loss carry-forward | 50 % of the tax base; losses lapse 5 years after NTP (one date) | v3 Control D102; Tao. 17. § | v3, simplification | ok |
| 72 | Tax on the RtB/COD gain | share deal, participation exemption (gain exempt); FEP pays 9 % on the 12 % return as interest | placeholder | P | the exemption needs the registered participation ("bejelentett részesedés") and a 1-year holding – confirm; asset deal = 9 % on the gain |
| 73 | Unlevered views | unlevered tax (no interest), own loss carry-forward | design | – | ok |

### 1.10 Exit and valuation
| # | Assumption | Value | Source | Status | Rec. |
|---|---|---|---|---|---|
| 74 | RtB price | 90 k€/MWp premium **on top of** reimbursements (pre-RtB costs 3,182, early works 785, deposits ≈ 1,776 → all-in ≈ 162 k€/MWp) | v3 Stat_inp E24; JV sheet 120 | v3, P | the model's implied premium at an 8 % buyer IRR is 34 k€/MWp (Aurora 4 h) and 0 under the offers – **decide the price basis (premium vs all-in) and the target** |
| 75 | COD price | implied at the buyer target IRR (EV) | – | P | ok |
| 76 | Buyer target IRR, RtB and COD | 8.0 % unlevered after tax | v3 Control D9 (NPV rate) | P | merchant HU PV+BESS buyers: 8.5–9.5 %; 9 % Base |
| 77 | FEP discount rate | 8 % | v3 Control D9 | v3 | ok |
| 78 | FX | 400 HUF/EUR (Aurora 397.77) | v3 Stat_inp E23 | v3 | ok |
| 79 | Sale costs (RtB / COD) | 0 / 0 | placeholder | P | 1.5 % of the price (≥ 200 k€) |
| 80 | Transaction costs (FEP) | 200 | v3 Stat_inp E37 | v3 | ok |
| 81 | Bank-guarantee cost | 1M EURIBOR 1.87 % + 1.75 % margin on guarantee balances | v3 Dyn_Inp G50:Y50, Stat_inp E110 | v3 | ok |
| 82 | Deposits and guarantees (HUF m) | connection-fee reimbursement 45; capacity guarantee 180; additional guarantee 485.55 (+ instalments 242.775 / 242.775 / 80.925) – returned at the exit (buyer reimburses) or at COD | v3 Stat_inp E84:E106 | v3 | ok; the v3 phantom 485.55 (fix 3) stays off |

### 1.11 JV (PSH)
| # | Assumption | Value | Source | Status | Rec. |
|---|---|---|---|---|---|
| 83 | Futureal share of the distributable | 60 % at ≤ 90 k€/MWp, 50 % above (step; v3: slide to 50 % at 100) | user 22-Oct | U | ok |
| 84 | Price basis for the threshold | price per MWp | user | U | the JV sheet uses 120 k€/MWp × 63 (= 7.56 m€); at 120 the step gives 50 % – confirm which price the JV discussion assumed |
| 85 | FEP SPV acquisition price | 549 (v3: 150, "connection rights") | JV sheet G19 | O, P | what is it? deducted before the split, credited to FEP |
| 86 | DSA loan to PSH | 15 k€/MWp = 945 | JV sheet F55 | O, P | repaid to FEP before the split; confirm |
| 87 | Non-devex loan 562.5 k€ (4,500,000 × 50/400) | **not modelled** | JV sheet G29 | O | explain; add if real |
| 88 | FEP return on loans and costs | 12 % p.a., act/365, capitalised 31-Dec (v3: flat / monthly) | JV sheet G30; v3 E69 | O, v3 | ok |
| 89 | PSH land apport | 150 in kind; not repaid before the split (v3: repaid with return) | v3 Stat_inp E34 | v3 | confirm |
| 90 | PSH own development costs (for PSH's IRR memo) | 15 k€/MWp | v3 Stat_inp E79 | v3 | memo only |
| 91 | COD sale split / hold | same step rule at COD; hold = FEP buys PSH out at the RtB entitlement | placeholder | P | the JV sheet covers the RtB sale only – confirm |

## 2. Offer and Aurora term sheet (as stated in the PDFs; everything indicative)

### 2.1 PV products
| Term | E.ON (Offer_EON_indicative_Futureal_PV.pdf, 09-Oct-2026, valid to 12-Dec-2026) | MET / E2 Hungary (Offer_E2_…_20261020.pdf, 20-Oct-2026) | MVM ONEnergy (e-mail 22-Sep-2026, scanned) |
|---|---|---|---|
| Product | spot-indexed pay-as-produced PPA on HUPX DAM 15-min (p4 §5.1) | fixed-price PaP (p3) **or** HUPX DAM-indexed PaP (p4), physical delivery into MET's balance group | fixed-price PaP "for the PV+BESS co-location near Nyirád" (one line) |
| Price | payable = DAM − handling fee; **fee = MAX(7.6 % × DAM, 5 €/MWh)** (p4; printed "7,6%") | **43.96 €/MWh** 01-Jan-2028 → 31-Dec-2030 (est. 235,368 MWh); **40.76** → 31-Dec-2031 (313,824 MWh) (p3); indexed: fee = MIN(10, MAX(5 % × DAM, 5.75)) (p4) | **45 €/MWh, 2028–2032** (the 2029–2034 series 96.88 → 82.82 is for a wind+BESS site) |
| Tenor / start | max 3 years (p1); COD 2027 Q4–2028 Q1 | 3 or 4 years from 01-Jan-2028 | 5 years (2028–2032) |
| Volume | 100 % of output; economic-curtailment option at ≤ 0 prices / DSO limits (p4) | 100 %; if DAM − fee ≤ 0 the reference volume is 0 (no penalty); downregulation compensation 10 €/MWh (p4) | generated volume; "takes the BESS optimisation effect into account" (said of the wind series) |
| Fees / other | balance-group fee 250 €/month/POD; GoO purchase option at 90 % of the HUPX GO auction price; payment 30 days (p4) | balancing, forecasting, scheduling, settlement by MET (p3) | scheduling and balancing borne by MVM; final price depends on the distance to the connection point |
| Indexation | n/a (spot) | none stated | none stated |
| Settlement point (our reading) | POD: all fed-in energy (p6 "PPA: physical offtake at the connection point") | PV meter: pay-as-produced on PV output (BESS tolled separately) | unclear – ask |
| Offeror's yield | P50 73,141 MWh/yr on 49 MW / 61 MWp (p3) | 78,456 MWh/yr on 63 MWdc (p3) | – |

### 2.2 BESS products
| Term | E.ON "Standalone BESS Flexibility and Optimization Agreement" | MET | MVM |
|---|---|---|---|
| Configuration offered | 50 MW / **100 MWh** (p1; p3 prints "10 MWh"), RTE 87 %, DoD 100 %, 2 cycles/day, COD 2027 Q4–2028 Q1 | 50 MW / **90 MWh**, 2 cycles/day, "green", start 01-Jan-2028 (p5–6) | none |
| Option A – floor + tiered share (2 h) | floor renegotiated every 2 years; **2027 bands** (€/kW/yr): ≤ 47 guaranteed, 0 % fee; 47–52 100 % to E.ON; 52–80 10 %; 80–109 15 %; > 109 18 %. **2028**: 38 / 43 / 67 / 91, same fees. Example: 120 → owner 105.87 (p5). Years without a floor: option B | **Floor + PSA**: floor **45.6 k€/MW/yr (3 y)** / **41.9 (4 y)**, upside 50 % Futureal / 50 % MET (p6) | – |
| Option B – full profit share | owner 90 % / E.ON 10 % of Net Optimization Profit (p5) | **Full tolling**: **49.7 k€/MW/yr (3 y)** / **46.4 (4 y)** (p5) | – |
| Profit definition | market revenues (DAM, IDM, balancing, aFRR/mFRR capacity + activation, imbalance) − direct energy procurement (grid charging, balancing energy) − regulatory charges (RHD, system fees, taxes) (p4–5) | not stated | – |
| Tenor | ≥ 2 years, can exceed 2; max 3 (p1, p5) | 3 or 4 years | – |

### 2.3 What each case earns per year (nominal; illustrative, from the CSV with the v2 rules; 2 h block for E.ON/MET)
| Year | Baseload | PV direct, 4 h | PV direct, 2 h | E.ON PPA (2 h) | MET index (2 h) | MET fixed | MVM | BESS uplift kept, 4 h | kept, 2 h | E.ON A/B owner | MET tolling | MET floor + PSA |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| | €/MWh | €/MWh | €/MWh | €/MWh | €/MWh | €/MWh | €/MWh | k€/MW | k€/MW | k€/MW | k€/MW | k€/MW |
| 2028 | 116.7 | 103.4 | 79.4 | 73.3 | 73.6 | 43.96 | 45 | 144 | 89 | 73–75 | 49.7 | 67 |
| 2029 | 111.6 | 98.1 | 76.2 | 70.4 | 70.5 | 43.96 | 45 | 120 | 71 | 59 | 49.7 | 58 |
| 2030 | 104.9 | 91.2 | 68.8 | 63.5 | 63.0 | 43.96 | 45 | 114 | 68 | 57 | 49.7 | 57 |
| 2031 | 111.0 | 93.4 | 71.5 | merchant | merchant | merchant (40.76 if 4 y) | 45 | 116 | 68 | merchant | merchant (46.4 if 4 y) | merchant |
| 2032 | 110.5 | 88.6 | 67.6 | | | | 45 | 123 | 72 | | | |

Blended PV value on all output (direct + stored), 2029: 60.4 €/MWh (4 h) / 63.8 (2 h) – the fixed offers (44–45) are
~30 % below Aurora Central; the indexed offers give up the fee only. BESS: the 4 h merchant uplift (120 k€/MW) is not
comparable with 2 h products (49.7–75); on the 2 h block the gap is 71 vs 50–75.

### 2.4 Discrepancies and simplifications in the current model (to fix in v2)
1. **4 h battery vs 2 h offers**: MET tolling 49.7 and the E.ON bands are applied to a 200 MWh battery (fee not scaled;
   the extra 110 MWh earn nothing) – the C case is penalised by construction. v2: hours per column.
2. **Balancing double count**: 5 €/MWh charged on top of fees that include balancing (all three offers).
3. **E.ON handling fee on discharged MWh** missing (the PPA settles all POD energy).
4. **MET 4-year options** (40.76 / 46.4 / 41.9) not selectable; E.ON option A vs B and MET tolling vs floor only via sets.
5. **MET pay-as-produced on stored PV energy**: the model pays the transfer price (≈ 38) instead of the PaP price on all PV output.
6. **MVM scope** unverified (PV meter vs POD; BESS inside the 45?).
7. **Profit-share base** (E.ON A/B, MET PSA): net uplift (Aurora storage total) – the gross reading would raise E.ON's
   fee by ≈ 4–6 k€/MW/yr; confirm with the offerors.
8. **GoO under E.ON** at 90 % of the HUPX GO auction vs 1.75 €/MWh flat.
9. E.ON floor after 2028 (renegotiated every two years) – modelled as none (option B); flag `EON_Bands_After_2028`.
10. The fee kinks (MAX/MIN) act on 15-min prices; the annual DAM proxy (Aurora direct-export price) smooths them.
11. E.ON 2027 bands never apply (COD Mar-2028); 2028 bands apply for nine months.
12. MET "if DAM − fee ≤ 0, volume = 0" and the 10 €/MWh downregulation compensation are not modelled (0 curtailment).
