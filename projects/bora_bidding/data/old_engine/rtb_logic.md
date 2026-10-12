# Old Bora engine: RtB side, logic map

Source: `reference/Bora_scenarios_meeting.xlsx` (meeting file, saved 11-Oct-2026), read as data. Sheets covered: `Bora` (front end), `Tender` §1–4,
`Inputs` (case 1 = RtB sale), `Deposits`, `Devex_FEP`, `BORA_Devex` (pasted developer file), `MOIC` (monthly RtB engine), `Run_Map`, `Calc` row 305 / data table rows 306–330,
plus the HTM rows that the price effect reads. This version differs from `docs/reference_analysis/bora_models.md` (older engine): **net MVA is 100, not 200**, the RtB
sale is **31-Dec-2029**, and the strategies are the meeting's S0–S4 (below). The workbook wins wherever the two disagree.

Shadow: `old_rtb.py` (stdlib) reproduces every Bora RtB output for S0–S4 to ≤ 4e-14 relative against the LibreOffice live values (IRR to 4e-16) and to ≤ 5e-15 against the
saved Excel values. The only exception is the IRR, where Excel's XIRR stops at ≈ 5e-9 and fails outright for S2 and S4 (see bug B1). Run `python3 old_rtb.py`.

Units: k€ unless stated. mHUF = HUF million. The Bora sheet shows €m (= k€/1000). Bora sign convention: proceeds +, deductions −. MOIC cash-flow blocks: + = inflow to the party.

---

## 0. Strategies and run codes

| Bora col | Strategy | Tender col | Code (Tender E5) | RtB run (Calc E10) | ÖSZ-1 k€/MW | ÖSZ-2 €/MWh | Overbuild | BESS % (basis) | h | ÖSZ-6 | Partial / shared | EU turbine / transf. |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| D | S0 Zero bid | I | 4 | 21 | 0 | 0 | 0 % | 0 | 4 | 1.0 | 0 / 0 | 0 / 0 |
| E | **S1 Current bid** (price-effect reference) | E | 0 | 1 | 20 | 2 | 30 % | 30 % (net MVA) | 4 | 1.5 | 0 / 0 | 1 / 0 |
| F | **S2 BORA 1** | F | 1 | 6 | 50 | 3 | 51 % | 51 % (net MVA) | 4 | 1.0 | 1 / 1 | 1 / 1 |
| G | **S3 BORA 2** | G | 2 | 11 | 50 | 2 | 40 % | 40 % (net MVA) | 4 | 1.0 | 1 / 1 | 1 / 1 |
| H | S4 Max score | H | 3 | 16 | 50 | 3 | 51 % | 51 % (net MVA) | 4 | 2.1 | 1 / 1 | 1 / 1 |

Inputs are typed in `Bora!D6:H17`, then linked to `Calc!E23:I34` and on to `Tender!E12:I25`. `Calc!E10` = run code = strategy code × 5 + route (route 1 = RtB sale).
`Inputs!E6` = `INDEX(Calc!E306:E330, E10)` sets the case, and `Tender!E5` = `INDEX(Calc!D306:D330, E10)` sets the strategy. Tender column J holds the active strategy and feeds
`Inputs!G20` (BESS MW), `G21` (h), `G90` (PB uplift), `G233` (ÖSZ-1), `G234` (ÖSZ-2) and `G297` (installed MW). Every other RtB input is the same for all strategies.

---

## 1. Physical project per strategy (Tender §2, rows 37–47)

| Quantity | Formula (cell) | S1 | S2 | S3 |
|---|---|---|---|---|
| Net MVA | `Tender!E37 = Calc!E74` = **100** (typed) | 100 | 100 | 100 |
| Installed wind MW | `E40 = ROUND(E37 × (1 + overbuild), 6)` → `Inputs!E297`, `E29 = E18 × E297/E281` | 130 | 151 | 140 |
| BESS MW | `E41 = ROUND(ratio × CHOOSE(basis+1, net MVA, installed), 6)`, with basis `Bora!row 10` = **0 (net MVA) for every strategy** | 30 | 51 | 40 |
| BESS MWh | `E42 = MW × h` → `Inputs!E22` (gross) / `E23` (× conversion 100 %) | 120 | 204 | 160 |
| Bottleneck loss share | `E45 = ob/(1+ob) × Bora!N18 (5 %)` | 1.154 % | 1.689 % | 1.429 % |

`Inputs!E16` "gross capacity entered in tender" = **100, typed separately** from `Calc!E74`. `E18 = E16 × E17` (conversion 100 %) is the "successful net MW"
that drives the TSO security, the connection fees, the premium (on the net basis) and the municipal price effect. See O3.

## 2. Tender scoring (Tender §3, rows 51–88 and 91–172)

- x per item, rounded to 6 dp: ÖSZ-1 `= k€/MW × 1000`; ÖSZ-2 `= €/MWh`; ÖSZ-3 `= overbuild × 100`; ÖSZ-4 `= BESS ÷ (installed if basis = 1, else net MVA) × 100`; ÖSZ-6 `= multiple × 100`.
- Band number: `SUMPRODUCT((lower < x) + (lower = x) × included)` over the band table. Inner points `= INDEX(points incl. "below" row, band + 1) × eligible`.
  Weighted points `= inner × weight / 100`.
- Band tables (`Tender!D93:F150`): ÖSZ-1 lower bounds 10,000 … 45,000 exclusive, then 50,000 inclusive, giving 10/20/30/40/55/70/80/90/100.
  ÖSZ-2 2.0 … 2.9 exclusive, then 3.0 inclusive, giving 5/10/20/…/90/100. ÖSZ-3 and ÖSZ-4: > 2 → 10, then ≥ 5, 10, …, 40 inclusive → 20 … 90, and > 50 → 100.
  ÖSZ-6: > 100 → 10, then ≥ 120 … ≥ 200 → 20 … 100.
- Eligibility: ÖSZ-1 10,000 ≤ x ≤ 50,000 and ÖSZ-2 2 ≤ x ≤ 3 (below the minimum scores 0 and the bid stays valid). ÖSZ-3 needs x ≥ 0.
  ÖSZ-4 needs x ≥ 5, h ≥ 4 and BESS MW ≤ installed MW.
- Yes/no items: ÖSZ-5 = 100 × 5 %, ÖSZ-7 = 100 × 10 %, ÖSZ-8 = 100 × 5 %, ÖSZ-9 inner = 80 × EU turbine + 20 × EU transformer, × 15 %. ÖSZ-9 counts on EU-turbine routes only.
- Weights (`E153:E161`): 20/15/10/15/5/5/10/5/15.
- Totals: `E84 = ROUND(Σ ÖSZ-1…8 + ÖSZ-9, 2)` (EU routes), `E85 = ROUND(Σ ÖSZ-1…8, 2)` (CN routes).
- Results (Bora rows 26–36): S1 4/0/7/10.5/0/2.5/0/0/12 = **36** (CN 24); S2 20/15/10/15/5/0/0/5/15 = **85** (70); S3 20/0/9/13.5/5/0/0/5/15 = **67.5** (52.5);
  S4 90 (75); S0 0.
- Scoring does not feed the RtB economics. It only informs the choice of strategy (there is no award probability).

## 3. Timing (Inputs B/D/G, all case 1)

| Date | Formula | Value |
|---|---|---|
| Month 0 / investment | `Inputs!E27` typed | 31-Jul-2026 |
| COD | `Bora!N16 → Calc!E91 → Inputs!E298` | 30-Sep-2031 |
| Months to RtB | `E55 = (Y(COD)−Y(E27))×12 + M(COD)−M(E27) − NTP lag 3 (E152) − construction 24 (E153)` | 35 |
| RtB | `E56 = EOMONTH(E27, E55)` | 30-Jun-2029 |
| RtB → sale lag | `E28 = months(Bora!N17) − months(E56)` | 6 |
| Sale month / date | `E57 = E55 + E28`, `E58 = EOMONTH(E27, E57)` | 41 / **31-Dec-2029** |
| NTP / COD (derived) | `E154 = EOMONTH(E56, 3)`, `E155 = EOMONTH(E154, 24)` | 30-Sep-2029 / 30-Sep-2031 |
| Bid (fee, bid bond) | `E95` | 30-Oct-2026 |
| Award | `E96` | 13-Dec-2026 |
| PB posting | `E96 + E102 (17 days)` | 30-Dec-2026 |
| Connection request (TSO security) | `E97` | 30-Dec-2026 |
| HCSSZ (connection fees) | `E99` | 08-Jan-2029 |
| Interest pause | `E64` start / `E73` end | 30-Jun-2029 / 01-Jan-2030 |
| DSA start | `E70` | 26-Oct-2026 |

The MOIC grid is monthly. Month m is dated `EOMONTH(31-Jul-2026, m)`, for m = 0…60 (H:BP). Each payment is booked in grid month
`(Y(d)−2026)×12 + M(d) − 7` at that month's month end, and these month-end dates drive XIRR/NPV. The interest claims, by contrast, use exact payment dates (§6).

## 4. Deposits (securities and grid fees), rows 7–18

Gross H = unit rate F × capacity G (mHUF). Payable J = H + netting I. Then `M = J × 1000 / Inputs!E11 (FX 365)` in k€, and `O = M` (or `× E109` = 0 if paid after the sale).
P is cash (1) or guarantee (0), Q the % recovered from the buyer, U the share funded on the recoverable 12 % loan, and W the FEP rate (12 % cash, 6 % guarantee).

| Row | Item | Basis and formula | Date | P | Q | U | S1 mHUF / k€ | S2 mHUF | S3 mHUF |
|---|---|---|---|---|---|---|---|---|---|
| 7 | Participation fee | 3 mHUF (E81) × 2 bids (E82 ← Calc E80) | bid 30-Oct-26 | 1 | 0 | 0 | 6.0 / 16.44 | 6.0 | 6.0 |
| 8 | Bid bond | 0.5 % (E76) × E91; `E91 = E16 × E29/E18 × 406 + E22 × 200` (installed wind + BESS MWh) | 30-Oct-26 | 0 | 1 | 0 | 383.9 / 1,051.78 | 510.53 | 444.2 |
| 9 | Bid bond refund on failed capacity | `−0.5 % × E19 × 406 × E29/E18 × E110` | award | 0 | 1 | 0 | 0 | 0 | 0 |
| 10 | Performance bond | 2.5 % (E77) × uplift (E90 ← Tender J17) × E92 (`E29 × 406 + E23 × 200`); **I10 = −H8 × E107** (bid bond credited) | award + 17 | 0 | 1 | 0 | gross 2,879.25, payable 2,495.35 / 6,836.58 | 2,552.65 / 2,042.12 | 2,221.0 / 1,776.8 |
| 11 | TSO connection-request security | 4.5 mHUF/MW (E83) × E18 | 30-Dec-26 | 1 | 1 | 1 | 450 / 1,232.88 | 450 | 450 |
| 12 | Direct connection fee | E84 = `Calc!E82/1e6/Calc!E74` = 5.0835 mHUF/MW × E18 (equal to the 508.35 lump sum while E18 = Calc E74) | HCSSZ 08-Jan-29 | 1 | 1 | 1 | 508.35 / 1,392.73 | same | same |
| 13–17 | Supplementary security (4 instalments + refund at permit) | 406 × 0 % (E78, GYIK 70) × share × E29 | 23-Apr-27 … 30 | 0 | 1 | 1 | 0 | 0 | 0 |
| 18 | Indirect connection fee | 73.071306 mHUF/MW (E85) × E18 **− J11** (TSO security credited) | HCSSZ | 1 | 1 | 1 | gross 7,307.13, net 6,857.13 / 18,786.66 | same | same |

Totals (S1): cash 21,428.70 k€ (fee 16.44 + TSO 1,232.88 + direct 1,392.73 + indirect 18,786.66). Guarantees 7,888.36 k€ (bid 1,051.78 + PB net 6,836.58 = PB gross).

Release and recovery:
- Guarantees are issued in the month of the date. They stay outstanding until the sale month and are then released (MOIC row 75); the buyer replaces them.
- Cash posted on or before the sale is recovered from the buyer at Q % (MOIC rows 87–99), i.e. 100 % except the fee.
- No security is refunded or forfeited before the sale (failure is not modelled). The TSO security is consumed against the indirect fee (row 18 netting), and the buyer reimburses the gross amount through rows 11 + 18.

FEP-convention return per line (Deposits X/Y, the same convention as BORA_Devex P):
```
seg(r, a, b) = 1                                                          if a > b
             = (1 + r·(31-Dec(Y(a)) − a + 1)/365) · (1+r)^(Y(b) − Y(a) − 1) · (1 + r·(b − 1-Jan(Y(b)) + 1)/365)
X = seg(W, pay, MIN(sale, pause start)) × [1 if sale < MAX(pay, pause end) else seg(W, MAX(pay, pause end), sale)]
Y = O × (X − 1)
```
Act/365 stubs count both ends, compounding is annual at 31-Dec, and whole years are (1+r) regardless of leap days. With the sale on 31-Dec-2029, which is before the pause end of 01-Jan-2030,
the second segment is always 1, so the pause acts as a **freeze at 30-Jun-2029** (the devex file's "freeze" variant).
- S1 factors: fee and TSO security (12 %) 1.356573 and 1.329919; bid bond and PB (6 %) 1.169013 and 1.157411; connection fees 1.057330 (the same-year quirk, B4).
- Sums (S1, k€): `Y22` guarantees 6 % = **1,253.92**; `Y23` cash excluding the fee, 12 % = **1,563.63** (TSO 406.75 + direct 79.84 + indirect 1,077.04); `Y24` fee 12 % = **5.86**.

## 5. Devex (Devex_FEP rows 5–38, from the pasted BORA_Devex)

- **Source:** `BORA_Devex` is a paste of **BORA_Devex_20260925.xlsx** (200 MW wind, 100 MW / 200 MWh BESS, FX **370**, AF6:AF18). It is **not** the 02-Oct-2026 method file.
  `Devex_FEP!E = BORA_Devex!J/1000 × scaling Inputs E294 (= 1)` and `F = BORA_Devex!E` (payment date). The continuous items (legal 20, financial 80, other 50 k€) are typed on 31-Dec-2027.
- **Flags:** G = FEP pays (devex lines), J = buyer reimburses at exit (land: rows 14, 15, 48). "Other costs (HR, car)" 50 k€ has G = 0 and J = 0, so it is excluded.
- **H = E × G** (FEP-borne principal), **K = E × J × (1 − G)** (land, buyer-reimbursed), **M = H + K** (principal earning 20 %).
- **Contingency** row 38 = `10 % (E68) × (ΣH + ΣK + Deposits!O7)` = 10 % × (2,380.26 + 425.38 + 16.44) = **282.21 k€**. It is dated 31-Dec-2028 and FEP-borne.
  Note that the base includes land and the participation fee.
- **Totals:** FEP devex principal (incl. contingency) ΣH = **2,662.47**; land ΣK = **425.38** (reservation 55.65 + first payment 123.24 on 13-Feb-27 + RtB payment 246.49 on 24-Mar-29). All strategies are the same.
- **Multiple** per line: `S = MAX(seg(20 %, pay, MIN(sale, pause start)) × seg2, MOIC!E131 = 2)`. The highest growth factor is 1.66× (row 20, paid 01-Oct-26), so **every line is at 2.0×**.
- **Waterfall claim** `MOIC!G131 = Σ_months≤sale (devex + land drawdown) × (principal-weighted S − 1)` = Σ M × (S−1) = **3,087.85** (= 1 × principal).
- The per-line HUF amounts (devex_lines.csv, `huf_native`) are divided by 370. The engine FX of 365 is never applied to devex (O1).

## 6. MOIC: the monthly RtB engine

**Block 1–2: devex and DSA.**
- Row 18 = Σ H by grid month, and row 20 = Σ K by grid month (E67 = 1).
- DSA (rows 22–25): `closing/MW = MIN(cap 10 (E37), 2.5 k€/MW/yr (E36) × N(date ≥ 26-Oct-26) × YEARFRAC(26-Oct-26, MIN(date, sale), basis 0 = US 30/360))`.
  `G25 = closing at sale × E29` (installed MW).
- YEARFRAC(26-Oct-2026, 31-Dec-2029) = 1,145/360 = 3.180556, which gives **7.951 k€/MW**: S1 1,033.68, S2 1,200.66, S3 1,113.19 k€. The cap does not bind, and the DSA is not paused.

**Blocks 3–5: securities, cash vs guarantee.**
- Rows 28–39 place Deposits O in its grid month. Cash lines (× P) go to rows 44–56 and guarantee lines (× (1−P)) to rows 60–72.
- Guarantee balance rows 73–76: opening + issued, released at the sale.
- **Guarantee fee (cost to FEP)** `H79 = (opening + issued) × 6 % (E113, flat, E114 = 1) / 12`. Every month counts in full, from the issue month to and including the sale month.
  S1: bid bond 39 months + PB 37 months = **1,469.86 k€** (S2 1,307.80; S3 1,137.88).
- Memo row 82: monthly-compounded 6 % "guarantee interest" with the accrual flag (row 10). It is **not** used in the waterfall.

**Block 6: recovery** `G99 = Σ cash posted ≤ sale × Q` = **21,412.26** (TSO + direct + indirect, 100 %). `G144 = G99 + G20 (land)` = **21,837.64** (Bora row 50).

**Blocks 7–8: shareholder loans** (FEP funds 100 %, E35 = 0).
- **Recoverable loan** = cash lines × U, i.e. the lines refunded if the project fails: TSO security, direct fee, indirect fee. Drawdown G110 = **21,412.26**.
- **Unrecoverable (at-risk) loan** = the cash fee (16.44) + devex + land (3,087.85).
- Monthly memo schedules: recoverable interest row 111 (`((1.12)^(1/12) − 1) × (open + draw) × flag`, G111 = 1,595.81); at-risk interest row 126 (0, because flat 2×, E66 = 1);
  top-up row 130 (`MAX(2 × cum. principal − balance, 0)`, G130 = 3,104.28). Accrual flag row 10: `N(OR(date ≤ 30-Jun-29, date > 01-Jan-30))`.
  **None of these memo amounts enters the waterfall.** The claims come from Deposits Y and Devex_FEP S (§4–5). Toggle: `interest_method='monthly'`.

**Block 9: exit waterfall** (G143–G154; Bora rows 49–61; every claim deducted in full, no priority).

| Bora row | MOIC | Content | S1 | S2 BORA 1 | S3 BORA 2 |
|---|---|---|---|---|---|
| 49 | G143 | Effective premium = `E141 × E150 + Tender!E34 × E20 + Tender!E219` (§7) | 15,000.00 | 3,927.82 | 10,082.19 |
| 50 | G144 | + buyer reimbursement: cash securities and connection fees (G99) + land (G20) | 21,837.64 | 21,837.64 | 21,837.64 |
| 51 | | = total proceeds | 36,837.64 | 25,765.46 | 31,919.83 |
| 52 | G146 | − sale costs (E26) | −500 | −500 | −500 |
| 53 | −(G110+G117) | − **recoverable loans, principal**: TSO 1,232.88 + direct 1,392.73 + indirect 18,786.66 | −21,412.26 | same | same |
| 54 | −Y23 | − **recoverable 12 % interest** on those three lines to 30-Jun-29 | −1,563.63 | same | same |
| 55 | G148 | − **unrecoverable loans, principal (land)** = G20: land reservation 55.65 + land 1st 123.24 + land RtB 246.49. Only land, because E71 = 0 (FEP carries its devex principal) | −425.38 | same | same |
| 56 | −(G131+Y24) | − **at-risk return**: devex (incl. contingency) + land at MAX(20 %, 2×) − principal = 3,087.85, **plus** the fee's 12 % = 5.86 | −3,093.71 | same | same |
| 57 | −Y22 | − **6 % on guarantees** (bid bond + PB net) to 30-Jun-29 | −1,253.92 | −1,117.09 | −971.96 |
| 58 | −G25 | − DSA | −1,033.68 | −1,200.66 | −1,113.19 |
| 59 | G152 | = distributable (**no floor**; a negative amount is shared) | 7,555.06 | **−3,547.27** | 2,839.70 |
| 61 | G154 | to the developer = `MAX(−DSA, 30 % × G152)` (floored at −DSA) | 2,266.52 | −1,064.18 | 851.91 |
| 60 | G153 | to FEP = G152 − G154 | 5,288.54 | −2,483.09 | 1,987.79 |

Rows 50, 53 and 55 net to zero: the buyer's reimbursement passes straight through to FEP. In substance the distributable equals the effective premium − sale costs − FEP's returns (12 % / 2× / 6 %) − DSA.

**Block 10: FEP cash flows** (months 0…60).
- Outflows: signing fee 250 and DD 150 at month 0; recoverable funding; unrecoverable funding (fee, devex, land); monthly guarantee fees.
- At the sale month, the inflows are rows 166–170 = (−G147 − G148 − G149 − G150) × FEP funding share (1) + G153. That is, FEP receives all of rows 53–57 plus its 70 % share.
- **Tax** (row 174, sale month only): `−9 % × MAX(cum. pre-tax CF − E14 (=1) × G153, 0)`. The 9 % rate is Inputs E12, and the participation exemption on the JV share is Inputs E14 = 1.
  Interest, the 2× return and the guarantee interest stay taxable. Signing fee, DD, devex principal and guarantee fees are deductible.
  **A negative G153 is added back** (see O7).
- Developer (block 12): inflows DSA + G154, tax `−9 % × MAX(inflows − G154, 0)`, i.e. on the DSA only.

**Block 11: returns.**
- Profit = Σ post-tax CF.
- MOIC = (inflows + tax) / outflows.
- IRR = `XIRR(H175:BP175, month-end dates)`.
- NPV = `Σ CF / 1.08^((date − 31-Jul-2026)/365)`. The rate is Inputs E253, and flows are discounted to month 0, not to today or to the sale.
- NPV/Equity = NPV / PV(negative flows).
- Peak exposure = −min cumulative post-tax.
- Economic profit = Σ (CF + MIN(prior cumulative, 0) × (1.08^(1/12) − 1)).

| Result (post-tax) | S0 | S1 | **S2 BORA 1** | **S3 BORA 2** | S4 |
|---|---|---|---|---|---|
| Pre-tax profit k€ | 17,691.3 | 6,651.03 | −1,095.36 | 3,400.30 | −3,295.51 |
| Tax at sale | −135.23 | −122.62 | −124.90 | −127.13 | −105.79 |
| **Net profit k€** | 17,556.02 | 6,528.41 | **−1,220.25** | **3,273.17** | −3,401.30 |
| **IRR** | 49.687 % | 19.373 % | **−3.913 %** (Excel: 2.98e-09) | **10.155 %** | −10.420 % (Excel: 2.98e-09) |
| MOIC | 1.690 | 1.247 | 0.953 | 1.126 | 0.877 |
| NPV @8 % | 11,582.68 | 3,016.29 | −2,923.11 | 546.51 | −4,736.20 |
| Cash invested | 25,436.56 | 26,386.41 | 26,224.34 | 26,054.43 | 27,647.53 |
| Developer post-tax | 7,661.57 | 3,207.17 | 28.42 | 1,864.92 | −108.06 |

## 7. Promise price effect (Tender §4, rows 176–227) and premium

**Gross premium** `Tender!E220 = Inputs!E141 (150 k€/MW ← Bora N5) × E150 + E34 (BESS value, Bora N21 = 0) × E20`. Here
`E150 = IF(E149 = 1, E29 installed, E18 net)`, and `E149 ← Calc!E86 ← Bora!N20 = 0`, so the **premium is on 100 net MVA = 15.0 €m for every strategy**.
Effective price = `MOIC!G143 = E220 + E219`.

**Reference = S1 "current bid"** (Tender column E). The price effect equals `E219 = pass-through E30 (100 %) × E218`, where E218 = Σ PV of six item rows: the
buyer's **extra after-tax cash cost of the live strategy's promises relative to S1**, as of the RtB sale date. The price effect is ≤ 0 when the promises exceed S1's, and S0 gets a positive one.

- **Valuation:** discount rate `E176 = Inputs!E224` = 8 % (the buyer's unlevered rate, linked to the single 8 % rate). Each year's flow sits at **31-Dec of the HTM year** (2026…2061).
  The discount factor is `1/1.08^(MAX(0, 31-Dec(y) − sale)/365)`, so flows on or before the sale date (31-Dec-2029) are not discounted.
- **Tax:** `E180 = HTM!E91 + E99` = 9 % CIT + Robin Hood 0. The depreciation profile is HTM row 90 / 20 years: the COD-year fraction 0.25 in 2031, then 1 per year, with 0.75 in 2051.

| Item (row) | Δ driver | Annual flow to the buyer (−) | Timing | S2 BORA 1 PV k€ | S3 BORA 2 PV k€ | S0 PV k€ |
|---|---|---|---|---|---|---|
| ÖSZ-1 municipal (209) | `E190 = (live − S1) k€/MW × E18 (net 100 MW)` | `−Δ × [y = COD year] + Δ × dep share × 9 %` | paid in the COD year (31-Dec-2031, 2 years of discounting); capitalised and depreciated 20 years | −2,456.16 | −2,456.16 | +1,637.44 |
| ÖSZ-2 community (210) | `G186 = live − S1 €/MWh` | `−Δ × net generation (GWh, HTM row 38) × (1 − 9 %)` | every operating year Q4-2031…Q3-2061, not escalated | −3,120.10 | 0 | +4,203.57 |
| ÖSZ-4 BESS (211) | `E191 = (live − S1 BESS MW) × 200 k€/MW (Inputs E168)` | `−Δ × construction share (HTM row 7) + Δ × dep share × 9 %` | capex NTP 30-Sep-29 → COD 30-Sep-31 by days: 2029 92/730 (**undiscounted**), 2030 365/730, 2031 273/730 | −3,658.18 | −1,741.99 | +5,225.97 |
| ÖSZ-6 security (212) | `E194 = (Deposits!H10 live − PB at S1 uplift, S1 BESS MWh and live installed MW) × 1000/FX` | `−Δ × 6 % (E182) × years outstanding after the sale` | from the sale to COD + 6 months (E179 = EDATE(COD, E31 = 6) = 30-Mar-2032): 2030 1, 2031 1, 2032 90/365. Not tax-deducted | +210.26 (ΔPB −1,770.75) | +263.66 (−2,220.55) | +457.87 |
| ÖSZ-9a EU turbine (213) | `Bora!N19 (0) × S1 installed × (live EU − S1 EU) × (950 − 500) × 1.05` | as BESS | switch off | 0 | 0 | 0 |
| ÖSZ-3 overbuild bottleneck (214) | `E205 = (clip_live − clip_S1)/(1 − clip_live)` | `−net gen × E205 × net margin (row 206) × (1 − 9 %)` | operating years | −2,048.00 | −983.32 | +2,998.46 |
| **Total = price effect** | | | | **−11,072.18** | **−4,917.81** | **+14,523.31** |

- **Net margin** (row 206, €/MWh):
  `(capture price × (1 − PPA share) + PPA × PPA share) × (1 − 0.1 % MEKH × 1.05 − 2 % HIPA) − (4 + 1) €/MWh × opex escalation × 1.05 − community €/MWh (live)`.
  - The capture price comes from Aurora Q2-26 **Central wind capture**, real 2025, inflated at 2 % (`Curves!K`, held flat in real terms after 2060).
  - The PPA share is 0, because case 1 is 100 % merchant (Inputs E190 = 1).
  - Opex escalation is 2.5 % from the COD year.
- **Generation** (HTM row 38): `installed MW × 2,460 h × (1 − 2 %) × (1 − clip_live) × (1 − 0.3 %)^(op. year − 1) × operating fraction`. COD year 0.25, last year 0.75.
  It uses the **live** strategy's MW and clip.
- **Not in the price effect:** the overbuild's extra generation and the extra wind capex and opex (only the bottleneck loss is charged), BESS revenue or opex, PB validity beyond COD + 6 months, and the TSO security.
- The ΔPB reference uses the **live installed MW**, so the extra PB caused by overbuild is not priced; only the uplift and BESS deltas are.
- **COD-route items** (E224/E225, community only) are outside the RtB scope.
- Bora Y23:AD28 = Calc DC:DH (Tender F209:F214) / 1000. Bora Y33:AD39 allocate Bora row 47 (ΔIRR vs S1) pro rata to the items. Because row 47 uses Excel's failed XIRR for S2 and S4, those columns are wrong (B1).

---

## 8. Bugs, inconsistencies and points to decide

**Bugs and errors**
- **B1 – Excel XIRR failure.** The saved IRR for S2 BORA 1 (`Bora!F45`, Calc I311) and S4 (H45) is 2.98e-09. The true root is **−3.913 %** for S2 and −10.420 % for S4 (LibreOffice and the shadow agree).
  The errors carry into Bora row 47 (ΔIRR, which is −19.37 pp for S2 instead of −23.29 pp) and into the item attribution in Y33:AD39. Excel's converged IRRs are also only accurate to about 5e-9.
- **B2 – ÖSZ-4 scored on net MVA.** Bora row 10 (BESS basis) is 0 for every strategy, so both the BESS MW and the scored ratio use **net MVA**. GYIK Q20 requires BESS MW ÷ installed wind MW.
  With the GYIK rule the totals become: S1 36 → **33.0**, BORA 1 85 → **80.5**, BORA 2 67.5 → **63.0**, S4 90 → 85.5. Toggle `osz4_basis_installed=True`.
- **B3 – Devex not on the project's configuration.**
  - The devex lines come from the pasted 25-Sep file: 200 MW wind, 100 MW / 200 MWh BESS; the MW-scaled design fees use AF7 = 200.
  - The project is 100 MVA / 130–151 MW. The figures are identical for every strategy and do not scale with MW or MWh.
  - They also differ from the 02-Oct devex method file, which uses 260 MW, 130 MW / 520 MWh and a different contingency base. Toggle `devex_source='devex_file_20261002'`: −259 k€ on BORA 1's profit.
- **B4 – Same-year stub quirk replicated.** The connection fees, paid 08-Jan-2029, are accrued to 30-Jun-2029 with `(1+r·358/365)(1.12)^-1(1+r·181/365)` = 1.057330, against a true stub of 1.057205.
  Deposits Z and Devex_FEP N copy the devex formula without a same-year branch. The effect is small (BORA 1: −0.5 k€ when fixed). Toggle `same_year_stub_fix`.
- **B5 – Inconsistent RtB timeline.**
  - NTP (30-Sep-2029 = RtB + 3) comes **before** the RtB sale (31-Dec-2029 = RtB + 6), yet the RtB route carries no construction cost.
  - In the price effect, 12.6 % of the ΔBESS capex falls in 2029 at a discount factor of 1. That charges the buyer for construction before it owns the project.
  - The devex RtB land payment (24-Mar-2029) and the devex RtB milestone come before the engine's RtB date of 30-Jun-2029.

**Inconsistent with the tender call**
- **T1 – PB posting date.** The engine posts the PB on 30-Dec-2026 (award + 17 days). Under the call (XII.5, GYIK 68) the deadline is ≤ 15 days after receipt of the decision, which is about mid-February 2027.
  Toggle `pb_date`: 15-Feb-2027 gives +41 k€ on BORA 1.
- **T2 – PB validity.** The price-effect security carry runs only to COD + 6 months. The PB guarantee must be valid to the **COD deadline + 6 months (≥ 31-Mar-2033)** (XII.8).
  Toggle `pb_hold_months_after_cod`.
- **T3 – TSO security.** 4.5 mHUF/MW, cash only, on net MVA. The call gives no amount (it defers to the grid code); the credit against the connection fee is consistent with GYIK 40.
- **T4 – Direct fee "per MW".** E84 = lump ÷ Calc E74 × E18, which equals the lump sum only while `Inputs!E16 × E17 = Calc!E74`. E16 is a separate typed 100 (see O3).
  The call charges per connection point (scaled only under ÖSZ-8 sharing, XIII.18). Toggle `direct_fee_basis`.
- **T5 – Bid-bond netting.**
  - Netting is **presentational in this engine**. The bid guarantee stays outstanding until the sale, alongside the PB net, so the exposure is the PB gross.
  - This is economically the same as "bid released at PB posting, PB issued gross", which is what XII.11 b implies for guarantees: the shadow gives Δ 0.
  - `bid_netting=False` without a release double-counts (−189 k€ on BORA 1).
  - This corrects `tender_2026_rules.md` §8.3: the engine does not understate the guarantee and its fee.
- **T6 – S0 and the expected law.** S0's +14.5 €m price effect assumes the buyer escapes the S1 promises entirely. Footnote 1 of the call expects 10 k€/MW and 2 €/MWh to become mandatory, and that would cut most of this.
  S0 also scores 0 points, so it would almost certainly not be awarded.
- **T7 – No forfeiture or failure scenarios.** The engine does not model the overbuild shortfall forfeiture, a full PB call (XIII.17 / GYIK 72), or a refund of the TSO security on failure.

**Other oddities**
- **O1 – Two FX rates.** Securities and fees convert at 365 (Inputs E11). Devex amounts are the pasted file's EUR values at **370**. Toggles: `fx`, `fx_devex`. FX 370 on securities: +6.7 k€; devex at 365: −18.2 k€ (BORA 1).
- **O2 – Guarantee cost vs credit mismatch.**
  - FEP pays a **6 % fee monthly**: full months, including the issue month, and running to the sale month without a pause.
  - It is credited only **6 % accrued per line to 30-Jun-2029**.
  - Net guarantee carry: S1 −216 k€, BORA 1 −191 k€. Toggle `guarantee_fee`.
- **O3 – Two MW inputs.** `Inputs!E16` = 100 (typed: tender MW) and `Calc!E74` = 100 (net MVA) must be kept equal by hand.
  - E16 drives E18, which in turn drives the TSO security, the fees, the net premium and the municipal price effect.
  - Calc E74 drives the strategy MW and the per-MW direct fee.
- **O4 – Premium on net MVA.** The overbuild earns no premium on the RtB route (`premium_basis='installed'` adds +5,355 k€ to BORA 1's profit). The devex file prices on installed MW (V34).
- **O5 – Devex principal is not repaid.** FEP recovers it only through the 2× return (E71 = 0), and the participation fee only through 12 %, with no 2×. The 2× floor binds on every line, so the 20 % rate never matters.
- **O6 – Developer floor and negative distributable.**
  - The developer floor at −DSA (G154) is not in the devex file.
  - With a negative distributable (BORA 1 −3.55 €m, S4 −5.67 €m), FEP's 70 % share is negative and eats into its interest income.
  - For S4 the floor binds: the developer bears −DSA and FEP bears the rest.
- **O7 – Tax on a pre-tax loss.** The participation exemption deducts G153 even when it is negative. BORA 1 therefore pays 124.9 k€ of tax on a −1,095 k€ pre-tax loss, because the JV-share loss is non-deductible while the interest income is taxed.
  The developer pays 9 % on its DSA even when its share is negative.
- **O8 – Memo schedules and labels.**
  - The MOIC monthly loan schedules (rows 109–140, incl. the "top-up" G130) and the guarantee interest row 82 are not used by the waterfall, yet checks G255–G257 test them. Monthly method: +16 k€ on BORA 1.
  - Inputs notes F61/F62 ("compounded monthly") describe the memo, not the claim.
  - Stale notes: F82 ("6 SPVs"), F72 ("8.9 %"), F85 ("200 MVA … 37,573"), F20 ("100 MW").
- **O9 – Cash-flow timing.** Interest claims use exact dates, but the cash flows behind IRR/NPV are moved to month ends.
- **O10 – Overbuild in the price effect.** Only the bottleneck loss on the live installed MW counts (`gen_live × (clip_live − clip_ref)/(1 − clip_live)`). The extra MWh from overbuild is never credited.
- **O11 – Deposits reconciliation check suppressed.** `Deposits!E48` = "FAIL", but E34 shows "n/a" because 6 drivers differ from the source, so the check is reported as n/a rather than failed.

## 9. Differences from `docs/domain/bora_devex_and_overbuild.md` §5–6 (bridge toggles in `old_rtb.py`)

| # | Old engine (default) | Devex file / proposed waterfall | Toggle |
|---|---|---|---|
| 1 | Devex lines from the 25-Sep paste (200 MW basis), contingency 10 % × (FEP devex + land + fee) on 31-Dec-28 | 02-Oct file, 260 MW basis, 10 % × all non-bond lines | `devex_source` |
| 2 | FX 365 for securities, 370 embedded in devex | 370 everywhere | `fx`, `fx_devex` |
| 3 | Premium × 100 net MVA | × installed MW (260 in the file) | `premium_basis` |
| 4 | Indirect fee × net MVA | × installed MW (file I18, probably an error) | `indirect_fee_basis` |
| 5 | Direct fee = lump/MVA × E18 | lump per connection point | `direct_fee_basis` |
| 6 | Bid bond netted into the PB, bid guarantee held to the sale | PB gross, bid released | `bid_netting`, `bid_release_date` |
| 7 | PB on 30-Dec-26; held after COD for 6 months in the price effect | ~15-Feb-27; to the COD deadline + 6 months | `pb_date`, `pb_hold_months_after_cod` |
| 8 | Per-line Act/365, 31-Dec compounding, pause = freeze at 30-Jun-29, same-year quirk | same convention (the file's "freeze") | `interest_method`, `pause`, `same_year_stub_fix` |
| 9 | Developer floored at −DSA | no floor (file); MIN-waterfall (§6.6 proposal) | `dev_floor` |
| 10 | FEP pays a 6 % guarantee fee monthly; signing 250 + DD 150; 9 % CIT with participation exemption | not in the file | `guarantee_fee`, `signing_dd`, `tax` |
| 11 | DSA 2.5 k€/MW × installed MW × YEARFRAC 30/360 (26-Oct-26 → sale), cap 10 k€/MW | same, no cap | – (cap not binding) |
| 12 | Buyer reimburses cash securities, fees and land in full; the guarantees are replaced | memo line incl. the guarantees | – |

## 10. Open questions
1. Which devex file and configuration should the new model use: the 25-Sep paste at 200 MW, or the 02-Oct file at 260 MW? The project is now 100 MVA / 130–151 MW, so the MW-scaled lines need rescaling either way.
2. Should the timeline be fixed so that NTP falls after the RtB sale, or the sale moved to RtB + 3 months? And should the RtB date follow the devex milestone (24-Mar-2029)?
3. Should the ÖSZ-4 basis follow GYIK 20 (installed MW)? That changes the BESS MW needed for the same points.
4. Should the tax treatment of a negative JV share (O7) and the developer floor (O6) be kept?
