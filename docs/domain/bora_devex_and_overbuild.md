# Bora: development expenditure (devex) funding base and overbuild penalty

Analysis of two developer workbooks to be used as the base for the new Bora bidding model
(house standard: `docs/MODELLING_STANDARDS.md`). Both files were read as untrusted data with openpyxl,
reading both formulas and cached values. No macros were run. Both are `.xlsx` files with no VBA and no external links.
The tender rules quoted below come from `reference/Kapacitas_palyazati_kiiras_2026_3.pdf` (the "kiírás", or tender call)
and `reference/Kapacitaspalyazat2026_GYIK.pdf` (the "GYIK", the tender FAQ).

| File | Sheets | Provenance (workbook metadata) | Language |
|---|---|---|---|
| A. `BORA_Devex_20261002_BG.xlsx` | 1 (`Sheet1`, used range B2:Y58, print area B2:S58, panes frozen at G5) | Created 21-Mar-2025. Last saved 06-Oct-2026 by B. Gugyella ("BG" in the file name). Saved under an `…\A.RES\04_PROJEKTEK\…\05_Bora\5_Finance\` path | Already English. No Hungarian strings in the file. HU terms below come from the IC model / tender, or are translations (tr.) |
| B. `BORA_tulepites.xlsx` ("túlépítés" = overbuild) | 1 (`Tulepites`, A1:R6) | Created 28-Sep-2026, last saved 12-Oct-2026 (B. Gugyella) | Hungarian headers, translated in §3.1 |

**Parties.** Neither file contains the name "ALSI". File A labels the two parties **"FGX"** (70 %) and **"Ventus+Solmentore"**
(30 % + DSA). Cell V21, which drives the 70 %, is labelled "FEP split", so FGX is the FEP side. The 30 % party is the developer. The IC model calls it A.RES, and this brief calls it **ALSI**. This report uses FEP and ALSI. **To confirm:** that Ventus+Solmentore = ALSI, and whether FGX is FEP's vehicle.

Units: file A works in **HUF** (column I) and **EUR**, not k€ (column J onwards). File B works in HUF, EUR, MW and **percentage points**
(C and G are ×100, not decimals). FX = **370 HUF/EUR** in both files.

---

## 1. Key findings

1. **FEP funds 100 % of every line.** Column N "FEP principal" = J on every row. ALSI funds nothing. It earns a DSA (2,500 €/MW/yr) and 30 % of the
   distributable proceeds.
2. **Total committed: 77.43 m€ (28,651 mHUF).** Of this, 73.96 m€ is bonds and connection fees. Within that, MEKH bank guarantees are 21.24 m€, the
   TSO cash deposit is 2.43 m€ and connection fees are 50.29 m€. The other 3.47 m€ is fees, land and devex (incl. a 10 % contingency).
3. **The return on funding is set per line, not per party:** 12 % on cash items (fee, TSO deposit, connection fees),
   6 % on the face value of bank guarantees, and 20 % on devex/land **with a 2.0× floor**. Interest compounds annually at 31-Dec, stubs are Act/365 with both ends counted,
   and accrual ends at the **interest freeze 30-Jun-2029** or, without the freeze, at the **exit 31-Dec-2029**. At the current dates the 2× floor binds on **every** 20 % line,
   so the 20 % rate never matters and devex/land return exactly 100 % of principal.
4. **Waterfall, as the file states it:** premium (150 k€/MW × **260 installed MW** = 39.0 m€) − sale costs 0.5 − **all FEP interest** 10.52 (14.47 without the freeze)
   − DSA 2.07 = distributable 25.91 (21.96) m€, split 70/30. FEP's net = interest + 70 % − devex/fee principal (3.05 m€)
   = **25.61 m€ (26.80 without the freeze)**. ALSI's = 30 % + DSA = **9.84 m€ (8.66)**. The buyer also reimburses land, bonds and connection fees (74.39 m€, memo line, not in the split).
   There is no hurdle, catch-up or promote, and no floor on the distributable amount.
5. **Probable error: the indirect connection fee is on 260 MW installed, not 200 MVA:** `I18 = 73,071,306 × V6 − I8`. The IC model uses 200 MVA.
   This overstates the fee by **4,384 mHUF / 11.85 m€** and inflates the 12 % interest on it.
6. **The MW basis is inconsistent inside file A.** The TSO bond uses a hardcoded 200 MW (`=4500000*200`). The connection fee, bid/performance bonds, design fees,
   premium and DSA use V6 = 260 MW. The 200 MVA grid capacity is not an input anywhere.
7. **Formula defects in the interest:** (a) items paid in the freeze year (2029) get `(1+r·a)(1+r·b)/(1+r)` instead of a simple stub, overstating by +6.1 k€ on the indirect
   connection fee; (b) there is no guard for payments after the end date; (c) the no-freeze formula is only correct because the exit falls on 31-Dec; (d) the 2× floor is switched
   by the rate `K ≥ V19`, not by category. All four are small at today's inputs, but fragile.
8. **FX: 370 (file A, file B and the pasted BORA_Devex in the IC model) vs 365 (IC model Inputs E11).** At 365 the file A total rises by **+1.05 m€** (HUF lines only).
9. **The overbuild file (B) models a 10 % cliff.** A shortfall up to 10 % of the committed installed MW costs nothing. Above that, the **whole shortfall** × 10.15 mHUF/MW is forfeited
   (2.5 % × 406 mHUF/MW). In scenario 1 (commit 300 MW = +50 %, build 217 MW) the forfeit is **842 mHUF = 2.28 m€**, and in scenario 3 it is 304.5 mHUF = 0.82 m€.
   The file ignores the **150 % ÖSZ-6 uplift** used in file A, which would raise the forfeit by one third (3.42 m€ in scenario 1). It also ignores the BESS test and has no timing.
10. **Legal risk the overbuild file does not model.** Overbuild is an ÖSZ-3 commitment. Under kiírás XIII.17 and GYIK 72, an unmet ÖSZ commitment means MEKH calls the **entire** performance bond.
    GYIK 71 adds a **3-year exclusion of the whole company group**, and GYIK 28/30 add the **pro-rata loss of the connection right**. GYIK 30/44 support the pro-rata 10 % regime that file B uses.
    The tender documents conflict here, and this must be cleared before the bid.

---

## 2. File A – `BORA_Devex_20261002_BG.xlsx`

### 2.1 Layout

| Block | Range | Content |
|---|---|---|
| Line items | B3:S48 | B task group (merged), C start, D end, **E payment date**, F year, G days in the first year, H item, **I HUF**, **J EUR**, **K rate**, **L category**, **M settlement**, **N FEP principal**, O comment, P/Q amount/interest with freeze, R/S amount/interest without freeze |
| Totals | rows 49–53 | 49 total incl. bonds and connection fee, 50 per wind MW, 51 total excl. bonds/connection fee, 52 per MW, 53 bonds + connection fee |
| Inputs | U4:V27 | Blue-font inputs (V6 highlighted yellow) |
| Waterfall | U30:Y44 | "FEP proposal – Exit: 2029.12.31": V/W freeze (EUR, EUR/MW), X/Y no freeze |
| Category summary | U47:X55 | Principal and interest by category (SUMIF/SUMIFS on L and M) |
| Disclaimers | B58 | 11 notes (summarised in §2.5) |

Rows 6–13 and 18–19 (bonds and connection fees) are shaded. These are also the rows excluded from the contingency base.
Yellow cells mark this version's edits: V6, I8, H9:H13, K18:K19.

### 2.2 Inputs (U5:V27)

| Cell | Label | Value | Unit / note |
|---|---|---|---|
| V5 | FY rate (HUF/EUR) | 370 | HUF per EUR |
| V6 | Wind capacity (MW) | 260 | installed MW (= 200 MVA × 1.30; IC S1 = 261) |
| V7 | BESS percentage | 50 % | of wind MW |
| V8 | BESS connection capacity (MW) | `=V6*V7` → 130 | |
| V9 | BESS storage capacity (MWh) – 4h | `=V8*4` → 520 | 4 h hardcoded |
| V10 | Benchmark CAPEX – wind | 406,000,000 | HUF/MW (kiírás Annex 2) |
| V11 | Benchmark CAPEX – BESS | 200,000,000 | HUF/MWh (Annex 2) |
| V12 | Tender application | 2 | number of bids |
| V13 | No. of top. lots | 28 | topographic lots |
| V14 | No. of WTGs | 28 | |
| V15 | No. of small power plants | 4 | permitting units |
| V16 | No. of municipalities | 10 | |
| V17 | Bank guarantee interest | 6 % | p.a. |
| V18 | Interest – cash deposits, fees, connection fee | 12 % | p.a. |
| V19 | Interest – devex and land | 20 % | p.a. |
| V20 | Minimum MOIC on 20 % items | 2.0× | |
| V21 | FEP split | 70 % | |
| V22 | DSA / success fee | 2,500 | EUR/MW/yr |
| V23 | DSA accrual start (signing) | `=DATE(2026,10,26)` | |
| V24 | Interest freeze date | `=DATE(2029,6,30)` | |
| V25 | Exit date | `=DATE(2029,12,31)` | |
| V26 | Purchase price | 150,000 | EUR/MW |
| V27 | Transaction costs | 500,000 | EUR |

### 2.3 Every line item

Rate codes: **C12** = V18 12 %, **G6** = V17 6 %, **A20** = V19 20 % with the 2× floor (V20). P = amount incl. interest to the freeze date (30-Jun-2029). R = amount incl. interest to the exit date (31-Dec-2029).
Every line is funded 100 % by FEP (N = J). HU names marked (tr.) are translations, because the source is in English.

| Row | Group / item (EN) | HU | Category (L) / settlement (M) | Amount formula (I or J) | HUF | EUR | Pay date (E) | Rate | P (freeze) | R (no freeze) | Comment (O) |
|---|---|---|---|---|---:|---:|---|---|---:|---:|---|
| 5 | MEKH capacity tender – Application fee | Részvételi díj | Fee / Cash | `=V12*3000000` | 6,000,000 | 16,216 | 30-Oct-26 | C12 | 21,998 | 23,254 | |
| 6 | MEKH – Application (bid) bond | Ajánlati biztosíték | Bond / Bank guarantee | `=V10*V6*0.005+V11*V9*0.005` | 1,047,800,000 | 2,831,892 | 30-Oct-26 | G6 | 3,310,519 | 3,407,758 | W+BESS |
| 7 | MEKH – Performance bond | Teljesítési biztosíték | Bond / Bank guarantee | `=((V10*V6*0.025+V11*V9*0.025)*1.5)-I6` | 6,810,700,000 | 18,407,297 | 30-Dec-26 | G6 | 21,304,815 | 21,930,593 | W+BESS; application bond deducted; incl. 50 % extra |
| 8 | Capacity allocation – Application bond (TSO) | Igénybejelentési biztosíték | Bond / Cash deposit | `=4500000*200` | 900,000,000 | 2,432,432 | 30-Dec-26 | C12 | 3,234,939 | 3,419,639 | Refunded if project fails; deducted from conn. fee |
| 9 | Supplementary grid bond – 2027 | Kiegészítő biztosíték | Bond / BG | blank | 0 | 0 | 23-Apr-27 | blank | 0 | 0 | Double amount; refunded if project fails |
| 10 | Supplementary grid bond – 2028 | Kiegészítő biztosíték | Bond / BG | blank | 0 | 0 | 23-Apr-28 | blank | 0 | 0 | Refunded if project fails |
| 11 | Supplementary grid bond – refund | … visszautalás | Refund / BG | blank (date `=D35+45`) | 0 | 0 | 13-Mar-28 | blank | 0 | 0 | Refund after BP (with time buffer) |
| 12 | Supplementary grid bond – 2029 | Kiegészítő biztosíték | Bond / BG | blank | 0 | 0 | 23-Apr-29 | blank | 0 | 0 | 50 % after building permit; refunded if project fails |
| 13 | Supplementary grid bond – 2030 | Kiegészítő biztosíték | Bond / BG | blank | 0 | 0 | 23-Apr-30 | blank | 0 | 0 | Paid by new owner; refunded if project fails |
| 14 | Awarding – Land reservation fee | Földterület-lekötési díj (tr.) | Land / Cash | `J=1325*V14*1.5` (EUR) | 20,590,500 | 55,650 | 13-Feb-27 | A20 | 111,300 | 111,300 | Reimbursed by buyer at exit; may be increased |
| 15 | Awarding – Land purchase price (1st part) | Földterület vételár (tr.) | Land / Cash | `=38000000*0.2*6` | 45,600,000 | 123,243 | 13-Feb-27 | A20 | 246,486 | 246,486 | Reimbursed by buyer at exit |
| 16 | Grid connection plan – Design fee | Tervezési díj (tr.) | Devex / Cash | `=72000000/280*V6*1.15` | 76,885,714 | 207,799 | 11-Aug-28 | A20 | 415,598 | 415,598 | |
| 17 | Grid connection plan – Administrative fee | Igazgatási díj (tr.) | Devex / Cash | `=187636*V15` | 750,544 | 2,028 | 11-Aug-28 | A20 | 4,057 | 4,057 | |
| 18 | Grid connection agreement – Connection fee, indirect | Csatlakozási díj – becsült (közvetett) | Connection fee / Cash | `=(73071306*V6)-I8` | 18,098,539,560 | 48,914,972 | 08-Jan-29 | C12 | 51,719,256 | 54,672,197 | TSO application bond deducted |
| 19 | Grid connection agreement – Connection fee, direct | Csatlakozási díj – közvetlen | Connection fee / Cash | `=508345804` | 508,345,804 | 1,373,908 | 08-Jan-29 | C12 | 1,452,673 | 1,535,615 | Depending on SPV's declaration and TSO's decision |
| 20 | Land contracting (additional land) – Legal fee, agent fee | Ügyvédi és ügynöki díj (tr.) | Devex / Cash | `=25*(1200000+180000)+3000*V13*V5` | 65,580,000 | 177,243 | 01-Oct-26 | A20 | 354,486 | 354,486 | To be paid |
| 21 | Wind measurement – Met masts | Szélmérő árboc (tr.) | Devex / Cash | `J=199000*2` (EUR) | 147,260,000 | 398,000 | 01-Mar-27 | A20 | 796,000 | 796,000 | For 2 MM |
| 22 | Wind measurement – Design fee | Tervezési díj (tr.) | Devex / Cash | `=1500000*2*1.2` | 3,600,000 | 9,730 | 01-Jan-27 | A20 | 19,459 | 19,459 | For 2 MM |
| 23 | Wind measurement – EYA | Hozambecslés (tr.) | Devex / Cash | `J=4500` typed (EUR) | 1,665,000 | 4,500 | 28-Jun-28 | A20 | 9,000 | 9,000 | |
| 24 | Land reclassification – Administrative fee | Igazgatási díj (tr.) | Devex / Cash | `=30000*(V14+V15)` | 960,000 | 2,595 | 25-Dec-27 | A20 | 5,189 | 5,189 | |
| 25 | Land reclassification – Report fee | Szakvélemény díja (tr.) | Devex / Cash | `=250000*(V14+V15)` | 8,000,000 | 21,622 | 28-Jul-27 | A20 | 43,243 | 43,243 | |
| 26 | Zoning plan modification – Design fee | Településrendezési terv (tr.) | Devex / Cash | `=5000000*V16` | 50,000,000 | 135,135 | 23-Feb-28 | A20 | 270,270 | 270,270 | |
| 27 | Municipality donation – Donation | Önkormányzati támogatás (tr.) | Devex / Cash | `=3500000*V16` | 35,000,000 | 94,595 | 23-Feb-28 | A20 | 189,189 | 189,189 | |
| 28 | Environmental assessment – Bird/bat monitoring | Madár-denevér monitoring (tr.) | Devex / Cash | `=12360000*1.2` | 14,832,000 | 40,086 | 30-Dec-27 | A20 | 80,173 | 80,173 | To be paid |
| 29 | Environmental assessment – Administrative fee | Igazgatási díj (tr.) | Devex / Cash | `=10000*V15` | 40,000 | 108 | 28-Apr-28 | A20 | 216 | 216 | |
| 30 | Environmental assessment – Assessment fee | Vizsgálati díj (tr.) | Devex / Cash | `=(7200000+10180000+4360000)*1.5` | 32,610,000 | 88,135 | 29-Mar-28 | A20 | 176,270 | 176,270 | |
| 31 | Other development – Geotechnical | Geotechnika (tr.) | Devex / Cash | `=250000*(V14+V15)` | 8,000,000 | 21,622 | 28-Jul-27 | A20 | 43,243 | 43,243 | |
| 32 | Other development – Archaeological | Régészet (tr.) | Devex / Cash | `=250000*(V14+V15)` | 8,000,000 | 21,622 | 28-Jun-27 | A20 | 43,243 | 43,243 | |
| 33 | Other development – Topography | Topográfia (tr.) | Devex / Cash | `=250000*(V14+V15)` | 8,000,000 | 21,622 | 28-Jun-27 | A20 | 43,243 | 43,243 | |
| 34 | Building permit, turbines – Administrative fee | Építési engedély, igazgatási díj (tr.) | Devex / Cash | `=1147500*V15` | 4,590,000 | 12,405 | 28-Jan-28 | A20 | 24,811 | 24,811 | |
| 35 | Building permit, turbines – Design fee | Tervezési díj (tr.) | Devex / Cash | `=30000000*V15*1.15` | 138,000,000 | 372,973 | 28-Jan-28 | A20 | 745,946 | 745,946 | |
| 36 | Cable right – Administrative fee | Vezetékjog, igazgatási díj (tr.) | Devex / Cash | `=150000*300` | 45,000,000 | 121,622 | 28-Jan-28 | A20 | 243,243 | 243,243 | |
| 37 | Cable right – Design fee | Tervezési díj (tr.) | Devex / Cash | `=(I40+I33+38000000)/280*V6*1.15` | 111,838,967 | 302,267 | 28-Jan-28 | A20 | 604,535 | 604,535 | |
| 38 | Cable right – Compensation | Kártalanítás (tr.) | Devex / Cash | `=50000*300` | 15,000,000 | 40,541 | 29-Nov-27 | A20 | 81,081 | 81,081 | |
| 39 | Building permit, substation/switching station – Administrative fee | Igazgatási díj (tr.) | Devex / Cash | `=1500000*V15` | 6,000,000 | 16,216 | 28-Jan-28 | A20 | 32,432 | 32,432 | |
| 40 | Building permit, substation – Design fee | Tervezési díj (tr.) | Devex / Cash | `=55000000/280*V6*1.15` | 58,732,143 | 158,736 | 28-Jan-28 | A20 | 317,471 | 317,471 | |
| 41 | Building permit, BESS – Administrative fee | Igazgatási díj (tr.) | Devex / Cash | `=1147500*V15` | 4,590,000 | 12,405 | 28-Jan-28 | A20 | 24,811 | 24,811 | |
| 42 | Building permit, BESS – Design fee | Tervezési díj (tr.) | Devex / Cash | `=7000000/40*V9` | 91,000,000 | 245,946 | 28-Jan-28 | A20 | 491,892 | 491,892 | |
| 43 | HEPURA permit – Administrative fee | Igazgatási díj (tr.) | Devex / Cash | `=2500000*V15*2` | 20,000,000 | 54,054 | 08-Jan-29 | A20 | 108,108 | 108,108 | W+BESS ("HEPURA" is not expanded in the source; likely the MEKH power-plant licence) |
| 44 | HEPURA permit – Report fee | Szakvélemény díja (tr.) | Devex / Cash | `=1500000*V15*2` | 12,000,000 | 32,432 | 09-Nov-28 | A20 | 64,865 | 64,865 | W+BESS |
| 45 | Legal costs | Jogi költségek (tr.) | Devex / Cash | `J=20000` typed | 7,400,000 | 20,000 | "Continuous" (F=2027, G=0) | A20 | 40,000 | 40,000 | |
| 46 | Financial costs | Pénzügyi költségek (tr.) | Devex / Cash | `J=80000` typed | 29,600,000 | 80,000 | "Continuous" (F=2027) | A20 | 160,000 | 160,000 | |
| 47 | RtB – Land purchase price (2nd part) | Földterület vételár (tr.) | Land / Cash | `=38000000*0.4*6` | 91,200,000 | 246,486 | 24-Mar-29 | A20 | 492,973 | 492,973 | |
| 48 | Contingency | Tartalék (tr.) | Devex / Cash | `=0.1*(SUM(I5:I47)-I6-I7-I8-I18-I19-I9-I10-I11-I12-I13)` | 116,832,487 | 315,763 | "Continuous" (F=2028) | A20 | 631,527 | 631,527 | |
| **49** | **Total with bonds and connection fee** | | | `=SUM(5:48)` | **28,650,542,719** | **77,433,899** | | | **87,958,565** | **91,903,421** | |
| 53 | of which bonds + connection fee | | | explicit cell list | 27,365,385,364 | 73,960,501 | | | 81,022,203 | 84,965,802 | |
| 51 | Total without bonds and connection fee | | | `=J49-J53` | 1,285,157,355 | 3,473,398 | | | 6,936,363 | 6,937,619 | |

**Date chain** (column C/D/E formulas). Bid on 30-Oct-26 (typed). The performance bond is paid 30-Dec-26 (typed start 15-Dec-26). The capacity application
`C8=C7+15`. Award `C14=E8+15`, `D14=C14+30` = 13-Feb-27. The grid connection plan is `C16=D14`, `D16=C16+365+180` = 11-Aug-28. The grid connection agreement
is `C18=E17+90+60` = 08-Jan-29. HEPURA runs `C43=D19`, `D43=+75`, and **RtB = 24-Mar-29** (`C47=D43`). The building permit is `C35` 01-Aug-27 (typed) plus 180 = 28-Jan-28,
and the supplementary refund `C11=D35+45` = 13-Mar-28.

**Category summary (U47:X55):**

| Category (L) | Principal € | Interest, freeze € | Interest, no freeze € |
|---|---:|---:|---:|
| Fee | 16,216 | 5,782 | 7,038 |
| Bond – bank guarantee | 21,239,189 | 3,376,145 | 4,099,162 |
| Bond – cash deposit | 2,432,432 | 802,506 | 987,207 |
| Refund – bank guarantee | 0 | 0 | 0 |
| Connection fee | 50,288,879 | 2,883,050 | 5,918,932 |
| Land | 425,380 | 425,380 | 425,380 |
| Devex (incl. contingency) | 3,031,802 | 3,031,802 | 3,031,802 |
| **Total** | **77,433,899** | **10,524,666** | **14,469,522** |

FEP's cash outlay is 56.19 m€ (the total less the 21.24 m€ of guarantees, which use no cash).

### 2.4 Deposits and securities payable to MEKH and to the TSO/DSO

| Item | Payee | Basis | Amount | Timing | Form / rate | Refundability (file comment and tender rule) | How the file computes it |
|---|---|---|---|---|---|---|---|
| Participation fee (Részvételi díj) | MEKH | 3 mHUF per bid × 2 bids (kiírás VI.4) | 6.0 mHUF / 16.2 k€ | 30-Oct-26 (bid deadline) | Cash, 12 % | **Not refundable** (only if MEKH withdraws the tender, VI.7) | `I5 = V12*3000000` |
| Bid bond (Ajánlati biztosíték) | MEKH | 0.5 % × (installed wind MW × 406 mHUF + BESS MWh × 200 mHUF) (VIII.2) = 0.5 % × (260 × 406 + 520 × 200) mHUF | 1,047.8 mHUF / 2.83 m€ | 30-Oct-26 | Bank guarantee, 6 % | Released after award once the performance bond is in place (XII.11 b). Called if the winner fails to post the performance bond (XII.10). The file nets it into the performance bond | `I6 = V10*V6*0.005 + V11*V9*0.005` |
| Performance bond (Teljesítési biztosíték) | MEKH | 2.5 % × same base (XII.2) × **1.5 ÖSZ-6 uplift**, less the bid bond | gross 7,858.5 mHUF. Payable line 6,810.7 mHUF / 18.41 m€ | 30-Dec-26 (≤ 15 days after the decision, XII.5) | Bank guarantee, 6 % | Released after COD verification (XII.13). The BG must be valid until **COD deadline + 6 months** (XII.8). **Forfeited in full** on loss of the connection right or an unmet ÖSZ commitment (XII.12, XIII.17, GYIK 72), and **pro rata** to a capacity shortfall above 10 % (XII.12, GYIK 28/30/31) | `I7 = ((V10*V6*0.025 + V11*V9*0.025)*1.5) - I6` |
| Connection request security (Igénybejelentési biztosíték) | TSO (MAVIR) | 4.5 mHUF/MW × **200 MW (hardcoded)** | 900 mHUF / 2.43 m€ | 30-Dec-26 | **Cash deposit** (no BG allowed, disclaimer 8), 12 % | Refunded if the project fails. **Credited against the indirect connection fee** | `I8 = 4500000*200` and `I18 = … - I8` |
| Supplementary grid security 2027–2030 and refund (Kiegészítő biztosíték) | TSO | n/a | **0** (rows kept as placeholders with dates) | 23-Apr each year; refund 13-Mar-28 | BG | Not payable in this procedure (GYIK 70). The older source version had 2 % / 1 % / 0.5 % of the wind benchmark, a 50 % refund at BP, and the 2030 instalment "paid by new owner" | I9:I13 blank |
| Indirect connection fee (Csatlakozási díj) | TSO | 73,071,306 HUF/MW × **V6 = 260 MW**, less the TSO security | 18,098.5 mHUF / 48.91 m€ (19.0 bn gross) | 08-Jan-29 (GCA) | Cash, 12 % | Not refundable. **The buyer reimburses it at exit** (V33). Should probably be on **200 MVA** (13,714.3 mHUF, as in the IC model) | `I18 = (73071306*V6) - I8` |
| Direct connection fee | TSO | lump sum | 508.3 mHUF / 1.37 m€ | 08-Jan-29 | Cash, 12 % | Not refundable. The buyer reimburses it. "Depending on the SPV's declaration and the TSO's decision" | `I19 = 508345804` |

The overbuild penalty (a partial call of the performance bond) is in §3.

### 2.5 Interest / return on each line

**Rate per line.** K5 (fee), K8 (TSO deposit), K18 and K19 (connection fees) = `$V$18` 12 %. K6 and K7 (MEKH bonds) = `$V$17` 6 %. Every devex and land line
(K14:K17, K20:K48) = `$V$19` 20 %. K9:K13 are blank, giving 0 %. The rate depends on the line, **not on the party**, and every line is FEP's.

**Formulas (row n):**

```
F n = YEAR(E n)                                   payment year
G n = DATE(YEAR(E n),12,31) - E n + 1             days from payment to 31-Dec, both ends counted
P n = $N n * MAX( (1+$K n*$G n/365)                               first-year stub
                * (1+$K n)^(YEAR($V$24)-$F n-1)                    whole calendar years
                * (1+$K n*($V$24-DATE(YEAR($V$24),1,1)+1)/365),    last-year stub to 30-Jun-2029
                  ($K n>=$V$19)*$V$20 )                            2.0x floor only if rate >= 20 %
Q n = P n - N n                                   interest (freeze case)
R n = $N n * MAX( (1+$K n*$G n/365) * (1+$K n)^(YEAR($V$25)-$F n), ($K n>=$V$19)*$V$20 )
S n = R n - N n                                   interest (no freeze, to 31-Dec-2029)
```

**Convention** (disclaimer 7: "exact dates, compounds yearly"):
- Simple interest within a calendar year. Interest capitalises on 31-Dec.
- Act/365 for stubs, counting both the start and the end day. Whole years compound at (1+r) regardless of leap years, so 2028 counts as one year.
- The accrual runs from the payment date to 30-Jun-2029 ("interest frozen"), or to the exit 31-Dec-2029 ("no freeze", disclaimer 11).
- The 2× floor applies to 20 % items: "minimum return of 2× the principal" (disclaimer 10).
- Guarantees earn 6 % on their face value (disclaimer 6). That is FEP's fee for posting them. The file has no line for the bank's own guarantee fee.

**Worked example by hand – row 8, TSO cash deposit, N = 2,432,432.43 €, K = 12 %, paid 30-Dec-2026:**
- G8 = 31-Dec − 30-Dec + 1 = **2 days**. The whole years to the freeze are 2029 − 2026 − 1 = **2** (2027, 2028). The last stub runs 1-Jan to 30-Jun-2029, which is **181 days**.
- The freeze factor is (1 + 0.12×2/365) × 1.12² × (1 + 0.12×181/365) = 1.000658 × 1.254400 × 1.059507 = **1.329919**.
  So P8 = 2,432,432.43 × 1.329919 = **3,234,938.80 €**, and Q8 = **802,506.37 €**. Both match the file.
- Without the freeze: (1.000658) × 1.12³ = 1.405852, so R8 = **3,419,639.48 €**, which matches.
- The floor term is (0.12 ≥ 0.20) × 2 = 0, so it does not apply.

**Example with the floor – row 35, turbine design fee, 372,973 € at 20 %, paid 28-Jan-2028:**
- The factor is (1 + 0.2×339/365) × 1.2⁰ × (1 + 0.2×181/365) = 1.3034. That is below 2, so MAX takes **2.0×**, P35 = 745,946 € and the interest = 1× principal.
- The earliest 20 % line (row 20, 01-Oct-2026) reaches only 1.663× to the freeze (1.815× to the exit), so the **2× floor binds on every devex and land line**.

**Quirk – row 18, paid in the freeze year (08-Jan-2029):**
- `YEAR(V24)-F-1 = -1`, so the formula gives (1 + 0.12×358/365) / 1.12 × (1 + 0.12×181/365) = **1.057330**.
- The true stub (8-Jan → 30-Jun, 174 days counted inclusively) is 1.057205.
- The overstatement is **+6.1 k€** on row 18 and +0.2 k€ on row 19. An item paid after 30-Jun-2029 would be mis-accrued in the same way, because there is no guard.
- The IC model's Deposits!Z column copies the same formula but adds an `IF(pay>end,1,…)` guard.

**Freeze vs no freeze.** Lifting the freeze adds 3.94 m€ of interest (cash and guarantee lines only). It moves **+1.18 m€ to FEP and −1.18 m€ from ALSI**, because each euro of FEP interest is worth 0.3 € net to FEP (it reduces the 70/30 pot).

### 2.6 Waterfall between FEP and ALSI (U30:Y44)

| Line | Cell | Formula | Freeze (V) € | No freeze (X) € | € per MW (260) |
|---|---|---|---:|---:|---:|
| Buyer reimbursement (land, bonds, connection fee) – memo | V33 | `=V49+V50+V51+V52+V53` | 74,385,881 | 74,385,881 | 286,100 |
| Exit premium | V34 | `=$V$26*$V$6` (150,000 × 260) | 39,000,000 | 39,000,000 | 150,000 |
| − Transaction costs | V35 | `=$V$27` | 500,000 | 500,000 | 1,923 |
| − Total interest (all lines, all categories) | V36 / X36 | `=Q49` / `=S49` | 10,524,666 | 14,469,522 | 40,479 / 55,652 |
| − DSA | V37 | `=$V$22*$V$6*YEARFRAC($V$23,$V$25)` (30/360 US: 3.1806 yr) | 2,067,361 | 2,067,361 | 7,951 |
| = Distributable proceeds | V38 | `=V34-V35-V36-V37` | 25,907,973 | 21,963,117 | 99,646 / 84,474 |
| Devex (deducted from FEP) | V39 | `=V48+V54` (Fee + Devex principal) | 3,048,019 | 3,048,019 | 11,723 |
| FGX (FEP) share 70 % | V41 | `=V38*$V$21` | 18,135,581 | 15,374,182 | |
| **FEP net = interest + 70 % − devex** | V42 | `=V36+V41-V39` | **25,612,228** | **26,795,685** | 98,509 / 103,060 |
| Ventus+Solmentore (ALSI) share 30 % | V43 | `=V38*(1-$V$21)` | 7,772,392 | 6,588,935 | |
| **ALSI total = 30 % + DSA** | V44 | `=V43+V37` | **9,839,753** | **8,656,296** | 37,845 / 33,293 |

The tie-out holds: FEP net + ALSI total + devex principal + transaction costs = 39.0 m€ = the premium.

**Explicit in the file:**
- FEP funds every line.
- Interest is deducted from the premium before the split, at 12 % cash, 6 % guarantee and 20 %/2× devex and land.
- The DSA is deducted before the split. The split is 70/30.
- FEP's devex and fee principal (3.05 m€) is borne by FEP, which recovers it only through the 2× "interest". There is no separate principal repayment from the premium.
- The buyer reimburses land, bonds and connection fee at exit (V33, comments O14/O15).
- There are two variants, freeze and no freeze. The exit is 31-Dec-2029.

**Inferred or not stated:**
- Order and priority. The file subtracts, it does not run a waterfall, and **V38 has no floor**: a negative distributable would be shared 70/30, so ALSI could end with a negative share.
- Who receives the buyer reimbursement. FEP is implied because it funded the lines. Note that V33 adds bank guarantees, which are not cash: the buyer replaces them.
- Treatment if the project fails. The comments say the TSO deposit and the supplementary bonds are "refunded if project fails". The fee, devex and land are lost, and the file gives FEP no claim on ALSI.
- The file has no hurdle, no IRR test, no catch-up and no promote, and FEP's costs are not shown.
  - FEP's bank-guarantee fee is missing. The IC model has a 6 % cost, which almost offsets the 6 % "interest".
  - FEP's cost of capital, signing fee (250 k€) and DD costs (150 k€) are also missing. The older 20260925 version had them as "Case B".
- The premium basis is 260 installed MW, whereas the IC model uses 200 MVA. The file does not cap the DSA (the IC model caps it at 10 k€/MW; 7,951 €/MW here, so it would not bind).

### 2.7 FX and units

- One rate, V5 = **370 HUF/EUR** ("FY rate"), is used in both directions.
  - HUF-native lines: `J = I/$V$5`.
  - EUR-native lines (rows 14, 21, 23, 45, 46; 558,150 € in total): `I = J*V5`.
  - Row 20 mixes the two: `25*(1,200,000+180,000) HUF + 3,000 € × 28 × FX`.
- The IC model uses 365 (Inputs E11), but its pasted BORA_Devex copy also used 370.
- At 365, the HUF-native lines rise by 1.37 %: the total becomes 78.49 m€ (**+1.05 m€**), the indirect connection fee +0.67 m€, and the performance bond +0.25 m€.
- Amounts are in units of HUF and EUR, not k€. Per-MW figures divide by V6 = 260 (installed wind), including for BESS-driven lines.

### 2.8 Formula quality

1. **Hardcodes inside formulas** (all should be named inputs):
   - 3,000,000 (fee per bid); 0.005, 0.025 and **1.5** (bond rates and ÖSZ-6 uplift); `4500000*200` (rate × **hardcoded 200 MW**).
   - 73,071,306; 508,345,804; 38,000,000 × 0.2 / 0.4 × 6 (land); 72,000,000/280 × 1.15; 55,000,000/280 × 1.15; 7,000,000/40; 199,000 × 2.
   - 1,325 × 1.5; 187,636; 30,000; 250,000; 5,000,000; 3,500,000; 12,360,000 × 1.2; (7.20 + 10.18 + 4.36 m) × 1.5; 1,147,500; 30,000,000 × 1.15.
   - 150,000 × 300 and 50,000 × 300 (unit of the "300" unknown); 2,500,000 × 2; 1,500,000 × 2.
   - 0.1 (contingency); 4 (BESS hours); 365; the 2029 dates inside `DATE()`.
   - Unexplained buffers: ×1.15, ×1.2, ×1.5.
2. **MW basis inconsistent:**
   - The TSO bond is on a hardcoded 200 MW.
   - The connection fee, bonds, design fees, premium and DSA are on V6 = 260 MW installed.
   - There is no "grid connection capacity (MVA)" input.
   - The indirect connection fee on 260 MW (I18) is very probably wrong: +11.85 m€ principal, plus 12 % interest on it.
3. **Interest formula:** the freeze-year bug and the missing after-end guard (§2.5). The no-freeze R column has no final stub, so it is correct only for a 31-Dec exit.
   The floor is switched by the rate `K ≥ V19` rather than by category. "Continuous" lines carry a typed year and G = 0.
   Whole years ignore leap days, while the stubs count both ends.
4. **Contingency (row 48):**
   - It is built by explicit cell subtraction, so a new bond row would silently enter the base.
   - Its base includes land and the fee (both buyer-reimbursed or MEKH items), yet the result is booked as non-reimbursed "Devex".
   - The I, J and N formulas list the excluded cells in different orders. This is harmless, but it is a copy-paste smell.
5. **Totals:** rows 51/53 list cells explicitly instead of using SUMIF on L (fragile). The summary block (V48:V54) is SUMIF-based, which is good, and it ties: V55 = J49.
6. **Waterfall:** no floor on V38. V33 adds non-cash guarantees to "reimbursement". The DSA uses YEARFRAC 30/360 while the interest uses Act/365.
   X33/X35/X39 point to the V column (harmless today).
7. **Supplementary bond rows 9–13:** the amounts and rates are blank but the dates and comments are kept. This is consistent with GYIK 70; the IC model also excludes them.
8. **Labels:** "FY rate (HUF/EUR)" is fine. Party labels do not match the brief (FGX and Ventus+Solmentore vs FEP and ALSI). "HEPURA" is not expanded.
9. **No errors (#REF!, #DIV/0!), circularity or external links.** Cached values match a Python recomputation of P and R to 0.00 € on every line.

---

## 3. File B – `BORA_tulepites.xlsx` (overbuild penalty)

### 3.1 Translation of the sheet

| Col | Hungarian header | English | Unit |
|---|---|---|---|
| A | – | Scenario no. | # |
| B | Csatlakozási kapacitás | Grid connection capacity | MW (≈ MVA) |
| C | Túlépítés mértéke | Overbuild rate | % points |
| D | Beépített kapacitás | Installed (committed) capacity | MW |
| E | Megengedett eltérés (10%) | Permitted deviation (10 %) | MW |
| F | Tényleges eltérés | Actual deviation (shortfall) | MW |
| G | Tényleges eltérés | Actual deviation (shortfall) | % points |
| H | Szükséges-e visszafizetés? ("nem"/"igen") | Is a forfeit due? ("no"/"yes"). It says "repayment", but it means forfeiture | text flag |
| I / J | Befizetett teljesítési biztosíték | Performance bond posted | HUF / EUR |
| K / L | Elvesztett teljesítési biztosíték | Performance bond forfeited | HUF / EUR |
| M | Biztosíték / teljes beruházási költség | Forfeit / total investment cost | % points |
| P1:R3 | Teljesítési biztosíték – Szél / BESS | Performance bond rate – wind / BESS | HUF/MW (BESS should be HUF/MWh) |
| P5:Q5 | EUR/HUF | FX (actually HUF per EUR) | 370 |

### 3.2 Inputs and formulas (row 3 shown; rows 4–6 identical except where noted)

| Cell | Formula / value | Meaning |
|---|---|---|
| Q2 | `=406000000*0.025` → 10,150,000 | Bond per wind MW = 2.5 % × 406 mHUF. **No ÖSZ-6 uplift** |
| Q3 | `=200000000*0.025` → 5,000,000 | BESS bond per **MWh** (labelled HUF/MW). **Not used** |
| Q5 | 370 | HUF/EUR |
| B3:B6 | 200 (typed) | Grid capacity |
| C3:C6 | 50, 50, 20, 10 (typed) | Overbuild % |
| D | `=B*(100+C)*0.01` | Committed installed MW |
| E | `=D*0.1` | Tolerance in MW. **Not used by the logic** |
| F3 | `=300-217` → 83 | Shortfall. The built MW (217) is embedded in the formula |
| F4:F6 | 10, 30, 7 (typed) | Shortfall MW |
| G | `=(F/D)*100` | Shortfall % of committed |
| H | `=IF(G<=10,"nem","igen")` | Cliff test (≤ 10 % means no penalty) |
| I | `=D*$Q$2` | Bond posted on committed MW |
| J | `=I/$Q$5` | in EUR |
| K | `=IF(G<=10,0,F*$Q$2)` | **Forfeit = the whole shortfall MW × bond per MW**, but only if the shortfall exceeds 10 % |
| L | `=K/$Q$5` | in EUR |
| M | `=L/(2000000*D)*100` | Forfeit as % of capex, assuming a **hardcoded 2.0 m€/MW** |

### 3.3 Penalty mechanics and results

Rule (kiírás V.2 m, XII.12, XIII.15; GYIK 28, 30, 44, 77):
- The realised installed MW may be up to 10 % below the MW committed in the bid (incl. overbuild). Within that, there is no sanction.
- Beyond 10 %, the bidder loses the performance bond **in proportion to the total shortfall**, not just the excess over 10 %. The forfeit is bond × shortfall / committed = shortfall MW × bond per MW.
- Wind and BESS are tested separately (GYIK 77).
- Only a downward deviation counts. Building more than committed carries no sanction.
- The test is made after **full commissioning** (GYIK 12: the protocol of the last turbine's first parallel connection).

The file's logic matches this for wind at the minimum bond rate.

| Scen. | Grid MW | Overbuild | Committed MW | No-penalty minimum build (90 %) | Shortfall MW | Built MW | Shortfall % | Forfeit? | Bond posted (file) | Forfeit (file) | Forfeit with 1.5 uplift (as in file A) | Forfeit % of capex (file) |
|---|---:|---:|---:|---:|---:|---:|---:|---|---:|---:|---:|---:|
| 1 | 200 | 50 % | 300 | 270 | 83 | 217 | 27.67 % | yes | 3,045.0 mHUF / 8.23 m€ | **842.45 mHUF / 2.28 m€** | 1,263.7 mHUF / 3.42 m€ | 0.379 % |
| 2 | 200 | 50 % | 300 | 270 | 10 | 290 | 3.33 % | no | 3,045.0 mHUF / 8.23 m€ | 0 | 0 | 0 |
| 3 | 200 | 20 % | 240 | 216 | 30 | 210 | 12.50 % | yes | 2,436.0 mHUF / 6.58 m€ | **304.5 mHUF / 0.82 m€** | 456.8 mHUF / 1.23 m€ | 0.171 % |
| 4 | 200 | 10 % | 220 | 198 | 7 | 213 | 3.18 % | no | 2,233.0 mHUF / 6.04 m€ | 0 | 0 | 0 |

There is a single threshold (10 %) with a **cliff**. Just above it, the forfeit jumps from 0 to 0.1 × committed MW × 10.15 mHUF:
304.5 mHUF for 300 MW, 243.6 mHUF for 240 MW and 223.3 mHUF for 220 MW (×1.5 with the uplift). Between the cliff and a 100 % shortfall the forfeit is linear in the shortfall MW.

**Timing.** The file has no dates. From the tender:
- The bond is posted within 15 days of the decision (30-Dec-2026 in file A).
- The shortfall is tested after full commissioning, by the COD deadline at the latest. The forfeit is cashed by a MEKH decision after that test.
- The BG must stay valid until **COD deadline + 6 months** (XII.8).
- On an RtB sale (31-Dec-2029) the exposure passes to the buyer, so it should enter the RtB price (as the IC model's "promise price effect" does for the other commitments). On the COD and HTM routes it is a project cost at COD.

### 3.4 Gaps and risks in file B

1. **The ÖSZ-6 uplift is missing.** File A posts 150 % of the minimum bond, so the bond per wind MW is 15.225 mHUF, not 10.15. The penalty is understated by one third.
2. **BESS is ignored.** Its bond is 520 MWh × 5 mHUF × 1.5 in the file A configuration. The separate 10 % BESS test (GYIK 31, 77) is not modelled.
3. **Conflicting regime – full call.** Overbuild is an **ÖSZ-3 preference commitment**. Kiírás XIII.17 and GYIK 72 say that if a commitment is not met, "in part or in full", MEKH has no discretion and calls the performance bond. The answer implies the **whole** bond, e.g. 4,567.5 mHUF / 12.34 m€ for 300 MW wind at 1.5×.
   GYIK 30, which concerns overbuild specifically, gives the pro-rata 10 % rule that file B uses. **Legal opinion needed.**
4. **Group exclusion.** A call of the performance bond excludes the **whole company group** from capacity tenders for 3 years (VII.2 n, GYIK 71). It is unclear whether a partial (pro-rata) forfeit counts as a call. **Strategic risk for FEP.**
5. **Partial loss of the connection right.** Above the tolerance, the connection right is also lost **in proportion to the total shortfall** (GYIK 28, 30, 31). If that were applied to the 200 MVA, scenario 1 would keep only ≈ 145 MVA and scenario 3 ≈ 175 MVA. The file does not model this. The mechanism (MVA or MW) is to be confirmed.
6. **Hardcodes and labels:**
   - The built MW is embedded in `=300-217`; B and C are typed per row; the 2.0 m€/MW capex is hardcoded.
   - E is computed but not used. The H flag is text ("igen"/"nem") rather than 1/0.
   - FX is labelled "EUR/HUF", and the BESS rate unit is wrong.
   - Percentages are in points (×100), contrary to the standard.
7. **Inconsistent with file A.** File B tests 300/240/220 MW, while file A is built on 260 MW. The two files are not linked.

---

## 4. Cross-check with the IC model (`docs/reference_analysis/bora_models.md`; Deposits, Devex_FEP, BORA_Devex and MOIC read directly)

| Topic | New devex (A, 02-Oct-26) | IC model v4 | Comment |
|---|---|---|---|
| FX | 370 | **365** (Inputs E11). The pasted BORA_Devex (20260925) copy uses 370 | Pick one. +1.05 m€ on the file A total at 365 |
| Wind / BESS | 260 MW; 130 MW / 520 MWh (50 %, 4 h) | S1: 261 MW; 90 MW / 360 MWh. Old source: 200 MW, 100 MW / 200 MWh | Different strategy |
| Bid bond | 1,047.8 mHUF | 889.8 mHUF (0.5 % × 177,966) | Base: wind MW × 406 + BESS MWh × 200 |
| Performance bond | gross 7,858.5; payable line 6,810.7 mHUF | gross 6,673.7; payable 5,783.9 mHUF | Both apply 1.5× and net off the bid bond. **Netting** a BG bid bond is presentational: MEKH releases a BG bid bond and converts only a cash one (XII.11) |
| TSO security | 900 mHUF, cash, 12 % | 900 mHUF, cash, 12 % | Same |
| Supplementary security | 0 | 0 (GYIK 70). The old source had 1,624 + 812 + 406 − 1,218 mHUF | Same |
| Indirect connection fee | 73.07 × **260** − 900 = 18,098.5 mHUF | 73.07 × **200** − 900 = 13,714.3 mHUF | **+4,384 mHUF in file A, likely an error** |
| Direct connection fee | 508.3 mHUF | 508.3 mHUF | Same |
| Dates | Bid 30-Oct-26; perf. bond and TSO 30-Dec-26; GCA 08-Jan-29; **RtB 24-Mar-29** | Same securities dates; **RtB 30-Jun-29**, RtB sale 31-Mar-30 | |
| Interest end | Freeze 30-Jun-29 (no restart) or to the exit 31-Dec-29 | Accrues to 30-Jun-29, **pauses until 01-Jan-30, restarts to the sale 31-Mar-30** (Deposits Z/AB, Devex_FEP N/P) | Same per-line convention (incl. the same-year quirk); the IC adds an after-end guard |
| Devex | 3,473 k€ excl. bonds (fee 16 + land 425 + devex 3,032 incl. contingency) | 3,138 k€ incl. contingency (FEP 2,662 + land 425) | MW-scaled design fees (260 MW, 520 MWh). "Other costs (HR, car)" 50 k€ dropped. Financial costs now FEP-paid. The contingency base now includes land and the fee |
| Premium | 150 k€/MW × **260 installed** = 39.0 m€ | 150 × **200 MVA** = 30.0 m€ (Inputs E149 switch) | +9.0 m€ |
| DSA | 2.067 m€ (260 MW, 26-Oct-26 → 31-Dec-29, 30/360) | 2.238 m€ (cap 10 k€/MW) | |
| Waterfall | premium − costs − interest − DSA, 70/30, no floor, devex principal borne by FEP | Same claims "in full" (MOIC G143:G154). **ALSI floored at −DSA** (G154). Devex principal repaid only if Inputs E71 = 1 (default: land only) | Same logic. The IC adds a buyer reimbursement of 41.9 m€ and the price effect |
| FEP costs outside the waterfall | none | Guarantee fee 6 % (−3.68 m€), signing fee 250 k€, DD 150 k€, 9 % CIT | Must be added in the new model |
| Overbuild penalty | file B (pro-rata, min. bond) | Not modelled ("no bond forfeiture or failure scenarios") | New in the bidding model |

---

## 5. Formula-ready line-item table (for `Inputs` and `Calc` in the new model)

Legend:
- **C12**: 12 % p.a. on cash, from the payment date to `Accrual_End`.
- **G6**: 6 % p.a. on the guarantee face value, from the issue date to `Accrual_End`.
- **A20×2**: 20 % p.a. with the floor `MAX(factor, Min_Multiple_AtRisk = 2.0)`.

In all cases interest compounds annually on 31-Dec, stubs are Act/365 with both ends counted, and
`Accrual_End = IF(Freeze_On=1, MIN(Exit_Date, Freeze_Date), Exit_Date)`.

Party = funder. All lines are 100 % FEP; ALSI's claims are the DSA and 30 % only. Amounts are native, at the file A configuration (260 MW, 130 MW / 520 MWh).

| Item | Party | Amount | Unit | Date | Interest rule | Refundable / recovered | Category | Driver (proposed named inputs) |
|---|---|---:|---|---|---|---|---|---|
| Tender participation fee | FEP 100 % | 6,000,000 | HUF | 30-Oct-2026 | C12 | No (MEKH fee) | MEKH fee | `Fee_Per_Bid × Bids` |
| Bid bond | FEP 100 % | 1,047,800,000 | HUF | 30-Oct-2026 | G6 | Yes – released after award; called if no performance bond | MEKH security (BG) | `Bid_Bond_Pct × (Wind_MW × Bench_Wind + BESS_MWh × Bench_BESS)` |
| Performance bond (payable line, net of the bid bond) | FEP 100 % | 6,810,700,000 | HUF | 30-Dec-2026 | G6 | Yes – released after COD verification; replaced by the buyer at exit; at risk (full call or pro rata) | MEKH security (BG) | `Perf_Bond_Pct × Uplift_OSZ6 × base − Bid_Bond` (recommend modelling it **gross**, without netting) |
| Connection request security | FEP 100 % | 900,000,000 | HUF | 30-Dec-2026 | C12 | Yes – credited against the indirect connection fee; refunded on failure | TSO deposit (cash) | `TSO_Sec_Per_MW × Grid_MVA` |
| Supplementary grid security 2027–30 and refund | FEP 100 % | 0 | HUF | 23-Apr-27/28/29/30; refund 13-Mar-28 | G6 | n/a (GYIK 70) | TSO security (BG) | keep as a zero line with a switch |
| Indirect connection fee (net of the TSO security) | FEP 100 % | 18,098,539,560 (file) / 13,714,261,200 (on 200 MVA) | HUF | 08-Jan-2029 | C12 | No; **reimbursed by the buyer at exit** | Connection fee | `Conn_Fee_Per_MVA × Grid_MVA − TSO_Security` (**to confirm MVA vs MW**) |
| Direct connection fee | FEP 100 % | 508,345,804 | HUF | 08-Jan-2029 | C12 | No; reimbursed by the buyer | Connection fee | lump-sum input |
| Land reservation fee | FEP 100 % | 55,650 | EUR | 13-Feb-2027 | A20×2 | Reimbursed by the buyer at exit | Land | `Land_Res_Per_WTG × WTGs × 1.5` |
| Land purchase, 1st part (20 %) | FEP 100 % | 45,600,000 | HUF | 13-Feb-2027 | A20×2 | Reimbursed by the buyer | Land | `Land_Price × 20 % × 6` |
| Land purchase, 2nd part (40 %, at RtB) | FEP 100 % | 91,200,000 | HUF | 24-Mar-2029 | A20×2 | Reimbursed by the buyer | Land | `Land_Price × 40 % × 6` |
| Grid connection plan – design | FEP 100 % | 76,885,714 | HUF | 11-Aug-2028 | A20×2 | No | Devex – design | `72 mHUF / 280 MW × Wind_MW × 1.15` |
| Grid connection plan – admin | FEP 100 % | 750,544 | HUF | 11-Aug-2028 | A20×2 | No | Devex – permits | `187,636 × Plants` |
| Additional land – legal and agent fee | FEP 100 % | 65,580,000 | HUF (mixed HUF/EUR) | 01-Oct-2026 | A20×2 | No | Devex – land/legal | `25 × 1.38 mHUF + 3,000 € × Lots × FX` |
| Met masts (2) | FEP 100 % | 398,000 | EUR | 01-Mar-2027 | A20×2 | No | Devex – studies | `199,000 × 2` |
| Met mast design | FEP 100 % | 3,600,000 | HUF | 01-Jan-2027 | A20×2 | No | Devex – studies | |
| EYA | FEP 100 % | 4,500 | EUR | 28-Jun-2028 | A20×2 | No | Devex – studies | |
| Land reclassification – admin | FEP 100 % | 960,000 | HUF | 25-Dec-2027 | A20×2 | No | Devex – permits | `30,000 × (WTGs + Plants)` |
| Land reclassification – report | FEP 100 % | 8,000,000 | HUF | 28-Jul-2027 | A20×2 | No | Devex – studies | `250,000 × (WTGs + Plants)` |
| Zoning plan modification – design | FEP 100 % | 50,000,000 | HUF | 23-Feb-2028 | A20×2 | No | Devex – permits | `5 mHUF × Municipalities` |
| Municipality donation | FEP 100 % | 35,000,000 | HUF | 23-Feb-2028 | A20×2 | No | Devex – community | `3.5 mHUF × Municipalities` |
| Bird/bat monitoring | FEP 100 % | 14,832,000 | HUF | 30-Dec-2027 | A20×2 | No | Devex – studies | |
| Environmental – admin fee | FEP 100 % | 40,000 | HUF | 28-Apr-2028 | A20×2 | No | Devex – permits | |
| Environmental – assessment fee | FEP 100 % | 32,610,000 | HUF | 29-Mar-2028 | A20×2 | No | Devex – permits | |
| Geotechnical | FEP 100 % | 8,000,000 | HUF | 28-Jul-2027 | A20×2 | No | Devex – studies | |
| Archaeological | FEP 100 % | 8,000,000 | HUF | 28-Jun-2027 | A20×2 | No | Devex – studies | |
| Topography | FEP 100 % | 8,000,000 | HUF | 28-Jun-2027 | A20×2 | No | Devex – studies | |
| Building permit, turbines – admin | FEP 100 % | 4,590,000 | HUF | 28-Jan-2028 | A20×2 | No | Devex – permits | `1,147,500 × Plants` |
| Building permit, turbines – design | FEP 100 % | 138,000,000 | HUF | 28-Jan-2028 | A20×2 | No | Devex – design | `30 mHUF × Plants × 1.15` |
| Cable right – admin | FEP 100 % | 45,000,000 | HUF | 28-Jan-2028 | A20×2 | No | Devex – permits | `150,000 × 300` (unit of the 300 unknown) |
| Cable right – design | FEP 100 % | 111,838,967 | HUF | 28-Jan-2028 | A20×2 | No | Devex – design | `(substation design + topography + 38 m)/280 × Wind_MW × 1.15` |
| Cable right – compensation | FEP 100 % | 15,000,000 | HUF | 29-Nov-2027 | A20×2 | No | Devex – land | `50,000 × 300` |
| Building permit, substation – admin | FEP 100 % | 6,000,000 | HUF | 28-Jan-2028 | A20×2 | No | Devex – permits | |
| Building permit, substation – design | FEP 100 % | 58,732,143 | HUF | 28-Jan-2028 | A20×2 | No | Devex – design | `55 mHUF / 280 × Wind_MW × 1.15` |
| Building permit, BESS – admin | FEP 100 % | 4,590,000 | HUF | 28-Jan-2028 | A20×2 | No | Devex – permits | |
| Building permit, BESS – design | FEP 100 % | 91,000,000 | HUF | 28-Jan-2028 | A20×2 | No | Devex – design | `7 mHUF / 40 × BESS_MWh` |
| HEPURA permit – admin (W+BESS) | FEP 100 % | 20,000,000 | HUF | 08-Jan-2029 | A20×2 | No | Devex – permits | `2.5 mHUF × Plants × 2` |
| HEPURA permit – report (W+BESS) | FEP 100 % | 12,000,000 | HUF | 09-Nov-2028 | A20×2 | No | Devex – permits | `1.5 mHUF × Plants × 2` |
| Legal costs | FEP 100 % | 20,000 | EUR | continuous (file: from 1-Jan-2028) | A20×2 | No | Devex – legal | |
| Financial costs | FEP 100 % | 80,000 | EUR | continuous (file: from 1-Jan-2028) | A20×2 | No | Devex – financing | |
| Contingency 10 % (of non-bond lines) | FEP 100 % | 315,763 | EUR | continuous (file: from 1-Jan-2029) | A20×2 | No | Devex – contingency | `Contingency_Pct × SUMIF(non-security lines)`. Decide whether land and the fee belong in the base |
| *Memo:* DSA to ALSI | ALSI receives | 2,500 | EUR/MW/yr | 26-Oct-2026 → exit | none | – | Developer fee | `DSA_Rate × Wind_MW × YEARFRAC(Signing, Exit)` |
| *Memo:* sale costs | deducted | 500,000 | EUR | exit | none | – | Transaction | |
| *Memo:* overbuild forfeit | buyer (RtB) or SPV (COD/HTM) | `IF(Shortfall/Committed>Tol,1,0) × Shortfall_MW × Bench_Wind × Perf_Bond_Pct × Uplift` | HUF | post-COD test | none | Lost | MEKH penalty | see §3 |

---

## 6. Proposed waterfall definition (plain English)

**What it is.** The FEP–ALSI funding and distribution agreement for the Bora development. FEP funds everything. ALSI develops the project. Both share the RtB sale.

1. **Funding.** FEP pays every devex, land, fee, cash-deposit and connection-fee line, and procures every bank guarantee required for MEKH and the TSO. ALSI funds nothing. Each line becomes an FEP claim with its own return class:
   - **Cash class, 12 % p.a.:** the tender fee, the TSO cash security and the connection fees.
   - **Guarantee class, 6 % p.a. on face value:** the bid bond, the performance bond and the supplementary securities.
   - **At-risk class, 20 % p.a. with a floor of 2.0× principal:** devex, land and contingency.
   - **Accrual:** from each payment date, compounding annually on 31-Dec, with Act/365 stubs. It runs to the exit date, or stops at the freeze date (30-Jun-2029) if the freeze option is chosen.
2. **At exit (RtB sale)** the buyer pays the RtB premium (premium per MW × the agreed MW basis) and reimburses the recoverable principal (land, cash deposits, connection fees). The buyer also replaces FEP's guarantees.
3. **Order of payment from the exit proceeds:**
   1. Sale and transaction costs.
   2. FEP's recoverable principal (paid from the buyer reimbursement).
   3. FEP's accrued return on every line (cash 12 %, guarantee 6 %, at-risk MAX(20 %, 2×)).
   4. ALSI's development success fee (DSA): 2,500 €/MW/yr on installed MW from signing to exit.
   5. The remainder is split **70 % FEP / 30 % ALSI**.
4. **Not repaid:** FEP's at-risk principal (devex and the fee) is **not** repaid separately. FEP recovers it only through the at-risk return in step 3.3. This is the file's convention ("Devex (deducted from FEP)") and the IC model's default.
5. **No other tiers:** there is no IRR hurdle, catch-up or promote. The split is a flat 70/30 at any outcome.
6. **Shortfall rule (not in the file; proposed):** each step is paid with `MIN(cash available, amount due)`. If the premium does not cover step 3.3, FEP's return is cut. If it does not cover step 3.4, the DSA is cut, and ALSI receives no negative share.
   The file itself shares a negative remainder 70/30. The IC model floors ALSI at −DSA.
7. **Project failure (inferred):**
   - FEP recovers only the refundable securities: the TSO cash deposit, the bid bond if not awarded, and any supplementary securities.
   - Fees, devex and land are lost, and FEP has no claim on ALSI.
   - Any forfeited MEKH bond (late COD, unmet ÖSZ commitment, overbuild shortfall) is a loss before the split if it falls before exit. After an RtB exit it is the buyer's, priced into the premium.

Standards-compliant build:
- One `Funding` block on Calc with the line items in rows (from Inputs).
- Scalar helper columns per line: stub-1 days, whole years, stub-2 days, factor, floor flag and return. Each formula stays short, and there is an `N(pay ≤ end)` guard.
- A `Waterfall` block with "cash available after X" subtotals.
- Checks:
  - FEP net + ALSI total + FEP at-risk principal + sale costs = premium.
  - The return total ties to the developer's file (Q49/S49) within `Tolerance`, after the agreed corrections.
  - The waterfall never pays more than the cash available.

**Decisions for the deal team** (before SPEC §4/§5):
1. Premium MW basis: 260 installed (developer) or 200 MVA (IC).
2. Indirect connection fee basis: 200 MVA or 260 MW.
3. FX: 365 or 370.
4. Freeze or no freeze, and the exit date. Is the RtB still 24-Mar-29 / 30-Jun-29?
5. Whether to fix the same-year interest quirk or replicate it for tie-out.
6. Negative distributable: the file's (shared), the IC's (ALSI floored at −DSA) or the MIN waterfall's (ALSI ≥ 0).
7. Is devex principal repaid before the split?
8. Is a DSA cap needed?
9. Are FEP's guarantee fee, signing fee and DD costs to be modelled?
10. Overbuild: the bond uplift in the penalty, the BESS test, the full-call (XIII.17 / GYIK 72) and group-exclusion (GYIK 71) risks, and the partial loss of the connection right (GYIK 28/30).
11. Confirm the parties: FGX = FEP; Ventus+Solmentore = ALSI.
