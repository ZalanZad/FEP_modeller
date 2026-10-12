# MEKH 2026 wind capacity tender (kapacitáspályázat) – rules for the financial model

Purpose: the exact rules of the 2026 MEKH capacity-allocation tender for wind (+ BESS) grid connection, written so a financial model can implement them as formulas. The focus is project **Bora** at the **Szombathely OVIT** node.

Sources (Hungarian originals; translations are mine):

| Short name | File | Notes |
|---|---|---|
| **Kiírás** | `reference/Kapacitas_palyazati_kiiras_2026_3.pdf` (83 pp, MEKH, 30-Aug-2026) | The binding call. Cited as "Kiírás §X.y, p.N". The printed page number equals the PDF page number. |
| **GYIK** | `reference/Kapacitaspalyazat2026_GYIK.pdf` (28 pp, 80 Q&As) | MEKH FAQ. Cited as "GYIK Qn". **Not binding:** "nem képez hivatkozási alapot, jogi, kötő-, illetve bizonyító erővel nem rendelkezik" (GYIK p.1). It does state how MEKH intends to apply the call. |
| **EMR** | 11/2025. (IV. 10.) EM rendelet | The underlying decree. It is quoted inside the Kiírás but was not read separately. |
| **IC model** | `reference/Bora_IC_model_v4.xlsx`, described in `docs/reference_analysis/bora_models.md` | Used only for the cross-check in §8. |

Notation used in the formulas:

| Symbol | Meaning (Hungarian term) |
|---|---|
| `MVA_req` | requested feed-in capacity in the bid (*igényelt betáplálásra rendelkezésre álló teljesítmény*). MVA and MW are treated as equal (cos φ = 1, Kiírás §X.4 bc). |
| `MVA_awd` | feed-in capacity awarded in the MEKH decision. It equals `MVA_req` unless the bid wins only in part. |
| `MW_W` | installed wind capacity, primary technology (*elsődleges technológia beépített teljesítménye*, BT₁) |
| `MW_B`, `MWh_B` | BESS rated power (BT₂) and energy capacity (*energiatárolási kapacitás*) |
| `BI` | benchmark investment = `MW_W × 406 mHUF + MWh_B × 200 mHUF` (Kiírás Annex 2, p.34) |
| `PB_min`, `PB` | minimum and offered performance security (*teljesítési biztosíték*) |
| wd | working days (*munkanap*) |

---

## 0. Bora at a glance (Annex 1, p.33)

| Field (Annex 1 column) | Szombathely OVIT (node no. 9) |
|---|---|
| Node ID / district / coordinates | SZHO B 132.00 / Szombathelyi járás / 47.231284, 16.679985 |
| Voltage | 132 kV. This is the actual connection voltage (GYIK Q36). |
| Spare bays in existing substation | 2 |
| Earliest connection date | 30 Sep 2030 |
| Primary technology | wind (*szélerőmű*) only. A BESS is the only permitted secondary technology (Kiírás §V.2 a–c, p.7). |
| Allocable feed-in capacity | **200 MVA** |
| Flexible connection | "Nem szükséges" (not needed; GYIK Q29) |
| Expected specific **indirect** connection fee | **73,071,306 HUF/MVA** (a maximum, §V.3 c) |
| Expected **direct** connection fee | **508,345,804 HUF per connection point** |
| Mandatory TMÁ (free transfer of the connection asset) | "Nem szükséges" (not required). Voluntary TMÁ is not possible either (GYIK Q55). |

**Critical point: one company group may request at most 100 MVA per node.** Kiírás §V.7, p.9: "ugyanazon társaságcsoportba tartozó Pályázók által ugyanazon hálózati csomópont vonatkozásában együttesen benyújtható betáplálásra igényelt rendelkezésre álló teljesítmény összesített mértéke legfeljebb 100 MVA" (Pályázók = applicants; társaságcsoport = company group). §VII.36, p.17 sets out the sanction. The group must declare in the deficiency-correction step (*hiánypótlás*) which bid to drop, because the MVA in a bid cannot be changed. If the group fails to do so, "a társaságcsoportba tartozó érintett Pályázók adott hálózati csomópontra benyújtott összes érintett pályázata érvénytelennek minősül": all of the group's bids at that node become invalid. GYIK Q33, Q43 and Q52 confirm the rule. A 200 MVA Bora bid from one group therefore cannot be valid. Splitting it across several SPVs of the same group does not help, because the cap is on the group total. Two bidders that are not affiliated with each other could each bid up to 100 MVA. Whether FEP, A.RES and the JV SPV count as *kapcsolt vállalkozás* (affiliates under the Accounting Act, Kiírás §II.1 dd; Annex 12 lists "közös ügyvezetés", i.e. joint management) is a legal question to settle before bidding.

---

## 1. Timeline and procedure

### 1.1 Dates and deadlines

| # | Event | Date / rule | Source and key Hungarian phrase |
|---|---|---|---|
| 1 | Call published | 30 Aug 2026 | Kiírás cover p.1 |
| 2 | Network licensees publish the formats for their "no-arrears" certificates and the TMÁ general terms (ÁSZF) | by 7 Sep 2026. The certificates take about 8 wd. | §VII.31–VII.32, p.17; §X.4 ab, p.23 |
| 3 | Consultation (online) | 10 Sep 2026. Register 1 wd before. Questions by email up to the 2nd wd after. | §XV.1–XV.3, p.32 |
| 4 | **Submission window** | **4 Sep 2026 00:00 – 30 Oct 2026 23:59**, electronic only (SZÜF, form 'MEKH_KAPACITAS_PALYAZAT'). One bid may cover only one node. | §VI.1, p.9: "2026. szeptember 4. 00:00-tól … 2026. október 30. 23:59-ig nyújtható be"; §VI.9, p.10 |
| 5 | Participation fee received and bid security in place | by the submission deadline | §VI.5, p.9; §VIII.4–VIII.5, p.18 |
| 6 | Withdrawal | only before the deadline (§VI.12, p.10). A new bid at the same node pays the fee again: "beszámításnak nincs helye" (no netting). | §VI.12, p.10 |
| 7 | Opening by a notary | the working day after the deadline, i.e. Mon 2 Nov 2026. A notice follows on the next wd. | §IX.3, p.19 |
| 8 | Compliance check | within 8 days of the deadline. One deficiency notice at most, with a deadline of up to 5 days. The items in §VII.12 cannot be corrected: the node, MVA, voltage, technology, earliest connection date, the TMÁ statement and the fee paid in one sum. | §IX.4–IX.5, p.19–20; §VII.12, p.14 |
| 9 | Ranking and Summary evaluation (*Összefoglaló értékelés*) | before the results date | §IX.18–IX.19, p.20–21 |
| 10 | **Results announcement** (*eredményhirdetés*) | **by 13 Dec 2026** | §V.4, p.8: "legkésőbbi időpontját, amely 2026. december 13. napja"; §XI.4, p.27 |
| 11 | **Award decision** per bid (*határozat*) | Formally opened the day after the Summary evaluation is handed over (§XI.2). **Expected at the end of January 2027** (GYIK Q68: "a … határozatok meghozatala később, 2027. január végén várható"). | §XI.2–XI.3, p.27; GYIK Q68 |
| 12 | **Performance security proven** | **within 15 days of receiving the decision** | §XII.5, p.28: "határozat kézhezvételétől számított 15 napon belül igazolja a teljesítés biztosíték nyújtását" |
| 13 | **Connection request to the network licensee** (*igénybejelentés*) **plus its security** | within 15 days of receiving the decision | §XIII.6 a–b, p.30: "határozat kézhezvételét követő 15 napon belül a csatlakozási igénybejelentést … nem nyújtja be" means loss of the right |
| 14 | Bid-security guarantee: minimum validity | from submission to ≥ 8 months after the deadline, i.e. **≥ 30 Jun 2027** | §VIII.5, p.18 |
| 15 | Connection plan; network connection contract (HCSSZ) | within the deadlines of the grid codes (*üzemi / elosztói szabályzat*). The call gives no dates. | §XIII.6 c–d, p.30 |
| 16 | Community purchase option | The renewable-energy community (REC) applies to the ministry within 16 months of the decision becoming final. The Government then designates RECs, and each designated REC notifies the winner within 1 month. The winner files the amended articles within 30 days and proves registration within 5 wd. **The REC must exercise within 2 years of the decision becoming final** ("E határidő jogvesztő": a preclusive deadline). | §VII.19–VII.24, p.15–16 |
| 17 | **Final building permit proven to MEKH** | **by 1 Jan 2030**. It must cover every permit-requiring element of the wind farm and of the BESS (GYIK Q10), for the rated capacity in the decision (GYIK Q11). | §V.2 l, p.8: "legkésőbb 2030. január 1-ig kell … végleges építési engedélyt a Hivatal részére megküldeni"; §XIII.6 g, p.30 |
| 18 | MEKH establishment licence (plants ≥ 50 MW) | Construction may start only after the licence is granted. | §XIII.3–XIII.4, p.29 |
| 19 | **Earliest connection date** | **30 Sep 2030**. By this date the network licensee must make connection possible (GYIK Q60). | Annex 1; GYIK Q60 |
| 20 | Connection window | 24 months from the earliest connection date. The actual date is fixed in the connection plan or HCSSZ. No date after 30 Sep 2032 may be set. | §XIII.5, p.29; §V.5, p.8; §XI.3 d, p.27 |
| 21 | Generation / storage licence application | by the commissioning deadline | §XIII.6 h, p.30 |
| 22 | BESS connection | not earlier than the wind farm's connection | §XIII.16, p.31 |
| 23 | Proof of all commitments that earned points | by commercial operation (COD) | §XIII.17, p.31 |
| 24 | **COD deadline** | **30 Sep 2032** for all nodes. COD = first parallel connection to the public grid, evidenced by the licensee's protocol. Commissioning in stages is possible if the licensee agrees, and the 10 % tolerance is tested after full commissioning (GYIK Q12). | §XIII.5, p.29: "egységesen 2032. szeptember 30. napja"; §XIII.10, p.30 |
| 25 | Performance-security guarantee: minimum validity | from posting to ≥ 6 months after the COD deadline, i.e. **≥ 31 Mar 2033** | §XII.8, p.28 |
| 26 | Release of the performance security | MEKH starts the release within 10 wd of its post-COD compliance decision (§XIII.9 b). That decision also ends the effect of the award decision (§XIII.14). | §XII.13, p.29 |

There is no "RtB" (ready-to-build) milestone in the call. The binding pre-construction milestone is the final building permit by 1 Jan 2030 (row 17), plus the MEKH establishment licence (row 18).

The "elnyert jogosultság teljes időtartama" (full term of the awarded right) runs from receipt of the decision until the decision ceases to have effect (GYIK Q78). For that whole period the winner must keep meeting the tender conditions (§VII.1 eh, p.12).

### 1.2 Evaluation algorithm (Kiírás §IX.18, p.20–21)

1. Remove invalid bids (§VI.2, VI.17, VII.10, VII.12, VII.36, VIII.12, IX.22).
2. Assign each bid to its node and rank by total score, highest first (§X).
3. Winners are the ranked bids whose **cumulative requested MVA** "nem éri el" the node's allocable capacity (literally "does not reach" it; see A21 in §7).
4. **Group cap of 200 MVA in total** (§V.6). If a group's provisional wins exceed 200 MVA, its bids are rejected one at a time, starting from the bottom of the group's own preference order (Annex 11). Several groups are handled in the same iteration, the declared partial-award minimum counts, and the process restarts from step 3. Inconsistent preference declarations within a group mean rejection by score, lowest first (§VII.17).
5. **Partial award:** the first bid that does not fit can win in part, up to the remaining capacity, if the bidder declared a minimum it would accept (§VII.14). That minimum must be ≥ 14 MVA. Otherwise the bid is rejected. A partial award applies only to the last bid that can still be declared a winner (GYIK Q35).
6. Bids whose cumulative MVA collides with the node cap are rejected (§IX.18 g).
7. A winner that drops out later does **not** pass its capacity to the next-ranked bidder. Freed capacity goes into future tenders (GYIK Q75).
8. MEKH does **not** assess project maturity or feasibility. Ranking is purely by commitments and points (GYIK Q48).

---

## 2. Deposits, securities, guarantees and fees

Summary of who is paid what:

| Item | Paid / posted to | Amount | When | Form | Refund or release | Forfeiture |
|---|---|---|---|---|---|---|
| Participation fee (*részvételi díj*) | MEKH | 3,000,000 HUF per bid | received by 30 Oct 2026 | cash, HUF | only in 3 cases (§2.1) | otherwise kept |
| Bid security (*ajánlati biztosíték*) | MEKH | 0.5 % × `BI_bid` | in place at submission | cash deposit or bank guarantee, HUF | 10 wd after the relevant event | only if a winner fails to post the PB |
| Performance security (*teljesítési biztosíték*) | MEKH | 2.5 % × `BI_awd` × `k` (k = ÖSZ-6 multiple, ≥ 1) | ≤ 15 days after receiving the decision | cash deposit or bank guarantee, HUF | 10 wd after the post-COD compliance decision | full or proportional (§2.3) |
| Connection-request security (*igénybejelentési biztosíték*) | network licensee | set by the grid code, **not by the call** | with the connection request, ≤ 15 days after receiving the decision | set by the grid code | credited against the connection fee (GYIK Q40) | failing to post it costs the whole right and the PB |
| Capacity-reservation / supplementary security | – | **none** | – | – | – | GYIK Q70 |
| Connection fee (*csatlakozási díj*) | network licensee | indirect HUF/MVA × MVA plus direct lump sum (§3) | as set in the HCSSZ | cash | – | failing to pay costs the whole right and the PB |

**Interest on cash deposits:** neither the call nor the FAQ provides any. Cash goes to MEKH's deposit account at the State Treasury. **Model: 0 % interest.**

### 2.1 Participation fee (*részvételi díj*) – to MEKH

- **Amount:** "A részvételi díj mértéke … 3.000.000 Ft. A részvételi díjat pályázatonként, egy összegben kell megfizetni." (3,000,000 HUF, per bid, in one sum; §VI.4, p.9.) The documentation fee is 0 HUF (§VI.3).
- **Currency and account:** HUF to MEKH's Treasury account HU22 1003 2000 0171 4373 0000 0000, reference "SZKP részvételi díj". It must arrive in full by the deadline, no instalments, and only the applicant itself may pay it (§VI.5, p.9).
- **Correction:** none. A fee not paid in one full sum makes the bid invalid (§VII.12 g, p.14).
- **Refund** within 10 wd, only if (§VI.7–VI.8, p.9–10):
  - MEKH withdraws the call;
  - the procedure fails because there were no bids, an act by some bidder voided it (except for that bidder), or ≥ 50 % of bidders are excluded;
  - the bid is withdrawn before the deadline.
- **Model:** `Fee = 3.0 mHUF × N_bids`, paid by 30-Oct-2026. 0 % recovery for losing, rejected or invalid bids.

### 2.2 Bid security (*ajánlati biztosíték*) – to MEKH

- **Definition:** security given before the winners are selected, "a pályázatban vállalt kötelezettségek biztosítékául" (to secure the commitments in the bid; §II.1 a, p.3).
- **Formula** (§VIII.2, p.17–18): "Az ajánlati biztosíték összege az adott pályázathoz tartozó beépített teljesítmény – villamosenergia-tároló esetén energiatárolási kapacitás – és a 2. mellékletben található fajlagos benchmark beruházási érték szorzata alapján számolt benchmark beruházási összeg 0,5%-a." (0.5 % of the benchmark investment, which is installed capacity, or storage energy for a BESS, times the Annex 2 unit value.) "…minden technológia után megfizetendő" (due for every technology).

  `BidSec = 0.5 % × (MW_W_bid × 406 + MWh_B_bid × 200)` mHUF, i.e. **2.03 mHUF per installed wind MW + 1.00 mHUF per BESS MWh.**

  - The basis is the **installed** MW and MWh in the bid, overbuild included. It is **not** the MVA. GYIK Q51: "indokolt a biztosíték alapjának a vállalt beépített teljesítményekhez kötése" (it is justified to tie the security base to the installed capacities committed). GYIK Q24 calls this basis "az igényelt" (the requested), meaning the capacities in the bid.
  - Annex 2 (p.34): wind 406 M Ft/MW; battery storage 200 M Ft/MWh, "a villamosenergia-tárolási kapacitás alapján" (on storage capacity).
- **Form:** a cash deposit, proven by a bank debit slip, or a bank guarantee on the **mandatory Annex 5 A/B template** (§VIII.3). Deviations are not accepted; the issuing bank must contact MEKH first (GYIK Q69).
  - Cash goes in one sum to MEKH deposit account HU22 1003 2000 0171 4373 2000 0002, reference "SZKP ajánlati biztosíték" (§VIII.4).
  - The guarantee must be electronically signed by the bank, unconditional, irrevocable and payable on first demand, in favour of MEKH (§VIII.6–VIII.7).
  - The issuer must be on the MNB (central bank) register and have a seat or branch in Hungary (§VIII.8).
  - An EU-seated bank or an EUR-denominated guarantee is accepted only with a counter-guarantee (*felülgarancia*) from a bank that meets the same conditions (§VIII.11; §V.2 k).
  - **Currency HUF** (§VIII.9): "Az ajánlati biztosítékot magyar forintban (HUF) kell teljesíteni."
  - One security per bid (§VIII.10). Only the applicant itself may post it; a parent company may not (GYIK Q67). Guarantees cannot be transferred, assigned or pledged (Annex 5, point 10).
- **Timing:** in place at submission. A guarantee must be "érvényesnek és hatályosnak … legalább a pályázat benyújtásának napjától legalább a benyújtási határidőt követő 8 hónapig" (valid from the submission day to at least 8 months after the deadline, i.e. ≥ 30 Jun 2027; §VIII.5). The template's expiry date is in 2027.
  - Proof can follow in the deficiency step, but the guarantee must be valid from the submission day (GYIK Q64).
  - Not posted, or not posted properly, makes the bid invalid (§VIII.12) and is an exclusion ground (§VII.2 j).
- **Release** within 10 wd (§VIII.13–VIII.15, p.19):
  - rejected or invalid bids: after the §XI.2 decision is sent (expected end of January 2027). This includes bids made invalid by failing the deficiency step (GYIK Q65).
  - withdrawn bids: after the deadline.
  - everyone: if the procedure fails or the call is withdrawn.
  - security that cannot be matched to any bid: 10 wd after the deadline.
  - Cash is returned to the account it came from.
- **Winners** (§XII.11, p.28–29), once MEKH confirms the performance security:
  - (a) if the bid security and the PB are **both cash**, "az ajánlati biztosítékot teljesítési biztosítékká minősíti át" (MEKH reclassifies the bid security as performance security), so only the top-up is new cash;
  - (b) in every other case MEKH starts releasing the bid security within 10 wd of its information letter.
  - Partial winner: the cash bid security is reclassified pro rata to the partial award and the excess is released.
  - The PB guarantee template (Annex 5 C/D, point 12) becomes **effective the day after the bank receives MEKH's waiver of the bid guarantee**, so the two guarantees hand over one to the other.
- **Forfeiture (*lehívás*):** only if a winner fails to post the PB on time or as required. MEKH then decides on loss of the connection right and on calling the bid security (§XII.10, p.28: "a csatlakozási jog elvesztéséről és az ajánlati biztosíték lehívásáról szóló határozatot hoz"). A partial winner forfeits pro rata to the partial award. The guarantee template (Annex 5 A/B, preamble and point 5) allows a call only on this ground.

### 2.3 Performance security (*teljesítési biztosíték*) – to MEKH

- **Definition:** given after winners are selected, "a nyertes pályázatban szereplő beruházás megvalósításának biztosítékául" (to secure delivery of the winning investment; §II.1 ff, p.5).
- **Minimum** (§XII.2, p.28): "…határozatban foglalt beépített teljesítmény – villamosenergia-tároló esetén energiatárolási kapacitás – és … fajlagos benchmark beruházási érték szorzata alapján számolt benchmark beruházási összeg 2,5%-a." Due per technology.

  `PB_min = 2.5 % × (MW_W_awd × 406 + MWh_B_awd × 200)` mHUF, i.e. **10.15 mHUF per wind MW + 5.00 mHUF per BESS MWh.**

  - The basis is the **decision's** installed wind MW and BESS MWh, including the overbuild and BESS committed for points (GYIK Q30).
  - For a partial winner, the decision's MVA, wind MW, BESS MW and MWh are all reduced in proportion (`f = MVA_awd / MVA_req`), and the PB is computed on the reduced values (GYIK Q23, Q24).
- **Actual amount:** `PB = PB_min × k`, where `k = x_ÖSZ6 / 100 ≥ 1` is the ÖSZ-6 commitment. The amount is fixed in the decision (§XII.3; §X.4 bf).
- **Deadline:** proof within **15 days of receiving the decision** (§XII.5). The decisions are expected at the end of January 2027 (GYIK Q68), so the deadline falls around mid-February 2027. MEKH advises approaching the bank right after the 13 Dec 2026 results.
- **Form:** as for the bid security (cash or guarantee, HUF, Annex 5 C/D templates, §VIII.6–VIII.11 apply, §XII.9). Cash goes to the same deposit account, reference "SZKP teljesítési biztosíték" plus the decision number, in one sum (§XII.7). **The form cannot be switched later** between guarantee and cash (GYIK Q66).
- **Guarantee validity:** "…a kereskedelmi üzem kezdetére vonatkozó határidőt követő legalább 6 hónapig" (to at least 6 months after the COD deadline, i.e. ≥ 31 Mar 2033; §XII.8). The minimum tenor is tied to the deadline, not to the actual COD.
- **Release:** MEKH starts the release within 10 wd of sending its §XIII.9 b) decision, which confirms after COD that all obligations were met (§XII.13). Cash goes back to the account it came from.
- **Forfeiture: three different mechanisms.**

| Trigger | Consequence for the PB | Consequence for the connection right | Source |
|---|---|---|---|
| Any §XIII.6 event: (a) no connection request within 15 days of the decision; (b) connection-request security not posted properly; (c) connection plan late under the grid codes; (d) HCSSZ not signed in time; (e) HCSSZ ends for the winner's reasons; (f) connection fee not paid under the HCSSZ; (g) no final building permit proven by 1 Jan 2030; (h) no generation/storage licence application by the commissioning deadline; (i) no COD by 30 Sep 2032, connection not built or accreditation not obtained | **100 % called** ("lehívásra kerül, azt a nyertes Pályázó elveszíti": it is called and the winner loses it) | **whole right lost** | §XIII.6, p.29–30; §XII.12, p.29 |
| Any commitment that earned points (§X.4: ÖSZ-1 … ÖSZ-9) not credibly proven to MEKH by COD | **100 % called.** "…akkor a Hivatal köteles lehívni a teljesítési biztosítékot" (MEKH must call it). MEKH has no discretion, and **partial** non-compliance is enough. | **not** lost. GYIK Q72: "A csatlakozási jogot elvesztése nem része a mulasztásból következő jogkövetkezménynek" (losing the right is not a consequence of this failure). | §XIII.17, p.31; GYIK Q72, Q9; Annex 13/14 statements |
| Failure attributable to the winner to implement the 25 % community purchase option (EMR 7 § (9)), e.g. no shareholder meeting, no resolution, late or abandoned filing at the court of registration | **100 % called** | **whole right lost** | §VII.26–VII.28, p.16 |
| Installed capacity at COD **> 10 % below** the decision's installed MW, wind and BESS **tested separately** | **Proportional to the full shortfall**: "a teljes hiányzó beépített teljesítményének arányában veszíti el" (lost in proportion to the whole missing installed capacity) | lost in the same proportion | §XII.12, p.29; §XIII.15, p.31; §V.2 m, p.8; GYIK Q28, Q30, Q31, Q77 |
| Wind only partly commissioned by the COD deadline | proportional rule above, if the shortfall is > 10 % | right tied to the uncommissioned part is lost | §XIII.7, p.30 |
| Unfair influence or false data | bid invalid or right lost (no explicit PB rule) | lost | §XIV.3, p.31–32 |

Model formulas for the capacity tolerance:

```
s_W  = MAX(0, 1 − MW_W_built / MW_W_awd)          # tested on the decision value (GYIK Q76)
s_B  = MAX(0, 1 − MW_B_built / MW_B_awd)          # BESS tested separately on MW (GYIK Q77)
Loss_W = IF(s_W > 10%, s_W × PB_W, 0)             # reading A: technology component of the PB
Loss_B = IF(s_B > 10%, s_B × PB_B, 0)             #   PB_W = 2.5%·k·406·MW_W_awd ; PB_B = 2.5%·k·200·MWh_B_awd
# reading B (conservative): Loss = s × PB_total for the technology that fails
```

- A shortfall of 10 % or less carries no sanction (GYIK Q28): "Amennyiben 10%-ot nem meghaladó mértékű a negatív irányú eltérés, úgy az szankcióval nem jár".
- With overbuild, the 10 % is measured on the **committed installed** MW. For 20 MVA requested and 30 MW committed, ≥ 27 MW is safe (GYIK Q44).
- Building more than the decision states is not sanctioned (GYIK Q77). Raising the installed MW needs a change to the decision, and feed-in MVA never increases (GYIK Q15, Q23).
- No force majeure relief: "A felsorolt esetek az üzleti kockázat részét képezik" (the listed cases are business risk; GYIK Q47; §XIII.2, p.29).
- **Group-wide knock-on:** if MEKH calls a PB, the bidder's whole company group is excluded from capacity tenders for 3 years (§VII.2 n, p.13; GYIK Q71: "Igen").

### 2.4 Connection-request security (*igénybejelentési biztosíték*) – to the network licensee

- **Definition** (§II.1 p, p.4): "a nyertes Pályázó által az igénybejelentési nyilatkozat benyújtásával egyidejűleg a hálózati engedélyes felé nyújtandó pénzügyi biztosíték" (security the winner gives the network licensee together with the connection request).
- **Amount, form, MW basis and release:** **not given in the call or the FAQ.** They are governed by the grid codes: "ennek szabályait az üzemi és elosztói szabályzatok tartalmazzák" (GYIK Q40).
- **Credited against the connection fee:** "A biztosítékok közül csak az igénybejelentési biztosíték számít bele a csatlakozási díjba" (of the securities, only this one counts towards the connection fee; GYIK Q40).
- **Timing:** together with the connection request, within 15 days of receiving the decision (§XIII.6 a, b).
- **Failure:** whole connection right lost and the PB called (§XIII.6 b; §XII.12).
- The network licensee for node 9 is not named in the call (Annex 7 names E.ON Észak-dunántúli for the other Szombathely node, no. 8). Confirm whether the request goes to E.ON or to MAVIR.

### 2.5 Securities that do not apply

The capacity-reservation security (*kapacitáslekötési biztosíték*) and the supplementary financial security (*kiegészítő pénzügyi biztosíték*, VET 177/E–G) are **not** payable for rights awarded in this procedure. GYIK Q70: "nem kell kapacitáslekötési és kiegészítő pénzügyi biztosítékot fizetni".

### 2.6 Local contributions (commitments, not securities)

These are payments to municipalities that the bidder commits to for points (§4, ÖSZ-1 and ÖSZ-2). They are not deposits, but they become real costs once committed.

Footnote 1 (Kiírás p.8) warns of an expected law that would make **10,000 €/MW (capacity-based) and 2 €/MWh (generation-based) mandatory** for winners, split among settlements within 5 km of the turbines. A "zero" offer may therefore still end up paying the minimum.

---

## 3. Connection fees (*csatlakozási díj*) – to the network licensee

- **Legal basis:** 13/2025. (VII. 31.) MEKH rendelet (connection-fee decree), quoted in GYIK Q37–Q38 and Q41. The licensee's cost estimate of (a) direct, (b) indirect and (c) switching-station connection investments is the basis for the preliminary fee and its payment.
- **Values for Szombathely OVIT** (Annex 1, p.33; §V.3 c, p.8):
  - **Indirect:** 73,071,306 HUF/MVA. §V.3 c calls it the *maximum* specific indirect fee: "a maximális közvetett fajlagos csatlakozási díj mértékét (Ft/MVA)".
  - **Direct:** 508,345,804 HUF **per connection point**. The column is "várható közvetlen csatlakozási díj … (Ft/csatlakozási pont)" (expected direct fee, per connection point). It covers only the investment needed at the advertised voltage. It excludes the TMÁ and the producer's own step-up substation, which the producer builds, owns and maintains (GYIK Q39).
- **Formula:**

  ```
  ConnFee_HUF = 73,071,306 × MVA_awd                         # indirect, on awarded feed-in MVA
              + 508,345,804 × share_direct                   # lump sum; share = 1, or the pro-rata share if shared
              − subsidies (decree 4 § (3) up front; 8 § (4) at ex-post settlement)
  Net cash  = ConnFee_HUF − igénybejelentési biztosíték      # GYIK Q40 credit
  ```

  - If ÖSZ-8 sharing applies and the licensee serves several requests through one shared producer line and connection point, the **direct connection investment cost** is shared "az egyes igénybejelentők által igényelt betáplálási csatlakozási teljesítmények egymáshoz viszonyított … arányában" (in proportion to each requester's requested feed-in capacity; §XIII.18, p.31). That gives `share_direct = MVA_awd / Σ MVA_awd(sharing winners)`.
  - EU or other grant money that the licensee receives is deducted, up front if paid ≥ 30 days before the technical-economic offer (MGT), or else at the ex-post settlement (GYIK Q38, Q41).
- **Timing and instalments:** **not set by the call.** They follow the HCSSZ and the connection-fee decree. Not paying "a hálózati csatlakozási szerződésben előírtaknak megfelelően" (as the HCSSZ requires) costs the whole right and the PB (§XIII.6 f, p.30).
- **Not applicable to Bora:** the mandatory-TMÁ cost estimates in Annex 1 (other nodes) and the Annex 7 technical scope. Flexible connection is not needed (GYIK Q29, Q54).

---

## 4. Scoring

### 4.1 Mechanics (Kiírás §X.1–X.3, p.22–23; Annex 3, p.35; Annex 4, p.36–41)

- There are two groups of criteria:
  - **ÉSZ**, the validity criteria (pass/fail): failing one invalidates the bid.
  - **ÖSZ**, the comparative criteria: each scores 0–100 inner points, weighted by the Annex 3 weights. "A súlyozást követően törtszámok előállhatnak, ezeket két tizedesjegyre kell kerekíteni" (fractions after weighting are rounded to 2 decimals; §X.1 b).
- `Score = ROUND( Σ_i w_i × p_i / 100 , 2 )`, with a maximum of 100.
- **Tie-break** (§X.3, p.23):
  - (1) If the 2-dp totals are equal, rank on the four highest-weighted criteria only. GYIK Q32 names them: **ÖSZ-1 (20), ÖSZ-2 (15), ÖSZ-4 (15), ÖSZ-9 (15)**, summing their weighted points.
  - (2) If still equal, a **notarial lottery**: "közjegyző jelenlétében, sorsolás útján".
  - The lottery is needed only if the tied bids together request more than the remaining capacity (GYIK Q22, Q35).
- Project maturity is not scored (GYIK Q48).

### 4.2 Validity criteria and hard thresholds

| Rule | Value | Source |
|---|---|---|
| ÉSZ-1: commitment, in a fully probative private document (Annex 16), to give a **25 % purchase option** on the SPV's shares to RECs designated by the Government | mandatory | §X.4 aa, p.23; §VII.18, p.15; Annex 16 |
| ÉSZ-2: commitment to mandatory TMÁ where Annex 1 requires it | n/a for Szombathely OVIT ("Nem szükséges") | §X.4 ab, p.23 |
| Minimum requested feed-in **14 MVA** and minimum installed wind **14 MW**. This floor also applies to a partial award. | "legalább 14 MVA … és legalább 14 MW … a részleges teljesítés esetén is érvényesül ez a korlát" | §V.2 d, p.7 |
| Installed wind ≥ requested MVA (cos φ = 1) | `MW_W ≥ MVA_req` | §VII.16, p.15 |
| BESS power ≤ installed wind | `MW_B ≤ MW_W` | §VII.13, p.14 |
| BESS is the only secondary technology; no third technology | – | §V.2 b–c, p.7; §VII.13 |
| One node per bid; one plant or park per bid; the decision cannot be split into units | – | §VI.9, p.10; GYIK Q53 |
| Site: land plots of ≥ 1 ha held under lease, building right, usufruct, ownership or a pre-contract. **≤ 19 km straight-line** from the node, to its farthest point (GYIK Q46). Proof of the land right is not needed at submission (GYIK Q2). | – | §VII.1 b, p.11 |
| **Group cap per node: 100 MVA submitted** | invalidity sanction (§0) | §V.7, p.9; §VII.36, p.17 |
| **Group cap in total: 200 MVA awarded** (preference order in Annex 11) | rejection from the bottom of the preference order | §V.6, p.9; §IX.18 e |
| Experienced staff: ≥ 1 full-time employee in the group with ≥ 5 years of energy experience | – | §VII.3, p.13 |
| Exclusion grounds (arrears, convictions, PB called within 3 years, HCSSZ terminated by own fault within 3 years, …) | 17 grounds (a–q) | §VII.2, p.12–13 |
| aFRR control capability (Üzemi Szabályzat 5.1.9 (G) dd) (1)). It may be provided through an aggregator. | technical obligation | §VII.15, p.15; GYIK Q14 |

### 4.3 Comparative criteria (ÖSZ)

Each entry gives the definition, formula, bands, verification and sanction.

**ÖSZ-1 – one-off local financial contribution to the host municipalities (weight 20)** (§X.4 ba, p.23–24; §V.2 f, p.7; Annex 13)

- "egy alkalommal teljesítendő helyi pénzügyi hozzájárulás … legalább 10.000,01 EUR/MW és legfeljebb 50.000 EUR/MW" (one-off local contribution of at least 10,000.01 and at most 50,000 EUR/MW).
- x = EUR/MW, as stated in the Annex 13 declaration.
- The MW basis (installed or requested), payment date and recipients are left to "a későbbiekben megjelenő jogszabály" (a later law).

**ÖSZ-2 – annual local contribution on energy fed in (weight 15)** (§X.4 bb, p.24; §V.2 g; Annex 14)

- "A szélerőmű által a villamosenergia-hálózatba betáplált villamosenergia mennyisége alapján … legalább 2,01 EUR/MWh, legfeljebb 3 EUR/MWh" (on the energy the wind farm feeds into the grid, at least 2.01 and at most 3 EUR/MWh).
- Paid every year to the municipalities. The detailed rules will come in a later law.

**ÖSZ-3 – overbuild: installed wind above the requested feed-in capacity (weight 10)** (§X.4 bc, p.24)

- The call prints the formula as "x % = x MVA / x MW₁". **GYIK Q19 corrects it:** `x = (MW_W / MVA_req) × 100 − 100`.
- MW₁ is the primary technology's installed capacity (BT₁), and MVA equals MW at cos φ = 1.

**ÖSZ-4 – BESS ratio (weight 15)** (§X.4 bd, p.24–25; Annex 8, p.57)

- The call's text swaps the labels MW₁ and MW₂ and refers to "megpályázott kapacitás" (the capacity bid for). **GYIK Q20 settles the formula:** `x = MW_B / MW_W × 100`, i.e. BESS rated power over **installed wind** (BT₂ / BT₁), **not over MVA**.
- **Technical minimums (Annex 8):**
  - continuous output for ≥ 4 h, i.e. 0.25C → `MWh_B ≥ 4 × MW_B`;
  - ≥ 6,000 full cycles (nominal, initial values);
  - four-quadrant inverter meeting MAVIR response times;
  - annual availability ≥ 98 % (GYIK Q6);
  - BMS/EMS/SCADA compatibility;
  - technically capable of FCR, aFRR and mFRR (general accreditation, Üzemi Szabályzat 5.1.9 (B) a) ac); GYIK Q6).
- §X.4 bd adds that the BESS power must reach at least 5 % of the bid capacity; see A3 in §7.
- The BESS gets **no feed-in capacity of its own**. It may use the MVA headroom when the wind output is below it (GYIK Q18).
- A BESS that already had a building permit before the call does not count as a new commitment (GYIK Q4).

**ÖSZ-5 – accept a partial award (weight 5)** (§X.4 be, p.25; §VII.14)

- Yes or no. The bidder declares only a minimum acceptable MVA, at least 14. No other conditions may be attached (GYIK Q25).

**ÖSZ-6 – performance security above the minimum (weight 5)** (§X.4 bf, p.25)

- `x = PB / PB_min × 100` (%). "A vállalt teljesítési biztosíték mértékét fajlagosan … kell megadni" (the committed performance security is stated as a ratio).

**ÖSZ-7 – brownfield site (weight 10)** (§X.4 bg, p.25–26)

- Yes or no. At least **70 %** of the realised wind site area (hectares from the land-registry deeds, covering every turbine location; a separate BESS site is excluded) must be in the BarKA brownfield register at **any** point between submission and COD.
- The 19 km rule applies. The test is done at COD (GYIK Q3, Q46).

**ÖSZ-8 – consent to shared use of the connection point and producer line (weight 5)** (§X.4 bh, p.26)

- Yes or no. The winner consents to another winner of the same procedure sharing the connection point and producer line (VET 39/G §) and commits to a sharing agreement (VET Vhr 13/H §).
- The sharing "nem terjed ki … csatlakozási teljesítményre": it does not extend to the awarded connection capacity.

**ÖSZ-9 – EU-made equipment (weight 15)** (§X.4 bi, p.26–27)

- Two separate commitments, each at ≥ 80 % of the purchase cost:
  - **turbine units: 12 points.** These include generators, gearboxes, main bearings, yaw and pitch drives, blades, towers, converter and inverter modules, and control systems.
  - **producer transformers: 3 points.**
- Transport, construction and services are excluded from the cost. **BESS is excluded** (GYIK Q9).
- The supplier must manufacture in the EU: "az Európai Unió területén gyártó ajánlattevőktől szerzi be".
- The commitment is made in the bid form, point 9.6 (GYIK Q74).

**Points tables (inner points; weighted = inner × weight / 100):**

| ÖSZ-1 x (EUR/MW) | inner | ÖSZ-2 x (EUR/MWh) | inner |
|---|---|---|---|
| x ≤ 10,000 | 0 | x ≤ 2.0 | 0 |
| 10,000 < x ≤ 15,000 | 10 | 2.0 < x ≤ 2.1 | 5 |
| 15,000 < x ≤ 20,000 | 20 | 2.1 < x ≤ 2.2 | 10 |
| 20,000 < x ≤ 25,000 | 30 | 2.2 < x ≤ 2.3 | 20 |
| 25,000 < x ≤ 30,000 | 40 | 2.3 < x ≤ 2.4 | 30 |
| 30,000 < x ≤ 35,000 | **55** | 2.4 < x ≤ 2.5 | 40 |
| 35,000 < x ≤ 40,000 | 70 | 2.5 < x ≤ 2.6 | 50 |
| 40,000 < x ≤ 45,000 | 80 | 2.6 < x ≤ 2.7 | 60 |
| 45,000 < x < 50,000 | 90 | 2.7 < x ≤ 2.8 | 70 |
| x = 50,000 | 100 | 2.8 < x ≤ 2.9 | 80 |
| x > 50,000 | outside the offered range (A8) | 2.9 < x < 3.0 | 90 |
| | | x = 3.0 | 100 |

| ÖSZ-3 x (% overbuild) and ÖSZ-4 x (% BESS / wind) | inner |
|---|---|
| x ≤ 2 | 0 |
| 2 < x < 5 | 10 (ÖSZ-4: see A3, the 5 % minimum) |
| 5 ≤ x < 10 | 20 |
| 10 ≤ x < 15 | 30 |
| 15 ≤ x < 20 | 40 |
| 20 ≤ x < 25 | 50 |
| 25 ≤ x < 30 | 60 |
| 30 ≤ x < 35 | 70 |
| 35 ≤ x < 40 | 80 |
| 40 ≤ x ≤ 50 | 90 (ÖSZ-4: Annex 4 prints "40 ≤ x < 50"; GYIK Q20 closes the gap at 50 → 90) |
| x > 50 | 100 |

| ÖSZ-6 x (% of PB_min) | inner | Yes/no items | inner if yes |
|---|---|---|---|
| x ≤ 100 | 0 | ÖSZ-5 partial award | 100 → 5 pts |
| 100 < x < 120 | 10 | ÖSZ-7 brownfield | 100 → 10 pts |
| 120 ≤ x < 130 | 20 | ÖSZ-8 shared connection | 100 → 5 pts |
| 130 ≤ x < 140 | 30 | ÖSZ-9a EU turbine units ≥ 80 % | 80 → 12 pts |
| 140 ≤ x < 150 | 40 | ÖSZ-9b EU transformers ≥ 80 % | 20 → 3 pts |
| 150 ≤ x < 160 | 50 | | |
| 160 ≤ x < 170 | 60 | | |
| 170 ≤ x < 180 | 70 | | |
| 180 ≤ x < 190 | 80 | | |
| 190 ≤ x < 200 | 90 | | |
| x ≥ 200 | 100 (GYIK Q34 corrects "200 < x") | | |

**Weights:** ÖSZ-1 20, ÖSZ-2 15, ÖSZ-3 10, ÖSZ-4 15, ÖSZ-5 5, ÖSZ-6 5, ÖSZ-7 10, ÖSZ-8 5, ÖSZ-9 15, total 100 (Annex 3, p.35).

Implementation tip: write each band as (lower bound, inclusive flag, points).
- ÖSZ-1 and ÖSZ-2 bands are **exclusive at the lower bound and inclusive at the upper**. Exception: the top band includes only the exact maximum, 50,000 or 3.0.
- ÖSZ-3, ÖSZ-4 and ÖSZ-6 bands are **inclusive at the lower bound**. Exception: their first band is exclusive (2 < x, 100 < x).
- So the cheapest offer in each ÖSZ-1/ÖSZ-2 band is "lower bound + 0.01". For example, 20,000.01 €/MW scores 30 inner points, while 20,000 scores 20, and 2.01 €/MWh scores 5 while 2.00 scores 0.

### 4.4 Verification and sanctions for points commitments

| ÖSZ | Verified when, how | Sanction if not met |
|---|---|---|
| 1, 2 | Proof to MEKH "az erőmű kereskedelmi üzembe helyezéséig" (by COD) (§XIII.17; Annex 13/14). The mechanics come from a later law. | Full PB call (§XIII.17; GYIK Q72). How the annual ÖSZ-2 payments are enforced after COD is unclear (A25). |
| 3 | Installed wind at COD, from the §XIII.10 protocol, compared with the decision's installed MW | ≤ 10 % short: none (GYIK Q28, Q44). > 10 % short: proportional PB and right loss on the full shortfall (§XII.12). Whether §XIII.17 can also trigger a full call is open (A9). |
| 4 | BESS MW at COD; Annex 8 compliance proven before BESS commissioning with manufacturer documents (declaration of conformity, datasheets, sizing, type tests) (GYIK Q5) | As for ÖSZ-3, measured on BESS MW (GYIK Q31, Q77). Failing Annex 8: full PB call under §XIII.17. |
| 5 | – | – (the reduction applies at award) |
| 6 | PB posted for the declared amount within 15 days of the decision | Bid security called and right lost (§XII.10) |
| 7 | BarKA register at COD (GYIK Q3) | Full PB call (§XIII.17) |
| 8 | Sharing agreement concluded when there is another winner (GYIK Q16, Q59) | Full PB call (§XIII.17), presumably |
| 9 | Contracts, invoices, delivery notes, manufacturers' EU-origin declarations and the technical handover protocol, after construction and before COD (§X.4 bi) | Full PB call; the right is kept (GYIK Q9, Q72). A permit can still be issued (GYIK Q9). |

The commitments become part of the decision and cannot be changed, except for the site (§XIII.12; GYIK Q49, Q50).

### 4.5 Partial award (ÖSZ-5)

- `f = MVA_awd / MVA_req`. The decision then sets `MW_W × f`, `MW_B × f` and `MWh_B × f` (GYIK Q23). The PB is computed on these values (GYIK Q24), and the 10 % tolerance is measured against them (GYIK Q76).
- **Points are not recalculated.** The ÖSZ-3 and ÖSZ-4 ratios of the original bid stand (GYIK Q26).
- The bidder may still build more MW or MWh, but the feed-in MVA does not increase (GYIK Q23).

### 4.6 Bora strategy scores re-derived from the call

These use the IC model's strategy inputs (`bora_models.md` §3.1). The scores depend on ratios only, so they are the same at 100 or 200 MVA.

| Item | S1 | S2 | S3 | S4 |
|---|---|---|---|---|
| ÖSZ-1 (k€/MW → inner → pts) | 20 → 20 → 4.0 | 20 → 4.0 | 50 → 100 → 20 | 50 → 20 |
| ÖSZ-2 (€/MWh) | 2.0 → **0** | 2.0 → 0 | 2.0 → 0 | 3.0 → 15 |
| ÖSZ-3 overbuild | 30.5 % → 70 → 7.0 | 40 % → 90 → 9.0 | 40 % → 9.0 | 54 % → 100 → 10 |
| ÖSZ-4 BESS / installed wind | 34.5 % → 70 → 10.5 | 90 / 280 = 32.1 % → 70 → 10.5 | 32.1 % → 10.5 | 51 % → 100 → 15 |
| ÖSZ-5 / ÖSZ-8 | 0 / 0 | 5 / 5 | 5 / 5 | 5 / 5 |
| ÖSZ-6 (× minimum) | 1.5 → 50 → 2.5 | 1.6 → 60 → 3.0 | 2.0 → 100 → 5.0 | 2.1 → 5.0 |
| ÖSZ-7 | 0 | 0 | 0 | 0 |
| ÖSZ-9 EU turbine / EU transformer | 12 / 0 | 12 / 3 | 12 / 3 | 12 / 3 |
| **Total (EU turbine)** | **36.0** | **51.5** | **69.5** | **90.0** |

These totals match the model (CP rows 52/53). The S2 and S3 BESS of 90 MW is "45 % of MVA" in the model's input, but it is scored, correctly, as 32.1 % of installed wind. The model's Tender rows 44 and 66 use installed wind, as GYIK Q20 requires.

---

## 5. Selling the project or the right before COD

The call has **no explicit change-of-control clause and no assignment mechanism.** The relevant rules are:

1. **The right is tied to the SPV.**
   - The right is "…csatlakozást a Hivatal határozata alapján biztosító jog és kötelezettség együttese" (a bundle of right **and obligation** under MEKH's decision; §II.1 h, p.4).
   - The applicant must be a Hungarian project company "kizárólag e célra létrehozott" (set up exclusively for this purpose; §II.1 aa–bb, p.5).
   - The call has no procedure to move the decision to another entity. A sale therefore means a **share deal** in the SPV.
2. **No legal succession until the decision is final.** "A pályázat benyújtásának napjától a Hivatal benyújtott pályázatra vonatkozó döntésének véglegessé válásáig a Pályázó személyében jogutódlás nem következhet be." (From submission until MEKH's decision is final, no legal succession may occur in the person of the applicant; §XIII.13, p.31.) This rules out a merger, demerger or conversion of the SPV in that window. A plain share transfer is not *jogutódlás*, but see the duties in item 3.
3. **Duties to notify.**
   - Changes to the declarations in form chapters 9–14: within 8 days. A change of authorised representative: within 3 days (§IX.17, p.20).
   - A change in affiliated companies: within 8 days (Annex 12).
   - The applicant must meet the tender conditions throughout the procedure **and the full term of the right** (§VII.1 eh, p.12; GYIK Q78). A buyer group must therefore not bring in an exclusion ground under §VII.2 (see A22).
4. **The decision is frozen.** "…a jogosult kérelmére nem módosítható, kivéve a telephely módosításának lehetőségét" (cannot be modified at the holder's request, except to move the site; §XIII.12, p.31). Every points commitment passes to the buyer.
   - The site can be moved, even to another district, from receipt of the decision until commissioning. The new site must still meet §VII.1 b (GYIK Q1, Q8).
   - The decision's site list must be updated to the permitted sites **before** the MEKH generation/storage licence application (GYIK Q46).
5. **Securities stay with the applicant.**
   - Only the applicant may post them (GYIK Q67).
   - Guarantees cannot be transferred, assigned or pledged (Annex 5, point 10) and can be amended only with MEKH's prior written consent (point 11).
   - The form (cash or guarantee) cannot be changed (GYIK Q66).
   - A buyer who wants to replace the seller-sponsored guarantee needs MEKH's cooperation (a waiver of the old guarantee and acceptance of a new one on the template). The call has no explicit procedure for this.
6. **The 25 % community purchase option survives a sale** (§VII.18–VII.25, p.15–16; Annex 16).
   - Designated RECs may buy 25 % of the SPV within 2 years of the decision becoming final.
   - **Price** = the auditor-certified "vagyoni hozzájárulás" (capital contribution) paid in up to exercise, × the share acquired (§VII.25). The REC also takes on its pro-rata share of later capital contributions.
   - The winner cannot refuse claims up to 25 % in total. Several claims are scaled pro rata.
   - An attributable failure to implement the option costs the right and the PB (§VII.26).
7. **The right cannot be split** into several plants (GYIK Q53). If abandoned, the capacity is lost rather than passed to the runner-up (GYIK Q75).
8. **The PB-forfeiture exclusion follows the group.** A PB call excludes the SPV's whole company group from capacity tenders for 3 years (§VII.2 n; GYIK Q71). The text implies the group at the time of the call, which may be the buyer's.
9. **Group caps are tested at bid stage only.** They are 100 MVA per node and 200 MVA in total (§V.6–V.7). The call does not address what happens if a buyer later consolidates several awarded SPVs.

---

## 6. Formula-ready summary

| Parameter | Value | Unit | Basis MW | Timing | Source |
|---|---|---|---|---|---|
| Participation fee | 3,000,000 | HUF per bid | per bid | received by 30-Oct-2026 23:59; non-refundable except §VI.7 | Kiírás §VI.4–VI.7, p.9–10 |
| Documentation fee | 0 | HUF | – | – | §VI.3, p.9 |
| Benchmark capex – wind | 406 | mHUF/MW | installed wind MW | – | Annex 2, p.34 |
| Benchmark capex – BESS | 200 | mHUF/MWh | BESS MWh | – | Annex 2, p.34 |
| Bid security rate | 0.5 % | of `BI_bid` | **bid's** installed wind MW + BESS MWh | in place at submission; guarantee valid ≥ 30-Jun-2027 | §VIII.2, VIII.5; GYIK Q24, Q51 |
| Bid security, per unit | 2.03 / 1.00 | mHUF per MW wind / per MWh BESS | as above | as above | derived |
| Bid security release (lost or invalid bid) | 10 | wd | – | after the §XI.2 decision is sent (decision expected end-Jan 2027) | §VIII.13; GYIK Q65, Q68 |
| Bid security, winner | cash+cash: reclassified (pro rata if partial award); else released | – | – | 10 wd after the PB compliance letter | §XII.11 |
| Bid security forfeiture | 100 % (× f for a partial winner) | of bid security | – | if the PB is not posted properly | §XII.10 |
| PB minimum rate | 2.5 % | of `BI_awd` | **decision's** installed wind MW + BESS MWh (× f if partial award) | – | §XII.2; GYIK Q24, Q30 |
| PB, per unit (k = 1) | 10.15 / 5.00 | mHUF per MW wind / per MWh BESS | as above | – | derived |
| PB multiple `k` | ≥ 1.0 (ÖSZ-6 x / 100) | × | – | fixed in the bid | §XII.3, §X.4 bf |
| PB posting deadline | 15 | calendar days | – | from receipt of the decision (≈ mid-Feb-2027) | §XII.5; GYIK Q68 |
| PB guarantee minimum validity | COD deadline + 6 months | date | – | ≥ 31-Mar-2033 | §XII.8 |
| PB release | 10 | wd | – | after MEKH's post-COD compliance decision | §XII.13, §XIII.9 b |
| PB full call | 100 % | of PB | – | any §XIII.6 event; §VII.26; §XIII.17 (ÖSZ not proven by COD) | §XII.12, §XIII.17; GYIK Q72 |
| Capacity tolerance | 10 % | negative only | decision wind MW, and BESS MW separately | after full commissioning | §XIII.15; GYIK Q12, Q44, Q77 |
| Proportional PB loss | `s × PB` if `s > 10 %`, where `s = 1 − built/decided` | of PB (component or total, A10) | as above | at the COD check | §XII.12; §V.2 m; GYIK Q28, Q31 |
| Interest on cash deposits | none provided | – | – | – | silent; model 0 % |
| Connection-request security | per grid code (model: 4.5 mHUF/MW, cash – not verifiable from the call) | – | grid code | with the connection request, ≤ 15 days after the decision | §II.1 p, §XIII.6 a–b; GYIK Q40 |
| Credit of the connection-request security against the connection fee | 100 % | – | – | at fee payment | GYIK Q40 |
| Supplementary / capacity-reservation security | 0 | – | – | – | GYIK Q70 |
| Indirect connection fee, Szombathely OVIT | 73,071,306 | HUF/MVA (maximum) | **awarded feed-in MVA** | per HCSSZ | Annex 1; §V.3 c |
| Direct connection fee, Szombathely OVIT | 508,345,804 | HUF per connection point | **lump sum** (pro rata to MVA if shared) | per HCSSZ | Annex 1; §XIII.18 |
| Node capacity, Szombathely OVIT | 200 | MVA | feed-in | – | Annex 1 |
| Group cap per node | 100 | MVA submitted | requested MVA | at bid | §V.7, §VII.36; GYIK Q52 |
| Group cap in total | 200 | MVA awarded | requested MVA | at evaluation | §V.6, §IX.18 e |
| Minimum bid size | 14 MVA and 14 MW wind | – | – | also applies to a partial award | §V.2 d |
| Partial-award factor `f` | `MVA_awd / MVA_req` | – | scales MW_W, MW_B, MWh_B | at award | GYIK Q23 |
| Site rule | ≥ 1 ha, ≤ 19 km straight line | – | – | at bid; again for any site change | §VII.1 b; GYIK Q1, Q46 |
| Building permit (final) | 1-Jan-2030 | date | decision's rated MW | proof to MEKH | §V.2 l, §XIII.6 g; GYIK Q10, Q11 |
| Earliest connection | 30-Sep-2030 | date | – | – | Annex 1; GYIK Q60 |
| COD deadline | 30-Sep-2032 | date | – | COD = first parallel connection | §XIII.5, §XIII.10 |
| ÖSZ-1 one-off contribution | 10,000.01 – 50,000 (points) | EUR/MW | not defined (A7) | one-off; timing by later law | §X.4 ba, §V.2 f |
| ÖSZ-2 annual contribution | 2.01 – 3.00 (points) | EUR/MWh | energy the wind farm feeds in | annual | §X.4 bb, §V.2 g |
| Expected mandatory contribution | 10,000 €/MW + 2 €/MWh | – | – | future law | Kiírás p.8, footnote 1 |
| ÖSZ-3 x | `(MW_W / MVA_req − 1) × 100` | % | installed wind vs requested MVA | fixed in the bid | GYIK Q19 |
| ÖSZ-4 x | `MW_B / MW_W × 100` | % | BESS MW vs installed wind | fixed in the bid | GYIK Q20 |
| ÖSZ-4 duration | ≥ 4 (0.25C) | h | `MWh_B ≥ 4 × MW_B` | proof before BESS commissioning | Annex 8; GYIK Q5, Q6 |
| ÖSZ-6 x | `PB / PB_min × 100` | % | – | at bid | §X.4 bf |
| ÖSZ-9 threshold | ≥ 80 % of purchase cost, turbines and transformers separately | % | – | proof before COD | §X.4 bi |
| Weights ÖSZ-1…9 | 20/15/10/15/5/5/10/5/15 | % | – | – | Annex 3 |
| Score rounding | 2 | decimals | – | – | §X.1 b |
| Community purchase option | 25 % of SPV, at paid-in capital contribution | % | – | exercise ≤ 2 years after the decision is final | §VII.18–VII.25 |

**Worked figures for Bora (S1 ratios: overbuild 30.5 %, BESS 34.48 % of wind, 4 h, k = 1.5).** HUF amounts in mHUF; € at 365 HUF/EUR as in the IC model.

| Item | 200 MVA (model case; not compliant with §V.7) | 100 MVA (one group's maximum) |
|---|---|---|
| Wind MW / BESS MW / MWh | 261 / 90 / 360 | 130.5 / 45 / 180 |
| Benchmark investment BI | 177,966 | 88,983 |
| Bid security (0.5 %) | 889.8 (2.44 m€) | 444.9 (1.22 m€) |
| PB minimum (2.5 %) | 4,449.2 | 2,224.6 |
| PB at k = 1.5 | 6,673.7 (18.28 m€) | 3,336.9 (9.14 m€) |
| Indirect connection fee | 14,614.3 (40.04 m€) | 7,307.1 (20.02 m€) |
| Direct connection fee | 508.3 (1.39 m€) | 508.3 (1.39 m€), **not halved** |
| Participation fee | 3.0 per bid | 3.0 per bid |

---

## 7. Ambiguities and how the FAQ resolves them

| # | Issue | Call text | FAQ resolution / status | Modelling choice |
|---|---|---|---|---|
| A1 | ÖSZ-3 formula | "x % = x MVA / x MW₁" (inverted, a ratio) (§X.4 bc) | **Resolved:** `x = MW₁/MVA × 100 − 100`, with worked examples: 103 → 10 pts, 150 → 90, 160 → 100 (GYIK Q19) | use the GYIK formula |
| A2 | ÖSZ-4 numerator and denominator | MW₁/MW₂ labels swapped; basis "megpályázott kapacitás" (§X.4 bd) | **Resolved:** `x = MW₂ (BESS) / MW₁ (installed wind) × 100` (GYIK Q20) | BESS / installed wind |
| A3 | ÖSZ-4 minimum 5 % vs the 2 < x < 5 band | The text requires BESS ≥ 5 % of the bid capacity, but Annex 4 awards 10 points for 2 < x < 5 | **Unresolved.** The GYIK Q20 example scores 3 % → 10 points, which supports Annex 4. | not relevant for Bora (all ratios > 30 %); keep the conservative 5 % gate |
| A4 | ÖSZ-4 at exactly 50 % | Annex 4: 40 ≤ x < 50 → 90; 50 < x → 100 (50 itself missing) | **Resolved:** 40 ≤ x ≤ 50 → 90 (GYIK Q20 table) | 50 → 90 |
| A5 | ÖSZ-6 at exactly 200 % | 190 ≤ x < 200 → 90; 200 < x → 100 | **Resolved:** "200 ≤ X esetén 100 pont" (GYIK Q34) | 200 → 100 |
| A6 | Offers below the ÖSZ-1/2 minimums | §V.2 f–g name minimums | **Resolved:** 0 points, bid still valid (GYIK Q17), but note the expected law (footnote 1) | 0 points; still model the expected mandatory 10 k€/MW + 2 €/MWh |
| A7 | ÖSZ-1 MW basis, payment date, currency | "EUR/MW"; details "a későbbiekben megjelenő jogszabály" (in a later law) | **Unresolved** | the model uses net MVA. Sensitivity: installed MW, which is 30.5 % more cost at S1. |
| A8 | ÖSZ-1/2 above the maximum (> 50,000 €/MW, > 3 €/MWh) | The range is stated as "legfeljebb" (at most) | **Unresolved.** Literally outside the range, so no band applies. | cap inputs at the maximum |
| A9 | Tolerance on capacity-type commitments (ÖSZ-3, ÖSZ-4) vs §XIII.17 full call | §XIII.15: 10 % tolerance. §XIII.17: an unproven commitment means the PB is called. | **Partly resolved.** GYIK Q28 and Q44: ≤ 10 % short has no sanction, measured on the committed overbuild MW. GYIK Q45: the 70 % (brownfield) and 80 % (EU content) thresholds are themselves the tolerance; no extra tolerance. GYIK Q72: any partial failure means the PB is called. | base case: ≤ 10 % short → 0; > 10 % → proportional loss. Downside: full PB call. |
| A10 | What the proportional loss applies to | "a teljes hiányzó beépített teljesítményének arányában" (§XII.12) | **Unresolved:** whole PB or the failing technology's component (GYIK Q31 says "a teljesítési biztosítékát") | base: technology component. Conservative: total PB × shortfall. |
| A11 | BESS tolerance on MW while the PB base is MWh | §XIII.15 "beépített teljesítmény" | GYIK Q77: tested on **MW**, wind and BESS separately. An MWh shortfall would breach the Annex 8 4-hour rule (§XIII.17). | test MW; treat MWh < 4 × MW as a breach |
| A12 | Bid-security basis "igényelt" vs installed | §VIII.2: installed | **Resolved:** the bid's installed MW and MWh, overbuild included (GYIK Q24, Q51) | installed wind + BESS MWh |
| A13 | Partial award: points and PB | – | **Resolved:** installed values scaled by f, PB on the scaled values, points unchanged, tolerance on the scaled values (GYIK Q23, Q24, Q26, Q76) | as stated |
| A14 | When the 15-day clocks start | "határozat kézhezvétele" (receipt of the decision) | **Resolved:** from the decisions expected at the end of January 2027, not from the 13-Dec-2026 announcement (GYIK Q68) | PB and connection request ≈ mid-Feb 2027 |
| A15 | Scope of the building permit | "a beruházás megvalósításához szükséges végleges építési engedély" (final permit needed for the investment) | **Resolved:** every permit-requiring element of the wind farm and BESS (GYIK Q10), covering the decision's rated capacity and holder (GYIK Q11) | – |
| A16 | Staged commissioning and the tolerance test | – | **Resolved:** possible if the licensee agrees; tolerance tested after full commissioning (GYIK Q12) | – |
| A17 | Tie-break "four largest weights" (three-way tie at 15) | §X.3 | **Resolved:** ÖSZ-1, ÖSZ-2, ÖSZ-4, ÖSZ-9 (GYIK Q32) | – |
| A18 | Does the BESS get its own MVA? | – | **Resolved:** no; it uses the wind's headroom (GYIK Q18) | – |
| A19 | Force majeure for the permit or for law changes | §XIII.2: no exemption | **Resolved:** business risk (GYIK Q47) | – |
| A20 | Group cap per node | §V.7, §VII.36 | **Resolved:** confirmed at 100 MVA (GYIK Q33, Q43, Q52) | Bora ≤ 100 MVA per group |
| A21 | §IX.18 d "nem éri el" (strictly below the cap?) and whether smaller, lower-ranked bids can fill the gap after a rejected bid (§IX.18 g) | – | **Unresolved** | assume ≤ the cap; no gap-filling |
| A22 | Eligibility after a sale | §VII.1 eh: conditions must be met for the whole term | **Unresolved** for buyer-group exclusion grounds (GYIK Q78 defines only the term) | treat as a buyer-diligence item |
| A23 | Community purchase price | "vagyoni hozzájárulás" (capital contribution) (§VII.25) | **Unresolved:** whether shareholder loans count; the price is pro rata, so very low if the SPV is mainly loan-funded | price = 25 % × paid-in equity; flag the dilution risk |
| A24 | EU bank or EUR guarantee | §V.2 k allows it; §VIII.9 requires HUF | Allowed only with a counter-guarantee meeting §VIII.6–VIII.10 (§VIII.11). The FAQ adds nothing. | assume a HUF guarantee or counter-guarantee |
| A25 | Enforcing ÖSZ-2 after COD | Proof is due by COD (§XIII.17), yet the payments are annual over the plant's life | **Unresolved;** expected to be covered by the future law | model as an opex line for the plant's life |
| A26 | Connection fee instalments and dates; amount of the connection-request security | silent | **Unresolved:** governed by the HCSSZ and grid codes (GYIK Q40) | keep as inputs |

---

## 8. Where the Bora IC model differs from the call

Cells refer to `Bora_IC_model_v4.xlsx`; see `bora_models.md`.

1. **Requested capacity of 200 MVA breaks the 100 MVA per-node group cap** (§V.7, §VII.36; GYIK Q52).
   - The model only flags this (Tender I353, CP E65); it is not counted in the integrity checks.
   - Every security and fee scales with it. At 100 MVA the bid security, PB and indirect fee all halve; the direct fee does not.
   - Inputs F82 mentions "6 SPVs". This does not get round the cap within one group.
2. **Direct connection fee scaled per MW.**
   - The model has Inputs G84 = 508.35 mHUF / 200 = 2.54 mHUF/MW × Inputs E18 (Deposits row 12).
   - Annex 1 gives a lump sum **per connection point**. At 100 MVA the model would show 254 mHUF instead of 508 mHUF.
   - It should be scaled only when ÖSZ-8 sharing applies (§XIII.18).
3. **Bid security netted into the PB although both are guarantees.**
   - Inputs G107 = 1 sets Deposits I10 = −H8, yet Deposits P8 = 0 and P10 = 0 (both guarantees).
   - The call allows reclassification only when **both are cash** (§XII.11 a). Otherwise the bid guarantee is released and the PB must be issued for the **full** 6,674 mHUF, not 5,784. The PB guarantee becomes effective when the bid guarantee is waived (Annex 5 C/D, point 12).
   - Effect: the PB guarantee and the fees on it are understated by 890 mHUF (≈ 2.4 m€ of guarantee line) once the bid guarantee lapses.
4. **PB and connection-request dates are too early.**
   - The model uses 30-Dec-2026 (award + 17 days; Inputs G97, G102).
   - The 15-day clocks run from **receipt of the decision**, expected at the end of January 2027 (GYIK Q68). The dates should be ≈ mid-February 2027.
   - The bid-bond refund on failed capacity (Deposits row 9, dated 13-Dec-2026) should likewise fall 10 wd after the decision is sent.
5. **Connection-request security of 4.5 mHUF/MW, cash only** (Inputs G83/F83).
   - The call and FAQ give no amount, form or basis; they defer to the grid codes.
   - The credit against the connection fee (Deposits I18) **is** consistent with GYIK Q40. Keep the input, but cite the grid code, not the call.
6. **How long the PB is held.**
   - The model holds the PB for 6 months after COD (Tender E31).
   - The call's minimum guarantee validity runs to **6 months after the COD deadline**, i.e. ≥ 31-Mar-2033, whatever the actual COD (§XII.8). Release comes 10 wd after MEKH's post-COD compliance decision (§XII.13).
   - For an early COD, the guarantee tenor and fee may run longer than the model assumes.
7. **The PB forfeiture logic is simplified** (Inputs F238: "loss of … performance bond (pro rata on partial COD)").
   - The call has three regimes (§2.3): a full call for §XIII.6 events and §VII.26; a **full call with the right kept** for any unproven points commitment (§XIII.17; GYIK Q72); and a proportional loss only when the capacity shortfall exceeds 10 %, wind and BESS tested separately.
   - The model has no failure scenarios (`bora_models.md` §6).
8. **Community purchase price "unknown"** (Inputs F235).
   - The call fixes it at the **paid-in capital contribution, pro rata** (§VII.25), exercisable within 2 years of the decision becoming final. Tender section 9 ("at cost") agrees; the Inputs note is out of date.
9. **ÖSZ-2 at 2.00 €/MWh earns 0 points** in S1–S3. The model is correct (Tender E59 = 0), but this is an inefficient offer: 2.01 €/MWh would add 0.75 points.
   - Likewise, ÖSZ-1 at 20,000 €/MW scores 20 inner points (4.0 weighted); 20,000.01 would score 30 (6.0 weighted).
   - S1's 30.5 % overbuild and 34.5 % BESS ratio both sit inside the 30–35 % band, so 30.0 % would earn the same points.
10. **ÖSZ-4 minimum of 5 %.** The model makes 2 < x < 5 ineligible (Tender E166). The call text supports this, but Annex 4 and GYIK Q20 (3 % → 10 points) do not (A3). No effect for Bora.
11. **ÖSZ-1 MW basis.** The model charges 20 k€/MW on net MVA (Inputs G233), but the call leaves the basis open (A7). Run an installed-MW sensitivity.
12. **Rules the model applies correctly:**
    - band tables and weights, including the GYIK Q20 and Q34 fixes; 2-dp rounding;
    - ÖSZ-4 measured on installed wind (Tender row 44);
    - security bases on installed wind MW + BESS MWh (Inputs E91/E92; GYIK Q51);
    - rates of 0.5 % / 2.5 % × k; the 406 / 200 benchmarks;
    - participation fee of 3 mHUF per bid, non-refundable;
    - no supplementary security (Inputs G78 = 0; GYIK Q70);
    - indirect fee per MVA;
    - dates of 1-Jan-2030 (permit), 30-Sep-2030 (earliest connection) and 30-Sep-2032 (COD deadline);
    - the re-derived S0–S4 scores (§4.6).
