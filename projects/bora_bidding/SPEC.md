# Bora – Bidding Scenarios: SPEC (DRAFT, awaiting deal details)

Status: skeleton. The sections marked **TBD** wait for the bidding details from the deal team.
Baseline facts come from `docs/reference_analysis/bora_models.md` (Bora_IC_model_v4).

## 1. Decision
TBD: the question the IC is asked to approve (bid level / strategy / walk-away price).

## 2. Baseline (from Bora IC model v4)
- 200 MVA net at the Szombathely OVIT node. The RtB premium is 150 k€/MW on net MVA (30 m€).
- Tender strategies S0–S4: S1 is the current bid. Scores (EU / CN): 0, 36/24, 51.5/36.5, 69.5/54.5, 90/75.
  Wind / BESS MW: 200/0, 261/90, 280/90, 280/90, 308/157.
- Routes: RtB sale, sale at COD (EU/CN turbine), hold to maturity (EU/CN).
- Key dates: RtB 30-Jun-2029, RtB sale 31-Mar-2030, COD 30-Sep-2031.

## 3. Known issues in v4 to fix, not copy
1. No probability of winning, so S0 looks best. Add a win-probability-weighted value.
2. The price effect is one-sided. Commitment costs go to the buyer, but the overbuild energy is never credited.
3. The buyer return is inconsistent between the RtB exit (11.1%) and the COD exit (7.2%). Use one buyer hurdle per route, set as an input.
4. Hidden FAILs (HTM G257, Deposits E48) and a lender downside min DSCR of 0.89x.
5. The node cap (200 vs 100 MVA) is a flag, not a check.
6. Run_Map F34/F35 are mislinked. The Dashboard figures are a stale snapshot.
7. FX is inconsistent (370 vs 365 HUF/EUR).

## 4. Scenarios: TBD
## 5. Model architecture: TBD (planned once §1 and §4 are fixed; must follow MODELLING_STANDARDS)
## 6. Outputs: TBD
## 7. Deck storyline: TBD (IC_DECK_STANDARDS storyline)
