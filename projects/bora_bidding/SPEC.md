# Bora – Bidding Scenarios: SPEC (DRAFT)

Status: the domain facts are in place. The decision, the scenario list and the open deal-team decisions (§6) are still needed
before the architecture is frozen.

Sources (read these, not the raw files):
- `docs/domain/tender_2026_rules.md`: the 2026 capacity tender (call + FAQ): deposits, fees, scoring, sanctions.
- `docs/domain/bora_devex_and_overbuild.md`: the devex lines, the FEP/ALSI waterfall and the overbuild penalty.
- `docs/reference_analysis/bora_models.md`: the Bora IC model v4 (routes, strategies, known issues).

## 1. Decision
TBD: the question the IC is asked to approve (bid configuration per strategy / walk-away level).

## 2. Hard constraints from the tender (formula inputs, each with a § citation in the model)
- **Max 100 MVA per company group per node** (§V.7, §VII.36, GYIK Q33/43/52). The v4 assumption of 200 MVA at
  Szombathely OVIT is invalid and must be a check in the model, not a flag.
- Wind installed ≥ requested MVA; BESS MW ≤ wind MW; BESS ≥ 4 h; ≥ 14 MVA.
- Score = ROUND(Σ weight × inner points / 100, 2) over ÖSZ-1…9. Use the band tables in tender_2026_rules.md §4.
  Bands are lookup tables on `Curves`, not nested IFs.

## 3. Cash and security items (timing and formula basis)
| Item | Payee | Basis | Timing | Return / release |
|---|---|---|---|---|
| Participation fee | MEKH | 3 mHUF per bid | 30-Oct-2026 | sunk |
| Bid security | MEKH | 0.5% × (wind MW × 406 + BESS MWh × 200) mHUF | at bid | released after award / converted |
| Performance security (PB) | MEKH | 2.5% × same basis × k (ÖSZ-6) | award + 15 days (~mid-Feb 2027) | to ≥ 31-Mar-2033; released after the COD compliance decision |
| Connection-request security | DSO/TSO | grid code (TBD) | award + 15 days | credited against the connection fee |
| Indirect connection fee | TSO | 73,071,306 HUF / awarded MVA | per connection contract | cost |
| Direct connection fee | TSO | 508,345,804 HUF per connection point (lump sum) | per connection contract | cost |

## 4. PB sanctions = overbuild scenario mechanics
- Shortfall vs. committed installed MW ≤ 10%: no sanction. Above 10%: shortfall share × PB, with wind and BESS tested separately
  (§XII.12, GYIK Q28/30/44/77).
- An unproven points commitment means a **full PB call** (§XIII.17, GYIK Q72), plus a 3-year group exclusion (GYIK Q71).
  The FAQ conflicts with the 10% rule. The model shows both readings as a switch until legal confirms which applies.

## 5. FEP / ALSI waterfall (from BORA_Devex_20261002_BG.xlsx; party names to be confirmed)
- FEP funds 100% of devex, guarantees, deposits and connection fees.
- Interest rates by line type: 12% on cash, 6% on guarantee face value, 20% on devex/land with a 2× floor.
  Compounded annually on 31-Dec, act/365, until the interest freeze date (switch).
- Proposed waterfall on the premium, each step capped at the cash available:
  1. sale costs
  2. FEP recoverable principal
  3. FEP return
  4. ALSI success fee
  5. remainder split 70/30
- The buyer reimburses land, bonds and connection fees separately (memo line).

## 6. Open decisions (deal team)
1. Requested MVA under the 100 MVA cap; premium and connection-fee MW basis (MVA vs installed MW).
2. FX: one input (the files use 370 and 365).
3. The interest freeze; the rule for a negative remainder; whether devex principal is repaid; the success-fee cap; FEP's own costs (6% guarantee fee, DD).
4. The overbuild sanction regime (10% pro rata vs full call) and whether the connection right is lost pro rata.
5. Confirm the parties: in the file, "FGX" (70%) = FEP? "Ventus+Solmentore" (30%) = ALSI?
6. The connection-request security amount (grid code).

## 7. Known v4 issues to fix, not copy
- No win probability.
- One-sided price effect.
- Buyer returns inconsistent between exits.
- Hidden FAILs.
- Direct connection fee scaled per MW.
- Bid security netted into the PB although both are guarantees (−890 mHUF).
- PB dated 30-Dec-2026 and held only COD + 6 months.
- S1 offers sit just below band edges.

## 8. Scenarios: TBD
## 9. Model architecture: TBD (after §1, §6 and §8; per MODELLING_STANDARDS)
## 10. Outputs: TBD
## 11. Deck storyline: TBD (per IC_DECK_STANDARDS)
