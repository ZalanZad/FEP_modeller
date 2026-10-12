Build the Bora bid strategy package (BORA 1 / BORA 2, 2 × 100 MVA): the Excel model, the IC deck, and the cover email in EN and HU.

Everything is specified in `projects/bora_bidding/SPEC.md`. Read it first, together with `docs/MODELLING_STANDARDS.md`,
`docs/IC_DECK_STANDARDS.md`, `docs/WRITING_STANDARDS.md` and the domain notes in `docs/domain/`.
The primary source is `reference/Bora_scenarios_meeting.xlsx`:
- `Bora` sheet: column F = BORA 1 and column G = BORA 2. Rows 6–17 hold the commitments, L5:N21 the key assumptions,
  rows 48–61 the RtB waterfall and rows 73–89 the cost to COD.
- `Devex_FEP` / `BORA_Devex` sheets: the devex lines.
- The method for fees, deposits, interest and the waterfall comes from `reference/BORA_Devex_20261002_BG.xlsx`.

## What I want
1. **A lean Excel model**, built the way a human modeller would build one (like Project Rock / Alpha): a few logical sheets.
   - Sheets: `Dashboard` (the presentation), `Inputs`, `Curves`, `Calc_1` / `Calc_2` (identical, one per bid), `Checks`.
     No data tables, no pasted snapshots. Everything is a formula; anything that must be stored as a value is written by a macro.
   - The Dashboard has two columns, BORA 1 | BORA 2, and runs top to bottom:
     1. the tender commitments, entered right there, with the points per item and the impact per item
        (Δ HTM IRR, Δ RtB price PV, upfront cash or guarantee, points per €m), plus a key-assumptions panel next to it
        with a selectable assumption set;
     2. the points total and validity;
     3. the securities and fees, computed with the devex-file method and projected over time;
     4. the RtB exit per the devex waterfall, with an editable exit date;
     5. the sale at COD;
     6. hold to maturity.
   - The fees follow the devex file. The RtB premium is calculated on the connection capacity (awarded MVA, 100 per bid).
   - Reconcile against the saved values in SPEC §3, and explain every difference (SPEC §8 lists the deliberate fixes).
2. **An IC deck** (~8–10 slides) and **a cover email** (EN + HU) explaining why we chose these points:
   - **Municipal contribution:** many points, little effect on IRR, and no financial commitment from us now.
   - **Overbuild:** many points, little downside, because we will not build it. Show the PB sanction under both readings.
   - **BESS power:** fits well with the overbuild, not expensive, many points.
   - **Performance security above the minimum: no.** It ties up a lot of money up front for few points.
   - **All other commitments:** accepted, because they carry no financial burden.
   Write in proper investment language, as in an IC email or deck, not chat style. State the risks honestly
   (the FAQ Q72 full-PB-call conflict, the 3-year exclusion, the connection right lost pro rata, the 100 MVA cap reading for two bids).

## How to run it
1. Plan (you): confirm or refine SPEC.md. If anything is genuinely ambiguous, ask me before building; otherwise proceed.
2. `model-builder`: extract the data to `projects/bora_bidding/data/`, write `build.py` and `vba/BoraImpact.bas`, build,
   run `python3 -m fepm audit projects/bora_bidding` until lint is clean and Master_Check = 0, then write `out/reconciliation.md`.
3. `model-auditor`: independent check (Python recomputation of HTM IRR, RtB price PV and FEP RtB proceeds for both bids,
   checks, logic). Fix every finding through `model-builder`.
4. `ic-writer`: the email in EN and HU, plus the slide titles and text, with every figure cell-referenced.
   `deck-builder`: build the deck from the recalculated workbook.
5. `model-auditor`: tie the deck and the email out to the workbook.
6. Commit on the session branch, push, and send me: the .xlsm, the .pptx, both emails, and a short summary in Hungarian
   (headline numbers per bid, the checks run, open points).
