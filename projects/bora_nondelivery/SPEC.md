# Bora – non-delivery case (separate workbook): SPEC v1

Owner: orchestrator. Builder: `model-builder`. Requested by the user on 15-Oct-2026.

## 1. Question
Suppose we bid for maximum points except a performance security above the minimum, and then deliver none of the commitments at COD.
How much do we save at COD, item by item, against the performance security (PB) that MEKH then calls? This is the bid strategy of
BORA 1 and BORA 2 (85 points each).

What we save:
- the overbuild, which is not built;
- the BESS, which is not built;
- the municipal one-off and the 3 €/MWh community contribution, which are not paid;
- the switch from EU turbines and transformers to Chinese ones.

Scope: both bids together (2 × 100 MVA). Per-bid figures are shown only where they help (one column).

## 2. Workbook (`out/Bora_NonDelivery.xlsx`, standard layout, lean)
Sheets: Summary, Inputs, Calc (annual timeline 2031–2061 for the community stream), Checks. No Cover text blocks.
Summary ends with a short "Sanction" note block (≤ 6 lines, with citations; explanation allowed here because the user asked for it).

## 3. Inputs (values from `projects/bora_bidding` model Inputs / Dashboard; source notes cite them)
- **Per bid:** MVA 100, overbuild 51 %, BESS 51 % of MVA × 4 h, built wind if delivered 151 MW, built wind if not delivered 100 MW.
- **Overbuild capex per MW** (old engine E163–E170 as split in the Bora model): turbine 600, installation 80, foundations 90, BoP 130,
  owner's 40, land/easements 15 k€/MW, plus contingency 5 % on these.
- **BESS capex:** 240 k€/MW (user).
- **Municipal one-off:** 50 k€/MW × MVA, paid at COD.
- **Community contribution:** 3 €/MWh on net generation of the built wind, over 30 years from COD 30-Sep-2031.
  Generation = built MW × 2,460 h × (1 − 2 % curtailment) × (1 − 0.3 %)^(year − 1), with COD-year and end-year fractions as in the Bora model.
  The PV is at 8 % to COD, pre-tax, with a CIT switch (default off).
- **EU to Chinese equipment ("white labelling"):** 100 k€ per built MW for turbines and transformers together (user, 15-Oct-2026).
  It is shown as its own line, separate from the overbuild and BESS capex items. It applies to the 100 MW that is built.
- **PB:** 2.5 % × (committed wind MW × 406 mHUF + BESS MWh × 200 mHUF) at FX 365 = the Bora model's PB per bid.
  - Full call: the whole PB. Unmet points commitments trigger it (Kiírás XIII.17, GYIK Q72).
  - Memo: the pro-rata reading (XII.12).

## 4. Outputs (Summary, both bids, € m, plus k€ per committed MW where useful)
- Savings at COD by item:
  - overbuild capex, by element (turbine, installation, foundations, BoP, owner's, land, contingency);
  - BESS capex;
  - municipal one-off;
  - community contribution (PV at 8 %);
  - EU → CN white labelling;
  - total.
- PB called (full call), plus the pro-rata memo.
- Net saving = total savings − PB called. Also: the savings excluding the community PV, and the savings ÷ PB multiple.
- Memo:
  - points lost after the fact: none, because the award stands;
  - 3-year group exclusion: yes, because the PB is called;
  - connection right: kept under a full call (GYIK Q72), lost pro rata under the capacity rule (GYIK 28/30) for the unbuilt wind.

## 5. Checks
- Overbuild capex by element = total.
- PB = the Bora model PB per bid (state the value).
- Community PV recomputed with XNPV = SUMPRODUCT.

## 6. Sanction note (Summary)
- Full PB call when points commitments are not proven by COD (Kiírás XIII.17; GYIK Q72, no discretion).
- A capacity shortfall above 10 % → PB and connection right lost pro rata (XII.12; GYIK 28/30/44).
- A called PB excludes the whole company group from capacity tenders for 3 years (Kiírás VII.2 n; GYIK Q71). It applies to the group at the
  time of the call, which may be the buyer's group after a sale.

## 7. COD buyer's project IRR (user, 15-Oct-2026)
Question: what project IRR does a buyer at COD earn when the developer delivers nothing? Compare it with the original project, i.e. the
zero bid (all commitments at their zero level).
- **Buyer's project IRR:** unlevered and post-tax, XIRR on: − purchase price at COD + the project's operating cash flows after COD.
  - Operating cash flow = EBITDA − CIT × MAX(0, EBITDA − tax depreciation).
  - In a share deal the SPV keeps its own asset tax base: capex incl. connection fees and municipal, over 20 years. The purchase price
    does not enter it.
- **Purchase price:** the same in both cases, 2,450 k€/MW × 100 MW built per bid (Bora model `In_COD_Price`). No price effect.
- **Original (zero bid):**
  - Take the zero-bid project's quarterly revenue, opex and EBITDA after COD from the Bora model. Use a private LibreOffice copy with
    every lever at its zero level (BESS hours kept), Calc_1 rows.
  - Enter them as a sourced data block on the Curves sheet (source "Bora model, zero bid, Calc_1 row …").
  - Use the zero bid's tax-depreciation base from the same run.
- **Non-delivery:** the same operating cash flows. There is no community or municipal payment, no BESS and no overbuild, as in the zero bid. Then:
  - the full PB called at the sanction date (COD + 6 months), borne by the buyer's SPV;
  - a tax-depreciation base lower by the white-labelling saving (100 k€/MW × 100 MW).
  - Memo: buyer IRR if the seller bears the PB call.
- **Output:** buyer project IRR for original and non-delivery, and the difference in pp. Show it per bid (= both bids, since the IRR
  is the same) and the NPV at 8 % to COD.
