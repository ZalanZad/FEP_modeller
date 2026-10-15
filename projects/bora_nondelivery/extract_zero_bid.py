"""Zero-bid operating cash flows from the Bora model (SPEC §7) → data/zero_bid.csv.

Opens a private LibreOffice copy of projects/bora_bidding/out/Bora_Bid_Model.xlsm (never saved), sets every
lever of both bids to its zero level (BoraImpact.bas zeros; BESS hours kept), recalculates and reads Calc_1.
"""
import csv
import datetime as dt
from pathlib import Path

from fepm.recalc import Office

HERE = Path(__file__).resolve().parent
SRC = HERE.parent / "bora_bidding" / "out" / "Bora_Bid_Model.xlsm"
ZEROS = {"Municipal": 0, "Community": 0, "Overbuild": 0, "BESS": 0, "PB_Mult": 1, "Partial": 0,
         "Brownfield": 0, "Shared": 0, "EU_Turbine": 0, "EU_Transformer": 0}
OPEX = ["Variable opex (balancing, grid)", "Fixed opex and administration", "Land lease", "MEKH fee and HIPA",
        "Opex contingency", "Community contribution (ÖSZ-2)"]
EPOCH = dt.date(1899, 12, 30)


def main():
    with Office(SRC) as lo:
        for k, v in ZEROS.items():
            for col in "EF":                                  # both bids
                r = lo._range(f"Dash_{k}").getCellByPosition(0, 0).CellAddress.Row + 1
                lo.set(f"Dashboard!{col}{r}", v)
        lo.calculate()
        ws = lo.doc.Sheets.getByName("Calc_1")
        rows = {ws.getCellByPosition(2, r).getString(): r for r in range(0, 400)}
        val = lambda r, c: ws.getCellByPosition(c, r).getValue()  # noqa: E731
        cod = EPOCH + dt.timedelta(lo.get("In_COD"))
        base_row = rows["Tax depreciation (E: base)"]
        print("zero-bid tax-depreciation base (Calc_1 E%d): %.6f" % (base_row + 1, val(base_row, 4)))
        out = []
        for c in range(9, 200):
            end = ws.getCellByPosition(c, 3)
            if not end.getFormula():
                break
            d = EPOCH + dt.timedelta(end.getValue())
            if d <= cod:
                continue
            rev = val(rows["Revenue"], c)
            opex = sum(val(rows[k], c) for k in OPEX)
            out.append((d.isoformat(), rev, opex, val(rows["EBITDA"], c)))
        with open(HERE / "data" / "zero_bid.csv", "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(["period_end", "revenue", "opex", "ebitda", "tax_base", "rows"])
            for i, o in enumerate(out):
                w.writerow(list(o) + ([val(base_row, 4), "Calc_1 rows %d, %s, %d; base E%d" % (
                    rows["Revenue"] + 1, "-".join(str(rows[k] + 1) for k in (OPEX[0], OPEX[-1])),
                    rows["EBITDA"] + 1, base_row + 1)] if i == 0 else ["", ""]))
        print(len(out), "quarters", out[0][0], "to", out[-1][0])


if __name__ == "__main__":
    main()
