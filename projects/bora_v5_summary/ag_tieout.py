"""SPEC §20.5: with the project selector on Bora the model reproduces the user's file it was built from.

Builds build.py (source reference/Bora_model_RTB_COD_HTM_271026.xlsm + the §20 edits) into out/ag_tieout/, then
recalculates both that build (Dash_Project = Bora) and the user's file in LibreOffice for each case (cases 1–3, set through
Scenario_Name) and compares, per case: seven headline metrics, the seven section-5b RtB results and every Bid_Points
value (tolerance 1e-6, relative to max(1, |value|); text must match). The user's file is only read (LibreOffice works
on a private copy).

    PYTHONPATH=. python3 projects/bora_v5_summary/ag_tieout.py
"""
import importlib.util
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))
from fepm.recalc import Office  # noqa: E402

spec = importlib.util.spec_from_file_location("bora_v5_summary_build", HERE / "build.py")
build = importlib.util.module_from_spec(spec)
spec.loader.exec_module(build)

TOL = 1e-6
METRICS = [("Project IRR, unlevered post-tax", "Out_proj_irr"),
           ("Equity IRR, levered post-tax (FEP)", "Out_eq_irr"),
           ("FEP NPV at 8 %, € m", "Out_fep_npv"),
           ("Minimum DSCR, whole life", "Out_dscr"),
           ("Debt at COD, € m", "Out_debt"),
           ("COD sale price, € m", "Out_cod_price"),
           ("Buyer IRR at the RtB premium", "Out_buyer_irr")]
RTB_5B = [("5b RtB premium, k€/MVA", "Out_rtb_prem"), ("5b FEP profit post-tax, € m", "Out_rtb_profit"),
          ("5b FEP IRR post-tax", "Out_rtb_irr"), ("5b FEP NPV at 8 %, € m", "Out_rtb_npv"),
          ("5b FEP MOIC", "Out_rtb_moic"), ("5b FEP peak exposure, € m", "Out_rtb_peak"),
          ("5b A.RES profit post-tax, € m", "Out_ares")]
GRID = "A1:T260"                                                      # Bid_Points, every value


def by_case(path):
    """{case name: ({name: value}, Bid_Points values)} for every case, recalculated in LibreOffice."""
    out = {}
    with Office(path) as lo:
        if lo.doc.NamedRanges.hasByName("Dash_Project"):
            lo.set("Dash_Project", "Bora")
        cases = [lo.get(f"Inputs!{c}7") for c in "GHI"]
        for case in cases:
            lo.set("Scenario_Name", case)
            lo.calculate()
            vals = {n: lo.get(n) for _, n in METRICS + RTB_5B}
            vals["Master_Check"] = lo.get("Master_Check")
            out[case] = (vals, lo.doc.Sheets.getByName("Bid_Points").getCellRangeByName(GRID).getDataArray())
    return out


def diff(a, b):
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) / max(1.0, abs(a))
    return 0.0 if a == b else float("inf")


def main():
    paths = build.build(HERE / "out" / "ag_tieout", name="Bora_model_RTB_COD_HTM_ag_tieout")
    ag, ours = by_case(build.SRC), by_case(paths["xlsm"])
    ok, worst, n = True, 0.0, 0
    for i, case in enumerate(ag, 1):
        (a_vals, a_grid), (b_vals, b_grid) = ag[case], ours[case]
        print(f"Case {i}: {case}  (Master_Check user {a_vals['Master_Check']}, rebuilt {b_vals['Master_Check']})")
        for label, name in METRICS + RTB_5B:
            d = diff(a_vals[name], b_vals[name])
            worst, n, ok = max(worst, d), n + 1, ok and d <= TOL
            print(f"  {label:<36} user {a_vals[name]:>14.8g}  rebuilt {b_vals[name]:>14.8g}  diff {d:.1e}"
                  f"{'' if d <= TOL else '  FAIL'}")
        bad = [(r + 1, c + 1, x, y) for r, (ra, rb) in enumerate(zip(a_grid, b_grid))
               for c, (x, y) in enumerate(zip(ra, rb)) if diff(x, y) > TOL]
        cells = sum(1 for row in a_grid for x in row if x != "")
        n += cells
        ok &= not bad and b_vals["Master_Check"] == 0
        print(f"  Bid_Points: {cells} values, {len(bad)} differ{'' if not bad else ' – FAIL ' + str(bad[:5])}")
    print(f"ag tie-out: {'PASS' if ok else 'FAIL'} ({n} values, largest relative difference {worst:.1e})")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
