"""SPEC §17.4: with the §17.1–17.2 levers set back to the AG values, the model reproduces the AG file.

Builds build.py with ag=True (overbuild 11 %, BESS 0, municipal 0, community 0 and the 5 % bottleneck rule in place of
the clipping curve) into out/ag_tieout/, then recalculates both that build and the user's AG file in LibreOffice for each
case (cases 1–3, set through Scenario_Name) and compares seven headline metrics (tolerance 1e-6, relative to
max(1, |value|)). The AG file is only read (LibreOffice works on a private copy).

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


def by_case(path):
    """{case name: {name: value}} for every case in Scenario_Names, recalculated in LibreOffice."""
    out = {}
    with Office(path) as lo:
        cases = [lo.get(f"Inputs!{c}7") for c in "GHI"]
        for case in cases:
            lo.set("Scenario_Name", case)
            lo.calculate()
            out[case] = {n: lo.get(n) for _, n in METRICS}
            out[case]["Master_Check"] = lo.get("Master_Check")
    return out


def main():
    paths = build.build(HERE / "out" / "ag_tieout", ag=True, name="Bora_model_RTB_COD_HTM_ag_tieout")
    ag, ours = by_case(build.SRC), by_case(paths["xlsm"])
    ok, worst = True, 0.0
    for i, case in enumerate(ag, 1):
        print(f"Case {i}: {case}  (Master_Check AG {ag[case]['Master_Check']}, rebuilt {ours[case]['Master_Check']})")
        for label, n in METRICS:
            a, b = ag[case][n], ours[case][n]
            d = abs(a - b) / max(1.0, abs(a))
            worst = max(worst, d)
            ok &= d <= TOL
            print(f"  {label:<36} AG {a:>14.8g}  rebuilt {b:>14.8g}  diff {d:.1e}{'' if d <= TOL else '  FAIL'}")
        ok &= ours[case]["Master_Check"] == 0
    print(f"ag tie-out: {'PASS' if ok else 'FAIL'} (21 values, largest relative difference {worst:.1e})")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
