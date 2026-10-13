"""Bridge from the new Bana model back to the old engine (reference/Bana_IC_model_v3.xlsm, deck of 09-Oct-2026).

    python3 projects/bana/reconcile.py          → projects/bana/out/reconciliation.md

Each step puts one old-engine input back into a copy of the packaged model (base column of Inputs), re-runs the tax
solver in LibreOffice and records the headline outputs. What is left after the last step is a change of method that
the new structure does not reproduce (listed with the reason).
"""
import shutil
import tempfile
from pathlib import Path

import openpyxl

from fepm.recalc import Office

HERE = Path(__file__).resolve().parent
OUT = HERE / "out"
MODEL = OUT / "Bana_Model.xlsm"

OUTPUTS = [("Total uses to COD", "Uses_Total", "m"), ("FEP equity to COD", "Equity_COD", "m"),
           ("FEP profit – sale at COD", "COD_Profit", "m"), ("FEP IRR – sale at COD", "COD_IRR", "%"),
           ("FEP MOIC – sale at COD", "COD_MOIC", "x"), ("FEP IRR – hold to maturity", "HTM_IRR", "%"),
           ("FEP MOIC – hold to maturity", "HTM_MOIC", "x"), ("Minimum DSCR", "Min_DSCR", "x")]
# old engine, Bana deck 09-Oct-2026 slides 5-6 (Bana sheet D22:G52): European, Chinese
OLD = {"Uses_Total": (101.75, 75.15), "Equity_COD": (44.23, 40.32), "COD_Profit": (20.25, 29.35),
       "COD_IRR": (0.264, 0.377), "COD_MOIC": (1.46, 1.73), "HTM_IRR": (0.092, 0.121), "HTM_MOIC": (5.64, 7.09),
       "Min_DSCR": (1.09, 1.77)}
STEPS = [  # label, {input name: old value}, source
    ("EU turbine 950 k€/MW", {"In_Turbine": 950}, "Bana v3 Inputs E162; new: Bora user decision 600"),
    ("Debt share cap 60 % EU / 50 % CN", {"In_Gearing": 0.60, "In_Gearing_CN": 0.50}, "Bana v3 Inputs E217"),
    ("Tenor 10 years", {"In_Debt_Tenor": 10}, "Bana v3 Inputs E220"),
]
RESIDUAL = [
    "Revenue: the old hold case had the 90 €/MWh PPA for 10 years and Aurora Central merchant prices after "
    "(average 134.7 €/MWh over the life); the new model has the PPA for the whole life (user, 13-Oct-2026).",
    "Debt: old = gearing-sized annuity from the first full operating year with a 50 % cash sweep for 5 years; new = "
    "debt share of cost capped by the DSCR capacity, sculpted repayment, no sweep (Rock / Alpha).",
    "Timing: old purchase price and DD on 31-Jul-2026 (monthly grid); new price at closing 31-Dec-2026 after the "
    "award and DD at the bid date, quarterly grid.",
    "Tax: old 50 % loss cap with devex expensed and decommissioning at 30 k€/MW; new capitalises the purchase price, "
    "DD and devex into the 20-year depreciation base and has no decommissioning (immaterial at 8 % over 30 years).",
    "Opex: the new model adds the 0.3 % innovation contribution on the HIPA base.",
    "Overbuild: the old bid had none (50 MW); the new bid commits 20 % and builds 10 % (55 MW, user 13-Oct-2026), "
    "which adds capex, generation, the COD price on the extra MW and a larger security base.",
    "Bid and BESS: the old bid was municipal 35 k€/MW and community 2.5 €/MWh with no BESS; the new bid is 50 and 3 "
    "with a 10 MW / 4 h BESS at capex only (user, 13-Oct-2026).",
]


def fmt(v, unit):
    if isinstance(v, str):
        return v
    return {"m": f"{v:,.2f}", "%": f"{v * 100:.1f}%", "x": f"{v:.2f}x"}[unit]


def base_cells(path):
    """Inputs base-column address (H) of every input row, by name."""
    wb = openpyxl.load_workbook(path)
    out = {}
    for name in ("In_Turbine", "In_Gearing", "In_Gearing_CN", "In_Debt_Tenor"):
        row = wb.defined_names[name].attr_text.split("$")[-1]
        out[name] = f"Inputs!H{row}"
    return out


def run():
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp) / MODEL.name
        shutil.copy(MODEL, work)
        cells = base_cells(work)
        results = []
        with Office(work) as lo:
            lo.run_macro("Scenarios", "RunSolvers")
            results.append(("New model, base", {o: [lo.get(f"Out_{o}_{i}") for i in (1, 2)] for _, o, _ in OUTPUTS}))
            for label, changes, _ in STEPS:
                for name, value in changes.items():
                    lo.set(cells[name], value)
                lo.run_macro("Scenarios", "RunSolvers")
                results.append((f"+ {label}", {o: [lo.get(f"Out_{o}_{i}") for i in (1, 2)] for _, o, _ in OUTPUTS}))
    lines = ["# Bana – reconciliation to the old engine", "",
             "Old engine: `reference/Bana_IC_model_v3.xlsm`, figures as shown in `reference/Bana_IC_20260109_vDSz.pptx` "
             "(slides 5–6). New: `out/Bana_Model.xlsm`, base set. Each step below is cumulative: one old input put back "
             "into the new model, tax solver re-run (`reconcile.py`).", ""]
    for r, route in enumerate(("European turbine", "Chinese turbine")):
        lines += [f"## {route}", "", "| Step | " + " | ".join(lbl for lbl, _, _ in OUTPUTS) + " |",
                  "|---|" + "---:|" * len(OUTPUTS)]
        for label, vals in results:
            lines.append(f"| {label} | " + " | ".join(fmt(vals[o][r], u) for _, o, u in OUTPUTS) + " |")
        lines.append("| **Old engine (deck 09-Oct-2026)** | " +
                     " | ".join(fmt(OLD[o][r], u) for _, o, u in OUTPUTS) + " |")
        lines.append("")
    lines += ["## Step sources", ""] + [f"- {label}: {src}." for label, _, src in STEPS]
    lines += ["", "## What the steps do not reproduce (change of method, by design)", ""]
    lines += [f"- {t}" for t in RESIDUAL]
    lines += ["", "The turbine step only moves the European route. Tax losses still unused 5 years after the NTP lapse "
                  "in every step (Tao. 17. §), so the steps carry no tax shield the rule would not allow.", ""]
    (OUT / "reconciliation.md").write_text("\n".join(lines))
    return OUT / "reconciliation.md"


if __name__ == "__main__":
    print(run())
