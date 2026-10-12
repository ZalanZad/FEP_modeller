"""Macro packaging: embed vba/*.bas into an .xlsm with LibreOffice and run them (see README, "Macros").

LibreOffice 24.2 writes a real MS-OVBA vbaProject.bin (PROJECT, VBA/dir, VBA/_VBA_PROJECT, one compressed
stream per module) when its experimental VBA exporter is switched on (Calc/Filter/Import/VBA/UseExport), the
library container is in VBA compatibility mode and each module carries a ModuleInfo. The same VBA runs inside
LibreOffice, so the build can execute RunAllScenarios headlessly. Fallback for Excel: tools/inject_vba.ps1.
"""
from __future__ import annotations

import sys
from pathlib import Path

from ..recalc import Office

VBA_DIR = Path(__file__).resolve().parents[2] / "vba"


def modules(vba_dir: Path = VBA_DIR) -> dict[str, str]:
    return {p.stem: p.read_text() for p in sorted(vba_dir.glob("*.bas"))}


def _calls(run) -> list[tuple[str, str]]:
    """`("Module", "Sub")`, `"Module.Sub"` or a list of either → [(module, sub), ...] in run order."""
    if not run:
        return []
    if isinstance(run, tuple) and len(run) == 2 and all(isinstance(x, str) and "." not in x for x in run):
        return [run]
    out = []
    for item in [run] if isinstance(run, str) else run:
        module, sub = item.split(".") if isinstance(item, str) else item
        out.append((module, sub))
    return out


SOLVER_MODULES = ("Solver", "GoalSeek")     # embedded only when the workbook defines Solver_Topics


def package(xlsx, xlsm, run=("Scenarios", "RunAllScenarios"), vba_dir: Path = VBA_DIR) -> Path:
    """Embed vba_dir/*.bas, run the macros in order, recalc and save as .xlsm.

    Solver.bas and GoalSeek.bas are embedded only if the workbook has the `Solver_Topics` name (Inputs
    "Macro control", written when the model has a COPY/PASTE topic or a goal seek); Scenarios.RunSolvers
    calls them late-bound, so it runs with or without them.

    `run`: one ("Module", "Sub") / "Module.Sub", a list of them (e.g. ["BoraImpact.RunCommitmentImpact",
    "Scenarios.RunAllScenarios"]), or None."""
    calls = _calls(run)
    with Office(xlsx) as lo:
        mods = modules(vba_dir)
        if not lo.doc.NamedRanges.hasByName("Solver_Topics"):
            mods = {k: v for k, v in mods.items() if k not in SOLVER_MODULES}
        lo.add_vba_modules(mods)
        for module, sub in calls:
            lo.run_macro(module, sub)   # LibreOffice swallows VBA runtime errors: the checks are the safety net
        lo.calculate()
        mc = lo.get("Master_Check")
        if mc != 0:
            ran = ", ".join(f"{m}.{s}" for m, s in calls) or "no macro"
            print(f"warning: Master_Check = {mc} after {ran} – see the Checks sheet", file=sys.stderr)
        return lo.save(xlsm, macro=True)


def embedded_modules(xlsm) -> dict[str, str]:
    """Module name -> source extracted from the file's vbaProject.bin (needs oletools)."""
    from oletools.olevba import VBA_Parser
    vp = VBA_Parser(str(xlsm))
    try:
        return {Path(name).stem: code for _, _, name, code in vp.extract_macros()}
    finally:
        vp.close()
