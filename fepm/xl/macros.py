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


def package(xlsx, xlsm, run: tuple[str, str] | None = ("Scenarios", "RunAllScenarios"),
            vba_dir: Path = VBA_DIR) -> Path:
    """Embed the .bas modules, optionally run a macro (module, sub), recalc and save as .xlsm."""
    with Office(xlsx) as lo:
        lo.add_vba_modules(modules(vba_dir))
        if run:
            lo.run_macro(*run)   # LibreOffice swallows VBA runtime errors: the checks are the safety net
        lo.calculate()
        mc = lo.get("Master_Check")
        if mc != 0:
            print(f"warning: Master_Check = {mc} after {'.'.join(run or ())} – see the Checks sheet",
                  file=sys.stderr)
        return lo.save(xlsm, macro=True)


def embedded_modules(xlsm) -> dict[str, str]:
    """Module name -> source extracted from the file's vbaProject.bin (needs oletools)."""
    from oletools.olevba import VBA_Parser
    vp = VBA_Parser(str(xlsm))
    try:
        return {Path(name).stem: code for _, _, name, code in vp.extract_macros()}
    finally:
        vp.close()
