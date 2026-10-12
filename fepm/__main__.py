"""fepm CLI.

    python3 -m fepm build projects/<name>   run projects/<name>/build.py → projects/<name>/out/
    python3 -m fepm audit projects/<name>   recalc (LibreOffice) + lint + headline outputs + master check
"""
from __future__ import annotations

import argparse
import importlib.util
import sys
from pathlib import Path

from .audit.lint import lint, lint_values, report
from .recalc import name_value, recalc, values


def _load(project: Path):
    spec = importlib.util.spec_from_file_location(f"fepm_project_{project.name}", project / "build.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _model_file(project: Path) -> Path:
    out = project / "out"
    files = sorted(out.glob("*.xlsm")) or sorted(out.glob("*.xlsx"))
    files = [f for f in files if not f.stem.endswith("_recalc")]
    if not files:
        sys.exit(f"no workbook in {out}; run `python3 -m fepm build {project}` first")
    return files[0]


def build(project: Path):
    res = _load(project).build(project / "out")
    for k, v in (res or {}).items():
        print(f"{k}: {v}")


def audit(project: Path) -> int:
    model = _model_file(project)
    rc = model.with_name(model.stem + "_recalc.xlsx")
    recalc(model, rc)
    findings = lint(model) + lint_values(rc)
    print(f"model:  {model}\nrecalc: {rc} (LibreOffice)")
    print(report(findings))
    wb = values(rc)
    print("headline outputs (live case):")
    outs = [(next(wb.defined_names[n].destinations), n) for n in wb.defined_names if n.startswith("Out_")]
    for (sheet, cell), name in sorted(outs, key=lambda x: (x[0][0], int(x[0][1].split("$")[-1]))):
        r = int(cell.split("$")[-1])
        label, unit = wb[sheet].cell(r, 3).value, wb[sheet].cell(r, 4).value
        print(f"  {label:<36} {unit:<8} {name_value(wb, name)}")
    mc = name_value(wb, "Master_Check")
    print(f"master check: {mc} ({'OK' if mc == 0 else 'CHECK'})")
    return 1 if findings else 0


def main(argv=None):
    ap = argparse.ArgumentParser(prog="fepm")
    ap.add_argument("command", choices=["build", "audit"])
    ap.add_argument("project", type=Path)
    a = ap.parse_args(argv)
    if not (a.project / "build.py").exists():
        sys.exit(f"{a.project}/build.py not found")
    if a.command == "build":
        build(a.project)
        return 0
    return audit(a.project)


if __name__ == "__main__":
    sys.exit(main())
