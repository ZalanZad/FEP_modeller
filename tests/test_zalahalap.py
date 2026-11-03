"""Zalahaláp model (SPEC v2): static lint of the build, the v3 shadow, the stored v1 targets and, with LibreOffice,
the tie-out of the independent shadow to the recalculated workbook."""
import csv
import importlib.util
import shutil
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
PROJECT = ROOT / "projects" / "zalahalap"


def _load(name, path):
    sys.path.insert(0, str(path.parent))
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    build = _load("zalahalap_build", PROJECT / "build.py")
    out = tmp_path_factory.mktemp("zhlp")
    return build, build.build(out)["model"]


def test_static_lint_clean(built):
    import openpyxl
    from fepm.audit.lint import lint
    _, path = built
    assert path.suffix == ".xlsx"
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames == ["Dashboard", "Inputs", "Revenue", "Calc", "Checks"]
    assert wb["Dashboard"]["H3"].value == "=Master_Check"
    assert lint(path) == []


def test_v3_shadow_ties_to_every_target():
    sys.path.insert(0, str(PROJECT / "data" / "old_engine"))
    v3 = _load("old_zhlp", PROJECT / "data" / "old_engine" / "old_zhlp.py")
    res, tg = v3.run(), v3.targets()
    bad = {k: (t, res.get(k)) for k, t in tg.items() if abs(res.get(k, 1e99) - t) > 1e-6 * max(1, abs(t))}
    assert not bad


def test_stored_v1_targets_match_the_v1_shadow():
    rec = _load("zalahalap_reconcile", PROJECT / "reconcile.py")
    sh = rec.run(rec.params(**rec.SET2))
    stored = {r["key"]: float(r["shadow"]) for r in csv.DictReader(open(PROJECT / "data" / "recon_targets.csv"))}
    assert set(stored) == set(rec.RECON_KEYS)
    for k, x in stored.items():
        assert sh[k] == pytest.approx(x, rel=1e-9, abs=1e-9)


@pytest.mark.skipif(shutil.which("soffice") is None, reason="LibreOffice not installed")
def test_shadow_ties_to_the_workbook(built, tmp_path):
    import openpyxl
    from fepm.recalc import recalc
    _, path = built
    rec = _load("zalahalap_reconcile", PROJECT / "reconcile.py")
    shadow = _load("zalahalap_shadow", PROJECT / "shadow.py")
    rec.XLSX = path
    rc = recalc(path, tmp_path / "rc.xlsx")
    ws = openpyxl.load_workbook(rc, data_only=True)["Dashboard"]
    assert ws["H3"].value == 0
    for i, case in enumerate(rec.CASES):
        sh = shadow.run(case, "Base")
        for key, (r, scale) in rec.dashboard_cells().items():
            assert ws.cell(r, 5 + i).value == pytest.approx(sh[key] / scale, rel=1e-9, abs=1e-7), (case, key)
