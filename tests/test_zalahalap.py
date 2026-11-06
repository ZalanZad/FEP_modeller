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


def test_gearing_zero_reproduces_the_unlevered_case():
    shadow = _load("zalahalap_shadow", PROJECT / "shadow.py")
    for case in shadow.CASES:
        r = shadow.run(case, "Base", gearing=0)
        assert r["debt_cod"] == 0
        for lev, unl in (("hold_irr_l", "hold_irr"), ("cod_irr_l", "cod_irr"), ("buyer_rtb_l", "buyer_rtb"),
                         ("buyer_cod_l", "buyer_cod"), ("cod_profit_l", "cod_profit")):
            assert r[lev] == pytest.approx(r[unl], abs=1e-12), (case, lev)


def test_excel_xirr_converges_from_the_workbook_guess():
    """Excel's XIRR (Newton) returns ~0 when it fails: from the workbook's start guess (MIN(10 %, MOIC − 1); buyer at
    RtB: MIN(target, MOIC − 1); buyer at COD: the target) every IRR of the shadow converges to the exact root, and
    where no root exists (the workbook shows n/a) the money multiple is below 1 (Base, Low, High; all cases, routes
    and lever flips)."""
    shadow = _load("zalahalap_shadow", PROJECT / "shadow.py")
    for st in ({}, dict(bess_price=0), dict(bess_income=3000), dict(fix_years=0)):
        for case in shadow.CASES:
            for sens in ("Base", "Low", "High"):
                res = shadow.run(case, sens, **st)
                for key, (flows, dates) in res["irr_flows"].items():
                    cap = res["buyer_start"] if key.startswith("buyer_rtb") else 0.1
                    start = res["buyer_start"] if key == "buyer_cod" else shadow.guess(flows, cap=cap)
                    exact = shadow.xirr(flows, dates, start=start)
                    if exact is None:
                        assert shadow.guess(flows, cap=1.0) < 0, (st, case, sens, key)
                        continue
                    got = shadow.xirr_excel(flows, dates, start)
                    assert got == pytest.approx(exact, abs=1e-7), (st, case, sens, key)
