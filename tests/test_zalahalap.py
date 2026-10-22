"""Zalahaláp model: fast checks without LibreOffice (static lint, the v3 shadow, the stored reconciliation targets)."""
import csv
import importlib.util
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
PROJECT = ROOT / "projects" / "zalahalap"


def _load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="module")
def build():
    return _load("zalahalap_build", PROJECT / "build.py")


def test_static_lint_clean(build, tmp_path):
    from fepm.audit.lint import lint
    m = build.model()
    build._fix_set_formulas(m)
    path = m.save(tmp_path / "z.xlsx")
    assert m.wb.sheetnames == ["Dashboard", "Inputs", "Curves", "Calc_1", "Calc_2", "Calc_3", "Checks"]
    assert len(m.scenarios) == 18
    assert lint(path) == []


def test_v3_shadow_ties_to_every_target():
    sys.path.insert(0, str(PROJECT / "data" / "old_engine"))
    v3 = _load("old_zhlp", PROJECT / "data" / "old_engine" / "old_zhlp.py")
    res, tg = v3.run(), v3.targets()
    bad = {k: (t, res.get(k)) for k, t in tg.items() if abs(res.get(k, 1e99) - t) > 1e-6 * max(1, abs(t))}
    assert not bad


def test_stored_reconciliation_targets_match_the_shadow():
    rec = _load("zalahalap_reconcile", PROJECT / "reconcile.py")
    sh = rec.run(rec.params(**rec.SET2))
    stored = {r["key"]: float(r["shadow"]) for r in csv.DictReader(open(PROJECT / "data" / "recon_targets.csv"))}
    assert set(stored) == set(rec.RECON_KEYS)
    for k, x in stored.items():
        assert sh[k] == pytest.approx(x, rel=1e-9, abs=1e-9)
