import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))


def _load_demo():
    import importlib.util
    spec = importlib.util.spec_from_file_location("demo_build", ROOT / "projects" / "_demo" / "build.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


@pytest.fixture(scope="session")
def demo(tmp_path_factory):
    """Build the demo once (model + macros run in LibreOffice + deck) and recalculate it."""
    from fepm.recalc import recalc
    out = tmp_path_factory.mktemp("demo")
    res = _load_demo().build(out)
    res["recalc"] = recalc(res["model"], out / "recalc.xlsx")
    res["out"] = out
    return res
