"""COPY/PASTE circularity solved by Solver.bas inside LibreOffice, and the result embedded in an .xlsm."""
import datetime as dt

from fepm.audit.lint import lint, lint_values
from fepm.recalc import name_value, recalc, values
from fepm.xl.builder import Model
from fepm.xl.macros import embedded_modules, package


def test_solver_closes_circularity(tmp_path):
    m = Model("Circ", periods=4, start=dt.date(2026, 1, 1))
    base = m.inputs.add_input("Base", "Base amount", "EUR k", 100, source="test")
    rate = m.inputs.add_input("Fee_Rate", "Fee rate on gross amount", "%", 0.2, source="test")
    c = m.calc
    c.section("Fee (circular: fee = rate x (base + fee))")
    c.paste("Fee", "Fee", "EUR k")
    c.row("gross", "Gross amount", "EUR k", f"={base}+{{fee_paste}}")
    c.copy("Fee", "Fee", "EUR k", f"={{gross}}*{rate}")
    xlsx = m.save(tmp_path / "circ.xlsx")
    assert lint(xlsx) == []
    xlsm = package(xlsx, tmp_path / "circ.xlsm", run=("Scenarios", "RunAllScenarios"))
    assert "Solver" in embedded_modules(xlsm)
    rc = recalc(xlsm, tmp_path / "circ_recalc.xlsx")
    assert lint_values(rc) == []
    wb = values(rc)
    paste = wb["Calc"].cell(m.calc.rows["fee_paste"].r, 10).value
    assert abs(paste - 100 * 0.2 / (1 - 0.2)) < 1e-6          # fee = 25
    assert name_value(wb, "Fee_Delta") < 1e-6


def test_no_topics_packages_without_solver_modules(tmp_path):
    """A model without COPY/PASTE or goal seek has no Macro control block, no Solver/GoalSeek modules, and
    RunAllScenarios (late-bound RunSolvers) still stores every scenario."""
    m = Model("Plain", periods=4, start=dt.date(2026, 1, 1), scenarios=["Base", "High"])
    base = m.inputs.add_input("Base", "Base amount", "EUR k", 100, {2: 200}, source="test")
    m.calc.section("Value")
    m.calc.row("v", "Value", "EUR k", f"={base}*{{yf}}")
    m.summary.add_output("V", "Value", "EUR k", m.calc.ref("v"))
    xlsx = m.save(tmp_path / "plain.xlsx")
    assert "Solver_Topics" not in m.names and "GoalSeek_Topics" not in m.names
    xlsm = package(xlsx, tmp_path / "plain.xlsm")
    assert {"Solver", "GoalSeek"}.isdisjoint(embedded_modules(xlsm)) and "Scenarios" in embedded_modules(xlsm)
    rc = recalc(xlsm, tmp_path / "plain_recalc.xlsx")
    assert lint_values(rc) == []
    wb = values(rc)
    stored = [wb["Summary"].cell(6, c).value for c in (7, 8)]
    assert abs(stored[1] / stored[0] - 2) < 1e-9
