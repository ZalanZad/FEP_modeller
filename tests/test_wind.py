"""Wind-format builder (fepm/xl/wind.py): grid, levers, Block, DatedBlock, lint exceptions, annual Model dates."""
import datetime as dt

import openpyxl
import pytest

from fepm.audit.lint import lint
from fepm.xl.builder import Model
from fepm.xl.wind import Block, Book, DatedBlock

CASES = ["A", "B", "C"]


def _book(tmp_path):
    b = Book(CASES, years=6)
    d = b.sheet("Dashboard", "Tiny wind book")
    d.bar("Key inputs: the cases", e="cases")
    d.lever("MW", "Capacity", "MW", [10, 20, 30], source="test")
    d.lever("Basis", "Revenue basis", "text", ["Merchant", "Tolling", "Merchant"], options=["Merchant", "Tolling"])
    d.blank()
    d.bar("Sensitivity case", e=["Base", "Low", "High"], h="Live", key="case_bar")
    r = d.row("sel", "Case selected", "case", h="Base", inputs=("H",))
    d.validate(f"H{r}", "Base,Low,High")
    b.define("Dash_Case", f"Dashboard!$H${r}")
    cb = d.row_of("case_bar")
    d.row("px", "Price change", "%", e=0, f=-0.2, g=0.2, inputs=("E", "F", "G"),
          h=f"=INDEX({{self.cases}},MATCH(Dash_Case,$E${cb}:$G${cb},0))")
    b.define("Dash_Case_Price", d.ref("px", "cases"))
    b.define("Case_Price", d.ref("px", "H"))
    i = b.sheet("Inputs", "Inputs")
    i.bar("Inputs", e="cases", h="Source")
    i.input("First_Year", "First year", "year", 2025, "test")
    i.input("Rate", "Return", "%", 0.12, "test")
    i.input("Days_In_Year", "Days", "days", 365, "test")
    i.input("Pay_Date", "Payment", "date", dt.date(2025, 11, 30), "test")
    i.input("Exit_Date", "Exit", "date", dt.date(2027, 3, 31), "test")
    c = b.sheet("Calc", "Calculations", timeline=True)

    def rows(n, blk):
        c.row("cap", "Capacity (E)", "MW", e="=" + blk.lever("Dash_MW"))
        c.row("cf", "Cash flow", "EUR k", t="={cap.E}*({year}-First_Year-1)*(1+Case_Price)", h="=SUM({cf.all})")
        c.row("irr", "IRR (E)", "%", e="=XIRR({cf.all},$J$5:$" + b.last_col + "$5)")
    blocks = Block(c, 3).build(lambda n: f"Case {n}", rows)
    c.end()
    db = DatedBlock(d, "Rate")
    d.blank()
    db.header("Dated flows")
    db.end("x", "Exit", "=Exit_Date", "L")
    first = db.line("pay", "Payment", "=-{c}$" + str(d.row_of("MW")), "=Pay_Date", "=1")
    db.line("rec", "Receipt", "={c}$" + str(d.row_of("MW")) + "*(1+1)", "=Exit_Date", "=0")
    d.row("irr", "IRR", "%", cases=f"=XIRR({{c}}{first}:{{c}}{first + 1},$J{first}:$J{first + 1})")
    db.claim("ret", "Return due at the exit", "x")
    d.end()
    path = b.save(tmp_path / "tiny.xlsx")
    return path, b, blocks


def test_grid_and_lint_clean(tmp_path):
    path, b, blocks = _book(tmp_path)
    assert lint(path) == []
    wb = openpyxl.load_workbook(path)
    ws = wb["Calc"]
    assert ws.freeze_panes == "E5" and wb["Dashboard"].freeze_panes == "E5"
    assert ws.column_dimensions["C"].width == 56 and ws.column_dimensions["J"].width == 10
    assert ws["J4"].value == "=First_Year+COLUMN()-COLUMN($J$1)" and ws["J5"].value == "=DATE(J4+1,1,1)-1"
    bar = wb["Dashboard"]["C" + str(b.sheets["Dashboard"].row_of("MW") - 1)]
    assert bar.fill.fgColor.rgb.endswith("00563C")
    # three blocks with the same formula text up to row numbers, case number 1 / previous + 1
    rows = [blk["cf"] for blk in blocks]
    assert len({ws.cell(r, 10).value.replace(str(blk["cap"]), "#") for r, blk in zip(rows, blocks)}) == 1
    assert ws.cell(blocks[0]["bar"], 5).value == "=1"
    assert ws.cell(blocks[1]["bar"], 5).value == f"=$E${blocks[0]['bar']}+1"


def test_subheader_fill_is_light_green():
    from fepm.xl import styles
    assert styles.SUB_FILL == "E6F4EC"


def test_lever_exception_needs_blue(tmp_path):
    path, b, _ = _book(tmp_path)
    r = b.sheets["Dashboard"].row_of("MW")
    wb = openpyxl.load_workbook(path)
    wb["Dashboard"].cell(r, 6).font = openpyxl.styles.Font(color="000000")    # a black number in a lever range
    wb["Dashboard"]["C99"] = 5                                                  # outside any lever range
    wb.save(path)
    where = {(x.rule, x.where) for x in lint(path)}
    assert ("a-hardcode", f"Dashboard!F{r}") in where and ("a-hardcode", "Dashboard!C99") in where
    assert ("a-hardcode", f"Dashboard!E{r}") not in where


def test_xlsx_only(tmp_path):
    b = Book(CASES, years=2)
    b.sheet("Dashboard", "x")
    with pytest.raises(ValueError):
        b.save(tmp_path / "x.xlsm")


def test_annual_model_has_calendar_years_and_year_end_dates(tmp_path):
    m = Model("Annual", periods=3, periodicity="annual", start=dt.date(2025, 1, 1))
    m.calc.section("x")
    m.calc.row("one", "One", "EUR k", "=1")
    m.checks.add_check("One is one", "=N({one}<>1)")
    ws = m.calc.ws
    m.save(tmp_path / "a.xlsx")
    assert ws["J4"].value == "=EOMONTH(J$3,Months_Per_Period-1)" and ws["J7"].value == "=YEAR(J$4)"
    ref = m.names["Months_Per_Period"].split("!")[1].replace("$", "")
    assert m.inputs.ws[ref].value == 12          # 1 Jan start + 12 months: calendar years ending 31-Dec


@pytest.mark.skipif(__import__("shutil").which("soffice") is None, reason="LibreOffice not installed")
def test_recalc_values(tmp_path):
    from fepm.recalc import recalc
    path, b, blocks = _book(tmp_path)
    rc = recalc(path, tmp_path / "tiny_recalc.xlsx")
    wb = openpyxl.load_workbook(rc, data_only=True)
    d = wb["Dashboard"]
    sh = b.sheets["Dashboard"]
    assert d.cell(sh.row_of("irr"), 5).value == pytest.approx((2 / 1) ** (365 / 486) - 1, abs=1e-6)
    # return due: 10 × [(1 + 12 % × 32/365) × 1.12 × (1 + 12 % × 90/365) − 1] (one full year, 2026)
    f = (1 + 0.12 * 32 / 365) * 1.12 * (1 + 0.12 * 90 / 365)
    assert d.cell(sh.row_of("ret"), 5).value == pytest.approx(10 * (f - 1), abs=1e-9)
    assert wb["Calc"]["J4"].value == 2025 and wb["Calc"]["J5"].value == dt.datetime(2025, 12, 31)
