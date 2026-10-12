"""Demo end to end: lint, master check, macros, independent recomputation, deck tie-out."""
import csv
import datetime as dt

from fepm.audit.lint import lint, lint_values
from fepm.ppt import deck as fmt
from fepm.recalc import name_value, values
from fepm.xl.macros import embedded_modules
from test_recalc import xirr

TOL = 1e-6


def test_lint_and_master_check(demo):
    assert lint(demo["model"]) == []
    assert lint_values(demo["recalc"]) == []


def test_vba_modules_embedded(demo):
    mods = embedded_modules(demo["model"])
    assert {"Solver", "GoalSeek", "Scenarios"} <= set(mods)
    assert "Sub RunAllScenarios" in mods["Scenarios"]


def _eomonth_year_end(year):
    return dt.date(year, 12, 31)


def test_independent_recomputation(demo):
    wb = values(demo["recalc"])
    v = lambda n: name_value(wb, n)  # noqa: E731
    start, cod, life = v("In_Model_Start"), v("In_COD"), v("In_Operating_Life")
    start, cod = start.date(), cod.date()
    ops_end = dt.date(cod.year + int(life), cod.month, cod.day)
    mw, ncf, capex_mw, opex_mw = v("In_Capacity"), v("In_Capacity_Factor"), v("In_Capex_per_MW"), v("In_Opex_per_MW")
    tariff, rate, base_year = v("Tariff_Input"), v("In_Discount_Rate"), v("In_Base_Year")
    rows = {r[2].value: [c.value for c in r][9:] for r in wb["Curves"].iter_rows() if r[2].value}
    cpi = dict(zip(rows["Year"], rows[v("In_CPI_Curve")]))
    cfs, dates, idx = [], [], 1.0
    constr_days = (cod - start).days + 1
    for year in range(start.year, start.year + 26):
        p0, p1 = dt.date(year, 1, 1), _eomonth_year_end(year)
        days = (p1 - p0).days + 1
        share = lambda a, b: max(0, ((min(b, p1) - max(a, p0)).days + 1) / days)  # noqa: E731
        constr, ops = share(start, cod), share(cod + dt.timedelta(days=1), ops_end)
        idx = 1.0 if year == base_year else idx * (1 + cpi[year])
        gen = ops * days * 24 * mw * ncf
        rev = gen * tariff / 1000
        opex = -opex_mw * mw * days / 365 * ops * idx
        capex = -capex_mw * mw * constr * days / constr_days
        cfs.append(rev + opex + capex)
        dates.append(p1)
    irr = xirr(cfs, dates)
    npv = sum(c / (1 + rate) ** ((d - dates[0]).days / 365) for c, d in zip(cfs, dates))
    moic = sum(c for c in cfs if c > 0) / -sum(c for c in cfs if c < 0)
    assert abs(irr - v("Out_Project_IRR")) < TOL
    assert abs(irr - v("In_Target_IRR")) < TOL              # goal seek hit the target
    assert abs(npv / 1000 - v("Out_NPV")) < TOL
    assert abs(moic - v("Out_MOIC")) < TOL
    assert abs(sum(cfs) / 1000 - v("Out_Net_CF")) < TOL


def test_stored_results_match_live(demo):
    wb = values(demo["recalc"])
    s = wb["Summary"]
    for r in range(6, 14):
        assert abs(s.cell(r, 5).value - s.cell(r, 7).value) < TOL     # Base stored = live
    assert s["H6"].value > s["G6"].value                                 # high capex needs a higher tariff


def test_deck_bindings_tie_out(demo):
    wb = values(demo["model"])
    with open(demo["out"] / "deck_bindings.csv") as fh:
        rows = list(csv.DictReader(fh))
    assert len(rows) > 20
    for r in rows:
        sheet, cell = r["cell"].split("!")
        assert abs(float(r["value"]) - wb[sheet][cell].value) < 1e-9


def test_formatters():
    assert fmt.fmt_eur_m(-5.26) == "(5.3)"
    assert fmt.fmt_eur_m(1234.56) == "1,234.6"
    assert fmt.fmt_pct(-0.075) == "−7.5%"
    assert fmt.fmt_pct(0.209) == "20.9%"
    assert fmt.fmt_mult(1.3149) == "1.31x"
    assert fmt.fmt_delta(3.3) == "+3.3"
    assert fmt.fmt_num(2450, 0) == "2,450"
    assert fmt.fmt_pct("n/a") == "–"
