"""Demo: 10 MW onshore wind CfD bid – which tariff earns the target project IRR?

Annual timeline, 3 scenarios, pre-tax project cash flows. The bid tariff is solved by GoalSeek.bas
(purple input) so that Project IRR = target IRR; RunAllScenarios stores every scenario on Summary.
"""
import datetime as dt
from pathlib import Path

from fepm.ppt.deck import Deck, fmt_eur_m, fmt_mult, fmt_pct, fmt_num
from fepm.xl.builder import Model
from fepm.xl.macros import package

NAME = "Demo_Wind_Bid"
SOURCE = "Demo wind bid model"


def build_model(out: Path) -> Path:
    m = Model("Project Demo", "10 MW wind CfD bid", periods=26, periodicity="annual",
              start=dt.date(2026, 1, 1), scenarios=["Base", "High capex", "Low yield"], curves=True,
              version="v1", author="FEP", date=dt.date(2026, 10, 12))
    i = m.inputs
    i.add_constant("Hours_In_Day", "Hours in day", "hours", 24, "Calendar")
    cod = i.add_input("COD", "Commercial operation date (COD)", "date", dt.date(2026, 12, 31), source="Demo plan")
    life = i.add_input("Operating_Life", "Operating life", "years", 25, source="Turbine design life")
    base_year = i.add_input("Base_Year", "Price base year", "year", 2026, source="Real terms basis", fmt="int")
    i.section("Project")
    mw = i.add_input("Capacity", "Installed capacity", "MW", 10, source="Grid connection")
    capex = i.add_input("Capex_per_MW", "Capex per MW", "EUR k/MW", 1300, {2: 1500}, source="EPC benchmark")
    ncf = i.add_input("Capacity_Factor", "Net capacity factor (P50)", "%", 0.30, {3: 0.26}, source="Yield report")
    opex = i.add_input("Opex_per_MW", "Opex per MW per year (real)", "EUR k/MW", 40, source="O&M benchmark")
    cpi_curve = i.add_input("CPI_Curve", "CPI curve", "text", "Central", {2: "High"}, source="Curves sheet")
    i.section("Returns")
    target = i.add_input("Target_IRR", "Target project IRR", "%", 0.08, source="FEP hurdle rate")
    rate = i.add_input("Discount_Rate", "Discount rate for NPV", "%", 0.07, source="FEP WACC")

    years = list(range(2026, 2052))
    m.curves.add_block("CPI", "CPI (annual inflation)", "%", years,
                       {"Low": [0.015] * 26, "Central": [0.025] * 26, "High": [0.035] * 26}, source="Demo curves")

    c = m.calc
    c.section("Timing")
    c.window("constr", "Construction period", m.inputs.model_start, cod)
    c.window("ops", "Operating period", f"{cod}+1", f"EDATE({cod},{life}*Months_In_Year)")
    c.curve("cpi", "CPI", "%", "CPI", cpi_curve)
    c.index("idx", "Opex index", base_year, "cpi")
    c.section("Operations")
    c.row("gen", "Net generation", "MWh", f"={{ops}}*{{days}}*Hours_In_Day*{mw}*{ncf}")
    c.row("rev", "Revenue", "EUR k", "={gen}*Tariff_Input/Thousand")
    c.row("opex", "Opex", "EUR k", f"=-{opex}*{mw}*{{yf}}*{{ops}}*{{idx}}")
    c.row("ebitda", "EBITDA", "EUR k", "={rev}+{opex}", bold=True)
    c.section("Investment")
    c.row("capex", "Capex", "EUR k", f"=-{capex}*{mw}*{{constr}}*{{days}}/{{self.E}}",
          e=f"={cod}-{m.inputs.model_start}+1", ef_fmt="int")
    c.section("Cash flow and returns")
    c.row("cf", "Project cash flow (pre-tax)", "EUR k", "={ebitda}+{capex}", bold=True)
    c.xirr("irr", "Project IRR", "cf")
    c.xnpv("npv", "Project NPV", rate, "cf")
    c.moic("moic", "MOIC", "cf")

    i.section("Solved by macro")
    m.goal_seek("Tariff", "Bid tariff (solved for target IRR)", "EUR/MWh", 60.0,
                achieved=c.ref("irr", "E"), target=target)

    k = m.checks
    k.add_check("Generation is not negative", "=N({gen}<0)")
    k.add_check("Construction and operations do not overlap", "=N({constr}+{ops}>1)")
    k.add_check("Capex spent = capex budget",
                f"=N(ABS(-{{capex.I}}-{capex}*{mw})>Tolerance)", per_period=False)
    k.add_check("Cash flow = revenue + opex + capex",
                "=N(ABS({cf.I}-{rev.I}-{opex.I}-{capex.I})>Tolerance)", per_period=False)

    s = m.summary
    s.add_output("Bid_Tariff", "Bid tariff", "EUR/MWh", "Tariff_Input", bold=True)
    s.add_output("Capex", "Capex", "EUR m", "=-" + c.ref("capex"), scale="Thousand")
    s.add_output("Revenue", "Revenue (lifetime)", "EUR m", c.ref("rev"), scale="Thousand")
    s.add_output("EBITDA", "EBITDA (lifetime)", "EUR m", c.ref("ebitda"), scale="Thousand")
    s.add_output("Net_CF", "Net project cash flow", "EUR m", c.ref("cf"), scale="Thousand")
    s.add_output("Project_IRR", "Project IRR", "%", c.ref("irr", "E"), bold=True)
    s.add_output("NPV", "Project NPV at discount rate", "EUR m", c.ref("npv", "E"), scale="Thousand")
    s.add_output("MOIC", "MOIC", "x", c.ref("moic", "E"))

    xlsx = m.save(out / f"{NAME}.xlsx")
    return package(xlsx, out / f"{NAME}.xlsm")      # embeds vba/*.bas, runs RunAllScenarios


def build_deck(xlsm: Path, out: Path) -> Path:
    d = Deck(xlsm, source=SOURCE, date=dt.date(2026, 10, 12), sections=["Recommendation", "Economics"])
    d.cover("Project Demo", "10 MW wind CfD bid", "12 October 2026")
    scen = ["Base", "High capex", "Low yield"]
    rows = [("Bid_Tariff", fmt_num), ("Capex", fmt_eur_m), ("Revenue", fmt_eur_m), ("Net_CF", fmt_eur_m),
            ("Project_IRR", fmt_pct), ("NPV", fmt_eur_m), ("MOIC", fmt_mult)]
    labels = {k: (d.wb["Summary"][c.replace("$E$", "C")].value, d.wb["Summary"][c.replace("$E$", "D")].value)
              for k, c in ((k, next(d.wb.defined_names[f"Out_{k}"].destinations)[1]) for k, _ in rows)}
    tariff = d.bind("Out_Bid_Tariff", fmt=fmt_num)
    target = d.bind("In_Target_IRR", fmt=lambda v: fmt_pct(v, 0))
    life = d.bind("In_Operating_Life", fmt=lambda v: f"{v:.0f}")
    s = d.content("Recommendation", f"Bid {tariff} EUR/MWh to earn the {target} target IRR in the base case",
                  f"Capex of EUR {d.bind('Out_Capex', fmt=fmt_eur_m)}m returns "
                  f"{d.bind('Out_MOIC', fmt=fmt_mult)} the investment over {life} years")
    s.table(["Headline outputs", "Unit"] + scen,
            [list(labels[k]) + [d.stored(k, i, f) for i in (1, 2, 3)] for k, f in rows],
            key_rows={0, 4}, top=1.75, height=3.6)
    s.card("Decision asked", [("Approve ", f"a bid of {tariff} EUR/MWh."),
                              ("Downside: ", f"high capex needs {d.stored('Bid_Tariff', 2, fmt_num)} EUR/MWh, "
                                             f"low yield {d.stored('Bid_Tariff', 3, fmt_num)} EUR/MWh.")],
           left=4.92, top=5.75, width=7.0, height=1.6)
    s.footnote(f"Pre-tax project cash flows over a {life}-year life, opex indexed with CPI. "
               "The bid tariff is solved by macro (GoalSeek.bas).")
    rate = d.bind("In_Discount_Rate", fmt=lambda v: fmt_pct(v, 0))
    worst = max((1, 2, 3), key=lambda i: d.stored("Bid_Tariff", i))
    s = d.content("Economics", f"{scen[worst - 1]} needs the highest tariff: "
                               f"{d.stored('Bid_Tariff', worst, fmt_num)} EUR/MWh",
                  f"Bid tariff per scenario and NPV at the {rate} discount rate")
    s.chart("column", scen, {"Bid tariff (EUR/MWh)": [d.stored("Bid_Tariff", i) for i in (1, 2, 3)]},
            left=4.92, top=1.9, width=7.3, height=5.0, title="Bid tariff (EUR/MWh)", number_format="0.0")
    s.chart("bar", scen, {"NPV (EUR m)": [d.stored("NPV", i) for i in (1, 2, 3)]},
            left=12.6, top=1.9, width=7.3, height=5.0, title="NPV at discount rate (EUR m)", number_format="0.0")
    s.footnote("NPV at the solved bid tariff, discounted to the end of the first period.")
    return d.save(out / f"{NAME}_IC.pptx")


def build(out: Path):
    out.mkdir(parents=True, exist_ok=True)
    xlsm = build_model(out)
    deck = build_deck(xlsm, out)
    return {"model": xlsm, "deck": deck}
