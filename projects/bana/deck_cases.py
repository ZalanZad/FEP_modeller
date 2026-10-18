"""Bana – short deck on the three cases (SPEC.md §22): base, full zero bid, 20 %+ IRR.

    python3 projects/bana/deck_cases.py        → projects/bana/out/cases/Bana_Cases.pptx (+ dated copy)

Every figure is bound to the packaged, recalculated workbook (live base, stored sets 17 and 18).
"""
import shutil
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))

from fepm.ppt.deck import Deck, fmt_eur_m, fmt_mult, fmt_num, fmt_pct  # noqa: E402

HERE = Path(__file__).resolve().parent
SECTIONS = ["Summary", "Assumptions", "Returns", "What it takes"]
TOP = 2.45
C1, C2, C3 = "Case 1: base", "Case 2: zero bid", "Case 3: 20 %+ IRR"


def eur(v):
    return fmt_eur_m(v)


def pct(v):
    return fmt_pct(v)


def pct0(v):
    return fmt_pct(v, 0)


def x2(v):
    return fmt_mult(v)


def n0(v):
    return fmt_num(v, 0)


def dp1(v):
    return f"{v:.1f}"


def pt(v):
    return f"{v:.0f}" if float(v).is_integer() else f"{v:.1f}"


def build_deck(xlsm: Path, out: Path) -> Path:
    d = Deck(xlsm, source=Path(xlsm).name, sections=SECTIONS, date_cell=("Dashboard", "E6"))
    d.source = f"{Path(xlsm).stem}_{d.date:%Y%m%d}{Path(xlsm).suffix}"
    f, names = d.f, d.wb.defined_names
    sets = [c.value for row in d.wb["Inputs"][next(names["Scenario_Names"].destinations)[1]] for c in row]
    ks = {C1: 1, C2: sets.index(C2) + 1, C3: sets.index(C3) + 1}

    def b(ref, fmt):
        return d.bind(*ref.split("!"), fmt=fmt) if "!" in ref else d.bind(ref, fmt=fmt)

    def st(key, case, fmt):        # stored result of a case (set 1 = live base, stored too)
        return d.stored(key, ks[case], fmt)

    def both(key, case, fmt):      # "European / Chinese"
        return f("{a} / {c}", a=st(f"{key}_1", case, fmt), c=st(f"{key}_2", case, fmt))

    def val(case, key, fmt):       # an input in a case: its set column, or the base when the set leaves it blank
        if case == C1:
            return b(key, fmt)
        row = next(names[key].destinations)[1].split("$")[-1]
        ref = f"Inputs!{chr(ord('H') + ks[case] - 1)}{row}"
        return b(ref, fmt) if d.wb["Inputs"][ref.split("!")[1]].value is not None else b(key, fmt)

    def lever(key, fmt):           # Dashboard bid lever, European route
        sheet, rng = next(names[f"Dash_{key}"].destinations)
        return d.bind(sheet, rng.split(":")[0].replace("$", ""), fmt=fmt)

    cases = (C1, C2, C3)

    # ---- 1 Cover ----------------------------------------------------------------------------------------------
    d.cover("Project Bana", "Three cases: base, zero bid and 20 %+ IRR",
            b("Dashboard!E6", lambda v: f"{v.day} {v:%B %Y}"))

    # ---- 2 Summary ---------------------------------------------------------------------------------------------
    s = d.content("Summary", f(
        "Case 3 lifts the hold IRR to {c} with Chinese turbines ({e} European), from {b1} ({b2}) in the base",
        c=st("HTM_IRR_2", C3, pct), e=st("HTM_IRR_1", C3, pct), b1=st("HTM_IRR_2", C1, pct),
        b2=st("HTM_IRR_1", C1, pct)), f(
        "Case 2 drops every bid commitment and assumes {h} h and a {p} €/MWh PPA; case 3 adds cheaper, longer debt and "
        "lower BoP capex. Both rest on a yield the data room does not yet support.",
        h=val(C2, "In_P50_Hours", n0), p=val(C2, "In_PPA_Price", n0)), title_lines=2)
    rows = [["Tender points", "points", *[both("Points", c, pt) for c in cases]],
            ["FEP equity to COD", "€ m", *[both("Equity_COD", c, eur) for c in cases]],
            ["FEP IRR, sale at COD", "%", *[both("COD_IRR", c, pct) for c in cases]],
            ["FEP IRR, hold", "%", *[both("HTM_IRR", c, pct) for c in cases]],
            ["FEP NPV, hold", "€ m", *[both("HTM_NPV", c, eur) for c in cases]],
            ["Minimum DSCR", "x", *[both("Min_DSCR", c, x2) for c in cases]]]
    s.table(["European / Chinese", "Unit", *cases], rows, key_rows={3}, top=TOP, width=10.0,
            height=0.59 + 6 * 0.6, col_widths=[2.9, 0.8, 2.1, 2.1, 2.1])
    s.chart("column", ["Case 1", "Case 2", "Case 3"],
            {f"{r} turbine": [st(f"HTM_IRR_{i}", c, None) for c in cases] for i, r in ((1, "European"), (2, "Chinese"))},
            left=15.2, top=TOP, width=4.7, height=4.2, title="FEP IRR, hold", number_format="0.0%", name="IRR chart")
    s.card("Reading", [
        ("Case 2: ", "the upside from yield and price, with the cost saved by bidding nothing."),
        ("Case 3: ", "the mildest combination we found that takes the Chinese route above 20 %."),
        ("Tender: ", f("both score {a} / {c} points against {b1} / {b2}: the award is the main risk.",
                       a=st("Points_1", C2, pt), c=st("Points_2", C2, pt), b1=st("Points_1", C1, pt),
                       b2=st("Points_2", C1, pt)))],
        left=4.92, top=6.9, width=10.0, height=1.7)
    s.footnote("Post-tax, FEP 100 %. Cases 2-3 are stored key-assumption sets of the model; everything not listed on "
               "the next slide is as in the base.")

    # ---- 3 Assumptions -------------------------------------------------------------------------------------------
    s = d.content("Assumptions", f(
        "Cases 2 and 3 change the yield, the bid, the PPA and the debt; case 3 also cuts BoP capex to {c} k€/MW",
        c=val(C3, "In_Other_Wind", n0)), "Everything else stays as in the base, including the turbine route, the "
        "sale price per MW and the tax.", title_lines=2)
    rows = [["What changes", "", "", "", ""],
            ["Net P50 hours", "h", *[val(c, "In_P50_Hours", n0) for c in cases]],
            ["Overbuild / BESS committed", "% of MVA", f("{o} / {s_}", o=lever("Overbuild", pct0),
                                                         s_=lever("BESS", pct0)), "none", "none"],
            ["Municipal one-off / community", "k€/MW, €/MWh", f("{m} / {c}", m=lever("Municipal", n0),
                                                                 c=lever("Community", dp1)), "none", "none"],
            ["PPA price, whole life", "€/MWh", *[val(c, "In_PPA_Price", n0) for c in cases]],
            ["Gearing (debt / construction cost)", "%", *[val(c, "In_Gearing", pct0) for c in cases]],
            ["Debt tenor", "years", *[val(c, "In_Debt_Tenor", n0) for c in cases]],
            ["All-in interest rate", "%", *[val(c, "In_Debt_Rate", pct) for c in cases]],
            ["BoP and other wind capex", "k€/MW", *[val(c, "In_Other_Wind", n0) for c in cases]],
            ["Same in all three", "", "", "", ""],
            ["Turbines, European / Chinese", "k€/MW", *[f("{e} / {c}", e=b("In_Turbine", n0),
                                                          c=b("In_Turbine_CN", n0)) for _ in cases]],
            ["Sale price at COD, European / Chinese", "k€/MW", *[f("{e} / {c}", e=b("In_COD_Price", n0),
                                                                  c=b("In_COD_Price_CN", n0)) for _ in cases]],
            ["Corporate income tax", "%", *[b("In_CIT", pct0) for _ in cases]]]
    s.table(["", "Unit", *cases], rows, sections={0, 9}, top=TOP, width=10.4, height=0.59 + 13 * 0.5,
            col_widths=[3.7, 1.3, 1.8, 1.8, 1.8])
    s.card("Why", [
        ("Yield: ", f("{h} h instead of {h0}; to be confirmed by the EYA, which is not in the data room.",
                      h=val(C2, "In_P50_Hours", n0), h0=b("In_P50_Hours", n0))),
        ("Zero bid: ", "no overbuild, BESS, municipal or community payment saves capex and opex."),
        ("PPA: ", "COD is 2030 at the earliest, leaving two years to sign a bankable long-term PPA; Aurora is "
                  "conservative and today's high prices may last longer."),
        ("Debt: ", "a signed long-term PPA should support more and longer debt than the banks' first view "
                   "(48–58 % LTC over 12–15 years)."),
        ("BoP: ", "a tighter BoP and installation budget; to be tested with contractors.")],
        left=15.45, top=TOP, width=4.45, height=7.4)
    s.footnote("Bid levers shown for the European route (the Chinese bid is the same, without the EU-turbine points). "
               "Bank terms as reported in the October 2026 meetings.")

    # ---- 4 Returns ------------------------------------------------------------------------------------------------
    s = d.content("Returns", f(
        "Lower capex and more debt cut the equity to COD to €{e}m in case 3 (€{c}m Chinese), from €{b1}m",
        e=st("Equity_COD_1", C3, eur), c=st("Equity_COD_2", C3, eur), b1=st("Equity_COD_1", C1, eur)), (
        "The sale at COD gains only from lower capex and more debt, because the sale price is fixed per MW; the "
        "hold gains from yield and price as well."), title_lines=2)

    def six(key, fmt):
        return [st(f"{key}_{i}", c, fmt) for c in cases for i in (1, 2)]

    rows = [["Total uses to COD", "€ m", *six("Uses_Total", eur)],
            ["Senior debt at COD", "€ m", *six("Debt_COD", eur)],
            ["Debt share of construction cost", "%", *six("Debt_Share", pct0)],
            ["FEP equity to COD", "€ m", *six("Equity_COD", eur)],
            ["Sale at COD", "", "", "", "", "", "", ""],
            ["FEP profit", "€ m", *six("COD_Profit", eur)],
            ["FEP IRR", "%", *six("COD_IRR", pct)],
            ["FEP MOIC", "x", *six("COD_MOIC", x2)],
            ["Buyer IRR, unlevered", "%", *six("Buyer_IRR_U", pct)],
            ["Hold to maturity", "", "", "", "", "", "", ""],
            ["FEP IRR", "%", *six("HTM_IRR", pct)],
            ["FEP MOIC", "x", *six("HTM_MOIC", x2)],
            ["FEP NPV", "€ m", *six("HTM_NPV", eur)],
            ["Minimum DSCR", "x", *six("Min_DSCR", x2)]]
    head = ["FEP, post-tax", "Unit"] + [f"{c.split(':')[0]} {r}" for c in cases for r in ("EU", "CN")]
    s.table(head, rows, sections={4, 9}, key_rows={6, 10}, top=TOP, width=15.0, height=0.59 + 14 * 0.48,
            col_widths=[3.4, 0.8] + [1.8] * 6)
    s.footnote(f("EU = European turbines, CN = Chinese turbines. NPV at {r}. Debt is sized on the gearing; the DSCR "
                 "is a floor of {x} and does not bind.", r=b("In_Discount_Rate", pct0), x=b("In_DSCR", x2)))

    # ---- 5 What it takes ------------------------------------------------------------------------------------------
    s = d.content("What it takes", f(
        "Case 3 needs a {p} €/MWh PPA, {g} gearing over {t} years and cheaper BoP, and still leaves the European "
        "route at {e}", p=val(C3, "In_PPA_Price", n0), g=val(C3, "In_Gearing", pct0), t=val(C3, "In_Debt_Tenor", n0),
        e=st("HTM_IRR_1", C3, pct)), "Price and debt move the hold IRR most; the zero bid and the yield need checking "
        "before any of this is relied on.", title_lines=2)
    s.card("What has to be true", [
        ("PPA: ", f("a bankable long-term PPA at {p} €/MWh (base {p0}) signed before financing.",
                    p=val(C3, "In_PPA_Price", n0), p0=b("In_PPA_Price", n0))),
        ("Debt: ", f("{g} gearing over {t} years at {r}; the minimum DSCR is {d} (Chinese), so the cash flows carry it.",
                     g=val(C3, "In_Gearing", pct0), t=val(C3, "In_Debt_Tenor", n0), r=val(C3, "In_Debt_Rate", pct),
                     d=st("Min_DSCR_2", C3, x2))),
        ("Capex: ", f("BoP, installation and other wind costs at {c} k€/MW, from {c0}.", c=val(C3, "In_Other_Wind", n0),
                      c0=b("In_Other_Wind", n0))),
        ("Yield: ", f("{h} h confirmed by an EYA.", h=val(C3, "In_P50_Hours", n0)))],
        left=4.92, top=TOP, width=7.3, height=3.2)
    s.card("Risks", [
        ("Award: ", f("the zero bid scores {a} / {c} points against {b1} / {b2}; it may not win the connection.",
                      a=st("Points_1", C3, pt), c=st("Points_2", C3, pt), b1=st("Points_1", C1, pt),
                      b2=st("Points_2", C1, pt))),
        ("Minimum payments: ", "a coming law may make a municipal and community minimum mandatory even for a zero "
                               "bid (Kiírás footnote 1)."),
        ("Banks: ", "both banks' first view is less debt over a shorter tenor; neither has financed a project like "
                    "this."),
        ("European route: ", f("{e} hold IRR in case 3; reaching 20 % would also need much cheaper turbines and a "
                               "higher PPA.", e=st("HTM_IRR_1", C3, pct))),
        ("Sale price: ", "fixed per MW in all cases; a buyer would pay more for a higher PPA, which the sale IRR "
                         "does not show.")],
        left=12.6, top=TOP, width=7.3, height=3.2)
    s.footnote(f("Model checks in all three cases: {ck}. Cases are stored by the scenario macro with the tax solver.",
                 ck=d.bind_sum([st(f"Checks_{i}", c, None) for c in cases for i in (1, 2)],
                               lambda v: "all pass" if v == 0 else "FAIL")))
    path = d.save(out / "Bana_Cases.pptx")
    shutil.copy(path, out / f"Bana_Cases_{d.date:%Y%m%d}.pptx")
    return path


if __name__ == "__main__":
    print(build_deck(HERE / "out" / "Bana_Model.xlsm", HERE / "out" / "cases"))
