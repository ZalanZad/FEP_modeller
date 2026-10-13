"""Bana IC deck – 7 slides to SPEC.md §9, streamlined from reference/Bana_IC_20260109_vDSz.pptx.

Every figure is `Deck.bind` on the packaged, LibreOffice-recalculated workbook. The Dentons DD risk table and the
deal narrative are qualitative and come from the original deck (the table is read from it at build time).
"""
from pathlib import Path

from pptx import Presentation
from pptx.util import Pt

from fepm.ppt.deck import DASH, ROOT, Deck, fmt_eur_m, fmt_mult, fmt_num, fmt_pct

SOURCE_DECK = ROOT / "reference" / "Bana_IC_20260109_vDSz.pptx"
SECTIONS = ["Executive summary", "Risks and mitigants", "Technical overview and tender bid", "Funding to COD",
            "Returns: sale at COD vs hold", "Sensitivities"]
TOP = 2.45                      # first free line under a two-line title and subtitle
ROUTES = ("European", "Chinese")


def eur(v):
    return fmt_eur_m(v)


def neg(v):                      # an amount shown as a deduction: (60.9)
    return fmt_eur_m(-v)


def pct(v):
    return fmt_pct(v)


def pct0(v):
    return fmt_pct(v, 0)


def x2(v):
    return fmt_mult(v)


def x1(v):
    return fmt_mult(v, 1)


def n0(v):
    return fmt_num(v, 0)


def n1(v):
    return fmt_num(v, 1)


def dp1(v):
    return f"{v:.1f}"


def k_to_m(v):                   # EUR k input shown in EUR m
    return fmt_eur_m(v / 1000, 2)


def date(v):                     # non-breaking hyphens: a date never wraps across two lines
    return f"{v:%d-%b-%y}".replace("-", "‑")


def day(v):
    return f"{v.day} {v:%B}"


def dd_table():
    """Header and rows of the Dentons risk table (slide 3 of the source deck), with its section rows."""
    prs = Presentation(SOURCE_DECK)
    shape = next(sh for sh in prs.slides[2].shapes if sh.name == "Risk table")
    foot = next(sh for sh in prs.slides[2].shapes if sh.name == "Footnote").text_frame.paragraphs
    rows = [[c.text.replace("\n", " ") for c in r.cells] for r in shape.table.rows]
    return rows[0], rows[1:], foot[0].text, foot[1].text


def build_deck(xlsm: Path, out: Path) -> Path:
    d = Deck(xlsm, source=Path(xlsm).name, sections=SECTIONS, date_cell=("Dashboard", "E6"))
    f = d.f
    names = d.wb.defined_names

    def b(ref, fmt):
        """Bind a workbook name (`Out_…`, `In_…`) or a `Sheet!Cell` reference."""
        return d.bind(*ref.split("!"), fmt=fmt) if "!" in ref else d.bind(ref, fmt=fmt)

    def o(key, fmt):             # both routes of an output: [European, Chinese]
        return [b(f"Out_{key}_{i}", fmt) for i in (1, 2)]

    def lever(key, i, fmt, grid=None):
        """Dashboard lever value of route i (1 EU, 2 CN), or its grid cell (`grid`: column letter)."""
        sheet, rng = next(names[f"Dash_{key}"].destinations)
        first, row = rng.split(":")[0].replace("$", "")[0], rng.split("$")[-1]
        col = grid or chr(ord(first) + i - 1)
        return d.bind(sheet, f"{col}{row}", fmt)

    def pts(key, i):             # points of one commitment: grid column I (EU route) or J (CN route)
        return lever(key, i, n0, grid="IJ"[i - 1])

    sets = [c.value for row in d.wb["Inputs"][next(names["Scenario_Names"].destinations)[1]] for c in row]

    # ---- 1 Cover -------------------------------------------------------------------------------------------
    d.cover("Project Bana", f("{mw} MW wind development project opportunity", mw=b("In_Wind_MW", n0)),
            b("Dashboard!E6", lambda v: f"{v.day} {v:%B %Y}"))

    # ---- 2 Executive summary -----------------------------------------------------------------------------------
    cod, htm, eq, p = o("COD_IRR", pct), o("HTM_IRR", pct), o("Equity_COD", eur), o("Points", n0)
    s = d.content("Executive summary", f(
        "At €{pp}m, Bana returns {a} (European) / {c} (Chinese) on a sale at COD and {h} / {h2} if held",
        pp=b("In_Purchase_Price", k_to_m), a=cod[0], c=cod[1], h=htm[0], h2=htm[1]), f(
        "FEP funds €{e1}m of equity to COD (€{e2}m Chinese) and bids {p1} points ({p2} Chinese) into the "
        "{bd} tender; the European route needs a permit amendment.",
        e1=eq[0], e2=eq[1], p1=p[0], p2=p[1], bd=b("In_Bid_Date", day)), title_lines=2)
    s.text("Narrative", "\n".join([
        "GEI has been developing a wind project near the Huayou Cobalt plant in Bana, first as a behind-the-meter "
        "supply. It holds a building permit for 6 turbines, is close to RtB and is for sale.",
        "",
        "The 2026 MEKH tender offers a grid connection at the Ács Katódgyár node. GEI cannot bid, so a share deal is "
        "not feasible: a transfer of rights lets FEP bid in its own name, with closing subject to the award. All land, "
        "rights and permits move to FEP. Dentons (sell side) has issued a legal DD report and a draft transfer "
        "agreement.",
        "",
        "RtB is not yet reached: cable-route permits are missing, the Pér airport flight-path agreement is open and "
        "only 6 of 9 turbine positions are permitted. The permits name Goldwind turbines; another manufacturer needs "
        "a permit amendment.",
        "",
        "The seller asks a near-RtB price, so an RtB exit is not modelled: value comes from building and selling at "
        "COD or holding the asset.",
        "",
        "FEP asks the IC to approve the purchase price and to give its view on buying and developing the project."]),
        4.92, TOP, 7.6, 7.9, size=15)
    rows = [["Tender points", "points", *p],
            ["FEP equity to COD", "€ m", *eq],
            ["Sale at COD", "", "", ""],
            ["FEP profit", "€ m", *o("COD_Profit", eur)],
            ["FEP IRR", "%", *cod],
            ["FEP MOIC", "x", *o("COD_MOIC", x2)],
            ["Hold to maturity", "", "", ""],
            ["FEP IRR", "%", *htm],
            ["FEP MOIC", "x", *o("HTM_MOIC", x2)],
            ["FEP NPV", "€ m", *o("HTM_NPV", eur)],
            ["Minimum DSCR", "x", *o("Min_DSCR", x2)]]
    s.table(["Key numbers", "Unit", *ROUTES], rows, sections={2, 6}, key_rows={4, 7}, left=12.85, top=TOP,
            width=7.05, height=0.59 + 11 * 0.55, col_widths=[3.05, 0.9, 1.55, 1.55])
    s.footnote(f("Post-tax, FEP 100 %. Purchase price paid at closing ({cl}), after the award. PPA {ppa} €/MWh for "
                 "the whole life; NPV at {r}. Chinese = the permitted Goldwind turbines.",
                 cl=b("In_Closing", date), ppa=b("In_PPA_Price", n0), r=b("In_Discount_Rate", pct0)))

    # ---- 3 Risks and mitigants (Dentons DD, unchanged) --------------------------------------------------------
    head, rows, note, src = dd_table()
    sections = {i for i, r in enumerate(rows) if not any(r[1:])}
    s = d.content("Risks and mitigants",
                  "Legal DD finds no high risks; four of the five medium risks are fixed before signing or closing",
                  "Dentons legal due diligence – no high risks, five medium and seven low", title_lines=2)
    t = s.table(head, rows, sections=sections, top=TOP, height=7.9, col_widths=[2.6, 6.0, 1.0, 5.4], align="LLLL",
                unit_col=None)
    for row in t.table.rows:                 # the DD wording is long: 11 pt as in the source deck
        for cell in row.cells:
            for run in cell.text_frame.paragraphs[0].runs:
                run.font.size = Pt(11)
    s.text("Footnote", f"{note}\n{src}", 4.92, 10.55, 14.4, 0.7, size=11, color="919191")

    # ---- 4 Technical overview and tender bid ---------------------------------------------------------------------
    s = d.content("Technical overview and tender bid", f(
        "{mw} MW producing {g} GWh a year; the bid scores {p1} points with European turbines and {p2} with Chinese",
        mw=b("In_Wind_MW", n0), g=b("Out_Gen_Year1_1", n1), p1=p[0], p2=p[1]), f(
        "European turbines add {t} points but need a permit amendment; the PB stays at the minimum and partial "
        "award, shared connection and EU transformers are not offered.", t=pts("EU_Turbine", 1)), title_lines=2)
    rows = [["Site and grid", ""],
            ["Grid node", "Ács Katódgyár (BAMO B 132.00)"],
            ["Net connection capacity / installed wind", f("{a} MVA / {w} MW", a=b("In_MVA", n0),
                                                           w=b("In_Wind_MW", n0))],
            ["Energy yield and price", ""],
            ["Net P50 equivalent hours", f("{h} h", h=b("In_P50_Hours", n0))],
            ["Curtailment / degradation", f("{c} / {g} p.a.", c=b("In_Curtailment", pct),
                                            g=b("In_Degradation", lambda v: fmt_pct(v, 2)))],
            ["Net generation, first full year", f("{g} GWh", g=b("Out_Gen_Year1_1", n1))],
            ["PPA, 100 % of output, whole life", f("{p} €/MWh at COD, +{e} p.a.", p=b("In_PPA_Price", n0),
                                                   e=b("In_PPA_Esc", pct))],
            ["Timeline", ""],
            ["Bid / award / closing", f("{a} / {w} / {c}", a=b("In_Bid_Date", date), w=b("In_Award_Date", date),
                                       c=b("In_Closing", date))],
            ["NTP / COD", f("{n} / {c}", n=b("In_NTP", date), c=b("In_COD", date))],
            ["Acquisition", ""],
            ["Purchase price / DD / devex", f("€{p}m / €{dd}m / €{dv}m", p=b("In_Purchase_Price", k_to_m),
                                             dd=b("In_DD", k_to_m), dv=b("In_Devex", k_to_m))]]
    s.table(["Item", "Assumption"], rows, sections={0, 3, 8, 11}, top=TOP, width=8.1, height=0.59 + 13 * 0.55,
            col_widths=[3.9, 4.2], align="LL", unit_col=None)
    rows = [["ÖSZ-1 municipal one-off", f("{a} k€/MW (€{m}m)", a=lever("Municipal", 1, n0),
                                          m=b("Out_Uses_Municipal_1", lambda v: fmt_eur_m(v, 2))),
             pts("Municipal", 1), pts("Municipal", 2)],
            ["ÖSZ-2 community contribution", f("{a} €/MWh", a=lever("Community", 1, dp1)),
             pts("Community", 1), pts("Community", 2)],
            ["ÖSZ-6 performance security", f("{a} of the minimum", a=lever("PB_Mult", 1, x1)),
             pts("PB_Mult", 1), pts("PB_Mult", 2)],
            ["ÖSZ-9 European turbines ≥ 80 %", "European route only", pts("EU_Turbine", 1), pts("EU_Turbine", 2)],
            ["ÖSZ-5 partial award, ÖSZ-8 shared connection, ÖSZ-9 EU transformers", "not offered",
             f("{a} / {b} / {c}", a=pts("Partial", 1), b=pts("Shared", 1), c=pts("EU_Transformer", 1)),
             f("{a} / {b} / {c}", a=pts("Partial", 2), b=pts("Shared", 2), c=pts("EU_Transformer", 2))],
            ["Total score", "", *p]]
    s.table(["Tender bid", "Commitment", "European", "Chinese"], rows, key_rows={5}, left=13.35, top=TOP,
            width=6.55, height=0.59 + 6 * 0.62, col_widths=[2.35, 1.9, 1.15, 1.15], align="LLRR", unit_col=None)
    s.footnote("Points are the model's own score on the MEKH 2026 bands (Kiírás Annex 3-4). Overbuild (ÖSZ-3), "
               "BESS (ÖSZ-4) and brownfield (ÖSZ-7) are not available at Bana. The municipal one-off is paid at COD.")

    # ---- 5 Funding to COD ------------------------------------------------------------------------------------------
    debt, uses = o("Debt_COD", eur), o("Uses_Total", eur)
    s = d.content("Funding to COD", f(
        "Uses to COD of €{u1}m with European turbines (€{u2}m Chinese); senior debt funds {g1} ({g2}) and FEP the "
        "rest", u1=uses[0], u2=uses[1], g1=b("Out_Gearing_Uses_1", pct0), g2=b("Out_Gearing_Uses_2", pct0)), f(
        "A €{g}m bid guarantee at the bid; after the award a €{pb}m PB guarantee and €{t}m of TSO cash; the €{c}m "
        "connection fees fall due at the HCSSZ, {hd}.", t=b("Out_Sec_tso_1", lambda v: fmt_eur_m(v, 2)),
        g=b("Out_Sec_bid_1", lambda v: fmt_eur_m(v, 2)), pb=b("Out_Sec_pb_1", lambda v: fmt_eur_m(v, 2)),
        c=d.bind_sum(["Out_Sec_direct_1", "Out_Sec_indirect_1"], eur), hd=b("In_HCSSZ_Date", date)), title_lines=2)
    rows = [["Purchase price, DD and devex", "€ m", *o("Uses_Development", eur)],
            ["Participation fee, TSO security, guarantee fees", "€ m", *o("Uses_Securities", eur)],
            ["Construction capex", "€ m", *o("Uses_Capex", eur)],
            ["Connection fees, net of TSO security", "€ m", *o("Uses_Connection", eur)],
            ["Municipal one-off (ÖSZ-1)", "€ m", *o("Uses_Municipal", eur)],
            ["Capitalised interest and debt fees", "€ m", *o("Uses_Financing", eur)],
            ["Land lease during construction", "€ m", *o("Uses_Lease", eur)],
            ["Total uses", "€ m", *uses],
            ["Senior debt at COD (incl. capitalised interest)", "€ m", *debt],
            ["FEP equity to COD", "€ m", *eq],
            ["FEP peak equity if held (incl. DSRA)", "€ m", *o("Equity_Peak_HTM", eur)]]
    s.table(["Uses and sources to COD", "Unit", *ROUTES], rows, key_rows={7, 9}, top=TOP, width=8.3,
            height=0.59 + 11 * 0.55, col_widths=[4.6, 0.9, 1.4, 1.4])
    mn = lambda v: fmt_eur_m(v, 2)  # noqa: E731
    sec = [("fee", "Participation fee", "cash", "not refunded"),
           ("bid", "Bid security", "guarantee", "released at PB posting"),
           ("pb", "Performance security", "guarantee", f("held to {r}", r=b("In_PB_Release", date))),
           ("tso", "TSO connection-request security", "cash", "credited to the indirect fee"),
           ("direct", "Direct connection fee", "cash", f("paid {h}", h=b("In_HCSSZ_Date", date))),
           ("indirect", "Indirect connection fee, net", "cash", f("paid {h}", h=b("In_HCSSZ_Date", date)))]
    rows = [[lbl, form, b(f"Out_Sec_{k}_1", mn), what] for k, lbl, form, what in sec]
    rows.append(["Guarantee fees to COD", "cash", b("Out_Guarantee_Fees_COD_1", mn),
                 f("{r} p.a. on guarantees", r=b("In_Guarantee_Fee", pct0))])
    s.table(["Tender securities and fees", "Form", "€ m", "What happens"], rows, left=13.6, top=TOP, width=6.3,
            height=0.59 + 7 * 0.6, col_widths=[2.3, 1.05, 0.8, 2.15], align="LLRL", unit_col=1)
    s.footnote(f("Debt share = senior debt at COD ÷ total uses. Debt draws pro rata with construction costs at up to "
                 "{g1} (European) / {g2} (Chinese) of cost; interest capitalised to COD. Securities are the same on both "
                 "routes; HUF amounts at {fx} HUF/EUR.", g1=b("In_Gearing", pct0), g2=b("In_Gearing_CN", pct0),
                 fx=b("In_FX", n0)))

    # ---- 6 Returns: sale at COD vs hold ---------------------------------------------------------------------------
    s = d.content("Returns: sale at COD vs hold", f(
        "A sale at COD turns €{e1}m of equity into €{n1}m ({m1}); holding earns {h1} at a {d1} DSCR (European)",
        e1=eq[0], n1=b("Out_COD_Proceeds_1", eur), m1=b("Out_COD_MOIC_1", x2), h1=htm[0],
        d1=b("Out_Min_DSCR_1", x2)), f(
        "With Chinese turbines the sale returns {c2} and the hold {h2}; debt is sized at {k1} / {k2} of "
        "construction cost.", c2=cod[1], h2=htm[1], k1=b("Out_Debt_Share_1", pct0),
        k2=b("Out_Debt_Share_2", pct0)), title_lines=2)

    def both_routes(key, fmt, cod_side=True, htm_side=True):
        v = o(key, fmt)
        return [*(v if cod_side else [DASH, DASH]), *(v if htm_side else [DASH, DASH])]

    rows = [["Sale price (enterprise value)", "€ m", *both_routes("COD_Price", eur, htm_side=False)],
            ["Senior debt assumed by the buyer", "€ m", *o("Debt_COD", neg), DASH, DASH],
            ["Net proceeds to FEP (after sale costs)", "€ m", *both_routes("COD_Proceeds", eur, htm_side=False)],
            ["FEP equity: to COD / peak if held", "€ m", *eq, *o("Equity_Peak_HTM", eur)],
            ["FEP profit", "€ m", *o("COD_Profit", eur), *o("HTM_Profit", eur)],
            ["FEP IRR", "%", *cod, *htm],
            ["FEP MOIC", "x", *o("COD_MOIC", x2), *o("HTM_MOIC", x2)],
            ["FEP NPV", "€ m", *o("COD_NPV", eur), *o("HTM_NPV", eur)],
            ["Project IRR, actual tax, before financing", "%", DASH, DASH, *o("Project_IRR", pct)],
            ["Minimum DSCR (sculpted repayment)", "x", DASH, DASH, *o("Min_DSCR", x2)],
            ["Corporate income tax over the life", "€ m", DASH, DASH, *o("Tax_Life", eur)]]
    s.table(["FEP, post-tax", "Unit", "COD – European", "COD – Chinese", "Hold – European", "Hold – Chinese"], rows,
            key_rows={5}, top=TOP, width=10.3, height=0.59 + 11 * 0.55, col_widths=[3.7, 0.8, 1.45, 1.45, 1.45, 1.45])
    nums = {r: [d.bind(f"Out_{k}_{i}") for i in (1, 2)] for r, k in (("cod", "COD_IRR"), ("htm", "HTM_IRR"))}
    s.chart("column", ["Sale at COD", "Hold to maturity"],
            {"European turbine": [nums["cod"][0], nums["htm"][0]], "Chinese turbine": [nums["cod"][1], nums["htm"][1]]},
            left=15.45, top=TOP, width=4.45, height=3.6, title="FEP IRR", number_format="0.0%", name="IRR chart")
    s.card("Debt and tax", [
        ("Debt: ", f("{dscr} target DSCR over {t} years, {r} fixed; sized at the lower of the DSCR capacity and "
                     "{g1} / {g2} of cost; {u1} / {u2} of the DSCR capacity is used.",
                     dscr=b("In_DSCR", x2), t=b("In_Debt_Tenor", n0), r=b("In_Debt_Rate", pct),
                     g1=b("Out_Gearing_Cap_1", pct0), g2=b("Out_Gearing_Cap_2", pct0),
                     u1=b("Out_Capacity_Used_1", pct0), u2=b("Out_Capacity_Used_2", pct0))),
        ("Tax: ", f("CIT {c}; {y}-year depreciation; losses capped at {l} of the base and lapse after {ly} years. "
                    "Gain on the COD share sale {e}.", c=b("In_CIT", pct0), y=b("In_Dep_Years", n0),
                    l=b("In_Loss_Cap", pct0), ly=b("In_Loss_Years", n0),
                    e=b("In_Sale_Gain_Taxed", lambda v: "exempt" if v == 0 else "taxed")))],
        left=15.45, top=6.25, width=4.45, height=3.9)
    s.footnote(f("Sale price {pe} / {pc} k€ per MW; the buyer takes the SPV with its debt and funds the DSRA. Hold: PPA {ppa} €/MWh, {life}-year life; DSRA of {dsra} months "
                 "funded at COD. HIPA {hipa} and the innovation contribution on revenue. NPV at {r}.",
                 pe=b("In_COD_Price", n0), pc=b("In_COD_Price_CN", n0), ppa=b("In_PPA_Price", n0),
                 life=b("In_Life", n0), dsra=b("In_DSRA_Months", n0), hipa=b("In_HIPA", pct0),
                 r=b("In_Discount_Rate", pct0)))

    # ---- 7 Sensitivities ------------------------------------------------------------------------------------------
    def st(key, k, fmt):
        return d.stored(key, k, fmt)

    s = d.content("Sensitivities", f(
        "A lower COD price is the largest risk to the sale ({c}); yield and PPA price drive the hold "
        "({y} / {pp})", c=st("COD_IRR_1", sets.index("COD price −10%") + 1, pct),
        y=st("HTM_IRR_1", sets.index("Yield −10%") + 1, pct), pp=st("HTM_IRR_1", sets.index("PPA −10%") + 1, pct)),
        "FEP IRR by key-assumption set, each change on its own; the sale at COD does not depend on yield or PPA price "
        "because its price is fixed per MW.", title_lines=2)
    rows = [[name, *[st(k, i, x2 if "DSCR" in k else pct) for k in
                     ("COD_IRR_1", "COD_IRR_2", "HTM_IRR_1", "HTM_IRR_2", "Min_DSCR_1", "Min_DSCR_2")]]
            for i, name in enumerate(sets, 1)]
    s.table(["Key-assumption set", "COD – EU", "COD – CN", "Hold – EU", "Hold – CN", "DSCR – EU", "DSCR – CN"],
            rows, key_rows={0}, top=TOP, width=10.3, height=0.59 + 8 * 0.6,
            col_widths=[2.8, 1.25, 1.25, 1.25, 1.25, 1.25, 1.25], unit_col=None)
    s.card("What to watch", [
        ("Turbine choice: ", "the permits name Goldwind; European turbines need a permit amendment and carry the "
                             "most capex risk."),
        ("COD price: ", "an estimate per MW, not a bid; test it with buyers before NTP."),
        ("PPA: ", f("{p} €/MWh for the whole life is an assumption; no offtake is signed.",
                    p=b("In_PPA_Price", n0))),
        ("Debt: ", f("a {t}-year tenor makes the DSCR bind and cuts the European hold IRR to {h}.",
                     t=b(f"Inputs!O{next(names['In_Debt_Tenor'].destinations)[1].split('$')[-1]}", n0),
                     h=st("HTM_IRR_1", sets.index("Tenor 10 years") + 1, pct)))],
        left=15.45, top=TOP, width=4.45, height=6.6)
    s.footnote(f("Each set changes one input, as named; tenor {t} years. Stored by the scenario macro with the tax "
                 "solver; base = live model. Model checks in every set: {ck}.",
                 t=b(f"Inputs!O{next(names['In_Debt_Tenor'].destinations)[1].split('$')[-1]}", n0),
                 ck=d.bind_sum([st(f"Checks_{r}", k, None) for r in (1, 2) for k in range(1, len(sets) + 1)],
                               lambda v: "all pass" if v == 0 else "FAIL")))
    return d.save(out / "Bana_IC.pptx")
