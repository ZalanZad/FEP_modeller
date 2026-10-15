"""Bana IC deck – 7 slides to SPEC.md §9, streamlined from reference/Bana_IC_20260109_vDSz.pptx.

Every figure is `Deck.bind` on the packaged, LibreOffice-recalculated workbook. The Dentons DD risk table and the
deal narrative are qualitative and come from the original deck (the table is read from it at build time).
"""
from pathlib import Path

from pptx import Presentation
from pptx.util import Pt

from fepm.ppt.deck import DASH, ROOT, Deck, fmt_eur_m, fmt_mult, fmt_num, fmt_pct

SOURCE_DECK = ROOT / "reference" / "Bana_IC_20260109_vDSz.pptx"
SECTIONS = ["Executive summary", "Risks and mitigants", "Technical overview and tender bid", "Funding to COD", "Bank feedback",
            "Returns: sale at COD vs hold", "Sale price and buyer return", "Promise or build",
            "Sensitivities"]
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


def pt(v):                       # tender points: 59.5, 12
    return f"{v:.0f}" if float(v).is_integer() else f"{v:.1f}"


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
    d.source = f"{Path(xlsm).stem}_{d.date:%Y%m%d}{Path(xlsm).suffix}"   # the dated copy that is sent
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
        return lever(key, i, pt, grid="IJ"[i - 1])

    sets = [c.value for row in d.wb["Inputs"][next(names["Scenario_Names"].destinations)[1]] for c in row]

    def set_cell(name, set_name):  # an input's value in one key-assumption set (Inputs column H = set 1)
        row = next(names[name].destinations)[1].split("$")[-1]
        return f"Inputs!{chr(ord('H') + sets.index(set_name))}{row}"

    def st(key, k, fmt):         # stored result of set k (1-based)
        return d.stored(key, k, fmt)

    # ---- 1 Cover -------------------------------------------------------------------------------------------
    d.cover("Project Bana", f("{mw} MW wind development project opportunity", mw=b("Out_Wind_Built_1", n0)),
            b("Dashboard!E6", lambda v: f"{v.day} {v:%B %Y}"))

    # ---- 2 Executive summary -----------------------------------------------------------------------------------
    cod, htm, eq, p = o("COD_IRR", pct), o("HTM_IRR", pct), o("Equity_COD", eur), o("Points", pt)
    s = d.content("Executive summary", f(
        "At €{pp}m, Bana returns {a} (European) / {c} (Chinese) on a sale at COD and {h} / {h2} if held",
        pp=b("In_Purchase_Price", k_to_m), a=cod[0], c=cod[1], h=htm[0], h2=htm[1]), f(
        "FEP funds €{e1}m of equity to COD (€{e2}m Chinese) at {g} gearing and bids {p1} points ({p2} Chinese) into "
        "the {bd} tender; the European route needs a permit amendment.", e1=eq[0], e2=eq[1],
        g=b("In_Gearing", pct0), p1=p[0], p2=p[1], bd=b("In_Bid_Date", day)), title_lines=2)
    s.text("Narrative", f("\n".join([
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
        "FEP asks the IC to approve the purchase price and to give its view on buying and developing the project.",
        "",
        "We are assuming a COD sale price of €{pe}m/MW with EU turbines and €{pc}m/MW with Chinese turbines"]),
        pe=b("In_COD_Price", lambda v: f"{v / 1000:.2f}"), pc=b("In_COD_Price_CN", lambda v: f"{v / 1000:.2f}")),
        4.92, TOP, 7.6, 6.29, size=16)
    rows = [["Tender points", "points", *p],
            ["Senior debt at COD", "€ m", *o("Debt_COD", eur)],
            ["FEP equity to COD", "€ m", *eq],
            ["Sale at COD", "", "", ""],
            ["FEP profit", "€ m", *o("COD_Profit", eur)],
            ["FEP IRR", "%", *cod],
            ["FEP MOIC", "x", *o("COD_MOIC", x2)],
            ["Buyer IRR, unlevered", "%", *o("Buyer_IRR_U", pct)],
            ["Hold to maturity", "", "", ""],
            ["FEP IRR", "%", *htm],
            ["FEP MOIC", "x", *o("HTM_MOIC", x2)],
            ["FEP NPV", "€ m", *o("HTM_NPV", eur)],
            ["Minimum DSCR", "x", *o("Min_DSCR", x2)]]
    s.table(["Key numbers", "Unit", *ROUTES], rows, sections={3, 8}, key_rows={5, 9}, left=12.85, top=TOP,
            width=7.05, height=0.59 + 13 * 0.5, col_widths=[3.05, 0.9, 1.55, 1.55])
    s.footnote(f("Post-tax, FEP 100 %. Purchase price paid at closing ({cl}), after the award. PPA {ppa} €/MWh for "
                 "the whole life; NPV at {r}. Chinese = the permitted Goldwind turbines. Debt: {g} of "
                 "construction cost, {t}-year annuity; bank feedback on slide 6.",
                 cl=b("In_Closing", date), ppa=b("In_PPA_Price", n0), r=b("In_Discount_Rate", pct0),
                 g=b("In_Gearing", pct0), t=b("In_Debt_Tenor", n0)))

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
        mw=b("Out_Wind_Built_1", n0), g=b("Out_Gen_Year1_1", n1), p1=p[0], p2=p[1]), f(
        "Overbuild {o} ({op} points, {w} of {c} MW built, inside the tolerance) and BESS {bs} of MVA ({bp} points); "
        "European turbines add {t} points but need a permit amendment.", o=lever("Overbuild", 1, pct0),
        c=b("Out_Wind_MW_1", n0), op=pts("Overbuild", 1), w=b("Out_Wind_Built_1", n0), bs=lever("BESS", 1, pct0),
        bp=pts("BESS", 1), t=pts("EU_Turbine", 1)),
        title_lines=2)
    rows = [["Site and grid", ""],
            ["Grid node", "Ács Katódgyár (BAMO B 132.00)"],
            ["Connection / wind committed / built", f("{a} MVA / {c} MW / {w} MW", a=b("In_MVA", n0),
                                                      c=b("Out_Wind_MW_1", n0), w=b("Out_Wind_Built_1", n0))],
            ["BESS committed and built", f("{m} MW / {h} MWh at {c} k€/MWh", m=b("Out_BESS_MW_1", n0),
                                           h=b("Out_BESS_MWh_1", n0), c=b("In_BESS_Capex", n0))],
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
    s.table(["Item", "Assumption"], rows, sections={0, 4, 9, 12}, top=TOP, width=8.1, height=0.59 + 14 * 0.52,
            col_widths=[3.9, 4.2], align="LL", unit_col=None)
    rows = [["ÖSZ-1 municipal one-off", f("{a} k€/MW (€{m}m)", a=lever("Municipal", 1, n0),
                                          m=b("Out_Uses_Municipal_1", lambda v: fmt_eur_m(v, 2))),
             pts("Municipal", 1), pts("Municipal", 2)],
            ["ÖSZ-2 community contribution", f("{a} €/MWh", a=lever("Community", 1, dp1)),
             pts("Community", 1), pts("Community", 2)],
            ["ÖSZ-3 overbuild", f("{a} of MVA; {w} built", a=lever("Overbuild", 1, pct0),
                                  w=b("In_Overbuild_Built", pct0)), pts("Overbuild", 1), pts("Overbuild", 2)],
            ["ÖSZ-4 BESS power", f("{a} of MVA, {h} h; built", a=lever("BESS", 1, pct0),
                                   h=lever("BESS_Hours", 1, n0)), pts("BESS", 1), pts("BESS", 2)],
            ["ÖSZ-6 performance security", f("{a} of the minimum", a=lever("PB_Mult", 1, x1)),
             pts("PB_Mult", 1), pts("PB_Mult", 2)],
            ["ÖSZ-9 European turbines ≥ 80 %", "European route only", pts("EU_Turbine", 1), pts("EU_Turbine", 2)],
            ["ÖSZ-5 partial award, ÖSZ-8 shared connection, ÖSZ-9 EU transformers", "not offered",
             f("{a} / {b} / {c}", a=pts("Partial", 1), b=pts("Shared", 1), c=pts("EU_Transformer", 1)),
             f("{a} / {b} / {c}", a=pts("Partial", 2), b=pts("Shared", 2), c=pts("EU_Transformer", 2))],
            ["Total score", "", *p]]
    s.table(["Tender bid", "Commitment", "European", "Chinese"], rows, key_rows={7}, left=13.35, top=TOP,
            width=6.55, height=0.59 + 8 * 0.6, col_widths=[2.35, 1.9, 1.15, 1.15], align="LLRR", unit_col=None)
    s.footnote(f("Points are the model's own score on the MEKH 2026 bands (Kiírás Annex 3-4). ÖSZ-4 is scored on net "
                 "MVA; BESS at capex only, no revenue or opex, not in the COD price. Brownfield (ÖSZ-7) is "
                 "not available. Build shortfall {sh} against the commitment, inside the "
                 "{t} tolerance (GYIK 28/30). The municipal one-off is paid at COD.",
                 sh=b("Out_Shortfall_1", pct), t=b("In_Tolerance_Cap", pct0)))

    # ---- 5 Funding to COD ------------------------------------------------------------------------------------------
    debt, uses = o("Debt_COD", eur), o("Uses_Total", eur)
    s = d.content("Funding to COD", f(
        "Uses to COD of €{u1}m European (€{u2}m Chinese); senior debt funds {g1} ({g2}), FEP the rest", u1=uses[0], u2=uses[1], g1=b("Out_Gearing_Uses_1", pct0), g2=b("Out_Gearing_Uses_2", pct0)), f(
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
    s.footnote(f("Debt share = senior debt at COD ÷ total uses. Debt = {g} of construction cost on both routes "
                 "(LTC), drawn pro rata with construction costs; interest capitalised to COD. HUF amounts at {fx} "
                 "HUF/EUR.", g=b("In_Gearing", pct0), fx=b("In_FX", n0)))

    # ---- 6 Bank feedback (indicative, verbal) and the model on each bank's terms -----------------------------------
    kr, ku = sets.index("Raiffeisen terms") + 1, sets.index("UniCredit terms") + 1
    s = d.content("Bank feedback", f(
        "Hold IRR {h1} / {h2} at {g} gearing; {r1} / {r2} on Raiffeisen and {u1} / {u2} on UniCredit terms", r1=st("HTM_IRR_1", kr, pct), r2=st("HTM_IRR_2", kr, pct),
        u1=st("HTM_IRR_1", ku, pct), u2=st("HTM_IRR_2", ku, pct), h1=htm[0], h2=htm[1], g=b("In_Gearing", pct0)),
        "Two bank meetings on the 50 MW wind + 10 MW BESS project; both indicative and verbal, with no commitment. "
        "The base case keeps 55 % gearing.", title_lines=2)
    rows = [["Debt amount", "~€0.9m/MW (±10 %) on a ~€1.7m/MW wind capex estimate (ours is slightly higher); "
             "~48–58 % LTC", "LTC lower than Raiffeisen, due to the shorter tenor and higher DSCR"],
            ["Repayment", "15-year annuity; significant cash sweep, up to full", "~12-year annuity"],
            ["Minimum DSCR", "1.20x on PPA revenue, 1.40x on merchant", "1.30x on PPA revenue, 1.65x on merchant"],
            ["Pricing and prices", "Merchant at the average of the Aurora low and central curves",
             "EUR loan, margin above 2.3 %"],
            ["Uncertainties", "Neither bank has financed a project like this and output starts late, so terms may "
             "move", "Chinese turbines not excluded but new to both; UniCredit finances one in Romania"]]
    s.table(["What the banks said", "Raiffeisen", "UniCredit"], rows, top=TOP, width=15.0,
            height=0.59 + 5 * 0.6, col_widths=[2.6, 6.2, 6.2], align="LLL", unit_col=None)
    rows = []
    for k, label, terms in (
            (1, "Base: 55 % gearing", f("{g} of cost, {t}-year annuity", g=b("In_Gearing", pct0),
                                        t=b("In_Debt_Tenor", n0))),
            (kr, "Raiffeisen terms", f("€{q}m/MW, {t}-year, {x}", q=b(set_cell("In_Debt_Per_MW", "Raiffeisen terms"),
                                                                    lambda v: f"{v / 1000:.2f}"),
                                       t=b("In_Debt_Tenor", n0), x=b("In_DSCR", x2))),   # = base tenor, DSCR
            (ku, "UniCredit terms", f("€{q}m/MW*, {t}-year, {x}", q=b(set_cell("In_Debt_Per_MW", "UniCredit terms"),
                                                                    lambda v: f"{v / 1000:.2f}"),
                                      t=b(set_cell("In_Debt_Tenor", "UniCredit terms"), n0),
                                      x=b(set_cell("In_DSCR", "UniCredit terms"), x2)))):
        rows.append([label, terms, st("Debt_COD_1", k, eur), st("Debt_COD_2", k, eur), st("COD_IRR_1", k, pct),
                     st("COD_IRR_2", k, pct), st("HTM_IRR_1", k, pct), st("HTM_IRR_2", k, pct),
                     st("Min_DSCR_1", k, x2), st("Min_DSCR_2", k, x2)])
    s.table(["The model on each basis", "Debt terms", "Debt EU € m", "Debt CN € m", "Sale IRR EU", "Sale IRR CN",
             "Hold IRR EU", "Hold IRR CN", "DSCR EU", "DSCR CN"], rows, key_rows={0}, top=6.35, width=15.0,
            height=0.59 + 3 * 0.55, col_widths=[2.6, 3.2] + [1.15] * 8, align="LLRRRRRRRR", unit_col=None)
    s.footnote("Bank terms as reported in the FEP notes of the Raiffeisen and UniCredit meetings (October 2026). "
               "*UniCredit gave no amount: Raiffeisen's quantum scaled to its tenor and DSCR. Bank cases: 70 % LTC "
               "cap, same 5 % rate; stored key-assumption sets.")

    # ---- 6 Returns: sale at COD vs hold ---------------------------------------------------------------------------
    s = d.content("Returns: sale at COD vs hold", f(
        "A sale at COD turns €{e1}m of equity into €{n1}m ({m1}); holding earns {h1} at a {d1} DSCR (European)",
        e1=eq[0], n1=b("Out_COD_Proceeds_1", eur), m1=b("Out_COD_MOIC_1", x2), h1=htm[0],
        d1=b("Out_Min_DSCR_1", x2)), f(
        "With Chinese turbines the sale returns {c2} and the hold {h2}; debt is {g} of construction cost, a "
        "{t}-year annuity.", c2=cod[1], h2=htm[1], g=b("In_Gearing", pct0), t=b("In_Debt_Tenor", n0)),
        title_lines=2)

    def both_routes(key, fmt, cod_side=True, htm_side=True):
        v = o(key, fmt)
        return [*(v if cod_side else [DASH, DASH]), *(v if htm_side else [DASH, DASH])]

    rows = [["Sale price (enterprise value)", "€ m", *both_routes("COD_Price", eur, htm_side=False)],
            ["Senior debt assumed by the buyer", "€ m", *o("Debt_COD", neg), DASH, DASH],
            ["Net proceeds to FEP (after sale costs)", "€ m", *both_routes("COD_Proceeds", eur, htm_side=False)],
            ["Buyer IRR, unlevered / levered", "%", *[f("{u} / {l}", u=u, l=l_) for u, l_ in
                                                     zip(o("Buyer_IRR_U", pct), o("Buyer_IRR_L", pct))], DASH, DASH],
            ["FEP equity: to COD / peak if held", "€ m", *eq, *o("Equity_Peak_HTM", eur)],
            ["FEP profit", "€ m", *o("COD_Profit", eur), *o("HTM_Profit", eur)],
            ["FEP IRR", "%", *cod, *htm],
            ["FEP MOIC", "x", *o("COD_MOIC", x2), *o("HTM_MOIC", x2)],
            ["FEP NPV", "€ m", *o("COD_NPV", eur), *o("HTM_NPV", eur)],
            ["Project IRR, actual tax, before financing", "%", DASH, DASH, *o("Project_IRR", pct)],
            ["Minimum DSCR (annuity)", "x", DASH, DASH, *o("Min_DSCR", x2)],
            ["Corporate income tax over the life", "€ m", DASH, DASH, *o("Tax_Life", eur)]]
    s.table(["FEP, post-tax", "Unit", "COD – European", "COD – Chinese", "Hold – European",
             "Hold – Chinese"], rows, key_rows={6}, top=TOP, width=10.3,
            height=0.59 + 12 * 0.52, col_widths=[3.7, 0.8, 1.45, 1.45, 1.45, 1.45])
    nums = {(r, i): d.bind(f"Out_{k}_{i}") for r, k in (("cod", "COD_IRR"), ("htm", "HTM_IRR")) for i in (1, 2)}
    s.chart("column", ["Sale at COD", "Hold to maturity"],
            {f"{route} turbine": [nums[("cod", i)], nums[("htm", i)]] for i, route in ((1, "European"), (2, "Chinese"))},
            left=15.45, top=TOP, width=4.45, height=3.1, title="FEP IRR", number_format="0.0%", name="IRR chart")
    s.card("Debt and tax", [
        ("Debt: ", f("{g} of construction cost on both routes (€{d1}m / €{d2}m), {t}-year annuity at {r} fixed; "
                     "sizing floor {dscr} DSCR on PPA revenue, {u1} / {u2} of that capacity used.",
                     g=b("In_Gearing", pct0), d1=debt[0], d2=debt[1], t=b("In_Debt_Tenor", n0),
                     r=b("In_Debt_Rate", pct), dscr=b("In_DSCR", x2),
                     u1=b("Out_Capacity_Used_1", pct0), u2=b("Out_Capacity_Used_2", pct0))),
        ("Maximum debt, no cap: ", f("€{a}m at {xl} / €{c}m at {xh} DSCR European ({pa} / {pc} of uses); "
                                    "€{b}m / €{d}m Chinese.", a=b("Out_Debt_Max_Low_1", eur),
                                    c=b("Out_Debt_Max_High_1", eur), xl=b("Out_DSCR_Max_Low_1", x2),
                                    xh=b("Out_DSCR_Max_High_1", x2), pa=b("Out_Debt_Max_Low_Pct_1", pct0),
                                    pc=b("Out_Debt_Max_High_Pct_1", pct0), b=b("Out_Debt_Max_Low_2", eur),
                                    d=b("Out_Debt_Max_High_2", eur))),
        ("Tax: ", f("CIT {c} (9 % + possible Pillar Two top-up); Robin Hood ({r}) only as a sensitivity. {y}-year "
                    "depreciation; losses capped at {l}, lapse after {ly} years. COD share-sale gain {e}.", c=b("In_CIT", pct0), y=b("In_Dep_Years", n0),
                    r=b("In_RH_Rate", pct0),
                    l=b("In_Loss_Cap", pct0), ly=b("In_Loss_Years", n0),
                    e=b("In_Sale_Gain_Taxed", lambda v: "exempt" if v == 0 else "taxed")))],
        left=15.45, top=5.7, width=4.45, height=4.75)
    s.footnote(f("Sale price {pe} / {pc} k€ per built MW, less the {bl} bottleneck loss; the buyer takes the SPV with its debt and funds the DSRA. Hold: PPA {ppa} €/MWh, {life}-year life; DSRA of {dsra} months "
                 "funded at COD. HIPA {hipa} and the innovation contribution on revenue. NPV at {r}.",
                 pe=b("In_COD_Price", n0), pc=b("In_COD_Price_CN", n0), ppa=b("In_PPA_Price", n0),
                 bl=b("Out_Bottleneck_Loss_1", lambda v: fmt_pct(v, 2)),
                 life=b("In_Life", n0), dsra=b("In_DSRA_Months", n0), hipa=b("In_HIPA", pct0),
                 r=b("In_Discount_Rate", pct0)))

    # ---- 7 Sale price and buyer return ---------------------------------------------------------------------------
    cases = ["COD price −10%", "Base", "COD price +10%"]
    lo = sets.index(cases[0]) + 1
    s = d.content("Sale price and buyer return", f(
        "At {pe} / {pc} k€/MW a buyer earns {u1} / {u2}; at the lower price {l1} / {l2}, but FEP's sale IRR falls "
        "to {c1} / {c2}", pe=b("In_COD_Price", n0), pc=b("In_COD_Price_CN", n0),
        u1=b("Out_Buyer_IRR_U_1", pct), u2=b("Out_Buyer_IRR_U_2", pct), l1=st("Buyer_IRR_U_1", lo, pct),
        l2=st("Buyer_IRR_U_2", lo, pct), c1=st("COD_IRR_1", lo, pct), c2=st("COD_IRR_2", lo, pct)), (
        "The buyer receives the same EBITDA on both routes, so the higher European price lowers its return; "
        "the price per MW is an estimate, not a bid."), title_lines=2)
    rows = []
    for name in cases:
        k = sets.index(name) + 1
        rows.append([name.replace("COD price ", "Price "),
                     b(set_cell("In_COD_Price", name), n0), st("Buyer_IRR_U_1", k, pct), st("Buyer_IRR_L_1", k, pct),
                     st("COD_IRR_1", k, pct),
                     b(set_cell("In_COD_Price_CN", name), n0), st("Buyer_IRR_U_2", k, pct), st("Buyer_IRR_L_2", k, pct),
                     st("COD_IRR_2", k, pct)])
    s.table(["Case", "EU price k€/MW", "Buyer unlevered", "Buyer levered", "FEP sale IRR",
             "CN price k€/MW", "Buyer unlevered", "Buyer levered", "FEP sale IRR"], rows, key_rows={1}, top=TOP,
            height=0.59 + 3 * 0.6, col_widths=[2.2] + [1.6] * 8, unit_col=None)
    s.footnote("Buyer unlevered: pays the enterprise value at COD and receives EBITDA less tax, guarantee fees and any "
               "PB call. Buyer levered: pays the price less the senior debt, funds the DSRA and services the debt. "
               "Share deal: SPV tax, debt and depreciation continue; the unlevered IRR uses the SPV's actual tax.")

    # ---- 8 Promise or build -----------------------------------------------------------------------------------------
    kn = sets.index("Promise all, deliver nothing") + 1
    s = d.content("Promise or build", f(
        "Delivering none of the commitments keeps the {p1} / {p2} points and costs the €{pb}m PB; sale IRR "
        "{c1} / {c2}", p1=st("Points_1", kn, pt), p2=st("Points_2", kn, pt),
        pb=st("PB_Call_Applied_1", kn, lambda v: fmt_eur_m(v, 2)), c1=st("COD_IRR_1", kn, pct),
        c2=st("COD_IRR_2", kn, pct)), (
        "Any commitment not delivered triggers the full PB call, so once one is missed the rest can be dropped too: "
        "no overbuild, no BESS, no municipal or community payments, Chinese turbines on the European route."),
        title_lines=2)
    rows = []
    for name, label in (("Base", "Base: deliver the bid"), ("Promise all, deliver nothing", "Promise all, deliver nothing")):
        k = sets.index(name) + 1
        rows.append([label, st("Points_1", k, pt), st("PB_Call_Applied_1", k, lambda v: fmt_eur_m(v, 2)),
                     st("Equity_COD_1", k, eur), st("Equity_COD_2", k, eur), st("COD_IRR_1", k, pct),
                     st("COD_IRR_2", k, pct), st("HTM_IRR_1", k, pct), st("HTM_IRR_2", k, pct),
                     st("Buyer_IRR_U_1", k, pct)])
    s.table(["Case", "Points EU", "PB call € m", "Equity EU € m", "Equity CN € m", "Sale IRR EU", "Sale IRR CN",
             "Hold IRR EU", "Hold IRR CN", "Buyer IRR EU"], rows, key_rows={0}, top=TOP, height=0.59 + 2 * 0.62,
            col_widths=[3.0] + [1.33] * 9, unit_col=None)
    s.card("Before relying on it", [
        ("Group exclusion: ", "a PB call may exclude the group from these tenders for three years."),
        ("Municipality and community: ", "the payments are a commitment to MEKH in the bid, not a contract with "
                                         "them; the PB call is the sanction."),
        ("Buyer: ", "a buyer at COD sees the called PB and the missing commitments in its due diligence."),
        ("Reading: ", "the full call on any missed commitment follows Kiírás XIII.17 and GYIK 72; the pro-rata "
                      "reading (GYIK 28/30) is the base.")],
        left=4.92, top=5.0, width=15.0, height=1.9)
    s.footnote(f("Deliver nothing: wind {w} MW (no overbuild), no BESS, no municipal one-off or community "
                 "contribution; the European route installs Chinese turbines and sells at the Chinese price. The whole "
                 "PB is called {lag} months after COD. Equity = FEP equity to COD.",
                 w=b("In_MVA", n0), lag=b("In_Sanction_Lag", n0)))

    # ---- 9 Sensitivities ------------------------------------------------------------------------------------------

    s = d.content("Sensitivities", f(
        "A lower COD price hits the sale most ({c}); Robin Hood tax ({r}) and lower yield ({y}) hit the hold", c=st("COD_IRR_1", sets.index("COD price −10%") + 1, pct),
        r=st("HTM_IRR_1", sets.index("Robin Hood 31%") + 1, pct), y=st("HTM_IRR_1", sets.index("Yield −10%") + 1, pct)),
        "FEP and buyer (unlevered) IRR by key-assumption set, each change on its own; the sale at COD does not depend on yield or PPA price "
        "because its price is fixed per MW.", title_lines=2)
    rows = [[name, st("Points_1", i, pt), *[st(k, i, x2 if "DSCR" in k else pct) for k in
                     ("COD_IRR_1", "COD_IRR_2", "Buyer_IRR_U_1", "Buyer_IRR_U_2", "HTM_IRR_1", "HTM_IRR_2",
                      "Min_DSCR_1", "Min_DSCR_2")]]
            for i, name in enumerate(sets, 1)]
    s.table(["Key-assumption set", "Points EU", "COD EU", "COD CN", "Buyer EU", "Buyer CN", "Hold EU", "Hold CN",
             "DSCR EU", "DSCR CN"], rows, key_rows={0}, top=TOP, width=10.4, height=0.59 + len(sets) * 0.42,
            col_widths=[2.0, 0.8] + [0.95] * 8, unit_col=None)
    s.card("What to watch", [
        ("Turbine choice: ", "the permits name Goldwind; European turbines need a permit amendment and carry the "
                             "most capex risk."),
        ("BESS 100 MWh: ", f("{p} points but COD IRR {c} and hold IRR {h} (European).",
                             p=st("Points_1", sets.index("BESS 100 MWh") + 1, pt),
                             c=st("COD_IRR_1", sets.index("BESS 100 MWh") + 1, pct),
                             h=st("HTM_IRR_1", sets.index("BESS 100 MWh") + 1, pct))),
        ("Robin Hood tax: ", f("not in the base; if the {r} levy applies (its exemption covers only KÁT/METÁR "
                               "plants), the hold IRR falls to {h} / {h2}.", r=b("In_RH_Rate", pct0),
                               h=st("HTM_IRR_1", sets.index("Robin Hood 31%") + 1, pct),
                               h2=st("HTM_IRR_2", sets.index("Robin Hood 31%") + 1, pct))),
        ("PPA: ", f("{p} €/MWh for the whole life is an assumption; no offtake is signed.",
                    p=b("In_PPA_Price", n0))),
        ("Debt: ", f("a {t}-year tenor lowers the European minimum DSCR to {ds} and the hold IRR to {h}.",
                     ds=st("Min_DSCR_1", sets.index("Tenor 10 years") + 1, x2),
                     t=b(set_cell("In_Debt_Tenor", "Tenor 10 years"), n0),
                     h=st("HTM_IRR_1", sets.index("Tenor 10 years") + 1, pct))),
        ("Cash sweep: ", f("Raiffeisen asks for a significant one, up to full; a full sweep cuts the hold IRR to "
                           "{s1} / {s2}.",
                           s1=st("HTM_IRR_1", sets.index("Full cash sweep") + 1, pct),
                           s2=st("HTM_IRR_2", sets.index("Full cash sweep") + 1, pct)))],
        left=15.45, top=TOP, width=4.45, height=6.4)
    s.footnote(f("Each set changes only the inputs its name describes, from the base ({t}-year tenor). Stored "
                 "by the scenario macro with the tax solver; base = live model. Model checks in every set: {ck}.",
                 t=b("In_Debt_Tenor", n0),
                 ck=d.bind_sum([st(f"Checks_{r}", k, None) for r in (1, 2) for k in range(1, len(sets) + 1)],
                               lambda v: "all pass" if v == 0 else "FAIL")))
    return d.save(out / "Bana_IC.pptx")
