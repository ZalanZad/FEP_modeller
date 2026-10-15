"""Bora IC deck – 8 slides to SPEC.md §0 items 19, 21, 22 and 23 (one strategy for both bids), text from
out/Bora_Bid_slide_text.md.

Every figure is `Deck.bind` on the packaged workbook (LibreOffice-recalculated when `package` saves it, so it holds
the same values as the audit's `_recalc.xlsx`). Composite texts use `Deck.f`, which records each bound part.
"""
from pathlib import Path

from fepm.ppt.deck import DASH, Deck, fmt_delta, fmt_eur_m, fmt_mult, fmt_num, fmt_pct

SECTIONS = ["Executive summary", "Commitments", "Economics – RtB exit", "Funding and capex",
            "Economics – COD and HTM", "Risks and mitigants"]
TOP = 2.45                      # first free line under a two-line title and subtitle
CARD_L, CARD_W = 15.30, 4.62    # right-hand card; a table beside it is CARD_L - 0.4 - LEFT wide
SIDE_W = CARD_L - 0.40 - 4.92


# ---- display formats of the slide text ("Formats" line of Bora_Bid_slide_text.md) -------------------------
def pts(v):
    return f"{v:.0f}" if float(v).is_integer() else f"{v:.2f}"


def pp(v):                       # signed percentage points: −1.2
    return fmt_num(v * 100, 1)


def pp_abs(v):
    return f"{abs(v) * 100:.1f}"


def eur(v):
    return fmt_eur_m(v)


def eur_abs(v):
    return fmt_eur_m(abs(v))


def eur3(v):
    return f"{v:.3f}"


def date(v):                     # non-breaking hyphens: a date never wraps across two lines
    return f"{v:%d-%b-%Y}".replace("-", "\u2011")


def pct(v):
    return fmt_pct(v)


def pct0(v):
    return fmt_pct(v, 0)


def x1(v):
    return fmt_mult(v, 1)


def n0(v):
    return fmt_num(v, 0)


def n1(v):
    return fmt_num(v, 1)


def dp2(v):
    return f"{v:.2f}"


def build_deck(xlsm: Path, out: Path) -> Path:
    d = Deck(xlsm, source=Path(xlsm).name, sections=SECTIONS, date_cell=("Dashboard", "E6"))
    f = d.f
    names = d.wb.defined_names
    ws = d.wb["Dashboard"]

    def b(ref, fmt):
        """Bind a workbook name (`Out_…`, `In_…`) or a `Sheet!Cell` reference."""
        return d.bind(*ref.split("!"), fmt=fmt) if "!" in ref else d.bind(ref, fmt=fmt)

    def block_head(row):         # header row of the Dashboard block holding `row` (column A = "x")
        return next(r for r in range(row, 0, -1) if ws[f"A{r}"].value == "x")

    def at(label, header="BORA 1", block=None):
        """Dashboard cell found by its row label (column C) and its block's column header: row moves never break."""
        rows = [c.row for c in ws["C"] if c.value == label
                and (block is None or str(ws[f"C{block_head(c.row)}"].value).startswith(block))]
        if len(rows) != 1:
            raise KeyError(f"Dashboard label {label!r} ({block}): {len(rows)} rows")
        col = next(c.column_letter for c in ws[block_head(rows[0])] if c.value == header)
        return f"Dashboard!{col}{rows[0]}"

    def lab(label, fmt, header="BORA 1", block=None):
        return b(at(label, header, block), fmt)

    def per(key, fmt):           # per bid: the BORA 1 cell (BORA 2 is identical, SPEC §0 item 21)
        return b(f"Out_{key}_1", fmt)

    def both(key, fmt):          # both bids: Dashboard Total (column H) where the block has one, else B1 + B2
        sheet, cell = next(names[f"Out_{key}_1"].destinations)
        row = int(cell.split("$")[-1])
        if ws[f"H{block_head(row)}"].value == "Total":
            return b(f"{sheet}!H{row}", fmt)
        return d.bind_sum([f"Out_{key}_1", f"Out_{key}_2"], fmt)

    def both_lab(label, fmt):    # both bids of a Securities and fees line: BORA 1 + BORA 2 cells
        return d.bind_sum([at(label, "BORA 1", "Securities"), at(label, "BORA 2", "Securities")], fmt)

    def zero(label, fmt):        # zero bid (Returns at a glance)
        return lab(label, fmt, "Zero bid", "Returns")

    def delta(label, fmt):       # Δ per bid vs the zero bid (Returns at a glance)
        return lab(label, fmt, "Δ B1", "Returns")

    def bid(label, header, fmt): # the bid block (commitments, points, impact vs the zero bid), BORA 1 columns
        return lab(label, fmt, header, "The bid")

    sets = [c.value for row in d.wb["Inputs"][next(names["Scenario_Names"].destinations)[1]] for c in row]

    def stored(key, set_name, fmt):   # BORA 1 result stored for one key-assumption set
        return d.stored(key, sets.index(set_name) + 1, fmt)

    curves = d.wb["Curves"]
    pb_row = next(c.row for c in curves["C"] if c.value == "ÖSZ-6 performance security multiple – lower bound")
    pb_top = [c for c in curves[pb_row] if isinstance(c.value, (int, float))][-1].coordinate   # top band: 2.0x
    pb2 = lambda: b(f"Curves!{pb_top}", x1)  # noqa: E731

    MUN, COM, OVB = "Municipal one-off (ÖSZ-1)", "Community contribution (ÖSZ-2)", "Overbuild (ÖSZ-3)"
    BESS, PBM = "BESS power, share of MVA (ÖSZ-4)", "Performance security multiple (ÖSZ-6)"
    PROFIT, NPV_COD, NPV_HTM = "FEP profit – RtB route", "FEP NPV – sale at COD", "FEP NPV – hold to maturity"
    IRR_RTB, IRR_COD, IRR_HTM = "FEP IRR – RtB route", "FEP IRR – sale at COD", "FEP IRR – hold to maturity"

    # ---- 1 Cover -------------------------------------------------------------------------------------------
    d.cover("Project Bora", "BORA 1 and BORA 2 capacity bids", b("Dashboard!E6", date),
            subtitle=f("Approval of the tender commitments,\n2 × {m} MVA, Szombathely OVIT", m=b("In_MVA", n0)))

    # ---- 2 Decision and timetable ----------------------------------------------------------------------------
    s = d.content("Executive summary", f(
        "Submit both bids with the same commitments at {p} points each: FEP earns €{r}m at the RtB exit, a {i} IRR",
        p=per("Points", pts), r=both("RtB_Profit_Buildup", eur), i=per("RtB_IRR", pct)), f(
        "Against the zero bid, the commitments cost €{a}m of RtB profit and {c} pp of HTM IRR per bid, "
        "and add €{e}m per bid to the all-in commitment.",
        a=delta(PROFIT, eur_abs), c=delta(IRR_HTM, pp_abs), e=delta("Total commitment", eur_abs)), title_lines=2)
    rows = [["Points (EU turbine)", "points", zero("Points, live turbine route", pts), per("Points", pts), DASH],
            [PROFIT, "€ m", zero(PROFIT, eur), per("RtB_Profit", eur), both("RtB_Profit_Buildup", eur)],
            [IRR_RTB, "%", zero(IRR_RTB, pct), per("RtB_IRR", pct), DASH],
            [IRR_COD, "%", zero(IRR_COD, pct), per("COD_IRR", pct), DASH],
            [IRR_HTM, "%", zero(IRR_HTM, pct), per("HTM_IRR", pct), DASH],
            ["All-in commitment if built (securities, development, construction, financing)", "€ m",
             zero("Total commitment", eur), per("Commitment_Total", eur), both("Commitment_Total", eur)]]
    s.table(["Item", "Unit", "Zero bid, per bid", "Per bid", "Both bids"], rows, key_rows={5}, top=TOP,
            width=SIDE_W, height=0.59 + 6 * 0.52, col_widths=[4.3, 1.1, 1.5, 1.5, 1.5])
    s.card("Approval requested and timetable", [
        (f("Submit BORA 1 and BORA 2 by {d}, each with: ", d=b("In_Bid_Date", date)), f(
            "municipal {m} k€/MW; community {c} €/MWh; overbuild {o} (not built); BESS {e} of MVA, {h} h; "
            "PB at the {pb} minimum; partial award, shared connection, EU turbines and transformers.",
            m=bid(MUN, "BORA 1", n0), c=bid(COM, "BORA 1", dp2), o=bid(OVB, "BORA 1", pct0),
            e=bid(BESS, "BORA 1", pct0), h=bid("BESS duration", "BORA 1", n0), pb=bid(PBM, "BORA 1", x1))),
        ("Day one, at submission (both bids): ", f("participation fees €{fee}m in cash and bid guarantees of €{g}m.",
                                                   fee=both_lab("Participation fee", eur3),
                                                   g=both_lab("Bid security", eur))),
        (f("After award ({d}), both bids: ", d=b("In_Award_Date", date)), f(
            "PB guarantees of €{p}m by {pd}, replacing the bid guarantees; TSO cash security of €{t}m by {td}.",
            p=both_lab("Performance security", eur), pd=lab("Performance security", date, "Posted", "Securities"),
            t=both_lab("TSO connection-request security", eur),
            td=lab("TSO connection-request security", date, "Posted", "Securities")))],
        left=CARD_L, top=TOP, width=CARD_W, height=4.1)
    s.footnote(f("Post-tax, FEP share, EU turbine, base case (overbuild not built, BESS built, PB lost pro rata to the "
                 "shortfall); RtB exit {x}. Zero bid = every commitment at its zero level. The two bids are identical, "
                 "so IRRs are the same per bid and for both.", x=b("Dash_Exit_Date", date)))

    # ---- 3 Why these commitments -----------------------------------------------------------------------------
    s = d.content("Commitments", f(
        "Municipal buys {p} points for {a} pp of HTM IRR; BESS ({e} pp) and community ({c} pp) carry the return cost",
        p=bid(MUN, "Points B1", pts), a=bid(MUN, "Δ HTM IRR B1", pp_abs), e=bid(BESS, "Δ HTM IRR B1", pp_abs),
        c=bid(COM, "Δ HTM IRR B1", pp_abs)), f(
        "Each commitment is measured alone against the zero bid, per bid; the PB stays at the minimum because {y} "
        "adds 5 points for one more PB of guarantee.", y=pb2()), title_lines=2)

    def impact(label):           # points, Δ RtB profit, Δ COD IRR, Δ HTM IRR, Δ deposits (BORA 1)
        return [bid(label, h, fm) for h, fm in (("Points B1", pts), ("Δ RtB profit €m B1", fmt_delta),
                                                ("Δ COD IRR B1", pp), ("Δ HTM IRR B1", pp),
                                                ("Δ securities €m B1", fmt_delta))]

    yes_rows = ["Partial award accepted (ÖSZ-5)", "Brownfield site (ÖSZ-7)", "Shared connection (ÖSZ-8)",
                "EU turbines ≥ 80 % (ÖSZ-9)", "EU transformers ≥ 80 % (ÖSZ-9)"]

    def zero_block(header, fm):  # a block of Δ cells that are all zero, shown once
        r0, r1 = (int(at(l, header, "The bid").split("!")[1][1:]) for l in (yes_rows[0], yes_rows[-1]))
        col = at(yes_rows[0], header, "The bid").split("!")[1][0]
        return b(f"Dashboard!{col}{r0}:{col}{r1}", fm)

    rows = [
        [MUN, f("{a} k€/MW", a=bid(MUN, "BORA 1", n0)), *impact(MUN),
         "Paid at COD, after the RtB exit: no FEP cash on the RtB route."],
        ["Community contribution (ÖSZ-2)", f("{a} €/MWh", a=bid(COM, "BORA 1", dp2)), *impact(COM),
         "Operating cost after COD."],
        [OVB, bid(OVB, "BORA 1", pct0), *impact(OVB),
         "Not built: the cost is the larger PB, lost pro rata to the shortfall."],
        [f("BESS power (ÖSZ-4), {h} h", h=bid("BESS duration", "BORA 1", n0)),
         f("{a} of MVA", a=bid(BESS, "BORA 1", pct0)), *impact(BESS),
         "Sized with the overbuild and built; valued at cost, with no BESS revenue."],
        ["PB multiple (ÖSZ-6)", bid(PBM, "BORA 1", x1), *impact(PBM),
         f("Not raised: {y} adds 5 points for one more PB of €{a}m of guarantee.", y=pb2(),
           a=lab("Performance security", eur, "BORA 1", "Securities"))],
        ["Partial award (ÖSZ-5), shared connection (ÖSZ-8), EU turbines / transformers (ÖSZ-9)", "yes",
         f("{a} / {b} / {c} / {e}", a=bid(yes_rows[0], "Points B1", pts), b=bid(yes_rows[2], "Points B1", pts),
           c=bid(yes_rows[3], "Points B1", pts), e=bid(yes_rows[4], "Points B1", pts)),
         *[zero_block(h, fm) for h, fm in (("Δ RtB profit €m B1", fmt_delta), ("Δ COD IRR B1", pp),
                                           ("Δ HTM IRR B1", pp), ("Δ securities €m B1", fmt_delta))],
         f("No cash or return impact. EU sourcing is in the base cost (turbine {t} k€/MW).", t=b("In_Turbine", n0))],
        ["Full bid vs zero bid", "", per("Points", pts), delta(PROFIT, fmt_delta), delta(IRR_COD, pp),
         delta(IRR_HTM, pp), DASH, "The rows do not add to the total because the commitments interact."],
    ]
    s.table(["Commitment (per bid)", "Offered", "Points", "Δ RtB profit € m", "Δ COD IRR pp", "Δ HTM IRR pp",
             "Δ deposits and guarantees € m", "Logic"], rows, key_rows={6}, top=TOP, height=0.59 + 7 * 0.62,
            col_widths=[3.2, 1.5, 1.3, 1.3, 1.25, 1.25, 1.75, 3.45], align="LRRRRRRL", unit_col=None)
    s.footnote(f("Δ = zero bid plus that commitment alone, minus the zero bid. Zero levels: municipal 0, "
                 "community 0, overbuild 0%, BESS 0%, PB 1.0x, yes/no items no. Δ deposits and guarantees = extra "
                 "security posted, not a PB loss. ÖSZ-6: 5 points at {y} (Kiírás X.4 bf, Annex 4). Brownfield (ÖSZ-7) "
                 "is not offered.", y=pb2()))

    # ---- 4 RtB exit: the devex-file waterfall (paid amounts) -------------------------------------------------
    prem_row = next(names["In_RtB_Premium"].destinations)[1].split("$")[-1]
    prem_col = next(names["Scenario_Names"].destinations)[1].split(":")[0].split("$")[1]
    low_col = chr(ord(prem_col) + sets.index("Low premium"))
    s = d.content("Economics – RtB exit", f(
        "At the RtB exit FEP earns €{p}m on both bids, a {m} MOIC; its interest claims are paid ahead of the "
        "{s}/30 split", p=both("RtB_Profit_Buildup", eur), m=per("RtB_MOIC", fmt_mult),
        s=b("In_FEP_Share", lambda v: f"{v * 100:.0f}")), f(
        "FEP keeps {fe} k€ and A.RES {ar} k€ per installed wind MW; at a {lp} k€/MVA RtB price, FEP's "
        "profit falls to €{lo}m for both bids.",
        fe=b("Out_Exit_FEP_Net_MW_1", n0), ar=b("Out_Exit_ARES_Total_MW_1", n0),
        lp=b(f"Inputs!{low_col}{prem_row}", n0),
        lo=d.bind_sum([stored("RtB_Profit_1", "Low premium", None), stored("RtB_Profit_2", "Low premium", None)],
                      eur)), title_lines=2)
    rows = [["RtB price", "€ m", per("RtB_Price", eur), both("RtB_Price", eur)],
            ["Less sale costs", "€ m", per("WF_Sale_Costs", eur), both("WF_Sale_Costs", eur)],
            [f("Less FEP interest claims ({a} deposits, {r} / {m} devex, {g} guarantees)", a=b("In_Cash_Rate", pct0),
               r=b("In_AtRisk_Rate", pct0), m=b("In_Min_Multiple", lambda v: f"{v:.0f}x"),
               g=b("In_Guar_Credit", pct0)), "€ m", per("Exit_Claims", eur), both("Exit_Claims", eur)],
            ["Less DSA", "€ m", per("WF_DSA", eur), both("WF_DSA", eur)],
            ["Distributable", "€ m", per("Distributable", eur), both("Distributable", eur)],
            [f("FEP net: interest claims + {s}, less FEP-borne devex", s=b("In_FEP_Share", pct0)), "€ m",
             per("Exit_FEP_Net", eur), both("Exit_FEP_Net", eur)],
            ["A.RES: 30% + DSA", "€ m", per("Exit_ARES_Total", eur), both("Exit_ARES_Total", eur)],
            ["Less guarantee fees, signing and DD, tax", "€ m", per("RtB_Route_Costs", eur),
             both("RtB_Route_Costs", eur)],
            ["FEP profit – RtB route", "€ m", per("RtB_Profit_Buildup", eur), both("RtB_Profit_Buildup", eur)]]
    s.table([f("Waterfall at the exit, {x}", x=b("Dash_Exit_Date", date)), "Unit", "Per bid", "Both bids"], rows,
            key_rows={8}, top=TOP, width=SIDE_W, height=0.59 + 9 * 0.46, col_widths=[5.58, 0.9, 1.75, 1.75])
    s.card("Cash to the exit", [
        ("", f("FEP funds the deposits, connection fees and land; the buyer reimburses €{r}m for both bids at the exit.",
               r=both("Exit_Reimbursement", eur))),
        ("", f("The largest outlay is the connection fees, €{a}m direct and €{c}m indirect, paid on {d}.",
               a=both_lab("Direct connection fee", eur), c=both_lab("Indirect connection fee, net", eur),
               d=lab("Direct connection fee", date, "Posted", "Securities"))),
        ("", "The buyer replaces the PB guarantee at the exit.")],
        left=CARD_L, top=TOP, width=CARD_W, height=2.75)
    s.footnote(f("RtB price {p} k€/MVA × {m} MVA per bid, per the Bora scenarios meeting file (Bora!N5), confirmed "
                 "{c}. Per-MW figures use {w} MW of installed wind per bid. The participation fee is not recovered.",
                 p=b("In_RtB_Premium", n0), m=b("In_MVA", n0), c=b("Dashboard!E6", date),
                 w=b("Out_Wind_MW_1", n0)))

    # ---- 5 Funding and capex (both bids; SPEC §0 item 22) ----------------------------------------------------
    fund_head = next(c.row for c in ws["C"] if str(c.value).startswith("Funding need to COD"))
    fund_total = int(next(names["Out_Funding_Equity_To_COD"].destinations)[1].split("$")[-1])
    r0, r1 = (int(x.split("$")[-1]) for x in next(names["Capex_Items"].destinations)[1].split(":"))
    capex_rows = list(range(r0, r1 + 1))
    exit_day = b("Dash_Exit_Date", date).value
    pre_exit = [r for r in range(fund_head + 1, fund_total) if ws[f"E{r}"].value <= exit_day]
    constr = next(r for r in range(fund_head + 1, fund_total) if str(ws[f"C{r}"].value).startswith("NTP → COD: construction"))
    s = d.content("Funding and capex", f(
        "Both bids need €{e}m of cash equity to COD and up to €{g}m of guarantees; all-in cost is {k} k€ per "
        "built MW", e=b("Out_Funding_Equity_To_COD", eur), g=b("Out_Funding_Peak_Guarantee", eur),
        k=b("Out_Capex_Total_MW", n0)), f(
        "€{x}m of the cash is spent by the RtB exit on {d}; on the build route, construction equity of €{c}m "
        "follows from NTP to COD.", x=b(f"Dashboard!I{pre_exit[-1]}", eur), d=b("Dash_Exit_Date", date),
        c=b(f"Dashboard!F{constr}", eur)), title_lines=2)

    def cash(v):                 # € m 1 dp; a small non-zero amount (fees) keeps 2 dp instead of showing 0.0
        return fmt_eur_m(v, 2) if v and abs(v) < 0.05 else fmt_eur_m(v)

    def money(ref):              # an empty cell is shown as a dash (no figure, no binding)
        return DASH if d.wb["Dashboard"][ref.split("!")[1]].value in (None, "") else b(ref, cash)

    rows = [[b(f"Dashboard!E{r}", date), ws[f"C{r}"].value, money(f"Dashboard!F{r}"), money(f"Dashboard!H{r}"),
             money(f"Dashboard!I{r}")] for r in range(fund_head + 1, fund_total)]
    rows.append(["", "Total cash equity to COD; peak guarantee", b("Out_Funding_Equity_To_COD", eur),
                 b("Out_Funding_Peak_Guarantee", eur), DASH])
    s.table(["Date", "Funding need – both bids", "Cash € m", "Guarantee € m", "Cumulative cash € m"], rows,
            key_rows={len(rows) - 1}, top=TOP, width=9.55, height=0.59 + len(rows) * 0.36,
            col_widths=[1.4, 4.05, 1.15, 1.45, 1.5], align="LLRRR", unit_col=None, name="Funding")
    rows = [[ws[f"C{r}"].value, b(f"Dashboard!E{r}", eur), b(f"Dashboard!F{r}", n0)] for r in capex_rows]
    rows.append(["All-in cost to COD", b("Out_Capex_Total", eur), b("Out_Capex_Total_MW", n0)])
    s.table([f("Capex to COD – both bids, {mw} MW", mw=d.bind_sum(["Out_Wind_Built_MW_1", "Out_Wind_Built_MW_2"],
                                                                    n0)), "€ m", "k€/MW"], rows, key_rows={len(rows) - 1},
            left=14.82, top=TOP, width=5.1, height=0.59 + len(rows) * 0.36, col_widths=[2.9, 1.1, 1.1],
            align="LRR", unit_col=None, name="Capex")
    s.footnote(f("Both bids together, base case, build route. Cash = FEP equity; guarantees = bank guarantees "
                 "outstanding (peak shown), not cash. IDC is debt-funded. k€/MW on {mw} MW built.",
                 mw=d.bind_sum(["Out_Wind_Built_MW_1", "Out_Wind_Built_MW_2"], n0)))

    # ---- 6 COD and HTM value -----------------------------------------------------------------------------------
    s = d.content("Economics – COD and HTM", f(
        "Both bids clear the {r} discount rate: HTM IRR {i} and NPV €{n}m for both bids; a sale at COD returns {c}",
        r=b("In_Discount_Rate", pct), i=per("HTM_IRR", pct), n=both("HTM_NPV", eur), c=per("COD_IRR", pct)), f(
        "Against the zero bid, the commitments reduce FEP's NPV by €{c}m on a sale at COD and €{h}m held to "
        "maturity, per bid.", c=delta(NPV_COD, eur_abs), h=delta(NPV_HTM, eur_abs)), title_lines=2)

    def ret(label, key, fm, dfm, total):
        return [label, "€ m" if fm is eur else "% / pp", zero(label, fm), per(key, fm), delta(label, dfm),
                both(key, fm) if total else DASH]

    rows = [ret(IRR_COD, "COD_IRR", pct, pp, False), ret(NPV_COD, "COD_NPV", eur, fmt_delta, True),
            ret(IRR_HTM, "HTM_IRR", pct, pp, False), ret(NPV_HTM, "HTM_NPV", eur, fmt_delta, True)]
    s.table(["Item", "Unit", "Zero bid, per bid", "Per bid", "Δ per bid", "Both bids"], rows, key_rows={3}, top=TOP,
            width=SIDE_W, height=0.59 + 4 * 0.5, col_widths=[4.38, 0.9, 1.25, 1.15, 1.15, 1.15])
    s.card("Per bid", [
        ("Sale at COD: ", f("FEP proceeds of €{p}m on a €{c}m price, a {m} MOIC.", p=per("COD_Proceeds", eur),
                            c=per("COD_Price", eur), m=per("COD_MOIC", fmt_mult))),
        ("Hold to maturity: ", f("{m} MOIC; minimum DSCR {d}.", m=per("HTM_MOIC", fmt_mult),
                                 d=per("Min_DSCR", fmt_mult)))],
        left=CARD_L, top=TOP, width=CARD_W, height=1.9)
    s.footnote(f("NPV at {r}. COD {cod}, {mw} MW built per bid; PPA {ppa} €/MWh; gearing {g}. HTM includes the "
                 "buy-out of the developer at the RtB exit date. Funding and capex to COD are on slide 5.",
                 r=b("In_Discount_Rate", pct), cod=b("In_COD", date), mw=b("Out_Wind_Built_MW_1", n0),
                 ppa=b("In_PPA_Price", n0), g=b("In_Gearing", pct)))

    # ---- 7 Nothing delivered after award (SPEC §0 item 23): stored Base vs "Nothing delivered", per bid --------
    BASE, NOTHING = "Base", "Nothing delivered"

    def case(key, fm, dfm):      # Base | Nothing delivered | Δ, BORA 1 (BORA 2 is identical)
        base, nd = stored(f"{key}_1", BASE, fm), stored(f"{key}_1", NOTHING, fm)
        return [base, nd, d.bind_diff(stored(f"{key}_1", NOTHING, None), stored(f"{key}_1", BASE, None), dfm)]

    def spp(v):                  # signed percentage points: +3.2
        return ("+" if round(v * 100, 1) > 0 else "") + pp(v)

    def spts(v):
        return ("+" if v > 0 else "") + pts(v)

    s = d.content("Risks and mitigants", f(
        "Delivering nothing after award costs the full PB of €{p}m per bid but lifts HTM IRR from {a} to {n}",
        p=stored("Commitment_Penalty_1", NOTHING, eur), a=stored("HTM_IRR_1", BASE, pct),
        n=stored("HTM_IRR_1", NOTHING, pct)), f(
        "The BESS, municipal, community and EU-equipment costs fall away (€{c}m less to COD per bid); the price is "
        "€{t}m of PB called on both bids and a possible group exclusion.",
        c=d.bind_diff(stored("Cost_To_COD_1", BASE, None), stored("Cost_To_COD_1", NOTHING, None), eur),
        t=d.bind_sum([stored("Commitment_Penalty_1", NOTHING, None), stored("Commitment_Penalty_2", NOTHING, None)],
                     eur)), title_lines=2)
    rows = [["Points", "points", *case("Points", pts, spts)],
            ["PB called (base: pro-rata loss)", "€ m", *case("Commitment_Penalty", eur, fmt_delta)],
            ["All-in cost to COD", "€ m", *case("Cost_To_COD", eur, fmt_delta)],
            [IRR_COD, "% / pp", *case("COD_IRR", pct, spp)],
            [IRR_HTM, "% / pp", *case("HTM_IRR", pct, spp)],
            [f("FEP NPV – hold to maturity at {r}", r=b("In_Discount_Rate", pct)), "€ m",
             *case("HTM_NPV", eur, fmt_delta)]]
    s.table(["Per bid", "Unit", "Base", "Nothing delivered", "Δ"], rows, key_rows={1, 4}, top=TOP, width=SIDE_W,
            height=0.59 + 6 * 0.5, col_widths=[4.48, 0.9, 1.4, 2.0, 1.2])
    s.card("Sanctions if nothing is delivered", [
        ("Award and points stand: ", "the sanction falls on the PB, not on the award or the score."),
        ("Full PB call (Kiírás XIII.17; GYIK Q72): ", "any unproven points commitment calls the whole PB, "
                                                       "wind and BESS."),
        ("Group exclusion (Kiírás VII.2 n; GYIK Q71): ", "a called PB can exclude the group from capacity tenders "
                                                          "for three years; it hits the group holding the SPV at the "
                                                          "time of the call."),
        ("Connection right kept (GYIK Q72): ", "a full call does not take the connection right.")],
        left=CARD_L, top=TOP, width=CARD_W, height=0.59 + 6 * 0.5)
    s.footnote(f("Nothing delivered: the bid as submitted; after award no BESS, overbuild, municipal one-off or "
                 "community contribution, and Chinese turbines and substation in capex. White labelling: the EU sale "
                 "price at COD and the EU gearing are kept. RtB route unaffected. Base PB = pro-rata loss on the "
                 "unbuilt overbuild. Zero bid for reference: COD IRR {c}, HTM IRR {h}.",
                 c=zero(IRR_COD, pct), h=zero(IRR_HTM, pct)))

    # ---- 8 Risks and mitigants -------------------------------------------------------------------------------
    s = d.content("Risks and mitigants", f(        # writer's title tightened to two lines
        "Unbuilt overbuild: both bids lose €{p}m of PB and {v} MVA of connection right pro rata; a full call is "
        "unresolved", p=both("Commitment_Penalty", eur), v=both("MVA_At_Risk", n1)),
        "The base COD and HTM returns include the pro-rata PB loss; the lost MVA, any RtB price deduction and a "
        "possible group exclusion are not in the returns.", title_lines=2)
    rows = [
        [f("Overbuild not built: PB lost pro rata to the shortfall above {t} (Kiírás XII.12; GYIK Q28/30/44) – base",
           t=b("In_Tolerance_Cap", pct0)),
         f("€{p}m, the PB on the extra MW.", p=per("Commitment_Penalty", eur)),
         f("PB at the {x} minimum. Part of the loss may be deducted from the RtB price.", x=bid(PBM, "BORA 1", x1))],
        ["Full PB call if the points-commitment rule (Kiírás XIII.17; GYIK Q72) also applies to the overbuild "
         "– unresolved",
         f("The whole PB: COD IRR {i}, HTM IRR {h}.", i=stored("COD_IRR_1", "Full PB call", pct),
           h=stored("HTM_IRR_1", "Full PB call", pct)),
         "Downside case, not the base. PB at the minimum limits it."],
        ["3-year group exclusion (Kiírás VII.2 n; GYIK Q71)",
         "Possible if the PB is called; unclear whether a pro-rata loss counts as a call. A call after COD would "
         "likely hit the group then holding the SPV.", "Allocate it in the RtB SPA."],
        ["Connection right lost pro rata (GYIK Q28/30)",
         f("{v} MVA, in the pro-rata reading only (GYIK Q72: a full call keeps the right). Not reflected in "
           "generation or price.", v=per("MVA_At_Risk", n1)),
         "The buyer can build the overbuild (card)."],
        ["ÖSZ-4 scored on installed wind (GYIK Q20)",
         f("Points fall to {p}; returns unchanged.", p=stored("Points_1", "ÖSZ-4 on wind", dp2)),
         "Ranking risk only; the commitment is the same."],
        ["BESS valued at cost only", "Its return cost (slide 3) excludes any BESS revenue or sale value.",
         "Conservative: any BESS revenue is upside."],
    ]
    s.table(["Risk", "Exposure (per bid)", "Mitigant"], rows, top=TOP, height=0.59 + 6 * 0.72,
            col_widths=[5.0, 5.4, 4.6], align="LLL", unit_col=None)
    s.card("Why the base does not build the overbuild", [
        ("", f("The grid connection stays at {m} MVA. Building the overbuild lifts the all-in cost to COD from €{a}m "
               "to €{c}m per bid. In that case COD IRR is {i} and HTM IRR is {h}. The build decision passes to the "
               "buyer at RtB.", m=b("In_MVA", n0), a=per("Cost_To_COD", eur),
               c=stored("Cost_To_COD_1", "Overbuild built", eur), i=stored("COD_IRR_1", "Overbuild built", pct),
               h=stored("HTM_IRR_1", "Overbuild built", pct)))],
        left=4.92, top=8.05, width=15.0, height=1.35)
    s.footnote(f("Shortfall tested {m} months after COD. Pro rata = shortfall share × the wind PB (GYIK Q28/30). "
                 "The full-call and overbuild-built figures are alternative cases.", m=b("In_Sanction_Lag", n0)))
    return d.save(out / "Bora_Bid_IC.pptx")
