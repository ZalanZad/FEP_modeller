"""Bora IC deck – 6 slides to SPEC.md §0 items 19 and 21 (one strategy for both bids), text from
out/Bora_Bid_slide_text.md.

Every figure is `Deck.bind` on the packaged workbook (LibreOffice-recalculated when `package` saves it, so it holds
the same values as the audit's `_recalc.xlsx`). Composite texts use `Deck.f`, which records each bound part.
"""
from pathlib import Path

from fepm.ppt.deck import DASH, Deck, fmt_delta, fmt_eur_m, fmt_mult, fmt_num, fmt_pct

SECTIONS = ["Executive summary", "Commitments", "Economics – RtB exit", "Economics – COD and HTM",
            "Risks and mitigants"]
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

    def b(ref, fmt):
        """Bind a workbook name (`Out_…`, `In_…`) or a `Sheet!Cell` reference."""
        return d.bind(*ref.split("!"), fmt=fmt) if "!" in ref else d.bind(ref, fmt=fmt)

    def per(key, fmt):           # per bid: the BORA 1 cell (BORA 2 is identical, SPEC §0 item 21)
        return b(f"Out_{key}_1", fmt)

    def both(key, fmt):          # both bids: Dashboard Total (column H) where the row has one, else B1 + B2
        sheet, cell = next(names[f"Out_{key}_1"].destinations)
        ws, row = d.wb[sheet], int(cell.split("$")[-1])
        head = next(r for r in range(row, 0, -1) if ws[f"A{r}"].value == "x")     # the block's header row
        if ws[f"H{head}"].value == "Total":
            return b(f"{sheet}!H{row}", fmt)
        return d.bind_sum([f"Out_{key}_1", f"Out_{key}_2"], fmt)

    def cells(c, fmt):           # both bids from a Dashboard cell pair (columns c and c+1), e.g. "E51"
        return d.bind_sum([f"Dashboard!{c}", f"Dashboard!{chr(ord(c[0]) + 1)}{c[1:]}"], fmt)

    sets = [c.value for row in d.wb["Inputs"][next(names["Scenario_Names"].destinations)[1]] for c in row]

    def stored(key, set_name, fmt):   # BORA 1 result stored for one key-assumption set
        return d.stored(key, sets.index(set_name) + 1, fmt)

    # ---- 1 Cover -------------------------------------------------------------------------------------------
    d.cover("Project Bora", "BORA 1 and BORA 2 capacity bids", b("Dashboard!E6", date),
            subtitle=f("Approval of the tender commitments,\n2 × {m} MVA, Szombathely OVIT", m=b("In_MVA", n0)))

    # ---- 2 Decision and timetable ----------------------------------------------------------------------------
    s = d.content("Executive summary", f(
        "Submit both bids with the same commitments at {p} points each: FEP earns €{r}m at the RtB exit, a {i} IRR",
        p=per("Points", pts), r=both("RtB_Profit_Buildup", eur), i=per("RtB_IRR", pct)), f(
        "Against the zero bid, the commitments cost €{a}m of RtB profit and {c} pp of HTM IRR per bid, "
        "and add €{e}m per bid to the all-in commitment.",
        a=b("Dashboard!I26", eur_abs), c=b("Dashboard!I30", pp_abs), e=b("Dashboard!I32", eur_abs)), title_lines=2)
    rows = [["Points (EU turbine)", "points", b("Dashboard!H25", pts), per("Points", pts), DASH],
            ["FEP profit – RtB route", "€ m", b("Dashboard!H26", eur), per("RtB_Profit", eur),
             both("RtB_Profit_Buildup", eur)],
            ["FEP IRR – RtB route", "%", b("Dashboard!H27", pct), per("RtB_IRR", pct), DASH],
            ["FEP IRR – sale at COD", "%", b("Dashboard!H28", pct), per("COD_IRR", pct), DASH],
            ["FEP IRR – hold to maturity", "%", b("Dashboard!H30", pct), per("HTM_IRR", pct), DASH],
            ["All-in commitment if built (securities, development, construction, financing)", "€ m",
             b("Dashboard!H32", eur), per("Commitment_Total", eur), both("Commitment_Total", eur)]]
    s.table(["Item", "Unit", "Zero bid, per bid", "Per bid", "Both bids"], rows, key_rows={5}, top=TOP,
            width=SIDE_W, height=0.59 + 6 * 0.52, col_widths=[4.3, 1.1, 1.5, 1.5, 1.5])
    s.card("Approval requested and timetable", [
        (f("Submit BORA 1 and BORA 2 by {d}, each with: ", d=b("In_Bid_Date", date)), f(
            "municipal {m} k€/MW; community {c} €/MWh; overbuild {o} (not built); BESS {e} of MVA, {h} h; "
            "PB at the {pb} minimum; partial award, shared connection, EU turbines and transformers.",
            m=b("Dashboard!E12", n0), c=b("Dashboard!E13", dp2), o=b("Dashboard!E14", pct0),
            e=b("Dashboard!E15", pct0), h=b("Dashboard!E16", n0), pb=b("Dashboard!E17", x1))),
        ("Day one, at submission (both bids): ", f("participation fees €{fee}m in cash and bid guarantees of €{g}m.",
                                                   fee=cells("E50", eur3), g=cells("E51", eur))),
        (f("After award ({d}), both bids: ", d=b("In_Award_Date", date)), f(
            "PB guarantees of €{p}m by {pd}, replacing the bid guarantees; TSO cash security of €{t}m by {td}.",
            p=cells("E52", eur), pd=b("Dashboard!I52", date), t=cells("E53", eur), td=b("Dashboard!I53", date)))],
        left=CARD_L, top=TOP, width=CARD_W, height=4.1)
    s.footnote(f("Post-tax, FEP share, EU turbine, base case (overbuild not built, BESS built, PB lost pro rata to the "
                 "shortfall); RtB exit {x}. Zero bid = every commitment at its zero level. The two bids are identical, "
                 "so IRRs are the same per bid and for both.", x=b("Dash_Exit_Date", date)))

    # ---- 3 Why these commitments -----------------------------------------------------------------------------
    s = d.content("Commitments", f(
        "Municipal buys {p} points for {a} pp of HTM IRR; BESS ({e} pp) and community ({c} pp) carry the return cost",
        p=b("Dashboard!I12", pts), a=b("Dashboard!O12", pp_abs), e=b("Dashboard!O15", pp_abs),
        c=b("Dashboard!O13", pp_abs)), f(
        "Each commitment is measured alone against the zero bid, per bid; the PB stays at the minimum because {y} "
        "adds 5 points for one more PB of guarantee.", y=b("Curves!S19", x1)), title_lines=2)

    def impact(r):               # points, Δ RtB profit, Δ COD IRR, Δ HTM IRR, Δ deposits of Dashboard row r (BORA 1)
        return [b(f"Dashboard!{c}{r}", fm) for c, fm in
                (("I", pts), ("K", fmt_delta), ("M", pp), ("O", pp), ("Q", fmt_delta))]

    rows = [
        ["Municipal one-off (ÖSZ-1)", f("{a} k€/MW", a=b("Dashboard!E12", n0)), *impact(12),
         "Paid at COD, after the RtB exit: no FEP cash on the RtB route."],
        ["Community contribution (ÖSZ-2)", f("{a} €/MWh", a=b("Dashboard!E13", dp2)), *impact(13),
         "Operating cost after COD."],
        ["Overbuild (ÖSZ-3)", b("Dashboard!E14", pct0), *impact(14),
         "Not built: the cost is the larger PB, lost pro rata to the shortfall."],
        [f("BESS power (ÖSZ-4), {h} h", h=b("Dashboard!E16", n0)), f("{a} of MVA", a=b("Dashboard!E15", pct0)),
         *impact(15), "Sized with the overbuild and built; valued at cost, with no BESS revenue."],
        ["PB multiple (ÖSZ-6)", b("Dashboard!E17", x1), *impact(17),
         f("Not raised: {y} adds 5 points for one more PB of €{a}m of guarantee.",
           y=b("Curves!S19", x1), a=b("Dashboard!E52", eur))],
        ["Partial award (ÖSZ-5), shared connection (ÖSZ-8), EU turbines / transformers (ÖSZ-9)", "yes",
         f("{a} / {b} / {c} / {e}", a=b("Dashboard!I18", pts), b=b("Dashboard!I20", pts),
           c=b("Dashboard!I21", pts), e=b("Dashboard!I22", pts)),
         *[b(f"Dashboard!{c}18:{c}22", fm) for c, fm in (("K", fmt_delta), ("M", pp), ("O", pp), ("Q", fmt_delta))],
         f("No cash or return impact. EU sourcing is in the base cost (turbine {t} k€/MW).",
           t=b("Dashboard!V18", n0))],
        ["Full bid vs zero bid", "", per("Points", pts), b("Dashboard!I26", fmt_delta), b("Dashboard!I28", pp),
         b("Dashboard!I30", pp), DASH, "The rows do not add to the total because the commitments interact."],
    ]
    s.table(["Commitment (per bid)", "Offered", "Points", "Δ RtB profit € m", "Δ COD IRR pp", "Δ HTM IRR pp",
             "Δ deposits and guarantees € m", "Logic"], rows, key_rows={6}, top=TOP, height=0.59 + 7 * 0.62,
            col_widths=[3.2, 1.5, 1.3, 1.3, 1.25, 1.25, 1.75, 3.45], align="LRRRRRRL", unit_col=None)
    s.footnote(f("Δ = zero bid plus that commitment alone, minus the zero bid. Zero levels: municipal 0, "
                 "community 0, overbuild 0%, BESS 0%, PB 1.0x, yes/no items no. Δ deposits and guarantees = extra "
                 "security posted, not a PB loss. ÖSZ-6: 5 points at {y} (Kiírás X.4 bf, Annex 4). Brownfield (ÖSZ-7) "
                 "is not offered.", y=b("Curves!S19", x1)))

    # ---- 4 RtB exit: the devex-file waterfall (paid amounts) -------------------------------------------------
    prem_row = next(names["In_RtB_Premium"].destinations)[1].split("$")[-1]
    prem_col = next(names["Scenario_Names"].destinations)[1].split(":")[0].split("$")[1]
    low_col = chr(ord(prem_col) + sets.index("Low premium"))
    share = b("In_FEP_Share", lambda v: f"{v * 100:.0f}")
    s = d.content("Economics – RtB exit", f(
        "At the RtB exit FEP earns €{p}m on both bids, a {m} MOIC; its interest claims are paid ahead of the "
        "{s}/30 split", p=both("RtB_Profit_Buildup", eur), m=per("RtB_MOIC", fmt_mult), s=share), f(
        "FEP keeps {fe} k€ and A.RES {ar} k€ per installed wind MW; at a {lp} k€/MVA RtB price, FEP's "
        "profit falls to €{lo}m for both bids.",
        fe=b("Out_Exit_FEP_Net_MW_1", n0), ar=b("Out_Exit_ARES_Total_MW_1", n0),
        lp=b(f"Inputs!{low_col}{prem_row}", n0),
        lo=d.bind_sum([stored("RtB_Profit_1", "Low premium", None), stored("RtB_Profit_2", "Low premium", None)],
                      eur)), title_lines=2)
    pct_ = pct0
    rows = [["RtB price", "€ m", per("RtB_Price", eur), both("RtB_Price", eur)],
            ["Less sale costs", "€ m", per("WF_Sale_Costs", eur), both("WF_Sale_Costs", eur)],
            [f("Less FEP interest claims ({a} deposits, {r} / {m} devex, {g} guarantees)", a=b("In_Cash_Rate", pct_),
               r=b("In_AtRisk_Rate", pct_), m=b("In_Min_Multiple", lambda v: f"{v:.0f}x"),
               g=b("In_Guar_Credit", pct_)), "€ m", per("Exit_Claims", eur), both("Exit_Claims", eur)],
            ["Less DSA", "€ m", per("WF_DSA", eur), both("WF_DSA", eur)],
            ["Distributable", "€ m", per("Distributable", eur), both("Distributable", eur)],
            [f("FEP net: interest claims + {s}, less FEP-borne devex", s=b("In_FEP_Share", pct_)), "€ m",
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
               a=cells("E54", eur), c=cells("E55", eur), d=b("Dashboard!I54", date))),
        ("", "The buyer replaces the PB guarantee at the exit.")],
        left=CARD_L, top=TOP, width=CARD_W, height=2.75)
    s.footnote(f("RtB price {p} k€/MVA × {m} MVA per bid, per the Bora scenarios meeting file (Bora!N5), confirmed "
                 "{c}. Per-MW figures use {w} MW of installed wind per bid. The participation fee is not recovered.",
                 p=b("In_RtB_Premium", n0), m=b("In_MVA", n0), c=b("Dashboard!E6", date),
                 w=b("Out_Wind_MW_1", n0)))

    # ---- 5 COD and HTM value and the commitment breakdown ----------------------------------------------------
    s = d.content("Economics – COD and HTM", f(
        "All-in cost of {c} k€ per built MW against a {p} k€/MW COD price; HTM NPV of €{n}m for both bids "
        "at {r}", c=per("Cost_To_COD_MW", n0), p=b("In_COD_Price", n0), n=both("HTM_NPV", eur),
        r=b("Dashboard!V22", pct)), f(
        "Against the zero bid, the commitments reduce FEP's NPV by €{c}m on a sale at COD and €{h}m held to "
        "maturity, per bid.", c=b("Dashboard!I29", eur_abs), h=b("Dashboard!I31", eur_abs)), title_lines=2)

    def npv(lbl, r, key):
        return [lbl, "€ m", b(f"Dashboard!H{r}", eur), per(key, eur), b(f"Dashboard!I{r}", fmt_delta), both(key, eur)]

    def line(lbl, unit, key, fm, total=True):
        return [lbl, unit, DASH, per(key, fm), DASH, both(key, fm) if total else DASH]

    rows = [npv("FEP NPV – sale at COD", 29, "COD_NPV"),
            npv("FEP NPV – hold to maturity", 31, "HTM_NPV"),
            line("All-in cost to COD per built MW", "k€/MW", "Cost_To_COD_MW", n0, total=False),
            ["COD sale price per built MW", "k€/MW", DASH, b("In_COD_Price", n0), DASH, DASH],
            line("Securities: cash posted and guarantees", "€ m", "Commitment_Securities", eur),
            line("Devex, land, signing and DD", "€ m", "Commitment_Devex", eur),
            line("Construction capex incl. connection fees, municipal and pre-COD lease", "€ m",
                 "Commitment_Construction", eur),
            line("Financing: IDC, debt fees, guarantee fees", "€ m", "Commitment_Financing", eur)]
    s.table(["Item", "Unit", "Zero bid, per bid", "Per bid", "Δ per bid", "Both bids"], rows, key_rows={1}, top=TOP,
            width=SIDE_W, height=0.59 + 8 * 0.48, col_widths=[4.38, 0.9, 1.25, 1.15, 1.15, 1.15])
    s.card("Per bid", [
        ("Sale at COD: ", f("FEP proceeds of €{p}m on a €{c}m price, a {m} MOIC.", p=per("COD_Proceeds", eur),
                            c=per("COD_Price", eur), m=per("COD_MOIC", fmt_mult))),
        ("Hold to maturity: ", f("{m} MOIC; minimum DSCR {d}.", m=per("HTM_MOIC", fmt_mult),
                                 d=per("Min_DSCR", fmt_mult)))],
        left=CARD_L, top=TOP, width=CARD_W, height=1.9)
    s.footnote(f("NPV at {r}. COD {cod}, {mw} MW built per bid; PPA {ppa} €/MWh; gearing {g}. HTM includes the "
                 "buy-out of the developer at the RtB exit date. The four breakdown lines sum to the all-in commitment "
                 "on slide 2.", r=b("Dashboard!V22", pct), cod=b("Dashboard!V23", date),
                 mw=b("Out_Wind_Built_MW_1", n0), ppa=b("Dashboard!V19", n0), g=b("Dashboard!V20", pct)))

    # ---- 6 Risks and mitigants -------------------------------------------------------------------------------
    s = d.content("Risks and mitigants", f(        # writer's title tightened to two lines
        "Unbuilt overbuild: both bids lose €{p}m of PB and {v} MVA of connection right pro rata; a full call is "
        "unresolved", p=both("Commitment_Penalty", eur), v=both("MVA_At_Risk", n1)),
        "The base COD and HTM returns include the pro-rata PB loss; the lost MVA, any RtB price deduction and a "
        "possible group exclusion are not in the returns.", title_lines=2)
    rows = [
        [f("Overbuild not built: PB lost pro rata to the shortfall above {t} (Kiírás XII.12; GYIK Q28/30/44) – base",
           t=b("In_Tolerance_Cap", pct0)),
         f("€{p}m, the PB on the extra MW.", p=per("Commitment_Penalty", eur)),
         f("PB at the {x} minimum. Part of the loss may be deducted from the RtB price.", x=b("Dashboard!E17", x1))],
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
