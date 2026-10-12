"""Bora IC deck – 6 slides to SPEC.md §0 item 19, text from out/Bora_Bid_slide_text.md.

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
    m2 = "{a}m / €{b}"           # "€{x}m" around a pair -> "€1.4m / €1.2m"

    def b(ref, fmt):
        """Bind a workbook name (`Out_…`, `In_…`) or a `Sheet!Cell` reference."""
        return d.bind(*ref.split("!"), fmt=fmt) if "!" in ref else d.bind(ref, fmt=fmt)

    def pair(r1, r2, fmt, tpl="{a} / {b}"):
        return f(tpl, a=b(r1, fmt), b=b(r2, fmt))

    def dash(c):                 # Dashboard cells of BORA 1 / BORA 2 in columns c and c+1 (e.g. "I12")
        col, row = c[0], c[1:]
        return f"Dashboard!{col}{row}", f"Dashboard!{chr(ord(col) + 1)}{row}"

    def both(key, fm):
        return [b(f"Out_{key}_1", fm), b(f"Out_{key}_2", fm)]

    # ---- 1 Cover -------------------------------------------------------------------------------------------
    d.cover("Project Bora", "BORA 1 and BORA 2 capacity bids", b("Dashboard!E6", date),
            subtitle=f("Approval of the tender commitments,\n2 × {m} MVA, Szombathely OVIT", m=b("In_MVA", n0)))

    # ---- 2 Decision and timetable ----------------------------------------------------------------------------
    s = d.content("Executive summary", f(
        "Submit both bids at {p} points: FEP earns €{r}m at the RtB exit, a {i} IRR",
        p=pair("Out_Points_1", "Out_Points_2", pts), r=pair("Out_RtB_Profit_1", "Out_RtB_Profit_2", eur, m2),
        i=pair("Out_RtB_IRR_1", "Out_RtB_IRR_2", pct)), f(
        "Against the zero bid, the commitments cost €{a}m of RtB profit and {c} pp of HTM IRR, "
        "and add €{e}m to the all-in commitment.",
        a=pair(*dash("I26"), eur_abs, m2), c=pair(*dash("I30"), pp_abs), e=pair(*dash("I32"), eur_abs, m2)),
        title_lines=2)
    rows = [("Points (EU turbine)", "points", "H25", "Points", pts),
            ("FEP profit – RtB route", "€ m", "H26", "RtB_Profit", eur),
            ("FEP IRR – RtB route", "%", "H27", "RtB_IRR", pct),
            ("FEP IRR – sale at COD", "%", "H28", "COD_IRR", pct),
            ("FEP IRR – hold to maturity", "%", "H30", "HTM_IRR", pct),
            ("All-in commitment if built (securities, development, construction, financing)", "€ m", "H32",
             "Commitment_Total", eur)]
    s.table(["Item", "Unit", "Zero bid", "BORA 1", "BORA 2"],
            [[lbl, unit, b(f"Dashboard!{zero}", fm), *both(key, fm)] for lbl, unit, zero, key, fm in rows],
            key_rows={5}, top=TOP, width=SIDE_W, height=0.59 + 6 * 0.52, col_widths=[4.3, 1.1, 1.5, 1.5, 1.5])
    s.card("Approval requested and timetable", [
        (f("Submit both bids by {d}: ", d=b("In_Bid_Date", date)), f(
            "municipal {m} k€/MW; community {c} €/MWh; overbuild {o} (not built); BESS {e} of MVA, {h} h; "
            "PB at the {pb} minimum; partial award, shared connection, EU turbines and transformers.",
            m=b("Dashboard!E12", n0), c=pair(*dash("E13"), dp2), o=pair(*dash("E14"), pct0),
            e=pair(*dash("E15"), pct0), h=b("Dashboard!E16", n0), pb=b("Dashboard!E17", x1))),
        ("Day one, at submission: ", f("participation fee €{fee}m per bid (cash) and bid guarantee €{g}m.",
                                       fee=b("Dashboard!E50", eur3), g=pair(*dash("E51"), eur, m2))),
        (f("After award ({d}): ", d=b("In_Award_Date", date)), f(
            "PB guarantee €{p}m by {pd}, which replaces the bid guarantee; TSO cash security €{t}m per bid by {td}.",
            p=pair(*dash("E52"), eur, m2), pd=b("Dashboard!I52", date), t=b("Dashboard!E53", eur),
            td=b("Dashboard!I53", date))),
        ("Before submission: ", f("legal confirmation that the two bids sit within the {c} MVA group cap per node.",
                                  c=b("In_Group_Cap", n0)))],
        left=CARD_L, top=TOP, width=CARD_W, height=4.7)
    s.footnote(f("Post-tax, FEP share, EU turbine, base case (overbuild not built, BESS built, pro-rata PB call); "
                 "RtB exit {x}. Zero bid = every commitment at its zero level.", x=b("Dashboard!E60", date)))

    # ---- 3 Why these commitments -----------------------------------------------------------------------------
    s = d.content("Commitments", f(
        "Municipal buys {p} points for {a} pp of HTM IRR; BESS ({e} pp) and community ({c} pp) carry the return "
        "cost on BORA 1", p=b("Dashboard!I12", pts), a=b("Dashboard!O12", pp_abs), e=b("Dashboard!O15", pp_abs),
        c=b("Dashboard!O13", pp_abs)), f(
        "Each commitment is measured alone against the zero bid; the PB stays at the {x} minimum because "
        "{y} adds 5 points for one more PB of guarantee.",
        x=b("Dashboard!E17", x1), y=b("Curves!S19", x1)), title_lines=2)

    def impact(r):               # points, Δ RtB profit, Δ COD IRR, Δ HTM IRR, Δ deposits of Dashboard row r
        return [pair(*dash(f"I{r}"), pts), pair(*dash(f"K{r}"), fmt_delta), pair(*dash(f"M{r}"), pp),
                pair(*dash(f"O{r}"), pp), pair(*dash(f"Q{r}"), fmt_delta)]

    fmts = [fmt_delta, pp, pp, fmt_delta]

    def zero(rows_):             # a block of Δ cells that are all zero, shown once (e.g. "17:L17")
        return [b(f"Dashboard!{c1}{rows_[0]}:{c2}{rows_[1]}", fm)
                for (c1, c2), fm in zip((("K", "L"), ("M", "N"), ("O", "P"), ("Q", "R")), fmts)]

    rows = [
        ["Municipal one-off (ÖSZ-1)", pair(*dash("E12"), n0, "{a} / {b} k€/MW"), *impact(12),
         "Paid at COD, after the RtB exit: no FEP cash on the RtB route."],
        ["Community contribution (ÖSZ-2)", pair(*dash("E13"), dp2, "{a} / {b} €/MWh"), *impact(13),
         f("Operating cost after COD. At the band floor, BORA 2 earns {p} points.", p=b("Dashboard!J13", pts))],
        ["Overbuild (ÖSZ-3)", pair(*dash("E14"), pct0), *impact(14),
         "Not built: the cost is the larger PB and its pro-rata call."],
        [f("BESS power (ÖSZ-4), {h} h", h=b("Dashboard!E16", n0)), pair(*dash("E15"), pct0, "{a} / {b} of MVA"),
         *impact(15), "Sized with the overbuild and built; valued at cost, with no BESS revenue."],
        ["PB multiple (ÖSZ-6)", pair(*dash("E17"), x1), pair(*dash("I17"), pts), *zero((17, 17)),
         f("Not raised: {y} adds 5 points for one more PB of €{a}m of guarantee.",
           y=b("Curves!S19", x1), a=pair(*dash("E52"), eur, m2))],
        ["Partial award (ÖSZ-5), shared connection (ÖSZ-8), EU turbines / transformers (ÖSZ-9)", "yes",
         f("{a} / {b} / {c} / {e} each bid", a=b("Dashboard!I18", pts), b=b("Dashboard!I20", pts),
           c=b("Dashboard!I21", pts), e=b("Dashboard!I22", pts)), *zero((18, 22)),
         f("No cash or return impact. EU sourcing is in the base cost (turbine {t} k€/MW).",
           t=b("Dashboard!V18", n0))],
        ["Full bid vs zero bid", "", pair("Out_Points_1", "Out_Points_2", pts), pair(*dash("I26"), fmt_delta),
         pair(*dash("I28"), pp), pair(*dash("I30"), pp), DASH,
         "The rows do not add to the total because the commitments interact."],
    ]
    s.table(["Commitment (B1 / B2)", "Offered", "Points", "Δ RtB profit € m", "Δ COD IRR pp", "Δ HTM IRR pp",
             "Δ deposits and guarantees € m", "Logic"], rows, key_rows={6}, top=TOP, height=0.59 + 7 * 0.66,
            col_widths=[3.1, 1.95, 1.2, 1.3, 1.2, 1.2, 1.55, 3.5], align="LRRRRRRL", unit_col=None)
    s.footnote(f("Δ = zero bid plus that commitment alone, minus the zero bid. Zero levels: municipal 0, "
                 "community 0, overbuild 0%, BESS 0%, PB 1.0x, yes/no items no. Δ deposits and guarantees = extra "
                 "security posted, not a call. ÖSZ-6: 5 points at {y} (Kiírás X.4 bf, Annex 4). Brownfield (ÖSZ-7) "
                 "is not offered.", y=b("Curves!S19", x1)))

    # ---- 4 Securities and cash to the RtB exit ---------------------------------------------------------------
    s = d.content("Economics – RtB exit", f(
        "At the RtB exit FEP recovers its cash and earns €{p}m, a {m} MOIC",
        p=pair("Out_RtB_Profit_1", "Out_RtB_Profit_2", eur, m2),
        m=pair("Out_RtB_MOIC_1", "Out_RtB_MOIC_2", fmt_mult)), f(
        "FEP funds €{o}m per bid to {x} and carries the PB guarantee. Of the €{t}m RtB price net of sale costs, "
        "FEP takes €{a}m.", o=b("Out_RtB_Outlay_1", eur_abs), x=b("Dashboard!E60", date),
        t=b("Out_Split_Total_1", eur), a=pair("Out_Split_FEP_1", "Out_Split_FEP_2", eur, m2)), title_lines=2)

    def split(key):              # "11.1 (73)": € m and k€ per installed wind MW
        return [f("{a} ({k})", a=b(f"Out_Split_{key}_{i}", eur), k=b(f"Out_Split_{key}_MW_{i}", n0))
                for i in (1, 2)]

    rows = [["RtB price", "€ m", *both("RtB_Price", eur)],
            ["Receipts at the exit", "€ m", *both("FEP_Proceeds", eur)],
            ["Outlays: devex, land, deposits, connection fees, signing", "€ m", *both("RtB_Outlay", eur)],
            [f("Guarantee fees ({g} carry)", g=b("In_Guarantee_Fee", pct0)), "€ m", *both("Guarantee_Fees", eur)],
            ["Tax at the exit", "€ m", *both("RtB_Tax", eur)],
            ["FEP profit – RtB route", "€ m", *both("RtB_Profit_Buildup", eur)],
            [f("FEP profit – RtB route, low-premium case ({p} k€/MVA)", p=b("Inputs!L74", n0)), "€ m",
             b("Dashboard!L105", eur), b("Dashboard!L106", eur)],
            [f("Exit split – FEP: claims + {s}", s=b("In_FEP_Share", pct0)), "€ m (k€/MW)", *split("FEP")],
            ["Exit split – A.RES: DSA + 30%", "€ m (k€/MW)", *split("ARES")],
            ["Exit split – total", "€ m (k€/MW)", *split("Total")]]
    s.table(["Item", "Unit", "BORA 1", "BORA 2"], rows, key_rows={5}, top=TOP, width=SIDE_W,
            height=0.59 + 10 * 0.46, col_widths=[5.06, 1.6, 1.6, 1.6])
    s.card("Securities to the exit", [
        ("Bid guarantee ", f("€{a}m, {p} to {r}.", a=pair(*dash("E51"), eur, m2),
                             p=b("Dashboard!I51", date), r=b("Dashboard!J51", date))),
        ("PB guarantee ", f("€{a}m from {p}; the buyer replaces it at the exit.",
                            a=pair(*dash("E52"), eur, m2), p=b("Dashboard!I52", date))),
        ("Cash: ", f("TSO security €{t}m ({td}) and connection fees €{c1}m + €{c2}m ({cd}). "
                     "The buyer reimburses €{r}m at the exit.",
                     t=b("Dashboard!E53", eur), td=b("Dashboard!I53", date), c1=b("Dashboard!E54", eur),
                     c2=b("Dashboard!E55", eur), cd=b("Dashboard!I54", date), r=b("Out_WF_Principal_1", eur)))],
        left=CARD_L, top=TOP, width=CARD_W, height=2.6)
    s.footnote(f("RtB price {p} k€/MVA × {m} MVA, per the Bora scenarios meeting file (Bora!N5), confirmed {c}. "
                 "Receipts = recoverable principal, land and claims, plus {s} of the distributable. Per-MW figures "
                 "use installed wind of {w} MW. The participation fee is not recovered.",
                 p=b("In_RtB_Premium", n0), m=b("In_MVA", n0), c=b("Dashboard!E6", date),
                 s=b("In_FEP_Share", pct0), w=pair("Out_Wind_MW_1", "Out_Wind_MW_2", n0)))

    # ---- 5 COD and HTM value and the commitment breakdown ----------------------------------------------------
    s = d.content("Economics – COD and HTM", f(
        "All-in cost of {c}\u00a0k€ per built MW against a {p}\u00a0k€/MW COD price; HTM NPV €{n}m at {r}",
        c=pair("Out_Cost_To_COD_MW_1", "Out_Cost_To_COD_MW_2", n0), p=b("In_COD_Price", n0),
        n=pair("Out_HTM_NPV_1", "Out_HTM_NPV_2", eur, m2), r=b("Dashboard!V22", pct)), f(
        "The commitments reduce HTM NPV by €{h}m against the zero bid; construction is €{c}m of the all-in "
        "commitment.", h=pair(*dash("I31"), eur_abs, m2),
        c=pair("Out_Commitment_Construction_1", "Out_Commitment_Construction_2", eur, m2)), title_lines=2)

    def npv(lbl, r, key):
        return [lbl, "€ m", b(f"Dashboard!H{r}", eur), *both(key, eur), *[b(c, fmt_delta) for c in dash(f"I{r}")]]

    def memo(lbl, unit, vals):
        return [lbl, unit, DASH, *vals, DASH, DASH]

    rows = [npv("FEP NPV – sale at COD", 29, "COD_NPV"),
            npv("FEP NPV – hold to maturity", 31, "HTM_NPV"),
            memo("All-in cost to COD per built MW", "k€/MW", both("Cost_To_COD_MW", n0)),
            memo("COD sale price per built MW", "k€/MW", [b("In_COD_Price", n0), b("In_COD_Price", n0)]),
            memo("Securities: cash posted and guarantees", "€ m", both("Commitment_Securities", eur)),
            memo("Devex, land, signing and DD", "€ m", both("Commitment_Devex", eur)),
            memo("Construction capex incl. connection fees, municipal and pre-COD lease", "€ m",
                 both("Commitment_Construction", eur)),
            memo("Financing: IDC, debt fees, guarantee fees", "€ m", both("Commitment_Financing", eur)),
            memo("Memo, not in the commitment: PB call, overbuild not built", "€ m", both("Commitment_Penalty", eur))]
    s.table(["Item", "Unit", "Zero bid", "BORA 1", "BORA 2", "Δ B1", "Δ B2"], rows, key_rows={1}, top=TOP,
            width=SIDE_W, height=0.59 + 9 * 0.46, col_widths=[4.06, 0.9, 1.0, 1.0, 1.0, 0.98, 0.98])
    s.card(None, [
        ("Sale at COD: ", f("FEP proceeds of €{p}m on a €{c}m price, a {m} MOIC.",
                            p=pair("Out_COD_Proceeds_1", "Out_COD_Proceeds_2", eur, m2),
                            c=b("Out_COD_Price_1", eur), m=pair("Out_COD_MOIC_1", "Out_COD_MOIC_2", fmt_mult))),
        ("Hold to maturity: ", f("{m} MOIC; minimum DSCR {d}.", m=pair("Out_HTM_MOIC_1", "Out_HTM_MOIC_2", fmt_mult),
                                 d=pair("Out_Min_DSCR_1", "Out_Min_DSCR_2", fmt_mult)))],
        left=CARD_L, top=TOP, width=CARD_W, height=1.75)
    s.footnote(f("NPV at {r}. COD {cod}, {mw} MW built; PPA {ppa} €/MWh; gearing {g}. HTM includes the buy-out of "
                 "the developer at the RtB exit date. The four breakdown lines sum to the all-in commitment on "
                 "slide 2.", r=b("Dashboard!V22", pct), cod=b("Dashboard!V23", date),
                 mw=b("Out_Wind_Built_MW_1", n0), ppa=b("Dashboard!V19", n0), g=b("Dashboard!V20", pct)))

    # ---- 6 Risks and mitigants -------------------------------------------------------------------------------
    s = d.content("Risks and mitigants", f(
        "Main risk: the unbuilt overbuild – a PB call, a 3-year group exclusion and {v} MVA of "
        "connection right lost", v=pair("Out_MVA_At_Risk_1", "Out_MVA_At_Risk_2", n1)),
        "The base COD and HTM returns include the pro-rata call; the lost MVA and any deduction from the RtB "
        "price are not in the returns.", title_lines=2)
    rows = [
        ["PB call on the unbuilt overbuild: pro rata (Kiírás XII.12) or full (XIII.17, GYIK Q72)",
         f("Pro rata €{p}m (base). Full call €{c}m: COD IRR {i}, HTM IRR {h}",
           p=pair("Out_Commitment_Penalty_1", "Out_Commitment_Penalty_2", eur, m2),
           c=pair("Out_Sanction_Full_1", "Out_Sanction_Full_2", eur, m2),
           i=pair("Dashboard!J107", "Dashboard!J108", pct), h=pair("Dashboard!J109", "Dashboard!J110", pct)),
         f("PB at the {x} minimum. Part of the call may be deducted from the RtB price.", x=b("Dashboard!E17", x1))],
        ["3-year group exclusion (Kiírás VII.2 n, GYIK Q71)",
         "Any PB call triggers it, the base pro-rata call included, for both bids. The call falls after COD, "
         "so it is likely to hit the group then holding the SPV.", "Allocate it in the RtB SPA."],
        ["Connection right lost pro rata (Kiírás XII.12, GYIK Q28)",
         f("{v} MVA; not reflected in generation or price.", v=pair("Out_MVA_At_Risk_1", "Out_MVA_At_Risk_2", n1)),
         "The buyer can build the overbuild (card)."],
        [f("{c} MVA group cap per node (Kiírás V.7, GYIK Q52)", c=b("In_Group_Cap", n0)),
         "Both bids are invalid if FEP, A.RES and the SPVs count as one company group.",
         "Legal confirmation before submission."],
        ["ÖSZ-4 scored on installed wind (GYIK Q20)",
         f("Points fall to {p}; returns unchanged.", p=pair("Dashboard!M103", "Dashboard!M104", dp2)),
         "Ranking risk only; the commitment is the same."],
        ["BESS valued at cost only",
         f("{a} pp of COD IRR, with no BESS revenue or sale value.", a=pair(*dash("M15"), pp)),
         "Conservative: any BESS revenue is upside."],
    ]
    s.table(["Risk", "Exposure (B1 / B2)", "Mitigant"], rows, top=TOP, height=0.59 + 6 * 0.66,
            col_widths=[4.4, 6.0, 4.6], align="LLL", unit_col=None)
    s.card("Why the base does not build the overbuild", [
        ("", f("The grid connection stays at {m} MVA. Building the overbuild lifts the all-in cost to COD from €{a}m "
               "to €{c}m. In that case COD IRR is {i} and HTM IRR is {h}. The build decision passes to the buyer "
               "at RtB.", m=b("In_MVA", n0), a=pair("Out_Cost_To_COD_1", "Out_Cost_To_COD_2", eur, m2),
               c=pair("Dashboard!I113", "Dashboard!I114", eur, m2), i=pair("Dashboard!I107", "Dashboard!I108", pct),
               h=pair("Dashboard!I109", "Dashboard!I110", pct)))],
        left=4.92, top=7.45, width=15.0, height=1.35)
    s.footnote(f("PB call tested {m} months after COD. Pro rata = shortfall share × the wind PB (GYIK Q28/Q30). "
                 "Full-call and overbuild-built figures are alternative cases.", m=b("In_Sanction_Lag", n0)))
    return d.save(out / "Bora_Bid_IC.pptx")
