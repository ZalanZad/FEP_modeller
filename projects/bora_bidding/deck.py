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

    sets = [c.value for row in d.wb["Inputs"][next(d.wb.defined_names["Scenario_Names"].destinations)[1]]
            for c in row]          # stored key-assumption sets, in the order of the Stored_Results columns

    def stored(key, set_name, fm, tpl="{a} / {b}"):   # BORA 1 / BORA 2 result stored for one key-assumption set
        i = sets.index(set_name) + 1
        return f(tpl, a=d.stored(f"{key}_1", i, fm), b=d.stored(f"{key}_2", i, fm))

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
    s.footnote(f("Post-tax, FEP share, EU turbine, base case (overbuild not built, BESS built, PB lost pro rata to the "
                 "shortfall); "
                 "RtB exit {x}. Zero bid = every commitment at its zero level.", x=b("Dash_Exit_Date", date)))

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
         "Not built: the cost is the larger PB, lost pro rata to the shortfall."],
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
                 "security posted, not a PB loss. ÖSZ-6: 5 points at {y} (Kiírás X.4 bf, Annex 4). Brownfield (ÖSZ-7) "
                 "is not offered.", y=b("Curves!S19", x1)))

    # ---- 4 Securities and cash to the RtB exit: the devex-file waterfall (paid amounts) ---------------------
    prem_row = next(d.wb.defined_names["In_RtB_Premium"].destinations)[1].split("$")[-1]
    prem_col = next(d.wb.defined_names["Scenario_Names"].destinations)[1].split(":")[0].split("$")[1]
    low_col = chr(ord(prem_col) + sets.index("Low premium"))
    s = d.content("Economics – RtB exit", f(
        "At the RtB exit FEP recovers its cash and earns €{p}m, a {m} MOIC",
        p=pair("Out_RtB_Profit_1", "Out_RtB_Profit_2", eur, m2),
        m=pair("Out_RtB_MOIC_1", "Out_RtB_MOIC_2", fmt_mult)), f(
        "Of the €{t}m exit premium per bid, FEP nets €{n}m and A.RES €{a}m; at a {lp} k€/MVA premium "
        "FEP's RtB profit falls to €{lo}m.", t=b("Out_RtB_Price_1", eur),
        n=pair("Out_Exit_FEP_Net_1", "Out_Exit_FEP_Net_2", eur, m2),
        a=pair("Out_Exit_ARES_Total_1", "Out_Exit_ARES_Total_2", eur, m2),
        lp=b(f"Inputs!{low_col}{prem_row}", n0), lo=stored("RtB_Profit", "Low premium", eur, m2)), title_lines=2)

    def wf(lbl, key):            # waterfall row: BORA 1 | BORA 2 (named outputs) | Total (column H of that row)
        row = next(d.wb.defined_names[f"Out_{key}_1"].destinations)[1].split("$")[-1]
        return [lbl, "€ m", *both(key, eur), b(f"Dashboard!H{row}", eur)]

    rows = [wf("Memo: buyer reimbursement (deposits, fees, land)", "Exit_Reimbursement"),
            wf("Exit premium (RtB price)", "RtB_Price"),
            wf("Less sale costs", "WF_Sale_Costs"),
            wf(f("Less FEP claim {r} on deposits and connection fees", r=b("In_Cash_Rate", pct0)), "WF_Return_12"),
            wf(f("Less FEP claim {r} / {m} on devex and land", r=b("In_AtRisk_Rate", pct0),
                 m=b("In_Min_Multiple", lambda v: f"{v:.0f}×")), "WF_Return_AtRisk"),
            wf(f("Less FEP claim {r} on guarantees", r=b("In_Guar_Credit", pct0)), "WF_Return_6"),
            wf("Less FEP interest claims, total", "Exit_Claims"),
            wf("Less DSA", "WF_DSA"),
            wf("Distributable", "Distributable"),
            wf(f("FEP {s}", s=b("In_FEP_Share", pct0)), "Exit_FEP_70"),
            wf("FEP-borne devex principal", "Exit_FEP_Devex"),
            wf(f("FEP net = claims + {s} − devex", s=b("In_FEP_Share", pct0)), "Exit_FEP_Net"),
            wf("A.RES 30%", "Exit_ARES_30"),
            wf("A.RES total = 30% + DSA", "Exit_ARES_Total"),
            wf("Guarantee fees, signing and DD, tax", "RtB_Route_Costs"),
            wf("FEP profit – RtB route", "RtB_Profit_Buildup")]
    s.table(["RtB exit waterfall (paid)", "Unit", "BORA 1", "BORA 2", "Total"], rows, key_rows={11, 13, 15},
            top=TOP, width=SIDE_W, height=0.59 + 16 * 0.36, col_widths=[4.68, 0.8, 1.5, 1.5, 1.5])
    s.card("Securities to the exit", [
        ("Bid guarantee ", f("€{a}m, {p} to {r}.", a=pair(*dash("E51"), eur, m2),
                             p=b("Dashboard!I51", date), r=b("Dashboard!J51", date))),
        ("PB guarantee ", f("€{a}m from {p}; the buyer replaces it at the exit.",
                            a=pair(*dash("E52"), eur, m2), p=b("Dashboard!I52", date))),
        ("Cash: ", f("TSO security €{t}m ({td}) and connection fees €{c1}m + €{c2}m ({cd}). The buyer "
                     "reimburses deposits, fees and land at the exit (memo line).",
                     t=b("Dashboard!E53", eur), td=b("Dashboard!I53", date), c1=b("Dashboard!E54", eur),
                     c2=b("Dashboard!E55", eur), cd=b("Dashboard!I54", date)))],
        left=CARD_L, top=TOP, width=CARD_W, height=2.6)
    s.footnote(f("RtB price {p} k€/MVA × {m} MVA, per the Bora scenarios meeting file (Bora!N5), confirmed {c}. "
                 "Paid amounts at the RtB exit on {x}, in devex-file order. The participation fee is not recovered.",
                 p=b("In_RtB_Premium", n0), m=b("In_MVA", n0), c=b("Dashboard!E6", date),
                 x=b("Dash_Exit_Date", date)))

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
            memo("Memo, not in the commitment: PB lost pro rata, overbuild not built", "€ m", both("Commitment_Penalty", eur))]
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
    pen = pair("Out_Commitment_Penalty_1", "Out_Commitment_Penalty_2", eur, m2)
    mva = pair("Out_MVA_At_Risk_1", "Out_MVA_At_Risk_2", n1)
    s = d.content("Risks and mitigants", f(        # writer's title tightened to two lines
        "Unbuilt overbuild: €{p}m of PB and {v} MVA of connection right lost pro rata; a full call is unresolved",
        p=pen, v=mva),
        "The base COD and HTM returns include the pro-rata PB loss; the lost MVA, any RtB price deduction and a "
        "possible group exclusion are not in the returns.", title_lines=2)
    rows = [
        [f("Overbuild not built: PB lost pro rata to the shortfall above {t} (Kiírás XII.12; GYIK Q28/30/44) – base",
           t=b("In_Tolerance_Cap", pct0)),
         f("€{p}m, the PB on the extra MW.", p=pen),
         f("PB at the {x} minimum. Part of the loss may be deducted from the RtB price.", x=b("Dashboard!E17", x1))],
        ["Full PB call if the points-commitment rule (Kiírás XIII.17; GYIK Q72) also applies to the overbuild "
         "– unresolved",
         f("€{c}m: COD IRR {i}, HTM IRR {h}.", c=pair("Out_Sanction_Full_1", "Out_Sanction_Full_2", eur, m2),
           i=stored("COD_IRR", "Full PB call", pct), h=stored("HTM_IRR", "Full PB call", pct)),
         "Downside case, not the base. PB at the minimum limits it."],
        ["3-year group exclusion (Kiírás VII.2 n; GYIK Q71)",
         "Possible if the PB is called; unclear whether a pro-rata loss counts as a call. A call after COD would "
         "likely hit the group then holding the SPV.", "Allocate it in the RtB SPA."],
        ["Connection right lost pro rata (GYIK Q28/30)",
         f("{v} MVA, in the pro-rata reading only (GYIK Q72: a full call keeps the right). Not reflected in "
           "generation or price.", v=mva),
         "The buyer can build the overbuild (card)."],
        [f("{c} MVA group cap per node (Kiírás V.7; GYIK Q52)", c=b("In_Group_Cap", n0)),
         "Both bids are invalid if FEP, A.RES and the SPVs count as one company group.",
         "Legal confirmation before submission."],
        ["ÖSZ-4 scored on installed wind (GYIK Q20); BESS valued at cost only",
         f("Points {p}, returns unchanged. BESS costs {a} pp of COD IRR with no BESS revenue.",
           p=stored("Points", "ÖSZ-4 on wind", dp2), a=pair(*dash("M15"), pp)),
         "Ranking risk only. Any BESS revenue is upside."],
    ]
    s.table(["Risk", "Exposure (B1 / B2)", "Mitigant"], rows, top=TOP, height=0.59 + 6 * 0.72,
            col_widths=[5.0, 5.4, 4.6], align="LLL", unit_col=None)
    s.card("Why the base does not build the overbuild", [
        ("", f("The grid connection stays at {m} MVA. Building the overbuild lifts the all-in cost to COD from €{a}m "
               "to €{c}m. In that case COD IRR is {i} and HTM IRR is {h}. The build decision passes to the buyer "
               "at RtB.", m=b("In_MVA", n0), a=pair("Out_Cost_To_COD_1", "Out_Cost_To_COD_2", eur, m2),
               c=stored("Cost_To_COD", "Overbuild built", eur, m2), i=stored("COD_IRR", "Overbuild built", pct),
               h=stored("HTM_IRR", "Overbuild built", pct)))],
        left=4.92, top=8.05, width=15.0, height=1.35)
    s.footnote(f("Shortfall tested {m} months after COD. Pro rata = shortfall share × the wind PB (GYIK Q28/30). "
                 "The full-call and overbuild-built figures are alternative cases.", m=b("In_Sanction_Lag", n0)))
    return d.save(out / "Bora_Bid_IC.pptx")
