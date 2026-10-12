"""Bora IC deck – 7 slides to SPEC.md §0 item 19, text from out/Bora_Bid_slide_text.md.

Every figure is `Deck.bind` on the packaged workbook (LibreOffice-recalculated when `package` saves it, so it holds
the same values as the audit's `_recalc.xlsx`). Composite texts use `Deck.f`, which records each bound part.
"""
from pathlib import Path

from fepm.ppt.deck import DASH, Deck, fmt_delta, fmt_eur_m, fmt_mult, fmt_num, fmt_pct

SECTIONS = ["Executive summary", "Commitments", "Economics – RtB exit", "Economics – COD and HTM",
            "Risks and mitigants", "Next steps"]
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


def date(v):
    return f"{v:%d-%b-%Y}"


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

    def b(ref, fmt):
        """Bind a workbook name (`Out_…`, `In_…`) or a `Sheet!Cell` reference."""
        return d.bind(*ref.split("!"), fmt=fmt) if "!" in ref else d.bind(ref, fmt=fmt)

    def pair(r1, r2, fmt, tpl="{a} / {b}"):
        return f(tpl, a=b(r1, fmt), b=b(r2, fmt))

    def dash(c):                 # Dashboard cells of BORA 1 / BORA 2 in columns c and c+1 (e.g. "I12")
        col, row = c[0], c[1:]
        return f"Dashboard!{col}{row}", f"Dashboard!{chr(ord(col) + 1)}{row}"

    # ---- 1 Cover -------------------------------------------------------------------------------------------
    d.cover("Project Bora", "BORA 1 and BORA 2 capacity bids", b("Dashboard!E6", date),
            subtitle=f("Approval of the tender commitments,\n2 × {m} MVA, Szombathely OVIT", m=b("In_MVA", n0)))

    # ---- 2 Executive summary ---------------------------------------------------------------------------------
    s = d.content("Executive summary", f(
        "Submit both bids at {p} points: FEP earns €{r}m at the RtB exit, a {i} IRR",
        p=pair("Out_Points_1", "Out_Points_2", pts),
        r=pair("Out_RtB_Profit_1", "Out_RtB_Profit_2", eur, "{a}m / €{b}"),
        i=pair("Out_RtB_IRR_1", "Out_RtB_IRR_2", pct)), f(
        "Against the zero bid, the commitments cost €{a}m of RtB profit and {c} pp of HTM IRR, "
        "and add €{e}m of commitment.",
        a=pair(*dash("I26"), eur_abs, "{a}m / €{b}"), c=pair(*dash("I30"), pp_abs),
        e=pair(*dash("I32"), eur_abs, "{a}m / €{b}")), title_lines=2)
    rows = [("Points (EU turbine)", "points", "H25", "Out_Points", pts),
            ("FEP profit – RtB route", "€ m", "H26", "Out_RtB_Profit", eur),
            ("FEP IRR – RtB route", "%", "H27", "Out_RtB_IRR", pct),
            ("FEP IRR – sale at COD", "%", "H28", "Out_COD_IRR", pct),
            ("FEP IRR – hold to maturity", "%", "H30", "Out_HTM_IRR", pct),
            ("Total commitment", "€ m", "H32", "Out_Commitment_Total", eur)]
    s.table(["Item", "Unit", "Zero bid", "BORA 1", "BORA 2"],
            [[lbl, unit, b(f"Dashboard!{zero}", fm), b(f"{out_}_1", fm), b(f"{out_}_2", fm)]
             for lbl, unit, zero, out_, fm in rows],
            key_rows={5}, top=TOP, width=SIDE_W, height=0.59 + 6 * 0.48, col_widths=[4.3, 1.1, 1.5, 1.5, 1.5])
    s.card("Approval requested", [
        (f("Submit BORA 1 and BORA 2 by {d} ", d=b("In_Bid_Date", date)), f(
            "with these commitments: municipal {m} k€/MW; community {c} €/MWh; overbuild {o}; "
            "BESS {e} of MVA, {h} h; PB at the {pb} minimum; partial award, shared connection, "
            "EU turbines and transformers.",
            m=b("Dashboard!E12", n0), c=pair(*dash("E13"), dp2), o=pair(*dash("E14"), pct0),
            e=pair(*dash("E15"), pct0), h=b("Dashboard!E16", n0), pb=b("Dashboard!E17", x1))),
        ("Main risk: ", f(
            "the overbuild is not built. That triggers a pro-rata PB call of €{p}m and loses {v} MVA of "
            "connection right. Building it removes both.",
            p=pair("Out_Commitment_Penalty_1", "Out_Commitment_Penalty_2", eur, "{a}m / €{b}"),
            v=pair("Out_MVA_At_Risk_1", "Out_MVA_At_Risk_2", n1)))],
        left=CARD_L, top=TOP, width=CARD_W, height=4.6)
    s.footnote(f("Post-tax, FEP share, EU turbine, base key-assumption set (overbuild not built, BESS built, "
                 "pro-rata sanction); RtB exit {x}. Zero bid = all commitments at their zero level.",
                 x=b("Dashboard!E60", date)))

    # ---- 3 Why these commitments -----------------------------------------------------------------------------
    s = d.content("Commitments", f(
        "Municipal and overbuild cost {a} and {o} pp of BORA 1 HTM IRR; BESS is the main return cost at {e} pp",
        a=b("Dashboard!O12", pp_abs), o=b("Dashboard!O14", pp_abs), e=b("Dashboard!O15", pp_abs)), f(
        "Each commitment is measured alone against the zero bid; the PB stays at the {x} minimum because "
        "{y} adds 5 points for one more PB of guarantee.",
        x=b("Dashboard!E17", x1), y=b("Curves!S19", x1)), title_lines=2)

    def impact(r):               # points, Δ RtB profit, Δ COD IRR, Δ HTM IRR, Δ securities of Dashboard row r
        return [pair(*dash(f"I{r}"), pts), pair(*dash(f"K{r}"), fmt_delta), pair(*dash(f"M{r}"), pp),
                pair(*dash(f"O{r}"), pp), pair(*dash(f"Q{r}"), fmt_delta)]

    def zero(rng):               # a block of Δ cells that are all zero, shown once
        return [b(f"Dashboard!{c}{rng}", fm) for c, fm in
                (("K", fmt_delta), ("M", pp), ("O", pp), ("Q", fmt_delta))]

    rng17 = ["Dashboard!K17:L17", "Dashboard!M17:N17", "Dashboard!O17:P17", "Dashboard!Q17:R17"]
    rng18 = ["Dashboard!K18:L22", "Dashboard!M18:N22", "Dashboard!O18:P22", "Dashboard!Q18:R22"]
    fmts = [fmt_delta, pp, pp, fmt_delta]
    rows = [
        ["Municipal one-off (ÖSZ-1)", pair(*dash("E12"), n0, "{a} / {b} k€/MW"), *impact(12),
         "Most points per bid; paid at COD, so no FEP cash before the RtB exit."],
        ["Community contribution (ÖSZ-2)", pair(*dash("E13"), dp2, "{a} / {b} €/MWh"), *impact(13),
         f("Operating cost after COD. At the band floor, BORA 2 earns {p} points.", p=b("Dashboard!J13", pts))],
        ["Overbuild (ÖSZ-3)", pair(*dash("E14"), pct0), *impact(14),
         "Points without capex: not built. The cost is the larger PB and its pro-rata call."],
        [f("BESS power (ÖSZ-4), {h} h", h=b("Dashboard!E16", n0)), pair(*dash("E15"), pct0, "{a} / {b} of MVA"),
         *impact(15), "Sized with the overbuild; built. Modelled as capex only, with no BESS revenue."],
        ["PB multiple (ÖSZ-6)", pair(*dash("E17"), x1), pair(*dash("I17"), pts),
         *[b(r, fm) for r, fm in zip(rng17, fmts)],
         f("Not raised: {y} adds 5 points for one more PB of €{a}m of guarantee.",
           y=b("Curves!S19", x1), a=pair(*dash("E52"), eur, "{a}m / €{b}"))],
        ["Partial award (ÖSZ-5), shared connection (ÖSZ-8), EU turbines / transformers (ÖSZ-9)", "yes",
         f("{a} / {b} / {c} / {e} each bid", a=b("Dashboard!I18", pts), b=b("Dashboard!I20", pts),
           c=b("Dashboard!I21", pts), e=b("Dashboard!I22", pts)),
         *[b(r, fm) for r, fm in zip(rng18, fmts)],
         f("No modelled cash or return cost. EU sourcing is in the base cost (turbine {t} k€/MW).",
           t=b("Dashboard!V18", n0))],
        ["Full bid vs zero bid", "", pair("Out_Points_1", "Out_Points_2", pts), pair(*dash("I26"), fmt_delta),
         pair(*dash("I28"), pp), pair(*dash("I30"), pp), DASH, "Rows do not add up to the total (interactions)."],
    ]
    s.table(["Commitment (B1 / B2)", "Offered", "Points", "Δ RtB profit € m", "Δ COD IRR pp", "Δ HTM IRR pp",
             "Δ securities € m", "Logic"], rows, key_rows={6}, top=TOP, height=0.59 + 7 * 0.66,
            col_widths=[3.25, 1.75, 1.3, 1.3, 1.2, 1.2, 1.3, 3.7], align="LRRRRRRL", unit_col=None)
    s.footnote(f("Δ = zero bid plus that commitment alone, minus the zero bid (macro RunCommitmentImpact). "
                 "Zero levels: municipal 0, community 0, overbuild 0%, BESS 0%, PB 1.0x, yes/no items 0. "
                 "ÖSZ-6 band: 5 points at {y} (Kiírás X.4 bf, Annex 4). Brownfield (ÖSZ-7) is not offered. "
                 "The model has no total Δ securities.", y=b("Curves!S19", x1)))

    # ---- 4 Securities and cash to the RtB exit ---------------------------------------------------------------
    s = d.content("Economics – RtB exit", f(
        "At the RtB exit FEP recovers its cash and earns €{p}m, a {m} MOIC",
        p=pair("Out_RtB_Profit_1", "Out_RtB_Profit_2", eur, "{a}m / €{b}"),
        m=pair("Out_RtB_MOIC_1", "Out_RtB_MOIC_2", fmt_mult)), f(
        "FEP funds €{o}m per bid to {x} and carries the PB guarantee. Of the €{t}m RtB price net of sale costs, "
        "FEP takes €{a}m.", o=b("Out_RtB_Outlay_1", eur_abs), x=b("Dashboard!E60", date),
        t=b("Out_Split_Total_1", eur), a=pair("Out_Split_FEP_1", "Out_Split_FEP_2", eur, "{a}m / €{b}")),
        title_lines=2)

    def split(key):              # "11.1 (73)": € m and k€ per installed wind MW
        return [f("{a} ({k})", a=b(f"Out_Split_{key}_{i}", eur), k=b(f"Out_Split_{key}_MW_{i}", n0))
                for i in (1, 2)]

    def both(key, fm):
        return [b(f"Out_{key}_1", fm), b(f"Out_{key}_2", fm)]

    rows = [["RtB price", "€ m", *both("RtB_Price", eur)],
            ["Receipts at the exit", "€ m", *both("FEP_Proceeds", eur)],
            ["Outlays: devex, land, deposits, connection fees, signing", "€ m", *both("RtB_Outlay", eur)],
            [f("Guarantee fees ({g} carry)", g=b("In_Guarantee_Fee", pct0)), "€ m", *both("Guarantee_Fees", eur)],
            ["Tax at the exit", "€ m", *both("RtB_Tax", eur)],
            ["FEP profit – RtB route", "€ m", *both("RtB_Profit_Buildup", eur)],
            [f("Exit split – FEP: claims + {s}", s=b("In_FEP_Share", pct0)), "€ m (k€/MW)", *split("FEP")],
            ["Exit split – A.RES: DSA + 30%", "€ m (k€/MW)", *split("ARES")],
            ["Exit split – total", "€ m (k€/MW)", *split("Total")]]
    s.table(["Item", "Unit", "BORA 1", "BORA 2"], rows, key_rows={5}, top=TOP, width=SIDE_W,
            height=0.59 + 9 * 0.46, col_widths=[5.06, 1.6, 1.6, 1.6])
    s.card("Securities to the exit", [
        ("Bid security ", f("€{a}m, guarantee, {p} to {r}.", a=pair(*dash("E51"), eur, "{a}m / €{b}"),
                            p=b("Dashboard!I51", date), r=b("Dashboard!J51", date))),
        ("PB ", f("€{a}m, guarantee from {p}; the buyer replaces it at the exit.",
                  a=pair(*dash("E52"), eur, "{a}m / €{b}"), p=b("Dashboard!I52", date))),
        ("Cash: ", f("TSO security €{t}m ({td}) and connection fees €{c1}m + €{c2}m ({cd}). "
                     "The buyer reimburses €{r}m at the exit.",
                     t=b("Dashboard!E53", eur), td=b("Dashboard!I53", date), c1=b("Dashboard!E54", eur),
                     c2=b("Dashboard!E55", eur), cd=b("Dashboard!I54", date), r=b("Out_WF_Principal_1", eur)))],
        left=CARD_L, top=TOP, width=CARD_W, height=4.0)
    s.footnote(f("RtB price {p} k€/MVA × {m} MVA. Receipts = recoverable principal, land and claims, plus {s} of "
                 "the distributable. Per-MW figures use installed wind of {w} MW. The participation fee "
                 "(€{fee}m per bid, cash) is not recovered.",
                 p=b("Dashboard!V16", n0), m=b("In_MVA", n0), s=b("In_FEP_Share", pct0),
                 w=pair("Out_Wind_MW_1", "Out_Wind_MW_2", n0), fee=b("Dashboard!E50", eur3)))

    # ---- 5 COD and HTM returns and the commitment ------------------------------------------------------------
    s = d.content("Economics – COD and HTM", f(
        "Held to maturity, both bids return above the {r} discount rate: HTM IRR {i}, NPV €{n}m",
        r=b("Dashboard!V22", pct), i=pair("Out_HTM_IRR_1", "Out_HTM_IRR_2", pct),
        n=pair("Out_HTM_NPV_1", "Out_HTM_NPV_2", eur, "{a}m / €{b}")), f(
        "Against the zero bid, the commitments cost {c} pp of COD IRR and {h} pp of HTM IRR, "
        "and raise the total commitment to €{t}m.",
        c=pair(*dash("I28"), pp_abs), h=pair(*dash("I30"), pp_abs),
        t=pair("Out_Commitment_Total_1", "Out_Commitment_Total_2", eur, "{a}m / €{b}")), title_lines=2)

    def ret(lbl, unit, r, key, fm, dfm):
        return [lbl, unit, b(f"Dashboard!H{r}", fm), *both(key, fm), *[b(c, dfm) for c in dash(f"I{r}")]]

    def memo(lbl, key):
        return [lbl, "€ m", DASH, *both(key, eur), DASH, DASH]

    rows = [ret("FEP IRR – sale at COD", "% / pp", 28, "COD_IRR", pct, pp),
            ret("FEP NPV – sale at COD", "€ m", 29, "COD_NPV", eur, fmt_delta),
            ret("FEP IRR – hold to maturity", "% / pp", 30, "HTM_IRR", pct, pp),
            ret("FEP NPV – hold to maturity", "€ m", 31, "HTM_NPV", eur, fmt_delta),
            memo("Securities: cash posted and guarantees", "Commitment_Securities"),
            memo("Devex, land, signing and DD", "Commitment_Devex"),
            memo("Construction capex incl. connection fees and municipal", "Commitment_Construction"),
            memo("Financing: IDC, debt fees, guarantee fees", "Commitment_Financing"),
            ret("Total commitment", "€ m", 32, "Commitment_Total", eur, fmt_delta),
            memo("Memo, not in the total: PB call, overbuild not built", "Commitment_Penalty")]
    s.table(["Item", "Unit", "Zero bid", "BORA 1", "BORA 2", "Δ B1", "Δ B2"], rows, key_rows={8}, top=TOP,
            width=SIDE_W, height=0.59 + 10 * 0.44, col_widths=[4.06, 0.9, 1.0, 1.0, 1.0, 0.98, 0.98])
    s.card(None, [
        ("BESS is the largest single driver: ", f("{a} pp of COD IRR on its own, modelled as capex with no revenue.",
                                                  a=pair(*dash("M15"), pp))),
        ("Sale at COD: ", f("FEP proceeds of €{p}m on a €{c}m price, a {m} MOIC.",
                            p=pair("Out_COD_Proceeds_1", "Out_COD_Proceeds_2", eur, "{a}m / €{b}"),
                            c=b("Out_COD_Price_1", eur), m=pair("Out_COD_MOIC_1", "Out_COD_MOIC_2", fmt_mult)))],
        left=CARD_L, top=TOP, width=CARD_W, height=2.6)
    s.footnote(f("NPV at {r}. COD {cod}; COD price {p} k€/MW on {mw} MW built; PPA {ppa} €/MWh; gearing {g}. "
                 "HTM includes the buy-out of the developer at the RtB exit date. The zero bid's total is shown "
                 "without a breakdown.",
                 r=b("Dashboard!V22", pct), cod=b("Dashboard!V23", date), p=b("Dashboard!V17", n0),
                 mw=b("Out_Wind_Built_MW_1", n0), ppa=b("Dashboard!V19", n0), g=b("Dashboard!V20", pct)))

    # ---- 6 Risks and mitigants -------------------------------------------------------------------------------
    s = d.content("Risks and mitigants", f(
        "The main risk is the unbuilt overbuild: a PB call, a 3-year group exclusion and {v} MVA of "
        "connection right lost", v=pair("Out_MVA_At_Risk_1", "Out_MVA_At_Risk_2", n1)),
        "The base COD and HTM returns include the pro-rata call; the lost MVA and any deduction from the RtB "
        "price are not modelled.", title_lines=2)
    m2 = "{a}m / €{b}"
    rows = [
        ["PB call on the unbuilt overbuild: pro rata (Kiírás XII.12) or full (XIII.17, GYIK Q72)",
         f("Pro rata €{p}m (base). Full call €{c}m: COD IRR {i}, HTM IRR {h}",
           p=pair("Out_Commitment_Penalty_1", "Out_Commitment_Penalty_2", eur, m2),
           c=pair("Out_Sanction_Full_1", "Out_Sanction_Full_2", eur, m2),
           i=pair("Dashboard!J107", "Dashboard!J108", pct), h=pair("Dashboard!J109", "Dashboard!J110", pct)),
         f("PB at the {x} minimum. Part of the call may be deducted from the RtB price (not modelled).",
           x=b("Dashboard!E17", x1))],
        ["3-year group exclusion (Kiírás VII.2 n, GYIK Q71)",
         f("Triggered by any PB call, including the base pro-rata call (flag {e}). The call falls after COD, "
           "so it is likely to hit the group then holding the SPV.",
           e=pair("Out_Exclusion_1", "Out_Exclusion_2", n0)),
         "Allocate it in the RtB SPA."],
        ["Connection right lost pro rata (Kiírás XII.12, GYIK Q28)",
         f("{v} MVA. This is not reflected in generation or price.",
           v=pair("Out_MVA_At_Risk_1", "Out_MVA_At_Risk_2", n1)),
         f('Building the overbuild removes it ("Overbuild built" set: COD IRR {i}, HTM IRR {h}).',
           i=pair("Dashboard!I107", "Dashboard!I108", pct), h=pair("Dashboard!I109", "Dashboard!I110", pct))],
        [f("{c} MVA group cap per node (Kiírás V.7, GYIK Q52)", c=b("In_Group_Cap", n0)),
         "Both bids are invalid if FEP, A.RES and the SPVs count as one company group.",
         "Legal confirmation before submission."],
        ["ÖSZ-4 scored on installed wind (GYIK Q20)",
         f("Points fall to {p}; returns unchanged.", p=pair("Dashboard!M103", "Dashboard!M104", dp2)),
         "Ranking risk only; the commitment is the same."],
        ["BESS modelled as capex only",
         f("{a} pp of COD IRR, with no BESS revenue or sale value.", a=pair(*dash("M15"), pp)),
         "Conservative. A BESS revenue case is not modelled."],
    ]
    s.table(["Risk", "Exposure (B1 / B2)", "Mitigant"], rows, top=TOP, height=0.59 + 6 * 0.78,
            col_widths=[4.4, 5.8, 4.8], align="LLL", unit_col=None)
    s.footnote(f("Sanction tested at COD + {m} months. Pro rata = shortfall share × the wind PB (GYIK Q28/Q30). "
                 'The full-call and "Overbuild built" figures are stored key-assumption sets '
                 "(macro RunAllScenarios).", m=b("In_Sanction_Lag", n0)))

    # ---- 7 Approval requested and next steps -----------------------------------------------------------------
    s = d.content("Next steps", f("Approval requested: submit BORA 1 and BORA 2 by {d}", d=b("In_Bid_Date", date)),
                  f("Submission ties up a €{a}m bid security; on award, the PB of €{p}m replaces it.",
                    a=pair(*dash("E51"), eur, m2), p=pair(*dash("E52"), eur, m2)))
    s.card(None, [
        ("Approve the commitments on slide 3: ", f("{p} points, EU turbine.",
                                                   p=pair("Out_Points_1", "Out_Points_2", pts))),
        ("Before submission: ", f("obtain legal confirmation of the {c} MVA group-cap position.",
                                  c=b("In_Group_Cap", n0))),
        (f("By {d}: ", d=b("In_Bid_Date", date)),
         f("pay the participation fee (€{fee}m per bid, cash) and post the bid security (€{a}m, guarantee).",
           fee=b("Dashboard!E50", eur3), a=pair(*dash("E51"), eur, m2))),
        (f("On award ({d}): ", d=b("In_Award_Date", date)),
         f("post the PB (€{p}m) by {pd} and the TSO security (€{t}m) by {td}. Agree with the RtB buyer whether "
           "to build the overbuild and how to price the PB call.",
           p=pair(*dash("E52"), eur, m2), pd=b("Dashboard!I52", date), t=b("Dashboard!E53", eur),
           td=b("Dashboard!I53", date)))],
        left=4.92, top=1.95, width=11.0, height=3.4)
    s.footnote("Dates from the tender timetable in the model inputs.")
    return d.save(out / "Bora_Bid_IC.pptx")
