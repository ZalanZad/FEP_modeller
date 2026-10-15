"""Bora non-delivery case – savings at COD vs the performance security called. Built to SPEC.md v1.

Cover | Summary | Inputs | Curves | Calc | Checks, annual 2031-2061 (community stream and the COD buyer's cash flows).
Calc rows carry E = per bid and F = both bids unless the label says otherwise. No macros (nothing is stored).
Curves holds the base-case and zero-bid operating cash flows read from the Bora model by extract_zero_bid.py (data/).
"""
import csv
import datetime as dt
from pathlib import Path

from fepm.xl.builder import Model
from fepm.xl.styles import Alignment, Font, FONT, SIZE, apply

D = dt.date
NAME = "Bora_NonDelivery"
BOTH = "{self.E}*In_Bids"                                    # F column: both bids
DATA = Path(__file__).resolve().parent / "data"
BLOCKS = {k: list(csv.DictReader(open(DATA / f"{f}.csv", newline=""))) for k, f in (("BC", "base_case"), ("ZB", "zero_bid"))}


# =========================================================================================== inputs
def inputs(m):
    i = m.inputs
    a = i.add_input
    a("COD", "Commercial operation date (COD)", "date", D(2031, 9, 30), source="Bora model Inputs (Bora!N16)")
    a("Life", "Community contribution term from COD", "years", 30, source="Bora model Inputs (E157)", fmt="int")

    i.section("Bids and commitments (per bid)")
    a("Bids", "Bids (BORA 1 and BORA 2)", "#", 2, source="SPEC §1", fmt="int")
    a("MVA", "Net connection capacity", "MVA", 100, source="Bora model Inputs", fmt="money")
    a("Overbuild", "Overbuild committed (ÖSZ-3)", "%", 0.51, source="Bora model Dashboard")
    a("BESS", "BESS power committed, share of MVA (ÖSZ-4)", "%", 0.51, source="Bora model Dashboard")
    a("BESS_Hours", "BESS duration", "h", 4, source="Bora model Dashboard", fmt="money1")
    a("Wind_ND", "Wind built if not delivered", "MW", 100, source="SPEC §3 (= MVA)", fmt="money1")

    i.section("Capex saved")
    a("Turbine", "Turbine supply, EU", "EUR k/MW", 600, source="Bora model Inputs (user 12-Oct-26)")
    a("Installation", "Turbine installation and erection", "EUR k/MW", 80, source="Bora model Inputs (old E163)")
    a("Foundations", "Foundations", "EUR k/MW", 90, source="Bora model Inputs (old E164)")
    a("BoP", "Balance of plant", "EUR k/MW", 130, source="Bora model Inputs (old E165:E166)")
    a("Owner", "Owner's costs", "EUR k/MW", 40, source="Bora model Inputs (old E169)")
    a("Land", "Land and easements", "EUR k/MW", 15, source="Bora model Inputs (old E170)")
    a("Contingency", "Contingency on overbuild capex", "%", 0.05, source="Bora model Inputs (old E171)")
    a("BESS_Capex", "BESS capex", "EUR k/MW", 240, source="Bora model Inputs (user 12-Oct-26)")
    a("White_Label", "EU to CN turbines and transformers, saving", "EUR k/MW", 100, source="User 15-Oct-26")

    i.section("Municipal and community commitments")
    a("Municipal", "Municipal one-off at COD (ÖSZ-1)", "EUR k/MW", 50, source="Bora model Dashboard")
    a("Community", "Community contribution (ÖSZ-2)", "EUR/MWh", 3, source="Bora model Dashboard")
    a("P50_Hours", "Net P50 hours", "h/yr", 2460, source="Bora model Inputs (old E179)", fmt="money")
    a("Curtailment", "Curtailment", "%", 0.02, source="Bora model Inputs (old E181)")
    a("Degradation", "Degradation per year", "%", 0.003, source="Bora model Inputs (old E182)", fmt="0.00%")
    a("Rate", "Discount rate (PV to COD)", "%", 0.08, source="Bora model Inputs (old E253)")
    a("CIT", "Corporate income tax", "%", 0.09, source="Bora model Inputs (old E12)")
    a("CIT_On", "Community PV after CIT (1 = yes)", "flag", 0, source="SPEC §3 (default off)", fmt="int")

    i.section("Performance security (PB)")
    a("Bench_Wind", "Security base: benchmark investment, wind", "mHUF/MW", 406, source="Kiírás Annex 2")
    a("Bench_BESS", "Security base: benchmark investment, BESS", "mHUF/MWh", 200, source="Kiírás Annex 2")
    a("PB_Rate", "Performance security, share of base", "%", 0.025, source="Kiírás XII.2", fmt="0.00%")
    a("FX", "FX rate", "HUF/EUR", 365, source="Bora model Inputs (old E11)", fmt="money")
    a("Tolerance_Cap", "Capacity shortfall tolerated without sanction", "%", 0.10, source="Kiírás XIII.15; GYIK 28, 44")
    a("Bora_PB", "PB per bid in the Bora model", "EUR k", 6993.56164383562,
      source="Bora model Dashboard (Out_Sanction_Full_1)", fmt="money1")
    a("Sanction_Lag", "PB called after COD", "months", 6, source="Kiírás XIII.17; Bora model Inputs", fmt="int")

    i.section("COD buyer (per bid)")
    a("COD_Price", "Purchase price at COD per built MW", "EUR k/MW", 2450, source="Bora model Inputs (In_COD_Price)")
    a("BC_Tax_Base", "Tax depreciation base, original plan (A)", "EUR k", float(BLOCKS["BC"][0]["tax_base"]),
      source="Bora model, base case, Calc_1 E122", fmt="money")
    a("ZB_Tax_Base", "Tax depreciation base, zero bid", "EUR k", float(BLOCKS["ZB"][0]["tax_base"]),
      source="Bora model, zero bid, Calc_1 E122", fmt="money")
    a("Dep_Years", "Tax depreciation period", "years", 20, source="Bora model Inputs (old E213)", fmt="int")


# =========================================================================================== curves
def curves(m):
    """Quarterly operating cash flows after COD, per bid (vertical blocks: date, revenue, opex, EBITDA)."""
    cv, ws = m.curves, m.curves.ws
    for key, title, case in (("BC", "Original plan (A, base case)", "base case"), ("ZB", "Zero bid", "zero bid")):
        cv.r += 1
        for col in range(1, m.last + 1):
            apply(ws.cell(cv.r, col), "section")
        ws.cell(cv.r, 1, "x")
        ws.cell(cv.r, 3, f"{title}: operating cash flows after COD, per bid")
        ws.cell(cv.r, 5, f"Bora model, {case}, Calc_1 rows 112-119")
        cv.r += 1
        apply(ws.cell(cv.r, 4, "EUR k"), "unit")
        for col, h in ((5, "Period end"), (6, "Revenue"), (9, "Opex"), (10, "EBITDA")):
            apply(ws.cell(cv.r, col, h), "header")
        first = cv.r + 1
        for d in BLOCKS[key]:
            cv.r += 1
            apply(ws.cell(cv.r, 5, D.fromisoformat(d["period_end"])), "input", "date")
            for col, k in ((6, "revenue"), (9, "opex"), (10, "ebitda")):
                apply(ws.cell(cv.r, col, float(d[k])), "input", "money1")
        for name, col in (("Date", "E"), ("Revenue", "F"), ("Opex", "I"), ("EBITDA", "J")):
            m.define(f"{key}_{name}", f"Curves!${col}${first}:${col}${cv.r}")
        cv.r += 1


# =========================================================================================== calc
def calc(m):
    c = m.calc
    c.section("Capacity per bid")
    c.row("wind", "Committed wind (built if delivered)", "MW", e="In_MVA*(1+In_Overbuild)", ef_fmt="money1")
    c.row("unbuilt", "Wind not built", "MW", e="{wind.E}-In_Wind_ND", ef_fmt="money1")
    c.row("bess", "BESS committed (E: MW, F: MWh)", "MW", e="In_BESS*In_MVA", f="{self.E}*In_BESS_Hours",
          ef_fmt="money1")
    c.row("pb_base", "Security base (E: wind, F: BESS)", "mHUF", e="{wind.E}*In_Bench_Wind",
          f="{bess.F}*In_Bench_BESS")

    c.section("Community contribution not paid (committed wind, per bid)")
    c.window("ops", "Operating period (COD to end of term)", "In_COD+1", "EDATE(In_COD,In_Life*Months_In_Year)")
    c.row("deg", "Degradation factor", "factor", "=(1-In_Degradation)^({year}-YEAR(In_COD))", total="none")
    c.row("gen", "Net generation", "MWh", "={wind.E}*In_P50_Hours*(1-In_Curtailment)*{deg}*{ops}*{yf}")
    c.row("df", "Discount factor to COD", "factor", "=1/(1+In_Rate)^(({end}-In_COD)/Days_In_Year)", total="none")
    c.row("com", "Community contribution not paid (E: PV by XNPV)", "EUR k",
          "=In_Community*{gen}*(1-In_CIT*In_CIT_On)/Thousand",
          e="XNPV(In_Rate,{self.all},{end.all})*INDEX({df.all},1)", ef_fmt="money")

    c.section("Savings at COD (E: per bid, F: both bids)")
    c.subheader("Overbuild capex not spent")
    for key, label, name in (("ob_tur", "Turbine supply, EU", "Turbine"),
                             ("ob_inst", "Installation and erection", "Installation"),
                             ("ob_fdn", "Foundations", "Foundations"), ("ob_bop", "Balance of plant", "BoP"),
                             ("ob_own", "Owner's costs", "Owner"), ("ob_land", "Land and easements", "Land")):
        c.row(key, label, "EUR k", e=f"In_{name}*{{unbuilt.E}}", f=BOTH)
    c.row("ob_cont", "Contingency", "EUR k", e="In_Contingency*SUM({ob_tur.E}:{ob_land.E})", f=BOTH)
    c.row("ob", "Overbuild capex", "EUR k",
          e="(In_Turbine+In_Installation+In_Foundations+In_BoP+In_Owner+In_Land)*{unbuilt.E}*(1+In_Contingency)",
          f=BOTH, bold=True)
    c.row("bess_capex", "BESS capex", "EUR k", e="In_BESS_Capex*{bess.E}", f=BOTH)
    c.row("municipal", "Municipal one-off", "EUR k", e="In_Municipal*In_MVA", f=BOTH)
    c.row("com_pv", "Community contribution (PV to COD)", "EUR k", e="SUMPRODUCT({com.all},{df.all})", f=BOTH)
    c.row("white", "EU to CN turbines and transformers (white labelling)", "EUR k", e="In_White_Label*In_Wind_ND",
          f=BOTH)
    c.row("total", "Total savings at COD", "EUR k", e="SUM({ob.E}:{white.E})", f=BOTH, bold=True)
    c.row("total_x", "Savings excluding the community PV", "EUR k", e="{total.E}-{com_pv.E}", f=BOTH)

    c.section("Performance security (E: per bid, F: both bids)")
    c.row("pb", "PB called (full call)", "EUR k", e="-In_PB_Rate*({pb_base.E}+{pb_base.F})*Thousand/In_FX", f=BOTH)
    c.row("net", "Net saving (savings less PB called)", "EUR k", e="{total.E}+{pb.E}", f=BOTH, bold=True)
    c.row("mult", "Savings ÷ PB called", "x", e="{total.E}/-{pb.E}", ef_fmt="mult")
    c.row("short", "Wind shortfall share", "%", e="{unbuilt.E}/{wind.E}", ef_fmt="pct")
    c.row("pb_pro", "PB lost, pro-rata reading (BESS in full, wind pro rata)", "EUR k",
          e="-In_PB_Rate*({short.E}*({short.E}>In_Tolerance_Cap)*{pb_base.E}+{pb_base.F})*Thousand/In_FX", f=BOTH)
    c.row("mva_lost", "Connection capacity lost, pro-rata reading", "MVA",
          e="({short.E}>In_Tolerance_Cap)*{short.E}*In_MVA", f=BOTH, ef_fmt="money1")
    c.row("excl", "3-year group exclusion", "text", e='IF({pb.E}<0,"Yes","No")', ef_fmt="text")

    c.section("COD buyer's project IRR, per bid (E: IRR, F: NPV at COD)")
    c.event("cod", "COD: purchase", "In_COD")
    c.event("sanc", "PB called (COD + lag)", "EDATE(In_COD,In_Sanction_Lag)")
    c.row("cf_date", "Cash-flow date (COD in the COD year)", "date", "=IF({cod},In_COD,{end})", total="none")
    sumifs = '=SUMIFS({0}_EBITDA,{0}_Date,">="&{{start}},{0}_Date,"<="&{{end}})'
    c.row("ebitda_a", "EBITDA, original plan (A)", "EUR k", sumifs.format("BC"))
    c.row("ebitda", "EBITDA, zero bid (B and reference)", "EUR k", sumifs.format("ZB"))
    c.row("dep", "Tax depreciation share", "factor",
          "=MAX(0,MIN({self.F},{end})-MAX({self.E},{start})+1)/({self.F}-{self.E}+1)", e="In_COD+1",
          f="EDATE(In_COD,In_Dep_Years*Months_In_Year)", ef_fmt="date")
    c.row("base", "Tax depreciation base, zero bid (F: non-delivery)", "EUR k", e="In_ZB_Tax_Base",
          f="{self.E}-{white.E}")
    tax = "={%s}-In_CIT*MAX(0,{%s}-%s*{dep})-In_COD_Price*In_Wind_ND*{cod}"
    irr = dict(e="XIRR({self.all},{cf_date.all})", f="XNPV(In_Rate,{self.all},{cf_date.all})", ef_fmt="money")
    c.row("cf_a", "Buyer cash flow: A, original plan, buyer bears the PB", "EUR k",
          tax % ("ebitda_a", "ebitda_a", "In_BC_Tax_Base") + "+{pb.E}*{sanc}", **irr, bold=True)
    c.row("cf_m", "Buyer cash flow: non-delivery, seller bears the PB", "EUR k",
          tax % ("ebitda", "ebitda", "{base.F}"), **irr)
    c.row("cf_n", "Buyer cash flow: B, non-delivery, buyer bears the PB", "EUR k", "={cf_m}+{pb.E}*{sanc}", **irr,
          bold=True)
    c.row("cf_o", "Buyer cash flow: zero bid, no call (reference)", "EUR k",
          tax % ("ebitda", "ebitda", "{base.E}"), **irr)
    c.row("d_irr", "IRR change, B less A", "%", e="{cf_n.E}-{cf_a.E}", ef_fmt="0.00%")


# =========================================================================================== checks
def checks(m):
    c, k = m.calc, m.checks
    first, last = c.rows["ob_tur"].r, c.rows["ob_cont"].r
    k.add_check("Overbuild capex: elements = total",
                f"=--(ABS(SUM(Calc!$E${first}:$E${last})-{{ob.E}})>Tolerance)", per_period=False)
    k.add_check("PB per bid = Bora model PB (EUR 6,993.6 k)", "=--(ABS(-{pb.E}-In_Bora_PB)>Tolerance)",
                per_period=False)
    k.add_check("Community PV: XNPV = SUMPRODUCT", "=--(ABS({com.E}-{com_pv.E})>Tolerance)", per_period=False)
    for key, case, row in (("BC", "Base-case", "ebitda_a"), ("ZB", "Zero-bid", "ebitda")):
        k.add_check(f"{case} data: revenue + opex = EBITDA",
                    f"=--(ABS(SUM({key}_Revenue)+SUM({key}_Opex)-SUM({key}_EBITDA))>Tolerance)", per_period=False)
        k.add_check(f"{case} EBITDA all on the timeline",
                    f"=--(ABS(SUM({key}_EBITDA)-{{{row}.I}})>Tolerance)", per_period=False)


# =========================================================================================== summary
SANCTION = [
    "Points commitments not proven by COD: full PB call, no discretion (Kiírás XIII.17; GYIK Q72).",
    "Capacity shortfall above 10 %: PB and connection right lost pro rata (XII.12; GYIK 28/30/44).",
    "A called PB excludes the whole group from capacity tenders for 3 years (VII.2 n; GYIK Q71).",
    "The exclusion hits the group at the time of the call: the buyer's group after a sale.",
]


def summary(m):
    """Savings table (E both bids, F per bid in EUR m, G per committed MW), memo and sanction note."""
    ws, c = m.summary.ws, m.calc
    for col in "EFG":
        ws.column_dimensions[col].width = 14
    r = [4]

    def bar(title, kind="section", headers=()):
        for col in range(1, 8):
            apply(ws.cell(r[0], col), kind)
        ws.cell(r[0], 1, "x" if kind == "section" else None)
        ws.cell(r[0], 3, title)
        r[0] += 1
        if headers:
            for col, h in zip(range(3, 8), headers):
                apply(ws.cell(r[0], col, h), "header")
            ws.cell(r[0], 3).alignment = ws.cell(r[0], 4).alignment = Alignment(horizontal="left")
            r[0] += 1

    def put(col, formula, fmt, bold=False):
        kind = "total" if bold else ("link" if "/" not in formula else "calc")
        apply(ws.cell(r[0], col, formula), kind, fmt)

    def line(label, unit=None, key=None, out=None, per_mw=True, bold=False, fmt="money1"):
        ws.cell(r[0], 3, label)
        if bold:
            ws.cell(r[0], 3).font = Font(name=FONT, size=SIZE, bold=True)
        if unit:
            apply(ws.cell(r[0], 4, unit), "unit")
        if key and unit == "EUR m":
            put(5, f"={c.ref(key, 'F')}/Thousand", fmt, bold)
            put(6, f"={c.ref(key, 'E')}/Thousand", fmt, bold)
            if per_mw:
                put(7, f"={c.ref(key, 'E')}/{c.ref('wind', 'E')}", "money", bold)
        elif key:
            put(5, f"={c.ref(key, 'F' if unit == 'MVA' else 'E')}", fmt, bold)
            if unit == "MVA":
                put(6, f"={c.ref(key, 'E')}", fmt, bold)
        if out:
            m.define(f"Out_{out}", f"Summary!$E${r[0]}")
        r[0] += 1

    bar("Savings at COD if no commitment is delivered", headers=("Item", "Unit", "Both bids", "Per bid", "EUR k/MW"))
    bar("Overbuild capex not spent (EUR k/MW: per committed MW)", "subheader")
    for key in ("ob_tur", "ob_inst", "ob_fdn", "ob_bop", "ob_own", "ob_land", "ob_cont"):
        line(c.rows[key].label, "EUR m", key)
    line("Overbuild capex", "EUR m", "ob", "Overbuild", bold=True)
    line("BESS capex", "EUR m", "bess_capex", "BESS")
    line("Municipal one-off", "EUR m", "municipal", "Municipal")
    line("Community contribution (PV at 8 % to COD)", "EUR m", "com_pv", "Community_PV")
    line("EU to CN turbines and transformers (white labelling)", "EUR m", "white", "White_Label")
    line("Total savings at COD", "EUR m", "total", "Savings_Total", bold=True)
    line("PB called (full call)", "EUR m", "pb", "PB_Called")
    m.define("Out_PB_Per_Bid", f"Summary!$F${r[0] - 1}")
    line("Net saving (savings less PB called)", "EUR m", "net", "Net_Saving", bold=True)
    r[0] += 1
    line("Savings excluding the community PV", "EUR m", "total_x", "Savings_Ex_Community")
    line("Savings ÷ PB called", "x", "mult", "Multiple", fmt="mult")
    line("Memo: PB lost under the pro-rata reading (XII.12)", "EUR m", "pb_pro", "PB_ProRata")

    r[0] += 1
    bar("COD buyer's project IRR (unlevered, post-tax; IRR per bid = both bids)",
        headers=("Item", "Unit", "A: original", "B: no delivery", "Zero bid"))
    for label, unit, col, out, fmt in (("Buyer project IRR", "%", "E", "IRR", "0.00%"),
                                       ("Buyer NPV at 8 % to COD, both bids", "EUR m", "F", "NPV", "money1")):
        ws.cell(r[0], 3, label)
        apply(ws.cell(r[0], 4, unit), "unit")
        for i, (key, case) in enumerate((("cf_a", "A"), ("cf_n", "B"), ("cf_o", "Zero"))):
            ref = c.ref(key, col)
            put(5 + i, f"={ref}" if col == "E" else f"={ref}*In_Bids/Thousand", fmt)
            m.define(f"Out_Buyer_{out}_{case}", f"Summary!${'EFG'[i]}${r[0]}")
        r[0] += 1
    line("Change in buyer IRR, B less A", "pp", "d_irr", "Buyer_IRR_Delta", fmt="0.00%")
    ws.cell(r[0], 3, "Memo: B if the seller bears the PB call")
    apply(ws.cell(r[0], 4, "%"), "unit")
    put(6, f"={c.ref('cf_m', 'E')}", "0.00%")
    m.define("Out_Buyer_IRR_Seller_PB", f"Summary!$F${r[0]}")
    r[0] += 1

    r[0] += 1
    bar("Memo")
    line("Points lost after the award: none, the award stands")
    line("3-year group exclusion, as the PB is called", "text", "excl", fmt="text")
    line("Connection right under a full call: kept (GYIK Q72)")
    line("Connection capacity lost, pro-rata reading (GYIK 28/30)", "MVA", "mva_lost", "MVA_Lost")

    r[0] += 1
    bar("Sanction")
    for text in SANCTION:
        line(text)
    m.summary.r = r[0]



def build_model(out: Path) -> Path:
    m = Model("Bora non-delivery case", "Savings at COD vs the performance security called", periods=31,
              periodicity="annual", start=D(2031, 1, 1), scenarios=["Base"], version="v1", author="FEP",
              date=D(2026, 10, 15), curves=True)
    inputs(m)
    curves(m)
    calc(m)
    checks(m)
    summary(m)
    m.summary.write_outputs = lambda: None          # own table above; no stored scenario panel (no macros)
    cover_finish = m.cover.finish

    def finish():                                    # plain .xlsx: drop the engine's macro line on the Cover
        cover_finish()
        m.cover.ws["B18"] = m.cover.ws["C18"] = None
    m.cover.finish = finish
    return m.save(out / f"{NAME}.xlsx")


def build(out: Path):
    out.mkdir(parents=True, exist_ok=True)
    return {"model": build_model(out)}
