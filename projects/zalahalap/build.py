"""Zalahaláp – 63 MWp PV, with or without a 50 MW BESS: investment model, SPEC v2.6 (wind-model format).

One workbook, `out/Zalahalap_Model.xlsx` (no macros): Dashboard · Inputs · Revenue · Calc · Checks, annual "to 2026",
2027 … 2064, three cases as columns (E PV only, F PV + 2 h BESS, G PV + 4 h BESS), unlevered and levered. In the fixed
term from COD the PV sells at a fixed €/MWh on all output and the BESS earns a fixed daily spread on its nameplate
energy (price × MWh × cycles × 365, or a yearly income override); after it both are merchant on Aurora central (BESS:
block uplift × the PV-only factor). Base / Low / High through `Dash_Case`. The pre-COD cash is dated on the Dashboard
(RtB route, XIRR) and mapped to the annual grid for the hold and COD routes. Every XIRR carries a start guess (Calc
column H): Excel's XIRR returns ~0 when it fails.
"""
from __future__ import annotations

import csv
import datetime as dt
from pathlib import Path

from openpyxl.styles import Font
from openpyxl.utils import get_column_letter as L

from fepm.xl.styles import FONT, SIZE
from fepm.xl.wind import Block, Book, DatedBlock, check_format

HERE = Path(__file__).resolve().parent
DATA = HERE / "data"
NAME = "Zalahalap_Model"
D = dt.date
FIRST_YEAR, N_YEARS = 2026, 39                          # "to 2026", 2027 … 2064 (Low: COD +12 months runs to 2064)
YEARS = list(range(FIRST_YEAR, FIRST_YEAR + N_YEARS))
CASES = ["PV only", "PV + 2h BESS", "PV + 4h BESS"]
BLOCKS = {"DC 4h": "2027 solar DC 1.0cy 4h", "DC 2h": "2027 solar DC 1.0cy 2h"}     # code -> Aurora block
PH = "placeholder – pending user"
EW_NOTE = "to be revisited, user 22-Oct-26"
TITLE = "Zalahaláp 63 MWp PV ± 50 MW BESS: RtB, COD or hold"

# dated tables (Inputs, flat rows): label, unit, amount, date, source
EARLY_WORKS = [("Land, 1st payment", "HUF m", 60, D(2025, 11, 30), "ZHLP v3 Stat_inp E44"),
               ("Land, 2nd payment", "HUF m", 60, D(2026, 2, 28), "ZHLP v3 Stat_inp E45"),
               ("ICS execution design, 1st payment", "EUR k", 25, D(2026, 2, 28), "ZHLP v3 Stat_inp E46"),
               ("Critical equipment, 1st payment (HV, ICS, transformer)", "EUR k", 460, D(2026, 4, 30),
                "ZHLP v3 Stat_inp E47")]
DEPOSITS = [("Connection-fee reimbursement", 45, D(2025, 11, 30), None, "ZHLP v3 Stat_inp E92; Dyn_Inp row 12"),
            ("Capacity guarantee, cash", 180, D(2025, 11, 30), None, "ZHLP v3 Stat_inp E93; Dyn_Inp row 13"),
            ("Additional guarantee, cash", 485.55, D(2025, 11, 30), D(2026, 1, 31),
             "ZHLP v3 Stat_inp E94; Dyn_Inp row 14")]
GUARANTEES = [("Additional guarantee, first half", 242.775, D(2026, 1, 31), D(2026, 3, 31),
               "ZHLP v3 Stat_inp E95; Dyn_Inp row 15"),
              ("Additional guarantee, second half", 242.775, D(2026, 1, 31), None, "ZHLP v3 Stat_inp E95"),
              ("Additional guarantee, final instalments", 80.925, D(2026, 4, 30), None,
               "ZHLP v3 Stat_inp E96; Dyn_Inp row 16")]
DSA_FIRST, RTB = D(2025, 12, 31), D(2026, 12, 31)
PRE_FIRST, PRE_LAST = D(2026, 10, 31), D(2027, 3, 31)


def _months(a, b):
    return (b.year - a.year) * 12 + b.month - a.month + 1


# =========================================================================================== Aurora
def aurora():
    out = {}
    for r in csv.DictReader(open(DATA / "aurora_curves.csv")):
        out[(r["series"], r["block"])] = ([float(r[str(y)]) for y in range(2026, 2061)], r["source"])
    return out


def _years(vals, grow=None):
    """Aurora 2026–2060 → model years 2026–2064 (the first column "to 2026" carries 2026), flat real after 2060 (or
    grown by `grow`)."""
    out = list(vals)
    while len(out) < N_YEARS:
        out.append(out[-1] * (1 + grow) if grow else out[-1])
    return out


# =========================================================================================== Dashboard top
def dashboard_top(book):
    d = book.sheet("Dashboard", TITLE)
    d.master_check()
    book.case_row = 6
    d.bar("Key inputs: the cases", e=_case_links(book, d))
    lv = d.lever
    lv("Name", "Case", "text", CASES, source="User 03-Nov-2026: PV only, 2 h, 4 h")
    lv("MWp", "PV capacity", "MWp", [63, 63, 63], source="Teaser p1")
    lv("MVA", "Grid export capacity", "MVA", [49.9, 49.9, 49.9], source="Teaser p1")
    lv("Yield", "Specific yield", "MWh/MWp", [1265, 1265, 1265], source="Teaser p1 (79.7 GWh)")
    lv("Clipping", "Clipping and curtailment (63 MWp on 49.9 MVA)", "%", [0.02, 0, 0],
       source="ZHLP v3 fix 5 (2 %); DC-coupled BESS 0; " + PH)
    lv("BESS_On", "BESS built (1 = yes)", "flag", [0, 1, 1], source="User 03-Nov-2026", options=[0, 1])
    lv("BESS_MW", "BESS power", "MW", [0, 50, 50], source="Teaser p1")
    lv("Hours", "BESS duration (2 or 4 hours)", "h", [0, 2, 4], source="User 03-Nov-2026", options=[0, 2, 4])
    d.row("MWh", "BESS energy", "MWh", cases="={BESS_On}*{BESS_MW}*{Hours}")
    book.define("BESS_MWh", d.ref("MWh", "cases"))
    d.row("block", "Aurora block (PV prices, BESS volumes; PV only: 2 h)", "text",
          cases='=IF({BESS_On}=1,"DC "&{Hours}&"h",PV_Only_Block)', fmt="general")
    book.define("Block_Code", d.ref("block", "cases"))
    usr = "User 06-Nov-2026"
    lv("PV_Price", "PV fixed price, nominal, on all PV output", "EUR/MWh", [45, 45, 45], source=usr, fmt="money2")
    lv("Fix_Years", "Fixed term from COD, PV and BESS (then merchant)", "years", [8, 8, 8], source=usr)
    lv("BESS_Price", "BESS fixed price, €/MWh of nameplate energy (0 = none)", "EUR/MWh", [0, 70, 55], source=usr,
       fmt="money2")
    lv("Fix_Cycles", "Fixed cycles per day (fixed BESS income)", "#", [0, 1, 1],
       source=usr + ": price × MWh × cycles × 365 days (user wrote 360 loosely)", fmt="money1")
    lv("BESS_Income", "BESS yearly income override (0 = from the €/MWh)", "EUR k/yr", [0, 0, 0],
       source="Full-year income; × operating share × in service")
    note = "Information: fixed income is contractual on nameplate energy; dispatch risk is the counterparty's"
    for key, label, unit, src in (("o_inc", "BESS yearly income at the fixed price, first full year", "EUR k", None),
                                  ("o_dis", "BESS discharge, first full year", "GWh", note),
                                  ("o_cyc", "BESS cycles per day, first full year", "#", note)):
        d.row(key, label, unit, cases="=0", fmt="money2" if key == "o_cyc" else None, source=src,
              src_col="J")                                                          # links: dashboard()
    lv("RtB_Price", "RtB price, premium (0 = implied)", "EUR k/MWp", [90, 90, 90],
       source="ZHLP v3 Stat_inp E24; " + PH)
    lv("COD_Price", "COD price (0 = implied)", "EUR k/MWp", [0, 0, 0], source="Implied at the buyer target IRR")
    fin = "User instruction 04-Nov-2026, not a bank term"
    lv("Gearing", "Gearing: debt share of capex to COD (0 = unlevered)", "%", [0.6, 0.6, 0.6], source=fin)
    lv("Debt_Rate", "Debt rate, all-in", "%", [0.05, 0.05, 0.05], source=fin)
    lv("Tenor", "Debt tenor: level annuity from the COD year end", "years", [20, 20, 20],
       source="Annuity; straight-line is the alternative (not built)")
    book.lever_rows = dict(d.rows)
    d.blank()
    cb = d.bar("Sensitivity case: Base, Low, High", e=["Base", "Low", "High"], h="Live", key="case_bar")
    r = d.row("case_sel", "Case selected (Base, Low or High)", "case", h="Base", inputs=("H",))
    d.ws.cell(r, 3).font = Font(name=FONT, size=SIZE, bold=True)
    d.validate(f"H{r}", "Base,Low,High")
    book.define("Dash_Case", f"Dashboard!$H${r}")
    live = f"=INDEX({{self.cases}},MATCH(Dash_Case,$E${cb}:$G${cb},0))"
    for key, label, unit, vals, src, fmt in [
            ("PV_Price", "PV price change, merchant years", "%", [0, -0.2, 0.2], "SPEC v2 §7", None),
            ("BESS", "BESS revenue change", "%", [0, -0.2, 0.2], "SPEC v2 §7", None),
            ("Yield", "Yield change", "%", [0, -0.02, 0.02], "SPEC v2 §7", None),
            ("Capex", "Capex change", "%", [0, 0.1, -0.05], "SPEC v2 §7", None),
            ("COD_Delay", "COD delay (capex stays in the capex year)", "months", [0, 12, 0], "SPEC v2 §7, §12.1", None),
            ("Buyer_IRR", "Buyer target IRR, unlevered after tax", "%", [0.08, 0.09, 0.07],
             "ZHLP v3 Control D9; " + PH, None)]:
        d.row("s_" + key, label, unit, e=vals[0], f=vals[1], g=vals[2], h=live, inputs=("E", "F", "G"), source=src,
              src_col="J", fmt=fmt)
        book.define(f"Dash_Case_{key}", d.ref("s_" + key, "cases"))
        book.define(f"Case_{key}", d.ref("s_" + key, "H"))
    return d


def _case_links(book, sh):
    p = "" if sh.ws.title == "Dashboard" else "Dashboard!"
    return [f"={p}{c}${book.case_row}" for c in "EFG"]


# =========================================================================================== Inputs
def inputs(book):
    i = book.sheet("Inputs", "Zalahaláp – inputs (blue), units and sources",
                   widths={"C": 54, "D": 11, "H": 60})
    i.bar("Inputs", e=_case_links(book, i), h="Source")
    a = i.input

    i.sub("Constants")
    a("Thousand", "Thousand", "#", 1000, "Scale")
    a("Days_In_Year", "Days in a year", "days", 365, "Day count act/365")
    a("Months_In_Year", "Months in a year", "months", 12, "Calendar")
    a("Hours_In_Year", "Hours in a year", "h", 8760, "Grid export check (365 × 24)")
    a("Tolerance", "Check tolerance", "#", 0.000001, "Checks", fmt="general")
    a("First_Year", "First grid column: spend to the model date", "year", FIRST_YEAR, "SPEC v2.1 §12.1 ('to 2026')")

    i.sub("Dates")
    a("Acq_Date", "Acquisition, JV start, deposits paid", "date", D(2025, 11, 30), "ZHLP v3 Stat_inp E39")
    a("RtB_Date", "Ready to build (last DSA draw)", "date", RTB, "Teaser p6; " + PH)
    a("Model_Date", "Model date (spend to date before it)", "date", D(2026, 12, 31), "User 04-Nov-2026 (= RtB)")
    a("Exit_Date", "RtB sale closing (exit)", "date", D(2027, 3, 31), PH)
    a("NTP_Date", "Notice to proceed (construction start)", "date", D(2027, 1, 1), "User 04-Nov-2026: costs in 2027")
    a("Capex_Year", "Construction capex paid in (all of it)", "year", 2027, "User 04-Nov-2026: costs in 2027")
    a("COD_Date", "COD, PV and BESS (first day of production)", "date", D(2028, 3, 1),
      "User 04-Nov-2026: production from Mar-2028")
    a("Life", "Operating life", "years", 35, "ZHLP v3 Control D30")
    a("Aurora_First", "Aurora's first year (2025 blank; flat real after 2060)", "year", 2026, "Aurora Q2-26 databook")

    i.sub("PV")
    a("PV_Avail", "PV availability", "%", 0.995, "ZHLP v3 Control D56")
    a("Degradation", "PV degradation a year (from the first full year)", "%", 0.003, "ZHLP v3 Control D57",
      fmt="0.00%")
    a("GoO_Price", "Guarantee of origin price, real 2025", "EUR/MWh", 1.75, "ZHLP v3 Control D58; " + PH, fmt="money2")
    a("PV_Only_Block", "PV-only price: Aurora block used as capture proxy", "text", "DC 2h",
      "Proxy: no solar capture in Aurora; series requested", text=True)

    i.sub("BESS (the Aurora blocks carry RTE 0.86 and availability 98 %)")
    a("BESS_RTE", "Round-trip efficiency (discharge ≤ PV charged × RTE)", "factor", 0.86, "Aurora Inputs C17")
    a("BESS_Haircut", "Haircut on the merchant BESS revenue", "%", 0, PH)
    a("AFRR_Share", "aFRR-down / FCR capacity deliverable from PV charge", "%", 0, "SPEC v1 §0.6; " + PH)
    a("BESS_Capex", "BESS capex and augmentation, real 2025", "EUR k/MWh", 110, "ZHLP v3 Control D94; " + PH)
    a("Repower", "BESS repowered at the repower year (0 = retired)", "flag", 1,
      "Aurora repower count s+1; ZHLP v3 Control D100")
    a("BESS_Opex_MW", "BESS fixed opex per MW, real 2025", "EUR k/MW/yr", 18.70,
      "Aurora Inputs r38-39 (2028): 2h 30.83, 4h 42.96 €/kW")
    a("BESS_Opex_MWh", "BESS fixed opex per MWh, real 2025", "EUR k/MWh/yr", 6.07,
      "Aurora Inputs r38-39: (42.96 − 30.83) ÷ 2 h")

    i.sub("Merchant years (after the fixed term)")
    a("Balancing", "Balancing cost, merchant years only, real 2025", "EUR/MWh", 5, "ZHLP v3 Control D60")

    i.sub("Capex to COD (2025 prices)")
    a("PV_Capex", "PV equipment and balance of plant", "EUR k/MWp", 350, "ZHLP v3 Control D82:D83 (330 + 20)")
    a("Cable_Capex", "MV cable to the Zalahaláp substation (12 km)", "EUR k", 3700, "ZHLP v3 Control D104; " + PH)
    a("Substation", "HV substation and bay", "EUR k", 9500, "ZHLP v3 Control D105; " + PH)
    a("Grid_Fee", "Grid connection fee", "EUR k", 0, "Teaser p1 (2018 rights)")
    a("Land_Capex", "Land to COD (not depreciated)", "EUR k", 620, "ZHLP v3 Control D106")
    a("EW_NTP", "Early works at NTP (E48 650 + E49 75)", "EUR k", 725, "ZHLP v3 Stat_inp E48:E49; " + EW_NOTE)

    i.sub("Opex (real 2025, × the Aurora euro index)")
    a("OM_Cost", "O&M, all-in (PV, BoP, MV cable, insurance, management)", "EUR k/MWp/yr", 10,
      "User 05-Nov-2026: 10 k€/MWp/yr all-in (v3: 14.96)")
    a("Lease", "Land lease from COD", "EUR k", 96, "ZHLP v3 Control D61")
    a("HIPA", "Local business tax (HIPA), on revenue", "%", 0.02, "Bana; " + PH)
    a("Innovation", "Innovation contribution, on revenue", "%", 0.003, "Act LXXVI/2014; " + PH, fmt="0.00%")

    i.sub("Development and JV")
    a("SPV_Price", "FEP SPV acquisition price (to FEP before the split)", "EUR k", 549, "JV sheet 221025 G19; " + PH)
    a("Transaction", "Transaction costs (FEP)", "EUR k", 200, "ZHLP v3 Stat_inp E37")
    a("DSA_Per_MWp", "DSA loan to PSH", "EUR k/MWp", 15, "JV sheet 221025 F55; " + PH)
    a("DSA_First", "DSA: first monthly draw (equal draws to RtB)", "date", DSA_FIRST, "v1 linear spread; " + PH)
    a("DSA_Months", "DSA: number of monthly draws (to RtB)", "months", _months(DSA_FIRST, RTB), "v1 linear spread; " + PH)
    a("PreRtB_Total", "Pre-RtB development costs (equal monthly)", "EUR k", 3182, "ZHLP v3 Control D115")
    a("PreRtB_Months", "Pre-RtB development: number of months", "months", _months(PRE_FIRST, PRE_LAST),
      "ZHLP v3 Control D116:D117")
    a("PreRtB_First", "Pre-RtB development: first month end", "date", PRE_FIRST, "ZHLP v3 Control D116")
    a("PreRtB_Last", "Pre-RtB development: last month end", "date", PRE_LAST, "ZHLP v3 Control D117")
    a("FX", "FX rate", "HUF/EUR", 400, "ZHLP v3 Stat_inp E23 (Aurora 397.77)")
    a("BG_Base", "Bank guarantee fee: 1M EURIBOR", "%", 0.0187, "ZHLP v3 Dyn_Inp G50:Y50", fmt="0.00%")
    a("BG_Margin", "Bank guarantee fee: margin", "%", 0.0175, "ZHLP v3 Stat_inp E110", fmt="0.00%")
    a("Draw_Timing", "Debt drawn mid-year: share of a year's IDC on draws", "%", 0.5, "Pro rata with capex in 2027")
    a("XIRR_Guess", "XIRR start guess (MOIC − 1 when lower)", "%", 0.1, "Excel returns ~0 when XIRR fails to converge")
    a("Loan_Rate", "FEP return on loans and costs (act/365, 31-Dec)", "%", 0.12, "JV sheet G30; v3 Stat_inp E69")
    a("Share_Low", "FEP share of the distributable, price ≤ threshold", "%", 0.6, "User 22-Oct-26")
    a("Share_High", "FEP share of the distributable, price > threshold", "%", 0.5, "User 22-Oct-26")
    a("Share_Threshold", "Share threshold (60 % at or below)", "EUR k/MWp", 90, "User 22-Oct-26: 60 % at exactly 90")
    a("PSH_Land", "PSH land apport (in kind)", "EUR k", 150, "ZHLP v3 Stat_inp E34")
    a("PSH_Devex", "PSH own development costs (memo)", "EUR k/MWp", 15, "ZHLP v3 Stat_inp E79")

    i.sub("Early works lines (paid in development; " + EW_NOTE + ")", extra={"F": "Paid"})
    rows = []
    for label, unit, amt, date, src in EARLY_WORKS:
        rows.append(i.row(None, label, unit, e=amt, f=date, inputs=("E", "F"), source=src,
                          fmts={"F": "date"}))
    book.ew_rows = rows
    i.sub("Cash deposits (HUF m; released blank = at the exit, buyer reimburses)", extra={"F": "Paid",
                                                                                        "G": "Released"})
    first = i.r
    for label, huf, paid, rel, src in DEPOSITS:
        i.row(None, label, "HUF m", e=huf, f=paid, g=rel, inputs=("E", "F", "G"), source=src,
              fmts={"F": "date", "G": "date"})
    book.dep_rows = list(range(first, i.r))
    for nm, col in (("Dep_HUF", "E"), ("Dep_Paid", "F"), ("Dep_Rel", "G")):
        book.define(nm, f"Inputs!${col}${first}:${col}${i.r - 1}")
    i.sub("Bank guarantees (HUF m; fee on the balance; released blank = at the exit)",
          extra={"F": "Issued", "G": "Released"})
    first = i.r
    for label, huf, paid, rel, src in GUARANTEES:
        i.row(None, label, "HUF m", e=huf, f=paid, g=rel, inputs=("E", "F", "G"), source=src,
              fmts={"F": "date", "G": "date"})
    book.bg_rows = list(range(first, i.r))
    for nm, col in (("BG_HUF", "E"), ("BG_Paid", "F"), ("BG_Rel", "G")):
        book.define(nm, f"Inputs!${col}${first}:${col}${i.r - 1}")

    i.sub("Return and reimbursement flags (1 = yes)")
    for nm, label, v, src in [
            ("Ret_SPV", "Earns the 12 % return: SPV price", 0, "JV sheet: price item, no interest"),
            ("Ret_Trans", "Earns the 12 % return: transaction costs", 0, "FEP's own cost"),
            ("Ret_Dep", "Earns the 12 % return: cash deposits", 1, "ZHLP v3 Calculation rows 75-78"),
            ("Ret_DSA", "Earns the 12 % return: DSA loan", 1, "JV sheet 221025 G30"),
            ("Ret_EW", "Earns the 12 % return: early works", 1, "ZHLP v3 rows 75-78; " + EW_NOTE),
            ("Ret_PreRtB", "Earns the 12 % return: pre-RtB development", 0, "ZHLP v3 / v1: reimbursed at cost"),
            ("Ret_Fees", "Earns the 12 % return: guarantee fees", 0, "Fees paid at release (simplification)"),
            ("Reimb_PreRtB", "Buyer reimburses pre-RtB development at the exit", 1, "ZHLP v3 Control D118"),
            ("Reimb_EW", "Buyer reimburses early works at the exit", 1, "ZHLP v3; " + EW_NOTE),
            ("Reimb_Dep", "Buyer reimburses cash deposits outstanding", 1, "ZHLP v3 Dyn_Inp rows 12-13")]:
        a(nm, label, "flag", v, src)

    i.sub("Tax (Hungary)")
    a("CIT_Rate", "Corporate income tax", "%", 0.09, "ZHLP v3 Control D66")
    a("Robin_Hood", "Robin Hood tax applies (1 = yes)", "flag", 0, "Tax adviser; " + PH)
    a("RH_Rate", "Robin Hood tax on the positive tax base", "%", 0.31, "Thtv. (energy suppliers' income tax)")
    a("Dep_Years", "Tax depreciation: PV, grid, development", "years", 30, "ZHLP v3 Control D65")
    a("Dep_BESS_Years", "Tax depreciation: BESS and augmentation", "years", 10, PH)
    a("Loss_Cap", "Tax losses usable, share of the tax base", "%", 0.5, "ZHLP v3 Control D102")
    a("Loss_Years", "Tax losses lapse NTP + years (one date)", "years", 5, "Tao. 17. §; simplification")
    a("Gain_Taxed", "Gain on the sale taxed (share deal: exempt)", "flag", 0, "Participation exemption; " + PH)

    i.sub("Valuation")
    a("Discount_Rate", "FEP discount rate (NPV)", "%", 0.08, "ZHLP v3 Control D9")
    a("Sale_Cost_RtB", "Sale costs at RtB, share of the consideration", "%", 0, PH)
    a("Sale_Cost_COD", "Sale costs at COD, share of the price", "%", 0, PH)
    i.note("Buyer target IRR: Dashboard sensitivity case (Case_Buyer_IRR); debt terms: Dashboard levers.")
    i.end()
    return i


# =========================================================================================== Revenue
def revenue(book):
    rv = book.sheet("Revenue", "Zalahaláp – revenue: Aurora and prices by case", timeline=True, widths={"E": 26, "F": 26, "G": 26, "H": 30})
    au = aurora()
    rv.bar("Aurora Q2-2026 Hungary, Central, real 2025 € (data/aurora_curves.csv)", e=["Block"], h="Source")
    series = [  # (key, label, unit, Aurora series, block code or None, sign, format)
        ("base", "Baseload price (information)", "EUR/MWh", "baseload", None, 1, "money1"),
        ("pvd4", "PV direct-export price, DC 4 h", "EUR/MWh", "direct_price", "DC 4h", 1, "money1"),
        ("pvd2", "PV direct-export price, DC 2 h (PV-only capture proxy)", "EUR/MWh", "direct_price", "DC 2h", 1,
         "money1"),
        ("pvt4", "PV transfer price (energy stored), DC 4 h", "EUR/MWh", "transfer_price", "DC 4h", 1, "money1"),
        ("pvt2", "PV transfer price (energy stored), DC 2 h", "EUR/MWh", "transfer_price", "DC 2h", 1, "money1"),
        ("st4", "BESS storage total (uplift net of the PV energy), DC 4 h", "EUR k/MW/yr", "storage_total", "DC 4h", 1,
         "money1"),
        ("st2", "BESS storage total, DC 2 h", "EUR k/MW/yr", "storage_total", "DC 2h", 1, "money1"),
        ("pvo4", "PV-only factor (no grid charging), DC 4 h", "factor", "pv_only_factor", "DC 4h", 1, "factor"),
        ("pvo2", "PV-only factor, DC 2 h", "factor", "pv_only_factor", "DC 2h", 1, "factor"),
        ("dcs4", "aFRR-down + FCR share of the storage total, DC 4 h", "factor", "down_cap_share", "DC 4h", 1,
         "factor"),
        ("dcs2", "aFRR-down + FCR share, DC 2 h", "factor", "down_cap_share", "DC 2h", 1, "factor"),
        ("chg4", "PV energy charged per MW of BESS, DC 4 h (−s+13)", "MWh/MW", "charge_res", "DC 4h", -1, "money"),
        ("chg2", "PV energy charged per MW of BESS, DC 2 h", "MWh/MW", "charge_res", "DC 2h", -1, "money"),
        ("dis4", "BESS discharge per MW, DC 4 h (s+11)", "MWh/MW", "discharge", "DC 4h", 1, "money"),
        ("dis2", "BESS discharge per MW, DC 2 h", "MWh/MW", "discharge", "DC 2h", 1, "money"),
    ]
    for key, label, unit, s, code, sign, fmt in series:
        blk = BLOCKS[code] if code else "Central"
        vals, src = au[(s, blk)]
        vals = _years([sign * v for v in vals])
        cells = {L(10 + k): v for k, v in enumerate(vals) if v is not None}
        r = rv.row(key, label, unit, e=code, inputs=tuple(cells) + (("E",) if code else ()), extra=cells,
                   fmt=fmt, source=_src(src), fmts={"E": "general"})
    vals = _years(au[("euro_index", "Central")][0], grow=0.02)
    cells = {L(10 + k): v for k, v in enumerate(vals)}
    rv.row("idx", "Aurora euro index (2025 = 1; +2 % a year after 2060)", "factor", inputs=tuple(cells),
           extra=cells, fmt="0.0000", source="Aurora Common Assumptions row 14")
    book.idx = rv.row_of("idx")
    rep = {}
    for code, blk in BLOCKS.items():
        rc = au[("repower_count", blk)][0]
        rep[code] = next(2026 + k for k, x in enumerate(rc) if x >= 1)
    rv.row("rep", "BESS repower year (E: DC 4 h, F: DC 2 h)", "year", e=rep["DC 4h"],
           f=rep["DC 2h"], inputs=("E", "F"), source="Aurora Degraded batteries s+1 (rows 140, 97)")
    book.define("Aurora_Blocks", f"Revenue!$E${rv.row_of('pvd4')}:$E${rv.row_of('pvd2')}")
    book.define("Repower_Year", f"Revenue!$E${rv.row_of('rep')}:$F${rv.row_of('rep')}")
    rv.note("Aurora's BESS stack is a fixed dispatch per block: it does not re-optimise when the fixed prices change. "
            "Rows 2061–2064 repeat 2060 (flat real).")
    rv.note("PV only: no standalone-solar capture in the databook; the DC 2 h block's direct price is the PV-only "
            "capture proxy (flagged). It is below v3's 65 % × baseload from 2031 (56 % on average 2030–60); the HU "
            "solar capture series is requested from Aurora.")

    rv.blank()
    rv.bar("Resulting prices by case, nominal – fixed price in the term, then the Aurora merchant tail",
           e=_case_links(book, rv))
    rv.row("yr", "Year", "year", t="={c}$4", fmt="int", bold=True)
    rv.sub("Operating months (live case): counted to each year end, then per year")
    rv.row("msince", "Months since COD at year end", "months",
           t="=({year}+1-YEAR(COD_Live))*Months_In_Year-MONTH(COD_Live)+1", fmt="int")
    rv.row("ocum", "Operating months to year end (E: life in months)", "months", e="=Life*Months_In_Year",
           t="=MIN({self.E},MAX(0,{msince}))", fmt="int")
    rv.row("opm", "Operating months in the year", "months", t="={ocum}-{ocum.prev}", fmt="int")
    rv.row("bfac", "BESS merchant factor: (1 − haircut) × (1 + case)", "factor",
           e="=(1-BESS_Haircut)*(1+Case_BESS)")
    idx = f"{{c}}${book.idx}"
    pair = lambda k: f"{{c}}${rv.row_of(k + '4')}:{{c}}${rv.row_of(k + '4') + 1}"      # noqa: E731 (4 h, 2 h rows)
    dl = book.lever_rows
    rows = []
    prev = None
    for k in range(1, 4):
        col = "EFG"[k - 1]
        dash = lambda key: f"Dashboard!${col}${dl[key]}"     # noqa: E731 (this case's lever)
        r = rv.sub("", e=["=1" if prev is None else f"=$E${prev}+1"], key=f"s{k}")
        rv.ws.cell(r, 3, f"=Dashboard!${col}${dl['Name']}")
        rv.ws.cell(r, 5).number_format = '"Case "0'
        rv.scope = f"c{k}_"
        rv.row("pos", "Merchant PV price, direct (E: block, F: BESS factor)", "EUR/MWh",
               e=f"=MATCH({dash('block')},Aurora_Blocks,0)", f=f"={dash('BESS_On')}*$E${rv.row_of('bfac')}",
               t=f"=INDEX({pair('pvd')},{{self.E}})*{idx}*(1+Case_PV_Price)", fmts={"E": "int", "F": "factor"})
        rv.row("pvs", "Merchant PV price, stored energy (transfer price)", "EUR/MWh",
               t=f"=INDEX({pair('pvt')},{{pos.E}})*{idx}*(1+Case_PV_Price)")
        rv.row("chg", "PV energy charged per MW of BESS (Aurora block)", "MWh/MW",
               t=f"=INDEX({pair('chg')},{{pos.E}})", fmt="money")
        rv.row("dmw", "BESS discharge per MW (Aurora block)", "MWh/MW", t=f"=INDEX({pair('dis')},{{pos.E}})",
               fmt="money")
        rv.row("fcum", "Fixed-term months to year end (E: term in months)", "months",
               e=f"={dash('Fix_Years')}*Months_In_Year", t="=MIN({self.E},{ocum})", fmt="int")
        rv.row("fm", "Fixed-term months in the year", "months", t="={fcum}-{fcum.prev}", fmt="int")
        rv.row("fts", "Fixed-term share of the operating months", "factor", t="=IF({opm}=0,0,{fm}/{opm})")
        rv.row("merch", "Merchant share of the operating months", "factor", t="=1-{fts}")
        rv.row("pxd", "PV price received, direct export", "EUR/MWh", t=f"={{fts}}*{dash('PV_Price')}+{{merch}}*{{pos}}",
               bold=True)
        rv.row("pxs", "PV price received, stored energy", "EUR/MWh", t=f"={{fts}}*{dash('PV_Price')}+{{merch}}*{{pvs}}",
               bold=True)
        rv.row("st", "BESS storage total, Aurora block (real)", "EUR k/MW", t=f"=INDEX({pair('st')},{{pos.E}})")
        rv.row("fac", "PV-only factor incl. the aFRR-down share", "factor",
               t=f"=INDEX({pair('pvo')},{{pos.E}})+INDEX({pair('dcs')},{{pos.E}})*AFRR_Share")
        rv.row("upl", "BESS merchant revenue, nominal (× F above)", "EUR k/MW", t=f"={{st}}*{{fac}}*{idx}*{{pos.F}}",
               bold=True)
        rows.append({k2[len(rv.scope):]: v for k2, v in rv.rows.items() if k2.startswith(rv.scope)})
        rv.scope = ""
        prev = r
    rv.end()
    book.rev = rows
    return rv


def _src(src):
    s = src.replace("Degraded batteries", "Aurora DB").replace("Common Assumptions", "Aurora CA")
    return s if len(s) <= 60 else s[:57] + "…"


# =========================================================================================== Calc
def calc(book, R):
    c = book.sheet("Calc", "Zalahaláp – calculations, one annual block per case",
                   timeline=True)
    last = book.last_col
    c.bar("Timeline and discounting (all cases)", e=["All cases"], h="Total")
    c.row("codcol", "COD year: column on the timeline (E), its year end (F)", "#",
          e=f"=MATCH(YEAR(COD_Live),$J$4:${last}$4,0)", f=f"=INDEX($J$5:${last}$5,{{self.E}})",
          fmts={"E": "int", "F": "date"})
    c.row("dfx", "Discount factor to the exit at the buyer target IRR", "factor",
          t="=(1+Case_Buyer_IRR)^((Exit_Date-{yend})/Days_In_Year)")
    c.row("dfc", "Discount factor to the COD year end (buyer target)", "factor",
          t="=(1+Case_Buyer_IRR)^(({codcol.F}-{yend})/Days_In_Year)")
    c.row("bxd", "Buyer at RtB: flow dates (year end, not before the exit)", "date", t="=MAX({yend},Exit_Date)")
    c.row("bcd", "Buyer at COD: flow dates (not before the COD year end)", "date", t="=MAX({yend},{codcol.F})")
    c.row("pye", "Previous year end (0 before the first column)", "date", t="={p}$5")
    c.row("dfa", "Discount factor, first year end to acquisition (NPV)", "factor",
          e="=1/(1+Discount_Rate)^(($J$5-Acq_Date)/Days_In_Year)")
    rev = book.rev
    idx = f"Revenue!{{c}}${book.idx}"

    dl = book.lever_rows
    a, z = R.get("l_spv", 1), R.get("l_fee_last", 1)       # the dated outlays on the Dashboard (first pass)

    def rows(i, blk):
        rr = rev[i - 1]
        col = "EFG"[i - 1]
        ref = lambda k: f"Revenue!{{c}}${rr[k]}"              # noqa: E731 (this case's Revenue row)
        lever = lambda k: f"Dashboard!${col}${dl[k]}"         # noqa: E731 (this case's Dashboard lever)
        dash = lambda k: f"Dashboard!${col}${R.get(k, 1)}"    # noqa: E731 (this case's Dashboard result)
        c.sub("Timing")
        c.row("yrs", "Years from the COD year (E: first full operating year)", "years",
              e="=YEAR(COD_Live)+1", t="={year}-YEAR(COD_Live)", fmts={"E": "int"})
        for key, label, t in [("f_pre", "Before the COD year (1/0)", "=N({yrs}<0)"),
                              ("f_from", "From the COD year (1/0)", "=N({yrs}>=0)"),
                              ("f_upto", "Up to the COD year (1/0)", "=N({yrs}<=0)"),
                              ("codf", "COD year (1/0)", "=N({yrs}=0)"),
                              ("f_capex", "Capex year (1/0)", "=N({year}=Capex_Year)"),
                              ("f_ntp", "NTP year (1/0)", "=N({year}=YEAR(NTP_Date))"),
                              ("f_exit", "Exit year (1/0)", "=N({year}=YEAR(Exit_Date))"),
                              ("f_first", "First column, dated at the exit or COD (1/0)", "=N({year}=First_Year)")]:
            c.row(key, label, "flag", t=t)
        c.row("msince", "Months since COD at year end", "months",
              t="=({year}+1-YEAR(COD_Live))*Months_In_Year-MONTH(COD_Live)+1")
        c.row("ocum", "Operating months to year end (E: life in months)", "months", e="=Life*Months_In_Year",
              t="=MIN({self.E},MAX(0,{msince}))")
        c.row("omonths", "Operating months in the year", "months", t="={ocum}-{ocum.prev}")
        c.row("oshare", "Operating share of the year", "%", h="=SUM({self.all})", t="={omonths}/Months_In_Year",
              fmts={"H": "money2"})
        c.row("deg", "Degradation factor (from the first full year)", "factor", t="=(1-Degradation)^MAX({yrs}-1,0)")
        c.row("retired", "BESS retired, no repower (E: repower year)", "flag",
              e=f"=INDEX(Repower_Year,1,Revenue!$E${rr['pos']})", t="=(1-Repower)*N({year}>={self.E})",
              fmts={"E": "int"})
        c.row("live", "BESS in service (1/0)", "flag", t="=1-{retired}")
        c.row("rep", "Repower year: augmentation (1/0)", "flag", t="=Repower*N({year}={retired.E})")
        c.sub("Energy")
        c.row("cap", "PV MWp (E); net yield after case and clipping (F)", "MWp", e=f"={lever('MWp')}",
              f=f"={lever('Yield')}*(1+Case_Yield)*(1-{lever('Clipping')})", fmts={"F": "money"})
        c.row("bess", "BESS power, MW (E); energy, MWh (F)", "MW", e=f"={lever('BESS_MW')}*{lever('BESS_On')}",
              f=f"={{self.E}}*{lever('Hours')}", fmts={"F": "money"})
        ffy = "=SUMIF({year.all},{yrs.E},{self.all})"
        tot = "=SUM({self.all})"
        c.row("gen", "PV generation (E: first full year)", "MWh", e=ffy, h=tot,
              t="={cap.E}*{cap.F}*PV_Avail*{deg}*{oshare}")
        c.row("chcap", "BESS charging capacity from PV", "MWh", t=f"={ref('chg')}*{{bess.E}}*{{oshare}}*{{live}}")
        c.row("sent", "PV energy sent to the BESS", "MWh", e=ffy, h=tot, t="=MIN({gen},{chcap})")
        c.row("direct", "PV energy exported directly", "MWh", e=ffy, h=tot, t="={gen}-{sent}")
        c.row("dcap", "BESS discharge capacity", "MWh", t=f"={ref('dmw')}*{{bess.E}}*{{oshare}}*{{live}}")
        c.row("dis", "BESS discharge (≤ PV charged × RTE, no grid charging)", "MWh", e=ffy, h=tot,
              t="=MIN({dcap},{sent}*BESS_RTE)")
        c.sub("Revenue")
        c.row("rpvd", "PV revenue, direct export", "EUR k", e=ffy, h=tot, t=f"={{direct}}*{ref('pxd')}/Thousand")
        c.row("rpvs", "PV revenue, energy sent to the BESS", "EUR k", e=ffy, h=tot, t=f"={{sent}}*{ref('pxs')}/Thousand")
        c.row("rgoo", "Guarantees of origin", "EUR k", e=ffy, h=tot, t=f"={{gen}}*GoO_Price*{idx}/Thousand")
        c.row("bpx", "BESS fixed price (E), cycles per day (F), override (G)", "EUR/MWh",
              e=f"={lever('BESS_Price')}", f=f"={lever('Fix_Cycles')}", g=f"={lever('BESS_Income')}",
              fmts={"E": "money1", "F": "money1", "G": "money"})
        c.row("bfull", "BESS fixed income, full year (E: override if set)", "EUR k",
              e="=IF({bpx.G}>0,{bpx.G},{bpx.E}*{bess.F}*{bpx.F}*Days_In_Year/Thousand)")
        c.row("bfy", "BESS income at the fixed price (E: first full year)", "EUR k", e=ffy,
              t="={bfull.E}*{oshare}*{live}")
        c.row("rbfix", "BESS revenue at the fixed price (fixed term)", "EUR k", e=ffy, h=tot,
              t=f"={{bfy}}*{ref('fts')}")
        c.row("rbm", "BESS revenue, merchant tail (after the fixed term)", "EUR k", e=ffy, h=tot,
              t=f"={{bess.E}}*{ref('upl')}*{{oshare}}*{{live}}*{ref('merch')}")
        c.row("rbess", "BESS revenue", "EUR k", e=ffy, h=tot, t="={rbfix}+{rbm}")
        c.row("rev", "Revenue", "EUR k", e=ffy, h=tot, t="={rpvd}+{rpvs}+{rgoo}+{rbess}", bold=True, top=True)
        c.sub("Opex (real 2025 × the Aurora euro index)")
        c.row("om", "O&M, all-in", "EUR k", e=ffy, h=tot, t=f"=-OM_Cost*{{cap.E}}*{idx}*{{oshare}}")
        c.row("bal", "Balancing (merchant share of the year)", "EUR k", e=ffy, h=tot,
              t=f"=-{{gen}}*Balancing*{idx}*{ref('merch')}/Thousand")
        c.row("lease", "Land lease", "EUR k", e=ffy, h=tot, t=f"=-Lease*{idx}*{{oshare}}")
        c.row("bopex", "BESS fixed opex (per MW and per MWh)", "EUR k", e=ffy, h=tot,
              t=f"=-(BESS_Opex_MW*{{bess.E}}+BESS_Opex_MWh*{{bess.F}})*{idx}*{{oshare}}*{{live}}")
        c.row("levy", "HIPA and innovation contribution (on revenue)", "EUR k", e=ffy, h=tot,
              t="=-{rev}*(HIPA+Innovation)")
        c.row("opex", "Opex", "EUR k", e=ffy, h=tot, t="=SUM({om}:{levy})", bold=True, top=True)
        c.row("ebitda", "EBITDA", "EUR k", e=ffy, h=tot, t="={rev}+{opex}", bold=True)
        c.sub("Capex")
        c.row("capex", "Construction capex (E: total uses, F: of which at NTP)", "EUR k", e=f"={dash('k_tot')}",
              f=f"={dash('k_ew')}", h=tot, t="=-({self.E}-{self.F})*{f_capex}")
        c.row("ewntp", "Early works at NTP (E48/E49)", "EUR k", h=tot, t="=-{capex.F}*{f_ntp}")
        c.row("aug", "BESS augmentation at the repower year (indexed)", "EUR k", h=tot,
              t=f"=-{{bess.F}}*BESS_Capex*{idx}*{{rep}}")
        c.row("capt", "Capex", "EUR k", h=tot, t="=SUM({capex}:{aug})", bold=True, top=True)
        c.sub("Tax (SPV, unlevered): straight-line from COD on operating months")
        dep = "=-{self.E}*(MIN({self.F},{ocum})-MIN({self.F},{ocum}-{omonths}))/{self.F}"   # no blank-cell MIN
        c.row("db", "Depreciation: BESS (E: base, F: months)", "EUR k", e=f"={dash('k_bess')}",
              f="=Dep_BESS_Years*Months_In_Year", h=tot, t=dep, fmts={"F": "int"})
        c.row("dpv", "Depreciation: PV, grid, development (E: base, F: months)", "EUR k",
              e=f"={{capex.E}}-{{db.E}}-{dash('k_land')}+EW_Total+PreRtB_Spent", f="=Dep_Years*Months_In_Year",
              h=tot, t=dep, fmts={"F": "int"})
        c.row("fda", "Augmentation depreciation years (1/0)", "flag",
              t="=N({year}>={retired.E})*N({year}<{retired.E}+Dep_BESS_Years)")
        c.row("da", "Depreciation: augmentation", "EUR k", h=tot, t="={aug.H}/Dep_BESS_Years*{fda}")
        c.row("taxb", "Tax base before losses", "EUR k", h=tot, t="={ebitda}+{db}+{dpv}+{da}")

        def losses(p, base, lapse_e=None):
            c.row(p + "lo", "Tax losses: opening (previous closing)", "EUR k", t=f"={{p}}{c.r + 4}")
            c.row(p + "ladd", "Losses added", "EUR k", h=tot, t=f"=MAX(0,-{{{base}}})")
            c.row(p + "luse", "Losses used (capped)", "EUR k", h=tot,
                  t=f"=-MIN({{{p}lo}},Loss_Cap*MAX(0,{{{base}}}))")
            c.row(p + "lapse", "Losses lapsed (E: year, NTP + carry-forward)", "EUR k", h=tot,
                  e=lapse_e or "={lapse.E}", t=f"=-SUM({{{p}lo}}:{{{p}luse}})*N({{year}}={{self.E}})",
                  fmts={"E": "int"})
            c.row(p + "lc", "Tax losses: closing", "EUR k", t=f"=SUM({{{p}lo}}:{{{p}lapse}})")
            c.row(p + "taxable", "Taxable profit after losses", "EUR k", h=tot, t=f"=MAX(0,{{{base}}})+{{{p}luse}}")
            c.row(p + "cit", "Corporate income tax", "EUR k", h=tot, t=f"=-CIT_Rate*{{{p}taxable}}")
            c.row(p + "rh", "Robin Hood tax (when switched on)", "EUR k", h=tot,
                  t=f"=-Robin_Hood*RH_Rate*MAX(0,{{{base}}})")
        losses("", "taxb", "=YEAR(NTP_Date)+Loss_Years")
        c.row("tax", "Tax", "EUR k", h=tot, bold=True, t="={cit}+{rh}")
        c.sub("Debt (levered case; no fees, no DSRA, no sweep)")
        c.row("fin", "Gearing (E), debt rate (F), tenor in years (G)", "%", e=f"={lever('Gearing')}",
              f=f"={lever('Debt_Rate')}", g=f"={lever('Tenor')}", fmts={"G": "int"})
        c.row("draw", "Debt drawn, pro rata with the construction capex", "EUR k", h=tot,
              t="=-{fin.E}*({capex}+{ewntp})")
        c.row("cbo", "Construction debt: opening (draws and IDC to date)", "EUR k", t=f"={{p}}{c.r + 2}")
        c.row("idc", "Interest during construction, capitalised (mid-year)", "EUR k", h=tot,
              t="=({cbo}+{draw}*Draw_Timing)*{fin.F}*{f_pre}")
        c.row("cbc", "Construction debt: closing (E: debt at COD)", "EUR k", e="=MAX({self.all})",
              t="={cbo}+{draw}+{idc}")
        c.row("dwin", "Repayment window: COD year + tenor (1/0)", "flag", t="={f_from}*N({yrs}<{fin.G})")
        c.row("dopen", "Debt: opening balance (previous closing)", "EUR k", t=f"={{p}}{c.r + 4}")
        c.row("dint", "Interest paid (deductible)", "EUR k", h=tot, t="={dopen}*{fin.F}*{dwin}")
        c.row("dann", "Annuity due (E: level annuity on the debt at COD)", "EUR k", h=tot,
              e="=PMT({fin.F},{fin.G},-{cbc.E})", t="={dwin}*{self.E}")
        c.row("dprin", "Principal repaid", "EUR k", h=tot, t="=MIN({dopen},{dann}-{dint})")
        c.row("dclose", "Debt: closing balance", "EUR k", t="={dopen}+{draw}+{idc}-{dprin}", bold=True)
        c.row("dserv", "Debt service", "EUR k", h=tot, t="={dint}+{dprin}", bold=True)
        c.sub("Tax, levered case (interest deductible; own loss carry-forward)")
        c.row("taxbl", "Tax base before losses, levered", "EUR k", h=tot, t="={taxb}-{dint}")
        losses("l_", "taxbl")
        c.row("taxl", "Tax, levered", "EUR k", h=tot, bold=True, t="={l_cit}+{l_rh}")
        c.row("cfads", "CFADS: EBITDA, levered tax, augmentation (from COD)", "EUR k", h=tot,
              t="=({ebitda}+{taxl}+{aug})*{f_from}")
        c.row("dscr", "DSCR (E: minimum, F: average; information)", "x", e='=IF(COUNT({self.all})=0,0,MIN({self.all}))',
              f='=IF(COUNT({self.all})=0,0,AVERAGE({self.all}))', t='=IF({dserv}>0,{cfads}/{dserv},"")')
        c.sub("Cash flows")
        c.row("dated", "Dated outlays by year: development, deposits, fees", "EUR k", h=tot,
              t=f'=SUMIFS(Dashboard!${col}${a}:${col}${z},Dated_Dates,">"&{{pye}},Dated_Dates,"<="&{{yend}})')
        c.row("codnet", "Net cash at COD: deposits back less fees", "EUR k", h=tot, t="={codf}*Hold_At_COD")
        c.row("pre", "Pre-COD outlays", "EUR k", h=tot, t="={dated}+{codnet}")
        c.row("pcf", "Project cash flow before tax", "EUR k", h=tot, t="={ebitda}+{capt}+{pre}", bold=True)
        c.row("pcft", "Project cash flow after tax", "EUR k", h=tot, t="={pcf}+{tax}", bold=True)
        c.row("lev", "Levered project cash flow after tax", "EUR k", h=tot, t="={pcf}+{taxl}+{draw}-{dserv}",
              bold=True)
        c.row("hold", "FEP hold: cash flow after tax (E: PSH buy-out)", "EUR k", h=tot,
              e=f"={dash('jv_pl')}", t="={pcft}-{self.E}*{f_exit}", bold=True)
        c.row("holdp", "FEP hold: cash flow before tax", "EUR k", h=tot, t="={hold}-{tax}")
        c.row("holdl", "FEP hold, levered: cash flow after tax", "EUR k", h=tot, bold=True,
              t="={lev}-{hold.E}*{f_exit}")
        c.row("holdlp", "FEP hold, levered: cash flow before tax", "EUR k", h=tot, t="={holdl}-{taxl}")
        c.row("cupto", "Capex up to the COD year", "EUR k", h=tot, t="={capt}*{f_upto}")
        c.row("cproc", "FEP proceeds at COD, JV split (E: total)", "EUR k", h=tot, e=f"={dash('c_fep')}",
              t="={self.E}*{codf}")
        c.row("codr", "FEP COD sale: flow before FEP tax (E: contributions)", "EUR k", h=tot,
              e="=-{pre.H}-{cupto.H}", t="={cupto}+{pre}+{cproc}")
        c.row("ctax", "FEP tax at COD (E: total)", "EUR k", h=tot, e=f"={dash('c_tax')}", t="={self.E}*{codf}")
        c.row("codt", "FEP COD sale: cash flow after FEP tax", "EUR k", h=tot, t="={codr}+{ctax}", bold=True)
        c.row("cprocl", "FEP proceeds at COD, levered split (E: total)", "EUR k", h=tot, e=f"={dash('cl_fep')}",
              t="={self.E}*{codf}")
        c.row("codrl", "FEP COD sale, levered: before FEP tax (E: equity)", "EUR k", h=tot,
              e="={codr.E}-{draw.H}", t="={cupto}+{pre}+{draw}+{cprocl}")
        c.row("ctaxl", "FEP tax at COD, levered (E: total)", "EUR k", h=tot, e=f"={dash('cl_tax')}",
              t="={self.E}*{codf}")
        c.row("codtl", "FEP COD sale, levered: after FEP tax", "EUR k", h=tot, t="={codrl}+{ctaxl}", bold=True)
        c.row("bx", "Buyer at RtB: flows after the exit (E: value at target)", "EUR k",
              e="=SUMPRODUCT({self.all},{dfx.all})", h=tot, t="={pcft}-{dated}")
        c.row("price", "Buyer price at the exit, in the first column (E: all-in)", "EUR k",
              e=f"={dash('allin')}*{{cap.E}}", t="={self.E}*{f_first}")
        c.row("bxp", "Buyer at RtB at our price, after tax", "EUR k", t="={bx}-{price}")
        c.row("bxpp", "Buyer at RtB at our price, before tax", "EUR k", t="={bxp}-{tax}")
        c.row("bxl", "Buyer at RtB, levered (raises the debt)", "EUR k", t="={lev}-{dated}-{price}")
        c.row("bc", "Buyer at COD: flows from COD year (E: value at target)", "EUR k",
              e="=SUMPRODUCT({self.all},{dfc.all})", h=tot, t="=({ebitda}+{tax}+{aug})*{f_from}")
        c.row("cprice", "Buyer price at COD, in the first column (E: EV)", "EUR k", e=f"={dash('ev')}",
              t="={self.E}*{f_first}")
        c.row("bcp", "Buyer at COD at our price, after tax", "EUR k", t="={bc}-{cprice}")
        c.row("bcpp", "Buyer at COD at our price, before tax", "EUR k", t="={bcp}-{tax}*{f_from}")
        c.row("cequity", "Buyer equity at COD, first column (E: EV − debt)", "EUR k", e="={cprice.E}-{cbc.E}",
              t="={self.E}*{f_first}")
        c.row("bcl", "Buyer at COD, levered (assumes the debt)", "EUR k",
              t="=({ebitda}+{taxl}+{aug})*{f_from}-{dserv}-{cequity}")
        c.sub("Returns (E: after tax; F: before tax; G: MOIC; H: XIRR start guess)")
        dates = "$J$5:$" + last + "$5"
        gx = "=MIN(XIRR_Guess,{self.G}-1)"                   # Excel's XIRR returns ~0 when it fails: start guess
        gb = "=MIN(Case_Buyer_IRR,{self.G}-1)"               # buyer at a fixed RtB premium: may sit far below target
        x = lambda k, d_=dates: f'=IFERROR(XIRR({{{k}.all}},{d_},{{self.H}}),"n/a")'   # noqa: E731
        mo = lambda k: f'=SUMIF({{{k}.all}},">0")/-SUMIF({{{k}.all}},"<0")'   # noqa: E731
        pct = {"G": "mult", "H": "pct"}
        c.row("r_proj", "Project IRR, unlevered", "%", e=x("pcft"), f=x("pcf"), g=mo("pcft"), h=gx, bold=True,
              fmts=pct)
        c.row("r_hold", "FEP IRR, hold", "%", e=x("hold"), f=x("holdp"), g=mo("hold"), h=gx, bold=True, fmts=pct)
        v = lambda k: f"=XNPV(Discount_Rate,{{{k}.all}},{dates})*$E${c.row_of('dfa')}"   # noqa: E731
        c.row("v_hold", "FEP NPV at the discount rate, hold (at acquisition)", "EUR k", e=v("hold"),
              f=v("holdp"))
        c.row("r_cod", "FEP IRR, sale at COD", "%", e=x("codt"), f=x("codr"), g=mo("codt"), h=gx, bold=True,
              fmts=pct)
        c.row("v_cod", "FEP NPV at the discount rate, COD sale (at acquisition)", "EUR k", e=v("codt"),
              f=v("codr"))
        bxd = f"$J${c.row_of('bxd')}:${last}${c.row_of('bxd')}"
        bcd = f"$J${c.row_of('bcd')}:${last}${c.row_of('bcd')}"
        c.row("r_bx", "Buyer IRR at RtB at our all-in price", "%", e=x("bxp", bxd), f=x("bxpp", bxd), g=mo("bxp"),
              h=gb, fmts=pct)
        c.row("r_bc", "Buyer IRR at COD at our price", "%", e=x("bcp", bcd), f=x("bcpp", bcd), g=mo("bcp"),
              h="=Case_Buyer_IRR", fmts=pct)
        c.row("r_holdl", "FEP IRR, hold, levered", "%", e=x("holdl"), f=x("holdlp"), g=mo("holdl"), h=gx, bold=True,
              fmts=pct)
        c.row("r_codl", "FEP IRR, sale at COD, levered", "%", e=x("codtl"), f=x("codrl"), g=mo("codtl"), h=gx,
              fmts=pct)
        c.row("r_bl", "Buyer IRR at RtB, levered, after tax", "%", e=x("bxl", bxd), g=mo("bxl"), h=gb, fmts=pct)
        c.row("r_bcl", "Buyer IRR at COD, levered (n/a: price below debt)", "%", e=x("bcl", bcd),
              h="=Case_Buyer_IRR", fmts={"H": "pct"})

    blocks = Block(c, 3).build(lambda i: f"Case {i}: annual calculation (EUR k)", rows)
    for i, blk in enumerate(blocks, 1):
        c.ws.cell(blk["bar"], 3, f'=Dashboard!${"EFG"[i - 1]}${book.lever_rows["Name"]}&": annual calculation (EUR k)"')
    c.end()
    book.blocks = blocks
    return c




# =========================================================================================== Dashboard body
def dashboard(book, d, R):
    """Sections 4–13 of SPEC §3.1. `R` = Dashboard row numbers from the first pass (rows referenced before they
    are written: the dated block and the JV splits sit below the results that read them)."""
    B = book.blocks
    cl = _case_links(book, d)

    def pc(label, unit, fn, key=None, **kw):
        """A row with one link per case: fn(block dict) -> formula."""
        e, f, g = (fn(b) for b in B)
        return d.row(key, label, unit, e=e, f=f, g=g, **kw)

    def r(key):
        return R.get(key, 1)

    LR = book.lever_rows
    for i, b in enumerate(B):
        C = "EFG"[i]
        d.ws[f"{C}{LR['o_dis']}"] = f"=Calc!$E${b['dis']}/Thousand"
        d.ws[f"{C}{LR['o_cyc']}"] = f"=IF({C}{LR['MWh']}=0,0,{C}{LR['o_dis']}*Thousand/{C}{LR['MWh']}/Days_In_Year)"
        d.ws[f"{C}{LR['o_inc']}"] = f"=Calc!$E${b['bfy']}"

    d.blank()
    d.bar("Capacity and energy, first full operating year", e=cl)
    pc("First full operating year", "year", lambda b: f"=Calc!$E${b['yrs']}")
    pc("PV generation", "GWh", lambda b: f"=Calc!$E${b['gen']}/Thousand")
    pc("PV energy sent to the BESS", "GWh", lambda b: f"=Calc!$E${b['sent']}/Thousand")
    pc("PV energy exported directly", "GWh", lambda b: f"=Calc!$E${b['direct']}/Thousand")
    pc("PV price received, all output (blended)", "EUR/MWh",
       lambda b: f"=(Calc!$E${b['rpvd']}+Calc!$E${b['rpvs']})*Thousand/Calc!$E${b['gen']}")
    pc("BESS revenue per MW", "EUR k/MW", lambda b: f"=IF(Calc!$E${b['bess']}>0,Calc!$E${b['rbess']}/Calc!$E${b['bess']},0)")

    d.blank()
    d.bar("Revenue and EBITDA (EUR k)", e=cl)
    for col, title in (("E", "First full operating year"), ("H", "Lifetime")):
        d.sub(title)
        pc("PV revenue (fixed price, then merchant)", "EUR k",
           lambda b: f"=Calc!${col}${b['rpvd']}+Calc!${col}${b['rpvs']}")
        pc("BESS revenue", "EUR k", lambda b: f"=Calc!${col}${b['rbess']}")
        pc("of which at the fixed price", "EUR k", lambda b: f"=Calc!${col}${b['rbfix']}", italic=True)
        pc("Guarantees of origin", "EUR k", lambda b: f"=Calc!${col}${b['rgoo']}")
        pc("Revenue", "EUR k", lambda b: f"=Calc!${col}${b['rev']}", key=f"rev_{col}", bold=True, top=True)
        pc("Opex", "EUR k", lambda b: f"=Calc!${col}${b['opex']}")
        pc("EBITDA", "EUR k", lambda b: f"=Calc!${col}${b['ebitda']}", key=f"ebitda_{col}", bold=True)
        d.row(None, "EBITDA margin", "%", cases=f"={{ebitda_{col}}}/{{rev_{col}}}")

    d.blank()
    d.bar("Cash to RtB by nature (EUR k; all cases)", e=["All cases"])
    sm = lambda a, b: f"=-E{r(a)}" if a == b else f"=-SUM(E{r(a)}:E{r(b)})"        # noqa: E731
    d.row("n_spv", "SPV acquisition price (FEP)", "EUR k", e=sm("l_spv", "l_spv"))
    d.row("n_trans", "Transaction costs (FEP)", "EUR k", e=sm("l_trans", "l_trans"))
    d.row("n_dsa", "DSA loan to PSH", "EUR k", e=sm("l_dsa1", "l_dsa_last"))
    d.row("n_ew", "Early works, four lines (to be revisited)", "EUR k", e=sm("l_ew1", "l_ew_last"))
    book.define("EW_Total", d.ref("n_ew", "E"))
    d.row("n_pre", "Pre-RtB development", "EUR k", e=sm("l_pre1", "l_pre_last"))
    book.define("PreRtB_Spent", d.ref("n_pre", "E"))
    d.row("n_dep", "Cash deposits, net of releases before the exit", "EUR k", e=sm("l_dep1", "l_dep_last"))
    d.row("n_fee", "Bank guarantee fees to the exit", "EUR k", e=sm("l_fee1", "l_fee_last"))
    d.row("n_tot", "Total cash to RtB", "EUR k", e="=SUM({n_spv}:{n_fee})", bold=True, top=True)
    d.row("n_reimb", "of which reimbursed (pre-RtB development, early works)", "EUR k",
          e="=Reimb_PreRtB*{n_pre}+Reimb_EW*{n_ew}", italic=True)
    d.row("n_ret", "of which returned at the exit (cash deposits)", "EUR k",
          e='=Reimb_Dep*SUMPRODUCT(Dep_HUF*(Dep_Rel=""))*Thousand/FX', italic=True)
    d.row("n_sunk", "of which sunk (recovered only through the consideration)", "EUR k",
          e="={n_tot}-{n_reimb}-{n_ret}", italic=True)
    d.row("n_buyer", "Buyer pays on top of the premium (reimbursed + returned)", "EUR k", e="={n_reimb}+{n_ret}")
    book.define("Buyer_Reimb", d.ref("n_buyer", "E"))
    d.row("n_exp", "FEP exposure at RtB (total less deposits returned)", "EUR k", e="={n_tot}-{n_ret}", bold=True)

    d.blank()
    d.bar("Capex to COD (EUR k, live sensitivity case)", e=cl)
    cx = "*(1+Case_Capex)"
    d.row("k_pv", "PV equipment and balance of plant", "EUR k", cases="=PV_Capex*{MWp}" + cx)
    d.row("k_cab", "MV cable", "EUR k", cases="=Cable_Capex" + cx)
    d.row("k_sub", "HV substation, bay and grid connection fee", "EUR k", cases="=(Substation+Grid_Fee)" + cx)
    d.row("k_land", "Land", "EUR k", cases="=Land_Capex" + cx)
    d.row("k_bess", "BESS (MWh × capex)", "EUR k", cases="=BESS_Capex*{MWh}" + cx)
    d.row("k_ew", "Early works at NTP (E48/E49; to be revisited)", "EUR k", cases="=EW_NTP" + cx)
    d.row("k_tot", "Total uses to COD (= equity, unlevered)", "EUR k", cases="=SUM({k_pv}:{k_ew})", bold=True,
          top=True)
    d.row(None, "Total uses per MWp", "EUR k/MWp", cases="={k_tot}/{MWp}")
    for nm, k in (("Capex_Land", "k_land"), ("Capex_BESS", "k_bess"), ("Capex_EW", "k_ew"), ("Capex_Total", "k_tot")):
        d.name_row(nm, k)

    d.blank()
    d.bar("Financing, levered case (EUR k; no fees, DSRA or sweep)", e=cl)
    pc("Debt drawn in the capex year (gearing × total uses)", "EUR k", lambda b: f"=Calc!$H${b['draw']}")
    pc("Interest during construction, capitalised", "EUR k", lambda b: f"=Calc!$H${b['idc']}")
    pc("Debt at COD", "EUR k", lambda b: f"=Calc!$E${b['cbc']}", key="f_debt", bold=True)
    pc("Equity to COD, levered (total uses less debt drawn)", "EUR k", lambda b: f"=Calc!$E${b['codrl']}")
    pc("Annual debt service (level annuity)", "EUR k", lambda b: f"=Calc!$E${b['dann']}")
    pc("Minimum DSCR (information)", "x", lambda b: f"=Calc!$E${b['dscr']}", key="f_dmin")
    pc("Average DSCR (information)", "x", lambda b: f"=Calc!$F${b['dscr']}")
    pc("Lifetime corporate income tax, levered", "EUR k", lambda b: f"=-Calc!$H${b['taxl']}")

    d.blank()
    d.bar("Returns", e=cl)
    d.sub("Sale at RtB (dated flows to the exit; XIRR)")
    d.row("p_rtb", "RtB price used (premium)", "EUR k/MWp", cases=f"=IF({{RtB_Price}}>0,{{RtB_Price}},{{c}}{r('imp')})", bold=True)
    pc("Buyer value at the exit (post-exit flows at the target)", "EUR k", lambda b: f"=Calc!$E${b['bx']}",
       key="bval")
    d.row("imp", "Implied premium at the buyer target IRR (floored at 0)", "EUR k/MWp",
          cases="=MAX(0,({bval}-Buyer_Reimb)/{MWp})")
    d.row("impneg", "Implied premium below zero (1 = floored)", "flag", cases="=N({bval}<Buyer_Reimb)")
    d.row("allin", "All-in price incl. reimbursements", "EUR k/MWp", cases="={p_rtb}+Buyer_Reimb/{MWp}")
    d.name_row("RtB_AllIn", "allin")
    pc("Buyer IRR at our price, unlevered after tax", "%", lambda b: f"=Calc!$E${b['r_bx']}")
    pc("Buyer IRR at our price, unlevered before tax", "%", lambda b: f"=Calc!$F${b['r_bx']}")
    pc("Buyer IRR at our price, levered after tax (raises debt)", "%", lambda b: f"=Calc!$E${b['r_bl']}")
    d.row(None, "FEP exposure (cash to RtB less deposits returned)", "EUR k",
          cases=f"=-{{c}}{r('dl_out')}-$E${r('n_ret')}")
    d.row(None, "FEP proceeds at the exit", "EUR k", cases=f"={{c}}{r('jv_fep')}")
    d.row(None, "FEP profit, pre-tax", "EUR k", cases=f"={{c}}{r('dl_pre')}")
    d.row(None, "FEP profit, post-tax", "EUR k", cases=f"={{c}}{r('dl_post')}", bold=True)
    d.row(None, "MOIC, pre-tax", "x", cases=f"={{c}}{r('dl_moic')}")
    d.row(None, "IRR, pre-tax", "%", cases=f"={{c}}{r('dl_irr_pre')}")
    d.row(None, "IRR, post-tax", "%", cases=f"={{c}}{r('dl_irr')}", bold=True)
    d.row(None, "NPV at the discount rate, post-tax (at acquisition)", "EUR k", cases=f"={{c}}{r('dl_npv')}")
    d.row(None, "PSH profit", "EUR k", cases=f"={{c}}{r('psh_profit')}")
    d.sub("Sale at COD (buyer pays at COD year end; receives the COD-year operating cash)")
    d.row("p_cod", "COD price used", "EUR k/MWp", cases=f"=IF({{COD_Price}}>0,{{COD_Price}},{{c}}{r('impc')})", bold=True)
    pc("Implied COD price at the buyer target IRR (floored at 0)", "EUR k/MWp",
       lambda b: f"=MAX(0,Calc!$E${b['bc']}/{{MWp}})", key="impc")
    d.row("ev", "Enterprise value used", "EUR k", cases="={p_cod}*{MWp}")
    d.name_row("COD_EV", "ev")
    pc("Buyer IRR at our price, unlevered after tax", "%", lambda b: f"=Calc!$E${b['r_bc']}")
    pc("Buyer IRR at our price, unlevered before tax", "%", lambda b: f"=Calc!$F${b['r_bc']}")
    pc("Buyer IRR at our price, levered after tax (takes debt)", "%", lambda b: f"=Calc!$E${b['r_bcl']}")
    pc("FEP contributions to COD", "EUR k", lambda b: f"=Calc!$E${b['codr']}")
    d.row(None, "FEP proceeds at COD", "EUR k", cases=f"={{c}}{r('c_fep')}")
    pc("FEP profit, pre-tax", "EUR k", lambda b: f"=Calc!$H${b['codr']}")
    pc("FEP profit, post-tax", "EUR k", lambda b: f"=Calc!$H${b['codt']}", bold=True)
    pc("MOIC, pre-tax", "x", lambda b: f'=SUMIF(Calc!$J${b["codr"]}:${book.last_col}${b["codr"]},">0")'
                                       f'/-SUMIF(Calc!$J${b["codr"]}:${book.last_col}${b["codr"]},"<0")')
    pc("IRR, pre-tax", "%", lambda b: f"=Calc!$F${b['r_cod']}")
    pc("IRR, post-tax", "%", lambda b: f"=Calc!$E${b['r_cod']}", bold=True)
    pc("NPV at the discount rate, post-tax (at acquisition)", "EUR k", lambda b: f"=Calc!$E${b['v_cod']}")
    d.row(None, "PSH profit", "EUR k", cases=f"={{c}}{r('c_pl')}+{{c}}{r('psh_net')}")
    pc("Levered: FEP profit, post-tax (debt repaid from price)", "EUR k", lambda b: f"=Calc!$H${b['codtl']}")
    pc("Levered: MOIC, pre-tax", "x", lambda b: f'=SUMIF(Calc!$J${b["codrl"]}:${book.last_col}${b["codrl"]},">0")'
                                              f'/-SUMIF(Calc!$J${b["codrl"]}:${book.last_col}${b["codrl"]},"<0")')
    pc("Levered: IRR, pre-tax", "%", lambda b: f"=Calc!$F${b['r_codl']}")
    pc("Levered: IRR, post-tax", "%", lambda b: f"=Calc!$E${b['r_codl']}", bold=True)
    d.row(None, "Levered: PSH profit", "EUR k", cases=f"={{c}}{r('cl_pl')}+{{c}}{r('psh_net')}")
    d.sub("Hold (FEP keeps the project; PSH bought out at the exit)")
    last = book.last_col
    pc("FEP equity invested (incl. the 2048 repower)", "EUR k",
       lambda b: f'=-SUMIF(Calc!$J${b["hold"]}:${last}${b["hold"]},"<0")', key="h_eq")
    pc("FEP distributions", "EUR k", lambda b: f'=SUMIF(Calc!$J${b["hold"]}:${last}${b["hold"]},">0")', key="h_dist")
    pc("FEP profit, pre-tax", "EUR k", lambda b: f"=Calc!$H${b['holdp']}")
    pc("FEP profit, post-tax", "EUR k", lambda b: f"=Calc!$H${b['hold']}", bold=True)
    pc("IRR, after tax", "%", lambda b: f"=Calc!$E${b['r_hold']}", bold=True)
    pc("IRR, before tax", "%", lambda b: f"=Calc!$F${b['r_hold']}")
    d.row(None, "MOIC, after tax", "x", cases="={h_dist}/{h_eq}")
    pc("NPV at the discount rate, after tax (at acquisition)", "EUR k", lambda b: f"=Calc!$E${b['v_hold']}")
    pc("Memo: project IRR, unlevered after tax (no buy-out)", "%",
       lambda b: f"=Calc!$E${b['r_proj']}", italic=True)
    pc("Memo: project IRR, unlevered before tax", "%", lambda b: f"=Calc!$F${b['r_proj']}", italic=True)
    pc("Levered: FEP equity invested (uses less debt)", "EUR k",
       lambda b: f'=-SUMIF(Calc!$J${b["holdl"]}:${last}${b["holdl"]},"<0")', key="hl_eq")
    pc("Levered: FEP distributions", "EUR k",
       lambda b: f'=SUMIF(Calc!$J${b["holdl"]}:${last}${b["holdl"]},">0")', key="hl_dist")
    pc("Levered: FEP profit, post-tax", "EUR k", lambda b: f"=Calc!$H${b['holdl']}")
    pc("Levered: IRR, after tax", "%", lambda b: f"=Calc!$E${b['r_holdl']}", bold=True)
    pc("Levered: IRR, before tax", "%", lambda b: f"=Calc!$F${b['r_holdl']}")
    d.row(None, "Levered: MOIC, after tax", "x", cases="={hl_dist}/{hl_eq}")
    d.row(None, "PSH buy-out at RtB entitlement, pending JV confirmation", "EUR k", cases=f"={{c}}{r('jv_pl')}")
    d.row(None, "PSH profit", "EUR k", cases=f"={{c}}{r('jv_pl')}+{{c}}{r('psh_net')}")

    d.blank()
    d.bar("JV split at the RtB exit (EUR k)", e=cl)
    d.row("jv_cons", "Consideration (premium × MWp)", "EUR k", cases="={p_rtb}*{MWp}")
    d.row("jv_reimb", "Buyer reimbursements (reimbursed and returned principal)", "EUR k", cases="=Buyer_Reimb")
    d.row("jv_avail", "Cash available", "EUR k", cases="={jv_cons}+{jv_reimb}", bold=True, top=True)
    d.row("jv_sale", "Less sale costs", "EUR k", cases="=-MIN({jv_avail},Sale_Cost_RtB*{jv_cons})")
    d.row("jv_prin", "Less FEP loan principal (DSA and reimbursed costs)", "EUR k",
          cases=f"=-MIN({{c}}{r('dl_dsa')}+Buyer_Reimb,SUM({{jv_avail}}:{{jv_sale}}))")
    d.row("jv_ret", "Less the 12 % return to FEP", "EUR k",
          cases=f"=-MIN({{c}}{r('dl_retx')},SUM({{jv_avail}}:{{jv_prin}}))")
    d.row("jv_spv", "Less the SPV price item (to FEP)", "EUR k", cases="=-MIN(SPV_Price,SUM({jv_avail}:{jv_ret}))")
    d.row("jv_dist", "Distributable", "EUR k", cases="=SUM({jv_avail}:{jv_spv})", bold=True, top=True)
    d.row("jv_share", "FEP share (60 % at ≤ 90 k€/MWp, else 50 %)", "%",
          cases="=IF({p_rtb}<=Share_Threshold,Share_Low,Share_High)")
    d.row("jv_fl", "FEP lot", "EUR k", cases="={jv_dist}*{jv_share}")
    d.row("jv_pl", "PSH lot", "EUR k", cases="={jv_dist}-{jv_fl}")
    d.name_row("PSH_Buyout", "jv_pl")
    d.row("jv_fep", "FEP total at the exit (principal, return, SPV, lot)", "EUR k",
          cases="=-SUM({jv_prin}:{jv_spv})+{jv_fl}", bold=True)
    d.row(None, "PSH total incl. the DSA received (JV sheet basis)", "EUR k", cases=f"={{jv_pl}}+{{c}}{r('dl_dsa')}",
          bold=True)

    d.blank()
    d.bar("Tax (EUR k)", e=cl)
    pc("Lifetime corporate income tax, hold route (SPV)", "EUR k", lambda b: f"=-Calc!$H${b['tax']}")
    pc("Tax losses lapsed", "EUR k", lambda b: f"=-Calc!$H${b['lapse']}")
    d.row(None, "FEP tax on the 12 % return, sale at RtB", "EUR k", cases=f"=-{{c}}{r('l_tax')}")
    d.row(None, "FEP tax on the 12 % return, sale at COD", "EUR k", cases=f"=-{{c}}{r('c_tax')}")

    d.blank()
    d.bar("Annual cash flow per case (EUR k)", h="Total")
    d.row("yr", "Year", "year", t="=Calc!{c}$4", fmt="int", bold=True)
    for i, b in enumerate(B):
        d.blank()
        rr = d.sub("", key=f"a{i}")
        d.ws.cell(rr, 3, f"=${'EFG'[i]}${book.case_row}")
        tot = "=SUM({self.all})"
        d.row(f"a{i}_inv", "Investment (capex, development, deposits, guarantees)", "EUR k", h=tot,
              t=f"=Calc!{{c}}${b['capt']}+Calc!{{c}}${b['pre']}")
        d.row(f"a{i}_rev", "Revenue", "EUR k", h=tot, t=f"=Calc!{{c}}${b['rev']}")
        d.row(f"a{i}_opex", "Opex", "EUR k", h=tot, t=f"=Calc!{{c}}${b['opex']}")
        d.row(f"a{i}_ebitda", "EBITDA", "EUR k", h=tot, t=f"=Calc!{{c}}${b['ebitda']}")
        d.row(f"a{i}_net", "Net cash flow after tax (project)", "EUR k", h=tot, t=f"=Calc!{{c}}${b['pcft']}", bold=True)
        d.row(f"a{i}_cum", "Cumulative net cash flow", "EUR k", t=f"=SUM($J{{a{i}_net.row}}:{{a{i}_net}})")

    d.blank()
    d.bar("Detailed calculations – expand the group", e=cl)
    g0 = d.r
    d.blank()
    db = DatedBlock(d, "Loan_Rate", returns={"x": "O", "c": "P"})
    book.dated = db
    db.header("RtB route: FEP dated cash flows and the 12 % return (EUR k)")
    db.end("x", "Accrual end: sale closing (L: factor in the exit year)", "=Exit_Date", "L")
    rc = db.end("c", "Accrual end: COD, live case (M: factor in COD year)",
                "=EDATE(COD_Date,Case_COD_Delay)", "M")
    book.define("COD_Live", f"Dashboard!$J${rc}")
    d.sub("Amounts used by the lines below", extra={"J": "Date"})
    d.row("dsa_amt", "DSA: amount per monthly draw", "EUR k", cases="=DSA_Per_MWp*{MWp}/DSA_Months")
    d.row("pre_amt", "Pre-RtB development: amount per month", "EUR k", e="=PreRtB_Total/PreRtB_Months")
    d.row("bg_rate", "Bank guarantee fee rate (EURIBOR + margin)", "%", e="=BG_Base+BG_Margin", fmt="0.00%")
    book.define("BG_Rate", d.ref("bg_rate", "E"))
    dep_h, bg_h = [], []
    for k, rr in enumerate(book.dep_rows):
        dep_h.append(d.row(None, f"Deposit (E: EUR k, F: released before exit): {DEPOSITS[k][0]}"[:56], "EUR k",
                           e=f"=Inputs!$E${rr}*Thousand/FX", f=f'=N(Inputs!$G${rr}<>"")',
                           extra={"J": f'=IF(Inputs!$G${rr}="",Inputs!$F${rr},Inputs!$G${rr})'},
                           fmts={"F": "flag", "J": "date"}))
    for k, rr in enumerate(book.bg_rows):
        bg_h.append(d.row(None, f"Guarantee (E: EUR k, F: days to release or exit) {k + 1}", "EUR k",
                          e=f"=Inputs!$E${rr}*Thousand/FX", f=f"=$J{d.r}-Inputs!$F${rr}",
                          extra={"J": f'=IF(Inputs!$G${rr}="",Exit_Date,Inputs!$G${rr})'},
                          fmts={"F": "int", "J": "date"}))
    d.sub("Outlays (all routes; the hold and COD routes take them by year)", extra={"J": "Date"})
    db.line("l_spv", "SPV acquisition price", "=-SPV_Price", "=Acq_Date", "=Ret_SPV")
    db.line("l_trans", "Transaction costs", "=-Transaction", "=Acq_Date", "=Ret_Trans")
    for k, (rr, h) in enumerate(zip(book.dep_rows, dep_h)):
        db.line("l_dep1" if k == 0 else None, f"Deposit paid: {DEPOSITS[k][0]}", f"=-$E${h}", f"=Inputs!$F${rr}",
                "=Ret_Dep")
    for k, h in enumerate(dep_h):
        db.line("l_dep_last" if k == len(dep_h) - 1 else None, f"Deposit released: {DEPOSITS[k][0]}",
                f"=$E${h}*$F${h}", f"=$J${h}", "=Ret_Dep")
    n_dsa = _months(DSA_FIRST, RTB)
    for k in range(n_dsa):
        key = "l_dsa1" if k == 0 else ("l_dsa_last" if k == n_dsa - 1 else None)
        date = "=DSA_First" if k == 0 else f"=EOMONTH(J{d.r - 1},1)"
        db.line(key, f"DSA draw {k + 1}", f"=-{{c}}${d.row_of('dsa_amt')}", date, "=Ret_DSA")
    for k, rr in enumerate(book.ew_rows):
        key = "l_ew1" if k == 0 else ("l_ew_last" if k == len(book.ew_rows) - 1 else None)
        huf = EARLY_WORKS[k][1] == "HUF m"
        db.line(key, f"Early works: {EARLY_WORKS[k][0].split(' (')[0]}",
                f"=-Inputs!$E${rr}" + ("*Thousand/FX" if huf else ""), f"=Inputs!$F${rr}", "=Ret_EW")
    n_pre = _months(PRE_FIRST, PRE_LAST)
    for k in range(n_pre):
        key = "l_pre1" if k == 0 else ("l_pre_last" if k == n_pre - 1 else None)
        date = "=PreRtB_First" if k == 0 else f"=EOMONTH(J{d.r - 1},1)"
        db.line(key, f"Pre-RtB development, month {k + 1}", f"=-$E${d.row_of('pre_amt')}", date, "=Ret_PreRtB")
    for k, h in enumerate(bg_h):
        key = "l_fee1" if k == 0 else ("l_fee_last" if k == len(bg_h) - 1 else None)
        db.line(key, f"Guarantee fee: {GUARANTEES[k][0]}", f"=-$E${h}*BG_Rate*$F${h}/Days_In_Year", f"=$J${h}",
                "=Ret_Fees")
    a, z = d.row_of("l_spv"), d.row_of("l_fee_last")
    book.define("Dated_Amounts", f"Dashboard!$E${a}:$G${z}")
    book.define("Dated_Dates", f"Dashboard!$J${a}:$J${z}")
    d.row("l_rec", "FEP receipts at the exit (JV split total)", "EUR k", cases=f"={{c}}{r('jv_fep')}",
          extra={"J": "=Exit_Date"}, fmts={"J": "date"})
    d.row("l_tax", "FEP tax: CIT on the 12 % return (and gain if taxed)", "EUR k",
          cases=f"=-CIT_Rate*({{c}}{r('t_ret')}+{{c}}{r('t_gain')})", extra={"J": "=Exit_Date"}, fmts={"J": "date"})
    rec, tx = d.row_of("l_rec"), d.row_of("l_tax")
    d.row("dl_pre", "FEP profit, pre-tax", "EUR k", cases=f"=SUM({{c}}{a}:{{c}}{rec})", bold=True, top=True)
    d.row("dl_post", "FEP profit, post-tax", "EUR k", cases=f"={{dl_pre}}+{{c}}{tx}", bold=True)
    d.row("dl_moic", "MOIC, pre-tax", "x", cases=f'=SUMIF({{c}}{a}:{{c}}{rec},">0")/-SUMIF({{c}}{a}:{{c}}{rec},"<0")')
    g = "MIN(XIRR_Guess,{dl_moic}-1)"                    # start guess: Excel's XIRR returns ~0 when it fails
    d.row("dl_irr_pre", "IRR, pre-tax", "%", cases=f"=XIRR({{c}}{a}:{{c}}{rec},$J{a}:$J{rec},{g})")
    d.row("dl_irr", "IRR, post-tax", "%", cases=f"=XIRR({{c}}{a}:{{c}}{tx},$J{a}:$J{tx},{g})", bold=True)
    d.row("dl_npv", "NPV at the discount rate, post-tax (at acquisition)", "EUR k",
          cases=f"=XNPV(Discount_Rate,{{c}}{a}:{{c}}{tx},$J{a}:$J{tx})")
    d.sub("Totals, taxable amounts and claims due (per case)")
    d.row("t_ret", "Taxable at FEP: the 12 % return received", "EUR k", cases=f"=-{{c}}{r('jv_ret')}")
    d.row("t_gain", "Taxable gain at FEP (if taxed)", "EUR k", cases="=Gain_Taxed*MAX(0,{dl_pre}-{t_ret})")
    d.row("dl_out", "Outlays to the exit, total (all routes)", "EUR k", cases=f"=SUM({{c}}{a}:{{c}}{z})")
    d.row("dl_dsa", "DSA loan principal (sum of the draws)", "EUR k",
          cases=f"=-SUM({{c}}{d.row_of('l_dsa1')}:{{c}}{d.row_of('l_dsa_last')})")
    db.claim("dl_retx", "12 % return due at the exit", "x", a, z)
    db.claim("dl_retc", "12 % return due at COD (sale at COD)", "c", a, z)
    d.sub("Hold and COD routes: at COD (all cases)")
    d.row("h_dep", "Cash deposits outstanding at the exit, returned at COD", "EUR k",
          e='=SUMPRODUCT(Dep_HUF*(Dep_Rel=""))*Thousand/FX')
    d.row("h_bg", "Bank guarantees outstanding at the exit", "EUR k", e='=SUMPRODUCT(BG_HUF*(BG_Rel=""))*Thousand/FX')
    d.row("h_fee", "Guarantee fees from the exit to COD", "EUR k", e="=-{h_bg}*BG_Rate*(COD_Live-Exit_Date)/Days_In_Year")
    d.row("h_net", "Net cash at COD (hold and COD routes)", "EUR k", e="={h_dep}+{h_fee}", bold=True, top=True)
    book.define("Hold_At_COD", d.ref("h_net", "E"))
    d.sub("PSH, sale at RtB (per case; DSA received over the DSA draws)", extra={"J": "Date"})
    d.row("psh_land", "PSH land apport (in kind)", "EUR k", cases="=-PSH_Land", extra={"J": "=Acq_Date"},
          fmts={"J": "date"})
    d.row("psh_dsa", "DSA received less PSH's own development costs", "EUR k", cases="=(DSA_Per_MWp-PSH_Devex)*{MWp}")
    d.row("psh_net", "PSH net own contribution (DSA − devex − land)", "EUR k", cases="={psh_land}+{psh_dsa}")
    d.row("psh_lot", "PSH lot at the exit", "EUR k", cases=f"={{c}}{r('jv_pl')}", extra={"J": "=Exit_Date"},
          fmts={"J": "date"})
    d.row("psh_profit", "PSH profit, sale at RtB", "EUR k", cases="={psh_lot}+{psh_net}", bold=True, top=True)

    d.blank()
    d.bar("Sale at COD: JV split at the COD year end (EUR k)", e=cl)
    d.row("c_ev", "Price (enterprise value)", "EUR k", cases="={ev}")
    pc("Less sale costs", "EUR k", lambda b: "=-MIN({c_ev},Sale_Cost_COD*{c_ev})", key="c_sale")
    pc("Less FEP funding returned (contributions to COD)", "EUR k",
       lambda b: f"=-MIN(Calc!$E${b['codr']},SUM({{c_ev}}:{{c_sale}}))", key="c_fund")
    d.row("c_ret", "Less the 12 % return to COD", "EUR k", cases="=-MIN({dl_retc},SUM({c_ev}:{c_fund}))")
    d.row("c_dist", "Distributable", "EUR k", cases="=SUM({c_ev}:{c_ret})", bold=True, top=True)
    d.row("c_share", "FEP share (RtB-price rule; pending JV confirmation)", "%", cases="={jv_share}")
    d.row("c_fl", "FEP lot", "EUR k", cases="={c_dist}*{c_share}")
    d.row("c_pl", "PSH lot", "EUR k", cases="={c_dist}-{c_fl}")
    d.row("c_fep", "FEP total at COD (funding, return, lot)", "EUR k", cases="=-SUM({c_fund}:{c_ret})+{c_fl}", bold=True)
    d.name_row("COD_FEP_Total", "c_fep")
    d.row("c_tret", "Taxable at FEP: the 12 % return at COD", "EUR k", cases="=-{c_ret}")
    pc("Taxable gain at COD (if taxed)", "EUR k", lambda b: f"=Gain_Taxed*MAX(0,Calc!$H${b['codr']}-{{c_tret}})",
       key="c_tgain")
    d.row("c_tax", "FEP tax at COD: CIT on the return (and gain if taxed)", "EUR k",
          cases="=-CIT_Rate*({c_tret}+{c_tgain})")
    d.name_row("COD_FEP_Tax", "c_tax")

    d.blank()
    d.bar("Sale at COD, levered: FEP repays the debt from the price, then the JV split", e=cl)
    pc("Price less the debt repaid at COD", "EUR k", lambda b: f"={{ev}}-Calc!$E${b['cbc']}", key="cl_ev")
    d.row("cl_sale", "Less sale costs (on the price)", "EUR k", cases="=-MIN({cl_ev},Sale_Cost_COD*{ev})")
    pc("Less FEP equity returned (equity to COD)", "EUR k",
       lambda b: f"=-MIN(Calc!$E${b['codrl']},SUM({{cl_ev}}:{{cl_sale}}))", key="cl_fund")
    d.row("cl_ret", "Less the 12 % return to COD", "EUR k", cases="=-MIN({dl_retc},SUM({cl_ev}:{cl_fund}))")
    d.row("cl_dist", "Distributable", "EUR k", cases="=SUM({cl_ev}:{cl_ret})", bold=True, top=True)
    d.row("cl_fl", "FEP lot (RtB-price share rule)", "EUR k", cases="={cl_dist}*{c_share}")
    d.row("cl_pl", "PSH lot", "EUR k", cases="={cl_dist}-{cl_fl}")
    d.row("cl_fep", "FEP total at COD, levered (equity, return, lot)", "EUR k",
          cases="=-SUM({cl_fund}:{cl_ret})+{cl_fl}", bold=True)
    d.name_row("CODL_FEP_Total", "cl_fep")
    d.row("cl_tret", "Taxable at FEP: the 12 % return at COD, levered", "EUR k", cases="=-{cl_ret}")
    pc("Taxable gain at COD, levered (if taxed)", "EUR k",
       lambda b: f"=Gain_Taxed*MAX(0,Calc!$H${b['codrl']}-{{cl_tret}})", key="cl_tgain")
    d.row("cl_tax", "FEP tax at COD, levered: CIT on the return (and gain)", "EUR k",
          cases="=-CIT_Rate*({cl_tret}+{cl_tgain})")
    d.name_row("CODL_FEP_Tax", "cl_tax")
    d.group(g0, d.r - 1)
    d.end()
    return d


# =========================================================================================== Checks
def checks(book):
    k = book.sheet("Checks", "Zalahaláp – checks (0 = OK), mirrored at Dashboard H3")
    k.bar("Checks (0 = OK)", e=_case_links(book, k), h="Total")
    B, D_, rv = book.blocks, book.sheets["Dashboard"], book.rev
    LR = book.lever_rows
    last = book.last_col
    dr = D_.row_of
    first = k.r
    a, z = dr("l_spv"), dr("l_fee_last")

    def ck(label, fn, shared=False):
        """fn(i, block, col) -> formula; shared checks are written in E only."""
        cells = {}
        for i, b in enumerate(B):
            if shared and i:
                break
            cells["efg"[i]] = fn(i, b, "EFG"[i])
        k.row(None, label, "flag", **cells, h="=SUM({self.cases})", fmt="flag")

    ck("Lever: BESS duration has an Aurora block (2 h or 4 h)",
       lambda i, b, C: f"=Dashboard!{C}${LR['BESS_On']}*ISNA(MATCH(Dashboard!{C}${LR['block']},Aurora_Blocks,0))")
    ck("Lever: revenue levers not negative",
       lambda i, b, C: f"=N(MIN(Dashboard!{C}${LR['PV_Price']}:{C}${LR['BESS_Income']})<0)")
    ck("Lever: fixed term within the operating life", lambda i, b, C: f"=N(Dashboard!{C}${LR['Fix_Years']}>Life)")
    ck("Timing: COD year not before Aurora's first year", lambda i, b, C: "=N(YEAR(COD_Live)<Aurora_First)",
       shared=True)
    ck("Timing: operating life inside the timeline", lambda i, b, C: f"=N(ABS(Calc!$H${b['oshare']}-Life)>Tolerance)")
    ck("Energy: sent + exported directly = generation",
       lambda i, b, C: f"=N(ABS(Calc!$H${b['sent']}+Calc!$H${b['direct']}-Calc!$H${b['gen']})>Tolerance)")
    ck("Energy: discharge ≤ PV charged × RTE, every year",
       lambda i, b, C: f"=SUMPRODUCT(--(Calc!J{b['dis']}:{last}{b['dis']}>Calc!J{b['sent']}:{last}{b['sent']}"
                       f"*BESS_RTE+Tolerance))")
    ck("Grid: direct export ≤ MVA × 8,760 h × operating share",
       lambda i, b, C: f"=SUMPRODUCT(--(Calc!J{b['direct']}:{last}{b['direct']}>Dashboard!{C}${LR['MVA']}*Hours_In_Year"
                       f"*Calc!J{b['oshare']}:{last}{b['oshare']}))")
    ck("BESS rows are zero when no BESS is built",
       lambda i, b, C: f"=(1-Dashboard!{C}${LR['BESS_On']})*N(ABS(Calc!$H${b['sent']})+ABS(Calc!$H${b['dis']})"
                       f"+ABS(Calc!$H${b['rbess']})>Tolerance)")
    ck("Revenue: PV price received = fixed price inside the term",
       lambda i, b, C: f"=SUMPRODUCT((Revenue!J{rv[i]['fts']}:{last}{rv[i]['fts']}=1)"
                       f"*(ABS(Revenue!J{rv[i]['pxd']}:{last}{rv[i]['pxd']}-Dashboard!{C}${LR['PV_Price']})>Tolerance))")
    ck("Capex: construction spread in full (= total uses)",
       lambda i, b, C: f"=N(ABS(Calc!$H${b['capex']}+Calc!$H${b['ewntp']}+Calc!$E${b['capex']})>Tolerance)")
    ck("Dated outlays all mapped to the annual grid",
       lambda i, b, C: f"=N(ABS(Calc!$H${b['pre']}-Hold_At_COD-Dashboard!{C}{dr('dl_out')})>Tolerance)")
    ck("Dated: DSA draws add up to the DSA loan",
       lambda i, b, C: f"=N(ABS(Dashboard!{C}{dr('dl_dsa')}-DSA_Per_MWp*Dashboard!{C}${LR['MWp']})>Tolerance)")
    ck("Dated: pre-RtB development months add up to the total",
       lambda i, b, C: "=N(ABS(PreRtB_Spent-PreRtB_Total)>Tolerance)", shared=True)
    ck("Dated: the cases share the outlays (returns per line)",
       lambda i, b, C: f"=--(SUMPRODUCT(ABS(Dashboard!$E${a}:$E${z}-Dashboard!{C}{a}:{C}{z}))>Tolerance)")
    ck("Dated: DSA and pre-RtB schedules end on their dates",
       lambda i, b, C: "=N(EOMONTH(DSA_First,DSA_Months-1)<>RtB_Date)"
                       "+N(EOMONTH(PreRtB_First,PreRtB_Months-1)<>PreRtB_Last)", shared=True)
    ck("Dated: return-earning lines paid before the exit year",
       lambda i, b, C: "=" + book.dated.check_before_end("x"), shared=True)
    ck("Dated: deposit and guarantee releases not after the exit",
       lambda i, b, C: '=N(SUMPRODUCT((Dep_Rel<>"")*(Dep_Rel>Exit_Date))+SUMPRODUCT((BG_Rel<>"")'
                       '*(BG_Rel>Exit_Date))>0)', shared=True)
    ck("JV at RtB: FEP + PSH + costs = price + reimbursements",
       lambda i, b, C: f"=N(ABS(Dashboard!{C}{dr('jv_fep')}+Dashboard!{C}{dr('jv_pl')}-Dashboard!{C}{dr('jv_sale')}"
                       f"-Dashboard!{C}{dr('jv_avail')})>Tolerance)")
    ck("JV at RtB: lots = distributable, distributable ≥ 0",
       lambda i, b, C: f"=N(ABS(Dashboard!{C}{dr('jv_fl')}+Dashboard!{C}{dr('jv_pl')}-Dashboard!{C}{dr('jv_dist')})"
                       f">Tolerance)+N(Dashboard!{C}{dr('jv_dist')}<-Tolerance)")
    ck("JV at COD: FEP total + PSH lot + sale costs = price",
       lambda i, b, C: f"=N(ABS(Dashboard!{C}{dr('c_fep')}+Dashboard!{C}{dr('c_pl')}-Dashboard!{C}{dr('c_sale')}"
                       f"-Dashboard!{C}{dr('c_ev')})>Tolerance)")
    ck("JV at COD: lots = distributable, distributable ≥ 0",
       lambda i, b, C: f"=N(ABS(Dashboard!{C}{dr('c_fl')}+Dashboard!{C}{dr('c_pl')}-Dashboard!{C}{dr('c_dist')})"
                       f">Tolerance)+N(Dashboard!{C}{dr('c_dist')}<-Tolerance)")
    ck("Tax: losses carried forward never negative",
       lambda i, b, C: f"=N(MIN(Calc!J{b['lc']}:{last}{b['lc']})<-Tolerance)")
    ck("Returns: IRR n/a only when the money multiple is below 1",
       lambda i, b, C: f'=COUNTIFS(Calc!$E${b["r_proj"]}:$E${b["r_bl"]},"n/a",Calc!$G${b["r_proj"]}:$G${b["r_bl"]},'
                       f'">=1")')
    ck("Returns: the RtB-route XIRRs are numbers",
       lambda i, b, C: f"=N(COUNT(Dashboard!{C}{dr('dl_irr_pre')}:{C}{dr('dl_irr')})<COUNTA(Dashboard!{C}"
                       f"{dr('dl_irr_pre')}:{C}{dr('dl_irr')}))")
    ck("Summary: net cash flow = investment + EBITDA + tax",
       lambda i, b, C: f"=N(ABS(Dashboard!$H${dr(f'a{i}_net')}-Dashboard!$H${dr(f'a{i}_inv')}"
                       f"-Dashboard!$H${dr(f'a{i}_ebitda')}-Calc!$H${b['tax']})>Tolerance)")
    gr = D_.row_of("Gearing")
    ck("Timeline: first grid year end = the model date", lambda i, b, C: "=N(Calc!$J$5<>Model_Date)", shared=True)
    ck("Timeline: first column = dated spend to the model date",
       lambda i, b, C: f'=--(ABS(Calc!$J${b["pre"]}-SUMIF(Dated_Dates,"<="&Model_Date,Dashboard!{C}{a}:{C}{z}))'
                       f'>Tolerance)')
    ck("Capex: all construction capex in the capex year",
       lambda i, b, C: f"=--(ABS(INDEX(Calc!$J${b['capex']}:${last}${b['capex']},Capex_Year-First_Year+1)"
                       f"-Calc!$H${b['capex']})>Tolerance)")
    ck("Debt: drawn = gearing × total uses to COD",
       lambda i, b, C: f"=N(ABS(Calc!$H${b['draw']}-Calc!$E${b['fin']}*Calc!$E${b['capex']})>Tolerance)")
    ck("Debt: balance before the COD year = debt at COD",
       lambda i, b, C: f"=--(ABS(INDEX(Calc!$J${b['dclose']}:${last}${b['dclose']},Calc!$E$8-1)"
                       f"-Calc!$E${b['cbc']})>Tolerance)")
    ck("Debt: repaid by the last annuity and never negative",
       lambda i, b, C: f"=N(ABS(Calc!${last}${b['dclose']})>Tolerance)+N(MIN(Calc!J{b['dclose']}:{last}"
                       f"{b['dclose']})<-Tolerance)")
    ck("Debt: gearing 0 reproduces the unlevered IRRs",
       lambda i, b, C: f"=N(Dashboard!{C}{gr}=0)*--(ABS(N(Calc!E{b['r_holdl']})-N(Calc!E{b['r_hold']}))"
                       f"+ABS(N(Calc!E{b['r_codl']})-N(Calc!E{b['r_cod']}))>Tolerance)")
    ck("JV at COD, levered: FEP + PSH + costs = price − debt",
       lambda i, b, C: f"=N(ABS(Dashboard!{C}{dr('cl_fep')}+Dashboard!{C}{dr('cl_pl')}-Dashboard!{C}{dr('cl_sale')}"
                       f"-Dashboard!{C}{dr('cl_ev')})>Tolerance)")
    lastc = k.r - 1
    k.blank()
    r = k.row(None, "Master check (sum of the check rows)", "flag", h=f"=SUM(H{first}:H{lastc})", bold=True,
              fmt="check")
    check_format(k.ws, f"H{r}")
    book.define("Master_Check", f"Checks!$H${r}")
    k.blank()
    k.sub("Information (economic outcomes, not errors: outside the master check)")
    ck("Implied RtB premium floored at 0",
       lambda i, b, C: f"=Dashboard!{C}{dr('impneg')}")
    ck("Levered buyer IRR at COD n/a (price below the debt)",
       lambda i, b, C: f"=N(ISTEXT(Calc!$E${b['r_bcl']}))")
    ck("IRRs n/a: flows never pay back (money multiple below 1)",
       lambda i, b, C: f'=COUNTIF(Calc!$E${b["r_proj"]}:$F${b["r_bl"]},"n/a")')
    ck("Years with CFADS below debt service (DSCR < 1)",
       lambda i, b, C: f"=SUMPRODUCT(--(Calc!J{b['cfads']}:{last}{b['cfads']}<Calc!J{b['dserv']}:{last}{b['dserv']}"
                       f"-Tolerance))")
    k.end()
    return k


# =========================================================================================== build
def model(R=None):
    """Build the workbook. The Dashboard reads rows written further down (dated block, JV splits): a first pass
    lays it out, the second writes it with the row numbers of the first (the layout does not depend on them)."""
    book = Book(CASES, years=N_YEARS)
    d = dashboard_top(book)
    inputs(book)
    book.rv = revenue(book)
    calc(book, R or {})
    dashboard(book, d, R or {})
    if R is None:
        return model(dict(d.rows))
    checks(book)
    outputs(book, d)
    _label_widths(book)
    for sh, r in (("Revenue", 4), ("Calc", 4), ("Dashboard", d.row_of("yr"))):
        book.sheets[sh].ws.cell(r, 10).number_format = '"to "0'      # first column: spend to the model date
    return book


def _label_widths(book):
    """Every labelled row (unit in D) and every title fits column C (56 characters; Inputs 54): no overflow."""
    for sh in book.sheets.values():
        width = sh.ws.column_dimensions["C"].width
        for r in range(1, sh.ws.max_row + 1):
            v = sh.ws.cell(r, 3).value
            if isinstance(v, str) and not v.startswith("=") and (r == 2 or sh.ws.cell(r, 4).value):
                assert len(v) <= width, f"{sh.ws.title}!C{r} {len(v)} > {width}: {v}"


HEADLINES = [  # Out_<name>_<case> for the audit: (name, Dashboard label, occurrence)
    ("Hold_IRR", "IRR, after tax", 0), ("Hold_IRR_Pre", "IRR, before tax", 0), ("Hold_MOIC", "MOIC, after tax", 0),
    ("RtB_IRR", "IRR, post-tax", 0), ("RtB_Profit", "FEP profit, post-tax", 0), ("RtB_MOIC", "MOIC, pre-tax", 0),
    ("COD_IRR", "IRR, post-tax", 1), ("COD_Profit", "FEP profit, post-tax", 1),
    ("Implied_RtB", "Implied premium at the buyer target IRR (floored at 0)", 0),
    ("Implied_COD", "Implied COD price at the buyer target IRR (floored at 0)", 0),
    ("Buyer_RtB_IRR", "Buyer IRR at our price, unlevered after tax", 0),
    ("Hold_IRR_Lev", "Levered: IRR, after tax", 0), ("Debt_COD", "Debt at COD", 0),
    ("Min_DSCR", "Minimum DSCR (information)", 0)]


def outputs(book, d):
    seen = {}
    for r in range(1, d.ws.max_row + 1):
        lab = d.ws.cell(r, 3).value
        if isinstance(lab, str) and not d.ws.row_dimensions[r].hidden:
            seen.setdefault(lab, []).append(r)
    for name, lab, nth in HEADLINES:
        for i, col in enumerate("EFG", 1):
            book.define(f"Out_{name}_{i}", f"Dashboard!${col}${seen[lab][nth]}")


def build_model(out: Path) -> Path:
    book = model()
    return Path(book.save(out / f"{NAME}.xlsx"))


def build(out: Path):
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    return {"model": build_model(out)}
