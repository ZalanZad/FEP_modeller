"""Zalahaláp – 63 MWp PV, with or without a 50 MW BESS: investment model, SPEC v2 (wind-model format).

One workbook, `out/Zalahalap_Model.xlsx` (no macros): Dashboard · Inputs · Revenue · Calc · Checks, annual 2025–2064,
three cases as columns (E PV only, F PV + 2 h BESS, G PV + 4 h BESS), unlevered. PV sells under a 5-year PPA at the
lever price on all output, then at Aurora's merchant prices; the BESS earns Aurora's block uplift × the PV-only
factor (or the MET tolling fee). Base / Low / High through `Dash_Case`. The pre-COD cash is dated on the Dashboard
(RtB route, XIRR) and mapped to the annual grid by year for the hold and COD routes.
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
FIRST_YEAR, N_YEARS = 2025, 40                          # 2025 … 2064 (the Low case's 12-month delay runs to 2064)
YEARS = list(range(FIRST_YEAR, FIRST_YEAR + N_YEARS))
CASES = ["PV only", "PV + 2h BESS", "PV + 4h BESS"]
BLOCKS = {"DC 4h": "2027 solar DC 1.0cy 4h", "DC 2h": "2027 solar DC 1.0cy 2h"}     # code -> Aurora block
PH = "placeholder – pending user"
EW_NOTE = "to be revisited, user 22-Oct-26"
TITLE = "Project Zalahaláp – 63 MWp PV, with or without a 50 MW BESS: sale at RtB, sale at COD or hold"

# dated tables (Inputs, flat rows): label, unit, amount, date, source
EARLY_WORKS = [("Land, 1st payment", "HUF m", 60, D(2025, 11, 30), "ZHLP v3 Stat_inp E44"),
               ("Land, 2nd payment", "HUF m", 60, D(2026, 2, 28), "ZHLP v3 Stat_inp E45"),
               ("ICS execution design, 1st payment", "EUR k", 25, D(2026, 2, 28), "ZHLP v3 Stat_inp E46"),
               ("Critical equipment, 1st payment (HV, ICS, 132 kV transformer)", "EUR k", 460, D(2026, 4, 30),
                "ZHLP v3 Stat_inp E47")]
DEPOSITS = [("Connection-fee reimbursement", 45, D(2025, 11, 30), None, "ZHLP v3 Stat_inp E92; Dyn_Inp row 12"),
            ("Capacity guarantee, cash", 180, D(2025, 11, 30), None, "ZHLP v3 Stat_inp E93; Dyn_Inp row 13"),
            ("Additional guarantee, cash instalments", 485.55, D(2025, 11, 30), D(2026, 1, 31),
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


def _years(vals, first=None, grow=None):
    """Aurora 2026–2060 → model years: 2025 = `first` (blank), flat real after 2060 (or grown by `grow`)."""
    out = [first] + list(vals)
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
    lv("Yield", "Specific yield", "MWh/MWp", [1265, 1265, 1265], source="Teaser p1 (79.7 GWh); offerors 1,199–1,245")
    lv("BESS_On", "BESS built (1 = yes)", "flag", [0, 1, 1], source="User 03-Nov-2026", options=[0, 1])
    lv("BESS_MW", "BESS power", "MW", [0, 50, 50], source="Teaser p1")
    lv("Hours", "BESS duration (2 or 4 hours)", "h", [0, 2, 4], source="User 03-Nov-2026", options=[0, 2, 4])
    d.row("MWh", "BESS energy", "MWh", cases="={c}$10*{c}$11*{c}$12")
    book.define("BESS_MWh", d.ref("MWh", "cases"))
    d.row("block", "Aurora block (BESS revenue and volumes)", "text",
              cases='=IF({c}$10=1,"DC "&{c}$12&"h","–")', fmt="general")
    book.define("Block_Code", d.ref("block", "cases"))
    d.row("pblock", "Aurora block for the PV prices (PV only: 2 h block, proxy)", "text",
          cases="=IF({c}$10=1,{c}$14,PV_Only_Block)", fmt="general")
    book.define("Price_Block", d.ref("pblock", "cases"))
    lv("PPA_Price", "PPA price, nominal as offered", "EUR/MWh", [45, 45, 45],
       source="MVM 5-y offer (Revenue: Offer_PPA_MVM)", fmt="money2")
    lv("PPA_Index", "PPA indexation", "%/yr", [0, 0, 0], source="MVM: none stated")
    lv("PPA_Years", "PPA tenor from COD", "years", [5, 5, 5], source="User 03-Nov-2026: 5-year PPA")
    lv("Basis", "BESS revenue basis (Merchant or Tolling)", "text", ["Merchant"] * 3,
       source="Aurora merchant or MET tolling", options="Merchant,Tolling")
    lv("RtB_Price", "RtB price, premium (0 = implied)", "EUR k/MWp", [90, 90, 90],
       source="ZHLP v3 Stat_inp E24; " + PH)
    lv("COD_Price", "COD price (0 = implied)", "EUR k/MWp", [0, 0, 0], source="Implied at the buyer target IRR")
    d.row("toll_cfg", "Tolling offer ≠ built configuration (1 = fee not scaled; MET: 90 MWh)", "flag",
          cases='=N(AND({c}$19="Tolling",{c}$13<>Offer_MET_MWh))')
    rows = {k: d.row_of(k) for k in ("MWp", "MVA", "BESS_On", "BESS_MW", "Hours", "MWh", "block", "pblock",
                                      "PPA_Years", "Basis", "RtB_Price", "COD_Price")}
    assert rows == {"MWp": 7, "MVA": 8, "BESS_On": 10, "BESS_MW": 11, "Hours": 12, "MWh": 13, "block": 14,
                    "pblock": 15, "PPA_Years": 18, "Basis": 19, "RtB_Price": 20, "COD_Price": 21}, rows
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
            ("COD_Delay", "COD delay", "months", [0, 12, 0], "SPEC v2 §7", None),
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
    i = book.sheet("Inputs", "Project Zalahaláp – inputs (blue) with units and sources; value in column E",
                   widths={"C": 54, "D": 11, "H": 60})
    i.bar("Inputs", e=_case_links(book, i), h="Source")
    a = i.input

    i.sub("Constants")
    a("Thousand", "Thousand", "#", 1000, "Scale")
    a("Days_In_Year", "Days in a year", "days", 365, "Day count act/365")
    a("Months_In_Year", "Months in a year", "months", 12, "Calendar")
    a("Hours_In_Year", "Hours in a year", "h", 8760, "Grid export check (365 × 24)")
    a("Tolerance", "Check tolerance", "#", 0.000001, "Checks", fmt="general")
    a("First_Year", "First year of the timeline", "year", FIRST_YEAR, "Acquisition year")

    i.sub("Dates")
    a("Acq_Date", "Acquisition, JV start, deposits paid", "date", D(2025, 11, 30), "ZHLP v3 Stat_inp E39")
    a("RtB_Date", "Ready to build (last DSA draw)", "date", RTB, "Teaser p6; " + PH)
    a("Exit_Date", "RtB sale closing (exit)", "date", D(2027, 3, 31), PH)
    a("NTP_Date", "Notice to proceed (construction start)", "date", D(2027, 6, 30), PH)
    a("COD_Date", "COD, PV and BESS (operations from the next day)", "date", D(2028, 3, 31), "Teaser p6; " + PH)
    a("Life", "Operating life", "years", 35, "ZHLP v3 Control D30")
    a("Aurora_First", "Aurora's first year (2025 blank; flat real after 2060)", "year", 2026, "Aurora Q2-26 databook")

    i.sub("PV")
    a("PV_Avail", "PV availability", "%", 0.995, "ZHLP v3 Control D56")
    a("Degradation", "PV degradation a year (from the first full year)", "%", 0.003, "ZHLP v3 Control D57",
      fmt="0.00%")
    a("Clipping", "Clipping and curtailment (DC-coupled BESS)", "%", 0, "Aurora s+29 = 0; " + PH)
    a("GoO_Price", "Guarantee of origin price, real 2025", "EUR/MWh", 1.75, "ZHLP v3 Control D58; " + PH, fmt="money2")
    a("PV_Only_Block", "PV-only price: Aurora block used as capture proxy", "text", "DC 2h",
      "PV-only capture proxy: no solar row in Aurora", text=True)

    i.sub("BESS (the Aurora blocks carry RTE 0.86 and availability 98 %)")
    a("BESS_RTE", "Round-trip efficiency (discharge ≤ PV charged × RTE)", "factor", 0.86, "Aurora Inputs C17")
    a("BESS_Haircut", "Haircut on the merchant BESS revenue", "%", 0, PH)
    a("AFRR_Share", "aFRR-down / FCR capacity deliverable from PV charge", "%", 0, "SPEC v1 §0.6; " + PH)
    a("BESS_Capex", "BESS capex and augmentation, real 2025", "EUR k/MWh", 110, "ZHLP v3 Control D94; " + PH)
    a("Repower", "BESS repowered at the block's repower year (0 = retired)", "flag", 1,
      "Aurora repower count s+1; ZHLP v3 Control D100")
    a("BESS_Opex", "BESS fixed opex, real 2025", "EUR k/MW/yr", 30.8, "ZHLP v3 Control D96")

    i.sub("Contracts (PPA price and tenor: Dashboard levers; offer terms: Revenue)")
    a("Balancing", "Balancing cost, merchant years only, real 2025", "EUR/MWh", 5, "ZHLP v3 Control D60")
    a("Toll_Start", "Tolling start (fixed fee, then merchant)", "date", D(2028, 1, 1), "MET offer p5")
    a("Toll_Years", "Tolling tenor", "years", 3, "MET offer p5 (4-year option 46.4)")
    i.note("PPA from COD, pay-as-produced on all PV output at the nominal lever price; tolling fee read from "
           "Revenue (Offer_Tolling_MET_3y).")

    i.sub("Capex to COD (2025 prices)")
    a("PV_Capex", "PV equipment and balance of plant", "EUR k/MWp", 350, "ZHLP v3 Control D82:D83 (330 + 20)")
    a("Cable_Capex", "MV cable to the Zalahaláp substation (12 km)", "EUR k", 3700, "ZHLP v3 Control D104; " + PH)
    a("Substation", "HV substation and bay", "EUR k", 9500, "ZHLP v3 Control D105; " + PH)
    a("Grid_Fee", "Grid connection fee", "EUR k", 0, "Teaser p1 (2018 rights)")
    a("Land_Capex", "Land to COD (not depreciated)", "EUR k", 620, "ZHLP v3 Control D106")
    a("EW_NTP", "Early works at NTP (E48 650 + E49 75)", "EUR k", 725, "ZHLP v3 Stat_inp E48:E49; " + EW_NOTE)

    i.sub("Opex (real 2025, × the Aurora euro index)")
    a("OM_Cost", "O&M and fixed opex", "EUR k/MWp/yr", 14.96, "ZHLP v3 Control D59")
    a("Cable_OM", "MV cable O&M", "EUR k/MWp/yr", 5, "ZHLP v3 Control D62")
    a("Lease", "Land lease from COD", "EUR k", 96, "ZHLP v3 Control D61")
    a("HIPA", "Local business tax (HIPA), on revenue", "%", 0.02, "Bana; " + PH)
    a("Innovation", "Innovation contribution, on revenue", "%", 0.003, "Act LXXVI/2014; " + PH, fmt="0.00%")

    i.sub("Development and JV")
    a("SPV_Price", "FEP SPV acquisition price (to FEP before the split)", "EUR k", 549, "JV sheet 221025 G19; " + PH)
    a("Transaction", "Transaction costs (FEP)", "EUR k", 200, "ZHLP v3 Stat_inp E37")
    a("DSA_Per_MWp", "DSA loan to PSH", "EUR k/MWp", 15, "JV sheet 221025 F55; " + PH)
    a("DSA_First", "DSA: first monthly draw (equal draws to RtB)", "date", DSA_FIRST, "v1 linear spread; " + PH)
    a("PreRtB_Total", "Pre-RtB development costs (equal monthly)", "EUR k", 3182, "ZHLP v3 Control D115")
    a("PreRtB_First", "Pre-RtB development: first month end", "date", PRE_FIRST, "ZHLP v3 Control D116")
    a("PreRtB_Last", "Pre-RtB development: last month end", "date", PRE_LAST, "ZHLP v3 Control D117")
    a("FX", "FX rate", "HUF/EUR", 400, "ZHLP v3 Stat_inp E23 (Aurora 397.77)")
    a("BG_Base", "Bank guarantee fee: 1M EURIBOR", "%", 0.0187, "ZHLP v3 Dyn_Inp G50:Y50", fmt="0.00%")
    a("BG_Margin", "Bank guarantee fee: margin", "%", 0.0175, "ZHLP v3 Stat_inp E110", fmt="0.00%")
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
    i.note("Buyer target IRR: Dashboard sensitivity case (Case_Buyer_IRR, unlevered after tax).")

    i.sub("Financing")
    i.note("Debt – out of scope for now (user 03-Nov-2026); the model is unlevered.", italic=False)
    i.end()
    return i


# =========================================================================================== Revenue
def revenue(book):
    rv = book.sheet("Revenue", "Project Zalahaláp – revenue: Aurora curves, offtake offers and the prices each case "
                               "earns", timeline=True, widths={"H": 30})
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
    vals = _years(au[("euro_index", "Central")][0], first=1.0, grow=0.02)
    cells = {L(10 + k): v for k, v in enumerate(vals)}
    rv.row("idx", "Aurora euro index (2025 = 1; +2 % a year after 2060)", "factor", inputs=tuple(cells),
           extra=cells, fmt="0.0000", source="Aurora Common Assumptions row 14")
    book.idx = rv.row_of("idx")
    rep = {}
    for code, blk in BLOCKS.items():
        rc = au[("repower_count", blk)][0]
        rep[code] = next(2026 + k for k, x in enumerate(rc) if x >= 1)
    rv.row("rep", "BESS repower year (E: DC 4 h, F: DC 2 h; first year with a repower)", "year", e=rep["DC 4h"],
           f=rep["DC 2h"], inputs=("E", "F"), source="Aurora Degraded batteries s+1 (rows 140, 97)")
    book.define("Aurora_Blocks", f"Revenue!$E${rv.row_of('pvd4')}:$E${rv.row_of('pvd2')}")
    book.define("Repower_Year", f"Revenue!$E${rv.row_of('rep')}:$F${rv.row_of('rep')}")
    rv.note("Aurora's BESS stack is a fixed dispatch per block: it does not re-optimise when prices or the PPA change. "
            "Rows 2061–2064 repeat 2060 (flat real).")
    rv.note("PV only: no standalone-solar capture in the databook; the DC 2 h block's direct price is the PV-only "
            "capture proxy (flagged).")

    rv.blank()
    rv.bar("Offtake offers as stated – indicative; E = E.ON, F = MET/E2, G = MVM; H = PDF page",
           e=["E.ON", "MET / E2", "MVM"], h="Source")
    offers(book, rv)

    rv.blank()
    rv.bar("Resulting prices by case, nominal – PPA in its tenor, then the Aurora merchant tail",
           e=_case_links(book, rv))
    rv.row("yr", "Year", "year", t="={c}$4", fmt="int", bold=True)
    rv.sub("Contract timing (live case)")
    rv.row("m0", "Months from COD to the start of the year", "months", t="=({year}-YEAR(COD_Live))*Months_In_Year"
                                                                          "-MONTH(COD_Live)", fmt="int")
    rv.row("opm", "Operating months in the year", "months",
           t="=MAX(0,MIN(Months_In_Year,Life*Months_In_Year-{m0})-MAX(0,-{m0}))", fmt="int")
    rv.row("t0", "Months from the tolling start to the start of the year", "months",
           t="=({year}-YEAR(Toll_Start))*Months_In_Year-MONTH(Toll_Start)+1", fmt="int")
    rv.row("bfac", "BESS merchant revenue factor: (1 − haircut) × (1 + case change)", "factor",
           e="=(1-BESS_Haircut)*(1+Case_BESS)")
    a, b = rv.row_of("pvd4"), rv.row_of("pvd2")
    idx = f"{{c}}${book.idx}"
    rows = []
    prev = None
    for k in range(1, 4):
        r = rv.sub("", e=["=1" if prev is None else f"=$E${prev}+1"], key=f"s{k}")
        rv.ws.cell(r, 3, f"=INDEX(Dash_Name,1,$E${r})")
        rv.ws.cell(r, 5).number_format = '"Case "0'
        rv.scope = f"c{k}_"
        cs = f"$E${r}"
        rv.row("pos", "Merchant PV price, direct export (E: Aurora block position; F: BESS factor)", "EUR/MWh",
               e=f"=MATCH(INDEX(Price_Block,1,{cs}),Aurora_Blocks,0)",
               f=f"=INDEX(Dash_BESS_On,1,{cs})*$E${rv.row_of('bfac')}",
               t=f"=INDEX({{c}}${a}:{{c}}${b},{{self.E}})*{idx}*(1+Case_PV_Price)",
               fmts={"E": "int", "F": "factor"})
        rv.row("pvs", "Merchant PV price, stored energy (transfer price)", "EUR/MWh",
               t=f"=INDEX({{c}}${rv.row_of('pvt4')}:{{c}}${rv.row_of('pvt2')},{{pos.E}})*{idx}*(1+Case_PV_Price)")
        rv.row("ppa", "PPA in force, share of the operating months (E: tenor)", "factor",
               e=f"=INDEX(Dash_PPA_Years,1,{cs})",
               t="=MAX(0,MIN(Months_In_Year,{self.E}*Months_In_Year-{m0})-MAX(0,-{m0}))/MAX(1,{opm})",
               fmts={"E": "int"})
        rv.row("ppx", "PPA price (E: lever, F: indexation)", "EUR/MWh", e=f"=INDEX(Dash_PPA_Price,1,{cs})",
               f=f"=INDEX(Dash_PPA_Index,1,{cs})", t="={self.E}*(1+{self.F})^({year}-YEAR(COD_Live))",
               fmts={"F": "pct"})
        rv.row("pxd", "PV price received, direct export", "EUR/MWh", t="={pos}+{ppa}*({ppx}-{pos})", bold=True)
        rv.row("pxs", "PV price received, stored energy", "EUR/MWh", t="={pvs}+{ppa}*({ppx}-{pvs})", bold=True)
        st, po, dc = (rv.row_of(x) for x in ("st4", "pvo4", "dcs4"))
        rv.row("upl", "BESS merchant uplift kept (block × PV-only factor; E: built × factor)", "EUR k/MW",
               t=f"={{pos.F}}*INDEX({{c}}${st}:{{c}}${st + 1},{{pos.E}})*(INDEX({{c}}${po}:{{c}}${po + 1},{{pos.E}})"
                 f"+INDEX({{c}}${dc}:{{c}}${dc + 1},{{pos.E}})*AFRR_Share)*{idx}")
        rv.row("tol", "Tolling in force, share of the year (E: basis is Tolling)", "factor",
               e=f'=N(INDEX(Dash_Basis,1,{cs})="Tolling")',
               t="={self.E}*MAX(0,MIN(Months_In_Year,Toll_Years*Months_In_Year-{t0})-MAX(0,-{t0}))/Months_In_Year",
               fmts={"E": "flag"})
        rv.row("bess", "BESS revenue received", "EUR k/MW", t="={upl}+{tol}*(Offer_Tolling_MET_3y-{upl})", bold=True)
        rows.append({k2[len(rv.scope):]: v for k2, v in rv.rows.items() if k2.startswith(rv.scope)})
        rv.scope = ""
        prev = r
    rv.end()
    book.rev = rows
    return rv


def _src(src):
    s = src.replace("Degraded batteries", "Aurora DB").replace("Common Assumptions", "Aurora CA")
    return s if len(s) <= 60 else s[:57] + "…"


def offers(book, rv):
    """Offer terms as stated (ASSUMPTIONS.md §2.1–2.2). Reference only; two named cells feed the model."""
    T = [  # label, unit, E.ON, MET/E2, MVM, source, fmt
        ("Offeror and date", "text", "E.ON, 09-Oct-26", "MET / E2, 20-Oct-26", "MVM ONEnergy, 22-Sep-26",
         "PDF p1; MVM e-mail", None),
        ("Validity", "text", "to 12-Dec-26", "–", "–", "E.ON p1", None),
        ("PV product", "text", "spot-indexed PaP", "fixed or DAM-indexed PaP", "fixed PaP", "E.ON p4; MET p3-4", None),
        ("Fixed price, 3 years (2028–2030)", "EUR/MWh", None, 43.96, None, "MET p3", "money2"),
        ("Fixed price, 4 years (2028–2031)", "EUR/MWh", None, 40.76, None, "MET p3", "money2"),
        ("Fixed price, 5 years (2028–2032)", "EUR/MWh", None, None, 45, "MVM e-mail", "money2"),
        ("Handling fee, share of DAM (indexed)", "%", 0.076, 0.05, None, "E.ON p4 §5.1; MET p4", "pct"),
        ("Handling fee floor", "EUR/MWh", 5, 5.75, None, "E.ON p4; MET p4", "money2"),
        ("Handling fee cap", "EUR/MWh", None, 10, None, "MET p4", "money2"),
        ("Start", "text", "COD (2027 Q4–2028 Q1)", "01-Jan-2028", "01-Jan-2028", "E.ON p1; MET p3; MVM", None),
        ("Tenor", "text", "max 3 years", "3 or 4 years", "5 years", "E.ON p1; MET p3; MVM", None),
        ("Volume", "text", "100 % of output", "100 % of output", "generated volume", "E.ON p4; MET p4", None),
        ("Balancing included (1 = yes)", "flag", 1, 1, 1, "E.ON p4; MET p3; MVM", "flag"),
        ("Balance-group fee", "EUR/month", 250, None, None, "E.ON p4", "money"),
        ("Guarantees of origin", "text", "option at 90 % of HUPX GO", "–", "–", "E.ON p4", None),
        ("Downregulation", "text", "curtailment at ≤ 0 prices", "10 €/MWh; volume 0 if DAM−fee ≤ 0", "–",
         "E.ON p4; MET p4", None),
        ("Settlement point (our reading)", "text", "POD", "PV meter", "unclear – ask", "ASSUMPTIONS §2.1", None),
        ("BESS product", "text", "flexibility agreement", "full tolling or floor + PSA", "none", "E.ON p3-5; MET p5-6",
         None),
        ("BESS configuration offered", "text", "50 MW / 100 MWh", "50 MW / 90 MWh", "–", "E.ON p1; MET p5", None),
        ("BESS energy offered", "MWh", 100, 90, None, "E.ON p1 (p3 prints 10 MWh); MET p5", "money"),
        ("Tolling fee, 3 years", "EUR k/MW/yr", None, 49.7, None, "MET p5", "money1"),
        ("Tolling fee, 4 years", "EUR k/MW/yr", None, 46.4, None, "MET p5", "money1"),
        ("Floor, 3 years (E.ON: 2027 / 2028 band)", "EUR k/MW/yr", "47 / 38", 45.6, None, "E.ON p5; MET p6", "money1"),
        ("Floor, 4 years", "EUR k/MW/yr", None, 41.9, None, "MET p6", "money1"),
        ("Share above the floor, to the owner", "text", "tiered fee 10–18 %", "50 %", "–", "E.ON p5; MET p6", None),
        ("Option B: owner share of the net profit", "%", 0.9, None, None, "E.ON p5", "pct"),
        ("Offeror's yield assumption", "text", "73,141 MWh on 61 MWp", "78,456 MWh on 63 MWp", "–",
         "E.ON p3; MET p3", None),
    ]
    for label, unit, e, f, g, src, fmt in T:
        cols = tuple(c for c, v in zip("EFG", (e, f, g)) if v is not None)
        r = rv.row(None, label, unit, e=e, f=f, g=g, inputs=cols, source=src, fmt=fmt)
        if label.startswith("Fixed price, 5 years"):
            book.define("Offer_PPA_MVM", f"Revenue!$G${r}")
        if label == "BESS energy offered":
            book.define("Offer_MET_MWh", f"Revenue!$F${r}")
        if label == "Tolling fee, 3 years":
            book.define("Offer_Tolling_MET_3y", f"Revenue!$F${r}")


# =========================================================================================== Calc
def calc(book):
    c = book.sheet("Calc", "Project Zalahaláp – calculations: three identical annual blocks, one per case (EUR k)",
                   timeline=True)
    last = book.last_col
    c.bar("Timeline and discounting (all cases)", e=["All cases"], h="Total")
    c.row("codcol", "COD year: column on the timeline (E), its year end (F)", "#",
          e=f"=MATCH(YEAR(COD_Live),$J$4:${last}$4,0)", f=f"=INDEX($J$5:${last}$5,{{self.E}})",
          fmts={"E": "int", "F": "date"})
    c.row("dfx", "Discount factor to the exit at the buyer target IRR", "factor",
          t="=(1+Case_Buyer_IRR)^((Exit_Date-{yend})/Days_In_Year)")
    c.row("dfc", "Discount factor to the COD year end at the buyer target IRR", "factor",
          t="=(1+Case_Buyer_IRR)^(({codcol.F}-{yend})/Days_In_Year)")
    c.row("bxd", "Buyer at RtB: flow dates (year end, not before the exit)", "date", t="=MAX({yend},Exit_Date)")
    c.row("dfa", "Discount factor from the first year end to the acquisition date (NPVs)", "factor",
          e="=1/(1+Discount_Rate)^(($J$5-Acq_Date)/Days_In_Year)")
    rev = book.rev
    idx = f"Revenue!{{c}}${book.idx}"

    def rows(i, blk):
        rr = rev[i - 1]
        ref = lambda k: f"Revenue!{{c}}${rr[k]}"              # noqa: E731
        lv = blk.lever
        c.sub("Timing")
        c.row("yrs", "Years from the COD year (E: first full operating year)", "years",
              e="=YEAR(COD_Live)+1", t="={year}-YEAR(COD_Live)", fmts={"E": "int"})
        c.row("cmon", "Construction months by year end (E: NTP to COD)", "months",
              e='=DATEDIF(NTP_Date,COD_Live,"m")', t="=({year}+1-YEAR(NTP_Date))*Months_In_Year-MONTH(NTP_Date)")
        c.row("cshare", "Construction share of the year (NTP to COD)", "%", h="=SUM({self.all})",
              t="=MAX(0,MIN({cmon},{cmon.E})-MAX({cmon}-Months_In_Year,0))/{cmon.E}", fmts={"H": "money2"})
        c.row("omon", "Months operated by year end (E: operating life)", "months", e="=Life*Months_In_Year",
              t="=({year}+1-YEAR(COD_Live))*Months_In_Year-MONTH(COD_Live)")
        c.row("oshare", "Operating share of the year (COD to the end of life)", "%", h="=SUM({self.all})",
              t="=MAX(0,MIN({omon},{omon.E})-MAX({omon}-Months_In_Year,0))/Months_In_Year", fmts={"H": "money2"})
        c.row("codf", "COD-year flag", "flag", t="=N({yrs}=0)")
        c.row("deg", "Degradation factor (from the first full year)", "factor", t="=(1-Degradation)^MAX({yrs}-1,0)")
        c.row("blk", "Aurora block for prices and volumes (E: code, F: position)", "text",
              e="=" + lv("Price_Block"), f="=MATCH({self.E},Aurora_Blocks,0)", fmts={"F": "int"})
        c.row("live", "BESS in service (E: repower year; 0 from it if not repowered)", "flag",
              e="=INDEX(Repower_Year,1,{blk.F})", t="=1-(1-Repower)*N({year}>={self.E})", fmts={"E": "int"})
        c.row("rep", "Repower-year flag (augmentation)", "flag", t="=Repower*N({year}={live.E})")
        c.sub("Energy")
        c.row("cap", "PV capacity, MWp (E); yield after the case change, MWh/MWp (F)", "MWp",
              e="=" + lv("Dash_MWp"), f=f"={lv('Dash_Yield')}*(1+Case_Yield)", fmts={"F": "money"})
        c.row("bess", "BESS power, MW (E); energy, MWh (F)", "MW", e=f"={lv('Dash_BESS_MW')}*{lv('Dash_BESS_On')}",
              f=f"={{self.E}}*{lv('Dash_Hours')}", fmts={"F": "money"})
        ra, rb = book.rv.row_of("chg4"), book.rv.row_of("chg2")
        c.row("chg", "PV energy charged per MW of BESS (Aurora block)", "MWh/MW",
              t=f"=INDEX(Revenue!{{c}}${ra}:{{c}}${rb},{{blk.F}})")
        ra, rb = book.rv.row_of("dis4"), book.rv.row_of("dis2")
        c.row("dmw", "BESS discharge per MW (Aurora block)", "MWh/MW", t=f"=INDEX(Revenue!{{c}}${ra}:{{c}}${rb},{{blk.F}})")
        ffy = "=SUMIF({year.all},{yrs.E},{self.all})"
        tot = "=SUM({self.all})"
        c.row("gen", "PV generation (E: first full year)", "MWh", e=ffy, h=tot,
              t="={cap.E}*{cap.F}*(1-Clipping)*PV_Avail*{deg}*{oshare}")
        c.row("sent", "PV energy sent to the BESS", "MWh", e=ffy, h=tot, t="=MIN({gen},{chg}*{bess.E}*{oshare}*{live})")
        c.row("direct", "PV energy exported directly", "MWh", e=ffy, h=tot, t="={gen}-{sent}")
        c.row("dis", "BESS discharge (no grid charging: at most the PV charged × RTE)", "MWh", e=ffy, h=tot,
              t="=MIN({dmw}*{bess.E}*{oshare}*{live},{sent}*BESS_RTE)")
        c.sub("Revenue")
        c.row("rpvd", "PV revenue, direct export", "EUR k", e=ffy, h=tot, t=f"={{direct}}*{ref('pxd')}/Thousand")
        c.row("rpvs", "PV revenue, energy sent to the BESS", "EUR k", e=ffy, h=tot, t=f"={{sent}}*{ref('pxs')}/Thousand")
        c.row("rgoo", "Guarantees of origin", "EUR k", e=ffy, h=tot, t=f"={{gen}}*GoO_Price*{idx}/Thousand")
        c.row("rbess", "BESS revenue", "EUR k", e=ffy, h=tot, t=f"={{bess.E}}*{ref('bess')}*{{oshare}}*{{live}}")
        c.row("rev", "Revenue", "EUR k", e=ffy, h=tot, t="=SUM({rpvd}:{rbess})", bold=True, top=True)
        c.sub("Opex")
        c.row("om", "O&M and MV cable O&M", "EUR k", e=ffy, h=tot, t=f"=-(OM_Cost+Cable_OM)*{{cap.E}}*{idx}*{{oshare}}")
        c.row("bal", "Balancing (merchant years only)", "EUR k", e=ffy, h=tot,
              t=f"=-{{gen}}*Balancing*{idx}*(1-{ref('ppa')})/Thousand")
        c.row("lease", "Land lease", "EUR k", e=ffy, h=tot, t=f"=-Lease*{idx}*{{oshare}}")
        c.row("bopex", "BESS fixed opex", "EUR k", e=ffy, h=tot, t=f"=-BESS_Opex*{{bess.E}}*{idx}*{{oshare}}*{{live}}")
        c.row("levy", "HIPA and innovation contribution (on revenue)", "EUR k", e=ffy, h=tot,
              t="=-{rev}*(HIPA+Innovation)")
        c.row("opex", "Opex", "EUR k", e=ffy, h=tot, t="=SUM({om}:{levy})", bold=True, top=True)
        c.row("ebitda", "EBITDA", "EUR k", e=ffy, h=tot, t="={rev}+{opex}", bold=True)
        c.sub("Capex")
        c.row("capex", "Construction capex, NTP to COD (E: total to COD, F: of which early works at NTP)", "EUR k",
              e="=" + lv("Capex_Total"), f="=" + lv("Capex_EW"), h=tot, t="=-({self.E}-{self.F})*{cshare}")
        c.row("ewntp", "Early works at NTP (E48/E49)", "EUR k", h=tot, t="=-{capex.F}*N({year}=YEAR(NTP_Date))")
        c.row("aug", "BESS augmentation at the repower year (indexed)", "EUR k", h=tot,
              t=f"=-{{bess.F}}*BESS_Capex*{idx}*{{rep}}")
        c.row("capt", "Capex", "EUR k", h=tot, t="=SUM({capex}:{aug})", bold=True, top=True)
        c.sub("Tax (SPV; unlevered = the only tax)")
        dep = "=-{self.E}*MAX(0,MIN({omon},{self.F})-MAX({omon}-Months_In_Year,0))/{self.F}"
        c.row("db", "Tax depreciation: BESS (E: base, F: months)", "EUR k", e="=" + lv("Capex_BESS"),
              f="=Dep_BESS_Years*Months_In_Year", h=tot, t=dep, fmts={"F": "int"})
        c.row("dpv", "Tax depreciation: PV, grid and development (E: base, F: months)", "EUR k",
              e=f"={{capex.E}}-{{db.E}}-{lv('Capex_Land')}+EW_Total+PreRtB_Spent", f="=Dep_Years*Months_In_Year",
              h=tot, t=dep, fmts={"F": "int"})
        c.row("da", "Tax depreciation: augmentation (BESS life from the repower year)", "EUR k", h=tot,
              t="={aug.H}/Dep_BESS_Years*N({year}>={live.E})*N({year}<{live.E}+Dep_BESS_Years)")
        c.row("taxb", "Tax base before losses", "EUR k", h=tot, t="={ebitda}+SUM({db}:{da})")
        c.row("lo", "Tax losses: opening (previous closing)", "EUR k", t=f"={{p}}{c.r + 4}")
        c.row("ladd", "Losses added", "EUR k", h=tot, t="=MAX(0,-{taxb})")
        c.row("luse", "Losses used (capped)", "EUR k", h=tot, t="=-MIN({lo},Loss_Cap*MAX(0,{taxb}))")
        c.row("lapse", "Losses lapsed (E: year, NTP + carry-forward; one date)", "EUR k", e="=YEAR(NTP_Date)+Loss_Years",
              h=tot, t="=-SUM({lo}:{luse})*N({year}={self.E})", fmts={"E": "int"})
        c.row("lc", "Tax losses: closing", "EUR k", t="=SUM({lo}:{lapse})")
        c.row("tax", "Corporate income tax (and Robin Hood when switched on)", "EUR k", h=tot, bold=True,
              t="=-CIT_Rate*(MAX(0,{taxb})+{luse})-Robin_Hood*RH_Rate*MAX(0,{taxb})")
        c.sub("Cash flows")
        c.row("pre", "Pre-COD outlays: development, deposits, guarantees (dated, by year)", "EUR k", h=tot,
              t=f"=SUMPRODUCT((YEAR(Dated_Dates)={{year}})*INDEX(Dated_Amounts,0,{blk.case}))+{{codf}}*Hold_At_COD")
        c.row("pcf", "Project cash flow before tax", "EUR k", h=tot, t="={ebitda}+{capt}+{pre}", bold=True)
        c.row("pcft", "Project cash flow after tax", "EUR k", h=tot, t="={pcf}+{tax}", bold=True)
        c.row("hold", "FEP, hold: cash flow after tax (E: PSH buy-out at the exit)", "EUR k", h=tot,
              e="=" + lv("PSH_Buyout"), t="={pcft}-{self.E}*N({year}=YEAR(Exit_Date))", bold=True)
        c.row("holdp", "FEP, hold: cash flow before tax", "EUR k", h=tot, t="={hold}-{tax}")
        c.row("codr", "FEP, sale at COD: cash flow before FEP tax (E: contributions to COD)", "EUR k", h=tot,
              e='=-{pre.H}-SUMIF({yrs.all},"<=0",{capt.all})',
              t=f"={{capt}}*N({{yrs}}<=0)+{{pre}}+{lv('COD_FEP_Total')}*{{codf}}")
        c.row("codt", "FEP, sale at COD: cash flow after FEP tax", "EUR k", h=tot,
              t=f"={{codr}}+{lv('COD_FEP_Tax')}*{{codf}}", bold=True)
        c.row("bx", "Buyer at RtB: flows after the exit, after tax (E: value at the exit at the target)", "EUR k",
              e="=SUMPRODUCT({self.all},{dfx.all})", h=tot, t="={ebitda}+{capt}+{tax}+{codf}*Hold_At_COD")
        c.row("bxp", "Buyer at RtB at our all-in price (E; paid at the exit, shown in 2025)", "EUR k",
              e=f"={lv('RtB_AllIn')}*{{cap.E}}", t="={bx}-{self.E}*N({year}=First_Year)")
        c.row("bxpp", "Buyer at RtB at our price, before tax", "EUR k", t="={bxp}-{tax}")
        c.row("bc", "Buyer at COD: flows from the COD year, after tax (E: value at the COD year end)", "EUR k",
              e="=SUMPRODUCT({self.all},{dfc.all})", h=tot, t="=({ebitda}+{tax}+{aug})*N({yrs}>=0)")
        c.row("bcp", "Buyer at COD at our price (E: price, paid at the COD year end)", "EUR k",
              e="=" + lv("COD_EV"), t="={bc}-{self.E}*{codf}")
        c.row("bcpp", "Buyer at COD at our price, before tax", "EUR k", t="={bcp}-{tax}*N({yrs}>=0)")
        c.sub("Returns (E: after tax, F: before tax)")
        dates = "$J$5:$" + last + "$5"
        x = lambda k: f"=XIRR({{{k}.all}},{dates})"          # noqa: E731
        c.row("r_proj", "Project IRR, unlevered", "%", e=x("pcft"), f=x("pcf"), bold=True)
        c.row("r_hold", "FEP IRR, hold", "%", e=x("hold"), f=x("holdp"), bold=True)
        v = lambda k: f"=XNPV(Discount_Rate,{{{k}.all}},{dates})*$E${c.row_of('dfa')}"   # noqa: E731
        c.row("v_hold", "FEP NPV at the discount rate, hold (to the acquisition date)", "EUR k", e=v("hold"),
              f=v("holdp"))
        c.row("r_cod", "FEP IRR, sale at COD", "%", e=x("codt"), f=x("codr"), bold=True)
        c.row("v_cod", "FEP NPV at the discount rate, sale at COD (to the acquisition date)", "EUR k", e=v("codt"),
              f=v("codr"))
        bxd = f"$J${c.row_of('bxd')}:${last}${c.row_of('bxd')}"
        c.row("r_bx", "Buyer IRR at RtB at our all-in price", "%", e=f"=XIRR({{bxp.all}},{bxd})",
              f=f"=XIRR({{bxpp.all}},{bxd})")
        cc = f"$E${c.row_of('codcol')}"
        rc, rp = c.row_of("bcp"), c.row_of("bcpp")
        xc = lambda r: f"=XIRR(INDEX($J${r}:${last}${r},{cc}):${last}${r},INDEX({dates},{cc}):${last}$5)"  # noqa: E731
        c.row("r_bc", "Buyer IRR at COD at our price", "%", e=xc(rc), f=xc(rp))

    blocks = Block(c, 3).build(lambda i: f"Case {i}: annual calculation (EUR k)", rows)
    for i, blk in enumerate(blocks, 1):
        c.ws.cell(blk["bar"], 3, f'=INDEX(Dash_Name,1,$E${blk["bar"]})&": annual calculation (EUR k)"')
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

    d.blank()
    d.bar("Capacity and energy, first full operating year", e=cl)
    pc("First full operating year", "year", lambda b: f"=Calc!$E${b['yrs']}")
    pc("PV generation", "GWh", lambda b: f"=Calc!$E${b['gen']}/Thousand")
    pc("PV energy sent to the BESS", "GWh", lambda b: f"=Calc!$E${b['sent']}/Thousand")
    pc("PV energy exported directly", "GWh", lambda b: f"=Calc!$E${b['direct']}/Thousand")
    pc("BESS discharge", "GWh", lambda b: f"=Calc!$E${b['dis']}/Thousand")
    pc("BESS full cycles", "#", lambda b: f"=IF(Calc!$F${b['bess']}>0,Calc!$E${b['dis']}/Calc!$F${b['bess']},0)",
       fmt="money1")
    pc("PV price received, all output (blended)", "EUR/MWh",
       lambda b: f"=(Calc!$E${b['rpvd']}+Calc!$E${b['rpvs']})*Thousand/Calc!$E${b['gen']}")
    pc("BESS revenue per MW", "EUR k/MW", lambda b: f"=IF(Calc!$E${b['bess']}>0,Calc!$E${b['rbess']}/Calc!$E${b['bess']},0)")

    d.blank()
    d.bar("Revenue and EBITDA (EUR k)", e=cl)
    for col, title in (("E", "First full operating year"), ("H", "Lifetime")):
        d.sub(title)
        pc("PV revenue (PPA, then merchant)", "EUR k", lambda b: f"=Calc!${col}${b['rpvd']}+Calc!${col}${b['rpvs']}")
        pc("BESS revenue", "EUR k", lambda b: f"=Calc!${col}${b['rbess']}")
        pc("Guarantees of origin", "EUR k", lambda b: f"=Calc!${col}${b['rgoo']}")
        pc("Revenue", "EUR k", lambda b: f"=Calc!${col}${b['rev']}", key=f"rev_{col}", bold=True, top=True)
        pc("Opex", "EUR k", lambda b: f"=Calc!${col}${b['opex']}")
        pc("EBITDA", "EUR k", lambda b: f"=Calc!${col}${b['ebitda']}", key=f"ebitda_{col}", bold=True)
        d.row(None, "EBITDA margin", "%", cases=f"={{ebitda_{col}}}/{{rev_{col}}}")

    d.blank()
    d.bar("Cash to RtB by nature (EUR k; all cases)", e=["All cases"])
    sm = lambda a, b: f"=-SUM($E${r(a)}:$E${r(b)})"        # noqa: E731
    d.row("n_spv", "SPV acquisition price (FEP)", "EUR k", e=sm("l_spv", "l_spv"))
    d.row("n_trans", "Transaction costs (FEP)", "EUR k", e=sm("l_trans", "l_trans"))
    d.row("n_dsa", "DSA loan to PSH", "EUR k", e=sm("l_dsa1", "l_dsa_last"))
    d.row("n_ew", "Early works, four lines (" + EW_NOTE + ")", "EUR k", e=sm("l_ew1", "l_ew_last"))
    book.define("EW_Total", d.ref("n_ew", "E"))
    d.row("n_pre", "Pre-RtB development", "EUR k", e=sm("l_pre1", "l_pre_last"))
    book.define("PreRtB_Spent", d.ref("n_pre", "E"))
    d.row("n_dep", "Cash deposits, net of releases before the exit", "EUR k", e=sm("l_dep1", "l_dep_last"))
    d.row("n_fee", "Bank guarantee fees to the exit", "EUR k", e=sm("l_fee1", "l_fee_last"))
    d.row("n_tot", "Total cash to RtB", "EUR k", e="=SUM({n_spv}:{n_fee})", bold=True, top=True)
    d.row("n_reimb", "of which reimbursed by the buyer (pre-RtB development, early works)", "EUR k",
          e="=Reimb_PreRtB*{n_pre}+Reimb_EW*{n_ew}", italic=True)
    d.row("n_ret", "of which returned at the exit (cash deposits outstanding)", "EUR k",
          e='=Reimb_Dep*SUMPRODUCT(Dep_HUF*(Dep_Rel=""))*Thousand/FX', italic=True)
    d.row("n_sunk", "of which sunk (recovered only through the consideration)", "EUR k",
          e="={n_tot}-{n_reimb}-{n_ret}", italic=True)
    d.row("n_buyer", "Buyer pays on top of the premium (reimbursed + returned)", "EUR k", e="={n_reimb}+{n_ret}")
    book.define("Buyer_Reimb", d.ref("n_buyer", "E"))
    d.row("n_exp", "FEP exposure at RtB (total less deposits returned)", "EUR k", e="={n_tot}-{n_ret}", bold=True)

    d.blank()
    d.bar("Capex to COD (EUR k, live sensitivity case)", e=cl)
    cx = "*(1+Case_Capex)"
    d.row("k_pv", "PV equipment and balance of plant", "EUR k", cases="=PV_Capex*{c}$7" + cx)
    d.row("k_cab", "MV cable", "EUR k", cases="=Cable_Capex" + cx)
    d.row("k_sub", "HV substation, bay and grid connection fee", "EUR k", cases="=(Substation+Grid_Fee)" + cx)
    d.row("k_land", "Land", "EUR k", cases="=Land_Capex" + cx)
    d.row("k_bess", "BESS (MWh × capex)", "EUR k", cases="=BESS_Capex*{c}$13" + cx)
    d.row("k_ew", "Early works at NTP (E48/E49; " + EW_NOTE + ")", "EUR k", cases="=EW_NTP" + cx)
    d.row("k_tot", "Total uses to COD (= equity, unlevered)", "EUR k", cases="=SUM({k_pv}:{k_ew})", bold=True,
          top=True)
    d.row(None, "Total uses per MWp", "EUR k/MWp", cases="={k_tot}/{c}$7")
    for nm, k in (("Capex_Land", "k_land"), ("Capex_BESS", "k_bess"), ("Capex_EW", "k_ew"), ("Capex_Total", "k_tot")):
        d.name_row(nm, k)

    d.blank()
    d.bar("Returns", e=cl)
    d.sub("Sale at RtB (dated flows to the exit; XIRR)")
    d.row("p_rtb", "RtB price used (premium)", "EUR k/MWp", cases=f"=IF({{c}}$20>0,{{c}}$20,{{c}}{r('imp')})", bold=True)
    pc("Implied premium at the buyer target IRR (floored at 0)", "EUR k/MWp",
       lambda b: f"=MAX(0,(Calc!$E${b['bx']}-Buyer_Reimb)/{{c}}$7)", key="imp")
    pc("Implied premium below zero (1 = floored)", "flag", lambda b: f"=N(Calc!$E${b['bx']}<Buyer_Reimb)", key="impneg")
    d.row("allin", "All-in price incl. reimbursements", "EUR k/MWp", cases="={p_rtb}+Buyer_Reimb/{c}$7")
    d.name_row("RtB_AllIn", "allin")
    pc("Buyer IRR at our price, unlevered after tax", "%", lambda b: f"=Calc!$E${b['r_bx']}")
    pc("Buyer IRR at our price, unlevered before tax", "%", lambda b: f"=Calc!$F${b['r_bx']}")
    d.row(None, "FEP exposure (cash to RtB less deposits returned)", "EUR k",
          cases=f"=-{{c}}{r('dl_out')}-$E${r('n_ret')}")
    d.row(None, "FEP proceeds at the exit", "EUR k", cases=f"={{c}}{r('jv_fep')}")
    d.row(None, "FEP profit, pre-tax", "EUR k", cases=f"={{c}}{r('dl_pre')}")
    d.row(None, "FEP profit, post-tax", "EUR k", cases=f"={{c}}{r('dl_post')}", bold=True)
    d.row(None, "MOIC, pre-tax", "x", cases=f"={{c}}{r('dl_moic')}")
    d.row(None, "IRR, pre-tax", "%", cases=f"={{c}}{r('dl_irr_pre')}")
    d.row(None, "IRR, post-tax", "%", cases=f"={{c}}{r('dl_irr')}", bold=True)
    d.row(None, "NPV at the discount rate, post-tax (to the acquisition date)", "EUR k", cases=f"={{c}}{r('dl_npv')}")
    d.row(None, "PSH profit", "EUR k", cases=f"={{c}}{r('psh_profit')}")
    d.sub("Sale at COD (EV at the COD year end; annual flows)")
    d.row("p_cod", "COD price used", "EUR k/MWp", cases=f"=IF({{c}}$21>0,{{c}}$21,{{c}}{r('impc')})", bold=True)
    pc("Implied COD price at the buyer target IRR (floored at 0)", "EUR k/MWp",
       lambda b: f"=MAX(0,Calc!$E${b['bc']}/{{c}}$7)", key="impc")
    d.row("ev", "Enterprise value used", "EUR k", cases="={p_cod}*{c}$7")
    d.name_row("COD_EV", "ev")
    pc("Buyer IRR at our price, unlevered after tax", "%", lambda b: f"=Calc!$E${b['r_bc']}")
    pc("Buyer IRR at our price, unlevered before tax", "%", lambda b: f"=Calc!$F${b['r_bc']}")
    pc("FEP contributions to COD", "EUR k", lambda b: f"=Calc!$E${b['codr']}")
    d.row(None, "FEP proceeds at COD", "EUR k", cases=f"={{c}}{r('c_fep')}")
    pc("FEP profit, pre-tax", "EUR k", lambda b: f"=Calc!$H${b['codr']}")
    pc("FEP profit, post-tax", "EUR k", lambda b: f"=Calc!$H${b['codt']}", bold=True)
    pc("MOIC, pre-tax", "x", lambda b: f'=SUMIF(Calc!$J${b["codr"]}:${book.last_col}${b["codr"]},">0")'
                                       f'/-SUMIF(Calc!$J${b["codr"]}:${book.last_col}${b["codr"]},"<0")')
    pc("IRR, pre-tax", "%", lambda b: f"=Calc!$F${b['r_cod']}")
    pc("IRR, post-tax", "%", lambda b: f"=Calc!$E${b['r_cod']}", bold=True)
    pc("NPV at the discount rate, post-tax (to the acquisition date)", "EUR k", lambda b: f"=Calc!$E${b['v_cod']}")
    d.row(None, "PSH profit", "EUR k", cases=f"={{c}}{r('c_pl')}+(DSA_Per_MWp-PSH_Devex)*{{c}}$7-PSH_Land")
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
    pc("NPV at the discount rate, after tax (to the acquisition date)", "EUR k", lambda b: f"=Calc!$E${b['v_hold']}")
    pc("Memo: project IRR, unlevered after tax (FEP's before the buy-out)", "%",
       lambda b: f"=Calc!$E${b['r_proj']}", italic=True)
    pc("Memo: project IRR, unlevered before tax", "%", lambda b: f"=Calc!$F${b['r_proj']}", italic=True)
    d.row(None, "PSH buy-out at RtB entitlement, pending JV confirmation", "EUR k", cases=f"={{c}}{r('jv_pl')}")
    d.row(None, "PSH profit", "EUR k", cases=f"={{c}}{r('jv_pl')}+(DSA_Per_MWp-PSH_Devex)*{{c}}$7-PSH_Land")

    d.blank()
    d.bar("JV split at the RtB exit (EUR k)", e=cl)
    d.row("jv_cons", "Consideration (premium × MWp)", "EUR k", cases="={p_rtb}*{c}$7")
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
    d.row("jv_fep", "FEP total at the exit (principal, return, SPV price, lot)", "EUR k",
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
        d.row(f"a{i}_cum", "Cumulative net cash flow", "EUR k", t=f"=SUM($J{{a{i}_net}}:{{a{i}_net}})")

    d.blank()
    d.bar("Detailed calculations – expand the group", e=cl)
    g0 = d.r
    d.blank()
    db = DatedBlock(d, "Loan_Rate")
    book.dated = db
    db.header("RtB route: FEP dated cash flows and the 12 % return (EUR k)")
    db.end("x", "Accrual end: sale closing; factor within the exit year (L)", "=Exit_Date", "L")
    rc = db.end("c", "Accrual end: COD, live case; factor within the COD year (M)",
                "=EOMONTH(COD_Date,Case_COD_Delay)", "M")
    book.define("COD_Live", f"Dashboard!$J${rc}")
    d.row("dsa_n", "DSA: monthly draws from the first draw to RtB", "months",
          e="=(YEAR(RtB_Date)-YEAR(DSA_First))*Months_In_Year+MONTH(RtB_Date)-MONTH(DSA_First)+1")
    d.row("pre_n", "Pre-RtB development: months", "months",
          e="=(YEAR(PreRtB_Last)-YEAR(PreRtB_First))*Months_In_Year+MONTH(PreRtB_Last)-MONTH(PreRtB_First)+1")
    d.sub("Outlays (all routes; the hold and COD routes take them by year)", extra={"J": "Date"})
    db.line("l_spv", "SPV acquisition price", "=-SPV_Price", "=Acq_Date", "=Ret_SPV")
    db.line("l_trans", "Transaction costs", "=-Transaction", "=Acq_Date", "=Ret_Trans")
    for k, rr in enumerate(book.dep_rows):
        db.line("l_dep1" if k == 0 else None, f"Cash deposit paid: {DEPOSITS[k][0]}",
                f"=-Inputs!$E${rr}*Thousand/FX", f"=Inputs!$F${rr}", "=Ret_Dep")
    for k, rr in enumerate(book.dep_rows):
        last_dep = k == len(book.dep_rows) - 1
        db.line("l_dep_last" if last_dep else None, f"Cash deposit released: {DEPOSITS[k][0]} (blank: at the exit)",
                f'=Inputs!$E${rr}*Thousand/FX*(Inputs!$G${rr}<>"")', f'=IF(Inputs!$G${rr}="",Inputs!$F${rr},'
                f'Inputs!$G${rr})', "=Ret_Dep")
    n_dsa = _months(DSA_FIRST, RTB)
    for k in range(n_dsa):
        key = "l_dsa1" if k == 0 else ("l_dsa_last" if k == n_dsa - 1 else None)
        date = "=DSA_First" if k == 0 else f"=EOMONTH(J{d.r - 1},1)"
        db.line(key, f"DSA draw {k + 1}", f"=-DSA_Per_MWp*{{c}}$7/$E${d.row_of('dsa_n')}*($J{{row}}<=RtB_Date)",
                date, "=Ret_DSA")
    for k, rr in enumerate(book.ew_rows):
        key = "l_ew1" if k == 0 else ("l_ew_last" if k == len(book.ew_rows) - 1 else None)
        huf = EARLY_WORKS[k][1] == "HUF m"
        db.line(key, f"Early works: {EARLY_WORKS[k][0]}", f"=-Inputs!$E${rr}" + ("*Thousand/FX" if huf else ""),
                f"=Inputs!$F${rr}", "=Ret_EW")
    n_pre = _months(PRE_FIRST, PRE_LAST)
    for k in range(n_pre):
        key = "l_pre1" if k == 0 else ("l_pre_last" if k == n_pre - 1 else None)
        date = "=PreRtB_First" if k == 0 else f"=EOMONTH(J{d.r - 1},1)"
        db.line(key, f"Pre-RtB development, month {k + 1}",
                f"=-PreRtB_Total/$E${d.row_of('pre_n')}*($J{{row}}<=PreRtB_Last)", date, "=Ret_PreRtB")
    for k, rr in enumerate(book.bg_rows):
        key = "l_fee1" if k == 0 else ("l_fee_last" if k == len(book.bg_rows) - 1 else None)
        db.line(key, f"Guarantee fee to release: {GUARANTEES[k][0]}",
                f"=-Inputs!$E${rr}*Thousand/FX*(BG_Base+BG_Margin)*($J{{row}}-Inputs!$F${rr})/Days_In_Year",
                f'=IF(Inputs!$G${rr}="",Exit_Date,Inputs!$G${rr})', "=Ret_Fees")
    a, z = d.row_of("l_spv"), d.row_of("l_fee_last")
    book.define("Dated_Amounts", f"Dashboard!$E${a}:$G${z}")
    book.define("Dated_Dates", f"Dashboard!$J${a}:$J${z}")
    d.sub("Receipts at the exit and FEP tax (per case)", extra={"J": "Date"})
    d.row("l_rec", "FEP receipts at the exit (JV split total)", "EUR k", cases=f"={{c}}{r('jv_fep')}",
          extra={"J": "=Exit_Date"}, fmts={"J": "date"})
    d.row("l_tax", "FEP tax: CIT on the 12 % return as interest (and the gain if taxed)", "EUR k",
          cases=f"=-CIT_Rate*(-{{c}}{r('jv_ret')}+Gain_Taxed*MAX(0,{{c}}{r('dl_pre')}+{{c}}{r('jv_ret')}))",
          extra={"J": "=Exit_Date"}, fmts={"J": "date"})
    rec, tx = d.row_of("l_rec"), d.row_of("l_tax")
    d.row("dl_pre", "FEP profit, pre-tax", "EUR k", cases=f"=SUM({{c}}{a}:{{c}}{rec})", bold=True, top=True)
    d.row("dl_post", "FEP profit, post-tax", "EUR k", cases=f"={{dl_pre}}+{{c}}{tx}", bold=True)
    d.row("dl_moic", "MOIC, pre-tax", "x", cases=f'=SUMIF({{c}}{a}:{{c}}{rec},">0")/-SUMIF({{c}}{a}:{{c}}{rec},"<0")')
    d.row("dl_irr_pre", "IRR, pre-tax", "%", cases=f"=XIRR({{c}}{a}:{{c}}{rec},$J{a}:$J{rec})")
    d.row("dl_irr", "IRR, post-tax", "%", cases=f"=XIRR({{c}}{a}:{{c}}{tx},$J{a}:$J{tx})", bold=True)
    d.row("dl_npv", "NPV at the discount rate, post-tax (to the acquisition date)", "EUR k",
          cases=f"=XNPV(Discount_Rate,{{c}}{a}:{{c}}{tx},$J{a}:$J{tx})")
    d.sub("Totals and claims due (per case)")
    d.row("dl_out", "Outlays to the exit, total (all routes)", "EUR k", cases=f"=SUM({{c}}{a}:{{c}}{z})")
    d.row("dl_dsa", "DSA loan principal (sum of the draws)", "EUR k",
          cases=f"=-SUM({{c}}{d.row_of('l_dsa1')}:{{c}}{d.row_of('l_dsa_last')})")
    db.claim("dl_retx", "12 % return due at the exit", "x", a, z)
    db.claim("dl_retc", "12 % return due at COD (sale at COD)", "c", a, z)
    d.sub("Hold and COD routes: at COD (all cases)")
    d.row("h_dep", "Cash deposits outstanding at the exit, returned at COD", "EUR k",
          e='=SUMPRODUCT(Dep_HUF*(Dep_Rel=""))*Thousand/FX')
    d.row("h_bg", "Bank guarantees outstanding at the exit", "EUR k", e='=SUMPRODUCT(BG_HUF*(BG_Rel=""))*Thousand/FX')
    d.row("h_fee", "Guarantee fees from the exit to COD", "EUR k",
          e="=-{h_bg}*(BG_Base+BG_Margin)*(COD_Live-Exit_Date)/Days_In_Year")
    d.row("h_net", "Net cash at COD (hold and COD routes)", "EUR k", e="={h_dep}+{h_fee}", bold=True, top=True)
    book.define("Hold_At_COD", d.ref("h_net", "E"))
    d.sub("PSH, sale at RtB (per case; DSA received over the DSA draws)", extra={"J": "Date"})
    d.row("psh_land", "PSH land apport (in kind)", "EUR k", cases="=-PSH_Land", extra={"J": "=Acq_Date"},
          fmts={"J": "date"})
    d.row("psh_dsa", "DSA received less PSH's own development costs", "EUR k", cases="=(DSA_Per_MWp-PSH_Devex)*{c}$7")
    d.row("psh_lot", "PSH lot at the exit", "EUR k", cases=f"={{c}}{r('jv_pl')}", extra={"J": "=Exit_Date"},
          fmts={"J": "date"})
    d.row("psh_profit", "PSH profit, sale at RtB", "EUR k", cases="=SUM({psh_land}:{psh_lot})", bold=True, top=True)

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
    pc("FEP tax at COD: CIT on the 12 % return (and the gain if taxed)", "EUR k",
       lambda b: f"=-CIT_Rate*(-{{c_ret}}+Gain_Taxed*MAX(0,Calc!$H${b['codr']}+{{c_ret}}))", key="c_tax")
    d.name_row("COD_FEP_Tax", "c_tax")
    d.group(g0, d.r - 1)
    d.end()
    return d


# =========================================================================================== Checks
def checks(book):
    k = book.sheet("Checks", "Project Zalahaláp – checks (0 = OK); the master check is mirrored at Dashboard H3")
    k.bar("Checks (0 = OK)", e=_case_links(book, k), h="Total")
    B, D_, rv = book.blocks, book.sheets["Dashboard"], book.rev
    last = book.last_col
    dr = D_.row_of
    first = k.r

    def ck(label, fn, shared=False):
        """fn(i, block, col) -> formula; shared checks are written in E only."""
        cells = {}
        for i, b in enumerate(B):
            if shared and i:
                break
            cells["efg"[i]] = fn(i, b, "EFG"[i])
        k.row(None, label, "flag", **cells, h="=SUM({self.cases})", fmt="flag")

    ck("Lever: BESS duration has an Aurora block when built (2 h or 4 h)",
       lambda i, b, C: f"=Dashboard!{C}$10*ISNA(MATCH(Dashboard!{C}$14,Aurora_Blocks,0))")
    ck("Lever: BESS revenue basis is Merchant or Tolling",
       lambda i, b, C: f'=N(AND(Dashboard!{C}$19<>"Merchant",Dashboard!{C}$19<>"Tolling"))')
    ck("Lever: PPA tenor within the operating life", lambda i, b, C: f"=N(Dashboard!{C}$18>Life)")
    ck("Timing: COD year not before Aurora's first year", lambda i, b, C: "=N(YEAR(COD_Live)<Aurora_First)",
       shared=True)
    ck("Timing: operating life inside the timeline", lambda i, b, C: f"=N(ABS(Calc!$H${b['oshare']}-Life)>Tolerance)")
    ck("Energy: sent to the BESS + exported directly = generation",
       lambda i, b, C: f"=N(ABS(Calc!$H${b['sent']}+Calc!$H${b['direct']}-Calc!$H${b['gen']})>Tolerance)")
    ck("Energy: discharge ≤ PV charged × RTE, every year",
       lambda i, b, C: f"=SUMPRODUCT(--(Calc!J{b['dis']}:{last}{b['dis']}>Calc!J{b['sent']}:{last}{b['sent']}"
                       f"*BESS_RTE+Tolerance))")
    ck("Grid: direct export ≤ MVA × 8,760 h × operating share (information)",
       lambda i, b, C: f"=SUMPRODUCT(--(Calc!J{b['direct']}:{last}{b['direct']}>Dashboard!{C}$8*Hours_In_Year"
                       f"*Calc!J{b['oshare']}:{last}{b['oshare']}))")
    ck("BESS rows are zero when no BESS is built",
       lambda i, b, C: f"=(1-Dashboard!{C}$10)*N(ABS(Calc!$H${b['sent']})+ABS(Calc!$H${b['dis']})"
                       f"+ABS(Calc!$H${b['rbess']})>Tolerance)")
    ck("Revenue: PV price received = PPA price inside the tenor",
       lambda i, b, C: f"=SUMPRODUCT((Revenue!J{rv[i]['ppa']}:{last}{rv[i]['ppa']}=1)"
                       f"*(ABS(Revenue!J{rv[i]['pxd']}:{last}{rv[i]['pxd']}-Revenue!J{rv[i]['ppx']}:{last}"
                       f"{rv[i]['ppx']})>Tolerance))")
    ck("Capex: construction spread in full (timeline = total uses)",
       lambda i, b, C: f"=N(ABS(Calc!$H${b['capex']}+Calc!$H${b['ewntp']}+Calc!$E${b['capex']})>Tolerance)")
    ck("Dated outlays all mapped to the annual grid",
       lambda i, b, C: f"=N(ABS(Calc!$H${b['pre']}-Hold_At_COD-Dashboard!{C}{dr('dl_out')})>Tolerance)")
    ck("Dated: DSA draws add up to the DSA loan",
       lambda i, b, C: f"=N(ABS(Dashboard!{C}{dr('dl_dsa')}-DSA_Per_MWp*Dashboard!{C}$7)>Tolerance)")
    ck("Dated: pre-RtB development months add up to the total",
       lambda i, b, C: "=N(ABS(PreRtB_Spent-PreRtB_Total)>Tolerance)", shared=True)
    ck("Dated: lines earning the return are paid before the exit year",
       lambda i, b, C: "=" + book.dated.check_before_end("x"), shared=True)
    ck("Dated: deposit and guarantee releases not after the exit",
       lambda i, b, C: '=N(SUMPRODUCT((Dep_Rel<>"")*(Dep_Rel>Exit_Date))+SUMPRODUCT((BG_Rel<>"")'
                       '*(BG_Rel>Exit_Date))>0)', shared=True)
    ck("JV at RtB: FEP total + PSH lot + sale costs = consideration + reimbursements",
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
    ck("Returns: every IRR on Calc is a number (project, hold, COD, buyers)",
       lambda i, b, C: f"=N(COUNT(Calc!$E${b['r_proj']}:$F${b['r_bc']})<COUNTA(Calc!$E${b['r_proj']}:$F${b['r_bc']}))")
    ck("Returns: the RtB-route XIRRs are numbers",
       lambda i, b, C: f"=N(COUNT(Dashboard!{C}{dr('dl_irr_pre')}:{C}{dr('dl_irr')})<COUNTA(Dashboard!{C}"
                       f"{dr('dl_irr_pre')}:{C}{dr('dl_irr')}))")
    ck("Summary: lifetime net cash flow = investment + EBITDA + tax",
       lambda i, b, C: f"=N(ABS(Dashboard!$H${dr(f'a{i}_net')}-Dashboard!$H${dr(f'a{i}_inv')}"
                       f"-Dashboard!$H${dr(f'a{i}_ebitda')}-Calc!$H${b['tax']})>Tolerance)")
    lastc = k.r - 1
    k.blank()
    r = k.row(None, "Master check (sum of the check rows)", "flag", h=f"=SUM(H{first}:H{lastc})", bold=True,
              fmt="check")
    check_format(k.ws, f"H{r}")
    book.define("Master_Check", f"Checks!$H${r}")
    k.blank()
    k.sub("Information (economic outcomes, not errors: outside the master check)")
    ck("Implied RtB premium floored at 0 (buyer value below the reimbursements)",
       lambda i, b, C: f"=Dashboard!{C}{dr('impneg')}")
    ck("Tolling offer applied to a battery other than the one offered (fee not scaled)",
       lambda i, b, C: f"=Dashboard!{C}{D_.row_of('toll_cfg')}")
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
    calc(book)
    dashboard(book, d, R or {})
    if R is None:
        return model(dict(d.rows))
    checks(book)
    outputs(book, d)
    return book


HEADLINES = [  # Out_<name>_<case> for the audit: (name, Dashboard label, occurrence)
    ("Hold_IRR", "IRR, after tax", 0), ("Hold_IRR_Pre", "IRR, before tax", 0), ("Hold_MOIC", "MOIC, after tax", 0),
    ("RtB_IRR", "IRR, post-tax", 0), ("RtB_Profit", "FEP profit, post-tax", 0), ("RtB_MOIC", "MOIC, pre-tax", 0),
    ("COD_IRR", "IRR, post-tax", 1), ("COD_Profit", "FEP profit, post-tax", 1),
    ("Implied_RtB", "Implied premium at the buyer target IRR (floored at 0)", 0),
    ("Implied_COD", "Implied COD price at the buyer target IRR (floored at 0)", 0),
    ("Buyer_RtB_IRR", "Buyer IRR at our price, unlevered after tax", 0)]


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
