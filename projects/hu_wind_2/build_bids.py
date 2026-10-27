"""HU wind bids: Bora (one 50 MVA low bid), Iváncsa, Dunaújváros – RtB exit, COD exit, hold to maturity (SPEC §23).

Two sheets, no macros (columns Bora | Iváncsa | Dunaújváros | Total in E:H; dates per project in J:L; years in J:AS):
  Dashboard  read top-down: key inputs and the case table (Base | Low | High, Live column) → bid points and capacity →
             deposit summary in the user's layout (HUF with dates, EUR k with dates, deductions shown, DEVEX line) →
             capex and devex breakdowns → the three routes (totals) → annual cash flow per project → the detailed
             calculations (claims, waterfall, dated FEP flows, timelines, COD tax, checks) in a collapsed row group.
  Inputs     rates, dates, fees, devex by category, capex, opex, claim terms; band tables in a collapsed row group.
Writes out/HU_Wind_bids.xlsx and out/rows_bids.json (row map for check.py).

    python3 projects/hu_wind_2/build_bids.py
"""
import json
import sys
from datetime import datetime as dt
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import Outline

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(ROOT))
from fepm.xl.styles import (BLACK, BLUE, CHECK_RED, DARK, FEP_DARK_GREEN, FMT, FONT, GREY, INPUT_FILL,  # noqa: E402
                            OK_GREEN, SUB_FILL, WHITE)

OUT = HERE / "out" / "HU_Wind_bids.xlsx"
ROWS = HERE / "out" / "rows_bids.json"
BID_CASH = ROOT / "reference" / "HU_Wind_bid_deposits_fees_20261027.xlsx"     # band tables (user's sheet)

P = ["E", "F", "G"]                                   # Bora, Iváncsa, Dunaújváros (same columns on both sheets)
NAMES = ["Bora", "Iváncsa", "Dunaújváros"]
DCOL = {"E": "J", "F": "J", "G": "J"}                 # payment date (one date column, common key dates)
TL = [L(10 + i) for i in range(36)]                   # J:AS, 2026-2061
NUM, NUM1, PCT, MULT, DATE, INT = FMT["money"], FMT["money1"], FMT["pct"], FMT["mult"], FMT["date"], "0"
NUM2, FAC, FLAG = FMT["money2"], FMT["factor"], FMT["flag"]
CHECK_FMT = '"CHECK" 0;"CHECK" -0;"OK"'
THIN = Side(style="thin", color="BFBFBF")
TOP = Border(top=Side(style="thin", color=BLACK))
HEADS = {"E": "Bora", "F": "Iváncsa", "G": "Dunaújváros", "H": "Total"}
DATES = {**HEADS, "J": "Date"}

ws = None
R: dict = {}          # Dashboard rows
I: dict = {}          # Inputs rows


def font(color=BLACK, bold=False, italic=False):
    return Font(name=FONT, size=10, color=color, bold=bold, italic=italic)


def put(ref, value, kind="calc", fmt=None, bold=False, top=False, sheet=None):
    c = (sheet or ws)[ref]
    c.value = value
    if kind == "input":
        c.font, c.fill = font(BLUE), PatternFill("solid", fgColor=INPUT_FILL)
        c.border = Border(THIN, THIN, THIN, THIN)
    elif kind == "note":
        c.font = font(GREY, italic=True)
    else:
        c.font = font(bold=bold)
        if top:
            c.border = TOP
    if fmt:
        c.number_format = fmt
    if kind != "note":
        c.alignment = Alignment(horizontal="right")
    return c


def label(r, text, unit=None, bold=False, sheet=None):
    s = sheet or ws
    s[f"C{r}"].value, s[f"C{r}"].font = text, font(bold=bold)
    if unit:
        s[f"D{r}"].value, s[f"D{r}"].font = unit, font(GREY, italic=True)


def bar(r, title, heads, extra=(), sheet=None):
    s = sheet or ws
    for col in ["B", "C", "D", "E", "F", "G", "H", *extra]:
        s[f"{col}{r}"].font = font(WHITE, bold=True)
        s[f"{col}{r}"].fill = PatternFill("solid", fgColor=FEP_DARK_GREEN)
        s[f"{col}{r}"].alignment = Alignment(horizontal=None if col in ("B", "C", "D") else "right")
    s[f"B{r}"].value, s[f"C{r}"].value = "x", title
    for col, h in heads.items():
        s[f"{col}{r}"].value = h


def sub(r, title, heads=None, extra=(), sheet=None):
    s = sheet or ws
    for col in ["C", "D", "E", "F", "G", "H", *extra]:
        s[f"{col}{r}"].font = font(DARK, bold=True)
        s[f"{col}{r}"].fill = PatternFill("solid", fgColor=SUB_FILL)
        s[f"{col}{r}"].alignment = Alignment(horizontal=None if col in ("C", "D") else "right")
    s[f"C{r}"].value = title
    for col, h in (heads or {}).items():
        s[f"{col}{r}"].value = h


def fx(t, **kw):
    return t.format(**{**R, **kw})


def name(wb, n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def rows3(r, key, text, unit, fmt, f, total=True, bold=False, top=False, date=None):
    """One row per project column (E:G) from a template ({c} = Dashboard / Inputs column), total in H, and the
    payment date in J (`date`: a date name on Inputs)."""
    R[key] = r
    label(r, text, unit, bold=bold)
    for c in P:
        put(f"{c}{r}", fx(f, c=c, d=DCOL[c]), fmt=fmt, bold=bold, top=top)
    if date:
        put(f"J{r}", "=" + date, fmt=DATE)
    if total:
        put(f"H{r}", f"=SUM(E{r}:G{r})", fmt=fmt, bold=bold, top=top)
    return r + 1


# ---------------------------------------------------------------------------------------------------------------------
# Inputs sheet.  Common inputs: (name, label, unit, fmt, value, source), column E, workbook name.
# Per-project inputs: (key, label, unit, fmt, [Bora, Iváncsa, Dunaújváros], source).
DEVEX = [("dv_grid", "Devex: grid connection plan and agreement", [40.47, 0, 0]),
         ("dv_land", "Devex: land (reservation, purchase, contracting)", [150.66, 0, 0]),
         ("dv_perm", "Devex: permits, zoning, cable rights", [333.93, 0, 0]),
         ("dv_stud", "Devex: studies (wind, environment, ground)", [151.36, 0, 0]),
         ("dv_legal", "Devex: legal, financial and other", [37.5, 0, 0]),
         ("dv_other", "Devex: not itemised", [0, 909.09, 909.09]),
         ("dv_cont", "Devex: contingency", [70.55, 90.91, 90.91])]
INPUTS = [
    ("Tender and grid (all projects; value in column E)", None),
    ("FX", "FX rate", "HUF/EUR", NUM, 365, "Bora Inputs E11; Bid_Cash E45"),
    ("Bench_Wind", "Security base: benchmark investment, wind", "HUF/MW", NUM, 406e6, "Bora Inputs E75; Bid_Cash E46"),
    ("Bench_BESS", "Security base: benchmark investment, BESS", "HUF/MWh", NUM, 200e6, "Bora Inputs E89; Bid_Cash E47"),
    ("Bid_Sec_Rate", "Bid security, share of base", "%", PCT, 0.005, "Bora Inputs E76; Bid_Cash E49"),
    ("PB_Rate", "Performance security, minimum share of base", "%", PCT, 0.025, "Bora Inputs E77; Bid_Cash E50"),
    ("Part_Fee", "Participation fee per bid", "HUF", NUM, 3e6, "Bora Inputs E81; Bid_Cash E48"),
    ("TSO_Rate", "TSO connection-request security", "HUF/MVA", NUM, 4.5e6, "Bora Inputs E83; Bid_Cash E51"),
    ("Per project: fees, devex, dates", "P"),
    ("ind", "Indirect connection fee", "HUF/MVA", NUM, [73071306, 19826680, 19826680], "Bora Inputs G290; Bid_Cash E52"),
    ("dir", "Direct connection fee per connection point", "HUF", NUM, [508345804, 460e6, 460e6],
     "Bora Inputs G291; Bid_Cash E53"),
    ("share_pt", "Bids sharing the connection point", "count", INT, [2, 1, 1], "Bora: 2 bids per point (SPEC §22.2)"),
    ("sdd", "Signing fee and DD", "EUR k", NUM, [100, 150, 150], "Bora E24+E25: 400 × 50/200; Bid_Cash E55"),
    *[(k, t, "EUR k", NUM1, v, "Bora Devex_FEP × 50/100 MVA; Bid_Cash E56 (10 % cont.)") for k, t, v in DEVEX],
    ("trafo", "EU transformers ≥ 80 % (ÖSZ-9)", "flag", INT, [0, 0, 0], "Bid_Cash E16; SPEC §23.1"),
    ("Key dates (all projects; value in column E)", None),
    ("Date_Sign", "Signing fee and DD paid", "date", DATE, dt(2026, 7, 31), "Bid_Cash E34"),
    ("Date_Bid", "Bid submission: fee paid, bid security posted", "date", DATE, dt(2026, 10, 30), "Bid_Cash E35"),
    ("Date_PB", "Performance security posted (bid security credited)", "date", DATE, dt(2027, 1, 15), "Bid_Cash E36"),
    ("Date_TSO", "TSO security paid", "date", DATE, dt(2027, 1, 15), "Bid_Cash E37"),
    ("Date_Devex", "Devex average payment date (spent continuously)", "date", DATE, dt(2027, 12, 31),
     "Bid_Cash H77 'continuous'; SPEC §21.1"),
    ("Date_HC", "Connection agreement (HCSSZ): fees paid", "date", DATE, dt(2028, 10, 1), "Bid_Cash E39"),
    ("Date_NTP", "Notice to proceed", "date", DATE, dt(2029, 6, 30), "Bid_Cash E40"),
    ("Date_COD", "COD", "date", DATE, dt(2031, 6, 30), "Bid_Cash E41"),
    ("Date_Release", "Performance security released", "date", DATE, dt(2033, 3, 31), "Bid_Cash E42 (estimate)"),
    ("RtB exit terms (all projects; value in column E)", None),
    ("Rate_Cash", "Claim rate, cash lines (incl. cash-posted securities)", "%", PCT, 0.12, "Devex file V18; SPEC §21.1"),
    ("Rate_Devex", "Claim rate, devex", "%", PCT, 0.20, "Devex file V19"),
    ("Devex_Floor", "Devex minimum return multiple", "x", MULT, 2.0, "Devex file V20"),
    ("Freeze_Date", "Interest freeze date", "date", DATE, dt(2029, 6, 30), "Devex file V24"),
    ("Sale_Date", "Sale date (RtB exit)", "date", DATE, dt(2029, 9, 30), "Bora Inputs E58; SPEC §21.1"),
    ("Sale_Costs", "Sale costs per project", "EUR k", NUM, 250, "SPEC §21.1"),
    ("DSA_Rate", "Developer success fee (on committed MW, signing to sale, 30/360)", "EUR k/MW/yr", NUM1, 2.5, "Bora Inputs E36"),
    ("DSA_Cap", "Developer success fee cap", "EUR k/MW", NUM1, 10, "Bora Inputs E37"),
    ("CIT_Rate", "Corporate income tax on exit gains", "%", PCT, 0.09, "SPEC §22.2 (exit gains only)"),
    ("Capex (per MW of wind built unless stated; turbines: case table on the Dashboard)", None),
    ("Capex_Install", "Turbine installation and erection", "EUR k/MW", NUM, 80, "Bora Inputs G163"),
    ("Capex_Fdn", "Foundations", "EUR k/MW", NUM, 90, "Bora Inputs G164"),
    ("Capex_Civil", "Balance of plant – civil", "EUR k/MW", NUM, 70, "Bora Inputs G165"),
    ("Capex_Elec", "Balance of plant – electrical", "EUR k/MW", NUM, 60, "Bora Inputs G166"),
    ("Capex_Owner", "Owner's costs, engineering, PM, construction insurance", "EUR k/MW", NUM, 40, "Bora Inputs G169"),
    ("Capex_Land", "Land purchase and easements", "EUR k/MW", NUM, 15, "Bora Inputs G170"),
    ("Subs_Lump", "On-site substation and 132 kV bay (lump sum)", "EUR k", NUM, 12000, "Bora Inputs G167"),
    ("Lump_Basis", "Source basis of the lump sums", "MVA", NUM, 200, "Bora Inputs E305"),
    ("Capex_Cont", "Capex contingency (not on BESS)", "%", PCT, 0.05, "Bora Inputs G173"),
    ("BESS_Capex", "BESS capex", "EUR k/MWh", NUM, 100, "CLAUDE.md standing assumption, 15-Oct-26"),
    ("Generation (hours and PPA price: case table on the Dashboard)", None),
    ("Curtailment", "Curtailment", "%", PCT, 0.02, "Bora Inputs G183"),
    ("Clipping", "Clipping", "%", PCT, 0.04, "Bora Dash_Clip_Override (user, 26-Oct-26)"),
    ("Degradation", "Degradation", "% p.a.", PCT, 0.003, "Bora Inputs G184"),
    ("PPA_Esc", "PPA escalation", "% p.a.", PCT, 0.01, "Bora Inputs G194"),
    ("Opex (COD-year prices; O&M contract: case table on the Dashboard)", None),
    ("Opex_BoP", "Balance of plant maintenance", "EUR k/MW/yr", NUM1, 3, "Bora Inputs G204"),
    ("Opex_Lease", "Land lease", "EUR k/MW/yr", NUM1, 3, "Bora Inputs G205"),
    ("Opex_Ins", "Operational insurance", "EUR k/MW/yr", NUM1, 2.5, "Bora Inputs G206"),
    ("Opex_Mgmt", "Asset and commercial management", "EUR k/MW/yr", NUM1, 1.5, "Bora Inputs G207"),
    ("Opex_PropTax", "Property tax", "EUR k/MW/yr", NUM1, 3, "Bora Inputs G208"),
    ("Opex_Admin", "Administration, audit, compliance (at the source basis)", "EUR k/yr", NUM, 150, "Bora Inputs G209"),
    ("Balancing", "Balancing cost", "EUR/MWh", NUM1, 4, "Bora Inputs G196"),
    ("Grid_Charges", "Grid use-of-system charges", "EUR/MWh", NUM1, 1, "Bora Inputs G197"),
    ("MEKH_Fee", "MEKH supervisory fee", "% of rev.", PCT, 0.001, "Bora Inputs G198"),
    ("HIPA_Rate", "HIPA local business tax", "% of rev.", PCT, 0.02, "Bora Inputs G212"),
    ("Opex_Cont", "Opex contingency (not on community, HIPA)", "%", PCT, 0.05, "Bora Inputs G210"),
    ("Opex_Esc", "Opex escalation", "% p.a.", PCT, 0.025, "Bora Inputs G211"),
    ("Valuation, timeline and constants", None),
    ("Life", "Operating life from COD", "years", INT, 30, "Bora Inputs G157"),
    ("Discount_Rate", "Discount rate (buyer's pre-tax return at COD)", "%", PCT, 0.08, "Bora Inputs G226"),
    ("First_Year", "First year of the annual timeline", "year", INT, 2026, "Year of the first signing"),
    ("Thousand", "Thousand", "x", NUM, 1000, "Constant"),
    ("Days_In_Year", "Days in a year (act/365)", "days", INT, 365, "Constant"),
    ("Months_In_Year", "Months in a year", "months", INT, 12, "Constant"),
    ("Tolerance", "Check tolerance", "EUR k", "0.000", 0.001, "Constant"),
]


def build_inputs(wb):
    s = wb.create_sheet("Inputs")
    s.sheet_view.showGridLines = False
    s.sheet_properties.outlinePr = Outline(summaryBelow=False)
    for col, w in zip("ABCDEFGH", (2, 2, 54, 11, 13, 13, 13, 50)):
        s.column_dimensions[col].width = w
    s["C2"].value, s["C2"].font = "HU wind bids – other inputs (blue) with sources", font(bold=True)
    r = 4
    bar(r, "Inputs", {"E": "Bora", "F": "Iváncsa", "G": "Dunaújváros", "H": "Source"}, sheet=s)
    s["H4"].alignment = Alignment(horizontal=None)
    r += 1
    for item in INPUTS:
        if item[1] in (None, "P"):
            sub(r, item[0], sheet=s)
            r += 1
            continue
        key, text, unit, fmt, v, note = item
        I[key] = r
        label(r, text, unit, sheet=s)
        if isinstance(v, list):
            for c, x in zip(P, v):
                put(f"{c}{r}", x, "input", fmt, sheet=s)
        else:
            put(f"E{r}", v, "input", fmt, sheet=s)
            name(wb, key, f"Inputs!$E${r}")
        put(f"H{r}", note, "note", sheet=s)
        r += 1
    name(wb, "Opex_Per_MW", f"Inputs!$E${I['Opex_BoP']}:$E${I['Opex_PropTax']}")
    # Band tables: the user's Bid_Cash J3:V28 (values and names), Kiírás Annex 4; collapsed row group.
    r += 1
    bar(r, "Band tables for the points (ÖSZ, Kiírás Annex 4) – expand the group", {"E": "Values →"}, sheet=s)
    r += 1
    b = openpyxl.load_workbook(BID_CASH)["Bid_Cash"]
    for rr in range(3, 29):
        n = b[f"J{rr}"].value
        vals = [b.cell(rr, c).value for c in range(11, 23) if b.cell(rr, c).value is not None]
        label(r, n, None, sheet=s)
        for j, x in enumerate(vals):
            put(f"{L(5 + j)}{r}", x, "input", "General", sheet=s)
        name(wb, n, f"Inputs!$E${r}:${L(4 + len(vals))}${r}" if len(vals) > 1 else f"Inputs!$E${r}")
        s.row_dimensions[r].outlineLevel, s.row_dimensions[r].hidden = 1, True
        r += 1
    r += 1
    s[f"C{r}"].value, s[f"C{r}"].font = "***END***", font(GREY, bold=True)
    s.freeze_panes = "E5"


# ---------------------------------------------------------------------------------------------------------------------
# Dashboard (a): key inputs (blue, Dash_* names) and the case table.
KEY = [("Dash_MVA", "Net connection capacity", "MVA", NUM, [50, 50, 50]),
       ("Dash_Overbuild", "Overbuild (ÖSZ-3)", "%", PCT, [0.20, 0.40, 0.40]),
       ("Dash_BESS_Share", "BESS power, share of MVA (ÖSZ-4)", "%", PCT, [0.10, 0.20, 0.20]),
       ("Dash_Hours", "BESS duration", "h", NUM, [4, 4, 4]),
       ("Dash_Shortfall", "Wind built below the bid", "%", PCT, [0.10, 0.10, 0.10]),
       ("Dash_Municipal", "Municipal one-off (ÖSZ-1), at COD", "EUR k/MVA", NUM, [20, 50, 50]),
       ("Dash_Community", "Community contribution (ÖSZ-2)", "EUR/MWh", NUM1, [0, 3, 3]),
       ("Dash_Partial", "Partial award accepted (ÖSZ-5)", "flag", INT, [1, 0, 0]),
       ("Dash_PB_Mult", "Performance security multiple (ÖSZ-6)", "x", MULT, [1, 1, 1]),
       ("Dash_Brownfield", "Brownfield site (ÖSZ-7)", "flag", INT, [0, 0, 0]),
       ("Dash_Shared", "Shared connection (ÖSZ-8)", "flag", INT, [1, 0, 0]),
       ("Dash_Turbine", "Turbine route (EU or CN; EU scores ÖSZ-9)", "text", "General", ["CN", "CN", "CN"]),
       ("Dash_Premium", "RtB premium", "EUR k/MVA", NUM, [150, 150, 150]),
       ("Dash_FEP_Share", "FEP share of distributable proceeds", "%", PCT, [0.70, 1.0, 1.0]),
       ("Dash_DSA", "Developer success fee (DSA) on", "flag", INT, [1, 0, 0])]
CASES = [("Case_EU", "Turbine supply incl. transport, EU route", "EUR k/MW", NUM, [850, 950, 800],
          "Bora Dashboard H8 / Inputs G162"),
         ("Case_CN", "Turbine supply incl. transport, CN route", "EUR k/MW", NUM, [500, 500, 500], "Bora Inputs H162"),
         ("Case_PPA", "PPA price in the COD year", "EUR/MWh", NUM1, [90, 90, 100], "Bora Dashboard H9"),
         ("Case_P50", "Net P50 hours", "h/yr", NUM, [2500, 2460, 2600], "Bora Dashboard H10"),
         ("Case_OM", "O&M full-service contract", "EUR k/MW/yr", NUM1, [12, 12, 10], "Bora Dashboard H11")]


def key_inputs(wb, r):
    bar(r, "Key inputs: the bids", HEADS)
    r += 1
    for key, text, unit, fmt, vals in KEY:
        R[key] = r
        label(r, text, unit)
        for c, x in zip(P, vals):
            put(f"{c}{r}", x, "input", fmt)
        name(wb, key, f"Dashboard!$E${r}:$G${r}")
        r += 1
    put(f"H{R['Dash_MVA']}", f"=SUM(E{R['Dash_MVA']}:G{R['Dash_MVA']})", fmt=NUM)
    r += 1
    bar(r, "Case: turbine price, PPA, yield, O&M", {"E": "Base", "F": "Low", "G": "High", "H": "Live"})
    R["case_hdr"] = r
    r += 1
    R["Dash_Case"] = r
    label(r, "Case selected (Base, Low or High)", "case", bold=True)
    put(f"H{r}", "Base", "input")
    name(wb, "Dash_Case", f"Dashboard!$H${r}")
    dv = DataValidation(type="list", formula1='"Base,Low,High"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(f"H{r}")
    r += 1
    for key, text, unit, fmt, vals, note in CASES:
        R[key] = r
        label(r, text, unit)
        for c, x in zip(P, vals):
            put(f"{c}{r}", x, "input", fmt)
        name(wb, "Dash_" + key, f"Dashboard!$E${r}:$G${r}")
        put(f"H{r}", f"=INDEX(E{r}:G{r},MATCH(Dash_Case,$E${R['case_hdr']}:$G${R['case_hdr']},0))", fmt=fmt, bold=True)
        name(wb, key, f"Dashboard!$H${r}")
        put(f"J{r}", note, "note")
        r += 1
    return r + 1


POINTS = [("pt_mun", "Points – municipal one-off",
           "=INDEX(Osz1_Pts,1+SUMPRODUCT((Osz1_LB<{c}{Dash_Municipal})+(Osz1_LB={c}{Dash_Municipal})*Osz1_Incl))*Osz1_W"),
          ("pt_com", "Points – community contribution",
           "=INDEX(Osz2_Pts,1+SUMPRODUCT((Osz2_LB<{c}{Dash_Community})+(Osz2_LB={c}{Dash_Community})*Osz2_Incl))*Osz2_W"),
          ("pt_ob", "Points – overbuild",
           "=INDEX(Osz3_Pts,1+SUMPRODUCT((Osz3_LB<{c}{Dash_Overbuild})+(Osz3_LB={c}{Dash_Overbuild})*Osz3_Incl))*Osz3_W"),
          ("pt_bess", "Points – BESS",
           "=INDEX(Osz3_Pts,1+SUMPRODUCT((Osz3_LB<{c}{Dash_BESS_Share})+(Osz3_LB={c}{Dash_BESS_Share})*Osz3_Incl))*Osz4_W"),
          ("pt_part", "Points – partial award", "={c}{Dash_Partial}*Osz5_Pts*Osz5_W"),
          ("pt_pb", "Points – performance security",
           "=INDEX(Osz6_Pts,1+SUMPRODUCT((Osz6_LB<{c}{Dash_PB_Mult})+(Osz6_LB={c}{Dash_PB_Mult})*Osz6_Incl))*Osz6_W"),
          ("pt_brown", "Points – brownfield", "={c}{Dash_Brownfield}*Osz7_Pts*Osz7_W"),
          ("pt_shared", "Points – shared connection", "={c}{Dash_Shared}*Osz8_Pts*Osz8_W"),
          ("pt_eut", "Points – EU turbines", '=N({c}{Dash_Turbine}="EU")*Osz9a_Pts*Osz9_W'),
          ("pt_eutr", "Points – EU transformers", "=Inputs!{c}$%d*Osz9b_Pts*Osz9_W")]


def points_capacity(r):
    bar(r, "Bid points and capacity", HEADS)
    r += 1
    for key, text, f in POINTS:
        r = rows3(r, key, text, "points", NUM1, f.replace("%d", str(I["trafo"])), total=False)
    r = rows3(r, "pt_tot", "Points total", "points", NUM1, "=SUM({c}{pt_mun}:{c}{pt_eutr})", total=False, bold=True,
              top=True)
    r = rows3(r, "wind", "Wind committed in the bid", "MW", NUM1, "={c}{Dash_MVA}*(1+{c}{Dash_Overbuild})")
    r = rows3(r, "mwh", "BESS energy", "MWh", NUM1, "={c}{Dash_MVA}*{c}{Dash_BESS_Share}*{c}{Dash_Hours}")
    r = rows3(r, "built", "Wind built (committed less the build shortfall)", "MW", NUM1,
              "={c}{wind}*(1-{c}{Dash_Shortfall})")
    return r + 1


# (b) deposit summary in the user's layout
def deposits(r):
    base = "({c}{wind}*Bench_Wind+{c}{mwh}*Bench_BESS)"
    bar(r, "Guarantees, securities and connection fees, DEVEX (HUF)", DATES, ["J"])
    r += 1
    huf = [("h_pf", "Participation fee", "=Part_Fee", "Date_Bid", False),
           ("h_bs", "Bid security", "=Bid_Sec_Rate*" + base, "Date_Bid", False),
           ("h_pbg", "Performance security, gross", "=PB_Rate*{c}{Dash_PB_Mult}*" + base, "Date_PB", False),
           ("h_pbc", "Less bid security credited", "=-{c}{h_bs}", None, False),
           ("h_pb", "Performance security payable (bid security deducted)", "={c}{h_pbg}+{c}{h_pbc}", "Date_PB", False),
           ("h_tso", "TSO connection-request security", "=TSO_Rate*{c}{Dash_MVA}", "Date_TSO", False),
           ("h_dir", "Direct connection fee (per point ÷ bids sharing it)",
            "=Inputs!{c}$%d/Inputs!{c}$%d" % (I["dir"], I["share_pt"]), "Date_HC", False),
           ("h_indg", "Indirect connection fee, gross", "=Inputs!{c}$%d*{c}{Dash_MVA}" % I["ind"], "Date_HC", False),
           ("h_tsoc", "Less TSO security credited", "=-{c}{h_tso}", None, False),
           ("h_ind", "Indirect connection fee, payable (TSO security deducted)", "={c}{h_indg}+{c}{h_tsoc}", "Date_HC",
            False),
           ("h_dev", "DEVEX", "=SUM(Inputs!{c}$%d:{c}$%d)*FX*Thousand" % (I["dv_grid"], I["dv_cont"]), "Date_Devex",
            False),
           ("h_tot", "Total cash: fee, securities, connection fees, DEVEX",
            "={c}{h_pf}+{c}{h_bs}+{c}{h_pb}+{c}{h_tso}+{c}{h_dir}+{c}{h_ind}+{c}{h_dev}", None, True)]
    for key, text, f, date, bold in huf:
        r = rows3(r, key, text, "HUF", NUM, f, bold=bold, top=bold, date=date)
    r += 1
    bar(r, "Guarantees, securities and connection fees, DEVEX (EUR k)", DATES, ["J"])
    r += 1
    eur = [("e_pf", "Participation fee", "h_pf", "Date_Bid"), ("e_bs", "Bid security", "h_bs", "Date_Bid"),
           ("e_pb", "Performance security payable (bid security deducted)", "h_pb", "Date_PB"),
           ("e_tso", "TSO connection-request security", "h_tso", "Date_TSO"),
           ("e_dir", "Direct connection fee", "h_dir", "Date_HC"),
           ("e_ind", "Indirect connection fee, payable (TSO security deducted)", "h_ind", "Date_HC"),
           ("e_dev", "DEVEX", "h_dev", "Date_Devex")]
    for key, text, src, date in eur:
        r = rows3(r, key, text, "EUR k", NUM1 if key == "e_pf" else NUM, "={c}{%s}/FX/Thousand" % src, date=date)
    r = rows3(r, "e_sdd", "Signing fee and DD", "EUR k", NUM, "=Inputs!{c}$%d" % I["sdd"], date="Date_Sign")
    r = rows3(r, "e_tot", "Total cash to RtB (incl. signing and DD)", "EUR k", NUM, "=SUM({c}{e_pf}:{c}{e_sdd})",
              bold=True, top=True)
    return r + 1


# (c) capex and devex breakdowns
def breakdowns(r):
    bar(r, "Capex breakdown to COD", HEADS)
    r += 1
    capex = [("k_turb", "Turbines (route price, live case)", '=IF({c}{Dash_Turbine}="EU",Case_EU,Case_CN)*{c}{built}'),
             ("k_inst", "Installation and erection", "=Capex_Install*{c}{built}"),
             ("k_fdn", "Foundations", "=Capex_Fdn*{c}{built}"),
             ("k_civ", "Balance of plant – civil", "=Capex_Civil*{c}{built}"),
             ("k_elec", "Balance of plant – electrical", "=Capex_Elec*{c}{built}"),
             ("k_own", "Owner's costs, engineering, PM, construction insurance", "=Capex_Owner*{c}{built}"),
             ("k_land", "Land purchase and easements", "=Capex_Land*{c}{built}"),
             ("k_subs", "Substation and 132 kV bay (scaled by MVA)", "=Subs_Lump*{c}{Dash_MVA}/Lump_Basis"),
             ("k_cont", "Contingency (not on BESS)", "=SUM({c}{k_turb}:{c}{k_subs})*Capex_Cont"),
             ("k_bess", "BESS (on BESS energy)", "=BESS_Capex*{c}{mwh}")]
    for key, text, f in capex:
        r = rows3(r, key, text, "EUR k", NUM, f)
    r = rows3(r, "k_capex", "Construction capex", "EUR k", NUM, "=SUM({c}{k_turb}:{c}{k_bess})", bold=True, top=True)
    r = rows3(r, "k_conn", "Connection fees (TSO security, direct, indirect)", "EUR k", NUM,
              "={c}{e_tso}+{c}{e_dir}+{c}{e_ind}")
    r = rows3(r, "k_muni", "Municipal one-off", "EUR k", NUM, "={c}{Dash_Municipal}*{c}{Dash_MVA}")
    r = rows3(r, "k_dev", "Development (participation fee, devex, signing and DD)", "EUR k", NUM,
              "={c}{e_pf}+{c}{e_dev}+{c}{e_sdd}")
    r = rows3(r, "k_allin", "All-in cost to COD (securities excluded: returned)", "EUR k", NUM,
              "={c}{k_capex}+SUM({c}{k_conn}:{c}{k_dev})", bold=True, top=True)
    R["k_mw"] = r
    label(r, "All-in cost per MW built", "EUR k/MW")
    for c in [*P, "H"]:
        put(f"{c}{r}", f"={c}{R['k_allin']}/{c}{R['built']}", fmt=NUM1)
    r += 1
    R["k_mva"] = r
    label(r, "All-in cost per MVA", "EUR k/MVA")
    for c in [*P, "H"]:
        put(f"{c}{r}", f"={c}{R['k_allin']}/{c}{R['Dash_MVA']}", fmt=NUM1)
    r += 2
    bar(r, "Devex breakdown (Bora: devex-file categories; tender fee and securities: above)", HEADS)
    r += 1
    for key, text, _ in DEVEX:
        r = rows3(r, "b_" + key, text[len("Devex: "):].capitalize(), "EUR k", NUM1, "=Inputs!{c}$%d" % I[key])
    r = rows3(r, "b_tot", "Devex total (= DEVEX line above)", "EUR k", NUM1, "=SUM({c}{b_dv_grid}:{c}{b_dv_cont})",
              bold=True, top=True)
    return r + 1


# (d) three routes and (e) annual cash flow: summary panels (links to the detail below)
EXEC = ["RtB exit (sale at RtB, 30-Sep-2029)",
        ("ex_cash", "Total FEP cash out to RtB (peak exposure)", "EUR k", NUM),
        ("ex_proc", "Total proceeds to FEP at the sale", "EUR k", NUM),
        ("ex_rp", "FEP profit, pre-tax", "EUR k", NUM), ("ex_rpost", "FEP profit, post-tax", "EUR k", NUM),
        ("ex_rmoic", "MOIC, pre-tax", "x", MULT), ("ex_rirr", "IRR, pre-tax", "%", PCT),
        "COD exit (sale at COD to a buyer earning the discount rate)",
        ("ex_allin", "Total capex and costs to COD", "EUR k", NUM), ("ex_price", "Sale price at COD", "EUR k", NUM),
        ("ex_xp", "FEP profit, pre-tax", "EUR k", NUM), ("ex_xpost", "FEP profit, post-tax", "EUR k", NUM),
        ("ex_xmoic", "MOIC, pre-tax", "x", MULT), ("ex_xirr", "IRR, pre-tax", "%", PCT),
        "Hold to maturity (EBITDA only, no debt, no tax)",
        ("ex_hcap", "Total capex and costs to COD", "EUR k", NUM),
        ("ex_rev", "Total revenue over the life", "EUR k", NUM), ("ex_opex", "Total opex over the life", "EUR k", NUM),
        ("ex_eb", "Total EBITDA over the life", "EUR k", NUM), ("ex_irr", "Project IRR, pre-tax", "%", PCT),
        ("ex_npv", "NPV at the discount rate, at the first year end", "EUR k", NUM)]
BOLD_EX = ("ex_rpost", "ex_xpost", "ex_irr")
CF = [("cf_spend", "Spending (capex, fees, securities, devex), net", "p_inv"), ("cf_rev", "Revenue", "p_rev"),
      ("cf_opex", "Opex", "p_opex"), ("cf_eb", "EBITDA", "p_ebitda"), ("cf_ncf", "Net cash flow", "p_cf")]


def routes_reserve(r):
    bar(r, "Three routes, totals", HEADS)
    r += 1
    for item in EXEC:
        if isinstance(item, str):
            sub(r, item)
        else:
            R[item[0]] = r
            label(r, item[1], item[2], bold=item[0] in BOLD_EX)
        r += 1
    r += 1
    bar(r, "Annual cash flow per project (EUR k)", {"H": "Total"}, TL)
    r += 1
    R["cf_yr"] = r
    label(r, "Year", "year")
    r += 2
    cf = {}
    for c, nm in zip(P, NAMES):
        sub(r, nm, extra=TL)
        r += 1
        rows = {}
        for key, text, _ in CF:
            rows[key] = r
            label(r, text, "EUR k", bold=key == "cf_ncf")
            r += 1
        rows["cf_cum"] = r
        label(r, "Cumulative net cash flow", "EUR k")
        r += 2
        cf[c] = rows
    R["cf"] = cf
    R["summary"] = (R["ex_cash"], r - 1)
    return r


# (f) detail: claims, waterfall, dated flows, timeline, COD tax, checks
def claims_waterfall(r):
    bar(r, "RtB exit: FEP claims at the sale (interest to the freeze, act/365, capitalised 31-Dec)", HEADS)
    r += 1
    r = rows3(r, "c_end", "Interest accrual end (earlier of freeze and sale)", "date", DATE,
              "=MIN(Freeze_Date,Sale_Date)", total=False)
    for key, rate, text in (("c_s12", "Rate_Cash", "cash rate"), ("c_s20", "Rate_Devex", "devex rate")):
        r = rows3(r, key, f"Accrual factor in the end year, {text}", "x", "0.0000",
                  "=1+%s*({c}{c_end}-DATE(YEAR({c}{c_end}),1,1)+1)/Days_In_Year" % rate, total=False)
    for ev, text, rate, stub, x in (("bid", "bid submission", "Rate_Cash", "c_s12", "Date_Bid"),
                                    ("pb", "performance security posting", "Rate_Cash", "c_s12", "Date_PB"),
                                    ("tso", "TSO security payment", "Rate_Cash", "c_s12", "Date_TSO"),
                                    ("hc", "connection agreement (HCSSZ)", "Rate_Cash", "c_s12", "Date_HC"),
                                    ("dev", "devex payment (average)", "Rate_Devex", "c_s20", "Date_Devex")):
        r = rows3(r, "c_d_" + ev, f"Days to the first 31-Dec from {text}", "days", INT,
                  "=DATE(YEAR(%s)+1,1,1)-%s" % (x, x), total=False)
        r = rows3(r, "c_f_" + ev, f"Accrual factor from {text}", "x", "0.0000",
                  "=(1+%s*{c}{c_d_%s}/Days_In_Year)*(1+%s)^(YEAR({c}{c_end})-YEAR(%s)-1)*{c}{%s}"
                  % (rate, ev, rate, x, stub), total=False)
    lines = [("c_pf", "Claim: participation fee (not reimbursed)", "={c}{e_pf}*({c}{c_f_bid}-1)"),
             ("c_bs", "Claim: bid security (cash)", "={c}{e_bs}*({c}{c_f_bid}-1)"),
             ("c_pb", "Claim: performance security payable (cash)", "={c}{e_pb}*({c}{c_f_pb}-1)"),
             ("c_tso", "Claim: TSO connection-request security", "={c}{e_tso}*({c}{c_f_tso}-1)"),
             ("c_dir", "Claim: direct connection fee", "={c}{e_dir}*({c}{c_f_hc}-1)"),
             ("c_ind", "Claim: indirect connection fee, payable", "={c}{e_ind}*({c}{c_f_hc}-1)"),
             ("c_dev", "Claim: devex, MAX(20 % accrued, 1× principal)", "={c}{e_dev}*MAX({c}{c_f_dev}-1,Devex_Floor-1)")]
    for key, text, f in lines:
        r = rows3(r, key, text, "EUR k", NUM1 if key == "c_pf" else NUM, f)
    r = rows3(r, "c_tot", "Total FEP claims", "EUR k", NUM, "=SUM({c}{c_pf}:{c}{c_dev})", bold=True, top=True)
    r = rows3(r, "c_dsa", "Developer success fee due (DSA, capped)", "EUR k", NUM,
              "={c}{Dash_DSA}*MIN(DSA_Rate*YEARFRAC(Date_Sign,Sale_Date),DSA_Cap)*{c}{wind}")
    r += 1
    bar(r, "RtB exit: waterfall at the sale", HEADS)
    r += 1
    wf = [("w_prem", "RtB premium", "={c}{Dash_Premium}*{c}{Dash_MVA}", False),
          ("w_reimb", "Buyer reimburses securities and connection fees at cost", "=SUM({c}{e_bs}:{c}{e_ind})", False),
          ("w_proc", "Total proceeds", "={c}{w_prem}+{c}{w_reimb}", True),
          ("w_sc", "Less sale costs", "=-MIN({c}{w_proc},Sale_Costs)", False),
          ("w_rp", "Less securities and fees repaid to FEP", "=-MIN(SUM({c}{w_proc}:{c}{w_sc}),{c}{w_reimb})", False),
          ("w_c12", "Less FEP claims at 12 %", "=-MIN(SUM({c}{w_proc}:{c}{w_rp}),SUM({c}{c_pf}:{c}{c_ind}))", False),
          ("w_dev", "Less FEP devex return", "=-MIN(SUM({c}{w_proc}:{c}{w_c12}),{c}{c_dev})", False),
          ("w_dsa", "Less developer success fee (DSA)", "=-MIN(SUM({c}{w_proc}:{c}{w_dev}),{c}{c_dsa})", False),
          ("w_dist", "Distributable proceeds", "=SUM({c}{w_proc}:{c}{w_dsa})", True),
          ("w_fep", "FEP share of distributable", "={c}{w_dist}*{c}{Dash_FEP_Share}", False),
          ("w_devsh", "Developer share of distributable", "={c}{w_dist}-{c}{w_fep}", False)]
    for key, text, f, bold in wf:
        r = rows3(r, key, text, "EUR k", NUM, f, bold=bold, top=key == "w_dist")
    r += 1
    sub(r, "FEP cash flows, dated (outlays, then receipts at the sale)", DATES, ["J"])
    r += 1
    flows = [("f_sdd", "Signing fee and DD", "e_sdd", "Date_Sign"), ("f_pf", "Participation fee", "e_pf", "Date_Bid"),
             ("f_bs", "Bid security", "e_bs", "Date_Bid"), ("f_pb", "Performance security payable", "e_pb", "Date_PB"),
             ("f_tso", "TSO connection-request security", "e_tso", "Date_TSO"),
             ("f_dir", "Direct connection fee", "e_dir", "Date_HC"),
             ("f_ind", "Indirect connection fee, payable", "e_ind", "Date_HC"), ("f_dev", "Devex", "e_dev", "Date_Devex")]
    for key, text, src, date in flows:
        r = rows3(r, key, text, "EUR k", NUM1 if key == "f_pf" else NUM, "=-{c}{%s}" % src, date=date)
    r = rows3(r, "f_rec", "FEP receipts at the sale (repaid, claims, share)", "EUR k", NUM,
              "={c}{w_fep}-SUM({c}{w_rp}:{c}{w_dev})", bold=True, date="Sale_Date")
    r = rows3(r, "r_profit", "FEP profit, pre-tax", "EUR k", NUM, "=SUM({c}{f_sdd}:{c}{f_rec})", bold=True, top=True)
    r = rows3(r, "r_cit", "CIT on the gain", "EUR k", NUM, "=-MAX(0,{c}{r_profit})*CIT_Rate")
    r = rows3(r, "r_post", "FEP profit, post-tax", "EUR k", NUM, "={c}{r_profit}+{c}{r_cit}", bold=True)
    return r + 1


SUBROWS = [
    ("t_n", "Years from the COD year", "years", INT, None, "={x}${t_yr}-YEAR(Date_COD)"),
    ("t_cm", "Construction months by year end (value: construction period)", "months", INT,
     '=DATEDIF(Date_NTP,Date_COD,"m")',
     "=({x}${t_yr}+1-YEAR(Date_NTP))*Months_In_Year-MONTH(Date_NTP)"),
    ("t_cs", "Construction share of the year (NTP to COD)", "%", PCT, None,
     "=MAX(0,MIN({x}{t_cm},${p}{t_cm})-MAX({x}{t_cm}-Months_In_Year,0))/${p}{t_cm}"),
    ("t_om", "Months operated by year end (value: operating life)", "months", INT, "=Life*Months_In_Year",
     "=({x}${t_yr}+1-YEAR(Date_COD))*Months_In_Year-MONTH(Date_COD)"),
    ("t_op", "Operating share of the year (COD to end of life)", "%", PCT, None,
     "=MAX(0,MIN({x}{t_om},${p}{t_om})-MAX({x}{t_om}-Months_In_Year,0))/Months_In_Year"),
    ("t_cod", "COD year flag", "flag", FLAG, None, "=N({x}{t_n}=0)"),
    ("p_dev", "FEP outlays at their dates (signing, fee, securities, fees, devex)", "EUR k", NUM, None,
     "=SUMPRODUCT((YEAR(${d}${f_sdd}:${d}${f_dev})={x}${t_yr})*${p}${f_sdd}:${p}${f_dev})"),
    ("p_pbr", "Performance security released (bid security included)", "EUR k", NUM, None,
     "=(${p}${e_bs}+${p}${e_pb})*({x}${t_yr}=YEAR(Date_Release))"),
    ("p_capex", "Construction capex (straight-line NTP to COD)", "EUR k", NUM, None, "=-${p}${k_capex}*{x}{t_cs}"),
    ("p_muni", "Municipal one-off at COD", "EUR k", NUM, None, "=-${p}${k_muni}*{x}{t_cod}"),
    ("p_inv", "Investment cash flow", "EUR k", NUM, None, "=SUM({x}{p_dev}:{x}{p_muni})"),
    ("p_deg", "Degradation factor (from the first full year)", "x", FAC, None, "=(1-Degradation)^MAX({x}{t_n}-1,0)"),
    ("p_gen", "Net generation", "GWh", NUM1, None,
     "=${p}${built}*Case_P50*(1-Curtailment)*(1-Clipping)*{x}{t_op}*{x}{p_deg}/Thousand"),
    ("p_px", "PPA price", "EUR/MWh", NUM1, None, "=Case_PPA*(1+PPA_Esc)^{x}{t_n}*({x}{t_op}>0)"),
    ("p_rev", "Revenue", "EUR k", NUM, None, "={x}{p_gen}*{x}{p_px}"),
    ("p_idx", "Opex index (COD year = 1)", "x", FAC, None, "=(1+Opex_Esc)^{x}{t_n}"),
    ("p_fix", "O&M, BoP, lease, insurance, management, property tax", "EUR k", NUM, None,
     "=-(Case_OM+SUM(Opex_Per_MW))*${p}${built}*{x}{p_idx}*{x}{t_op}"),
    ("p_adm", "Administration, audit, compliance", "EUR k", NUM, None,
     "=-Opex_Admin*${p}${Dash_MVA}/Lump_Basis*{x}{p_idx}*{x}{t_op}"),
    ("p_bg", "Balancing and grid charges", "EUR k", NUM, None, "=-{x}{p_gen}*(Balancing+Grid_Charges)*{x}{p_idx}"),
    ("p_mekh", "MEKH supervisory fee", "EUR k", NUM, None, "=-{x}{p_rev}*MEKH_Fee"),
    ("p_oc", "Opex contingency", "EUR k", NUM, None, "=SUM({x}{p_fix}:{x}{p_mekh})*Opex_Cont"),
    ("p_comm", "Community contribution", "EUR k", NUM, None, "=-{x}{p_gen}*${p}${Dash_Community}"),
    ("p_hipa", "HIPA local business tax", "EUR k", NUM, None, "=-{x}{p_rev}*HIPA_Rate"),
    ("p_opex", "Total opex", "EUR k", NUM, None, "=SUM({x}{p_fix}:{x}{p_hipa})"),
    ("p_ebitda", "EBITDA", "EUR k", NUM, None, "={x}{p_rev}+{x}{p_opex}"),
    ("p_cf", "Unlevered pre-tax cash flow", "EUR k", NUM, None, "={x}{p_inv}+{x}{p_ebitda}"),
    ("p_df", "Discount factor to COD (buyer's rate)", "x", FAC, None,
     "=(1+Discount_Rate)^((Date_COD-{x}${t_end})/Days_In_Year)"),
    ("p_pv", "EBITDA valued at COD (total = COD price)", "EUR k", NUM, None, "={x}{p_ebitda}*{x}{p_df}"),
    ("p_cod", "COD exit: FEP cash flow (to COD, price and securities at COD)", "EUR k", NUM, None,
     "={x}{p_inv}*({x}{t_n}<=0)+{x}{t_cod}*($H{p_pv}+${p}${e_bs}+${p}${e_pb})"),
]
NO_TOTAL = {"t_n", "t_cm", "t_om", "t_cod", "p_deg", "p_px", "p_idx", "p_df"}
BOLD = {"p_inv", "p_rev", "p_opex", "p_ebitda", "p_cf", "p_cod"}


def timeline(r):
    bar(r, "Annual timeline", {"H": "Total"}, TL)
    r += 1
    R["t_yr"] = r
    label(r, "Year", "year")
    for x in TL:
        put(f"{x}{r}", "=First_Year+COLUMN()-COLUMN($J$1)", fmt=INT)
    r += 1
    R["t_end"] = r
    label(r, "Year end", "date")
    for x in TL:
        put(f"{x}{r}", f"=DATE({x}{R['t_yr']}+1,1,1)-1", fmt=DATE)
    r += 2
    proj = {}
    for c, nm in zip(P, NAMES):
        bar(r, f"{nm}: annual cash flow (EUR k)", {c: "Value", "H": "Total"}, TL)
        r += 1
        rows = {}
        for key, text, unit, fmt, const, f in SUBROWS:
            rows[key] = R[key] = r
            bold = key in BOLD
            label(r, text, unit, bold=bold)
            kw = dict(p=c, d=DCOL[c])
            if const:
                put(f"{c}{r}", fx(const, **kw), fmt=INT)
            for x in TL:
                put(f"{x}{r}", fx(f, x=x, **kw), fmt=fmt, bold=bold, top=bold and key not in ("p_inv", "p_ebitda"))
            if key not in NO_TOTAL:
                put(f"H{r}", f"=SUM(J{r}:AS{r})", fmt=NUM2 if fmt == PCT else fmt, bold=bold,
                    top=bold and key not in ("p_inv", "p_ebitda"))
            r += 1
        proj[c] = rows
        r += 1
    R["proj"] = proj
    return r


def cod_tax(r):
    bar(r, "COD exit: tax on the gain (price and profit: totals of the cash-flow blocks above)", HEADS)
    r += 1
    pr = R["proj"]
    for key, text, f, bold in (("x_cit", "CIT on the gain", "=-MAX(0,$H${cod})*CIT_Rate", False),
                               ("x_post", "FEP profit, post-tax", "=$H${cod}+{c}{x_cit}", True)):
        R[key] = r
        label(r, text, "EUR k", bold=bold)
        for c in P:
            put(f"{c}{r}", fx(f, c=c, cod=pr[c]["p_cod"]), fmt=NUM, bold=bold)
        put(f"H{r}", f"=SUM(E{r}:G{r})", fmt=NUM, bold=bold)
        r += 1
    return r + 1


def checks(r):
    bar(r, "Checks (0 = OK)", HEADS)
    r += 1
    pr = R["proj"]
    first = r
    defs = [("ch_wf", "RtB waterfall closes (proceeds = all amounts paid out)",
             lambda c: fx("=N(ABS({c}{w_proc}+{c}{w_sc}+{c}{w_dsa}-{c}{w_devsh}-{c}{f_rec})>Tolerance)", c=c)),
            ("ch_end", "Claim lines paid before the accrual end year",
             lambda c: fx("=SUMPRODUCT(--(YEAR({d}{f_pf}:{d}{f_dev})>=YEAR({c}{c_end})))", c=c, d=DCOL[c])),
            ("ch_route", "Turbine route is EU or CN",
             lambda c: fx('=N(AND({c}{Dash_Turbine}<>"EU",{c}{Dash_Turbine}<>"CN"))', c=c)),
            ("ch_devex", "Devex breakdown = DEVEX line",
             lambda c: fx("=N(ABS({c}{b_tot}-{c}{e_dev})>Tolerance)", c=c)),
            ("ch_capex", "Construction capex spread in full (timeline = capex)",
             lambda c: fx("=N(ABS({c}{k_capex}+$H$%d)>Tolerance)" % pr[c]["p_capex"], c=c)),
            ("ch_allin", "Investment cash flows = all-in cost to COD",
             lambda c: fx("=N(ABS({c}{k_allin}+$H$%d)>Tolerance)" % pr[c]["p_inv"], c=c)),
            ("ch_life", "Operating life inside the timeline",
             lambda c: "=N(ABS($H$%d-Life)>Tolerance)" % pr[c]["t_op"])]
    for key, text, f in defs:
        R[key] = r
        label(r, text, "flag")
        for c in P:
            put(f"{c}{r}", f(c), fmt=FLAG)
        put(f"H{r}", f"=SUM(E{r}:G{r})", fmt=FLAG)
        r += 1
    R["master"] = r
    label(r, "Master check", "flag", bold=True)
    put(f"H{r}", f"=SUM(H{first}:H{r - 1})", fmt=CHECK_FMT, bold=True, top=True)
    r += 2
    ws[f"C{r}"].value, ws[f"C{r}"].font = "***END***", font(GREY, bold=True)
    return r


def fill_summary():
    """Route totals and the annual cash flow: links to the detail; IRR, MOIC and NPV on the detail cash-flow rows."""
    pr = R["proj"]
    ye = f"$J${R['t_end']}:$AS${R['t_end']}"
    fmts = {k[0]: k[3] for k in EXEC if not isinstance(k, str)}
    for c in P:
        q = pr[c]
        tl = lambda k: f"$J${q[k]}:$AS${q[k]}"  # noqa: E731
        fl = f"{c}{R['f_sdd']}:{c}{R['f_rec']}"
        f = {"ex_cash": f"={c}{R['e_tot']}", "ex_proc": f"={c}{R['f_rec']}", "ex_rp": f"={c}{R['r_profit']}",
             "ex_rpost": f"={c}{R['r_post']}", "ex_rmoic": f'=SUMIF({fl},">0")/-SUMIF({fl},"<0")',
             "ex_rirr": f"=XIRR({fl},{DCOL[c]}{R['f_sdd']}:{DCOL[c]}{R['f_rec']})",
             "ex_allin": f"={c}{R['k_allin']}", "ex_price": f"=$H${q['p_pv']}", "ex_xp": f"=$H${q['p_cod']}",
             "ex_xpost": f"={c}{R['x_post']}", "ex_xmoic": f'=SUMIF({tl("p_cod")},">0")/-SUMIF({tl("p_cod")},"<0")',
             "ex_xirr": f"=XIRR({tl('p_cod')},{ye})", "ex_hcap": f"={c}{R['k_allin']}",
             "ex_rev": f"=$H${q['p_rev']}", "ex_opex": f"=$H${q['p_opex']}", "ex_eb": f"=$H${q['p_ebitda']}",
             "ex_irr": f"=XIRR({tl('p_cf')},{ye})", "ex_npv": f"=XNPV(Discount_Rate,{tl('p_cf')},{ye})"}
        for key, form in f.items():
            put(f"{c}{R[key]}", form, fmt=fmts[key], bold=key in BOLD_EX)
        cf = R["cf"][c]
        for key, _, src in CF:
            for x in TL:
                put(f"{x}{cf[key]}", f"={x}{q[src]}", fmt=NUM, bold=key == "cf_ncf")
            put(f"H{cf[key]}", f"=SUM(J{cf[key]}:AS{cf[key]})", fmt=NUM, bold=key == "cf_ncf")
        for x in TL:
            put(f"{x}{cf['cf_cum']}", f"=SUM($J{cf['cf_ncf']}:{x}{cf['cf_ncf']})", fmt=NUM)
    for key in ("ex_cash", "ex_proc", "ex_rp", "ex_rpost", "ex_allin", "ex_price", "ex_xp", "ex_xpost", "ex_hcap",
                "ex_rev", "ex_opex", "ex_eb", "ex_npv"):
        put(f"H{R[key]}", f"=SUM(E{R[key]}:G{R[key]})", fmt=NUM, bold=key in BOLD_EX)
    hf = f"H{R['f_sdd']}:H{R['f_rec']}"
    put(f"H{R['ex_rmoic']}", f'=SUMIF({hf},">0")/-SUMIF({hf},"<0")', fmt=MULT)
    for x in TL:
        put(f"{x}{R['cf_yr']}", f"={x}{R['t_yr']}", fmt=INT, bold=True)


def build():
    global ws
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    build_inputs(wb)
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False)
    for col, w in zip("ABCDEFGHI", (2, 2, 56, 10, 13, 13, 13, 13, 2)):
        ws.column_dimensions[col].width = w
    for x in TL:
        ws.column_dimensions[x].width = 10
    ws["C2"].value = "HU wind bids: Bora (one 50 MVA low bid), Iváncsa, Dunaújváros – RtB exit, COD exit, hold"
    ws["C2"].font = font(bold=True)
    r = key_inputs(wb, 5)
    r = points_capacity(r)
    r = deposits(r)
    r = breakdowns(r)
    r = routes_reserve(r)
    bar(r, "Detailed calculations: claims, waterfall, dated flows, timelines, checks – expand the group", HEADS)
    R["detail"] = r
    r = claims_waterfall(r + 2)
    r = timeline(r)
    r = cod_tax(r)
    end = checks(r)
    for rr in range(R["detail"] + 1, end + 1):
        ws.row_dimensions[rr].outlineLevel, ws.row_dimensions[rr].hidden = 1, True
    fill_summary()
    label(3, "Master check (0 = OK; checks at the end of the detail)", "flag", bold=True)
    put("H3", "=Master_Check", fmt=CHECK_FMT, bold=True)
    for op, colr in (("equal", OK_GREEN), ("notEqual", CHECK_RED)):
        cf = CellIsRule(operator=op, formula=["0"], font=Font(color=colr, bold=True))
        ws.conditional_formatting.add("H3", cf)
        ws.conditional_formatting.add(f"H{R['master']}", cf)
    name(wb, "Master_Check", f"Dashboard!$H${R['master']}")
    ws.freeze_panes = "E4"
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    ROWS.write_text(json.dumps({"R": R, "I": I}, indent=1))
    print(f"wrote {OUT}")


if __name__ == "__main__":
    build()
