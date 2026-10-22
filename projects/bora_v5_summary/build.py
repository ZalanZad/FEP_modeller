"""Bora v5 Summary – one Dashboard sheet on the colleague's MOIC / HTM workbook. Built to SPEC.md.

Loads reference/HUN_Wind_MOIC_HTM_IC_Bora_v5_20261022.xlsm (read-only), re-points every link to the duplicate
sheets (Cover, Control_Panel, the old Dashboard, Bora, Run_Map) and deletes them, adds Inputs block U (points bands
and constants, rows 300+), edits only the engine cells of SPEC §4, §8, §9 and §10, adds Cash_Flow (HTM annual grid as links) and the
Dashboard (levers in section 1, two cases) as the first two sheets and packages the workbook with
vba/Scenarios.bas. The MOIC, HTM and Deposits engines are otherwise untouched (no rows or columns inserted).
regression.py runs the same build with the source lever values (SPEC §7 step 1).
"""
import re
import shutil
import tempfile
import zipfile
from copy import copy
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.dimensions import ColumnDimension
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

from fepm.recalc import recalc
from fepm.xl.builder import STORED_CHECK, _bar, _banner, _columns, _end, _ok_format, _status, _validate, _write
from fepm.xl.macros import VBA_DIR, package
from fepm.xl.styles import FMT, FONT, SIZE, STYLES, apply

HERE = Path(__file__).resolve().parent
SRC = HERE.parents[1] / "reference" / "HUN_Wind_MOIC_HTM_IC_Bora_v5_20261022.xlsm"
NAME = "Bora_v5_Summary"
DELETE = ("Cover", "Control_Panel", "Dashboard", "Bora", "Run_Map")
DELETED_REF = re.compile(r"(Cover|Control_Panel|Dashboard|Bora|Run_Map)'?!")
U0 = 300                                                   # first row of Inputs block U
BAND_COL = 17                                              # Q: points bands (J:P are hidden spare cases)

# Levers: the Dashboard section 1 inputs (Dash_*) and the case cells of SPEC §4 / §8.3 (Inputs G = EU, H = CN turbine).
NEW = {"Dash_MVA": 100, "In_Exit_Profit_Charge": 0, "Dash_Overbuild": 0.40, "Dash_BESS_Share": 0.40, "Dash_PB_Mult": 1.0, "Dash_Build_Shortfall": 0.10,
       "cases": {"G7": "EU turbine", "H7": "CN turbine",
                 "G8": "EU turbines 950 k€/MW, 60 % gearing, 90 €/MWh PPA on all output for 30 years",
                 "H8": "As EU turbine, but CN turbines 500 k€/MW and 50 % gearing",
                 "G78": 0, "G233": 50, "G234": 3, "G162": 950, "H162": 500, "G190": 0, "H190": None,
                 "G193": 30, "G217": 0.60, "H217": 0.50, "G11": 365, "G268": 0, "G269": 0}}
SOURCE = {"Dash_MVA": 200, "In_Exit_Profit_Charge": 1, "Dash_Overbuild": 0.0,    # the v5 file as received:
          "Dash_BESS_Share": 0.50, "Dash_PB_Mult": 1.5, "Dash_Build_Shortfall": 0.0,
          "cases": {"G7": "EU turbine", "H7": "CN turbine",                         # its cases 2 and 3
                    "G8": "v5 case 2 (regression)", "H8": "v5 case 3 (regression)",
                    "G78": 0.05, "G233": 20, "G234": 2, "G162": 950, "H162": 500, "G190": 0, "H190": 1,
                    "G193": 10, "G217": 0.50, "H217": 0, "G11": 370, "G268": 6, "G269": 0.5}}
# Dashboard section 1 levers that are the same in both builds: name -> value
LEVERS = {"Dash_BESS_Capex_MWh": 100, "Dash_Bottleneck_Loss": 0.05, "Dash_OSZ4_Basis": 0,
          "Dash_Partial": 1, "Dash_Brownfield": 0, "Dash_Shared": 1}

# Points tables (docs/domain/tender_2026_rules.md "Points tables"; layout of projects/bora_bidding/build.py).
# key: label, unit, lower bounds, lower bound included, inner points (first entry = below the first bound)
BANDS = {
    "1": ("ÖSZ-1 municipal one-off", "k€/MVA", [10, 15, 20, 25, 30, 35, 40, 45, 50], [0] * 8 + [1],
          [0, 10, 20, 30, 40, 55, 70, 80, 90, 100], "Kiírás Annex 4"),
    "2": ("ÖSZ-2 community contribution", "€/MWh", [2, 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.8, 2.9, 3],
          [0] * 10 + [1], [0, 5, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100], "Kiírás Annex 4"),
    "3": ("ÖSZ-3 overbuild and ÖSZ-4 BESS ratio", "%", [x / 100 for x in (2, 5, 10, 15, 20, 25, 30, 35, 40, 50)],
          [0] + [1] * 8 + [0], list(range(0, 101, 10)), "Kiírás Annex 4; GYIK Q20"),
    "6": ("ÖSZ-6 performance security multiple", "x", [1.0, 1.2, 1.3, 1.4, 1.5, 1.6, 1.7, 1.8, 1.9, 2.0],
          [0] + [1] * 9, list(range(0, 101, 10)), "Kiírás Annex 4; GYIK Q34"),
}


def u_rows(lv):
    """Block U scalars: (name, label, unit, value or formula, note, number format). The levers are on Dashboard."""
    return [
        ("In_MVA_Per_Bid", "Connection capacity per bid", "MVA", 100, "Bora bids are 100 MVA each", "money"),
        ("In_Bids", "Number of bids", "#", "=ROUNDUP(Dash_MVA/In_MVA_Per_Bid,0)", "Inputs E82 links here", "int"),
        ("In_Source_MVA", "Source basis of the lump sums", "MVA", 200, "Bora 2 × 100 MVA", "money"),
        ("In_Scale", "Lump-sum scale factor (capacity ÷ source basis)", "x", "=Dash_MVA/In_Source_MVA",
         "Inputs E24:E26, E167, E207, E294", "0.000"),
        ("In_Exit_Profit_Charge", "HTM charged the FEP profit forgone at RtB (1 = yes)", "flag",
         lv["In_Exit_Profit_Charge"], "SPEC §11.3: cash IRR, no opportunity cost", "flag"),
        ("In_Wind_Built", "Wind built (committed less the build shortfall)", "MW",
         "=E18*(1-Dash_Build_Shortfall)", "GYIK Q28/30/76; HTM E35, E27", "money1"),
        ("In_Bottleneck_Factor", "Generation kept after the bottleneck loss", "x",
         "=1-Dash_Bottleneck_Loss*(In_Wind_Built-Dash_MVA)/In_Wind_Built", "Used in HTM E36",
         "0.0000"),
        ("In_Min_Hours", "Minimum BESS duration", "h", 4, "Kiírás Annex 8; GYIK Q5", "int"),
        ("Points_Decimals", "Points rounded to decimals", "#", 2, "Kiírás X.1 b", "int"),
        ("Osz1_W", "ÖSZ-1 weight", "%", 0.20, "Kiírás Annex 3", "pct"),
        ("Osz2_W", "ÖSZ-2 weight", "%", 0.15, "Kiírás Annex 3", "pct"),
        ("Osz3_W", "ÖSZ-3 weight", "%", 0.10, "Kiírás Annex 3", "pct"),
        ("Osz4_W", "ÖSZ-4 weight", "%", 0.15, "Kiírás Annex 3", "pct"),
        ("Osz5_W", "ÖSZ-5 weight", "%", 0.05, "Kiírás Annex 3", "pct"),
        ("Osz6_W", "ÖSZ-6 weight", "%", 0.05, "Kiírás Annex 3", "pct"),
        ("Osz7_W", "ÖSZ-7 weight", "%", 0.10, "Kiírás Annex 3", "pct"),
        ("Osz8_W", "ÖSZ-8 weight", "%", 0.05, "Kiírás Annex 3", "pct"),
        ("Osz9_W", "ÖSZ-9 weight", "%", 0.15, "Kiírás Annex 3", "pct"),
        ("Osz5_Pts", "ÖSZ-5 partial award – inner points if yes", "points", 100, "Kiírás Annex 4", "int"),
        ("Osz7_Pts", "ÖSZ-7 brownfield – inner points if yes", "points", 100, "Kiírás Annex 4", "int"),
        ("Osz8_Pts", "ÖSZ-8 shared connection – inner points if yes", "points", 100, "Kiírás Annex 4", "int"),
        ("Osz9a_Pts", "ÖSZ-9 EU turbines – inner points if yes", "points", 80, "Kiírás Annex 4", "int"),
        ("Osz9b_Pts", "ÖSZ-9 EU transformers – inner points if yes", "points", 20, "Kiírás Annex 4", "int"),
        ("Osz1_Min", "ÖSZ-1 offerable range – minimum", "k€/MVA", 10, "Kiírás Annex 4", "money"),
        ("Osz1_Max", "ÖSZ-1 offerable range – maximum", "k€/MVA", 50, "Kiírás Annex 4 (A8)", "money"),
        ("Osz2_Min", "ÖSZ-2 offerable range – minimum", "€/MWh", 2, "Kiírás Annex 4", "money1"),
        ("Osz2_Max", "ÖSZ-2 offerable range – maximum", "€/MWh", 3, "Kiírás Annex 4 (A8)", "money1"),
        ("Osz4_Min", "ÖSZ-4 minimum BESS ratio", "%", 0.05, "Kiírás X.4 bd (A3)", "pct"),
        ("Thousand", "Thousand", "#", 1000, "Unit conversion", "money"),
        ("Hours_In_Year", "Hours in a year", "h", 8760, "Unit conversion", "money"),
        ("Tolerance", "Check tolerance", "#", 0.000001, "House standard", "0.000000"),
    ]


# =========================================================================================== helpers
def _fmt(key):
    return FMT.get(key, key)


def define(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def text(v):
    """Formula text of a cell value (array formulas included), else the value."""
    return v.text if isinstance(v, ArrayFormula) else v


def add_styles(wb):
    """FEP named styles for the new cells; the colleague's default font is left as it is."""
    for st in STYLES.values():
        if st.name not in wb.named_styles:
            wb.add_named_style(copy(st))


def like(cell, model, fmt=None, color=None):
    """Give `cell` the colleague's style of `model` (Inputs sheet look), optional number format and font colour."""
    cell._style = copy(model._style)
    if fmt:
        cell.number_format = fmt
    if color:
        font = copy(cell.font)
        font.color = color
        cell.font = font
    return cell


# =========================================================================================== §1 re-point, delete
def repoint(wb, cached):
    """Inputs G:I links to Control_Panel become blue values; the case number E6 reads the Dashboard selector (§5).
    A blank that another Inputs formula reads (H253 via H13 `=H$253`) is written as ="" so it stays a text blank,
    as in the source. G292 (devex target) carried the Control_Panel E22 formula: moved here so the devex scale
    factor G294 stays 1."""
    ws, cp = wb["Inputs"], wb["Control_Panel"]
    read = {m.group(1) + m.group(2) for row in ws.iter_rows(min_col=7, max_col=16) for c in row
            if isinstance(c.value, str) for m in re.finditer(r"(?<![A-Z!$])\$?([G-P])\$?(\d+)", c.value)}
    blue = ws["G268"]
    for row in ws.iter_rows():
        for c in row:
            f = text(c.value)
            if not (isinstance(f, str) and "Control_Panel!" in f):
                continue
            if c.coordinate == "E6":
                c.value = "=MATCH(Scenario_Name,Scenario_Names,0)"
                ws["F6"] = "Dashboard case selector (Scenario_Name)"
                continue
            if c.coordinate == "G292":
                c.value = ArrayFormula("G292", text(cp["E22"].value).replace("Inputs!", ""))
                continue
            v = cached["Inputs"][c.coordinate].value
            c.value = v if v not in (None, "") else ('=""' if c.coordinate in read else None)
            like(c, blue, c.number_format)


def delete_duplicates(wb):
    for name in DELETE:
        del wb[name]
    left = [f"{ws.title}!{c.coordinate}" for ws in wb.worksheets for row in ws.iter_rows() for c in row
            if isinstance(text(c.value), str) and str(text(c.value)).startswith("=")
            and DELETED_REF.search(text(c.value))]
    assert not left, f"formulas still refer to deleted sheets: {left}"


# =========================================================================================== §3 Inputs block U
def block_u(wb, lv):
    ws = wb["Inputs"]
    label, unit, note, calc, blue = ws["C268"], ws["D268"], ws["F268"], ws["E154"], ws["G268"]
    like(ws.cell(U0, 2, "x"), ws["B267"])
    like(ws.cell(U0, 3, "U. Tender points tables and constants (levers Dash_* are on the Dashboard)"), ws["C267"])
    r = U0 + 1
    for name, lab, un, value, src, fmt in u_rows(lv):
        like(ws.cell(r, 3, lab), label)
        like(ws.cell(r, 4, un), unit)
        cell = ws.cell(r, 5, value)
        if isinstance(value, str):                         # calculation
            like(cell, calc, _fmt(fmt))
        else:
            like(cell, blue, _fmt(fmt))
        like(ws.cell(r, 6, src), note)
        define(wb, name, f"Inputs!$E${r}")
        r += 1

    # Points bands across columns Q onward: one row each for lower bounds, inclusive flags and inner points.
    r += 1
    like(ws.cell(r, 3, "Points bands (band 1 = below the first lower bound)"), ws["C277"])
    for i in range(12):
        like(ws.cell(r, BAND_COL + i, i + 1), ws["G5"])
    for col in range(BAND_COL, BAND_COL + 12):
        ws.column_dimensions[L(col)].width = 7
    r += 1
    for k, (lab, un, lb, incl, pts, src) in BANDS.items():
        for suffix, what, u, vals, fmt in (("LB", "lower bound", un, lb, "General"),
                                           ("Incl", "lower bound included", "flag", incl, "0"),
                                           ("Pts", "inner points", "points", pts, "0")):
            like(ws.cell(r, 3, f"{lab} – {what}"), label)
            like(ws.cell(r, 4, u), unit)
            like(ws.cell(r, 6, src), note)
            for i, v in enumerate(vals):
                like(ws.cell(r, BAND_COL + i, v), blue, fmt)
            define(wb, f"Osz{k}_{suffix}", f"Inputs!${L(BAND_COL)}${r}:${L(BAND_COL + len(vals) - 1)}${r}")
            r += 1
    return r


# =========================================================================================== §4 engine cells
def engine(wb, lv):
    inp, dep, htm, moic = wb["Inputs"], wb["Deposits"], wb["HTM"], wb["MOIC"]
    inp["E16"] = "=Dash_MVA*(1+Dash_Overbuild)"          # installed wind = committed
    inp["E20"] = "=Dash_BESS_Share*Dash_MVA"
    inp["E168"] = "=Dash_BESS_Capex_MWh*E21"                            # 100 k€/MWh × hours
    inp["E90"] = "=Dash_PB_Mult"
    inp["E82"] = "=In_Bids"                                           # §11.1 bids follow the capacity
    for ref in ("E24", "E25", "E26", "E167", "E207"):                  # §11.6 lump sums quoted for 200 MVA
        inp[ref] = f"=({text(inp[ref].value)[1:]})*In_Scale"
    inp["E294"] = "=In_Scale"                                          # devex line items (Devex_FEP column E)
    inp["E173"] = ("=(SUM(E162:E166)+E169+E170)*(1+E171)+E167*(1+E171)/In_Wind_Built"   # §10.2 per built MW
                   "+E168*E20/In_Wind_Built")
    inp["E174"] = "=E173+E172/In_Wind_Built"
    inp["D173"] = inp["D174"] = "k€/MW built"
    for r in (16, 20):                                                 # stale case values, now unused
        inp[f"G{r}"] = None
        inp[f"F{r}"] = f"Driven by the Dashboard levers (E{r})"
    for ref, v in lv["cases"].items():                                 # case values (blue inputs already)
        inp[ref] = v
    for ref in ("G11", "G12", "G18"):                                  # TSO security and connection fees per MVA
        dep[ref] = "=Dash_MVA*Inputs!$E$17"
    for col in range(8, 44):                                           # H:AQ municipal one-off per MVA
        c = L(col)
        htm[f"{c}29"] = f"=-IF({c}5=YEAR(Inputs!$E$155),Inputs!$E$233*Dash_MVA,0)"
    htm["E21"] = "=In_Exit_Profit_Charge*Inputs!$E$228"              # §11.3 no opportunity cost in a cash IRR
    htm["E35"] = "=In_Wind_Built"                                      # §10.2 generation, opex on built MW
    htm["E27"] = "=Inputs!$E$173*In_Wind_Built"
    htm["G32"] = "=-G30/In_Wind_Built"                                 # G150 links G32
    htm["C32"] = htm["C150"] = "Capex per MW built"
    htm["E36"] = "=Inputs!$E$179*(1-Inputs!$E$180)*(1-Inputs!$E$181)*Inputs!$E$183*In_Bottleneck_Factor"
    moic["G143"] = "=Inputs!$E$141*Dash_MVA"            # RtB price per MVA
    rtb_basis(wb)
    cod_sale(htm)
    deposits_check(dep)


def rtb_basis(wb):
    """§8.1: the RtB premium sensitivities and the implied maximum premium on the MVA basis (k€/MVA)."""
    moic, htm = wb["MOIC"], wb["HTM"]
    for col in "HIJKLMN":
        moic[f"{col}226"] = f"={col}225*Dash_MVA"
    for r in list(range(227, 234)) + list(range(237, 244)):
        htm[f"E{r}"] = f"=D{r}*Dash_MVA+MOIC!$G$144"
    htm["G221"] = "=G220+G219/Dash_MVA"
    htm["G222"] = "=G220+G217/Dash_MVA"
    for ws, ref in ((moic, "D225"), (htm, "D220"), (htm, "D221"), (htm, "D222")):
        ws[ref] = "k€/MVA"
    for ref in ("D226", "D236"):
        htm[ref] = "Premium k€/MVA"


COD = 332                                                              # first row of the new HTM block


def cod_sale(htm):
    """§11.2: the buyer earns the unlevered discount rate (Inputs E224) on the post-COD unlevered flows; FEP sells at
    31-Dec of the COD year, keeps the COD-year flows and repays the debt outstanding. Styles follow HTM block 16."""
    cod = "YEAR(Inputs!$E$155)"
    rows = [  # label, unit, G result, timeline formula ({c} = column), G style model
        ("Post-COD unlevered post-tax cash flow", "k€", "=SUM(H{r}:AQ{r})", "={c}110*({c}$5>" + cod + ")", "G322"),
        ("COD sale price (buyer NPV = 0 at Inputs E224), 31-Dec of the COD year", "k€",
         f"=SUMPRODUCT(H{COD + 1}:AQ{COD + 1},H12:AQ12)/INDEX(H12:AQ12,MATCH({cod},H5:AQ5,0))", None, "G322"),
        ("COD sale price per MW built", "k€/MW", f"=G{COD + 2}/In_Wind_Built", None, "G322"),
        ("COD sale price per MVA", "k€/MVA", f"=G{COD + 2}/Dash_MVA", None, "G322"),
        ("Debt repaid at the sale (closing balance, COD year)", "k€", f"=INDEX(H84:AQ84,MATCH({cod},H5:AQ5,0))", None,
         "G322"),
        ("FEP cash flow – sale at COD", "k€", None,
         "={c}123*({c}$5<=" + cod + f")+($G${COD + 2}-$G${COD + 5})*" + "({c}$5=" + cod + ")", "G322"),
        ("FEP IRR – sale at COD", "%", f'=IFERROR(IRR(H{COD + 6}:AQ{COD + 6}),"n/a")', None, "G326"),
        ("FEP profit – sale at COD", "k€", f"=SUM(H{COD + 6}:AQ{COD + 6})", None, "G322"),
        ("FEP MOIC – sale at COD", "x", f'=SUMIF(H{COD + 6}:AQ{COD + 6},">0")/-SUMIF(H{COD + 6}:AQ{COD + 6},"<0")', None,
         "G329")]
    htm[f"C{COD}"] = "17. Sale at COD – buyer earns the unlevered discount rate"
    htm[f"C{COD}"]._style = copy(htm["C309"]._style)
    for r, (label, unit, g, line, model) in enumerate(rows, COD + 1):
        for col, v, m in (("C", label, "C322"), ("D", unit, "D322"), ("G", g, model)):
            htm[f"{col}{r}"] = v.format(r=r) if col == "G" and v else v
            htm[f"{col}{r}"]._style = copy(htm[m]._style)
        for i in range(8, 44) if line else ():
            c = L(i)
            htm[f"{c}{r}"] = line.replace("{c}", c)
            htm[f"{c}{r}"]._style = copy(htm[f"{c}322"]._style)


def deposits_check(dep):
    """§8.2: E34 checks the schedule itself (the source totals are a regression test in regression.py).
    Payable is compared in k€ (O19, the flows included in FEP cash flows), the unit of MOIC G40."""
    rows = {28: ("Gross total – schedule", "mHUF", "=H19"),
            29: ("Gross total – sum of the item rows", "mHUF", "=SUM(H7:H18)"),
            30: ("Difference – gross", "mHUF", None),
            31: ("Payable included in FEP cash flows – schedule", "k€", "=O19"),
            32: ("Payable flows in MOIC (MOIC G40)", "k€", "=MOIC!$G$40"),
            33: ("Difference – payable", "k€", None),
            34: ("Internal check", "Status", None)}
    for r, (label, unit, f) in rows.items():
        dep[f"C{r}"], dep[f"D{r}"] = label, unit
        if f:
            dep[f"E{r}"] = f


# =========================================================================================== §5 cases
def cases(wb):
    """§8.3: two cases. The CN turbine case moves from column I to H (H's own values go), I:P are cleared and hidden.
    Formula cells in H (case number, `=H$253` links) stay as they are."""
    ws = wb["Inputs"]
    for r in range(7, U0):
        h, i = ws[f"H{r}"], ws[f"I{r}"]
        if not str(text(h.value)).startswith("="):
            h.value = text(i.value) if not str(text(i.value)).startswith("=") else None
    for row in ws.iter_rows(min_row=7, max_row=U0 - 1, min_col=9, max_col=16):    # I:P spare cases
        for c in row:
            c.value = None
    g = ws.column_dimensions["G"]                                     # the source has one <col> for G:P
    g.max = 8
    ws.column_dimensions["I"] = ColumnDimension(ws, index="I", width=g.width, hidden=True, min=9, max=16)
    define(wb, "Scenario_Names", "Inputs!$G$7:$H$7")


# =========================================================================================== §6 Dashboard
class Sheet:
    """Row writer for Dashboard: C label | D unit | E F G values (no notes column). `rows` maps keys to rows."""

    def __init__(self, ws):
        self.ws, self.r, self.rows = ws, 4, {}

    def bar(self, title):
        self.r += 1
        _bar(self.ws, self.r, title, 7)
        self.r += 1

    def header(self, *cols):
        for c, h in zip("EFG", cols):
            if h:
                apply(self.ws[f"{c}{self.r}"], "header").value = h
        self.r += 1

    def line(self, key, label, unit, e=None, f=None, g=None, fmt="money1", kind="calc", bold=False):
        r = self.rows[key] = self.r
        ws = self.ws
        apply(ws.cell(r, 3, label), "label")
        apply(ws.cell(r, 4, unit), "unit")
        for col, v in zip((5, 6, 7), (e, f, g)):
            if isinstance(v, str):
                v = v.format(**{k: f"$E${x}" for k, x in self.rows.items()}, r=r)
            if v is not None:
                _write(ws, r, col, v, "total" if bold else kind, fmt)
        if bold:
            ws.cell(r, 3).font = Font(name=FONT, size=SIZE, bold=True)
        self.r += 1
        return r


def band(k, w=None, col="E"):
    """Weighted points of the offer in `col` (lookup of projects/bora_bidding/build.py)."""
    x = f"{col}{{r}}"
    return f"=INDEX(Osz{k}_Pts,1+SUMPRODUCT((Osz{k}_LB<{x})+(Osz{k}_LB={x})*Osz{k}_Incl))*Osz{w or k}_W"


def first_year(row, scale="", sign=""):
    return f"={sign}SUMIF(HTM!$H$5:$AQ$5,{{yr}},HTM!$H${row}:$AQ${row}){scale}"


PER_MW = "=E{r}*Thousand/{built}"
PER_MVA = "=E{r}*Thousand/{mva}"

# Section 5 (Live_Results): key, label, unit, formula, format
RETURNS = [                                 # key, label, unit, E formula, format[, F, G]; label only = sub-header
    ("proj_irr_pre", "Project IRR, unlevered pre-tax", "%", "=HTM!$G$134", "pct"),
    ("proj_irr", "Project IRR, unlevered post-tax", "%", "=HTM!$G$135", "pct"),
    ("proj_npv", "Project NPV, unlevered post-tax", "€ m", "=HTM!$G$136/Thousand", "money1"),
    ("eq_irr", "Equity IRR, levered post-tax (FEP, 100 %)", "%", "=HTM!$G$137", "pct"),
    ("fep_npv", "FEP NPV, build and operate", "€ m", "=HTM!$G$140/Thousand", "money1"),
    ("fep_moic", "FEP MOIC, build and operate", "x", "=HTM!$G$147", "mult"),
    ("fep_peak", "FEP peak exposure, build and operate", "€ m", "=HTM!$G$146/Thousand", "money1"),
    ("dscr", "Minimum DSCR", "x", "=HTM!$G$126", "mult"),
    ("debt", "Debt at COD", "€ m", "=HTM!$G$157/Thousand", "money1"),
    ("gearing", "Debt at COD ÷ all-in cost to COD", "%", "={debt}/{allin}", "pct"),
    ("payback", "Payback year (FEP equity)", "year", "=HTM!$G$148", "int"),
    ("sub_cod", "Buyer at 8 % – FEP sells at COD (31-Dec of the COD year)"),
    ("cod_price", "COD sale price (€ m | k€/MW built | k€/MVA)", "€ m | k€", "=HTM!$G$334/Thousand", "money1",
     "=HTM!$G$335", "=HTM!$G$336"),
    ("cod_irr", "FEP IRR – sale at COD", "%", "=HTM!$G$339", "pct"),
    ("cod_profit", "FEP profit – sale at COD", "€ m", "=HTM!$G$340/Thousand", "money1"),
    ("cod_moic", "FEP MOIC – sale at COD", "x", "=HTM!$G$341", "mult"),
    ("sub_rtb", "Buyer at 8 % – FEP sells at RtB (price Inputs E141)"),
    ("buyer_irr", "Buyer IRR, unlevered post-tax, at the RtB price", "%", "=HTM!$G$215", "pct"),
    ("max_rtb", "Implied maximum RtB premium (buyer NPV = 0 at 8 %)", "k€/MVA", "=HTM!$G$221", "money"),
    ("rtb_irr", "RtB: FEP IRR post-tax", "%", "=MOIC!$G$187", "pct"),
    ("rtb_profit", "RtB: FEP net profit post-tax", "€ m", "=MOIC!$G$182/Thousand", "money2"),
    ("rtb_moic", "RtB: FEP MOIC", "x", "=MOIC!$G$186", "mult"),
    ("rtb_npv", "RtB: FEP NPV", "€ m", "=MOIC!$G$192/Thousand", "money2"),
    ("ares", "RtB: A.RES net profit post-tax", "€ m", "=MOIC!$G$220/Thousand", "money2"),
]


def summary(wb, lv):
    ws = wb.create_sheet("Dashboard", 1)
    ws.sheet_properties.tabColor = "00563C"
    ws.sheet_view.showGridLines = False
    _columns(ws, {"A": 3, "B": 3, "C": 46, "D": 11, "E": 13, "F": 13, "G": 13})
    ws.column_dimensions.group("H", "XFD", hidden=True)
    ws["B2"] = '=Inputs!$A$1&" – Dashboard"'
    _banner(ws, 2, 7)
    apply(ws.cell(3, 3, "Master check (section 8)"), "label").font = Font(name=FONT, size=SIZE, bold=True)
    _write(ws, 3, 5, _status("Master_Check"), "total")
    _ok_format(ws, "E3")
    s = Sheet(ws)

    # 1. Bid strategy: the levers are blue inputs here (names Dash_*); points next to each scored item.
    s.bar("1. Bid strategy and tender points (levers)")
    s.header("Offer", "Points", "Scored")
    lv = {**LEVERS, **lv}

    def lever(key, name, label, unit, f=None, fmt="flag"):
        r = s.line(key, label, unit, lv[name], f, fmt=fmt, kind="input")
        define(wb, name, f"Dashboard!$E${r}")

    lever("mva", "Dash_MVA", "Connection capacity (100 MVA per bid)", "MVA", fmt="money")
    s.line("bids", "Number of bids", "#", "=In_Bids", fmt="int")
    lever("ob", "Dash_Overbuild", "Overbuild (ÖSZ-3)", "% of MVA", band(3), fmt="pct")
    lever("bess", "Dash_BESS_Share", "BESS power (ÖSZ-4); G = ratio scored", "% of MVA", band(3, 4, "G"), fmt="pct")
    _write(ws, s.rows["bess"], 7, f"=E{s.rows['bess']}/(1+$E${s.rows['ob']}*Dash_OSZ4_Basis)", fmt="pct")
    lever("basis", "Dash_OSZ4_Basis", "ÖSZ-4 basis (0 = MVA, 1 = installed wind)", "flag")
    s.line("hours", "BESS duration", "h", "=Inputs!$E$21", fmt="money1")
    lever("bess_capex", "Dash_BESS_Capex_MWh", "BESS capex", "k€/MWh", fmt="money")
    lever("pb", "Dash_PB_Mult", "Performance security multiple (ÖSZ-6)", "x", band(6), fmt="mult")
    s.line("mun", "Municipal one-off (ÖSZ-1)", "k€/MVA", "=Inputs!$E$233", band(1), fmt="money")
    s.line("com", "Community contribution (ÖSZ-2)", "€/MWh", "=Inputs!$E$234", band(2), fmt="money2")
    for key, label, name, pts, w in (("partial", "Partial award accepted (ÖSZ-5)", "Dash_Partial", "Osz5_Pts", "Osz5_W"),
                                     ("brown", "Brownfield site (ÖSZ-7)", "Dash_Brownfield", "Osz7_Pts", "Osz7_W"),
                                     ("shared", "Shared connection (ÖSZ-8)", "Dash_Shared", "Osz8_Pts", "Osz8_W")):
        lever(key, name, label, "flag", f"=E{{r}}*{pts}*{w}")
    for key, label, name, pts in (("eu_t", "EU turbines, at least 80 % (ÖSZ-9)", "Dash_EU_Turbine", "Osz9a_Pts"),
                                  ("eu_tr", "EU transformers, at least 80 % (ÖSZ-9)", "Dash_EU_Transformer", "Osz9b_Pts")):
        r = s.line(key, label, "flag", '=N(Scenario_Name="EU turbine")', f"=E{{r}}*{pts}*Osz9_W", fmt="flag")
        define(wb, name, f"Dashboard!$E${r}")
    lever("short", "Dash_Build_Shortfall", "Wind built below the bid (no sanction up to 10 %)", "%", fmt="pct")
    lever("loss_in", "Dash_Bottleneck_Loss", "Energy lost on overbuilt MW (grid bottleneck)", "%", fmt="pct")
    for k in ("ob", "bess", "pb", "mun", "com", "partial", "brown", "shared", "eu_t", "eu_tr"):
        ws[f"F{s.rows[k]}"].number_format = "0.00"
    case = wb["Inputs"]["G7"].value
    sel = s.line("route", "Turbine route (case)", "case", case, fmt="@", kind="input")
    ws[f"E{sel}"].alignment = Alignment(horizontal="left")
    define(wb, "Scenario_Name", f"Dashboard!$E${sel}")
    _validate(ws, f"E{sel}", "Scenario_Names")
    s.line("pts", "Total points", "points", None,
           f"=ROUND(SUM(F{s.rows['ob']}:F{s.rows['eu_tr']}),Points_Decimals)", fmt="0.00", bold=True)

    # 2. Capacity bases.
    s.bar("2. Capacity and basis")
    s.line("inst", "Wind committed (bid)", "MW", "=Inputs!$E$18", fmt="money")
    s.line("built", "Wind built (bid less 10 %, no sanction)", "MW", "=In_Wind_Built", fmt="money")
    s.line("bmw", "BESS power", "MW", "=Inputs!$E$20", fmt="money")
    s.line("bmwh", "BESS energy", "MWh", "=Inputs!$E$22", fmt="money")
    s.line("over", "Wind built above connection capacity", "MW", "={built}-{mva}", fmt="money")
    s.line("loss", "Bottleneck loss applied to generation", "%", "=1-In_Bottleneck_Factor", fmt="0.00%")
    apply(ws.cell(s.r, 3, "k€/MW = per built wind MW; k€/MVA = per connection MVA"), "unit")
    s.r += 1

    # 3. Build cost to COD.
    s.bar("3. Build cost to COD")
    s.header("€ m", "k€/MW built", "k€/MVA")
    items = [("turb", "Turbines", "=Inputs!$E$162*{built}/Thousand"),
             ("inst_c", "Installation", "=Inputs!$E$163*{built}/Thousand"),
             ("found", "Foundations", "=Inputs!$E$164*{built}/Thousand"),
             ("civil", "Balance of plant, civil", "=Inputs!$E$165*{built}/Thousand"),
             ("elec", "Balance of plant, electrical", "=Inputs!$E$166*{built}/Thousand"),
             ("subst", "Substation and 132 kV bay", "=Inputs!$E$167/Thousand"),
             ("owner", "Owner's costs", "=Inputs!$E$169*{built}/Thousand"),
             ("land", "Land and easements", "=Inputs!$E$170*{built}/Thousand"),
             ("cont", "Contingency", "=SUM(E{turb_r}:E{land_r})*Inputs!$E$171"),
             ("bess_c", "BESS", "=Inputs!$E$168*{bmw}/Thousand"),
             ("fee_ind", "Connection fee, indirect (net of TSO security)", "=Deposits!$M$18/Thousand"),
             ("fee_dir", "Connection fee, direct", "=Deposits!$M$12/Thousand"),
             ("muni", "Municipal one-off", "=-HTM!$G$29/Thousand"),
             ("idc", "Interest during construction", "=HTM!$G$79/Thousand"),
             ("fees", "Debt fees (upfront, commitment)", "=-HTM!$G$85/Thousand")]
    for key, label, e in items:
        if key == "cont":
            e = e.replace("{turb_r}", str(s.rows["turb"])).replace("{land_r}", str(s.rows["land"]))
        s.line(key, label, "€ m", e, PER_MW, PER_MVA)
    s.line("allin", "All-in cost to COD", "€ m", f"=SUM(E{s.rows['turb']}:E{s.rows['fees']})", PER_MW, PER_MVA,
           bold=True)
    for k in [i[0] for i in items] + ["allin"]:
        for c in "FG":
            ws[f"{c}{s.rows[k]}"].number_format = _fmt("money")

    # 4. Generation and operations.
    s.bar("4. Generation and operations")
    s.header("Value", "per MW built", "per MVA")
    s.line("yr", "First full operating year", "year", "=Inputs!$E$156", fmt="int")
    s.line("gen", "Net generation, first full year", "GWh", first_year(38), fmt="money")
    s.line("eqh", "Equivalent full-load hours, first full year", "h/yr", None, "={gen}*Thousand/{built}",
           "={gen}*Thousand/{mva}", fmt="money")
    s.line("ncf", "Net capacity factor, first full year", "%", None, f"=F{s.rows['eqh']}/Hours_In_Year",
           f"=G{s.rows['eqh']}/Hours_In_Year", fmt="pct")
    s.line("lost", "Energy lost to the bottleneck, first full year", "GWh",
           "={gen}*(1-In_Bottleneck_Factor)/In_Bottleneck_Factor", fmt="money1")
    s.line("price", "Achieved price, first full year", "€/MWh", first_year(49), fmt="money1")
    s.line("rev", "Revenue, first full year", "€ m", first_year(54, "/Thousand"), fmt="money1")
    s.line("ebitda", "EBITDA (NOI), first full year", "€ m", first_year(70, "/Thousand"), fmt="money1")
    s.line("margin", "EBITDA margin, first full year", "%", first_year(71), fmt="pct")
    s.line("opex", "Operating costs, first full year (k€ per MW / MVA)", "€ m | k€", first_year(69, "/Thousand", "-"),
           PER_MW, PER_MVA, fmt="money1")
    for c in "FG":
        ws[f"{c}{s.rows['opex']}"].number_format = _fmt("money1")
    s.line("life_ebitda", "EBITDA (NOI), lifetime", "€ m", "=HTM!$G$70/Thousand", fmt="money1")
    s.line("life_gen", "Lifetime net generation", "GWh", "=HTM!$G$38", fmt="money")
    s.line("lcoe", "LCOE", "€/MWh", "=HTM!$G$156", fmt="money1")

    # 5. Returns = Live_Results (one column); section 7 holds the stored copies.
    s.bar("5. Returns, live case (Live_Results)")
    first = s.r
    live = [("points", "Tender points", "points", f"=F{s.rows['pts']}", "0.00")] + RETURNS
    for key, label, *rest in live:
        if not rest:                                                   # sub-header inside Live_Results (E empty)
            _bar(ws, s.r, label, 7, "subheader")
            s.r += 1
            continue
        unit, e, fmt, *fg = rest
        s.line(key, label, unit, e, *fg, fmt=fmt)
        define(wb, f"Out_{key}", f"Dashboard!$E${s.rows[key]}")
    define(wb, "Live_Results", f"Dashboard!$E${first}:$E${s.r - 1}")

    # 6. Securities and fees (Deposits, tender currency).
    s.bar("6. Securities and fees (Deposits)")
    for key, label, ref in (("part", "Participation fee", "H7"), ("bid", "Bid security", "H8"),
                            ("perf", "Performance security (gross; bid credit in Deposits J10)", "H10"),
                            ("tso", "TSO connection-request security", "H11"),
                            ("ind", "Connection fee, indirect (before TSO credit)", "H18"),
                            ("dir", "Connection fee, direct", "H12")):
        s.line(key, label, "mHUF", f"=Deposits!${ref[0]}${ref[1:]}", fmt="money")
    s.line("gross", "Total gross", "mHUF", "=Deposits!$H$19", fmt="money", bold=True)
    s.line("payable", "Total payable after netting", "mHUF", "=Deposits!$J$19", fmt="money")
    s.line("cash", "Of which cash, in FEP cash flows", "€ m", "=Deposits!$E$26/Thousand", fmt="money1")
    s.line("guar", "Of which bank guarantee, in FEP cash flows", "€ m", "=Deposits!$E$27/Thousand", fmt="money1")

    # 7. Stored results by case, written by RunAllScenarios (purple).
    s.bar("7. Stored results by case (Stored_Results, written by RunAllScenarios)")
    for i in range(2):
        _write(ws, s.r, 5 + i, f"=Inputs!{L(7 + i)}$7", "header")
        ws.cell(s.r, 5 + i).alignment = Alignment(horizontal="right", wrap_text=True)
    s.r += 1
    top = s.r
    for key, label, *rest in live:
        if not rest:
            _bar(ws, s.r, label, 7, "subheader")
        else:
            apply(ws.cell(s.r, 3, label), "label")
            apply(ws.cell(s.r, 4, rest[0]), "unit")
            for col in (5, 6):
                apply(ws.cell(s.r, col), "macro", rest[2])
        s.r += 1
    define(wb, "Stored_Results", f"Dashboard!$E${top}:$F${s.r - 1}")

    # 8. Checks: one row per test (E = 1 if failing), Master_Check = number of failing rows.
    s.bar("8. Checks")
    s.header("Fails", "Status")
    checks = [("c_moic", "MOIC engine checks (MOIC G270)", '=N(MOIC!$G$270<>"ALL OK")'),
              ("c_htm", "HTM engine checks (HTM G172)", '=N(HTM!$G$172<>"ALL OK")'),
              ("c_dep", "Deposits reconciliation (Deposits E34)", '=N(Deposits!$E$34<>"OK")'),
              ("c_capex", "Capex breakdown = HTM capex + IDC + fees",
               "=N(ABS({allin}-(HTM!$G$79-HTM!$G$30-HTM!$G$85)/Thousand)>Tolerance)"),
              ("c_stored", "Stored results = live case", STORED_CHECK[1]),
              ("c_osz1", "ÖSZ-1 offer within the offerable range", "=N((E{m}<>0)*(E{m}<Osz1_Min)+(E{m}>Osz1_Max)>0)"),
              ("c_osz2", "ÖSZ-2 offer within the offerable range", "=N((E{c}<>0)*(E{c}<Osz2_Min)+(E{c}>Osz2_Max)>0)"),
              ("c_osz4", "ÖSZ-4 BESS ratio at least the minimum", "=N((G{b}>0)*(G{b}<Osz4_Min)>0)"),
              ("c_cap", "Capacity: BESS MW ≤ wind MW, duration ≥ minimum",
               "=N(({bmw}>{inst})+({hours}<In_Min_Hours)>0)")]
    for key, label, f in checks:
        for token, row in (("{m}", "mun"), ("{c}", "com"), ("{b}", "bess")):
            f = f.replace(token, str(s.rows[row]))
        r = s.line(key, label, "#", f, fmt="int")
        _write(ws, r, 6, _status(f"E{r}"))
    first, last = s.rows["c_moic"], s.rows["c_cap"]
    r = s.line("master", "Master check (number of failing rows)", "#", f"=SUM(E{first}:E{last})", fmt="int",
               bold=True)
    _write(ws, r, 6, _status(f"E{r}"), "total")
    define(wb, "Master_Check", f"Dashboard!$E${r}")
    _ok_format(ws, f"F{first}:F{r}")

    # 9. Key assumptions (links to the engine inputs).
    s.bar("9. Key assumptions, live case")
    for key, label, unit, ref, fmt in (("a_turb", "Turbine supply", "k€/MW", "E162", "money"),
                                       ("a_ppa", "PPA price, first operating year", "€/MWh", "E191", "money1"),
                                       ("a_tenor", "PPA tenor", "years", "E193", "int"),
                                       ("a_merch", "Merchant share of output", "%", "E190", "pct"),
                                       ("a_gear", "Gearing (debt on capex incl. connection)", "%", "E217", "pct"),
                                       ("a_rate", "Debt interest rate", "%", "E218", "pct"),
                                       ("a_disc", "Discount rate (NPV)", "%", "E253", "pct"),
                                       ("a_fx", "FX", "HUF/EUR", "E11", "money"),
                                       ("a_rtb", "RtB date", "date", "E56", "date"),
                                       ("a_sale", "Sale date (RtB route)", "date", "E58", "date"),
                                       ("a_cod", "COD", "date", "E155", "date"),
                                       ("a_life", "Operating life", "years", "E157", "int"),
                                       ("a_exit", "A.RES bought out at RtB in HTM (1 = yes)", "flag", "E228", "flag"),
                                       ("a_dsra", "DSRA", "months", "E268", "int"),
                                       ("a_sweep", "Cash sweep", "%", "E269", "pct")):
        s.line(key, label, unit, f"=Inputs!${ref[0]}${ref[1:]}", fmt=fmt)
    s.line("a_scale", "Lump sums scaled from the 200 MVA source (factor)", "x", "=In_Scale", fmt="0.000")
    _end(ws, s.r + 1, 7)
    return s.rows


# =========================================================================================== §9.4 Cash_Flow
CF_FIRST, CF_YEARS = 9, 36                 # I.. timeline = HTM H:AQ (2026-2061)
CF_ROWS = [                                # key, label, HTM row(s) linked
    ("spend", "Total spending", "24+30+117"), ("rev", "Revenue", "54"), ("opex", "Operating costs", "69"),
    ("ebitda", "EBITDA (NOI)", "70"), ("tax", "Tax", "100"), ("ucf", "Unlevered post-tax cash flow", "110"),
    ("draw", "Debt drawdown", "77"), ("ds", "Debt service", "119"), ("ecf", "Equity cash flow", "123"),
    ("ecum", "Cumulative equity cash flow", "124"),
]
CF_BOLD, CF_NO_TOTAL = {"spend", "ebitda", "ucf", "ecf"}, {"ecum"}


def cash_flow(wb):
    """The HTM 2026-2061 grid as links (SPEC §9.4, §10.3): one formula per row copied across, E total € m,
    F per built MW, G per MVA. Total spending = development phase + capex + fees paid by equity (HTM 24 + 30 + 117);
    drawdown points at its source HTM 77 (HTM 116 is only a link)."""
    ws = wb.create_sheet("Cash_Flow", 0)
    ws.sheet_view.showGridLines = False
    last = CF_FIRST + CF_YEARS - 1
    cols = range(CF_FIRST, last + 1)
    _columns(ws, {"A": 3, "B": 3, "C": 34, "D": 8, "E": 11, "F": 11, "G": 11, "H": 2,
                  **{L(c): 9 for c in cols}})
    ws.column_dimensions.group(L(last + 1), "XFD", hidden=True)
    ws["B2"] = '=Inputs!$A$1&" – Cash flow, build and operate (HTM, EUR k)"'
    _banner(ws, 2, last)
    for col, h in zip("EFG", ("Total € m", "k€/MW built", "k€/MVA")):
        apply(ws[f"{col}4"], "header").value = h
    apply(ws.cell(5, 3, "Year"), "label")
    apply(ws.cell(5, 4, "year"), "unit")
    for c in cols:
        _write(ws, 5, c, f"=HTM!{L(c - 1)}$5", "header", "0")
    ws.freeze_panes = "I6"
    for r, (key, label, src) in enumerate(CF_ROWS, 7):
        kind = "total" if key in CF_BOLD else "calc"
        apply(ws.cell(r, 3, label), "label")
        if key in CF_BOLD:
            ws.cell(r, 3).font = Font(name=FONT, size=SIZE, bold=True)
        apply(ws.cell(r, 4, "EUR k"), "unit")
        for c in cols:
            _write(ws, r, c, "=" + "+".join(f"HTM!{L(c - 1)}{x}" for x in src.split("+")), kind, "money")
        if key not in CF_NO_TOTAL:
            _cf_totals(ws, r, last, kind)
    _end(ws, 7 + len(CF_ROWS) + 1, last)


def _cf_totals(ws, r, last, kind):
    _write(ws, r, 5, f"=SUM({L(CF_FIRST)}{r}:{L(last)}{r})/Thousand", kind, "money1")
    _write(ws, r, 6, f"=E{r}*Thousand/In_Wind_Built", kind, "money")
    _write(ws, r, 7, f"=E{r}*Thousand/Dash_MVA", kind, "money")


# =========================================================================================== Change_Log
def change_log(wb):
    ws, r = wb["Change_Log"], 131
    for col in "BCDEFGHIJ":
        ws[f"{col}{r}"]._style = copy(ws[f"{col}{r - 1}"]._style)
    for col, v in zip("BCDEFGHIJ", (
            f"=B{r - 1}+1", date(2026, 10, 22), "Dashboard sheet", "Cash_Flow; Dashboard; Inputs; Deposits; HTM; MOIC",
            "Inputs E6, E16, E20, E90, E168, G:H case cells, rows 300+; Deposits G11, G12, G18, C28:E34; HTM E36, "
            "H29:AQ29, E227:E243, G221:G222; MOIC G143, H226:N226",
            "Dashboard rebuilt as one IC page and Cash_Flow (HTM annual grid) added; Cover and the old "
            "Control_Panel, Dashboard, Bora, Run_Map deleted (links → values, "
            "G292 formula moved from Control_Panel E22). Levers on Dashboard, points bands on Inputs block U; capacity, BESS, PB "
            "and RtB price on the MVA basis; two cases (EU, CN turbine); Deposits E34 is an internal check; "
            "Module1/Module2 replaced by Scenarios.bas.",
            "200 MW, BESS 100 MW, PB ×1.5, merchant base, 950 k€/MW, 3 cases",
            "200 MVA / 280 MW installed, BESS 80 MW, PB ×1.0, PPA 90 €/MWh 30 yrs, 950 / 500 k€/MW",
            "User 22-Oct-2026; projects/bora_v5_summary/SPEC.md")):
        ws[f"{col}{r}"] = v


# =========================================================================================== build
def general(path, sheet, cells, like):
    """§8.6: give `cells` the style index of `like` (a General cell of the same kind) inside the saved file.
    LibreOffice infers a 0.00 % format for these sums on save; Excel would then show 3000.00 %."""
    path = Path(path)
    with zipfile.ZipFile(path) as z:
        parts = {n: z.read(n) for n in z.namelist()}
    wbx, rels = parts["xl/workbook.xml"].decode(), parts["xl/_rels/workbook.xml.rels"].decode()
    rid = re.search(rf'<sheet name="{sheet}"[^>]*r:id="(\w+)"', wbx).group(1)
    target = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"|Target="([^"]+)"[^>]*Id="{rid}"', rels)
    xml_name = "xl/" + (target.group(1) or target.group(2)).lstrip("/").removeprefix("xl/")
    xml = parts[xml_name].decode()
    style = re.search(rf'<c r="{like}" s="(\d+)"', xml).group(1)
    for c in cells:
        xml, n = re.subn(rf'<c r="{c}" s="\d+"', f'<c r="{c}" s="{style}"', xml)
        assert n == 1, c
    parts[xml_name] = xml.encode()
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        for n, data in parts.items():
            z.writestr(n, data)
    tmp.replace(path)


def edit(levers=NEW):
    """The edited workbook (openpyxl), before packaging."""
    wb = openpyxl.load_workbook(SRC, keep_vba=False)
    cached = openpyxl.load_workbook(SRC, data_only=True)
    add_styles(wb)
    repoint(wb, cached)
    delete_duplicates(wb)
    block_u(wb, levers)
    cases(wb)
    engine(wb, levers)
    cash_flow(wb)
    summary(wb, levers)
    change_log(wb)
    for ws in wb.worksheets:                                   # open on Dashboard
        ws.sheet_view.tabSelected = ws.title == "Dashboard"
    wb.active = wb.sheetnames.index("Dashboard")
    return wb


def build(out=HERE / "out", levers=NEW, name=NAME):
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    xlsx, xlsm = out / f"{name}.xlsx", out / f"{name}.xlsm"
    edit(levers).save(xlsx)
    with tempfile.TemporaryDirectory() as vba:                 # Scenarios.bas only (as is)
        shutil.copy(VBA_DIR / "Scenarios.bas", vba)
        package(xlsx, xlsm, run=["Scenarios.RunAllScenarios"], vba_dir=Path(vba))
    general(xlsm, "HTM", ("G7", "G10", "G14"), like="G8")
    rc = recalc(xlsm, out / f"{name}_recalc.xlsx")
    return {"xlsx": xlsx, "xlsm": xlsm, "recalc": rc}
