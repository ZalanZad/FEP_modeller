"""Bora v5 Summary – one Dashboard sheet on the colleague's MOIC / HTM workbook. Built to SPEC.md.

Loads reference/HUN_Wind_MOIC_HTM_IC_Bora_v5_20261022.xlsm (read-only), re-points every link to the duplicate
sheets (Control_Panel, the old Dashboard, Bora, Run_Map) and deletes them, adds Inputs block U (points bands and constants,
rows 300+), edits only the engine cells of SPEC §4 and §8, adds Dashboard after Cover (levers in section 1, two cases)
and packages the workbook with
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
DELETE = ("Control_Panel", "Dashboard", "Bora", "Run_Map")
DELETED_REF = re.compile(r"(Control_Panel|Dashboard|Bora|Run_Map)'?!")
U0 = 300                                                   # first row of Inputs block U
BAND_COL = 17                                              # Q: points bands (J:P are hidden spare cases)

# Levers: the Dashboard section 1 inputs (Dash_*) and the case cells of SPEC §4 / §8.3 (Inputs G = EU, H = CN turbine).
NEW = {"Dash_Overbuild": 0.40, "Dash_BESS_Share": 0.40, "Dash_PB_Mult": 1.0,
       "cases": {"G7": "EU turbine", "H7": "CN turbine",
                 "G8": "EU turbines 600 k€/MW, 60 % gearing, 90 €/MWh PPA on all output for 30 years",
                 "H8": "As EU turbine, but CN turbines 500 k€/MW and 50 % gearing",
                 "G78": 0, "G233": 50, "G234": 3, "G162": 600, "H162": 500, "G190": 0, "H190": None,
                 "G193": 30, "G217": 0.60, "H217": 0.50, "G11": 365, "G268": 0, "G269": 0}}
SOURCE = {"Dash_Overbuild": 0.0, "Dash_BESS_Share": 0.50, "Dash_PB_Mult": 1.5,                # the v5 file as received:
          "cases": {"G7": "EU turbine", "H7": "CN turbine",                         # its cases 2 and 3
                    "G8": "v5 case 2 (regression)", "H8": "v5 case 3 (regression)",
                    "G78": 0.05, "G233": 20, "G234": 2, "G162": 950, "H162": 500, "G190": 0, "H190": 1,
                    "G193": 10, "G217": 0.50, "H217": 0, "G11": 370, "G268": 6, "G269": 0.5}}
# Dashboard section 1 levers that are the same in both builds: name -> value
LEVERS = {"Dash_MVA_Per_Bid": 100, "Dash_BESS_Capex_MWh": 100, "Dash_Bottleneck_Loss": 0.05, "Dash_OSZ4_Basis": 0,
          "Dash_Partial": 1, "Dash_Brownfield": 0, "Dash_Shared": 1, "Dash_EU_Turbine": 1, "Dash_EU_Transformer": 1}

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


def u_rows():
    """Block U scalars: (name, label, unit, value or formula, note, number format). The levers are on Dashboard."""
    return [
        ("In_Bottleneck_Factor", "Generation kept after the bottleneck loss", "x",
         "=1-Dash_Bottleneck_Loss*(E18-Dash_MVA_Per_Bid*In_Bids)/E18", "Used in HTM E36", "0.0000"),
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
def block_u(wb):
    ws = wb["Inputs"]
    label, unit, note, calc, blue = ws["C268"], ws["D268"], ws["F268"], ws["E154"], ws["G268"]
    like(ws.cell(U0, 2, "x"), ws["B267"])
    like(ws.cell(U0, 3, "U. Tender points tables and constants (levers Dash_* are on the Dashboard)"), ws["C267"])
    r = U0 + 1
    for name, lab, un, value, src, fmt in u_rows():
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
    inp["E16"] = "=Dash_MVA_Per_Bid*In_Bids*(1+Dash_Overbuild)"          # installed wind = committed
    inp["E20"] = "=Dash_BESS_Share*Dash_MVA_Per_Bid*In_Bids"
    inp["E168"] = "=Dash_BESS_Capex_MWh*E21"                            # 100 k€/MWh × hours
    inp["E90"] = "=Dash_PB_Mult"
    for r in (16, 20):                                                 # stale case values, now unused
        inp[f"G{r}"] = None
        inp[f"F{r}"] = f"Driven by the Dashboard levers (E{r})"
    for ref, v in lv["cases"].items():                                 # case values (blue inputs already)
        inp[ref] = v
    for ref in ("G11", "G12", "G18"):                                  # TSO security and connection fees per MVA
        dep[ref] = "=Dash_MVA_Per_Bid*In_Bids*Inputs!$E$17"
    for col in range(8, 44):                                           # H:AQ municipal one-off per MVA
        c = L(col)
        htm[f"{c}29"] = f"=-IF({c}5=YEAR(Inputs!$E$155),Inputs!$E$233*Dash_MVA_Per_Bid*In_Bids,0)"
    htm["E36"] = "=Inputs!$E$179*(1-Inputs!$E$180)*(1-Inputs!$E$181)*Inputs!$E$183*In_Bottleneck_Factor"
    moic["G143"] = "=Inputs!$E$141*Dash_MVA_Per_Bid*In_Bids"            # RtB price per MVA
    rtb_basis(wb)
    deposits_check(dep)


def rtb_basis(wb):
    """§8.1: the RtB premium sensitivities and the implied maximum premium on the MVA basis (k€/MVA)."""
    moic, htm = wb["MOIC"], wb["HTM"]
    for col in "HIJKLMN":
        moic[f"{col}226"] = f"={col}225*Dash_MVA_Per_Bid*In_Bids"
    for r in list(range(227, 234)) + list(range(237, 244)):
        htm[f"E{r}"] = f"=D{r}*Dash_MVA_Per_Bid*In_Bids+MOIC!$G$144"
    htm["G221"] = "=G220+G219/(Dash_MVA_Per_Bid*In_Bids)"
    htm["G222"] = "=G220+G217/(Dash_MVA_Per_Bid*In_Bids)"
    for ws, ref in ((moic, "D225"), (htm, "D220"), (htm, "D221"), (htm, "D222")):
        ws[ref] = "k€/MVA"
    for ref in ("D226", "D236"):
        htm[ref] = "Premium k€/MVA"


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
    """Row writer for Dashboard: C label | D unit | E F G values | H note. `rows` maps keys to row numbers."""

    def __init__(self, ws):
        self.ws, self.r, self.rows = ws, 4, {}

    def bar(self, title):
        self.r += 1
        _bar(self.ws, self.r, title, 8)
        self.r += 1

    def header(self, *cols):
        for c, h in zip("EFG", cols):
            if h:
                apply(self.ws[f"{c}{self.r}"], "header").value = h
        self.r += 1

    def line(self, key, label, unit, e=None, f=None, g=None, note=None, fmt="money1", kind="calc", bold=False):
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
        if note:
            apply(ws.cell(r, 8, note), "label")
        self.r += 1
        return r


def band(k, w=None, col="E"):
    """Weighted points of the offer in `col` (lookup of projects/bora_bidding/build.py)."""
    x = f"{col}{{r}}"
    return f"=INDEX(Osz{k}_Pts,1+SUMPRODUCT((Osz{k}_LB<{x})+(Osz{k}_LB={x})*Osz{k}_Incl))*Osz{w or k}_W"


def first_year(row, scale="", sign=""):
    return f"={sign}SUMIF(HTM!$H$5:$AQ$5,{{yr}},HTM!$H${row}:$AQ${row}){scale}"


PER_MW = "=E{r}*Thousand/{inst}"
PER_MVA = "=E{r}*Thousand/{mva}"

# Section 5 (Live_Results): key, label, unit, formula, format, source note
RETURNS = [
    ("proj_irr_pre", "Project IRR, unlevered pre-tax", "%", "=HTM!$G$134", "pct", "HTM G134"),
    ("proj_irr", "Project IRR, unlevered post-tax", "%", "=HTM!$G$135", "pct", "HTM G135"),
    ("proj_npv", "Project NPV, unlevered post-tax", "€ m", "=HTM!$G$136/Thousand", "money1", "HTM G136"),
    ("eq_irr", "Equity IRR, levered post-tax (FEP, 100 %)", "%", "=HTM!$G$137", "pct", "HTM G137 (= G139)"),
    ("fep_npv", "FEP NPV, build and operate", "€ m", "=HTM!$G$140/Thousand", "money1", "HTM G140"),
    ("fep_moic", "FEP MOIC, build and operate", "x", "=HTM!$G$147", "mult", "HTM G147"),
    ("fep_peak", "FEP peak exposure, build and operate", "€ m", "=HTM!$G$146/Thousand", "money1", "HTM G146"),
    ("dscr", "Minimum DSCR", "x", "=HTM!$G$126", "mult", "HTM G126"),
    ("debt", "Debt at COD", "€ m", "=HTM!$G$157/Thousand", "money1", "HTM G157"),
    ("gearing", "Debt at COD ÷ all-in cost to COD", "%", "={debt}/{allin}", "pct", "Rows above"),
    ("payback", "Payback year (FEP equity)", "year", "=HTM!$G$148", "int", "HTM G148"),
    ("rtb_profit", "RtB: FEP net profit post-tax", "€ m", "=MOIC!$G$182/Thousand", "money2", "MOIC G182"),
    ("rtb_irr", "RtB: FEP IRR post-tax", "%", "=MOIC!$G$187", "pct", "MOIC G187"),
    ("rtb_moic", "RtB: FEP MOIC", "x", "=MOIC!$G$186", "mult", "MOIC G186"),
    ("rtb_npv", "RtB: FEP NPV", "€ m", "=MOIC!$G$192/Thousand", "money2", "MOIC G192"),
    ("ares", "RtB: A.RES net profit post-tax", "€ m", "=MOIC!$G$220/Thousand", "money2", "MOIC G220"),
    ("buyer_irr", "Buyer IRR, unlevered post-tax, at the RtB price", "%", "=HTM!$G$215", "pct", "HTM G215"),
    ("max_rtb", "Implied maximum RtB premium (buyer NPV = 0)", "k€/MVA", "=HTM!$G$221", "money", "HTM G221"),
]


def summary(wb, lv):
    ws = wb.create_sheet("Dashboard", 1)
    ws.sheet_properties.tabColor = "00563C"
    ws.sheet_view.showGridLines = False
    _columns(ws, {"A": 3, "B": 3, "C": 46, "D": 11, "E": 13, "F": 13, "G": 13, "H": 44})
    ws.column_dimensions.group("I", "XFD", hidden=True)
    ws["B2"] = '=Cover!$B$2&" – Dashboard"'
    _banner(ws, 2, 8)
    apply(ws.cell(3, 3, "Live case: pick the case in E{sel}; section 7 holds every case (RunAllScenarios)"), "unit")
    s = Sheet(ws)

    # 1. Bid strategy: the levers are blue inputs here (names Dash_*); points next to each scored item.
    s.bar("1. Bid strategy and tender points (levers)")
    s.header("Offer", "Points", "Scored")
    lv = {**LEVERS, **lv}

    def lever(key, name, label, unit, f=None, fmt="flag", note="User 22-Oct-26"):
        r = s.line(key, label, unit, lv[name], f, fmt=fmt, kind="input", note=note)
        define(wb, name, f"Dashboard!$E${r}")

    lever("mva_bid", "Dash_MVA_Per_Bid", "Connection capacity per bid", "MVA", fmt="money", note="Kiírás Annex 1")
    r = s.line("bids", "Number of bids", "#", "=Inputs!$E$82", fmt="int", note="Inputs E82")
    define(wb, "In_Bids", f"Dashboard!$E${r}")
    lever("ob", "Dash_Overbuild", "Overbuild (ÖSZ-3)", "% of MVA", band(3), fmt="pct")
    lever("bess", "Dash_BESS_Share", "BESS power (ÖSZ-4); G = ratio scored", "% of MVA", band(3, 4, "G"), fmt="pct",
          note="User 22-Oct-26; G on the ÖSZ-4 basis below")
    _write(ws, s.rows["bess"], 7, f"=E{s.rows['bess']}/(1+$E${s.rows['ob']}*Dash_OSZ4_Basis)", fmt="pct")
    lever("basis", "Dash_OSZ4_Basis", "ÖSZ-4 basis (0 = MVA, 1 = installed wind)", "flag", note="Meeting file basis")
    s.line("hours", "BESS duration", "h", "=Inputs!$E$21", fmt="money1", note="Inputs E21")
    lever("bess_capex", "Dash_BESS_Capex_MWh", "BESS capex", "k€/MWh", fmt="money", note="User 15-Oct-26 (CLAUDE.md)")
    lever("pb", "Dash_PB_Mult", "Performance security multiple (ÖSZ-6)", "x", band(6), fmt="mult",
          note="User 12-Oct-26")
    s.line("mun", "Municipal one-off (ÖSZ-1)", "k€/MVA", "=Inputs!$E$233", band(1), fmt="money",
           note="Inputs G233; offerable 10–50 k€/MVA")
    s.line("com", "Community contribution (ÖSZ-2)", "€/MWh", "=Inputs!$E$234", band(2), fmt="money2",
           note="Inputs G234; offerable 2–3 €/MWh")
    for key, label, name, pts, w in (("partial", "Partial award accepted (ÖSZ-5)", "Dash_Partial", "Osz5_Pts", "Osz5_W"),
                                     ("brown", "Brownfield site (ÖSZ-7)", "Dash_Brownfield", "Osz7_Pts", "Osz7_W"),
                                     ("shared", "Shared connection (ÖSZ-8)", "Dash_Shared", "Osz8_Pts", "Osz8_W"),
                                     ("eu_t", "EU turbines, at least 80 % (ÖSZ-9)", "Dash_EU_Turbine", "Osz9a_Pts",
                                      "Osz9_W"),
                                     ("eu_tr", "EU transformers, at least 80 % (ÖSZ-9)", "Dash_EU_Transformer",
                                      "Osz9b_Pts", "Osz9_W")):
        lever(key, name, label, "flag", f"=E{{r}}*{pts}*{w}", note="Bid")
    lever("loss_in", "Dash_Bottleneck_Loss", "Energy lost on overbuilt MW (grid bottleneck)", "%", fmt="pct",
          note="Meeting file Bora!N18")
    for k in ("ob", "bess", "pb", "mun", "com", "partial", "brown", "shared", "eu_t", "eu_tr"):
        ws[f"F{s.rows[k]}"].number_format = "0.00"
    case = wb["Inputs"]["G7"].value
    sel = s.line("route", "Turbine route (case)", "case", case, fmt="@", kind="input", note="Cases: Inputs G7:H7")
    ws[f"E{sel}"].alignment = Alignment(horizontal="left")
    define(wb, "Scenario_Name", f"Dashboard!$E${sel}")
    _validate(ws, f"E{sel}", "Scenario_Names")
    ws["C3"] = ws["C3"].value.format(sel=sel)
    s.line("pts", "Total points", "points", None,
           f"=ROUND(SUM(F{s.rows['ob']}:F{s.rows['eu_tr']}),Points_Decimals)", fmt="0.00", bold=True,
           note="Kiírás Annex 3-4")

    # 2. Capacity bases.
    s.bar("2. Capacity and basis, both bids")
    s.line("mva", "Connection capacity", "MVA", "={mva_bid}*{bids}", fmt="money", note="Tender right")
    s.line("inst", "Installed wind", "MW", "=Inputs!$E$18", fmt="money", note="Inputs E18")
    s.line("bmw", "BESS power", "MW", "=Inputs!$E$20", fmt="money", note="Inputs E20")
    s.line("bmwh", "BESS energy", "MWh", "=Inputs!$E$22", fmt="money", note="Inputs E22")
    s.line("over", "Overbuilt wind above connection capacity", "MW", "={inst}-{mva}", fmt="money")
    s.line("loss", "Bottleneck loss applied to generation", "%", "=1-In_Bottleneck_Factor", fmt="0.00%",
           note="Inputs block U (In_Bottleneck_Factor)")
    apply(ws.cell(s.r, 3, "k€/MW = per installed wind MW; k€/MVA = per connection MVA"), "unit")
    s.r += 1

    # 3. Build cost to COD.
    s.bar("3. Build cost to COD, both bids")
    s.header("€ m", "k€/MW inst.", "k€/MVA")
    items = [("turb", "Turbines", "=Inputs!$E$162*{inst}/Thousand", "Inputs E162"),
             ("inst_c", "Installation", "=Inputs!$E$163*{inst}/Thousand", "Inputs E163"),
             ("found", "Foundations", "=Inputs!$E$164*{inst}/Thousand", "Inputs E164"),
             ("civil", "Balance of plant, civil", "=Inputs!$E$165*{inst}/Thousand", "Inputs E165"),
             ("elec", "Balance of plant, electrical", "=Inputs!$E$166*{inst}/Thousand", "Inputs E166"),
             ("subst", "Substation and 132 kV bay", "=Inputs!$E$167/Thousand", "Inputs E167"),
             ("owner", "Owner's costs", "=Inputs!$E$169*{inst}/Thousand", "Inputs E169"),
             ("land", "Land and easements", "=Inputs!$E$170*{inst}/Thousand", "Inputs E170"),
             ("cont", "Contingency", "=SUM(E{turb_r}:E{land_r})*Inputs!$E$171", "Inputs E171"),
             ("bess_c", "BESS", "=Inputs!$E$168*{bmw}/Thousand", "Inputs E168 (k€/MW BESS)"),
             ("fee_ind", "Connection fee, indirect (net of TSO security)", "=Deposits!$M$18/Thousand",
              "Deposits M18"),
             ("fee_dir", "Connection fee, direct", "=Deposits!$M$12/Thousand", "Deposits M12"),
             ("muni", "Municipal one-off", "=-HTM!$G$29/Thousand", "HTM G29"),
             ("idc", "Interest during construction", "=HTM!$G$79/Thousand", "HTM G79"),
             ("fees", "Debt fees (upfront, commitment)", "=-HTM!$G$85/Thousand", "HTM G85")]
    for key, label, e, note in items:
        if key == "cont":
            e = e.replace("{turb_r}", str(s.rows["turb"])).replace("{land_r}", str(s.rows["land"]))
        s.line(key, label, "€ m", e, PER_MW, PER_MVA, note=note)
    s.line("allin", "All-in cost to COD", "€ m", f"=SUM(E{s.rows['turb']}:E{s.rows['fees']})", PER_MW, PER_MVA,
           bold=True, note="HTM total capex + IDC + fees")
    for k in [i[0] for i in items] + ["allin"]:
        for c in "FG":
            ws[f"{c}{s.rows[k]}"].number_format = _fmt("money")

    # 4. Generation and operations.
    s.bar("4. Generation and operations, both bids")
    s.header("Value", "per MW inst.", "per MVA")
    s.line("yr", "First full operating year", "year", "=Inputs!$E$156", fmt="int", note="Inputs E156")
    s.line("gen", "Net generation, first full year", "GWh", first_year(38), fmt="money", note="HTM row 38")
    s.line("eqh", "Equivalent full-load hours, first full year", "h/yr", None, "={gen}*Thousand/{inst}",
           "={gen}*Thousand/{mva}", fmt="money")
    s.line("ncf", "Net capacity factor, first full year", "%", None, f"=F{s.rows['eqh']}/Hours_In_Year",
           f"=G{s.rows['eqh']}/Hours_In_Year", fmt="pct")
    s.line("lost", "Energy lost to the bottleneck, first full year", "GWh",
           "={gen}*(1-In_Bottleneck_Factor)/In_Bottleneck_Factor", fmt="money1")
    s.line("price", "Achieved price, first full year", "€/MWh", first_year(49), fmt="money1", note="HTM row 49")
    s.line("rev", "Revenue, first full year", "€ m", first_year(54, "/Thousand"), fmt="money1", note="HTM row 54")
    s.line("ebitda", "EBITDA, first full year", "€ m", first_year(70, "/Thousand"), fmt="money1", note="HTM row 70")
    s.line("margin", "EBITDA margin, first full year", "%", first_year(71), fmt="pct", note="HTM row 71")
    s.line("opex", "Operating costs, first full year (k€ per MW / MVA)", "€ m | k€", first_year(69, "/Thousand", "-"),
           PER_MW, PER_MVA, fmt="money1", note="HTM row 69")
    for c in "FG":
        ws[f"{c}{s.rows['opex']}"].number_format = _fmt("money1")
    s.line("life_gen", "Lifetime net generation", "GWh", "=HTM!$G$38", fmt="money", note="HTM G38")
    s.line("lcoe", "LCOE", "€/MWh", "=HTM!$G$156", fmt="money1", note="HTM G156")

    # 5. Returns = Live_Results (one column); section 7 holds the stored copies.
    s.bar("5. Returns, live case (Live_Results)")
    first = s.r
    for key, label, unit, f, fmt, note in RETURNS:
        s.line(key, label, unit, f, fmt=fmt, note=note)
        define(wb, f"Out_{key}", f"Dashboard!$E${s.rows[key]}")
    define(wb, "Live_Results", f"Dashboard!$E${first}:$E${s.r - 1}")

    # 6. Securities and fees (Deposits, tender currency).
    s.bar("6. Securities and fees, both bids (Deposits)")
    for key, label, ref in (("part", "Participation fee", "H7"), ("bid", "Bid security", "H8"),
                            ("perf", "Performance security (gross; bid credit in Deposits J10)", "H10"),
                            ("tso", "TSO connection-request security", "H11"),
                            ("ind", "Connection fee, indirect (before TSO credit)", "H18"),
                            ("dir", "Connection fee, direct", "H12")):
        s.line(key, label, "mHUF", f"=Deposits!${ref[0]}${ref[1:]}", fmt="money", note=f"Deposits {ref}")
    s.line("gross", "Total gross", "mHUF", "=Deposits!$H$19", fmt="money", bold=True,
           note="Deposits H19")
    s.line("payable", "Total payable after netting", "mHUF", "=Deposits!$J$19", fmt="money", note="Deposits J19")
    s.line("cash", "Of which cash, in FEP cash flows", "€ m", "=Deposits!$E$26/Thousand", fmt="money1",
           note="Deposits E26")
    s.line("guar", "Of which bank guarantee, in FEP cash flows", "€ m", "=Deposits!$E$27/Thousand", fmt="money1",
           note="Deposits E27")

    # 7. Stored results by case, written by RunAllScenarios (purple).
    s.bar("7. Stored results by case (Stored_Results, written by RunAllScenarios)")
    for i in range(2):
        _write(ws, s.r, 5 + i, f"=Inputs!{L(7 + i)}$7", "header")
        ws.cell(s.r, 5 + i).alignment = Alignment(horizontal="right", wrap_text=True)
    s.r += 1
    top = s.r
    for key, label, unit, _, fmt, _ in RETURNS:
        apply(ws.cell(s.r, 3, label), "label")
        apply(ws.cell(s.r, 4, unit), "unit")
        for col in (5, 6):
            apply(ws.cell(s.r, col), "macro", fmt)
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
                                       ("a_exit", "Deemed exit at RtB charged to HTM (1 = yes)", "flag", "E228", "flag"),
                                       ("a_dsra", "DSRA", "months", "E268", "int"),
                                       ("a_sweep", "Cash sweep", "%", "E269", "pct")):
        s.line(key, label, unit, f"=Inputs!${ref[0]}${ref[1:]}", fmt=fmt, note=f"Inputs {ref}")
    _end(ws, s.r + 1, 8)
    return s.rows


# =========================================================================================== Cover, Change_Log
SHEETS = [("Dashboard", "One page for the IC: bid strategy and points, capacity bases, build cost, generation, "
                      "returns, securities, stored cases, checks.", "Live – case selector here"),
          ("Inputs", "Static inputs per case (G = EU turbine, H = CN turbine) and block U (points tables, constants, "
                     "row 300+).", "Live – edit here"),
          ("Curves", "EURIBOR forward curves and the Aurora Hungary wind capture prices.", "Live"),
          ("Deposits", "MEKH tender securities and grid-connection fees.", "Live"),
          ("MOIC", "Sell-at-RtB monthly engine (months 0–60). No inputs.", "Live"),
          ("HTM", "Annual build-and-operate engine 2026–2061. No inputs.", "Live"),
          ("Open_Items", "Open assumptions, interim treatment and owner.", "Hidden"),
          ("Change_Log", "Audit trail of every change.", "Hidden"),
          ("Devex_FEP, BORA_Devex", "Devex line items (BORA Devex 02-Oct-26).", "Hidden")]


def cover(wb):
    ws = wb["Cover"]
    for i, (sheet, purpose, status) in enumerate(SHEETS):
        for col, v in zip("BCD", (sheet, purpose, status)):
            ws[f"{col}{6 + i}"] = v
    for r in range(26, 44):                                   # headline block (duplicate of Dashboard) and old notes
        for col in "BC":
            ws[f"{col}{r}"].value = None
            ws[f"{col}{r}"].style = "Normal"
    ws["B26"], ws["B27"] = "Model status", "Master check (Dashboard)"
    ws["B26"].font = copy(ws["B18"].font)
    _write(ws, 27, 3, _status("Master_Check"))
    _ok_format(ws, "C27")
    ws["B29"] = "Update note"
    ws["B29"].font = copy(ws["B18"].font)
    ws["B30"] = ("22-Oct-26: Dashboard rebuilt as one page for the IC. The old Control_Panel, Dashboard, Bora and "
                 "Run_Map were deleted as duplicates; their Inputs links are now values. Bid levers and points bands on Inputs "
                 "block U; levers and case selector on Dashboard; Scenarios.RunAllScenarios stores the two cases. "
                 "See Change_Log #128.")
    ws["B31"] = "Hidden working sheets: Open_Items, Change_Log, Devex_FEP, BORA_Devex (right-click a tab → Unhide)."


def change_log(wb):
    ws, r = wb["Change_Log"], 131
    for col in "BCDEFGHIJ":
        ws[f"{col}{r}"]._style = copy(ws[f"{col}{r - 1}"]._style)
    for col, v in zip("BCDEFGHIJ", (
            f"=B{r - 1}+1", date(2026, 10, 22), "Dashboard sheet", "Dashboard; Inputs; Deposits; HTM; MOIC; Cover",
            "Inputs E6, E16, E20, E90, E168, G:H case cells, rows 300+; Deposits G11, G12, G18, C28:E34; HTM E36, "
            "H29:AQ29, E227:E243, G221:G222; MOIC G143, H226:N226",
            "Dashboard rebuilt as one IC page; old Control_Panel, Dashboard, Bora, Run_Map deleted (links → values, "
            "G292 formula moved from Control_Panel E22). Levers on Dashboard, points bands on Inputs block U; capacity, BESS, PB "
            "and RtB price on the MVA basis; two cases (EU, CN turbine); Deposits E34 is an internal check; "
            "Module1/Module2 replaced by Scenarios.bas.",
            "200 MW, BESS 100 MW, PB ×1.5, merchant base, 950 k€/MW, 3 cases",
            "200 MVA / 280 MW installed, BESS 80 MW, PB ×1.0, PPA 90 €/MWh 30 yrs, 600 / 500 k€/MW",
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
    block_u(wb)
    cases(wb)
    engine(wb, levers)
    summary(wb, levers)
    cover(wb)
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
