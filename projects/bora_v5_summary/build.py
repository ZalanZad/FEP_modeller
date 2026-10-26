"""Bora model RtB / COD / HTM – SPEC §18: the user's case made proper, Bora only, euros first.

Source (SPEC §18): reference/Bora_model_RTB_COD_HTM_261026_1.xlsm, the user's edit of our §17 output (scenario-table
tweaks, Inputs G162 typed 800 over the link, Inputs E322 typed 4 % over the interpolation, Bid_Points columns H:I deleted
by hand, stored results re-run). Every other cell stays as it is; the §18 edits are:
- §18.1 turbine: scenario-table Base turbine 800 and Inputs G162 back to the link; clipping: Dashboard lever
  Dash_Clip_Override (4 %, blank = memo curve) read by In_Bottleneck_Factor, E322 back to the interpolation;
- §18.2 Bora only: Bid_Points rebuilt from the §17.3 copy of the Wind HU sheet with a per-bid and a total column only;
  Inputs E9 / block M wording, Open_Items and Change_Log texts without other projects;
- §18.3 euros: Dashboard section 6 in € m (k€ where small); Bid_Points amounts in k€ (HUF only for the tender unit
  rates, one FX line);
- Change_Log: one line.
Rows are found by names and labels. Then vba/Scenarios.bas is re-embedded and RunAllScenarios stores the three cases.
build_v1.py … build_v3.py are the history (build_v3.py = §17); ag_tieout.py ties the cases out to the source file.
"""
import re
import shutil
import tempfile
import zipfile
from copy import copy
from datetime import date
from pathlib import Path

import openpyxl
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils import get_column_letter as L
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

from fepm.recalc import recalc
from fepm.xl.macros import VBA_DIR, package

HERE = Path(__file__).resolve().parent
REF = HERE.parents[1] / "reference"
SRC = REF / "Bora_model_RTB_COD_HTM_261026_1.xlsm"
WIND_HU = REF / "Wind_HU_points_deposits_devex_20261026.xlsx"
WIND_HU_SHEET = "Bid_Deposits_Devex"
NAME = "Bora_model_RTB_COD_HTM"

# §18.1 the user's case: EU turbine 800 k€/MW in the Base scenario, clipping 4 % instead of the memo curve.
BASE_TURBINE = 800
CLIP_OVERRIDE = 0.04
CLIP_NOTE = "Clipping 4 % (user, 26-Oct-26): the memo curve (9.9 % at this ratio) is for a 36 % CF park"

# §18.2 / §17.3 Bid_Points: Wind HU row numbers (the copy is rebuilt from the user's sheet of 26-Oct-2026).
# Column E = Bora, per bid; column F = total over In_Bids bids (amount rows only).
PER_BID, TOTAL = "Bora, per bid", '="Total, "&In_Bids&" bid(s)"'
BID_LEVERS = {8: "=Dash_MVA/In_Bids", 9: "=Dash_Overbuild", 10: "=Dash_BESS_Share", 11: "=Inputs!$E$21",
              12: "=Dash_Municipal", 13: "=Dash_Community", 14: "=Dash_Partial", 15: "=Dash_PB_Mult",
              16: "=Dash_Brownfield", 17: "=Dash_Shared", 18: "=Dash_EU_Turbine", 19: "=Dash_EU_Transformer"}
BID_RATES = {48: "=Inputs!$E$11", 49: "=Inputs!$E$75*Million", 50: "=Inputs!$E$89*Million",
             51: "=Inputs!$E$81*Million", 52: "=Inputs!$E$76", 53: "=Inputs!$E$77", 54: "=Inputs!$E$83*Million",
             55: "=Inputs!$E$85*Million", 56: "=Inputs!$E$84*Million*E$8", 57: "=Inputs!$E$113",
             58: "=(Inputs!$E$24+Inputs!$E$25)/In_Bids"}
BID_DATES = {37: "=Inputs!$E$27", 38: "=Inputs!$E$95", 39: "=Deposits!$K$10", 40: "=Inputs!$E$97",
             41: "=Devex_FEP!$F$38", 42: "=Inputs!$E$99", 43: "=Inputs!$E$154", 44: "=Inputs!$E$155"}
EUR_K = {75: "=E$58", 76: "=E$51/E$48/Thousand", 77: "=E$52*(E$20*E$49+E$22*E$50)/E$48/Thousand",
         78: "=E$53*E$15*(E$20*E$49+E$22*E$50)/E$48/Thousand", 79: "=E$54*E$8/E$48/Thousand",
         80: "=E$56/E$48/Thousand", 81: "=(E$55-E$54)*E$8/E$48/Thousand"}
# Devex lines: Wind HU row → Devex_FEP row (the model's k€ per line, scaled by In_Scale), per bid = ÷ In_Bids.
DEVEX = {**{r: r - 81 for r in range(86, 117)}, 117: 37, 118: 38}
DEVEX_OTHER = 36                                                   # no Wind HU line: inserted as row 117
TOTAL_ROWS = (8, 20, 21, 22, *range(75, 85), *range(86, 120), *range(123, 129), *range(134, 155))
HEADER_ROWS = (7, 47, 74, 85, 121, 133)
CASH_LINES = (75, 76, 79, 80, 81, *range(86, 119))                 # the lines the quarterly cash rows add up
DROP = {"rows": [(59, 2), (62, 12)], "summary": (129, 2)}           # devex for other sites, HUF block, HUF summary
PB_RELEASE_NOTE = "model has no PB release date; Kiírás XII.13, estimate"
CONSTANTS = {"R": ("Million", "#", "Unit conversion"), "S": ("Days in a year", "days", "Guarantee-fee day count"),
             "T": ("Months in a quarter", "months", "Calendar")}

# §18.2 texts that named other projects (Change_Log history condensed to Bora; Open_Items E43 wording).
TEXTS = {
    ("Inputs", "C9"): "Bora", ("Inputs", "F9"): "Fixed at 1",
    ("Inputs", "C279"): "M. Bora inputs – grid node fees from MEKH/MAVIR Kapacitás pályázati kiírás 2026, Annex 1 "
                        "(1. melléklet)",
    ("Inputs", "E280"): "Bora",
    ("Open_Items", "E43"): "Single asset with one COD versus several SPVs with staggered RTB / COD. Determines whether "
                           "Stage 2 / 3 run per SPV or on a blended asset.",
    ("Change_Log", "F99"): "Inputs E9, rows 16/20/24/31/32/38/82/84–88, block M C277:J296; Devex_FEP E5:E37, L4:L39; "
                           "Dashboard G14",
    ("Change_Log", "G99"): "Selector added at Inputs E9 (fixed to Bora since #127). Grid fees from Kapacitás pályázati "
                           "kiírás 2026 Annex 1. Restored Dashboard G14 formula (had been overwritten with 0).",
    ("Change_Log", "I99"): "Bora unchanged; RTB FEP IRR 27.5%",
    ("Change_Log", "E100"): "Dashboard; Inputs; Run_Map",
    ("Change_Log", "G100"): "Added Dashboard scenario control panel: case selector and case inputs (Base / European / "
                            "Chinese: EYA, capex, PPA, Aurora, opex, gearing, rate …). Inputs now link to the panel "
                            "(green). Run_Map holds live result formulas; RunAllScenarios VBA (text in Run_Map B100+) "
                            "rewrites Dashboard section 11. Button added – macro must be pasted and assigned by the user.",
    ("Change_Log", "E103"): "Dashboard; Run_Map",
    ("Change_Log", "F103"): "Dashboard E61/E63, C61/C63, K76, K90, H90; Run_Map C22:C23",
    ("Change_Log", "G107"): "Stale cell-address labels updated after row inserts; 400 kV → 132 kV notes; E220 tenor "
                            "label; Dashboard F8 stale-snapshot warning; Open_Items D6/E7/F9/E10/F12 refreshed",
    ("Change_Log", "E109"): "MOIC; Dashboard; Run_Map",
    ("Change_Log", "F109"): "MOIC rows 270, 272–278 deleted; Dashboard row 150 deleted; Dashboard C65:E66; Run_Map "
                            "B36:C37",
    ("Change_Log", "E111"): "Inputs; Run_Map",
    ("Change_Log", "F111"): "Inputs C9, J278, J279, F283, F284, F292; Run_Map C8",
    ("Change_Log", "G111"): "Name spelling aligned across Inputs and Run_Map on user instruction (superseded by #127).",
    ("Change_Log", "H111"): "n/a", ("Change_Log", "I111"): "n/a",
    ("Change_Log", "G113"): "Base turbine price 500 → 950 k€/MW (European) to score the ≥80% EU-equipment tender "
                            "points.",
    ("Change_Log", "E123"): "Dashboard; Run_Map",
    ("Change_Log", "F123"): "Dashboard E7:E8; Dashboard section 11 E/F; Run_Map C3; Control_Panel E8",
    ("Change_Log", "E125"): "Inputs; Dashboard; Control_Panel; Run_Map",
    ("Change_Log", "F125"): "Inputs H7:I276 ↔ K7:L276; Dashboard F7:G8, E241:F241; macro",
    ("Change_Log", "F127"): "Control_Panel E6, F6, F11:H22; Inputs E9, H278:J296; Dashboard E211:F240; Run_Map C6:C8, "
                            "VBA text",
    ("Change_Log", "G127"): "Model reduced to Bora only: selector fixed to 1, inputs and notes not for Bora removed. "
                            "Dashboard section 11 re-populated with Bora case 2 / case 3 results.",
    ("Change_Log", "H127"): "Selector over several sites",
}
OTHER_PROJECTS = re.compile(r"Iváncsa|Dunaújváros|\bBana\b|Bána|Ócsa|Comparison sheet|project #|project sheet", re.I)

# §18.3 Dashboard section 6 (label → unit, formula): € m at the Inputs FX (mHUF ÷ HUF/EUR), k€ where small.
SECTION6 = {"Participation fee": ("k€", "=Deposits!$M$7", "#,##0.0"),
            "Bid security": ("€ m", "=Deposits!$H$8/Inputs!$E$11", None),
            "Performance security (gross": ("€ m", "=Deposits!$H$10/Inputs!$E$11", None),
            "TSO connection-request security": ("€ m", "=Deposits!$H$11/Inputs!$E$11", None),
            "Connection fee, indirect (before TSO credit)": ("€ m", "=Deposits!$H$18/Inputs!$E$11", None),
            "Connection fee, direct": ("€ m", "=Deposits!$H$12/Inputs!$E$11", None),
            "Total gross": ("€ m", "=Deposits!$H$19/Inputs!$E$11", None),
            "Total payable after netting": ("€ m", "=Deposits!$J$19/Inputs!$E$11", None)}

GREEN = "FF00B050"


# =========================================================================================== helpers
def define(wb, name, ref):
    wb.defined_names[name] = DefinedName(name, attr_text=ref)


def like(cell, model, value=None, fmt=None):
    """Write `value` with the style of `model` (and an optional number format)."""
    cell.value = value
    cell._style = copy(model._style)
    if fmt:
        cell.number_format = fmt
    return cell


def name_cell(wb, name):
    sheet, ref = next(wb.defined_names[name].destinations)
    return wb[sheet][ref.replace("$", "").split(":")[0]]


def label_row(ws, text, col=3, start=1):
    """First row at or below `start` whose label in `col` starts with `text`."""
    rows = sorted(c.row for c in ws._cells.values()
                  if c.column == col and c.row >= start and isinstance(c.value, str) and c.value.startswith(text))
    assert rows, (ws.title, text)
    return rows[0]


# =========================================================================================== row insertion / deletion
_ADDR = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")
_REF = re.compile(r"\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?")


def _shift_addr(addr, at, n):
    return _ADDR.sub(lambda m: f"{m[1]}{m[2]}{m[3]}{int(m[4]) + n if int(m[4]) >= at else m[4]}", addr)


def _shift_formula(f, host, sheet, at, n):
    """Shift every cell reference to `sheet` rows ≥ `at` by `n` in formula `f` written on sheet `host`."""
    out = []
    for t in Tokenizer(f).items:
        v = t.value
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            ref_sheet, addr = (v.rsplit("!", 1) if "!" in v else (None, v))
            if (ref_sheet.strip("'") if ref_sheet else host) == sheet and _REF.fullmatch(addr):
                v = (ref_sheet + "!" if ref_sheet else "") + _shift_addr(addr, at, n)
        out.append(v)
    return "=" + "".join(out)


def _refs_rows(f, host, sheet):
    """Rows of `sheet` referred to as single cells or range ends in formula `f` (for the deletion guard)."""
    rows = []
    for t in Tokenizer(f).items:
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            ref_sheet, addr = (t.value.rsplit("!", 1) if "!" in t.value else (None, t.value))
            if (ref_sheet.strip("'") if ref_sheet else host) == sheet and _REF.fullmatch(addr):
                rows += [int(m[4]) for m in _ADDR.finditer(addr)]
    return rows


def shift_rows(wb, sheet, at, n):
    """Insert `n` rows before row `at` (n > 0) or delete −n rows from row `at` (n < 0) on `sheet`, re-pointing every
    formula, name, merged range, conditional format, validation and array-formula range (openpyxl moves cells only)."""
    ws = wb[sheet]
    first = at if n > 0 else at - n                                    # first row that moves
    gone = range(at, at - n) if n < 0 else range(0)
    for w in wb.worksheets:
        for c in list(w._cells.values()):
            v = c.value
            text = v.text if isinstance(v, ArrayFormula) else (v if c.data_type == "f" else None)
            if not text:
                continue
            assert not set(_refs_rows(text, w.title, sheet)) & set(gone), (w.title, c.coordinate, text)
            new = _shift_formula(text, w.title, sheet, first, n)
            if isinstance(v, ArrayFormula):
                v.text = new
            elif new != v:                                             # text that starts with "=" stays text
                c.value = new
    for name, dn in list(wb.defined_names.items()):
        if dn.attr_text and f"{sheet}!" in dn.attr_text:
            dn.attr_text = _shift_formula("=" + dn.attr_text, "", sheet, first, n)[1:]
    merged = [str(m) for m in ws.merged_cells.ranges]
    for m in merged:
        ws.unmerge_cells(m)
    heights = {r: d.height for r, d in ws.row_dimensions.items() if d.height and r >= at}
    if n > 0:
        ws.insert_rows(at, n)
    else:
        ws.delete_rows(at, -n)
    for r in heights:
        ws.row_dimensions[r].height = None
    for r, h in heights.items():
        if r >= first:
            ws.row_dimensions[r + n].height = h
    for m in merged:
        ws.merge_cells(_shift_addr(m, first, n))
    for c in ws._cells.values():                                     # array formulas carry their own range
        if isinstance(c.value, ArrayFormula) and c.value.ref:
            c.value.ref = c.coordinate if ":" not in c.value.ref else _shift_addr(c.value.ref, first, n)
    cf = ws.conditional_formatting
    ws.conditional_formatting = ConditionalFormattingList()
    for rng in cf:
        for rule in rng.rules:
            ws.conditional_formatting.add(_shift_addr(str(rng.sqref), first, n), rule)
    for dv in ws.data_validations.dataValidation:
        dv.sqref = MultiCellRange(_shift_addr(str(dv.sqref), first, n))
    if ws.freeze_panes:
        ws.freeze_panes = _shift_addr(ws.freeze_panes, first, n)

# =========================================================================================== §17.3 Bid_Points
def _copy_cell(src, dst):
    v = src.value
    if isinstance(v, ArrayFormula):
        v = ArrayFormula(dst.coordinate, v.text)
    dst.value = v
    if src.has_style:
        dst.font, dst.fill, dst.border = copy(src.font), copy(src.fill), copy(src.border)
        dst.alignment, dst.protection, dst.number_format = copy(src.alignment), copy(src.protection), src.number_format


def _link(cell, formula):
    """A link from another sheet: green font, no input fill (MODELLING_STANDARDS §5)."""
    cell.value = formula
    f = copy(cell.font)
    cell.font = Font(name=f.name, size=f.sz, bold=f.b, italic=f.i, color=GREEN)
    cell.fill = PatternFill(fill_type=None)


def _pure_link(f):
    toks = [t for t in Tokenizer(f).items if t.type != Token.WSPACE] if isinstance(f, str) and f.startswith("=") else []
    return toks[0].value if len(toks) == 1 and toks[0].type == Token.OPERAND and toks[0].subtype == Token.RANGE \
        else None


def _target(wb, host, op):
    """(sheet, coordinate) of a single-cell reference or single-cell name, else None."""
    if op in wb.defined_names:
        sheet, ref = next(wb.defined_names[op].destinations)
    else:
        sheet, ref = op.rsplit("!", 1) if "!" in op else (host, op)
    ref = ref.replace("$", "")
    return (sheet.strip("'"), ref) if re.fullmatch(r"[A-Z]{1,3}\d+", ref) else None


def collapse_chains(wb, sheet):
    """Every pure link on `sheet` that points at another pure link is re-pointed at the source (no chains)."""
    ws = wb[sheet]
    for c in ws._cells.values():
        op = _pure_link(c.value)
        while op:
            tgt = _target(wb, sheet, op)
            if not tgt:
                break
            s, coord = tgt
            inner = _pure_link(wb[s][coord].value)
            if not inner:
                break
            op = inner if ("!" in inner or inner in wb.defined_names or s == sheet) else f"{s}!{inner}"
            c.value = "=" + op


# =========================================================================================== §18.2/18.3 Bid_Points
def _clear(ws, rows, cols):
    for r in rows:
        for col in cols:
            ws[f"{col}{r}"].value = None


def bid_points(wb):
    """The Wind HU sheet copied in (§17.3), then Bora only: column E per bid, column F the total over In_Bids bids,
    amounts in k€ (the HUF unit rates of the tender stay HUF, converted at the one FX line E48)."""
    at = wb.sheetnames.index("Bid_Points")
    del wb["Bid_Points"]
    src = openpyxl.load_workbook(WIND_HU)[WIND_HU_SHEET]
    ws = wb.create_sheet("Bid_Points", at)
    for row in src.iter_rows():
        for c in row:
            if c.value is not None or c.has_style:
                _copy_cell(c, ws.cell(c.row, c.column))
    for k, dim in src.column_dimensions.items():
        ws.column_dimensions[k].width, ws.column_dimensions[k].hidden = dim.width, dim.hidden
    for r, dim in src.row_dimensions.items():
        if dim.height:
            ws.row_dimensions[r].height = dim.height
    ws.freeze_panes = src.freeze_panes
    ws.sheet_view.showGridLines = src.sheet_view.showGridLines
    for rng in src.conditional_formatting:
        for rule in rng.rules:
            ws.conditional_formatting.add(str(rng.sqref).replace("E156:I156", "E156:F156"), rule)

    # Bora only: the other sites' columns (G BORA S, H Iváncsa, I Dunaújváros, J total, N:O source split) go;
    # the scoring tables (rows 158+) keep their own columns E, F, H, I.
    _clear(ws, range(1, 157), "GHIJNO")
    _clear(ws, [r for r in range(1, 61) if r not in (7, 8, 20, 21, 22)], "F")
    ws["C3"].value = "Bora"
    ws["C4"].value = '="Bid strategy – "&TEXT($E$9,"0 %")&" overbuild"'
    for r in HEADER_ROWS:
        ws[f"E{r}"].value, ws[f"F{r}"].value = PER_BID, (TOTAL if r != 47 else None)
    ws["C34"].value = ws["C122"].value = "Points total, live case"

    # Constants: the workbook's Thousand (Inputs block U); Million, days in a year and the quarter length here.
    ws["Q4"].value = None
    _copy_cell(src["S4"], ws["T4"])
    ws["T4"].value = 3
    for col, (label, unit, source) in CONSTANTS.items():
        like(ws[f"{col}4"], ws["H8"], ws[f"{col}4"].value, "#,##0")
        like(ws[f"{col}3"], ws["C8"], label)
        like(ws[f"{col}5"], ws["D8"], unit)
        like(ws[f"{col}6"], ws["D8"], source)
    define(wb, "Million", "Bid_Points!$R$4")
    define(wb, "Days_In_Year", "Bid_Points!$S$4")
    define(wb, "Qtr_Months", "Bid_Points!$T$4")

    for r, f in {**BID_LEVERS, **BID_RATES, **BID_DATES}.items():   # the model's bid, rates and dates
        _link(ws[f"E{r}"], f)
    ws["C56"].value = "Direct connection fee per bid (model: per MVA of the shared point)"
    like(ws["F45"], ws["D45"], PB_RELEASE_NOTE)
    ws["F222"].value = 1                                              # ÖSZ-6: x ≥ 2.0 scores 100 (GYIK Q34)

    for r, f in EUR_K.items():                                         # securities and fees in k€
        ws[f"E{r}"].value = f
    for r, d in DEVEX.items():                                         # devex per line = the model's k€ ÷ bids
        _link(ws[f"E{r}"], f"=Devex_FEP!$E${d}/In_Bids")
        if r != 118:
            _link(ws[f"L{r}"], f"=Devex_FEP!$F${d}")
    for r in TOTAL_ROWS:
        ws[f"F{r}"].value = f"=E{r}*In_Bids"
        ws[f"F{r}"]._style = copy(ws[f"E{r}"]._style)
    ws["F122"].value = ws["F131"].value = None

    # Quarterly cash: one SUMPRODUCT per cell over the cash lines (flag column M); cumulative and guarantees on the
    # total column.
    like(ws["M74"], ws["L74"], "In cash")
    for r in range(75, 119):
        if ws[f"E{r}"].value is not None and r != 85:
            like(ws[f"M{r}"], ws["F222"], 1 if r in CASH_LINES else 0, "0")
    define(wb, "Bid_Dates", "Bid_Points!$L$75:$L$118")
    define(wb, "Bid_Cash", "Bid_Points!$M$75:$M$118")
    first, last = 134, label_row(ws, "Total", start=134) - 1
    for r in range(first, last + 1):
        if r > first:
            ws[f"C{r}"].value = f"=EOMONTH(C{r - 1},Qtr_Months)"
        ws[f"E{r}"].value = f"=SUMPRODUCT((Bid_Dates>EOMONTH($C{r},-Qtr_Months))*(Bid_Dates<=$C{r})*Bid_Cash,E$75:E$118)"
        ws[f"L{r}"].value = f"=F{r}" if r == first else f"=L{r - 1}+F{r}"
        ws[f"M{r}"].value = (f"=($C{r}<=$E$44)*(($E$38<=$C{r})*($E$39>$C{r})*$F$77"
                             f"+($E$39<=$C{r})*($E$45>$C{r})*$F$78)")
    chk = label_row(ws, "Check: quarters vs summary")
    for col in "EF":
        ws[f"{col}{chk - 1}"].value = f"=SUM({col}{first}:{col}{last})"
        ws[f"{col}{chk}"].value = f'=IF(ABS({col}{chk - 1}-{col}126)<Tolerance,"OK","CHECK")'
    mx = label_row(ws, "Maximum achievable points")
    ws[f"I{mx}"].value = "=(E159+E171+E185+E198+E211+E225+E226+E227+E229)*MAX(H160:H231)"

    other(wb, ws)
    start, n = DROP["summary"]                                         # HUF rows go (bottom up: rows stay valid)
    for r0, k in [(start + 1, n), *reversed(DROP["rows"])]:
        _clear(ws, range(r0, r0 + k), [L(c) for c in range(1, 21)])
        shift_rows(wb, "Bid_Points", r0, -k)
    collapse_chains(wb, "Bid_Points")
    return ws


def other(wb, ws):
    """Devex_FEP row 36 (other costs) has no Wind HU line: a row is inserted above the RtB land purchase."""
    at = 117
    shift_rows(wb, "Bid_Points", at, 1)
    for col in range(2, 14):
        like(ws.cell(at, col), ws.cell(at + 1, col))
    ws[f"C{at}"].value, ws[f"D{at}"].value = "Other costs (HR, car etc.)", ws[f"D{at + 1}"].value
    _link(ws[f"E{at}"], f"=Devex_FEP!$E${DEVEX_OTHER}/In_Bids")
    ws[f"F{at}"].value = f"=E{at}*In_Bids"
    _link(ws[f"L{at}"], f"=Devex_FEP!$F${DEVEX_OTHER}")
    ws[f"M{at}"].value = 1


# =========================================================================================== §18.1 the user's case
def turbine(wb):
    """Base turbine 800 in the scenario table; Inputs G162 links to it again (the user typed 800 over the link)."""
    name_cell(wb, "Dash_Key_Inputs").value = BASE_TURBINE
    inp = wb["Inputs"]
    inp[f"G{label_row(inp, 'Turbine supply incl. transport')}"].value = "=Dashboard!$H$8"


def clip_override(wb):
    """Dashboard lever Dash_Clip_Override (blank = memo curve); E322 back to the interpolation (typed 4 % over it)."""
    inp, d = wb["Inputs"], wb["Dashboard"]
    clip = label_row(inp, "Clipped share at the built ratio")
    inp[f"E{clip}"].value = f"=MAX(0,E{clip - 3}+E{clip - 1}*(E{clip - 8}-E{clip - 5}))"
    inp[f"G{clip}"].value = None                                      # the user's stray copy of the formula
    at = label_row(d, "Turbine route (case)")
    shift_rows(wb, "Dashboard", at, 1)
    for col in range(3, 9):
        like(d.cell(at, col), d.cell(at - 1, col))
    d.cell(at, 3).value, d.cell(at, 4).value = "Clipped share of gross energy (blank = OVIT memo curve)", "%"
    d.cell(at, 5).value = CLIP_OVERRIDE
    d.cell(at, 5).number_format = "0.0%"
    define(wb, "Dash_Clip_Override", f"Dashboard!$E${at}")
    factor = name_cell(wb, "In_Bottleneck_Factor")
    factor.value = f'=1-IF(Dash_Clip_Override="",E{clip},Dash_Clip_Override)'
    inp.cell(factor.row, 3).value = "Generation kept after clipping (1 − clipped share applied)"
    inp.cell(factor.row, 6).value = "Dash_Clip_Override, else the curve"
    d.cell(label_row(d, "Clipped share of gross energy (OVIT"), 3).value = "Clipped share of gross energy applied"
    d.cell(label_row(d, "Clipping per OVIT memo"), 3).value = CLIP_NOTE


# =========================================================================================== §18.2 Bora only
def bora_only(wb):
    for (sheet, ref), text in TEXTS.items():
        wb[sheet][ref].value = text
    left = [(ws.title, c.coordinate) for ws in wb.worksheets if ws.title != "BORA_Devex" for c in ws._cells.values()
            if isinstance(c.value, str) and OTHER_PROJECTS.search(c.value)]
    assert not left, left


# =========================================================================================== §18.3 euros first
def euros(wb):
    d = wb["Dashboard"]
    top = label_row(d, "6. Securities and fees")
    d.cell(top, 3).value = "6. Securities and fees (Deposits; € m at the Inputs FX)"
    model = label_row(d, "Of which cash", start=top)
    for label, (unit, f, fmt) in SECTION6.items():
        r = label_row(d, label, start=top)
        d.cell(r, 4).value = unit
        like(d.cell(r, 5), d.cell(model, 5), f, fmt)


def change_log(wb):
    ws = wb["Change_Log"]
    r = max(c.row for c in ws._cells.values() if c.column == 2 and c.value is not None) + 1
    for col in "BCDEFGHIJ":
        ws[f"{col}{r}"]._style = copy(ws[f"{col}{r - 1}"]._style)
    for col, v in zip("BCDEFGHIJ", (
            f"=B{r - 1}+1", date(2026, 10, 26), "User case made proper; Bora only; euros",
            "Dashboard; Inputs; Bid_Points; Open_Items; Change_Log",
            "Dashboard E8, clipping lever row, section 6, key assumptions; Inputs G162, E322, In_Bottleneck_Factor, "
            "C9, F9, C279, E280; Bid_Points rebuilt",
            "Base turbine 800 via the scenario table; Dash_Clip_Override 4 % (blank = memo curve) replaces the typed "
            "E322; Bid_Points per bid and total only, amounts in k€; Dashboard section 6 in € m; texts naming other "
            "sites removed.",
            "G162 and E322 typed over formulas; Bid_Points with other sites; section 6 in mHUF",
            "Same numbers from levers and links; Bora only; € on the presentation sheets",
            "User 26-Oct-2026; SPEC §18")):
        ws[f"{col}{r}"] = v



# =========================================================================================== build
def general(path, sheet, cells, like):
    """SPEC §8.6: give `cells` the style index of `like` (a General cell) inside the saved file; LibreOffice infers
    a 0.00 % format for these sums on save (HTM G10 would show 3000.00 %)."""
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




def edit():
    wb = openpyxl.load_workbook(SRC, keep_vba=False)
    turbine(wb)
    clip_override(wb)
    bid_points(wb)
    bora_only(wb)
    euros(wb)
    change_log(wb)
    return wb


def build(out=HERE / "out", name=NAME):
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    xlsx, xlsm = out / f"{name}.xlsx", out / f"{name}.xlsm"
    edit().save(xlsx)
    with tempfile.TemporaryDirectory() as vba:                         # Scenarios.bas only (as is)
        shutil.copy(VBA_DIR / "Scenarios.bas", vba)
        package(xlsx, xlsm, run=["Scenarios.RunAllScenarios"], vba_dir=Path(vba))
    general(xlsm, "HTM", ("G7", "G10", "G14"), like="G8")
    rc = recalc(xlsm, out / f"{name}_recalc.xlsx")
    return {"xlsx": xlsx, "xlsm": xlsm, "recalc": rc}


if __name__ == "__main__":
    print(build())
