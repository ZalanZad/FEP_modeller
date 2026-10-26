"""Bora model RtB / COD / HTM – SPEC §16: a real refinancing at COD + 1 year for case 3. Built to SPEC.md §16.

Source (SPEC §16.1): reference/Bora_model_RTB_COD_HTM_20261022.xlsm, the user's re-save of our last output (incl. the
user's "Sum" column on Cash_Flow). Every cell stays as it is except the §16 edits:
- Inputs block U: the refinancing inputs (In_Refi_*; In_Refi_On is case-driven, columns G:I);
- Inputs G7:I8 / I217:I220: case 3 becomes "EU turbine, refinanced at COD + 1 year" on the case 1 debt terms;
- HTM block 18 (rows 357+): refinancing year, new loan, new annuity, flags, rate row and the principal helpers;
  HTM rows 77, 81, 83, 85, 87 and G158 read them (with In_Refi_On = 0 every value is unchanged);
- Dashboard: four returns-table rows and four key-assumption rows (rows inserted, every reference re-pointed);
- Change_Log: one line.
Then vba/Scenarios.bas is re-embedded and RunAllScenarios stores the three cases. build_v1.py is the history.
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
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.formula import ArrayFormula

from fepm.recalc import recalc
from fepm.xl.macros import VBA_DIR, package

HERE = Path(__file__).resolve().parent
SRC = HERE.parents[1] / "reference" / "Bora_model_RTB_COD_HTM_20261022.xlsm"
NAME = "Bora_model_RTB_COD_HTM"
REFI = 357                                         # first row of the new HTM block 18
U_REFI = 350                                       # first row of the refinancing inputs on Inputs block U
CASE3 = "EU turbine, refinanced at COD + 1 year"

# Refinancing inputs (Inputs block U): name, label, unit, value (list = case values G:I), note, number format
REFI_INPUTS = [
    ("In_Refi_On", "Refinancing at COD + lag (1 = yes; case-driven, columns G:I)", "flag", [0, 0, 1], "SPEC §16",
     "0"),
    ("In_Refi_Lag", "Refinancing lag after the COD year", "years", 1, "SPEC §16", "0"),
    ("In_Refi_Gearing_Add", "New loan gearing on top of the base gearing", "%", 0.04, "SPEC §16 (+4 pp)", "0.0%"),
    ("In_Refi_Rate", "New loan all-in fixed rate", "%", 0.045, "SPEC §16 (−50 bp)", "0.00%"),
    ("In_Refi_Tenor", "New loan tenor (annuity from the year after the refinancing)", "years", 15, "SPEC §16", "0"),
    ("In_Refi_Fee", "Upfront fee on the new loan (paid by equity)", "% of debt", "=E219", "Inputs E219", "0.00%"),
]

# Dashboard section 0 (Dash_Key_Inputs, rows: turbine EU k€/MW, PPA €/MWh, net P50 h, O&M k€/MW/yr, RtB premium
# k€/MVA; columns Base | Low | High). User 22-Oct-26: Low = the former Base; new Base turbine 850, PPA 95 (lower
# balancing fee, GoO 2–2.5 €/MWh), P50 2,500; High turbine 800.
KEY_TABLE = {"Base": [850, 95, 2500, 12, 150], "Low": [950, 90, 2460, 12, 150], "High": [800, 100, 2600, 10, 200]}

# Regression only (regression.py): the v5 source levers written onto the returned file, refinancing off.
SOURCE = {"names": {"Dash_MVA": 200, "Dash_Municipal": 20, "Dash_Community": 2, "Dash_Overbuild": 0.0,
                    "Dash_BESS_Share": 0.5, "Dash_PB_Mult": 1.5, "Dash_Build_Shortfall": 0.0,
                    "In_Exit_Profit_Charge": 1},
          "cells": {"Inputs!G78": 0.05, "Inputs!G190": 0, "Inputs!H190": 1, "Inputs!G193": 10, "Inputs!G217": 0.5,
                    "Inputs!H217": 0, "Inputs!G220": 10, "Inputs!G11": 370, "Inputs!G268": 6, "Inputs!G269": 0.5,
                    f"Inputs!I{U_REFI}": 0},
          "gearing_steps": [0, 0.3, 0.4, 0.5, 0.55, 0.6, 0.7],
          "key_base": [950, 90, 2460, 12, 150]}


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


def _text(v):
    return v.text if isinstance(v, ArrayFormula) else v


# =========================================================================================== row insertion
_ADDR = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def _shift_addr(addr, at, n):
    return _ADDR.sub(lambda m: f"{m[1]}{m[2]}{m[3]}{int(m[4]) + n if int(m[4]) >= at else m[4]}", addr)


def _shift_formula(f, host, sheet, at, n):
    """Shift every reference to `sheet` rows ≥ `at` by `n` in formula `f` written on sheet `host`."""
    out = []
    for t in Tokenizer(f).items:
        v = t.value
        if t.type == Token.OPERAND and t.subtype == Token.RANGE:
            ref_sheet, addr = (v.rsplit("!", 1) if "!" in v else (None, v))
            if (ref_sheet.strip("'") if ref_sheet else host) == sheet:
                v = (ref_sheet + "!" if ref_sheet else "") + _shift_addr(addr, at, n)
        out.append(v)
    return "=" + "".join(out)


def insert_rows(wb, sheet, at, n):
    """Insert `n` rows on `sheet` before row `at` and re-point every formula, name and format range that refers to
    the moved rows (openpyxl's insert_rows moves cells only)."""
    ws = wb[sheet]
    for w in wb.worksheets:
        for row in w.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, ArrayFormula):
                    v.text = _shift_formula(v.text, w.title, sheet, at, n)
                elif c.data_type == "f":                              # text that starts with "=" stays text
                    new = _shift_formula(v, w.title, sheet, at, n)
                    if new != v:
                        c.value = new
    for name, dn in list(wb.defined_names.items()):
        if dn.attr_text and f"{sheet}!" in dn.attr_text:
            dn.attr_text = _shift_formula("=" + dn.attr_text, "", sheet, at, n)[1:]
    ws.insert_rows(at, n)
    heights = {r: d.height for r, d in ws.row_dimensions.items() if d.height and r >= at}
    for r in sorted(heights, reverse=True):
        ws.row_dimensions[r + n].height, ws.row_dimensions[r].height = heights[r], None
    cf = ws.conditional_formatting
    ws.conditional_formatting = ConditionalFormattingList()
    for rng in cf:
        for rule in rng.rules:
            ws.conditional_formatting.add(_shift_addr(str(rng.sqref), at, n), rule)
    for dv in ws.data_validations.dataValidation:
        dv.sqref = openpyxl.worksheet.cell_range.MultiCellRange(_shift_addr(str(dv.sqref), at, n))


# =========================================================================================== §16 inputs and case 3
def refi_inputs(wb):
    """Inputs block U rows for the refinancing; In_Refi_On picks its case value like the engine's case rows."""
    ws = wb["Inputs"]
    label, unit, value, note, case = ws["C301"], ws["D301"], ws["E301"], ws["F301"], ws["G217"]
    like(ws.cell(U_REFI - 1, 3), ws["C300"], "Refinancing at COD + 1 year (SPEC §16; case 3)")
    for r, (name, lab, un, v, src, fmt) in enumerate(REFI_INPUTS, U_REFI):
        like(ws.cell(r, 3), label, lab)
        like(ws.cell(r, 4), unit, un)
        like(ws.cell(r, 6), note, src)
        if isinstance(v, list):                                    # case-driven: G:I values, E picks the live case
            for i, x in enumerate(v):
                like(ws.cell(r, 7 + i), case, x, fmt)
            like(ws.cell(r, 5), ws["E302"], f'=IF(INDEX($G{r}:$P{r},$E$6)="",$G{r},INDEX($G{r}:$P{r},$E$6))', fmt)
        else:
            like(ws.cell(r, 5), ws["E302"] if isinstance(v, str) else value, v, fmt)
        define(wb, name, f"Inputs!$E${r}")


def case3(wb):
    """Case 3 is case 1's debt (60 % at 5 %, 15 years) until the refinancing: its own debt cells are cleared."""
    ws = wb["Inputs"]
    ws["I7"] = CASE3
    ws["I8"] = ("As the EU turbine case; at the end of the year after the COD year the loan is refinanced: "
                "+4 pp gearing at 4.5 % over 15 years")
    for r in (217, 218, 220):
        ws[f"I{r}"] = None
    ws["F217"] = "Case 3 refinances to gearing + In_Refi_Gearing_Add (Inputs block U)"


# =========================================================================================== §16 HTM
def htm_refi(wb):
    """HTM block 18 and the debt-block edits. Timeline H:AQ; one formula per row; In_Refi_On = 0 changes nothing.
    The old loan is repaid at the end of the refinancing year through the prepayment row 87 (unscheduled, so the
    DSCR keeps measuring scheduled debt service); the new loan is drawn in row 77 the same year."""
    htm = wb["HTM"]
    R, N, A = f"$G${REFI + 1}", f"$G${REFI + 2}", f"$G${REFI + 3}"
    flag, after, rate, old, new, sweep, repaid, fee, draw = (REFI + i for i in range(4, 13))
    rows = [  # row, label, unit, G (scalar), timeline formula ({c} = column), style model row
        (REFI + 1, "Refinancing year (end of year; COD year + In_Refi_Lag)", "year",
         "=Inputs!$E$156-1+In_Refi_Lag", None, 345),
        (REFI + 2, "New loan: (gearing + In_Refi_Gearing_Add) × the sizing base (E75)", "k€",
         "=(Inputs!$E$217+In_Refi_Gearing_Add)*-$G$30", None, 336),
        (REFI + 3, "New loan annuity (In_Refi_Rate, In_Refi_Tenor)", "k€", f"=PMT(In_Refi_Rate,In_Refi_Tenor,-{N})",
         None, 336),
        (flag, "Refinancing year flag", "flag", None, f"=In_Refi_On*({{c}}$5={R})", 335),
        (after, "After-refinancing flag", "flag", None, f"=In_Refi_On*({{c}}$5>{R})", 335),
        (rate, "Interest rate on the opening balance", "%", None, f"=$E$74+{{c}}{after}*(In_Refi_Rate-$E$74)", 335),
        (old, "Old loan principal (scheduled annuity, up to the refinancing)", "k€", "=SUM(H{r}:AQ{r})",
         f"=IF(AND({{c}}9>=1,{{c}}9<=Inputs!$E$220),MIN(MAX($E$82-{{c}}81,0),{{c}}76),0)*(1-{{c}}{after})", 335),
        (new, "New loan principal (annuity after the refinancing)", "k€", "=SUM(H{r}:AQ{r})",
         f"={{c}}{after}*({{c}}$5<={R}+In_Refi_Tenor)*MIN(MAX({A}-{{c}}81,0),{{c}}76)", 335),
        (sweep, "Cash sweep (row 87 as before)", "k€", "=SUM(H{r}:AQ{r})", None, 335),
        (repaid, "Old loan repaid at the refinancing (prepayment)", "k€", "=SUM(H{r}:AQ{r})",
         f"={{c}}{flag}*({{c}}76-{{c}}{old}-{{c}}{sweep})", 335),
        (fee, "Refinancing fee (upfront fee on the new loan)", "k€", "=SUM(H{r}:AQ{r})",
         f"=-{{c}}{flag}*{N}*In_Refi_Fee", 335),
        (draw, "New loan drawn at the refinancing (in row 77; excluded from E82 and G157)", "k€", "=SUM(H{r}:AQ{r})",
         f"={{c}}{flag}*{N}", 335),
        (REFI + 13, "Last debt-service year", "year",
         f"=MAX(Inputs!$E$156+Inputs!$E$220-1,In_Refi_On*({R}+In_Refi_Tenor))", None, 345),
        (REFI + 14, "Closing debt balance in the last debt-service year", "k€",
         f"=SUMIF(H5:AQ5,$G${REFI + 13},H84:AQ84)", None, 336),
        (REFI + 15, "Debt after refinancing (closing balance, refinancing year)", "k€",
         f"=In_Refi_On*SUMIF(H5:AQ5,{R},H84:AQ84)", None, 336),
        (REFI + 16, "Gearing after refinancing (debt ÷ sizing base)", "%", f"=$G${REFI + 15}/-$G$30", None, 339),
        (REFI + 17, "Minimum DSCR after refinancing", "x",
         f'=IF(In_Refi_On=1,_xlfn.MINIFS(H126:AQ126,H5:AQ5,">"&{R}),"n/a")', None, 343),
        (REFI + 18, "Released to equity at the refinancing (new loan − old loan repaid − fee)", "k€",
         f"=In_Refi_On*{N}-$G${repaid}+$G${fee}", None, 336),
    ]
    like(htm[f"C{REFI}"], htm["C332"], "18. Refinancing at COD + 1 year (case 3; In_Refi_On = 0 elsewhere)")
    sweep_orig = {L(i): htm[f"{L(i)}87"].value for i in range(8, 44)}  # the original sweep, moved to the helper
    for r, label, unit, g, line, model in rows:
        like(htm[f"C{r}"], htm[f"C{model}"], label)
        like(htm[f"D{r}"], htm[f"D{model}"], unit)
        if g:
            like(htm[f"G{r}"], htm[f"G{model}"], g.replace("{r}", str(r)),
                 {"year": "0", "%": "0.0%", "x": "0.00x", "flag": "0"}.get(unit))
        for i in range(8, 44):
            c = L(i)
            f = sweep_orig[c] if r == sweep else (line.replace("{c}", c) if line else None)
            if f:
                like(htm[f"{c}{r}"], htm[f"{c}{model}"], f, {"%": "0.00%", "flag": "0"}.get(unit))
    for i in range(8, 44):                                             # the debt block reads the helpers
        c = L(i)
        htm[f"{c}77"] = f"=$E$75*{c}7+{c}{draw}"
        htm[f"{c}81"] = f"=IF({c}10>0,IF({c}9=0,({c}76+{c}77+{c}79)*{c}{rate}*{c}10,{c}76*{c}{rate}),0)"
        htm[f"{c}83"] = f"={c}{old}+{c}{new}"
        htm[f"{c}85"] = f"{htm[f'{c}85'].value}+{c}{fee}"
        htm[f"{c}87"] = f"={c}{sweep}+{c}{repaid}"
    htm["C87"] = "Prepayment – cash sweep (Inputs E269) and the old loan repaid at the refinancing (block 18)"
    htm["G158"] = f'=IF(ABS($G${REFI + 14})<0.01,"OK","CHECK")'
    htm["E82"] = f"=IF(G77+G79=0,0,PMT($E$74,Inputs!$E$220,-(G77+G79-$G${draw})))"   # old loan annuity
    htm["G157"] = f"=G77+G79-$G${draw}"                                            # debt at COD


def _sweep_columns_check(htm):
    """The original sweep formula must be written per column (it refers to its own column only)."""
    assert re.fullmatch(r"=IF\(AND\(H\$9>=1,H\$9<=Inputs!\$E\$270\),.*", htm["H87"].value)


# =========================================================================================== §16 Dashboard
def dashboard_refi(wb):
    """Four key-assumption rows (above the closing note) and four returns-table rows (after "Debt at COD")."""
    d = wb["Dashboard"]
    rows = {d.cell(r, 3).value: r for r in range(1, d.max_row + 1) if d.cell(r, 3).value}
    note = next(r for lab, r in rows.items() if str(lab).startswith("Wind built = bid"))
    insert_rows(wb, "Dashboard", note, 4)
    model = note - 1
    for r, (label, unit, f, fmt) in enumerate((
            ("Refinancing year (case 3)", "year", f"=HTM!$G${REFI + 1}", "0"),
            ("Refinancing: new loan rate", "%", "=In_Refi_Rate", "0.00%"),
            ("Refinancing: new loan tenor", "years", "=In_Refi_Tenor", "0"),
            ("Refinancing: gearing added to the base gearing", "%", "=In_Refi_Gearing_Add", "0.0%")), note):
        for col, v in ((3, label), (4, unit), (5, f)):
            like(d.cell(r, col), d.cell(model, col), v, fmt if col == 5 else None)

    out = {k: int(next(wb.defined_names[f"Out_{k}"].destinations)[1].split("$")[-1]) for k in ("debt", "dscr", "gearing")}
    debt = out["debt"]
    d.cell(out["dscr"], 3).value = "Minimum DSCR, whole life"
    insert_rows(wb, "Dashboard", debt + 1, 4)
    for r, (key, label, unit, f, fmt) in enumerate((
            ("refi_on", "Refinanced at COD + 1 (1 = yes)", "flag", "=In_Refi_On", "0"),
            ("debt_refi", "Debt after refinancing", "€ m", f"=HTM!$G${REFI + 15}/Thousand", None),
            ("gear_refi", "Gearing after refinancing", "%", f"=HTM!$G${REFI + 16}", None),
            ("dscr_refi", "Minimum DSCR after refinancing", "x", f"=HTM!$G${REFI + 17}", "0.00x")), debt + 1):
        src = {"€ m": debt, "%": out["gearing"], "x": out["dscr"], "flag": debt}[unit]
        for col in range(1, 10):
            like(d.cell(r, col), d.cell(src, col))
        d.cell(r, 3).value, d.cell(r, 4).value, d.cell(r, 5).value = label, unit, f
        d.cell(r, 9).value = f"=IFERROR(INDEX(F{r}:H{r},Inputs!$E$6)-E{r},0)"
        if fmt:
            for col in (5, 6, 7, 8):
                d.cell(r, col).number_format = fmt
        define(wb, f"Out_{key}", f"Dashboard!$E${r}")


def change_log(wb):
    ws, r = wb["Change_Log"], 132
    for col in "BCDEFGHIJ":
        ws[f"{col}{r}"]._style = copy(ws[f"{col}{r - 1}"]._style)
    for col, v in zip("BCDEFGHIJ", (
            f"=B{r - 1}+1", date(2026, 10, 22), "Refinancing at COD + 1", "Inputs; HTM; Dashboard",
            f"Inputs I7:I8, I217:I220, F217, rows {U_REFI - 1}–{U_REFI + 5}; HTM rows 77, 81, 83, 85, 87, C87, G158, "
            f"block 18 (rows {REFI}+); Dashboard returns and key-assumption rows",
            "Case 3 refinances at the end of the year after the COD year: old loan repaid (prepayment row), new loan "
            "at base gearing + 4 pp, 4.5 %, 15-year annuity, upfront fee; release to equity via the debt rows.",
            "Case 3: 64 % at 4.5 % from NTP", "Case 3: 60 % at 5 %, refinanced at COD + 1 to 64 % at 4.5 %",
            "User 22-Oct-2026; SPEC §16")):
        ws[f"{col}{r}"] = v


def key_table(wb, table):
    """Write the Base, Low and High columns of the Dashboard key-input table (columns E, F, G)."""
    sheet, ref = next(wb.defined_names["Dash_Key_Inputs"].destinations)
    first = int(ref.split(":")[0].split("$")[-1])
    for col, name in ((5, "Base"), (6, "Low"), (7, "High")):
        for i, v in enumerate(table.get(name, [])):
            wb[sheet].cell(first + i, col).value = v


def apply_levers(wb, lv):
    """Regression only: write the source levers into the named cells, case cells and gearing steps."""
    for name, v in lv["names"].items():
        sheet, ref = next(wb.defined_names[name].destinations)
        wb[sheet][ref.replace("$", "")] = v
    for ref, v in lv["cells"].items():
        sheet, cell = ref.split("!")
        wb[sheet][cell] = v
    sheet, ref = next(wb.defined_names["Dash_Gearing_Steps"].destinations)
    first = ref.split(":")[0].replace("$", "")
    row, col = int(re.sub(r"\D", "", first)), openpyxl.utils.column_index_from_string(re.sub(r"\d", "", first))
    for i, v in enumerate(lv["gearing_steps"]):
        wb[sheet].cell(row, col + i).value = v
    key_table(wb, {"Base": lv["key_base"]})


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


def edit(levers=None):
    wb = openpyxl.load_workbook(SRC, keep_vba=False)
    _sweep_columns_check(wb["HTM"])
    refi_inputs(wb)
    case3(wb)
    htm_refi(wb)
    dashboard_refi(wb)
    change_log(wb)
    key_table(wb, KEY_TABLE)
    if levers:
        apply_levers(wb, levers)
    return wb


def build(out=HERE / "out", levers=None, name=NAME):
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    xlsx, xlsm = out / f"{name}.xlsx", out / f"{name}.xlsm"
    edit(levers).save(xlsx)
    with tempfile.TemporaryDirectory() as vba:                         # Scenarios.bas only (as is)
        shutil.copy(VBA_DIR / "Scenarios.bas", vba)
        package(xlsx, xlsm, run=["Scenarios.RunAllScenarios"], vba_dir=Path(vba))
    general(xlsm, "HTM", ("G7", "G10", "G14"), like="G8")
    rc = recalc(xlsm, out / f"{name}_recalc.xlsx")
    return {"xlsx": xlsx, "xlsm": xlsm, "recalc": rc}
