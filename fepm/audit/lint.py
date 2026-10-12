"""Static lint of a workbook against MODELLING_STANDARDS.md, plus post-recalc value checks.

    python3 -m fepm.audit.lint model.xlsx [recalculated.xlsx]
"""
from __future__ import annotations

import re
import sys
from dataclasses import dataclass

import openpyxl
from openpyxl.formula.tokenizer import Token, Tokenizer
from openpyxl.utils import column_index_from_string as col_idx, range_boundaries

CALC_SHEET = re.compile(r"^Calc(_\d+)?$")           # Calc, Calc_1, Calc_2 ...: direction rule, timeline
CALC_TYPE = ("Checks", "Summary", "Dashboard")      # + calc sheets: no hardcoded numbers
TIMELINE_COL = 10                        # J
ALLOWED_NUMBERS = {0.0, 1.0}
FORBIDDEN = {"OFFSET", "INDIRECT"}
MACRO_NAMES = re.compile(r"(_Paste$|^Stored_Results$|_Stored$)")   # written by vba/*.bas
LEVER_NAMES = re.compile(r"^Dash_")                                # deal levers, on the Dashboard only
MAX_LEN, MAX_DEPTH, MAX_NOTE = 100, 2, 60                          # g-long, g-note (MODELLING_STANDARDS §6.3)
ERRORS = ("#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#N/A", "#NUM!", "#NULL!", "Err:")
_CELL = re.compile(r"(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])")


@dataclass
class Finding:
    rule: str
    where: str
    detail: str

    def __str__(self):
        return f"[{self.rule}] {self.where}: {self.detail}"


def _is_formula(v):
    return isinstance(v, str) and v.startswith("=")


def _operands(formula):
    """(token value) of every range operand in a formula."""
    return [t.value for t in Tokenizer(formula).items if t.type == Token.OPERAND and t.subtype == Token.RANGE]


def _split_sheet(ref):
    if "!" in ref:
        sheet, addr = ref.rsplit("!", 1)
        return sheet.strip("'"), addr
    return None, ref


def to_r1c1(formula: str, row: int, col: int) -> str:
    """Relative form of a formula, so copies along a row compare equal."""
    def sub(m):
        cabs, c, rabs, r = m.groups()
        cc = f"C{col_idx(c)}" if cabs else f"C[{col_idx(c) - col}]"
        rr = f"R{r}" if rabs else f"R[{int(r) - row}]"
        return rr + cc
    out = []
    for t in Tokenizer(formula).items:
        out.append(_CELL.sub(sub, t.value) if t.type == Token.OPERAND and t.subtype == Token.RANGE else t.value)
    return "".join(out)


def is_calc(title):
    return bool(CALC_SHEET.match(title))


def is_calc_type(title):
    return is_calc(title) or title in CALC_TYPE


def is_timeline(title):
    return is_calc(title) or title == "Checks"


def _bid_sheet(title):
    """Bid number of a `Calc_<n>` sheet, else None."""
    m = CALC_SHEET.match(title)
    return int(m.group(1)[1:]) if m and m.group(1) else None


def _constant_cells(wb):
    """Cells on calc-type sheets that may hold numbers: macro-written names (PASTE rows, stored results),
    Dashboard deal levers (`Dash_*`) and the bid number in `Calc_<n>!A1`."""
    cells = {(ws.title, 1, 1) for ws in wb.worksheets if _bid_sheet(ws.title)}
    for name, dn in wb.defined_names.items():
        if MACRO_NAMES.search(name) or LEVER_NAMES.search(name):
            for sheet, rng in dn.destinations:
                if LEVER_NAMES.search(name) and sheet != "Dashboard":
                    continue
                c1, r1, c2, r2 = range_boundaries(rng.replace("$", ""))
                cells |= {(sheet, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1)}
    return cells


def _name_targets(wb):
    out = {}
    for name, dn in wb.defined_names.items():
        dest = list(dn.destinations)
        if len(dest) == 1:
            out[name] = dest[0]
    return out


def lint(path) -> list[Finding]:
    wb = openpyxl.load_workbook(path)
    f: list[Finding] = []
    allowed = _constant_cells(wb)
    names = _name_targets(wb)
    if wb._external_links:  # noqa: SLF001
        f.append(Finding("e-external", "workbook", f"{len(wb._external_links)} external link(s)"))  # noqa: SLF001
    if wb.calculation.iterate:
        f.append(Finding("e-iterate", "workbook", "iterative calculation is on"))
    for name, dn in wb.defined_names.items():
        if "#REF" in (dn.attr_text or "") or "[" in (dn.attr_text or ""):
            f.append(Finding("e-name", name, f"broken or external name {dn.attr_text}"))
    for ws in wb.worksheets:
        for rng in ws.merged_cells.ranges:
            f.append(Finding("e-merged", f"{ws.title}!{rng}", "merged cells"))
        bid = _bid_sheet(ws.title)
        if bid and ws["A1"].value != bid:
            f.append(Finding("a-bid", f"{ws.title}!A1", f"bid number is {ws['A1'].value!r}, expected {bid}"))
        f += _notes(ws)
        rows: dict[int, dict[int, str]] = {}
        for row in ws.iter_rows():
            for cell in row:
                v, where = cell.value, f"{ws.title}!{cell.coordinate}"
                if v is None:
                    continue
                if not _is_formula(v):
                    if (is_calc_type(ws.title) and isinstance(v, (int, float)) and not isinstance(v, bool)
                            and (ws.title, cell.row, cell.column) not in allowed):
                        f.append(Finding("a-hardcode", where, f"number {v} on a calculation sheet"))
                    continue
                try:
                    toks = Tokenizer(v).items
                except Exception as exc:  # noqa: BLE001
                    f.append(Finding("parse", where, str(exc)))
                    continue
                for t in toks:
                    if t.type == Token.OPERAND and t.subtype == Token.NUMBER and float(t.value) not in ALLOWED_NUMBERS:
                        f.append(Finding("b-literal", where, f"literal {t.value} in {v}"))
                    if t.type == Token.FUNC and t.subtype == Token.OPEN and t.value[:-1].upper() in FORBIDDEN:
                        f.append(Finding("e-function", where, f"{t.value[:-1]} in {v}"))
                    if t.type == Token.OPERAND and t.subtype == Token.RANGE and "[" in t.value:
                        f.append(Finding("e-external", where, f"external reference {t.value}"))
                f += _long(where, v, toks)
                f += _chain(wb, ws, cell, v, toks, names)
                if ws.title != "Inputs":
                    for op in _operands(v):
                        sheet, addr = _split_sheet(op)
                        if sheet == "Inputs" and _CELL.search(addr):
                            c = col_idx(_CELL.search(addr).group(2))
                            header = ws.title in ("Summary", "Dashboard") and _CELL.search(addr).group(4) == "7"
                            if 8 <= c <= 15 and not header:
                                f.append(Finding("live-only", where, f"refers to scenario column {op}"))
                if is_calc(ws.title):
                    f += _direction(ws.title, cell.row, cell.column, v, names)
                if is_timeline(ws.title) and cell.column >= TIMELINE_COL:
                    rows.setdefault(cell.row, {})[cell.column] = to_r1c1(v, cell.row, cell.column)
        for r, cols in rows.items():
            forms = set(cols.values())
            if len(forms) > 1:
                first = cols[min(cols)]
                bad = [c for c, x in cols.items() if x != first]
                f.append(Finding("c-row", f"{ws.title}!row {r}",
                                 f"{len(forms)} different formulas on the timeline (first differs at column "
                                 f"{openpyxl.utils.get_column_letter(min(bad))})"))
            elif ws.max_column >= TIMELINE_COL and len(cols) < _timeline_len(ws):
                f.append(Finding("c-row", f"{ws.title}!row {r}", "timeline row is not filled to the last column"))
    f += _same_sheets(wb)
    return f


def _formulas(ws):
    return {(c.row, c.column): c.value for row in ws.iter_rows() for c in row if _is_formula(c.value)}


def _same_sheets(wb) -> list[Finding]:
    """`Calc_2`.. must carry exactly the formula text of `Calc_1` (one finding per differing row)."""
    sheets = sorted((ws for ws in wb.worksheets if _bid_sheet(ws.title)), key=lambda w: _bid_sheet(w.title))
    if len(sheets) < 2:
        return []
    ref = _formulas(sheets[0])
    out = []
    for ws in sheets[1:]:
        cur = _formulas(ws)
        diff = sorted(k for k in ref.keys() | cur.keys() if ref.get(k) != cur.get(k))
        for r in sorted({r for r, _ in diff}):
            c = min(c for rr, c in diff if rr == r)
            out.append(Finding("c-sheet", f"{ws.title}!row {r}",
                               f"formula text differs from {sheets[0].title} (first at column "
                               f"{openpyxl.utils.get_column_letter(c)})"))
    return out


def _long(where, formula, toks) -> list[Finding]:
    """g-long: over MAX_LEN characters or more than MAX_DEPTH nested function levels."""
    out, depth, deepest = [], 0, 0
    for t in toks:
        if t.type == Token.FUNC:
            depth += 1 if t.subtype == Token.OPEN else -1
            deepest = max(deepest, depth)
    if len(formula) > MAX_LEN:
        out.append(Finding("g-long", where, f"{len(formula)} characters (max {MAX_LEN}): {formula}"))
    if deepest > MAX_DEPTH:
        out.append(Finding("g-long", where, f"{deepest} nested function levels (max {MAX_DEPTH}): {formula}"))
    return out


def _link_target(wb, sheet, op, names):
    """(sheet, row, col) of a single-cell operand or single-cell name, else None."""
    s, addr = _split_sheet(op)
    if s is None and op in names:
        s, addr = names[op]
    addr = addr.replace("$", "")
    if ":" in addr or not re.fullmatch(r"[A-Z]{1,3}\d+", addr):
        return None
    s = s or sheet
    if s not in wb.sheetnames:
        return None
    c1, r1, _, _ = range_boundaries(addr)
    return s, r1, c1


def _pure_link(toks):
    ops = [t for t in toks if t.type != Token.WSPACE]
    if len(ops) == 1 and ops[0].type == Token.OPERAND and ops[0].subtype == Token.RANGE:
        return ops[0].value
    return None


def _chain(wb, ws, cell, formula, toks, names) -> list[Finding]:
    """g-chain: a pure link (`=Sheet!A1`, `=Name`) to a cell that is itself a pure link, and a calc row that only
    copies another row in the same column (`={other}`). Corkscrew openings (`={close@prev}`) are exempt."""
    op = _pure_link(toks)
    if op is None:
        return []
    tgt = _link_target(wb, ws.title, op, names)
    if tgt is None:
        return []
    s, r, c = tgt
    where = f"{ws.title}!{cell.coordinate}"
    same_sheet = s == ws.title
    if same_sheet and is_calc(s) and cell.column >= TIMELINE_COL and c == cell.column:
        return [Finding("g-chain", where, f"row only copies row {r} ({formula}); use the original row")]
    if same_sheet and cell.column >= TIMELINE_COL and c == cell.column - 1:
        return []                                         # corkscrew opening = previous closing / seed
    inner = wb[s].cell(r, c).value
    if _is_formula(inner):
        try:
            if _pure_link(Tokenizer(inner).items):
                return [Finding("g-chain", where, f"{formula} links to {s}!{wb[s].cell(r, c).coordinate}, "
                                                  f"itself a link ({inner}); link to the source")]
        except Exception:  # noqa: BLE001
            pass
    return []


def _notes(ws) -> list[Finding]:
    """g-note: source / note cells (Inputs column D, the Dashboard "Source" column) over MAX_NOTE characters."""
    if ws.title == "Inputs":
        col = 4
    elif ws.title == "Dashboard":
        col = next((c.column for c in ws[9] if c.value == "Source"), None)
    else:
        return []
    out = []
    for r in range(1, ws.max_row + 1) if col else ():
        v = ws.cell(r, col).value
        if isinstance(v, str) and not v.startswith("=") and len(v) > MAX_NOTE:
            out.append(Finding("g-note", f"{ws.title}!{ws.cell(r, col).coordinate}",
                               f"{len(v)} characters (max {MAX_NOTE}); cite the source, do not explain"))
    return out


def _timeline_len(ws):
    return sum(1 for c in range(TIMELINE_COL, ws.max_column + 1) if _is_formula(ws.cell(4, c).value))


def _direction(sheet, row, col, formula, names) -> list[Finding]:
    """Calc refers only up / left (E, F, I may also read their own row and any timeline column above)."""
    out = []
    targets = []
    for op in _operands(formula):
        s, addr = _split_sheet(op)
        if s is None and op in names:
            s, addr = names[op]
        if (s or sheet) != sheet:
            continue
        try:
            c1, r1, c2, r2 = range_boundaries(addr.replace("$", ""))
        except ValueError:
            continue
        targets.append((c2, r2, op))
    for c, r, op in targets:
        if col >= TIMELINE_COL:
            bad = c > col or (c == col and r >= row)
        else:
            bad = r > row or (r == row and c == col)
        if bad:
            out.append(Finding("d-direction", f"{sheet}!{openpyxl.utils.get_column_letter(col)}{row}",
                               f"refers right/down to {op}"))
    return out


def lint_values(recalculated) -> list[Finding]:
    """After recalculation: Excel errors anywhere and Master_Check = 0."""
    wb = openpyxl.load_workbook(recalculated, data_only=True)
    f = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if isinstance(cell.value, str) and cell.value.startswith(ERRORS):
                    f.append(Finding("f-error", f"{ws.title}!{cell.coordinate}", cell.value))
    if "Master_Check" not in wb.defined_names:
        f.append(Finding("f-master", "workbook", "no Master_Check name"))
    else:
        sheet, ref = next(wb.defined_names["Master_Check"].destinations)
        v = wb[sheet][ref.replace("$", "")].value
        if v != 0:
            f.append(Finding("f-master", f"{sheet}!{ref}", f"Master_Check = {v}"))
    return f


def report(findings: list[Finding], limit=50) -> str:
    if not findings:
        return "lint: clean"
    lines = [f"lint: {len(findings)} finding(s)"] + [f"  {x}" for x in findings[:limit]]
    if len(findings) > limit:
        lines.append(f"  ... {len(findings) - limit} more")
    return "\n".join(lines)


if __name__ == "__main__":
    fs = lint(sys.argv[1]) + (lint_values(sys.argv[2]) if len(sys.argv) > 2 else [])
    print(report(fs))
    sys.exit(1 if fs else 0)
