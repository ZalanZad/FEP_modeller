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

CALC_SHEETS = ("Calc", "Checks", "Summary")
TIMELINE_SHEETS = ("Calc", "Checks")
TIMELINE_COL = 10                        # J
ROW_RESULT_COLS = (5, 6, 9)              # E, F, I may refer to their own row's timeline
ALLOWED_NUMBERS = {0.0, 1.0}
FORBIDDEN = {"OFFSET", "INDIRECT"}
MACRO_NAMES = re.compile(r"(_Paste|^Stored_Results)$")
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


def _macro_cells(wb):
    """Cells inside names written by macros (PASTE rows, stored results)."""
    cells = set()
    for name, dn in wb.defined_names.items():
        if MACRO_NAMES.search(name):
            for sheet, rng in dn.destinations:
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
    macro = _macro_cells(wb)
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
        rows: dict[int, dict[int, str]] = {}
        for row in ws.iter_rows():
            for cell in row:
                v, where = cell.value, f"{ws.title}!{cell.coordinate}"
                if v is None:
                    continue
                if not _is_formula(v):
                    if (ws.title in CALC_SHEETS and isinstance(v, (int, float)) and not isinstance(v, bool)
                            and (ws.title, cell.row, cell.column) not in macro):
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
                if ws.title != "Inputs":
                    for op in _operands(v):
                        sheet, addr = _split_sheet(op)
                        if sheet == "Inputs" and _CELL.search(addr):
                            c = col_idx(_CELL.search(addr).group(2))
                            if 8 <= c <= 15 and not (ws.title == "Summary" and _CELL.search(addr).group(4) == "7"):
                                f.append(Finding("live-only", where, f"refers to scenario column {op}"))
                if ws.title == "Calc":
                    f += _direction(ws.title, cell.row, cell.column, v, names)
                if ws.title in TIMELINE_SHEETS and cell.column >= TIMELINE_COL:
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
    return f


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
