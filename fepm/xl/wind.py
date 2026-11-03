"""Wind-format workbook builder: the layout of `reference/HU_Wind_bids_271026.xlsx` ("the wind model").

Dashboard-centred, annual, one column per case. The same column grid on every sheet:

    A 2 · B 3 (x on bars) · C label · D unit (grey italic) · E, F, G the cases · H Total / Live / Source · I 2 ·
    J → one column per calendar year

Row 2 holds a plain bold title in C2; panes freeze at E5. Timeline sheets carry the year in row 4
(`=First_Year+COLUMN()-COLUMN($J$1)`) and the year-end date in row 5 (`=DATE(J4+1,1,1)-1`).

    book = Book(["PV only", "PV + 2h", "PV + 4h"], years=40)
    d = book.sheet("Dashboard", "Project X – ...")
    d.bar("Key inputs: the cases", e="cases")
    d.lever("MWp", "PV capacity", "MWp", [63, 63, 63], source="Teaser p1")       # blue, named Dash_MWp
    c = book.sheet("Calc", "...", timeline=True)
    blocks = Block(c, 3).build(title, rows)                                         # one block per case
    book.save("out/model.xlsx")

Templates (`=` formulas with tokens), rendered once per cell:

    {key}          this sheet's row `key`, current column          J15 (timeline) or E15 (a case column)
    {key.E}..{key.H}  the row's E..H cell, column absolute         $E15
    {key.all}      the row's timeline                               $J15:$AW15
    {key.prev}     previous column                                  I15
    {key.cases}    the row's case columns                           $E15:$G15
    {year} {yend}  the timeline header (row 4 year, row 5 year end) J$4, J$5
    {c}            the current column letter                        J  (for cross-sheet refs: "Revenue!{c}$23")
    {self...}      the row being written

Inside a `Block` an unqualified key resolves to the block's own row first. Workbook names are written as text.
Inputs (blue on yellow) are values; everything else is a formula, one template per row across the timeline and one
template across E:G for case columns (`cases=`).
"""
from __future__ import annotations

import re

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import Outline

from . import styles
from .styles import FMT, apply

J = 10                                   # first year column
CASES = (5, 6, 7)                        # E, F, G
H = 8
YEAR_ROW, YEND_ROW = 4, 5
WIDTHS = {"A": 2, "B": 3, "C": 56, "D": 10, "E": 13, "F": 13, "G": 13, "H": 15, "I": 2}
_TOK = re.compile(r"\{(\w+)(?:\.(E|F|G|H|all|prev|cases))?\}")
_LINK = re.compile(r"^=(?:[A-Za-z_][\w.]*|'?[\w ]+'?!\$?[A-Z]{1,3}\$?\d+)$")
UNIT_FMT = {"EUR k": "money", "EUR m": "money1", "%": "pct", "x": "mult", "date": "date", "flag": "flag",
            "MW": "money1", "MWp": "money1", "MVA": "money1", "MWh": "money", "GWh": "money1", "EUR/MWh": "money1",
            "years": "int", "year": "int", "months": "int", "days": "int", "h": "money", "factor": "factor",
            "text": "general", "#": "general", "case": "general", "EUR k/MWp": "money1", "EUR k/MW": "money1",
            "EUR k/MWh": "money", "MWh/MWp": "money", "MWh/MW": "money", "HUF m": "money2", "HUF/EUR": "money",
            "%/yr": "pct", "EUR k/MW/yr": "money1", "EUR k/MWp/yr": "money2", "position": "int"}
GREY_ITALIC = Font(name=styles.FONT, size=styles.SIZE, color=styles.GREY, italic=True)


def fmt_of(unit, fmt=None):
    key = fmt or UNIT_FMT.get(unit, "money")
    return FMT.get(key, key)


class Book:
    """The workbook: sheets, names, save as plain `.xlsx` (no macros)."""

    def __init__(self, cases, *, years: int, first_year_name: str = "First_Year"):
        if len(cases) != len(CASES):
            raise ValueError("the wind format has three case columns (E:G)")
        self.wb = Workbook()
        styles.register(self.wb)
        self.cases = list(cases)
        self.n = years
        self.last = J + years - 1
        self.first_year = first_year_name
        self.names: dict[str, str] = {}
        self.sheets: dict[str, Sheet] = {}
        self._first = True

    def sheet(self, title, heading, *, timeline=False, widths=None, top=None) -> "Sheet":
        ws = self.wb.active if self._first else self.wb.create_sheet(title)
        ws.title = title
        self._first = False
        sh = Sheet(self, ws, heading, timeline, widths or {}, top)
        self.sheets[title] = sh
        return sh

    def define(self, name, ref) -> str:
        if name in self.names:
            raise ValueError(f"name {name} already defined")
        self.names[name] = ref
        self.wb.defined_names[name] = DefinedName(name, attr_text=ref)
        return name

    @property
    def last_col(self):
        return L(self.last)

    def save(self, path):
        path = str(path)
        if not path.endswith(".xlsx"):
            raise ValueError("the wind format is a plain .xlsx (no macros)")
        for sh in self.sheets.values():
            sh.finish()
        self.wb.calculation.fullCalcOnLoad = True
        self.wb.save(path)
        return path


class Sheet:
    """One sheet on the wind grid. `r` is the next free row."""

    def __init__(self, book: Book, ws, heading, timeline, widths, top):
        self.b, self.ws, self.timeline = book, ws, timeline
        self.rows: dict[str, int] = {}
        self.scope = ""
        self.groups: list[tuple[int, int]] = []
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 85
        ws.freeze_panes = "E5"
        ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False)
        for col, w in {**WIDTHS, **widths}.items():
            ws.column_dimensions[col].width = w
        for c in range(J, book.last + 1):
            ws.column_dimensions[L(c)].width = 10
        cell = ws.cell(2, 3, heading)
        cell.font = Font(name=styles.FONT, size=styles.SIZE, bold=True)
        self.r = top or 5
        if timeline:
            self._label(YEAR_ROW, "Year", "year", bold=True)
            self._label(YEND_ROW, "Year end", "date")
            for c in range(J, book.last + 1):
                y = ws.cell(YEAR_ROW, c, f"={book.first_year}+COLUMN()-COLUMN($J$1)")
                apply(y, "calc", "int")
                y.font = Font(name=styles.FONT, size=styles.SIZE, bold=True)
                apply(ws.cell(YEND_ROW, c, f"=DATE({L(c)}{YEAR_ROW}+1,1,1)-1"), "calc", "date")
            self.rows["year"], self.rows["yend"] = YEAR_ROW, YEND_ROW
            self.r = top or 7

    # ---- references -----------------------------------------------------------------------
    def key(self, key) -> str:
        k = self.scope + key
        if k in self.rows:
            return k
        if key in self.rows:
            return key
        raise KeyError(f"{self.ws.title}: unknown row key {key!r}")

    def row_of(self, key) -> int:
        return self.rows[self.key(key)]

    def ref(self, key, col="E", absolute=True, sheet=True) -> str:
        """`Sheet!$E$15` (or a range `col` = "all" → `Sheet!$J$15:$AW$15`, "cases" → `Sheet!$E$15:$G$15`)."""
        r = self.row_of(key)
        p = f"{self.ws.title}!" if sheet else ""
        if col == "all":
            return f"{p}$J${r}:${self.b.last_col}${r}"
        if col == "cases":
            return f"{p}$E${r}:$G${r}"
        return f"{p}${col}${r}" if absolute else f"{p}{col}{r}"

    def col_ref(self, key) -> str:
        """Same-column reference to this sheet's row from another sheet's template: `Revenue!{c}$23`."""
        return f"{self.ws.title}!{{c}}${self.row_of(key)}"

    def render(self, tpl, col, row=None) -> str:
        last = self.b.last_col

        def sub(m):
            key, mod = m.group(1), m.group(2)
            if key == "c":
                return L(col)
            if key == "year":
                return f"{L(col)}${YEAR_ROW}"
            if key == "yend":
                return f"{L(col)}${YEND_ROW}"
            r = row if key == "self" else self.row_of(key)
            if mod in ("E", "F", "G", "H"):
                return f"${mod}{r}"
            if mod == "all":
                return f"$J{r}:${last}{r}"
            if mod == "cases":
                return f"$E{r}:$G{r}"
            if mod == "prev":
                return f"{L(col - 1)}{r}"
            return f"{L(col)}{r}"
        return _TOK.sub(sub, tpl)

    # ---- layout ---------------------------------------------------------------------------
    def _label(self, r, label, unit, bold=False, italic=False):
        c = self.ws.cell(r, 3, label)
        if bold or italic:
            c.font = Font(name=styles.FONT, size=styles.SIZE, bold=bold, italic=italic)
        if unit:
            apply(self.ws.cell(r, 4, unit), "unit")

    def blank(self, n=1):
        self.r += n

    def bar(self, title, *, e=None, h=None, extra=None, key=None, kind="section") -> int:
        """Section bar B..H (green, white bold) with `e` over E:G ("cases" = the case names, a list of three, or one
        value for E), `h` over H, and `extra` = {column: text} beyond H (filled the same)."""
        ws, r = self.ws, self.r
        first = 2 if kind == "section" else 3
        cols = list(range(first, H + 1)) + sorted(_col(c) for c in (extra or {}))
        for c in cols:
            apply(ws.cell(r, c), kind)
        if kind == "section":
            ws.cell(r, 2, "x")
        ws.cell(r, 3, title)
        vals = self.b.cases if e == "cases" else (e if isinstance(e, (list, tuple)) else [e])
        for c, v in zip(CASES, vals or []):
            if v is not None:
                ws.cell(r, c, v).alignment = styles.RIGHT
        if h is not None:
            ws.cell(r, H, h).alignment = styles.RIGHT
        for c, v in (extra or {}).items():
            ws.cell(r, _col(c), v).alignment = styles.RIGHT
        if key:
            self.rows[self.scope + key] = r
        self.r += 1
        return r

    def sub(self, title, **kw) -> int:
        """Sub-header C..H bold on light green (wind C105)."""
        return self.bar(title, kind="subheader", **kw)

    def note(self, text, italic=True):
        c = self.ws.cell(self.r, 3, text)
        c.font = GREY_ITALIC if italic else Font(name=styles.FONT, size=styles.SIZE)
        self.r += 1

    def row(self, key, label, unit, *, t=None, cases=None, e=None, f=None, g=None, h=None, fmt=None, fmts=None,
            bold=False, top=False, italic=False, inputs=(), source=None, src_col="H", extra=None) -> int:
        """One row. `t` = timeline template; `cases` = one template for E:G; `e`..`h` = a template or a value per
        cell; `inputs` = columns written as blue inputs (values); `fmts` = {column: format} overrides;
        `source` = text in `src_col` (grey italic); `extra` = {column: template or value} beyond H."""
        ws, r = self.ws, self.r
        if key is not None:
            k = self.scope + key
            if k in self.rows:
                raise ValueError(f"{ws.title}: duplicate row key {k}")
            self.rows[k] = r
        self._label(r, label, unit, bold=bold, italic=italic)
        base = fmt_of(unit, fmt)
        cells = {}
        if cases is not None:
            cells.update({c: cases for c in CASES})
        for c, v in zip((5, 6, 7, 8), (e, f, g, h)):
            if v is not None:
                cells[c] = v
        for c, v in (extra or {}).items():
            cells[_col(c)] = v
        inp = {_col(c) for c in inputs}
        for c, v in cells.items():
            self._cell(r, c, v, c in inp, fmt_of(unit, (fmts or {}).get(L(c))) if fmts and L(c) in fmts else base,
                       bold, top)
        if t is not None:
            for c in range(J, self.b.last + 1):
                self._cell(r, c, t, False, base, bold, top)
        if source:
            sc = ws.cell(r, _col(src_col), source)
            sc.font = GREY_ITALIC
        self.r += 1
        return r

    def _cell(self, r, c, v, is_input, fmt, bold, top):
        ws = self.ws
        if isinstance(v, str) and v.startswith("="):
            text = self.render(v, c, r)
            cell = ws.cell(r, c, text)
            kind = "link" if (self.ws.title != "Dashboard" and _LINK.match(text) and "!" in text) else "calc"
            apply(cell, "total" if top else kind, fmt)
        else:
            cell = ws.cell(r, c, v)
            apply(cell, "input" if is_input else "calc", fmt)
        if bold:
            f = cell.font
            cell.font = Font(name=f.name, size=f.size, color=f.color, bold=True, italic=f.italic)
        if top and not (isinstance(v, str) and v.startswith("=")):
            cell.border = Border(top=Side(style="thin", color=styles.BLACK))
        return cell

    # ---- inputs and levers ----------------------------------------------------------------
    def input(self, name, label, unit, value, source, *, fmt=None, col="E", text=False) -> str:
        """Blue input in E (or `col`), named `name`, source in H."""
        r = self.row(None, label, unit, **{col.lower(): value}, inputs=(col,), fmt=fmt, source=source)
        if text:
            self.ws.cell(r, _col(col)).number_format = "@"
        return self.b.define(name, f"{self.ws.title}!${col}${r}")

    def lever(self, key, label, unit, values, *, source=None, fmt=None, options=None, src_col="J") -> str:
        """Per-case lever on the Dashboard: blue E:G, named `Dash_<key>` over E:G, optional list validation."""
        r = self.row(key, label, unit, e=values[0], f=values[1], g=values[2], inputs=("E", "F", "G"), fmt=fmt,
                     source=source, src_col=src_col)
        if options:
            self.validate(f"E{r}:G{r}", options)
        return self.b.define(f"Dash_{key}", f"{self.ws.title}!$E${r}:$G${r}")

    def validate(self, cells, options):
        f = options if isinstance(options, str) and not isinstance(options, list) and "," not in options \
            else '"' + (options if isinstance(options, str) else ",".join(str(o) for o in options)) + '"'
        dv = DataValidation(type="list", formula1=f, allow_blank=False)
        self.ws.add_data_validation(dv)
        dv.add(cells)

    def name_row(self, name, key, cols="cases") -> str:
        return self.b.define(name, self.ref(key, cols))

    def name_cell(self, name, key, col="E") -> str:
        return self.b.define(name, self.ref(key, col))

    # ---- master check, groups, end --------------------------------------------------------
    def master_check(self, r=3, label="Master check (0 = OK; checks on the Checks sheet)", ref="Master_Check"):
        """C3 label, H3 `=Master_Check` shown as OK / CHECK, green when 0, red otherwise (wind H3)."""
        self._label(r, label, "flag", bold=True)
        cell = self.ws.cell(r, H, f"={ref}")
        apply(cell, "calc", "check")
        cell.font = Font(name=styles.FONT, size=styles.SIZE, bold=True)
        check_format(self.ws, f"H{r}")

    def group(self, first, last, hidden=True):
        """Outline level 1 (collapsed by default), summary row above (wind rows 162–345)."""
        self.groups.append((first, last))
        for r in range(first, last + 1):
            self.ws.row_dimensions[r].outline_level = 1
            self.ws.row_dimensions[r].hidden = hidden

    def end(self):
        self.r += 1
        c = self.ws.cell(self.r, 3, "***END***")
        apply(c, "end")
        self.r += 1

    def finish(self):
        pass


def check_format(ws, rng):
    green = Font(color=styles.OK_GREEN, bold=True)
    red = Font(color=styles.CHECK_RED, bold=True)
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=["0"], font=green))
    ws.conditional_formatting.add(rng, CellIsRule(operator="notEqual", formula=["0"], font=red))


def _col(c) -> int:
    return c if isinstance(c, int) else __import__("openpyxl").utils.column_index_from_string(c)


# =========================================================================================== Block
class Block:
    """The same row set written `n` times on one sheet, one block per case (wind rows 250–341).

    Each block opens with a section bar whose E cell holds the case number (1, then previous + 1) and H "Total".
    `build(i, blk)` writes the block's rows with plain keys; `blk.case` is the `$E$<bar>` reference and
    `blk.lever(name)` gives `INDEX(name,1,$E$<bar>)`, so every block carries the same formula text up to row numbers.
    Returns one {key: row} dict per block."""

    def __init__(self, sheet: Sheet, n: int = 3):
        self.sh, self.n = sheet, n
        self.blocks: list[dict[str, int]] = []

    def build(self, title, build):
        sh = self.sh
        prev = None
        for i in range(1, self.n + 1):
            sh.scope = f"b{i}_"
            sh.blank()
            r = sh.bar(title(i), h="Total", key="bar")
            num = sh.ws.cell(r, 5, "=1" if prev is None else f"=$E${prev}+1")
            num.number_format = '"Case "0'
            num.alignment = styles.RIGHT
            ctx = _BlockCtx(i, f"$E${r}")
            build(i, ctx)
            self.blocks.append({k[len(sh.scope):]: v for k, v in sh.rows.items() if k.startswith(sh.scope)})
            prev = r
        sh.scope = ""
        return self.blocks


class _BlockCtx:
    def __init__(self, i, case):
        self.i, self.case = i, case

    def lever(self, name):
        return f"INDEX({name},1,{self.case})"


# =========================================================================================== DatedBlock
class DatedBlock:
    """Dated cash flows (wind rows 165–213): amounts per case in E:G, the date in J, days to the first 31-Dec in K,
    the accrual factor to each end date in L, M (closed form act/365 with 31-Dec capitalisation, wind E169:
    (1 + r × days₁/365) × (1 + r)^(full years) × the end-year stub factor) and the "earns the return" flag in N.

    `end()` writes an accrual end row (date in J, its in-year stub factor in the factor column); `line()` a cash flow;
    `claim()` the return due per case; `xirr()` / `moic()` on a contiguous range of lines."""

    DATE, DAYS, FLAG = "J", "K", "N"

    def __init__(self, sheet: Sheet, rate: str, days="Days_In_Year"):
        self.sh, self.rate, self.days = sheet, rate, days
        self.ends: dict[str, tuple[int, str]] = {}          # name -> (row, factor column)
        self.lines: list[int] = []

    def header(self, title, ends=("exit", "COD")):
        extra = {"J": "Date", "K": "Days to 31-Dec"}
        for col, e in zip("LM", ends):
            extra[col] = f"Factor to {e}"
        extra["N"] = "Earns return"
        return self.sh.bar(title, e="cases", extra=extra)

    def end(self, key, label, date, col):
        """Accrual end: `date` (template) in J, stub factor 1 + r × (day of year)/365 in `col`."""
        sh = self.sh
        r = sh.row(key, label, "date", extra={"J": date})
        stub = f"=1+{self.rate}*(J{r}-DATE(YEAR(J{r}),1,1)+1)/{self.days}"
        sh._cell(r, _col(col), stub, False, FMT["accrual"], False, False)
        self.ends[key] = (r, col)
        return r

    def line(self, key, label, amount, date, earns, unit="EUR k", fmt=None):
        """A dated flow: `amount` template for E:G (same formula), `date` template, `earns` = flag template."""
        sh = self.sh
        r = sh.row(key, label, unit, cases=amount, extra={"J": date}, fmt=fmt)
        sh.ws.cell(r, 10).number_format = FMT["date"]
        sh._cell(r, 11, f"=DATE(YEAR(J{r})+1,1,1)-J{r}", False, FMT["int"], False, False)
        for e_row, col in self.ends.values():
            f = (f"=(1+{self.rate}*K{r}/{self.days})*(1+{self.rate})^(YEAR($J${e_row})-YEAR(J{r})-1)"
                 f"*${col}${e_row}")
            sh._cell(r, _col(col), f, False, FMT["accrual"], False, False)
        sh._cell(r, 14, earns, False, FMT["flag"], False, False)
        self.lines.append(r)
        return r

    def span(self, first=None, last=None, col="{c}"):
        a, b = first or self.lines[0], last or self.lines[-1]
        return a, b

    def claim(self, key, label, end_key, first=None, last=None):
        """Return due at an end date per case: SUMPRODUCT(−amount × (factor − 1) × flag) over the lines."""
        a, b = first or self.lines[0], last or self.lines[-1]
        col = self.ends[end_key][1]
        t = f"=SUMPRODUCT(-{{c}}{a}:{{c}}{b}*(${col}${a}:${col}${b}-1)*$N${a}:$N${b})"
        return self.sh.row(key, label, "EUR k", cases=t)

    def check_before_end(self, end_key, first=None, last=None):
        """Formula (no `=`): number of lines earning the return paid in or after the end year (wind check 349)."""
        a, b = first or self.lines[0], last or self.lines[-1]
        e_row = self.ends[end_key][0]
        sh = self.sh.ws.title
        return f"SUMPRODUCT({sh}!$N${a}:$N${b}*(YEAR({sh}!$J${a}:$J${b})>=YEAR({sh}!$J${e_row})))"
