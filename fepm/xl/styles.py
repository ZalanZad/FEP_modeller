"""FEP named cell styles and number formats (docs/MODELLING_STANDARDS.md §5)."""
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Side

FONT = "Aptos Narrow"
SIZE = 10

BLUE, GREEN, BLACK, PURPLE, WHITE, DARK = "0000FF", "008000", "000000", "7030A0", "FFFFFF", "3C3C3B"
FEP_DARK_GREEN, SUB_FILL, INPUT_FILL, MACRO_FILL, SEED_FILL = "00563C", "F2F2F2", "FFF2CC", "EDEDED", "F2F2F2"
GREY = "808080"
OK_GREEN, CHECK_RED = "00B050", "FF0000"

FMT = {
    "money": '#,##0;(#,##0);"–"',
    "money1": '#,##0.0;(#,##0.0);"–"',
    "money2": '#,##0.00;(#,##0.00);"–"',
    "pct": '0.0%;(0.0%);"–"',
    "mult": '0.00"x"',
    "date": "dd-mmm-yy",
    "flag": '0;-0;"–"',
    "factor": '0.000;(0.000);"–"',
    "int": "0",
    "text": "@",
    "general": "General",
}

_thin = Side(style="thin", color="BFBFBF")


def _font(color=BLACK, bold=False, italic=False):
    return Font(name=FONT, size=SIZE, color=color, bold=bold, italic=italic)


def _fill(color):
    return PatternFill("solid", fgColor=color)


def _style(name, font, fill=None, border=None, align=None):
    st = NamedStyle(name=name, font=font)
    if fill:
        st.fill = fill
    if border:
        st.border = border
    if align:
        st.alignment = align
    return st


RIGHT = Alignment(horizontal="right")

STYLES = {
    "input": _style("fep_input", _font(BLUE), _fill(INPUT_FILL), Border(_thin, _thin, _thin, _thin), RIGHT),
    "link": _style("fep_link", _font(GREEN), align=RIGHT),
    "calc": _style("fep_calc", _font(BLACK), align=RIGHT),
    "macro": _style("fep_macro", _font(PURPLE), _fill(MACRO_FILL), Border(_thin, _thin, _thin, _thin), RIGHT),
    "seed": _style("fep_seed", _font(BLACK), _fill(SEED_FILL), align=RIGHT),
    "section": _style("fep_section", _font(WHITE, bold=True), _fill(FEP_DARK_GREEN)),
    "subheader": _style("fep_subheader", _font(DARK, bold=True), _fill(SUB_FILL)),
    "total": _style("fep_total", _font(BLACK, bold=True), border=Border(top=Side(style="thin", color=BLACK)),
                    align=RIGHT),
    "unit": _style("fep_unit", _font(GREY, italic=True)),
    "label": _style("fep_label", _font(BLACK)),
    "banner": _style("fep_banner", Font(name=FONT, size=12, bold=True, color=WHITE), _fill(FEP_DARK_GREEN)),
    "header": _style("fep_header", _font(DARK, bold=True), border=Border(bottom=Side(style="medium", color=DARK)),
                     align=RIGHT),
    "end": _style("fep_end", _font(GREY, bold=True)),
}


def register(wb):
    """Add the FEP named styles to a workbook and make Aptos Narrow 10 the default font."""
    for st in STYLES.values():
        if st.name not in wb.named_styles:
            wb.add_named_style(st)
    # Default ("Normal") font for every unstyled cell.
    wb._fonts[0] = _font()  # noqa: SLF001 - openpyxl keeps the default font at index 0
    normal = next(s for s in wb._named_styles if s.name == "Normal")  # noqa: SLF001
    normal.font = _font()


def apply(cell, kind: str, fmt: str | None = None):
    """Apply a named style (`input`, `link`, `calc`, ...) and an optional number format key or code."""
    cell.style = STYLES[kind].name
    if fmt:
        cell.number_format = FMT.get(fmt, fmt)
    return cell
