"""Build the 2026 MEKH wind capacity tender bid summary workbook (out/Tender_2026_bid_summary.xlsx).

Sources
- inputs/bids.json: the 47 bids from the MEKH bid-opening record (Bontási adatok, 3-Nov-2026), see extract_pdf.py.
- Scoring bands and weights: Kiírás Annex 3-4 with GYIK Q19/Q20 (docs/domain/tender_2026_rules.md §4.3); they
  reproduce the user's sheet-3 points (HU_Wind_bids_strategy..., Sheet1) band for band.
- Node capacities and fees: Kiírás Annex 1 (reference/Kapacitas_palyazati_kiiras_2026_3.pdf p.33).

Style: FEP layout (Aptos Narrow 10, inputs blue on an input fill, formulas black, links green,
section bars, one formula per column) in an orange / grey palette.
"""
import json
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

HERE = Path(__file__).parent
OUT = HERE / "out" / "Tender_2026_bid_summary.xlsx"

# ---------------------------------------------------------------- palette and styles
FONT, SIZE = "Aptos Narrow", 10
ORANGE, DARK_ORANGE, LIGHT_ORANGE = "ED7D31", "C55A11", "FBE5D6"
DARK, GREY, MID_GREY, LIGHT_GREY, WHITE = "3C3C3B", "595959", "BFBFBF", "F2F2F2", "FFFFFF"
BLUE, GREEN, BLACK, RED = "0000FF", "00B050", "000000", "FF0000"

FMT = {
    "money": '#,##0;(#,##0);"–"',
    "money1": '#,##0.0;(#,##0.0);"–"',
    "money2": '#,##0.00;(#,##0.00);"–"',
    "money3": '#,##0.000;(#,##0.000);"–"',
    "pct2": '0.00"%";(0.00"%");"–"',
    "mult": '0.00"x";(0.00"x");"–"',
    "int": '0;-0;"–"',
    "flag": '0;-0;"–"',
    "datetime": "dd-mmm-yy hh:mm",
    "text": "@",
}
THIN = Side(style="thin", color=MID_GREY)


def font(color=BLACK, bold=False, italic=False, size=SIZE):
    return Font(name=FONT, size=size, color=color, bold=bold, italic=italic)


def fill(color):
    return PatternFill("solid", fgColor=color)


def style(cell, kind, fmt=None, wrap=False, center=False):
    """Apply one of the FEP cell kinds in the orange/grey palette."""
    if kind == "input":
        cell.font, cell.fill = font(BLUE), fill(LIGHT_ORANGE)
        cell.border = Border(THIN, THIN, THIN, THIN)
    elif kind == "link":
        cell.font = font(GREEN)
    elif kind == "calc":
        cell.font = font(BLACK)
    elif kind == "label":
        cell.font = font(BLACK)
    elif kind == "unit":
        cell.font = font(GREY, italic=True)
    elif kind == "note":
        cell.font = font(GREY, italic=True)
    elif kind == "header":
        cell.font = font(DARK, bold=True)
        cell.fill = fill(LIGHT_GREY)
        cell.border = Border(bottom=Side(style="medium", color=ORANGE))
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    elif kind == "section":
        cell.font, cell.fill = font(WHITE, bold=True), fill(ORANGE)
    elif kind == "subheader":
        cell.font, cell.fill = font(DARK, bold=True), fill(LIGHT_GREY)
    elif kind == "banner":
        cell.font, cell.fill = font(WHITE, bold=True, size=12), fill(DARK)
    elif kind == "total":
        cell.font = font(BLACK, bold=True)
        cell.border = Border(top=Side(style="thin", color=DARK))
    elif kind == "end":
        cell.font = font(GREY, bold=True)
    if fmt:
        cell.number_format = FMT.get(fmt, fmt)
    if kind in ("input", "link", "calc", "total") and cell.number_format not in (FMT["text"], "General") and not wrap:
        cell.alignment = Alignment(horizontal="right")
    if wrap:
        cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if center else None)
    return cell


def banner(ws, text, width):
    ws["B2"] = text
    for c in range(2, width + 1):
        style(ws.cell(2, c), "banner")
    ws.row_dimensions[2].height = 20


def section(ws, row, text, width):
    ws.cell(row, 2, text)
    for c in range(2, width + 1):
        style(ws.cell(row, c), "section")


def subheader(ws, row, text, width):
    ws.cell(row, 2, text)
    for c in range(2, width + 1):
        style(ws.cell(row, c), "subheader")


def headers(ws, row, col, names, height=42):
    for i, n in enumerate(names):
        style(ws.cell(row, col + i, n), "header")
    ws.row_dimensions[row].height = height


def sheet_setup(ws, widths, freeze=None):
    ws.sheet_view.showGridLines = False
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    if freeze:
        ws.freeze_panes = freeze


# ---------------------------------------------------------------- data
bids = json.loads((HERE / "inputs" / "bids.json").read_text())
N = len(bids)

NODES = [  # Kiírás Annex 1 (p.33): no, name, ID, district, allocable MVA, indirect fee HUF/MVA, direct fee HUF/connection point
    (1, "Ács Katódgyár (BAMO)", "BAMO B 132.00", "Komáromi járás", 50, 54_830_209, 460_000_000),
    (2, "Dunaújváros Dél", "DUJD GY1 132.00", "Dunaújvárosi járás", 50, 19_826_680, 460_000_000),
    (3, "Iváncsa", "IVAN B 132.00", "Dunaújvárosi járás", 50, 19_826_680, 460_000_000),
    (4, "Kisbér", "KISB 221 22.000", "Kisbéri járás", 16, 97_552_584, 134_150_000),
    (5, "Komárom Ipari Park", "KOIP 221 22.000", "Komáromi járás", 21, 54_690_516, 64_200_000),
    (6, "Ócsa", "OCSA 1 220.00", "Gyáli járás", 255, 32_642_400, 619_998_683),
    (7, "Székesfehérvár Nyugat", "SZFN GY1 132.00", "Székesfehérvári járás", 40, 91_021_127, 460_000_000),
    (8, "Szombathely Észak", "SZHE 221 22.000", "Szombathelyi járás", 20, 128_297_272, 182_350_000),
    (9, "Szombathely OVIT", "SZHO B 132.00", "Szombathelyi járás", 200, 73_071_306, 508_345_804),
]

# ÖSZ-3 / ÖSZ-4 inner points (Kiírás Annex 4 as read in docs/domain/tender_2026_rules.md §4.3).
# x in %, rounded to 2 dp before the lookup, so "2 < x" becomes a 2.01 lower bound and "x > 50" a 50.01 lower bound.
BANDS = [
    ("x ≤ 2 %", 0, 0), ("2 < x < 5 %", 2.01, 10), ("5 ≤ x < 10 %", 5, 20), ("10 ≤ x < 15 %", 10, 30),
    ("15 ≤ x < 20 %", 15, 40), ("20 ≤ x < 25 %", 20, 50), ("25 ≤ x < 30 %", 25, 60), ("30 ≤ x < 35 %", 30, 70),
    ("35 ≤ x < 40 %", 35, 80), ("40 ≤ x ≤ 50 %", 40, 90), ("x > 50 %", 50.01, 100),
]

# Bidder groups: an assumption, not a fact of the record. Basis: shared name stem, consecutive tax numbers,
# and the labels in the user's sheet 3 (ATH Windkraft = ATMP / VTT / AT Wind; Trustify = Euronergy Arrabona).
GROUPS = {
    "MVM Szélenergia": (["MVM Szélenergia"], "same company"),
    "ATH Windkraft (AT)": (["ATMP Szélenergia", "VTT Szél Energia", "AT Wind"], "sheet 3 label; AT Wind name stem"),
    "Euronergy (Elektron Holding)": (["Euronergy Elara", "Euronergy Europa", "Euronergy Thebe", "Euronergy Kore"],
                                     "name stem, tax 331368xx"),
    "Trustify / Euronergy Arrabona": (["Euronergy Arrabona"], "sheet 3 label (2 × 50 MW, 25 % BESS)"),
    "Greenergy": (["Greenergy"], "name stem, tax 331471xx"),
    "Trantor SU": (["Trantor SU"], "name stem, tax 331425xx"),
    "WOC Development": (["WOC"], "name stem"),
    "Hanság / Sárbag-Wind Energetikai": (["Hanság-Wind", "Sárbag-Wind"], "name pattern, tax 327586xx-327587xx"),
    "DALMA / GORDON / HELENA / IMRE Energy": (["DALMA Energy", "GORDON Energy", "HELENA Energy", "IMRE Energy"],
                                              "name pattern, tax 331464xx-331465xx"),
    "Altek Energy (TR)": (["Altek"], "sheet 3 label"),
    "Cossava": (["Cossava"], "single bidder"),
    "Magyar Power": (["Magyar Power"], "single bidder"),
    "Pannon Ventus Energy": (["Pannon Ventus"], "single bidder"),
    "BRK Wind": (["BRK Wind"], "single bidder"),
    "TVI 01 Wind": (["TVI 01"], "single bidder"),
    "8G Energy": (["8G Energy"], "single bidder"),
    "Bloch3 Alfa": (["Bloch3"], "single bidder"),
    "CashlineWindfarm": (["Cashline"], "single bidder"),
    "Future Green": (["Future Green"], "single bidder"),
    "n/a (empty submission)": (["-"], "no bid form in the e-file"),
}


def group_of(name):
    for g, (stems, _) in GROUPS.items():
        if any(name.startswith(s) or s in name for s in stems if s != "-") or (name == "-" and stems == ["-"]):
            return g
    raise ValueError(name)


# ---------------------------------------------------------------- workbook
wb = Workbook()
wb.remove(wb.active)
ws_cover = wb.create_sheet("Cover")
ws_in = wb.create_sheet("Inputs")
ws_b = wb.create_sheet("Bids")
ws_s = wb.create_sheet("Summary")
ws_c = wb.create_sheet("Checks")

# ================================================================ Inputs
ws = ws_in
sheet_setup(ws, {"A": 2, "B": 64, "C": 10, "D": 64, "E": 2, "F": 12, "G": 16, "H": 22, "I": 20, "J": 20, "K": 22})
banner(ws, "2026 MEKH wind capacity tender – bid summary | Inputs", 11)
ws["B3"] = "Blue on orange = hard input with unit and source. Everything else in the workbook is a formula."
style(ws["B3"], "note")

section(ws, 5, "Scoring (Kiírás Annex 3, Annex 4; GYIK Q19, Q20)", 11)
headers(ws, 6, 2, ["Item", "Unit", "Source / note", "", "Value"], height=18)
scoring = [
    ("ÖSZ-3 weight – overbuild (installed wind above requested MVA)", "%", "Kiírás Annex 3", 10, "int"),
    ("ÖSZ-4 weight – BESS ratio", "%", "Kiírás Annex 3", 15, "int"),
    ("ÖSZ-4 basis: 1 = BESS MW / installed wind MW (GYIK Q20); 0 = BESS MW / requested MVA (user's sheet 3)",
     "flag", "GYIK Q20 settles the formula as BESS / installed wind; sheet 3 used MVA", 1, "flag"),
    ("Group cap per node (requested MVA, one company group)", "MVA", "Kiírás §V.7, §VII.36; GYIK Q52", 100, "money"),
    ("Bids listed in the bid-opening record", "count", "Bontási adatok, MEKH, 3-Nov-2026 (47 rows)", N, "int"),
]
for i, (lab, unit, src, val, fmt) in enumerate(scoring):
    r = 7 + i
    style(ws.cell(r, 2, lab), "label")
    style(ws.cell(r, 3, unit), "unit")
    style(ws.cell(r, 4, src), "note")
    style(ws.cell(r, 6, val), "input", fmt)
R_W3, R_W4, R_BASIS, R_GCAP, R_NBIDS = 7, 8, 9, 10, 11

section(ws, 13, "Points band for ÖSZ-3 and ÖSZ-4 (inner points; weighted = inner × weight / 100)", 11)
headers(ws, 14, 2, ["Band (x in %)", "Unit", "Source / note", "", "Lower bound", "Inner points"], height=18)
BAND_R0 = 15
for i, (lab, lb, pts) in enumerate(BANDS):
    r = BAND_R0 + i
    style(ws.cell(r, 2, lab), "label")
    style(ws.cell(r, 3, "%"), "unit")
    style(ws.cell(r, 4, "Kiírás Annex 4; x rounded to 2 dp, bounds inclusive" if i == 0 else ""), "note")
    style(ws.cell(r, 6, lb), "input", "money2" if lb % 1 else "money")
    ws.cell(r, 6).number_format = "0.00"
    style(ws.cell(r, 7, pts), "input", "int")
BAND_R1 = BAND_R0 + len(BANDS) - 1
BAND_LB = f"Inputs!$F${BAND_R0}:$F${BAND_R1}"
BAND_PTS = f"Inputs!$G${BAND_R0}:$G${BAND_R1}"
style(ws.cell(BAND_R1 + 1, 2, "ÖSZ-4 at exactly 50 % scores 90 (GYIK Q20 closes the Annex 4 gap)."), "note")

NODE_R_SEC = BAND_R1 + 3
section(ws, NODE_R_SEC, "Nodes offered (Kiírás Annex 1, p.33)", 11)
headers(ws, NODE_R_SEC + 1, 2, ["Node", "No.", "Node ID", "", "Allocable MVA", "Indirect fee (HUF/MVA)",
                                "Direct fee (HUF/connection point)", "District"], height=44)
NODE_R0 = NODE_R_SEC + 2
for i, (no, name, nid, district, mva, ind, direct) in enumerate(NODES):
    r = NODE_R0 + i
    style(ws.cell(r, 2, name), "input", "text")
    ws.cell(r, 2).alignment = Alignment(horizontal="left")
    style(ws.cell(r, 3, no), "input", "int")
    style(ws.cell(r, 4, nid), "input", "text")
    ws.cell(r, 4).alignment = Alignment(horizontal="left")
    style(ws.cell(r, 6, mva), "input", "money")
    style(ws.cell(r, 7, ind), "input", "money")
    style(ws.cell(r, 8, direct), "input", "money")
    style(ws.cell(r, 9, district), "note")
NODE_R1 = NODE_R0 + len(NODES) - 1
NODE_NAMES = f"Inputs!$B${NODE_R0}:$B${NODE_R1}"
NODE_CAPS = f"Inputs!$F${NODE_R0}:$F${NODE_R1}"
r = NODE_R1 + 1
style(ws.cell(r, 2, "Total allocable"), "total")
style(ws.cell(r, 6, f"=SUM(F{NODE_R0}:F{NODE_R1})"), "total", "money")

GRP_R_SEC = NODE_R1 + 3
section(ws, GRP_R_SEC, "Bidder groups (assumption – affiliation inferred, not stated in the record)", 11)
headers(ws, GRP_R_SEC + 1, 2, ["Group", "", "Basis of the grouping"], height=18)
GRP_R0 = GRP_R_SEC + 2
for i, (g, (_, basis)) in enumerate(GROUPS.items()):
    r = GRP_R0 + i
    style(ws.cell(r, 2, g), "input", "text")
    ws.cell(r, 2).alignment = Alignment(horizontal="left")
    style(ws.cell(r, 4, basis), "note")
GRP_R1 = GRP_R0 + len(GROUPS) - 1
GRP_NAMES = f"Inputs!$B${GRP_R0}:$B${GRP_R1}"
style(ws.cell(GRP_R1 + 2, 2, "***END***"), "end")

# ================================================================ Bids
ws = ws_b
COLS = [  # (header, width, kind, fmt)
    ("Bontási sorszám", 8, "input", "int"),                       # B
    ("Érkeztetési szám", 26, "input", "text"),                    # C
    ("Feladás dátuma", 14, "input", "datetime"),                  # D
    ("Pályázó neve", 38, "input", "text"),                        # E
    ("Adószám", 14, "input", "text"),                             # F
    ("Hálózati csomópont", 22, "input", "text"),                  # G
    ("Csomópont azonosító", 16, "input", "text"),                 # H
    ("Igényelt teljesítmény (MVA)", 11, "input", "money3"),       # I
    ("Szélerőmű beépített (MWe)", 11, "input", "money3"),         # J
    ("BESS teljesítmény (MWe)", 11, "input", "money3"),           # K
    ("BESS kapacitás (MWh)", 11, "input", "money3"),              # L
    ("Telephely járás", 18, "input", "text"),                     # M
    ("Telephely település, hrsz", 44, "input", "text"),           # N
    ("Megjegyzés (MEKH)", 30, "input", "text"),                   # O
    ("Group (assumption, see Inputs)", 30, "input", "text"),      # P
    ("Valid bid form (1 = yes)", 9, "input", "flag"),             # Q
    ("Overbuild x (%)", 10, "calc", "pct2"),                      # R
    ("BESS / installed wind (%)", 10, "calc", "pct2"),            # S
    ("BESS / requested MVA (%)", 10, "calc", "pct2"),             # T
    ("ÖSZ-4 x used (%)", 10, "calc", "pct2"),                     # U
    ("ÖSZ-3 inner points", 9, "calc", "int"),                     # V
    ("ÖSZ-4 inner points", 9, "calc", "int"),                     # W
    ("ÖSZ-3 points (weighted)", 9, "calc", "money1"),             # X
    ("ÖSZ-4 points (weighted)", 9, "calc", "money1"),             # Y
    ("Points ÖSZ-3 + ÖSZ-4", 9, "total", "money1"),               # Z
    ("BESS duration (h)", 9, "calc", "money1"),                   # AA
    ("Rank in node (points, then sorszám)", 9, "calc", "int"),    # AB
    ("Cumulative MVA in node", 11, "calc", "money3"),             # AC
    ("Node allocable MVA", 10, "link", "money"),                  # AD
    ("Fits in node cap (1 = yes)", 9, "calc", "flag"),            # AE
    ("Key (node | rank)", 24, "calc", "text"),                    # AF
]
col = {name: 2 + i for i, (name, *_rest) in enumerate(COLS)}
C = {k: L(v) for k, v in col.items()}  # letter by header
sheet_setup(ws, {"A": 2, **{L(2 + i): w for i, (_, w, *_r) in enumerate(COLS)}}, freeze="F8")
W = len(COLS) + 1
banner(ws, "Bids as opened (MEKH Bontási adatok, 3-Nov-2026) and indicative ÖSZ-3 / ÖSZ-4 scoring", W)
ws["B3"] = ("Columns B–Q are the record as published (blue). Group and the validity flag are assumptions. "
            "Columns R–AF are formulas. The ranking uses only ÖSZ-3 + ÖSZ-4 (25 of 100 points); "
            "the other criteria are not in the record.")
style(ws["B3"], "note")
section(ws, 5, "Bid list", W)
ws.cell(6, 2, "Record (input)")
style(ws.cell(6, 2), "subheader")
for c in range(2, col["Overbuild x (%)"]):
    style(ws.cell(6, c), "subheader")
ws.cell(6, col["Overbuild x (%)"], "Scoring and ranking (formulas)")
for c in range(col["Overbuild x (%)"], W + 1):
    style(ws.cell(6, c), "subheader")
headers(ws, 7, 2, [n for n, *_ in COLS], height=54)
R0 = 8
R1 = R0 + N - 1


def rng(name):
    return f"${C[name]}${R0}:${C[name]}${R1}"


for i, b in enumerate(bids):
    r = R0 + i
    valid = 1 if b["name"] != "-" and b["node"] != "n/a" else 0
    sent = datetime.strptime(b["sent"], "%Y.%m.%d. %H:%M") if b["sent"] else None
    values = [b["no"], b["reg"], sent, b["name"], b["tax"], b["node"], b["node_id"], b["mva"], b["wind_mw"],
              b["bess_mw"], b["bess_mwh"], b["district"], b["site"], b["note"], group_of(b["name"]), valid]
    for j, v in enumerate(values):
        name, _w, kind, fmt = COLS[j]
        c = ws.cell(r, 2 + j, v)
        style(c, kind, fmt)
        if fmt == "text":
            c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=name.startswith("Telephely"))
        if v is None and fmt == "datetime":
            c.value = "-"
            c.alignment = Alignment(horizontal="right")
    mva, mw, bmw, bmwh = (f"{C[k]}{r}" for k in ("Igényelt teljesítmény (MVA)", "Szélerőmű beépített (MWe)",
                                                   "BESS teljesítmény (MWe)", "BESS kapacitás (MWh)"))
    node, val, pts, no, rank = (f"{C[k]}{r}" for k in ("Hálózati csomópont", "Valid bid form (1 = yes)",
                                                        "Points ÖSZ-3 + ÖSZ-4", "Bontási sorszám",
                                                        "Rank in node (points, then sorszám)"))
    f = {
        "Overbuild x (%)": f"=IF({mva}=0,0,ROUND(({mw}/{mva}-1)*100,2))",
        "BESS / installed wind (%)": f"=IF({mw}=0,0,ROUND({bmw}/{mw}*100,2))",
        "BESS / requested MVA (%)": f"=IF({mva}=0,0,ROUND({bmw}/{mva}*100,2))",
        "ÖSZ-4 x used (%)": f"=IF(Inputs!$F${R_BASIS}=1,{C['BESS / installed wind (%)']}{r},{C['BESS / requested MVA (%)']}{r})",
        "ÖSZ-3 inner points": f"=INDEX({BAND_PTS},MATCH({C['Overbuild x (%)']}{r},{BAND_LB},1))",
        "ÖSZ-4 inner points": f"=INDEX({BAND_PTS},MATCH({C['ÖSZ-4 x used (%)']}{r},{BAND_LB},1))",
        "ÖSZ-3 points (weighted)": f"={C['ÖSZ-3 inner points']}{r}*Inputs!$F${R_W3}/100",
        "ÖSZ-4 points (weighted)": f"={C['ÖSZ-4 inner points']}{r}*Inputs!$F${R_W4}/100",
        "Points ÖSZ-3 + ÖSZ-4": f"={C['ÖSZ-3 points (weighted)']}{r}+{C['ÖSZ-4 points (weighted)']}{r}",
        "BESS duration (h)": f"=IF({bmw}=0,0,{bmwh}/{bmw})",
        "Rank in node (points, then sorszám)": (
            f"=IF({val}=0,0,1+COUNTIFS({rng('Hálózati csomópont')},{node},{rng('Valid bid form (1 = yes)')},1,"
            f"{rng('Points ÖSZ-3 + ÖSZ-4')},\">\"&{pts})"
            f"+COUNTIFS({rng('Hálózati csomópont')},{node},{rng('Valid bid form (1 = yes)')},1,"
            f"{rng('Points ÖSZ-3 + ÖSZ-4')},{pts},{rng('Bontási sorszám')},\"<\"&{no}))"),
        "Cumulative MVA in node": (
            f"=IF({val}=0,0,SUMIFS({rng('Igényelt teljesítmény (MVA)')},{rng('Hálózati csomópont')},{node},"
            f"{rng('Rank in node (points, then sorszám)')},\"<=\"&{rank},{rng('Valid bid form (1 = yes)')},1))"),
        "Node allocable MVA": f"=IF({val}=0,0,INDEX({NODE_CAPS},MATCH({node},{NODE_NAMES},0)))",
        "Fits in node cap (1 = yes)": (
            f"=IF(AND({val}=1,{C['Cumulative MVA in node']}{r}<={C['Node allocable MVA']}{r}),1,0)"),
        "Key (node | rank)": f"=IF({val}=0,\"n/a | \"&{no},{node}&\" | \"&{rank})",
    }
    for name, formula in f.items():
        _, _w, kind, fmt = COLS[col[name] - 2]
        c = style(ws.cell(r, col[name], formula), kind, fmt)
        if fmt == "text":
            c.alignment = Alignment(horizontal="left")
    ws.row_dimensions[r].height = 27

RT = R1 + 1
style(ws.cell(RT, 2, f"=COUNT({rng('Bontási sorszám')})"), "total", "int")
style(ws.cell(RT, 3, "Total / count"), "total")
for name in ("Igényelt teljesítmény (MVA)", "Szélerőmű beépített (MWe)", "BESS teljesítmény (MWe)",
             "BESS kapacitás (MWh)", "Valid bid form (1 = yes)", "Fits in node cap (1 = yes)"):
    fmt = COLS[col[name] - 2][3]
    style(ws.cell(RT, col[name], f"=SUM({rng(name)})"), "total", fmt)
style(ws.cell(RT + 2, 2, "***END***"), "end")
ws.auto_filter.ref = f"B7:{L(W)}{R1}"

# ================================================================ Summary
ws = ws_s
sheet_setup(ws, {"A": 2, "B": 36, "C": 12, "D": 12, "E": 12, "F": 12, "G": 12, "H": 12, "I": 12, "J": 12, "K": 12,
                 "L": 11, "M": 10, "N": 11, "O": 10, "P": 10, "Q": 10, "R": 9, "S": 9}, freeze="C7")
WS = 19
banner(ws, "2026 MEKH wind capacity tender – who bid where, and how they score on ÖSZ-3 + ÖSZ-4", WS)
ws["B3"] = ("All figures are formulas on Bids and Inputs. Ranking and the cap test use only the overbuild and "
            "BESS points (25 of 100); local contributions, brownfield, EU content and the other criteria are not public.")
style(ws["B3"], "note")

# ---- Panel 1: by node
section(ws, 5, "1. Bids by node", WS)
hdr1 = ["Node", "Allocable MVA", "Bids", "Valid bids", "Bidders (distinct)", "Groups (distinct)",
        "Requested MVA", "Oversub-scription (x)", "Wind installed (MWe)", "BESS (MWe)", "BESS (MWh)",
        "Bids within cap", "MVA within cap", "Spare MVA after cap", "Best points", "Cut-off points"]
headers(ws, 6, 2, hdr1, height=54)
P1_R0 = 7
B = "Bids!"
r_node, r_val, r_name, r_grp, r_mva, r_mw, r_bmw, r_bmwh, r_fit, r_key, r_pts = (
    B + rng(k) for k in ("Hálózati csomópont", "Valid bid form (1 = yes)", "Pályázó neve", "Group (assumption, see Inputs)",
                         "Igényelt teljesítmény (MVA)", "Szélerőmű beépített (MWe)", "BESS teljesítmény (MWe)",
                         "BESS kapacitás (MWh)", "Fits in node cap (1 = yes)", "Key (node | rank)", "Points ÖSZ-3 + ÖSZ-4"))
for i in range(len(NODES) + 1):
    r = P1_R0 + i
    last = i == len(NODES)
    if last:
        style(ws.cell(r, 2, "n/a"), "input", "text")
        ws.cell(r, 2).alignment = Alignment(horizontal="left")
    else:
        style(ws.cell(r, 2, f"=Inputs!B{NODE_R0 + i}"), "link", "text")
        ws.cell(r, 2).alignment = Alignment(horizontal="left")
    n = f"$B{r}"
    cells = {
        3: (f"=IF({n}=\"n/a\",0,Inputs!F{NODE_R0 + i})" if not last else "=0", "link", "money"),
        4: (f"=COUNTIF({r_node},{n})", "calc", "int"),
        5: (f"=COUNTIFS({r_node},{n},{r_val},1)", "calc", "int"),
        6: (f"=SUMPRODUCT(({r_node}={n})*({r_val}=1)/COUNTIFS({r_node},{r_node},{r_name},{r_name}))", "calc", "int"),
        7: (f"=SUMPRODUCT(({r_node}={n})*({r_val}=1)/COUNTIFS({r_node},{r_node},{r_grp},{r_grp}))", "calc", "int"),
        8: (f"=SUMIFS({r_mva},{r_node},{n})", "calc", "money1"),
        9: (f"=IF(C{r}=0,0,H{r}/C{r})", "calc", "mult"),
        10: (f"=SUMIFS({r_mw},{r_node},{n})", "calc", "money1"),
        11: (f"=SUMIFS({r_bmw},{r_node},{n})", "calc", "money1"),
        12: (f"=SUMIFS({r_bmwh},{r_node},{n})", "calc", "money1"),
        13: (f"=COUNTIFS({r_node},{n},{r_fit},1)", "calc", "int"),
        14: (f"=SUMIFS({r_mva},{r_node},{n},{r_fit},1)", "calc", "money1"),
        15: (f"=C{r}-N{r}", "calc", "money1"),
        16: (f"=IF(E{r}=0,0,INDEX({r_pts},MATCH({n}&\" | 1\",{r_key},0)))", "calc", "money1"),
        17: (f"=IF(M{r}=0,0,INDEX({r_pts},MATCH({n}&\" | \"&M{r},{r_key},0)))", "calc", "money1"),
    }
    for cidx, (formula, kind, fmt) in cells.items():
        style(ws.cell(r, cidx, formula), kind, fmt)
P1_R1 = P1_R0 + len(NODES)
P1_RT = P1_R1 + 1
style(ws.cell(P1_RT, 2, "Total"), "total")
for cidx, fmt in [(3, "money"), (4, "int"), (5, "int"), (8, "money1"), (10, "money1"), (11, "money1"),
                  (12, "money1"), (13, "int"), (14, "money1"), (15, "money1")]:
    style(ws.cell(P1_RT, cidx, f"=SUM({L(cidx)}{P1_R0}:{L(cidx)}{P1_R1})"), "total", fmt)
style(ws.cell(P1_RT, 9, f"=IF(C{P1_RT}=0,0,H{P1_RT}/C{P1_RT})"), "total", "mult")
style(ws.cell(P1_RT + 1, 2, "Oversubscription = requested MVA ÷ allocable MVA. Distinct bidders are counted by "
                            "company name, groups by the assumed grouping on Inputs. Cut-off = points of the last "
                            "bid whose cumulative MVA still fits the node cap (ties broken by bontási sorszám)."),
      "note")

# ---- Panel 2: group × node matrix
P2_SEC = P1_RT + 3
section(ws, P2_SEC, "2. Requested MVA by bidder group and node (red = above the 100 MVA group cap per node)", WS)
hdr2 = ["Group"] + [n[1] for n in NODES] + ["n/a", "Total MVA", "Bids", "Wind installed (MWe)", "BESS (MWe)", "BESS (MWh)",
                                           "Best points"]
style(ws.cell(P2_SEC + 1, 2, "Group cap per node (MVA)"), "label")
style(ws.cell(P2_SEC + 1, 3, f"=Inputs!F{R_GCAP}"), "link", "money")
P2_HDR = P2_SEC + 2
headers(ws, P2_HDR, 2, hdr2, height=54)
P2_R0 = P2_HDR + 1
NG = len(GROUPS)
for i in range(NG):
    r = P2_R0 + i
    style(ws.cell(r, 2, f"=Inputs!B{GRP_R0 + i}"), "link", "text")
    ws.cell(r, 2).alignment = Alignment(horizontal="left")
    g = f"$B{r}"
    for j in range(len(NODES) + 1):
        cidx = 3 + j
        node_ref = f"{L(cidx)}${P2_HDR}"
        style(ws.cell(r, cidx, f"=SUMIFS({r_mva},{r_grp},{g},{r_node},{node_ref})"), "calc", "money1")
    c_tot = 3 + len(NODES) + 1
    style(ws.cell(r, c_tot, f"=SUM(C{r}:{L(c_tot - 1)}{r})"), "total", "money1")
    style(ws.cell(r, c_tot + 1, f"=COUNTIF({r_grp},{g})"), "calc", "int")
    style(ws.cell(r, c_tot + 2, f"=SUMIFS({r_mw},{r_grp},{g})"), "calc", "money1")
    style(ws.cell(r, c_tot + 3, f"=SUMIFS({r_bmw},{r_grp},{g})"), "calc", "money1")
    style(ws.cell(r, c_tot + 4, f"=SUMIFS({r_bmwh},{r_grp},{g})"), "calc", "money1")
    style(ws.cell(r, c_tot + 5, f"=SUMPRODUCT(MAX(({r_grp}={g})*{r_pts}))"), "calc", "money1")
P2_R1 = P2_R0 + NG - 1
P2_RT = P2_R1 + 1
style(ws.cell(P2_RT, 2, "Total"), "total")
for cidx in range(3, c_tot + 5):
    fmt = "int" if cidx == c_tot + 1 else "money1"
    style(ws.cell(P2_RT, cidx, f"=SUM({L(cidx)}{P2_R0}:{L(cidx)}{P2_R1})"), "total", fmt)
matrix = f"C{P2_R0}:{L(2 + len(NODES))}{P2_R1}"
ws.conditional_formatting.add(matrix, CellIsRule(operator="greaterThan", formula=[f"$C${P2_SEC + 1}"],
                                                 font=Font(name=FONT, size=SIZE, color=RED, bold=True)))
style(ws.cell(P2_RT + 1, 2, "Header row of the matrix links to the node names on Inputs (column headers are text "
                            "that the SUMIFS keys on). A cell above 100 MVA means the group must drop a bid at that "
                            "node in the deficiency step, or all its bids there become invalid (Kiírás §VII.36)."),
      "note")
for j, n in enumerate(NODES):  # header cells must equal the node names exactly: write them as links
    c = ws.cell(P2_HDR, 3 + j, f"=Inputs!B{NODE_R0 + j}")
    style(c, "header")
    c.font = font(GREEN, bold=True)

# ---- Panel 3: ranking at a selected node
P3_SEC = P2_RT + 4
section(ws, P3_SEC, "3. Indicative ranking at one node (change the node in the orange cell)", WS)
style(ws.cell(P3_SEC + 1, 2, "Node"), "label")
style(ws.cell(P3_SEC + 1, 3, "Szombathely OVIT"), "input", "text")
ws.cell(P3_SEC + 1, 3).alignment = Alignment(horizontal="left")
ws.merge_cells(start_row=P3_SEC + 1, start_column=3, end_row=P3_SEC + 1, end_column=5)
style(ws.cell(P3_SEC + 1, 6, "Allocable MVA"), "label")
style(ws.cell(P3_SEC + 1, 8, f"=INDEX({NODE_CAPS},MATCH($C${P3_SEC + 1},{NODE_NAMES},0))"), "link", "money")
SEL = f"$C${P3_SEC + 1}"
hdr3 = ["Rank", "Bidder", "Group", "Sorszám", "Requested MVA", "Wind installed (MWe)", "BESS (MWe)", "BESS (MWh)",
        "Overbuild x (%)", "ÖSZ-4 x used (%)", "ÖSZ-3 points", "ÖSZ-4 points", "Points ÖSZ-3 + ÖSZ-4",
        "Cumulative MVA", "Fits in cap (1 = yes)", "Row on Bids"]
headers(ws, P3_SEC + 3, 2, hdr3, height=54)
P3_R0 = P3_SEC + 4
NRANK = 20
pull = [  # (summary col, Bids header, kind, fmt)
    (3, "Pályázó neve", "link", "text"), (4, "Group (assumption, see Inputs)", "link", "text"),
    (5, "Bontási sorszám", "link", "int"), (6, "Igényelt teljesítmény (MVA)", "link", "money3"),
    (7, "Szélerőmű beépített (MWe)", "link", "money3"), (8, "BESS teljesítmény (MWe)", "link", "money3"),
    (9, "BESS kapacitás (MWh)", "link", "money3"), (10, "Overbuild x (%)", "link", "pct2"),
    (11, "ÖSZ-4 x used (%)", "link", "pct2"), (12, "ÖSZ-3 points (weighted)", "link", "money1"),
    (13, "ÖSZ-4 points (weighted)", "link", "money1"), (14, "Points ÖSZ-3 + ÖSZ-4", "link", "money1"),
    (15, "Cumulative MVA in node", "link", "money3"), (16, "Fits in node cap (1 = yes)", "link", "flag"),
]
for i in range(NRANK):
    r = P3_R0 + i
    style(ws.cell(r, 2, i + 1), "input", "int")
    rowref = f"$Q{r}"
    style(ws.cell(r, 17, f"=IFERROR(MATCH({SEL}&\" | \"&$B{r},{r_key},0),0)"), "calc", "int")
    for cidx, hdr, kind, fmt in pull:
        c = style(ws.cell(r, cidx, f"=IF({rowref}=0,\"\",INDEX({B}{rng(hdr)},{rowref}))"), kind, fmt)
        if fmt == "text":
            c.alignment = Alignment(horizontal="left")
P3_R1 = P3_R0 + NRANK - 1
style(ws.cell(P3_R1 + 1, 2, "Rank 1 = most ÖSZ-3 + ÖSZ-4 points; ties by bontási sorszám (the call breaks ties "
                            "differently, on the four highest-weighted criteria). A bid fits while its cumulative "
                            "MVA is within the allocable MVA. Column Q is the row index used by the lookups."),
      "note")
style(ws.cell(P3_R1 + 3, 2, "***END***"), "end")
ws.column_dimensions["Q"].width = 8

# ================================================================ Checks
ws = ws_c
sheet_setup(ws, {"A": 2, "B": 70, "C": 8, "D": 14, "E": 2, "F": 12})
banner(ws, "Checks – every row must be 0", 6)
section(ws, 4, "Checks", 6)
headers(ws, 5, 2, ["Check", "Unit", "Status", "", "Value"], height=18)
S = "Summary!"
checks = [
    ("Bid count on Bids equals the record (Inputs)", f"={B}B{RT}-Inputs!F{R_NBIDS}"),
    ("Requested MVA: node panel total equals the bid list", f"=ROUND({S}H{P1_RT}-{B}{C['Igényelt teljesítmény (MVA)']}{RT},6)"),
    ("Requested MVA: group × node matrix equals the bid list", f"=ROUND({S}{L(c_tot)}{P2_RT}-{B}{C['Igényelt teljesítmény (MVA)']}{RT},6)"),
    ("Bid count: node panel equals the bid list", f"={S}D{P1_RT}-{B}B{RT}"),
    ("Bid count: group matrix equals the bid list", f"={S}{L(c_tot + 1)}{P2_RT}-{B}B{RT}"),
    ("Every bid has a group listed on Inputs", f"=SUMPRODUCT(--(COUNTIF({GRP_NAMES},{r_grp})=0))"),
    ("Every valid bid names a node listed on Inputs", f"=SUMPRODUCT(({r_val}=1)*(COUNTIF({NODE_NAMES},{r_node})=0))"),
    ("Ranking keys are unique (no duplicate node | rank)", f"=SUMPRODUCT(--(COUNTIF({r_key},{r_key})>1))"),
    ("Ranks run 1..n in every node (sum of ranks = n(n+1)/2)",
     f"=SUM({B}{rng('Rank in node (points, then sorszám)')})-SUMPRODUCT({S}E{P1_R0}:E{P1_R1}*({S}E{P1_R0}:E{P1_R1}+1))/2"),
    ("MVA within cap never exceeds the allocable MVA", f"=SUMPRODUCT(--({S}N{P1_R0}:N{P1_R1}>{S}C{P1_R0}:C{P1_R1}))"),
    ("Band table lower bounds ascending", f"=SUMPRODUCT(--({BAND_LB.replace(f'$F${BAND_R0}', f'$F${BAND_R0 + 1}')}<=Inputs!$F${BAND_R0}:$F${BAND_R1 - 1}))"),
]
CK_R0 = 7
for i, (lab, formula) in enumerate(checks):
    r = CK_R0 + i
    style(ws.cell(r, 2, lab), "label")
    style(ws.cell(r, 3, "–"), "unit")
    style(ws.cell(r, 4, f"=IF(ABS(F{r})<0.000001,\"OK\",\"CHECK\")"), "calc")
    ws.cell(r, 4).alignment = Alignment(horizontal="center")
    style(ws.cell(r, 6, formula), "calc", "money3")
CK_R1 = CK_R0 + len(checks) - 1
RM = CK_R1 + 2
style(ws.cell(RM, 2, "Master check (sum of absolute values)"), "total")
style(ws.cell(RM, 4, f"=IF(ABS(F{RM})<0.000001,\"OK\",\"CHECK\")"), "total")
ws.cell(RM, 4).alignment = Alignment(horizontal="center")
style(ws.cell(RM, 6, f"=SUMPRODUCT(ABS(F{CK_R0}:F{CK_R1}))"), "total", "money3")
ws.conditional_formatting.add(f"D{CK_R0}:D{RM}", CellIsRule(operator="equal", formula=['"OK"'], font=Font(name=FONT, size=SIZE, color=GREEN, bold=True)))
ws.conditional_formatting.add(f"D{CK_R0}:D{RM}", CellIsRule(operator="equal", formula=['"CHECK"'], font=Font(name=FONT, size=SIZE, color=RED, bold=True)))
style(ws.cell(RM + 2, 2, "***END***"), "end")

# ================================================================ Cover
ws = ws_cover
sheet_setup(ws, {"A": 2, "B": 28, "C": 100})
banner(ws, "2026 MEKH wind capacity tender – bid-opening summary", 3)
rows = [
    ("Purpose", "Who bid at which grid node, how much capacity they requested, and how the bids compare on the two "
                "criteria the bid-opening record reveals: overbuild (ÖSZ-3) and BESS ratio (ÖSZ-4)."),
    ("Source", "MEKH bid-opening record (Bontási adatok), published 3-Nov-2026: 47 submissions, 45 with a bid form. "
               "Node capacities and fees: Kiírás Annex 1. Scoring bands and weights: Kiírás Annex 3-4, GYIK Q19-Q20."),
    ("Sheets", "Inputs – weights, band table, nodes, bidder groups. Bids – the record plus scoring and in-node ranking. "
               "Summary – by node, group × node matrix, ranking at a selected node. Checks – one master check = 0."),
    ("Scoring shown", "Only ÖSZ-3 (weight 10) and ÖSZ-4 (weight 15), 25 of the 100 points. Local contributions (ÖSZ-1, 2), "
                      "partial award, extra security, brownfield, shared connection and EU content (ÖSZ-5 to 9) are not "
                      "in the record, so the ranking is indicative and not the official one."),
    ("ÖSZ-4 basis", "Default is BESS MW / installed wind MW (GYIK Q20). Inputs!F9 = 0 switches to BESS MW / requested "
                    "MVA, which reproduces the points in the strategy workbook's sheet 3."),
    ("Groups", "Affiliations are inferred from names, consecutive tax numbers and the strategy workbook's labels. "
               "They drive the distinct-group counts and the 100 MVA group-cap flag only."),
    ("Colours", "Blue on light orange = hard input. Black = formula. Green = link to another sheet. Orange bar = section. "
                "Red in the group matrix = above the 100 MVA per-node group cap."),
    ("Built", datetime.now().strftime("%d-%b-%Y") + " with projects/tender_2026_bids/build.py (FEP_modeller)."),
]
for i, (k, v) in enumerate(rows):
    r = 4 + i
    style(ws.cell(r, 2, k), "subheader")
    c = style(ws.cell(r, 3, v), "label")
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = 32
style(ws.cell(4 + len(rows) + 1, 2, "***END***"), "end")

for s in wb.worksheets:
    s.sheet_properties.tabColor = ORANGE if s.title in ("Summary",) else GREY
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToWidth = 1

OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print("written", OUT)
