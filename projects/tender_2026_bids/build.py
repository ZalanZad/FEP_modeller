"""Build the single-sheet 2026 MEKH wind tender bid summary (out/Tender_2026_bids_by_node.xlsx).

Sources
- inputs/bids.json: the 47 bids from the MEKH bid-opening record (Bontási adatok, 3-Nov-2026), see extract_pdf.py.
- inputs/owners.json: owner / interest group per bidder from company-registry mirrors and press (source in the
  Notes column); an assumption where the registry is silent.
- Scoring bands and weights: Kiírás Annex 3-4 with GYIK Q19/Q20 (docs/domain/tender_2026_rules.md §4.3).
- Node capacities: Kiírás Annex 1 (reference/Kapacitas_palyazati_kiiras_2026_3.pdf p.33).

One sheet, FEP green style (fepm/xl/styles.py): blue inputs, black formulas, green section bars.
"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter as L

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from fepm.xl import styles  # noqa: E402

HERE = Path(__file__).parent
OUT = HERE / "out" / "Tender_2026_bids_by_node.xlsx"
FMT = styles.FMT | {"money3": '#,##0.000;(#,##0.000);"–"', "pct2": '0.0"%";(0.0"%");"–"', "x": '0.0"x";(0.0"x");"–"'}

bids = json.loads((HERE / "inputs" / "bids.json").read_text())
owners = json.loads((HERE / "inputs" / "owners.json").read_text())

NODES = {  # Kiírás Annex 1 (p.33): name -> (no, ID, allocable MVA)
    "Szombathely OVIT": (9, "SZHO B 132.00", 200), "Dunaújváros Dél": (2, "DUJD GY1 132.00", 50),
    "Iváncsa": (3, "IVAN B 132.00", 50), "Ács Katódgyár (BAMO)": (1, "BAMO B 132.00", 50),
    "Szombathely Észak": (8, "SZHE 221 22.000", 20), "Kisbér": (4, "KISB 221 22.000", 16),
    "Komárom Ipari Park": (5, "KOIP 221 22.000", 21), "Ócsa": (6, "OCSA 1 220.00", 255),
    "Székesfehérvár Nyugat": (7, "SZFN GY1 132.00", 40), "n/a": (0, "", 0),
}
BANDS = [("x ≤ 2", 0, 0), ("2 < x < 5", 2.01, 10), ("5 ≤ x < 10", 5, 20), ("10 ≤ x < 15", 10, 30),
         ("15 ≤ x < 20", 15, 40), ("20 ≤ x < 25", 20, 50), ("25 ≤ x < 30", 25, 60), ("30 ≤ x < 35", 30, 70),
         ("35 ≤ x < 40", 35, 80), ("40 ≤ x ≤ 50", 40, 90), ("x > 50", 50.01, 100)]


def inner(x):
    return max(p for _, lb, p in BANDS if x >= lb)


def points(b):  # python-side, only to order the rows; the sheet recomputes everything
    if b["mva"] == 0 or b["wind_mw"] == 0:
        return 0
    return inner(round((b["wind_mw"] / b["mva"] - 1) * 100, 2)) * 0.1 + inner(round(b["bess_mw"] / b["wind_mw"] * 100, 2)) * 0.15


wb = Workbook()
styles.register(wb)
ws = wb.active
ws.title = "Bids by node"
ws.sheet_view.showGridLines = False
WIDTHS = {"A": 2, "B": 20, "C": 6, "D": 40, "E": 36, "F": 18, "G": 10, "H": 10, "I": 9, "J": 9, "K": 9, "L": 9,
          "M": 8, "N": 8, "O": 8, "P": 7, "Q": 10, "R": 7, "S": 70}
for c, w in WIDTHS.items():
    ws.column_dimensions[c].width = w
WIDE = 19


def put(cell, value, kind, fmt=None, left=False, wrap=False):
    cell.value = value
    styles.apply(cell, kind, FMT.get(fmt, fmt))
    if left or wrap:
        cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=wrap)
    return cell


def row_style(row, kind, c0=2, c1=WIDE):
    for c in range(c0, c1 + 1):
        styles.apply(ws.cell(row, c), kind)


# ---------------------------------------------------------------- title and assumptions
put(ws["B2"], "2026 MEKH wind capacity tender – bids by grid node (bid-opening record, 3-Nov-2026)", "banner")
row_style(2, "banner")
put(ws["B3"], "Blue = input (record, Annex 1, registry). Black = formula. Points use ÖSZ-3 overbuild and ÖSZ-4 BESS only "
              "(25 of 100); the other criteria are not public, so rank and cap test are indicative.", "unit")

put(ws.cell(5, 2), "Assumptions", "section")
row_style(5, "section", 2, 9)
A = [("ÖSZ-3 weight (overbuild)", "%", 10, "int", "Kiírás Annex 3"),
     ("ÖSZ-4 weight (BESS ratio)", "%", 15, "int", "Kiírás Annex 3"),
     ("ÖSZ-4 basis: 1 = BESS / installed wind (GYIK Q20), 0 = BESS / requested MVA", "flag", 1, "flag", "GYIK Q20"),
     ("Group cap per node (requested MVA, one company group)", "MVA", 100, "money", "Kiírás §V.7; GYIK Q52"),
     ("Bids in the record", "count", len(bids), "int", "Bontási adatok, 47 rows")]
for i, (lab, unit, val, fmt, src) in enumerate(A):
    r = 6 + i
    put(ws.cell(r, 2), lab, "label")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    put(ws.cell(r, 6), unit, "unit")
    put(ws.cell(r, 7), val, "input", fmt)
    put(ws.cell(r, 8), src, "unit")
R_W3, R_W4, R_BASIS, R_GCAP, R_N = 6, 7, 8, 9, 10

put(ws.cell(5, 11), "Points band (inner points; x in %, 2 dp)", "section")
row_style(5, "section", 11, 14)
put(ws.cell(6, 11), "Band", "header")
put(ws.cell(6, 13), "From", "header")
put(ws.cell(6, 14), "Points", "header")
for i, (lab, lb, p) in enumerate(BANDS):
    r = 7 + i
    put(ws.cell(r, 11), lab, "label")
    put(ws.cell(r, 13), lb, "input", "0.00")
    put(ws.cell(r, 14), p, "input", "int")
BAND_LB, BAND_P = f"$M$7:$M${6 + len(BANDS)}", f"$N$7:$N${6 + len(BANDS)}"
put(ws.cell(7 + len(BANDS), 11), "Kiírás Annex 4; GYIK Q20: x = 50 scores 90", "unit")

# ---------------------------------------------------------------- bid table
HDR = ["Node", "#", "Bidder", "Owner / interest group (registry)", "Site", "Requested MVA",
       "Wind installed MWe", "BESS MW", "BESS MWh", "Overbuild %", "BESS ratio %", "ÖSZ-3 pts", "ÖSZ-4 pts",
       "Points", "Rank", "Cum. MVA", "Fits cap", "Notes (duplicates, group cap, registry source)"]
R_HDR = 20
put(ws.cell(R_HDR - 1, 2), "Bids by node – ranked on ÖSZ-3 + ÖSZ-4 points, ties by bontási sorszám", "section")
row_style(R_HDR - 1, "section")
for i, h in enumerate(HDR):
    put(ws.cell(R_HDR, 2 + i), h, "header")
    ws.cell(R_HDR, 2 + i).alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.row_dimensions[R_HDR].height = 42
ws.freeze_panes = ws.cell(R_HDR + 1, 5)

by_node = {}
for b in bids:
    by_node.setdefault(b["node"], []).append(b)
order = ["Szombathely OVIT"] + sorted((n for n in NODES if n not in ("Szombathely OVIT", "n/a") and n in by_node),
                                      key=lambda n: -sum(b["mva"] for b in by_node[n]))
order += [n for n in NODES if n not in order and n != "n/a"] + ["n/a"]

r = R_HDR + 1
data_rows, blocks = [], []  # blocks: (first, last, header_row, total_row)
for node in order:
    no, nid, cap = NODES[node]
    rows = sorted(by_node.get(node, []), key=lambda b: (-points(b), b["no"]))
    hr = r
    row_style(hr, "subheader")
    put(ws.cell(hr, 2), node, "subheader")
    put(ws.cell(hr, 4), f"Node {no} · {nid}" if no else "No node named in the bid form", "subheader")
    put(ws.cell(hr, 5), "Allocable MVA (Annex 1)", "subheader")
    ws.cell(hr, 5).alignment = Alignment(horizontal="right")
    put(ws.cell(hr, 7), cap, "input", "money")
    r += 1
    first = r
    for b in rows:
        o = owners.get(b["name"], {})
        valid = 1 if b["name"] != "-" and node != "n/a" else 0
        put(ws.cell(r, 2), node, "unit")
        put(ws.cell(r, 3), b["no"], "input", "int")
        put(ws.cell(r, 4), b["name"], "input", "text", left=True)
        put(ws.cell(r, 5), o.get("group", "n/a"), "input", "text", left=True)
        put(ws.cell(r, 6), b["site"].split(" ")[0] if b["site"] else "-", "input", "text", left=True)
        for c, k in zip((7, 8, 9, 10), ("mva", "wind_mw", "bess_mw", "bess_mwh")):
            put(ws.cell(r, c), b[k], "input", "money3")
        mva, mw, bmw = f"G{r}", f"H{r}", f"I{r}"
        put(ws.cell(r, 11), f"=IF({mva}=0,0,ROUND(({mw}/{mva}-1)*100,2))", "calc", "pct2")
        put(ws.cell(r, 12), f"=IF({mw}=0,0,ROUND({bmw}/IF($G${R_BASIS}=1,{mw},{mva})*100,2))", "calc", "pct2")
        put(ws.cell(r, 13), f"=INDEX({BAND_P},MATCH(K{r},{BAND_LB},1))*$G${R_W3}/100", "calc", "money1")
        put(ws.cell(r, 14), f"=INDEX({BAND_P},MATCH(L{r},{BAND_LB},1))*$G${R_W4}/100", "calc", "money1")
        put(ws.cell(r, 15), f"=M{r}+N{r}", "calc", "money1").font = Font(name=styles.FONT, size=styles.SIZE, bold=True)
        note = " ".join(x for x in (b["note"], o.get("note", ""), o.get("source", "")) if x)
        put(ws.cell(r, 19), note, "label", "text", wrap=True)
        data_rows.append((r, valid, hr))
        r += 1
    last = r - 1
    tr = r
    row_style(tr, "total")
    put(ws.cell(tr, 3), f"=COUNT(C{first}:C{last})" if rows else 0, "total", "int")
    put(ws.cell(tr, 4), "Total", "total")
    if rows:
        put(ws.cell(tr, 5), f"=SUMPRODUCT(1/COUNTIF(D{first}:D{last},D{first}:D{last}))&\" bidders, \"&"
                            f"SUMPRODUCT(1/COUNTIF(E{first}:E{last},E{first}:E{last}))&\" groups\"", "total")
        ws.cell(tr, 5).alignment = Alignment(horizontal="left")
    for c in (7, 8, 9, 10):
        put(ws.cell(tr, c), f"=SUM({L(c)}{first}:{L(c)}{last})" if rows else 0, "total", "money1")
    put(ws.cell(tr, 11), f"=IF(G{hr}=0,0,G{tr}/G{hr})", "total", "x")
    put(ws.cell(tr, 12), "oversubscribed", "unit")
    if rows and node != "n/a":
        put(ws.cell(tr, 17), f"=SUMIF(R{first}:R{last},1,G{first}:G{last})", "total", "money1")
        put(ws.cell(tr, 18), f"=SUM(R{first}:R{last})", "total", "int")
        put(ws.cell(tr, 19), f"=R{tr}&\" bid(s) fit the \"&G{hr}&\" MVA cap with \"&TEXT(Q{tr},\"0.0\")&\" MVA; \"&"
                             f"TEXT(G{tr}-Q{tr},\"0.0\")&\" MVA requested beyond it\"", "total")
        ws.cell(tr, 19).alignment = Alignment(horizontal="left")
    else:
        put(ws.cell(tr, 17), 0, "total", "money1")
        put(ws.cell(tr, 18), 0, "total", "int")
    blocks.append((first, last, hr, tr))
    r += 2

# rank, cumulative MVA and cap test: one formula per column over the whole table
R_FIRST, R_LAST = data_rows[0][0], data_rows[-1][0]
NODE_R, PTS_R, NO_R, RANK_R, MVA_R = (f"$B${R_FIRST}:$B${R_LAST}", f"$O${R_FIRST}:$O${R_LAST}",
                                      f"$C${R_FIRST}:$C${R_LAST}", f"$P${R_FIRST}:$P${R_LAST}",
                                      f"$G${R_FIRST}:$G${R_LAST}")
for rr, valid, hr in data_rows:
    if not valid:
        for c, fmt in ((16, "int"), (17, "money1"), (18, "flag")):
            put(ws.cell(rr, c), 0, "calc", fmt)
        continue
    put(ws.cell(rr, 16), f"=1+COUNTIFS({NODE_R},B{rr},{PTS_R},\">\"&O{rr})"
                         f"+COUNTIFS({NODE_R},B{rr},{PTS_R},O{rr},{NO_R},\"<\"&C{rr})", "calc", "int")
    put(ws.cell(rr, 17), f"=SUMIFS({MVA_R},{NODE_R},B{rr},{RANK_R},\"<=\"&P{rr})", "calc", "money1")
    put(ws.cell(rr, 18), f"=IF(Q{rr}<=$G${hr},1,0)", "calc", "flag")

# grand total and check
R_GT = r
row_style(R_GT, "total")
put(ws.cell(R_GT, 2), "All nodes", "total")
put(ws.cell(R_GT, 3), "=" + "+".join(f"C{tr}" for *_, tr in blocks), "total", "int")
put(ws.cell(R_GT, 4), "Total", "total")
for c in (7, 8, 9, 10, 17):
    put(ws.cell(R_GT, c), "=" + "+".join(f"{L(c)}{tr}" for *_, tr in blocks), "total", "money1")
put(ws.cell(R_GT, 18), "=" + "+".join(f"R{tr}" for *_, tr in blocks), "total", "int")
put(ws.cell(R_GT, 11), f"=G{R_GT}/({'+'.join(f'G{hr}' for _, _, hr, _ in blocks)})", "total", "x")
put(ws.cell(R_GT, 12), "oversubscribed (all nodes)", "unit")

R_CK = R_GT + 2
put(ws.cell(R_CK, 2), "Check", "label")
put(ws.cell(R_CK, 4), "Bid count = record; one rank per valid bid; MVA within cap ≤ allocable at every node", "unit")
cap_terms = "+".join(f"(Q{tr}>G{hr})" for _, _, hr, tr in blocks)
check = (f"=ABS(C{R_GT}-$G${R_N})"
         f"+SUMPRODUCT(({RANK_R}>0)*(COUNTIFS({NODE_R},{NODE_R},{RANK_R},{RANK_R})>1))"
         f"+({cap_terms})")
put(ws.cell(R_CK, 7), check, "calc", "money")
put(ws.cell(R_CK, 8), f"=IF(G{R_CK}=0,\"OK\",\"CHECK\")", "calc")
ok = Font(name=styles.FONT, size=styles.SIZE, color=styles.OK_GREEN, bold=True)
bad = Font(name=styles.FONT, size=styles.SIZE, color=styles.CHECK_RED, bold=True)
ws.conditional_formatting.add(f"H{R_CK}", CellIsRule(operator="equal", formula=['"OK"'], font=ok))
ws.conditional_formatting.add(f"H{R_CK}", CellIsRule(operator="equal", formula=['"CHECK"'], font=bad))
ws.conditional_formatting.add(f"R{R_FIRST}:R{R_LAST}", CellIsRule(operator="equal", formula=["1"], font=ok))
put(ws.cell(R_CK + 2, 2), "***END***", "end")

ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = f"{R_HDR}:{R_HDR}"

OUT.parent.mkdir(exist_ok=True)
wb.save(OUT)
print("written", OUT, "| data rows", R_FIRST, "-", R_LAST, "| check row", R_CK)
