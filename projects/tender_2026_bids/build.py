"""Build the 2026 MEKH wind tender bid summary (out/Kapacitastender_2026_palyazatok.xlsx).

Two sheets, Hungarian labels, plain black on white with FEP green bars:
- "Összesítő": bids by grid node, ranked on ÖSZ-3 + ÖSZ-4 points, with owner / interest group and notes.
- "Inputok": weights, points band, node capacities, sources.

Sources
- inputs/bids.json: the 47 bids from the MEKH bid-opening record (Bontási adatok, 3-Nov-2026), see extract_pdf.py.
- inputs/owners.json: owner / interest group per bid from registry mirrors and press (search snippets; the registry
  sites were not reachable from the build session), with the confidence stated in the label.
- Scoring bands and weights: Kiírás Annex 3-4 with GYIK Q19/Q20 (docs/domain/tender_2026_rules.md §4.3).
- Node capacities: Kiírás Annex 1 (reference/Kapacitas_palyazati_kiiras_2026_3.pdf p.33).
"""
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

sys.path.insert(0, str(Path(__file__).resolve().parents[2]))
from fepm.xl import styles  # noqa: E402

HERE = Path(__file__).parent
OUT = HERE / "out" / "Kapacitastender_2026_palyazatok.xlsx"
FMT = {"n0": '#,##0;(#,##0);"–"', "n1": '#,##0.0;(#,##0.0);"–"', "n3": '#,##0.000;(#,##0.000);"–"',
       "pct": '0.0"%";(0.0"%");"–"', "x": '0.0"x";(0.0"x");"–"', "int": '0;-0;"–"', "n2": "0.00"}
BLACK, GREY, WHITE = "000000", "808080", "FFFFFF"
GREEN, LIGHT = styles.FEP_DARK_GREEN, styles.SUB_FILL
THIN = Side(style="thin", color="BFBFBF")

bids = json.loads((HERE / "inputs" / "bids.json").read_text())
owners = json.loads((HERE / "inputs" / "owners.json").read_text())
groups = json.loads((HERE / "inputs" / "groups.json").read_text())

NODES = [  # Kiírás Annex 1 (p.33): name, no, ID, allocable MVA
    ("Szombathely OVIT", 9, "SZHO B 132.00", 200), ("Dunaújváros Dél", 2, "DUJD GY1 132.00", 50),
    ("Iváncsa", 3, "IVAN B 132.00", 50), ("Ács Katódgyár (BAMO)", 1, "BAMO B 132.00", 50),
    ("Szombathely Észak", 8, "SZHE 221 22.000", 20), ("Kisbér", 4, "KISB 221 22.000", 16),
    ("Komárom Ipari Park", 5, "KOIP 221 22.000", 21), ("Ócsa", 6, "OCSA 1 220.00", 255),
    ("Székesfehérvár Nyugat", 7, "SZFN GY1 132.00", 40),
]
BANDS = [("x ≤ 2", 0, 0), ("2 < x < 5", 2.01, 10), ("5 ≤ x < 10", 5, 20), ("10 ≤ x < 15", 10, 30),
         ("15 ≤ x < 20", 15, 40), ("20 ≤ x < 25", 20, 50), ("25 ≤ x < 30", 25, 60), ("30 ≤ x < 35", 30, 70),
         ("35 ≤ x < 40", 35, 80), ("40 ≤ x ≤ 50", 40, 90), ("x > 50", 50.01, 100)]


def inner(x):
    return max(p for _, lb, p in BANDS if x >= lb)


def points(b):  # python-side, only to order the rows; the sheet recomputes everything
    if b["mva"] == 0 or b["wind_mw"] == 0:
        return 0
    return inner(round((b["wind_mw"] / b["mva"] - 1) * 100, 2)) * 0.1 + inner(round(b["bess_mw"] / b["wind_mw"] * 100, 2)) * 0.15


def font(bold=False, italic=False, color=BLACK, size=styles.SIZE):
    return Font(name=styles.FONT, size=size, bold=bold, italic=italic, color=color)


def put(ws, row, col, value, fmt=None, bold=False, italic=False, color=BLACK, align=None, wrap=False, size=styles.SIZE):
    c = ws.cell(row, col, value)
    c.font = font(bold, italic, color, size)
    if fmt:
        c.number_format = FMT.get(fmt, fmt)
        c.alignment = Alignment(horizontal="right")
    if align or wrap:
        c.alignment = Alignment(horizontal=align, vertical="top", wrap_text=wrap)
    return c


def bar(ws, row, text, c1, kind="section"):
    """FEP green section bar (kind='section') or light green sub-header (kind='sub')."""
    fill = PatternFill("solid", fgColor=GREEN if kind == "section" else LIGHT)
    for c in range(2, c1 + 1):
        cell = ws.cell(row, c)
        cell.fill = fill
        cell.font = font(True, color=WHITE if kind == "section" else styles.DARK)
    ws.cell(row, 2).value = text


def header(ws, row, names, c0=2, height=42):
    for i, n in enumerate(names):
        c = ws.cell(row, c0 + i, n)
        c.font = font(True, color=styles.DARK)
        c.border = Border(bottom=Side(style="medium", color=styles.DARK))
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = height


def total_row(ws, row, c1):
    for c in range(2, c1 + 1):
        cell = ws.cell(row, c)
        cell.font = font(True)
        cell.border = Border(top=Side(style="thin", color=BLACK))


wb = Workbook()
styles.register(wb)
ws = wb.active
ws.title = "Összesítő"
wg = wb.create_sheet("Cégcsoportok")
wi = wb.create_sheet("Inputok")
for s in (ws, wg, wi):
    s.sheet_view.showGridLines = False
    s.page_setup.orientation = "landscape"
    s.page_setup.fitToWidth = 1
    s.page_setup.fitToHeight = 0
    s.sheet_properties.pageSetUpPr.fitToPage = True

# ================================================================ Inputok
for c, w in {"A": 2, "B": 58, "C": 10, "D": 12, "E": 12, "F": 14, "G": 60}.items():
    wi.column_dimensions[c].width = w
WI = 7
put(wi, 2, 2, "Inputok és források", bold=True, color=WHITE, size=12)
bar(wi, 2, "Inputok és források", WI)
wi.cell(2, 2).font = font(True, color=WHITE, size=12)

bar(wi, 4, "Pontozás (Kiírás 3. és 4. melléklet; GYIK Q19, Q20)", WI)
header(wi, 5, ["Tétel", "Egység", "Érték", "", "", "Forrás"], height=18)
A = [("ÖSZ-3 súly – túlépítés (beépített szél az igényelt MVA fölött)", "%", 10, "int", "Kiírás 3. melléklet"),
     ("ÖSZ-4 súly – BESS-arány", "%", 15, "int", "Kiírás 3. melléklet"),
     ("ÖSZ-4 alap: 1 = BESS MW / beépített szél MW (GYIK Q20); 0 = BESS MW / igényelt MVA", "kapcsoló", 1, "int",
      "GYIK Q20 szerint a beépített szél az alap; a stratégiai munkafüzet 3. lapja MVA-ra vetített"),
     ("Csoportlimit csomópontonként (igényelt MVA, egy társaságcsoport)", "MVA", 100, "n0", "Kiírás V.7, VII.36; GYIK Q52"),
     ("Összesített nyertességi plafon csoportonként", "MVA", 200, "n0", "Kiírás V.6"),
     ("Pályázatok a bontási jegyzőkönyvben", "db", len(bids), "int", "MEKH Bontási adatok, 2026.11.03. (47 sor)")]
for i, (lab, unit, val, fmt, src) in enumerate(A):
    r = 6 + i
    put(wi, r, 2, lab)
    put(wi, r, 3, unit, italic=True, color=GREY)
    put(wi, r, 4, val, fmt)
    put(wi, r, 7, src, italic=True, color=GREY)
R_W3, R_W4, R_BASIS, R_GCAP, R_TCAP, R_N = 6, 7, 8, 9, 10, 11
I = "Inputok!"
W3, W4, BASIS, GCAP, NBIDS = (f"{I}$D${r}" for r in (R_W3, R_W4, R_BASIS, R_GCAP, R_N))

bar(wi, 13, "Pontsáv ÖSZ-3 és ÖSZ-4 (belső pont; súlyozott = belső × súly / 100; x %-ban, 2 tizedesre kerekítve)", WI)
header(wi, 14, ["Sáv (x %)", "", "Alsó határ", "Belső pont", "", "Forrás"], height=18)
B0 = 15
for i, (lab, lb, p) in enumerate(BANDS):
    r = B0 + i
    put(wi, r, 2, lab)
    put(wi, r, 4, lb, "n2")
    put(wi, r, 5, p, "int")
    if i == 0:
        put(wi, r, 7, "Kiírás 4. melléklet; ÖSZ-4-nél x = 50 pontosan 90 pont (GYIK Q20)", italic=True, color=GREY)
B1 = B0 + len(BANDS) - 1
BAND_LB, BAND_P = f"{I}$D${B0}:$D${B1}", f"{I}$E${B0}:$E${B1}"

N0 = B1 + 3
bar(wi, N0 - 1, "Csomópontok (Kiírás 1. melléklet, 33. o.)", WI)
header(wi, N0, ["Csomópont", "Sorszám", "Azonosító", "", "Kiosztható MVA", "Forrás"], height=18)
cap_cell = {}
for i, (name, no, nid, cap) in enumerate(NODES):
    r = N0 + 1 + i
    put(wi, r, 2, name)
    put(wi, r, 3, no, "int")
    put(wi, r, 4, nid, align="left")
    put(wi, r, 6, cap, "n0")
    put(wi, r, 7, "Kiírás 1. melléklet" if i == 0 else "", italic=True, color=GREY)
    cap_cell[name] = f"{I}$F${r}"
N1 = N0 + len(NODES)
put(wi, N1 + 1, 2, "Összesen kiosztható", bold=True)
put(wi, N1 + 1, 6, f"=SUM(F{N0 + 1}:F{N1})", "n0", bold=True)
total_row(wi, N1 + 1, WI)

S0 = N1 + 4
bar(wi, S0 - 1, "Források és módszer", WI)
notes = [
    "Pályázati adatok: MEKH „Bontási adatok” (közjegyzői bontás 2026.11.02., közzététel 2026.11.03.), 47 sor; a #012 üres e-aktát, a #014 űrlap nélküli beadványt tartalmazott.",
    "Pontozás: csak az ÖSZ-3 (túlépítés, 10 súly) és ÖSZ-4 (BESS-arány, 15 súly) számolható a bontási adatokból, a 100-ból 25 pont. "
    "A többi szempont (helyi hozzájárulás, részleges nyertesség, többletbiztosíték, barnamező, közös csatlakozás, EU-tartalom) nem nyilvános, "
    "ezért a rangsor és a kapacitás-teszt tájékoztató jellegű.",
    "Rangsor: pont szerint csökkenő, holtverseny bontási sorszám szerint (a kiírás a négy legnagyobb súlyú szempont alapján bont holtversenyt). "
    "Egy pályázat akkor „fér be”, ha a csomóponton belüli kumulált igényelt MVA nem haladja meg a kiosztható MVA-t.",
    "Tulajdonos / érdekcsoport: cégadatbázis-tükrök (companywall, Opten, Nemzeti Cégtár) és sajtó keresőkivonatai alapján; a cégjegyzék közvetlenül nem volt elérhető. "
    "A címke jelzi a megbízhatóságot: igazolt / valószínű / feltételezés / nem található. Cégkivonattal ellenőrizendő, mielőtt döntés épül rá.",
    "Csoportlimit: egy társaságcsoport csomópontonként legfeljebb 100 MVA-t igényelhet (Kiírás V.7); túllépés esetén hiánypótlásban kell pályázatot elejteni, különben a csoport minden pályázata érvénytelen azon a csomóponton (VII.36).",
]
for i, t in enumerate(notes):
    put(wi, S0 + i, 2, t, wrap=True)
    wi.merge_cells(start_row=S0 + i, start_column=2, end_row=S0 + i, end_column=WI)
    wi.row_dimensions[S0 + i].height = 44
put(wi, S0 + len(notes) + 1, 2, "***END***", bold=True, color=GREY)


# ================================================================ Cégcsoportok
for c, w in {"A": 2, "B": 52, "C": 15, "D": 50, "E": 20, "F": 9, "G": 11, "H": 90}.items():
    wg.column_dimensions[c].width = w
WG = 8
bar(wg, 2, "Projektcégek és cégcsoportok – ki áll a pályázó mögött", WG)
wg.cell(2, 2).font = font(True, color=WHITE, size=12)
put(wg, 3, 2, "Cégadatbázis-tükrök (companywall, Opten, Nemzeti Cégtár) és sajtó keresőkivonatai alapján; a cégjegyzék közvetlenül nem volt elérhető. "
              "Megbízhatóság: igazolt / valószínű / feltételezés / nem található. Cégkivonattal ellenőrizendő.", italic=True, color=GREY)
header(wg, 5, ["Projektcég (pályázó)", "Adószám", "Cégcsoport / tulajdonos", "Megbízhatóság", "Pályázat (db)", "Igényelt MVA", "Alap / forrás"], height=30)
for i, gr in enumerate(groups):
    r = 6 + i
    put(wg, r, 2, gr["company"], align="left")
    put(wg, r, 3, gr["tax"], align="left")
    put(wg, r, 4, gr["group"], align="left")
    put(wg, r, 5, gr["confidence"], align="left")
    put(wg, r, 8, gr["basis"], wrap=True)
    for c in range(2, 8):
        al = wg.cell(r, c).alignment
        wg.cell(r, c).alignment = Alignment(horizontal=al.horizontal, vertical="top")
GR0, GR1 = 6, 5 + len(groups)
put(wg, GR1 + 1, 2, "Összesen", bold=True)
put(wg, GR1 + 1, 6, f"=SUM(F{GR0}:F{GR1})", "int", bold=True)
put(wg, GR1 + 1, 7, f"=SUM(G{GR0}:G{GR1})", "n1", bold=True)
total_row(wg, GR1 + 1, WG)
put(wg, GR1 + 3, 2, "***END***", bold=True, color=GREY)
wg.freeze_panes = "C6"

# ================================================================ Összesítő
WIDTHS = {"A": 2, "B": 7, "C": 44, "D": 36, "E": 16, "F": 10, "G": 10, "H": 9, "I": 9, "J": 9, "K": 9,
          "L": 8, "M": 8, "N": 8, "O": 7, "P": 10, "Q": 8, "R": 110, "S": 18}
for c, w in WIDTHS.items():
    ws.column_dimensions[c].width = w
WS = 19  # column S carries the node name for the formulas
put(ws, 2, 2, "Szélerőművi kapacitáspályázat 2026 – pályázatok csomópontonként (MEKH bontási adatok, 2026.11.03.)",
    bold=True, color=WHITE, size=12)
bar(ws, 2, ws.cell(2, 2).value, WS)
ws.cell(2, 2).font = font(True, color=WHITE, size=12)
put(ws, 3, 2, "Pont = ÖSZ-3 túlépítés + ÖSZ-4 BESS-arány (a 100-ból 25); a többi szempont nem nyilvános, a rangsor tájékoztató. "
              "Súlyok, pontsáv, csomóponti kapacitás és források az Inputok lapon.", italic=True, color=GREY)

HDR = ["Sorsz.", "Pályázó", "Cégcsoport (Cégcsoportok lap)", "Telephely", "Igényelt MVA", "Szél MWe", "BESS MW", "BESS MWh",
       "Túlépítés %", "BESS-arány %", "ÖSZ-3 pont", "ÖSZ-4 pont", "Pont", "Rang", "Kum. MVA", "Befér a kapacitásba", "Megjegyzés (dupla beadás, csoportlimit)", "Csomópont"]
R_HDR = 5
header(ws, R_HDR, HDR)
ws.cell(R_HDR, 19).font = font(True, color=GREY)
ws.freeze_panes = ws.cell(R_HDR + 1, 4)

by_node = {}
for b in bids:
    by_node.setdefault(b["node"], []).append(b)
order = ["Szombathely OVIT"] + sorted((n for n, *_ in NODES if n not in ("Szombathely OVIT",) and n in by_node),
                                      key=lambda n: -sum(b["mva"] for b in by_node[n]))
order += [n for n, *_ in NODES if n not in order] + ["n/a"]
node_meta = {n: (no, nid) for n, no, nid, _ in NODES}

r = R_HDR + 1
data_rows, blocks = [], []
for node in order:
    rows = sorted(by_node.get(node, []), key=lambda b: (-points(b), b["no"]))
    hr = r
    bar(ws, hr, "", WS, kind="sub")
    if node == "n/a":
        put(ws, hr, 3, "Csomópont nélkül (üres vagy űrlap nélküli beadvány)", bold=True, color=styles.DARK)
        put(ws, hr, 6, 0, "n0", bold=True)
    else:
        no, nid = node_meta[node]
        put(ws, hr, 3, f"{node}  ·  {no}. csomópont, {nid}", bold=True, color=styles.DARK)
        put(ws, hr, 5, "Kiosztható MVA", bold=True, color=styles.DARK, align="right")
        put(ws, hr, 6, f"={cap_cell[node]}", "n0", bold=True)
    ws.cell(hr, 3).font = font(True, color=styles.DARK)
    ws.cell(hr, 19).value = node
    ws.cell(hr, 19).font = font(color=GREY)
    r += 1
    first = r
    for b in rows:
        o = owners.get(str(b["no"]), {})
        valid = 1 if b["name"] != "-" and node != "n/a" else 0
        put(ws, r, 2, b["no"], "int")
        put(ws, r, 3, b["name"], align="left")
        put(ws, r, 4, f"=INDEX(Cégcsoportok!$D$6:$D${5 + len(groups)},MATCH(C{r},Cégcsoportok!$B$6:$B${5 + len(groups)},0))", align="left")
        put(ws, r, 5, b["site"].split(" ")[0] if b["site"] else "-", align="left")
        for c, k in zip((6, 7, 8, 9), ("mva", "wind_mw", "bess_mw", "bess_mwh")):
            put(ws, r, c, b[k], "n3")
        mva, mw, bmw = f"F{r}", f"G{r}", f"H{r}"
        put(ws, r, 10, f"=IF({mva}=0,0,ROUND(({mw}/{mva}-1)*100,2))", "pct")
        put(ws, r, 11, f"=IF({mw}=0,0,ROUND({bmw}/IF({BASIS}=1,{mw},{mva})*100,2))", "pct")
        put(ws, r, 12, f"=INDEX({BAND_P},MATCH(J{r},{BAND_LB},1))*{W3}/100", "n1")
        put(ws, r, 13, f"=INDEX({BAND_P},MATCH(K{r},{BAND_LB},1))*{W4}/100", "n1")
        put(ws, r, 14, f"=L{r}+M{r}", "n1", bold=True)
        note = " ".join(x for x in (b["note"], o.get("note", "")) if x)
        put(ws, r, 18, note, align="left")
        put(ws, r, 19, node, color=GREY)
        data_rows.append((r, valid, hr))
        r += 1
    last = r - 1
    tr = r
    total_row(ws, tr, WS)
    put(ws, tr, 2, f"=COUNT(B{first}:B{last})" if rows else 0, "int", bold=True)
    put(ws, tr, 3, "Összesen", bold=True)
    if rows:
        put(ws, tr, 4, f"=SUMPRODUCT(1/COUNTIF(C{first}:C{last},C{first}:C{last}))&\" pályázó, \"&"
                       f"SUMPRODUCT(1/COUNTIF(D{first}:D{last},D{first}:D{last}))&\" csoport\"", bold=True, align="left")
    for c in (6, 7, 8, 9):
        put(ws, tr, c, f"=SUM({L(c)}{first}:{L(c)}{last})" if rows else 0, "n1", bold=True)
    put(ws, tr, 10, f"=IF(F{hr}=0,0,F{tr}/F{hr})", "x", bold=True)
    put(ws, tr, 11, "túljegyzés", italic=True, color=GREY)
    if rows and node != "n/a":
        put(ws, tr, 16, f"=SUMIF(Q{first}:Q{last},1,F{first}:F{last})", "n1", bold=True)
        put(ws, tr, 17, f"=SUM(Q{first}:Q{last})", "int", bold=True)
        put(ws, tr, 18, f"=Q{tr}&\" pályázat fér a \"&F{hr}&\" MVA-ba (\"&TEXT(P{tr},\"0,0\")&\" MVA); \"&"
                        f"TEXT(F{tr}-P{tr},\"0,0\")&\" MVA igény marad fölötte\"", bold=True, align="left")
    else:
        put(ws, tr, 16, 0, "n1", bold=True)
        put(ws, tr, 17, 0, "int", bold=True)
    ws.cell(tr, 19).value = node
    ws.cell(tr, 19).font = font(color=GREY)
    blocks.append((first, last, hr, tr))
    r += 2

R_FIRST, R_LAST = data_rows[0][0], data_rows[-1][0]
NODE_R, PTS_R, NO_R, RANK_R, MVA_R = (f"$S${R_FIRST}:$S${R_LAST}", f"$N${R_FIRST}:$N${R_LAST}",
                                      f"$B${R_FIRST}:$B${R_LAST}", f"$O${R_FIRST}:$O${R_LAST}",
                                      f"$F${R_FIRST}:$F${R_LAST}")
for rr, valid, hr in data_rows:
    if not valid:
        for c, fmt in ((15, "int"), (16, "n1"), (17, "int")):
            put(ws, rr, c, 0, fmt)
        continue
    put(ws, rr, 15, f"=1+COUNTIFS({NODE_R},S{rr},{PTS_R},\">\"&N{rr})"
                    f"+COUNTIFS({NODE_R},S{rr},{PTS_R},N{rr},{NO_R},\"<\"&B{rr})", "int")
    put(ws, rr, 16, f"=SUMIFS({MVA_R},{NODE_R},S{rr},{RANK_R},\"<=\"&O{rr})", "n1")
    put(ws, rr, 17, f"=IF(P{rr}<=$F${hr},1,0)", '"igen";"igen";"–"')

for i in range(len(groups)):
    rg = 6 + i
    put(wg, rg, 6, f"=COUNTIF(Összesítő!$C${R_FIRST}:$C${R_LAST},B{rg})", "int")
    put(wg, rg, 7, f"=SUMIF(Összesítő!$C${R_FIRST}:$C${R_LAST},B{rg},Összesítő!$F${R_FIRST}:$F${R_LAST})", "n1")
    for c in (6, 7):
        wg.cell(rg, c).alignment = Alignment(horizontal="right", vertical="top")

R_GT = r
total_row(ws, R_GT, WS)
put(ws, R_GT, 2, "=" + "+".join(f"B{tr}" for *_, tr in blocks), "int", bold=True)
put(ws, R_GT, 3, "Mind a 9 csomópont összesen", bold=True)
for c in (6, 7, 8, 9, 16):
    put(ws, R_GT, c, "=" + "+".join(f"{L(c)}{tr}" for *_, tr in blocks), "n1", bold=True)
put(ws, R_GT, 17, "=" + "+".join(f"Q{tr}" for *_, tr in blocks), "int", bold=True)
put(ws, R_GT, 10, f"=F{R_GT}/({'+'.join(f'F{hr}' for _, _, hr, _ in blocks)})", "x", bold=True)
put(ws, R_GT, 11, "túljegyzés", italic=True, color=GREY)

R_CK = R_GT + 2
put(ws, R_CK, 2, "Ellenőrzés", bold=True)
put(ws, R_CK, 3, "Darabszám = jegyzőkönyv; egy rang / érvényes pályázat; beférő MVA ≤ kiosztható; a Cégcsoportok lap lefedi minden pályázót",
    italic=True, color=GREY)
cap_terms = "+".join(f"(P{tr}>F{hr})" for _, _, hr, tr in blocks)
put(ws, R_CK, 6, f"=ABS(B{R_GT}-{NBIDS})+SUMPRODUCT(({RANK_R}>0)*(COUNTIFS({NODE_R},{NODE_R},{RANK_R},{RANK_R})>1))"
                 f"+({cap_terms})+ABS(Cégcsoportok!F{GR1 + 1}-B{R_GT})", "n0")
put(ws, R_CK, 7, f"=IF(F{R_CK}=0,\"OK\",\"HIBA\")", bold=True, align="center")
ok = font(True, color=styles.OK_GREEN)
bad = font(True, color=styles.CHECK_RED)
ws.conditional_formatting.add(f"G{R_CK}", CellIsRule(operator="equal", formula=['"OK"'], font=ok))
ws.conditional_formatting.add(f"G{R_CK}", CellIsRule(operator="equal", formula=['"HIBA"'], font=bad))

put(ws, R_CK + 2, 2, "***END***", bold=True, color=GREY)
ws.print_title_rows = f"{R_HDR}:{R_HDR}"

OUT.parent.mkdir(exist_ok=True)
for old in OUT.parent.glob("Tender_2026_*.xlsx"):
    old.unlink()
wb.save(OUT)
print("written", OUT, "| rows", R_FIRST, "-", R_LAST, "| check", R_CK)
