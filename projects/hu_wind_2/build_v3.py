"""HU wind bids v3: operating CIT and a 60 % levered hold case on the user's v3 file (SPEC bora_v5_summary §27).

Base: reference/HU_Wind_bids_v3_281026.xlsx (the user's corrected copy; every cell of it stays as it is). Edited in
LibreOffice through UNO like build_v2.py: rows inserted with Rows.insertByIndex (references, names and formats shift),
new rows formatted from a neighbouring row of the same kind, one formula per timeline row set as one array on J:AS.
  Inputs     relabel row 39 (CIT); new block "Tax and debt (hold case)" after row 40: Tax_Dep_Life, Gearing,
             Debt_Rate, Debt_Tenor.
  Dashboard  per project block, after "Unlevered pre-tax cash flow": tax depreciation, CIT unlevered, post-tax cash
             flow, debt drawn, annuity, interest, principal, closing balance, CIT levered, equity cash flow, DSCR;
             hold summary: five rows after the NPV row (heading relabelled); per-project annual summary: post-tax and
             equity cash flow rows; check "debt repaid by the end of the tenor" into Master_Check.
The exits (RtB, COD) are not touched. Then verify(): recalc, errors, Master_Check, before / after against the
recalculated base (out/rows_v3.json), independent Python recomputation (Bora), lint on the new rows, out/dump_v3.txt.

    python3 projects/hu_wind_2/build_v3.py
"""
import json
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))
from build_lev import LAST, TL, add_name, col, letter, xirr  # noqa: E402
from build_lev2 import Rows, Sheet  # noqa: E402
from fepm.audit.lint import lint, lint_values  # noqa: E402
from fepm.recalc import Office  # noqa: E402

ROOT = HERE.parents[1]
SRC = ROOT / "reference" / "HU_Wind_bids_v3_281026.xlsx"
OUT = HERE / "out" / "HU_Wind_bids_v3.xlsx"
RC = HERE / "out" / "HU_Wind_bids_v3_recalc.xlsx"
BASE_RC = HERE / "out" / "HU_Wind_bids_v3_base_recalc.xlsx"
MAPS = HERE / "out" / "rows_v3.json"
DUMP = HERE / "out" / "dump_v3.txt"
P = ["E", "F", "G"]
NAMES = ["Bora", "Iváncsa", "Dunaújváros"]
CIT_LABEL = "Corporate income tax (operating profit; exit gains)"
HOLD_LABEL = "Hold to maturity (as the RtB buyer; EBITDA, CIT 15 %, optional 60 % debt)"

# name, label, unit, value, template ("pct" | "years"), source note
NEW_INPUTS = [
    ("Tax_Dep_Life", "Tax depreciation life (straight line on the all-in cost to COD, from the COD year)", "years", 20,
     "years", "HU CIT Act Annex 2, 5 % p.a. power plants – to confirm"),
    ("Gearing", "Gearing (debt ÷ all-in cost to COD)", "%", 0.60, "pct", "user, 27-Oct-2026"),
    ("Debt_Rate", "Debt interest rate", "%", 0.05, "pct", "Bora model Inputs G220"),
    ("Debt_Tenor", "Debt tenor (annuity from the first full operating year)", "years", 15, "years",
     "Bora model Inputs G222"),
]

# Timeline rows after "Unlevered pre-tax cash flow": key, label, unit, template, formula, H column.
# {x} this column, {w} the column to the left (I for J: empty, so the opening balance is 0), {a} the all-in cost cell.
TIMELINE = [
    ("dep", "Tax depreciation (straight line on the all-in cost from the COD year)", "EUR k", "money",
     "=-{a}/Tax_Dep_Life*({x}{t}>=0)*({x}{t}<Tax_Dep_Life)", "sum"),
    ("cit", "CIT, unlevered (no loss carry-forward)", "EUR k", "money",
     "=-MAX(0,{x}{eb}+{x}{dep})*CIT_Rate*({x}{op}>0)", "sum"),
    ("post", "Unlevered post-tax cash flow", "EUR k", "bold", "={x}{pre}+{x}{cit}", "sum"),
    ("draw", "Debt drawn at COD (gearing × all-in cost; no interest during construction)", "EUR k", "money",
     "={a}*Gearing*{x}{cod}", "sum"),
    ("ds", "Debt service, annuity (interest + principal; paid, shown positive)", "EUR k", "money",
     "=-PMT(Debt_Rate,Debt_Tenor,{a}*Gearing)*({x}{t}>=1)*({x}{t}<=Debt_Tenor)", "sum"),
    ("int", "Interest (on the opening balance)", "EUR k", "money", "={w}{bal}*Debt_Rate*({x}{t}>=1)", "sum"),
    ("prin", "Principal repaid (annuity − interest)", "EUR k", "money", "={x}{ds}-{x}{int}", "sum"),
    ("bal", "Debt balance, closing", "EUR k", "money", "={w}{bal}+{x}{draw}-{x}{prin}", None),
    ("citl", "CIT, levered (interest deductible; no loss carry-forward)", "EUR k", "money",
     "=-MAX(0,{x}{eb}+{x}{dep}-{x}{int})*CIT_Rate*({x}{op}>0)", "sum"),
    ("eq", "Equity cash flow, post-tax (60 % debt)", "EUR k", "bold", "={x}{pre}+{x}{draw}-{x}{ds}+{x}{citl}", "sum"),
    ("dscr", "DSCR (EBITDA ÷ debt service)", "x", "ratio", '=IF({x}{ds}>0,{x}{eb}/{x}{ds},"")', "min"),
]
SUMMARY = [("s_cit", "CIT over the life, unlevered", "EUR k", "money"),
           ("s_irr", "Project IRR, post-tax", "%", "irr"),
           ("s_debt", "Debt drawn at COD (60 %)", "EUR k", "money"),
           ("s_eirr", "Equity IRR, post-tax (60 % debt)", "%", "irr"),
           ("s_dscr", "Minimum DSCR", "x", "ratio")]
PER_PROJECT = [("post", "Net cash flow, post-tax"), ("eq", "Equity cash flow, post-tax (60 % debt)")]


def set_row(d, r, template):
    """One formula for the whole timeline row: the same template filled for every column, set as one array."""
    d.cell(f"J{r}:{LAST}{r}").setFormulaArray(
        (tuple(template.format(x=x, w=letter(col(x) - 1)).replace(",", ";") for x in TL),))


# ---------------------------------------------------------------------------------------------------------------------
def edit_inputs(doc):
    inp = Sheet(doc, "Inputs")
    m = Rows()
    r_cit = inp.find("Corporate income tax on exit gains")
    r_ex = inp.find("Share deal, participation exemption", exact=False)
    tmpl = {"head": r_ex + 1, "pct": inp.find("Discount rate", exact=False), "years": inp.find("Operating life from COD")}
    assert (r_cit, r_ex, tmpl["head"]) == (39, 40, 41), (r_cit, r_ex, tmpl)
    inp.text(f"C{r_cit}", CIT_LABEL)
    inp.text(f"H{r_cit}", "SPEC §27 (exits exempt by the flag below)")
    inp.insert(r_ex + 1, 1 + len(NEW_INPUTS))
    m.add(r_ex + 1, 1 + len(NEW_INPUTS))
    tmpl = {k: m(v) for k, v in tmpl.items()}
    rows = m.block(r_ex + 1)
    inp.copy_format(tmpl["head"], rows[0], last="H")
    inp.text(f"C{rows[0]}", "Tax and debt (hold case)")
    idx = [doc.Sheets.getByIndex(i).Name for i in range(doc.Sheets.Count)].index("Inputs")
    out = {"head": rows[0]}
    for (name, lab, unit, v, kind, note), r in zip(NEW_INPUTS, rows[1:]):
        inp.copy_format(tmpl[kind], r, last="H")
        inp.text(f"C{r}", lab)
        inp.text(f"D{r}", unit)
        inp.number(f"E{r}", v)
        inp.text(f"H{r}", note)
        add_name(doc, name, idx, col("E"), r)
        out[name] = r
    return m, out


def edit_dashboard(doc):
    d = Sheet(doc, "Dashboard")
    m = Rows()
    f = lambda t, s=1: d.find(t, start=s, exact=False)  # noqa: E731
    a = {"allin": f("All-in cost to COD for a buyer"), "hold": f("Hold to maturity"), "year": d.find("Year", 300),
         "ye": d.find("Year end"), "irr": f("Project IRR, pre-tax"), "moic": f("MOIC, pre-tax", 120),
         "money": f("Construction capex (straight-line"), "chk1": f("Checks (0 = OK)"), "master": d.find("Master check")}
    a["npv"] = f("NPV at the discount rate", a["hold"])
    a["chk_last"] = a["master"] - 1
    s_head = f("Annual cash flow per project")
    blocks = []
    for nm in NAMES:
        h = f(f"{nm}: annual cash flow")
        b = {"t": d.find("Years from the COD year", h), "op": f("Operating share of the year", h),
             "cod": d.find("COD year flag", h), "eb": d.find("EBITDA", h), "pre": d.find("Unlevered pre-tax cash flow", h)}
        b["net"] = d.find("Net cash flow", d.find(nm, s_head))
        blocks.append(b)
    expect = {"allin": 99, "hold": 134, "irr": 139, "npv": 140, "year": 328, "ye": 329, "master": 442}
    assert all(a[k] == v for k, v in expect.items()), {k: a[k] for k in expect}
    assert [(b["t"], b["op"], b["cod"], b["eb"], b["pre"], b["net"]) for b in blocks] == [
        (332 + k, 336 + k, 337 + k, 356 + k, 357 + k, 153 + n) for k, n in ((0, 0), (32, 8), (64, 16))], blocks

    # Inserts, bottom-up (positions in the user's numbering).
    ops = [(a["master"], 1)] + [(b["pre"] + 1, len(TIMELINE)) for b in reversed(blocks)] + \
          [(b["net"] + 1, len(PER_PROJECT)) for b in reversed(blocks)] + [(a["npv"] + 1, len(SUMMARY))]
    for p, n in ops:
        d.insert(p, n)
        m.add(p, n)
    R = {k: m(v) for k, v in a.items()}
    B = [{k: m(v) for k, v in b.items()} for b in blocks]
    tmpl = {"money": R["money"], "bold": B[0]["pre"], "ratio": R["money"], "irr": R["irr"], "moic": R["moic"],
            "npv": R["npv"]}
    d.text(f"C{R['hold']}", HOLD_LABEL)

    # Project timeline rows.
    for c, b, blk in zip(P, B, blocks):
        new = m.block(blk["pre"] + 1)
        b.update({key: r for (key, *_), r in zip(TIMELINE, new)})
        for key, lab, unit, kind, fml, tot in TIMELINE:
            r = b[key]
            d.copy_format(tmpl[kind], r)
            d.text(f"C{r}", lab)
            d.text(f"D{r}", unit)
            refs = {k: v for k, v in b.items()}
            set_row(d, r, fml.replace("{a}", f"${c}${R['allin']}").replace("{x}", "{{x}}").replace("{w}", "{{w}}")
                    .format(**refs))
            if tot == "sum":
                d.formula(f"H{r}", f"=SUM(J{r}:{LAST}{r})")
            elif tot == "min":
                d.formula(f"H{r}", f"=MIN(J{r}:{LAST}{r})")
            if kind == "ratio":
                for ref in (f"H{r}", f"J{r}:{LAST}{r}"):
                    d.cell(ref).NumberFormat = d.cell(f"E{R['moic']}").NumberFormat

    # Hold summary rows after the NPV row; IRRs in the user's array style (row 139: XIRR on the non-zero flows).
    irr_src = {c: d.cell(f"{c}{R['irr']}").getFormula() for c in P}       # "{=...GETXIRR(...)}", API grammar
    for (key, lab, unit, kind), r in zip(SUMMARY, m.block(a["npv"] + 1)):
        d.copy_format({"money": tmpl["npv"], "irr": tmpl["irr"], "ratio": tmpl["moic"]}[kind], r, last="H")
        d.text(f"C{r}", lab)
        d.text(f"D{r}", unit)
        for c, b in zip(P, B):
            if kind == "irr":
                row = b["post"] if key == "s_irr" else b["eq"]
                fml = irr_src[c].strip("{}").replace(f"$J${b['pre']}:$AS${b['pre']}", f"$J${row}:$AS${row}")
                assert fml != irr_src[c].strip("{}"), fml
                d.cell(f"{c}{r}").setArrayFormula(fml)
            else:
                d.formula(f"{c}{r}", "=$H${}".format(b[{"s_cit": "cit", "s_debt": "draw", "s_dscr": "dscr"}[key]]))
        if key in ("s_cit", "s_debt"):
            d.formula(f"H{r}", f"=SUM(E{r}:G{r})")
        R[key] = r

    # Per-project annual summary: post-tax and equity cash flow after "Net cash flow" (cumulative stays pre-tax).
    for b, blk in zip(B, blocks):
        for (key, lab), r in zip(PER_PROJECT, m.block(blk["net"] + 1)):
            d.copy_format(b["net"], r)
            d.text(f"C{r}", lab)
            d.text(f"D{r}", "EUR k")
            set_row(d, r, f"={{x}}{b[key]}")
            d.formula(f"H{r}", f"=SUM(J{r}:{LAST}{r})")
            b["s_" + key] = r

    # Check: debt repaid by the end of the tenor (closing balance in the year with years from COD = tenor).
    r = m.block(a["master"])[0]
    d.copy_format(R["chk_last"], r, last="H")
    d.text(f"C{r}", "Debt repaid by the end of the tenor (closing balance in the last timeline year = 0)")
    d.text(f"D{r}", "flag")
    for c, b in zip(P, B):                                     # no movement after the tenor: last year = tenor end
        d.formula(f"{c}{r}", f"=N(ABS(${LAST}${b['bal']})>Tolerance)")
    d.formula(f"H{r}", f"=SUM(E{r}:G{r})")
    R["check"] = r
    old = d.cell(f"H{R['master']}").getFormula()
    assert old == f"=SUM(H{R['chk1'] + 1}:H{r - 1})", old
    d.formula(f"H{R['master']}", f"=SUM(H{R['chk1'] + 1}:H{r})")
    return m, R, B


def array_cells(path):
    """The user's array formulas (E36:G36, E133:G133, E139:G139), as (sheet, column, row)."""
    wb = openpyxl.load_workbook(path)
    return [(ws.title, c.column_letter, c.row) for ws in wb.worksheets for row in ws.iter_rows() for c in row
            if isinstance(c.value, ArrayFormula)]


def keep_arrays(doc, cells, maps):
    """LibreOffice 24.2 exports an array formula as a plain one once a named range has been added (seen here); set
    each of the user's array formulas again with its own, unchanged formula so it is saved as an array."""
    for sheet, c, r in cells:
        cell = doc.Sheets.getByName(sheet).getCellRangeByName(f"{c}{maps[sheet](r)}")
        f = cell.getFormula()
        assert f.startswith("{=") and f.endswith("}"), (sheet, c, r, f)
        cell.setArrayFormula(f[1:-1])
        assert cell.getFormula() == f, (sheet, c, r)


def restore_filter_case(path):
    """LibreOffice 24.2 does not know FILTER and writes the user's `_xlfn._xlws.FILTER(` in lower case; put it back."""
    tmp = path.with_suffix(".tmp")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.startswith("xl/worksheets/"):
                data = data.replace(b"_xlfn._xlws.filter(", b"_xlfn._xlws.FILTER(")
            zout.writestr(item, data)
    shutil.move(tmp, path)


def build():
    with Office(SRC) as lo:                                   # Office recalculates on load
        lo.save(BASE_RC)
        mi, ri = edit_inputs(lo.doc)
        md, R, B = edit_dashboard(lo.doc)
        keep_arrays(lo.doc, array_cells(SRC), {"Inputs": mi, "Dashboard": md})
        lo.calculate()
        lo.save(OUT)
    restore_filter_case(OUT)
    with Office(OUT) as lo:
        lo.save(RC)
    info = {"Inputs": {"inserts": mi.ins, "rows": ri},
            "Dashboard": {"inserts": md.ins, "rows": R, "blocks": dict(zip(NAMES, B))}}
    MAPS.write_text(json.dumps(info, indent=1, ensure_ascii=False))
    print(f"wrote {OUT}")


# ---------------------------------------------------------------------------------------------------------------------
def remap(sheet):
    m = Rows()
    for p, n in json.loads(MAPS.read_text())[sheet]["inserts"]:
        m.add(p, n)
    return m


def compare(b_wb, a_wb):
    """Every cell of the user's two sheets (values after recalc) against its shifted position in the new file."""
    out = []
    for name in ("Dashboard", "Inputs"):
        b, a, m = b_wb[name], a_wb[name], remap(name)
        for r in range(1, b.max_row + 1):
            for c in range(1, b.max_column + 1):
                x, y = b.cell(r, c).value, a.cell(m(r), c).value
                if not ((x == y) or (isinstance(x, (int, float)) and isinstance(y, (int, float))
                                     and abs(x - y) <= 1e-9 * max(1, abs(x)))):
                    out.append((name, r, m(r), b.cell(r, c).column_letter, x, y))
    return out


def cells_check(b_wb, a_wb):
    """Formulas: same count of formula cells per old row after the shift (a cheap structure check)."""
    bad = []
    for name in ("Dashboard", "Inputs"):
        b, a, m = b_wb[name], a_wb[name], remap(name)
        for r in range(1, b.max_row + 1):
            nb = sum(1 for c in range(1, b.max_column + 1) if b.cell(r, c).value is not None)
            na = sum(1 for c in range(1, b.max_column + 1) if a.cell(m(r), c).value is not None)
            if nb != na:
                bad.append((name, r, m(r), nb, na))
    return bad


def recompute(wb, info):
    """Bora (and the other two), independently: from the recalculated EBITDA, pre-tax CF, years and all-in cost."""
    dash, inp = wb["Dashboard"], wb["Inputs"]
    iv = lambda name: inp[f"E{info['Inputs']['rows'][name]}"].value  # noqa: E731
    cit_r = inp[f"E{next(r for r in range(1, 200) if inp.cell(r, 3).value == CIT_LABEL)}"].value
    life, gear, rate, ten = iv("Tax_Dep_Life"), iv("Gearing"), iv("Debt_Rate"), iv("Debt_Tenor")
    R = info["Dashboard"]["rows"]
    cols = range(col("J") + 1, col(LAST) + 2)
    ye = [dash.cell(R["ye"], c).value.date() for c in cols]
    res, ok = {}, True
    for c, nm in zip(P, NAMES):
        b = info["Dashboard"]["blocks"][nm]
        row = lambda r: [dash.cell(r, k).value for k in cols]  # noqa: E731
        tn, op, eb, pre = row(b["t"]), row(b["op"]), row(b["eb"]), row(b["pre"])
        allin = dash[f"{c}{R['allin']}"].value
        dep = [allin / life * (0 <= n < life) for n in tn]
        cit = [-max(0, e - x) * cit_r * (o > 0) for e, x, o in zip(eb, dep, op)]
        post = [p + t for p, t in zip(pre, cit)]
        debt = gear * allin
        ann = debt * rate / (1 - (1 + rate) ** -ten)
        bal, ints, citl, eq, dscr = 0.0, [], [], [], []
        for n, e, x, o, p in zip(tn, eb, dep, op, pre):
            i = bal * rate if n >= 1 else 0.0
            a_ = ann if 1 <= n <= ten else 0.0
            drawn = debt if n == 0 else 0.0
            bal = bal + drawn - (a_ - i)
            ct = -max(0, e - x - i) * cit_r * (o > 0)
            ints.append(i)
            citl.append(ct)
            eq.append(p + drawn - a_ + ct)
            if a_ > 0:
                dscr.append(e / a_)
            if n == ten:
                bal_end = bal
        nz = lambda v: [(t, x) for t, x in zip(ye, v) if x != 0]  # noqa: E731
        py = {"all-in cost": allin, "EBITDA total": sum(eb), "CIT total": sum(cit), "project IRR pre-tax": xirr(nz(pre)),
              "project IRR post-tax": xirr(nz(post)), "debt at COD": debt, "equity IRR": xirr(nz(eq)),
              "min DSCR": min(dscr), "CIT levered total": sum(citl), "balance at tenor end": bal_end}
        # Model values: summary cells; IRR cells (FILTER arrays) cannot be evaluated by LibreOffice 24.2, so the
        # model's own post-tax / equity rows are run through XIRR here instead.
        mr = lambda r: [dash.cell(r, k).value for k in cols]  # noqa: E731
        model = {"all-in cost": dash[f"{c}{R['hold'] + 1}"].value, "EBITDA total": dash[f"{c}{R['hold'] + 4}"].value,
                 "CIT total": dash[f"{c}{R['s_cit']}"].value, "project IRR pre-tax": xirr(nz(mr(b["pre"]))),
                 "project IRR post-tax": xirr(nz(mr(b["post"]))), "debt at COD": dash[f"{c}{R['s_debt']}"].value,
                 "equity IRR": xirr(nz(mr(b["eq"]))), "min DSCR": dash[f"{c}{R['s_dscr']}"].value,
                 "CIT levered total": dash[f"H{b['citl']}"].value,
                 "balance at tenor end": next(x for n, x in zip(tn, mr(b["bal"])) if n == ten)}
        for k in py:
            tol = 1e-6 if ("IRR" in k or "DSCR" in k) else 0.01
            ok &= abs(py[k] - model[k]) <= tol
        diff = max(abs(py[k] - model[k]) for k in py)
        res[nm] = model
        print(f"   {nm:<11} " + " | ".join(f"{k} {model[k]:,.4f}" for k in model) + f"   (max python diff {diff:.1e})")
    print("independent recomputation:", "PASS" if ok else "FAIL")
    return ok, res


def lo_irr_check(info):
    """LibreOffice cross-check: plain XIRR of the model's rows in a scratch session (not saved)."""
    R, out = info["Dashboard"]["rows"], {}
    with Office(OUT) as lo:
        s = lo.doc.Sheets.getByName("Dashboard")
        for k, (c, nm) in enumerate(zip(P, NAMES)):
            b = info["Dashboard"]["blocks"][nm]
            for j, key in enumerate(("pre", "post", "eq")):
                cell = s.getCellByPosition(col("AU") + j, k)
                cell.setFormula(f"=XIRR($J${b[key]}:$AS${b[key]};$J${R['ye']}:$AS${R['ye']})")
        lo.calculate()
        for k, nm in enumerate(NAMES):
            out[nm] = [s.getCellByPosition(col("AU") + j, k).getValue() for j in range(3)]
    return out


def dump(wb, info):
    """Labels | unit | E | F | G | H | J (values), then the J formula (or E formula) of the row."""
    fwb = openpyxl.load_workbook(OUT)

    def fmt(x):
        if isinstance(x, datetime):
            return x.strftime("%d-%b-%Y")
        if isinstance(x, float):
            return f"{x:,.6f}" if abs(x) < 10 else f"{x:,.2f}"
        return "" if x is None else str(x)

    def ftxt(v):
        return v.text if hasattr(v, "text") else ("" if v is None else str(v))

    out = []
    for sheet, cols in (("Dashboard", "DEFGHJ"), ("Inputs", "DEFGH")):
        ws, fs = wb[sheet], fwb[sheet]
        out.append(f"#### {sheet}: row | label | " + " | ".join(cols) + " | formula (E, else J)")
        for r in range(1, ws.max_row + 1):
            lab = ws[f"C{r}"].value
            if lab is None:
                continue
            fe = ftxt(fs[f"E{r}"].value)
            fj = ftxt(fs[f"J{r}"].value)
            fh = ftxt(fs[f"H{r}"].value)
            f = next((x for x in (fe, fj, fh) if x.startswith("=")), "")
            out.append(" | ".join([str(r), str(lab)] + [fmt(ws[f"{c}{r}"].value) for c in cols] + [f]))
    DUMP.write_text("\n".join(out) + "\n")


def verify():
    info = json.loads(MAPS.read_text())
    a_wb = openpyxl.load_workbook(RC, data_only=True)
    b_wb = openpyxl.load_workbook(BASE_RC, data_only=True)
    R = info["Dashboard"]["rows"]
    irr_rows = {R["irr"] - 6, R["irr"], R["s_irr"], R["s_eirr"]}          # rows 133 / 139 (user) and the two new
    errs = lint_values(RC)
    known = [e for e in errs if e.rule == "f-error" and re.fullmatch(r"Dashboard![EFG](\d+)", e.where)
             and int(e.where[11:]) in irr_rows]
    other = [e for e in errs if e not in known]
    print(f"recalc: {len(other)} error(s) / Master_Check issue(s) outside the FILTER IRR cells", [str(e) for e in other])
    print(f"   FILTER array IRR cells LibreOffice 24.2 cannot evaluate (#NAME?): {[e.where for e in known]}")
    print("   Master_Check (Dashboard H3):", a_wb["Dashboard"]["H3"].value, "| H", R["master"], "=",
          a_wb["Dashboard"][f"H{R['master']}"].value)
    base_err = [e.where for e in lint_values(BASE_RC) if e.rule == "f-error"]
    print("   base file, same LibreOffice:", base_err)
    diffs = compare(b_wb, a_wb)
    print(f"before / after: {len(diffs)} changed cell(s) in the user's cells")
    for sh, r0, r1, cl, x, y in diffs:
        print(f"   {sh}!{cl}{r0} → {cl}{r1}: {x!r} → {y!r}")
    struct = cells_check(openpyxl.load_workbook(SRC), openpyxl.load_workbook(OUT))
    print(f"   non-empty cells per row (formulas / text) vs the user's file: {len(struct)} row(s) differ", struct[:5])
    want = {(sh, c, remap(sh)(r)) for sh, c, r in array_cells(SRC)} | {
        ("Dashboard", c, R[k]) for c in P for k in ("s_irr", "s_eirr")}
    got = set(array_cells(OUT))
    print(f"array formulas saved as arrays: {len(got)} (user's {len(want) - 6} + 6 new)",
          "OK" if got == want else f"MISMATCH {sorted(got ^ want)}")
    wb = openpyxl.load_workbook(OUT)
    print("hidden / grouped rows or columns:", sum(1 for s in wb.worksheets for _, x in
                                                   list(s.row_dimensions.items()) + list(s.column_dimensions.items())
                                                   if x.hidden or x.outlineLevel),
          "| comments:", [(s.title, c.coordinate) for s in wb.worksheets for row in s.iter_rows() for c in row
                          if c.comment])
    ok, res = recompute(a_wb, info)
    lo = lo_irr_check(info)
    for nm in NAMES:
        x = lo[nm]
        d = max(abs(x[0] - res[nm]["project IRR pre-tax"]), abs(x[1] - res[nm]["project IRR post-tax"]),
                abs(x[2] - res[nm]["equity IRR"]))
        ok &= d < 1e-6
        print(f"   LibreOffice XIRR on the model rows, {nm}: pre {x[0]:.6%} post {x[1]:.6%} equity {x[2]:.6%}"
              f" (diff {d:.1e})")
    # Lint, new rows only.
    new = {"Inputs": set(info["Inputs"]["rows"].values()),
           "Dashboard": {R[k] for k in ("s_cit", "s_irr", "s_debt", "s_eirr", "s_dscr", "check", "master")}
           | {v for b in info["Dashboard"]["blocks"].values() for k, v in b.items()
              if k in [t[0] for t in TIMELINE] or k.startswith("s_")}}
    found = []
    for x in lint(OUT):
        sh, _, where = x.where.partition("!")
        hit = re.search(r"(\d+)$", where)
        if sh in new and hit and int(hit.group(1)) in new[sh]:
            found.append(str(x))
    print("lint (new rows):", "clean" if not found else f"{len(found)} finding(s)")
    for x in found[:20]:
        print("  ", x)
    dump(a_wb, info)
    print("dump:", DUMP)
    return ok and not other and not diffs and got == want


if __name__ == "__main__":
    build()
    sys.exit(0 if verify() else 1)
