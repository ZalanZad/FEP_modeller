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
§28 (CIT on the interest of FEP's shareholder loans to the JV, RtB cases, taxed yearly as capitalised at 31-Dec):
             factor columns M:O on the claim lines of both dated memo lists; per project block, after the COD-exit
             row, a loan sub-block (per case: balances cash / land+devex, drawn, interest capitalised); dated CIT rows
             (31-Dec-2026/27/28, rest at the sale) and post-tax MOIC / IRR after each memo list; loan CIT row in both
             profit stacks; post-tax MOIC / IRR in the summary, "FEP profit, post-tax"; check "rest ≤ 0".
The COD exit is not touched; the user's row groups are kept as they are. Then verify(): recalc, errors, Master_Check, before / after against the
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
# §28: FEP shareholder loans to the JV (RtB cases), CIT on the interest as capitalised at 31-Dec.
LOAN_HEAD = "FEP shareholder loans to the JV (RtB cases): balance at 31-Dec and interest capitalised, taxed yearly"
LOAN_ROWS = [("cash", "loan balance at 31-Dec, cash lines (12 %)"), ("dev", "loan balance at 31-Dec, land and devex (20 %)"),
             ("drawn", "loans drawn in the year"), ("int", "interest capitalised in the year (taxable at FEP)")]
CIT_YEARS = 3                                               # 31-Dec-2026, -2027, -2028: timeline columns J, K, L
LOAN_CIT = "CIT on the interest of FEP's shareholder loans (JV only; 15 %, paid yearly as capitalised)"
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

    q = anchors28(d)
    for b, nm in zip(blocks, NAMES):
        b["codx"] = f("COD exit: buyer's cash flow", f(f"{nm}: annual cash flow"))
    assert [b["codx"] for b in blocks] == [361, 393, 425], blocks

    # Inserts, bottom-up (positions in the user's numbering). §27: hold rows; §28: loan rows, CIT rows, summary.
    ops = [(a["master"], 2)] + [(b["codx"] + 1, 1 + 2 * len(LOAN_ROWS)) for b in reversed(blocks)] + \
          [(b["pre"] + 1, len(TIMELINE)) for b in reversed(blocks)] + \
          [(q[k]["rec"] + 1, CIT_YEARS + 3) for k in "BA"] + [(q[k]["post"], 1) for k in "BA"] + \
          [(b["net"] + 1, len(PER_PROJECT)) for b in reversed(blocks)] + [(a["npv"] + 1, len(SUMMARY))] + \
          [(q[k]["s_irr"] + 1, 2) for k in "BA"]
    ops.sort(key=lambda x: -x[0])
    assert len({p for p, _ in ops}) == len(ops), ops
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
    d.formula(f"H{R['master']}", f"=SUM(H{R['chk1'] + 1}:H{r + 1})")
    R["check28"] = r + 1
    Q = edit28(d, m, R, B, blocks, q)
    return m, R, B, Q


def user_groups(path):
    """The user's grouped Dashboard rows (outline level 1, not collapsed), as runs (first, last)."""
    rows = sorted(r for r, x in openpyxl.load_workbook(path)["Dashboard"].row_dimensions.items() if x.outlineLevel)
    runs = []
    for r in rows:
        if runs and r == runs[-1][1] + 1:
            runs[-1][1] = r
        else:
            runs.append([r, r])
    return runs


def keep_groups(doc, runs, m):
    """Rows inserted right after a group join it in LibreOffice; reset the outline to the user's groups, shifted."""
    import uno
    from com.sun.star.table import CellRangeAddress
    orient = uno.Enum("com.sun.star.table.TableOrientation", "ROWS")
    s = doc.Sheets.getByName("Dashboard")
    s.clearOutline()
    for r1, r2 in runs:
        assert m(r2) - m(r1) == r2 - r1, (r1, r2)                  # no insert inside a user's group
        s.group(CellRangeAddress(Sheet=0, StartColumn=0, StartRow=m(r1) - 1, EndColumn=0, EndRow=m(r2) - 1), orient)


def anchors28(d):
    """§28 anchors in the user's numbering, per RtB case: claims, memo list, profit stack, summary."""
    f = lambda t, s=1: d.find(t, start=s, exact=False)  # noqa: E731
    q = {"B": {"end": f("Interest accrual end (earlier of freeze and sale)"),
               "tot": f("Total FEP claims on its shareholder loans (JV only)"), "memo": f("Memo: FEP dated cash flows, Case B"),
               "s_head": f("RTB exit, Case B"), "sale": "Sale_Date"},
         "A": {"end": f("Case A: interest accrual end"), "tot": f("Total FEP claims on its shareholder loans, Case A"),
               "memo": f("Memo: FEP dated cash flows, Case A"), "s_head": f("RTB exit, Case A"), "sale": "Sale_Date_A"}}
    for k, x in q.items():
        x["post"] = d.find("FEP profit, post-tax", x["memo"] - 12)
        x["land"] = d.find("Land", x["memo"])
        x["rec"] = x["land"] + 2
        x["first"] = x["memo"] + 1                               # signing fee and DD (no claim: no factor)
        x["s_irr"] = d.find("IRR, pre-tax", x["s_head"])
        x["s_profit"] = d.find("FEP profit", x["s_head"])
        assert d.cell(f"C{x['first']}").getString() == "Signing fee and DD", k
        assert d.cell(f"C{x['rec']}").getString().startswith("FEP receipts at"), k
        assert d.cell(f"C{x['post'] - 1}").getString().startswith("CIT on the exit gain"), k
    expect = {"B": (175, 198, 315, 312, 323, 325, 121, 119), "A": (210, 233, 270, 266, 276, 278, 115, 113)}
    got = {k: tuple(x[i] for i in ("end", "tot", "memo", "post", "land", "rec", "s_irr", "s_profit")) for k, x in q.items()}
    assert got == expect, got
    return q


def edit28(d, m, R, B, blocks, q):
    """§28 rows: factor columns M:O, loan sub-blocks per project, dated CIT rows, profit stack, summary, check."""
    Q = {k: {n: m(v) for n, v in x.items() if isinstance(v, int)} | {"sale": x["sale"]} for k, x in q.items()}
    fcols, dcols = ["M", "N", "O"], ["J", "K", "L"]
    # 1. Factor to the first 31-Dec on the claim lines (M:O; dates in J:L), one formula per row.
    for k, x in Q.items():
        for fc, dc, nm in zip(fcols, dcols, NAMES):
            d.s.copyRange(d.cell(f"{fc}{x['memo']}").getCellAddress(), d.cell(f"J{x['memo']}").getRangeAddress())
            d.text(f"{fc}{x['memo']}", f"Factor to the first 31-Dec: {nm}")
        for r in range(x["first"] + 1, x["rec"]):
            rate = "Rate_Devex" if r >= x["land"] else "Rate_Cash"
            for fc, dc in zip(fcols, dcols):
                d.formula(f"{fc}{r}", f"=1+{rate}*(DATE(YEAR({dc}{r})+1,1,1)-{dc}{r})/Days_In_Year")
                d.cell(f"{fc}{r}").NumberFormat = d.cell(f"E{Q['B']['end'] + 1}").NumberFormat   # 0.0000
    # 2. Loan sub-blocks after each project's COD-exit row.
    for c, fc, dc, b, blk in zip(P, fcols, dcols, B, blocks):
        rows = m.block(blk["codx"] + 1)
        d.copy_format(R["hold"], rows[0])
        d.text(f"C{rows[0]}", LOAN_HEAD)
        it = iter(rows[1:])
        for k in "BA":
            x = Q[k]
            rr = {key: next(it) for key, _ in LOAN_ROWS}
            fm = {"cash": f"=({{w}}{rr['cash']}*(1+Rate_Cash)+SUMPRODUCT(-${c}{x['first'] + 1}:${c}{x['land'] - 1}"
                          f"*${fc}{x['first'] + 1}:${fc}{x['land'] - 1}*(YEAR(${dc}{x['first'] + 1}:${dc}{x['land'] - 1})"
                          f"={{x}}{R['year']})))*({{x}}{R['ye']}<${c}{x['end']})",
                  "dev": f"=({{w}}{rr['dev']}*(1+Rate_Devex)+SUMPRODUCT(-${c}{x['land']}:${c}{x['land'] + 1}"
                         f"*${fc}{x['land']}:${fc}{x['land'] + 1}*(YEAR(${dc}{x['land']}:${dc}{x['land'] + 1})"
                         f"={{x}}{R['year']})))*({{x}}{R['ye']}<${c}{x['end']})",
                  "drawn": f"=-SUMPRODUCT(${c}{x['first'] + 1}:${c}{x['land'] + 1}*(YEAR(${dc}{x['first'] + 1}:"
                           f"${dc}{x['land'] + 1})={{x}}{R['year']}))*({{x}}{R['ye']}<${c}{x['end']})",
                  "int": f"=({{x}}{rr['cash']}+{{x}}{rr['dev']}-{{w}}{rr['cash']}-{{w}}{rr['dev']}-{{x}}{rr['drawn']})"
                         f"*({{x}}{R['ye']}<${c}{x['end']})"}
            for key, lab in LOAN_ROWS:
                r = rr[key]
                d.copy_format(R["money"], r)
                d.text(f"C{r}", f"Case {k}: {lab}")
                d.text(f"D{r}", "EUR k")
                set_row(d, r, fm[key])
                if key in ("drawn", "int"):
                    d.formula(f"H{r}", f"=SUM(J{r}:{LAST}{r})")
                b[f"loan{k}_{key}"] = r
    # 3. Dated CIT rows after the receipts row, then post-tax MOIC and IRR on the full list.
    for k, x in Q.items():
        rows = m.block(q[k]["rec"] + 1)
        years, rest, moic, irr = rows[:CIT_YEARS], rows[CIT_YEARS], rows[CIT_YEARS + 1], rows[CIT_YEARS + 2]
        for i, r in enumerate(years):
            yc = TL[i]
            d.copy_format(x["land"], r, last="L")
            d.text(f"C{r}", f"CIT on the interest capitalised at 31-Dec-{2026 + i}")
            d.text(f"D{r}", "EUR k")
            for c, dc, b in zip(P, dcols, B):
                d.formula(f"{c}{r}", f"=-{yc}${b[f'loan{k}_int']}*CIT_Rate*{c}$21")
                d.formula(f"{dc}{r}", f"={yc}${R['ye']}")
            d.formula(f"H{r}", f"=SUM(E{r}:G{r})")
        d.copy_format(x["land"], rest, last="L")
        d.text(f"C{rest}", "CIT on the rest at the sale (stub to the accrual end, devex floor top-up)")
        d.text(f"D{rest}", "EUR k")
        for c, dc in zip(P, dcols):
            d.formula(f"{c}{rest}", f"=-{c}{x['tot']}*CIT_Rate*{c}$21-SUM({c}{years[0]}:{c}{years[-1]})")
            d.formula(f"{dc}{rest}", f"={x['sale']}")
        d.formula(f"H{rest}", f"=SUM(E{rest}:G{rest})")
        pre_moic = d.find("MOIC, pre-tax (dated", x["post"], exact=False)
        pre_irr = d.find("IRR, pre-tax (dated", x["post"], exact=False)
        d.copy_format(pre_moic, moic, last="H")
        d.text(f"C{moic}", "MOIC, post-tax (dated cash flows incl. the CIT rows)")
        d.text(f"D{moic}", "x")
        for c in P + ["H"]:
            rng = f"{c}{x['first']}:{c}{rest}"
            d.formula(f"{c}{moic}", f'=SUMIF({rng},">0")/-SUMIF({rng},"<0")')
        d.copy_format(pre_irr, irr, last="H")
        d.text(f"C{irr}", "IRR, post-tax (dated cash flows incl. the CIT rows)")
        d.text(f"D{irr}", "%")
        for c, dc in zip(P, dcols):
            d.formula(f"{c}{irr}", f"=XIRR({c}{x['first']}:{c}{rest},{dc}{x['first']}:{dc}{rest})")
        x.update(cit_years=years, rest=rest, moic=moic, irr=irr)
    # 4. Profit stacks: loan CIT before "FEP profit, post-tax" (accrued claims, not the waterfall payment).
    for k, x in Q.items():
        r, post = m.block(q[k]["post"])[0], x["post"]
        d.copy_format(r - 1, r, last="H")
        d.text(f"C{r}", LOAN_CIT)
        d.text(f"D{r}", "EUR k")
        for c in P:
            d.formula(f"{c}{r}", f"=-{c}{x['tot']}*CIT_Rate*{c}$21")
            old = d.cell(f"{c}{post}").getFormula()
            assert old == f"={c}{r - 2}+{c}{r - 1}", old
            d.formula(f"{c}{post}", f"{old}+{c}{r}")
        d.formula(f"H{r}", f"=SUM(E{r}:G{r})")
        x["loan_cit"] = r
    # 5. Summary: post-tax MOIC and IRR under each case's pre-tax IRR; "FEP profit" relabelled.
    for k, x in Q.items():
        r1, r2 = m.block(q[k]["s_irr"] + 1)
        for r, src, lab, unit in ((r1, x["s_irr"] - 1, "MOIC, post-tax", "x"), (r2, x["s_irr"], "IRR, post-tax", "%")):
            d.copy_format(src, r, last="H")
            d.text(f"C{r}", lab)
            d.text(f"D{r}", unit)
            for c in P:
                d.formula(f"{c}{r}", f"={c}{x['moic'] if unit == 'x' else x['irr']}")
        d.formula(f"H{r1}", f"=H{x['moic']}")
        d.text(f"C{x['s_profit']}", "FEP profit, post-tax")
        x.update(s_moic=r1, s_irr_post=r2)
    # 6. Check: the rest taxed at the sale is <= 0 in both cases.
    r = R["check28"]
    d.copy_format(R["check"], r, last="H")
    d.text(f"C{r}", "RtB loan CIT: the rest taxed at the sale is ≤ 0 (yearly schedule does not over-accrue), "
                    "Case B + Case A")
    d.text(f"D{r}", "flag")
    for c in P:
        d.formula(f"{c}{r}", f"=N({c}{Q['B']['rest']}>Tolerance)+N({c}{Q['A']['rest']}>Tolerance)")
    d.formula(f"H{r}", f"=SUM(E{r}:G{r})")
    return Q


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
        md, R, B, Q = edit_dashboard(lo.doc)
        keep_arrays(lo.doc, array_cells(SRC), {"Inputs": mi, "Dashboard": md})
        keep_groups(lo.doc, user_groups(SRC), md)
        lo.calculate()
        lo.save(OUT)
    restore_filter_case(OUT)
    with Office(OUT) as lo:
        lo.save(RC)
    info = {"Inputs": {"inserts": mi.ins, "rows": ri},
            "Dashboard": {"inserts": md.ins, "rows": R, "blocks": dict(zip(NAMES, B)), "rtb": Q}}
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


# User cells that §27 / §28 change on purpose (user's numbering): relabels, the post-tax profit (+ loan CIT) and its
# summary link, and the new factor cells M:O on the two memo lists (heading and claim lines).
RELABELS = {("Dashboard", "C134"), ("Dashboard", "C113"), ("Dashboard", "C119"), ("Inputs", "C39"), ("Inputs", "H39")}
PROFIT_ROWS = {266, 312, 113, 119}
FACTOR_ROWS = {270, *range(272, 278), 315, *range(317, 325)}


def expected(sh, cl, r0):
    if (sh, f"{cl}{r0}") in RELABELS:
        return "relabel"
    if sh == "Dashboard" and r0 in PROFIT_ROWS and cl in "EFGH":
        return "post-tax profit incl. loan CIT"
    if sh == "Dashboard" and r0 in FACTOR_ROWS and cl in "MNO":
        return "new factor cell"
    return None


def recompute28(wb, info):
    """Bora (and the others): yearly capitalised loan interest and the post-tax RtB returns, from the memo lists."""
    from datetime import date
    dash, Q = wb["Dashboard"], info["Dashboard"]["rtb"]
    name = lambda n: (lambda s, c: wb[s][c.replace("$", "")].value)(*next(wb.defined_names[n].destinations))  # noqa
    rc, rd, cit, freeze = name("Rate_Cash"), name("Rate_Devex"), name("CIT_Rate"), name("Freeze_Date").date()
    ok, out = True, {}
    for c, dc, nm in zip(P, "JKL", NAMES):
        jv = dash[f"{c}21"].value
        for k, x in Q.items():
            sale = name(x["sale"]).date()
            end = min(freeze, sale)
            lines = [(dash[f"{c}{r}"].value, dash[f"{dc}{r}"].value, rd if r >= x["land"] else rc)
                     for r in range(x["first"] + 1, x["rec"])]          # datetimes: amount-weighted dates keep the
            bal = {}                                                     # fraction of a day, as the claim rows do
            for y in range(2026, 2026 + CIT_YEARS + 1):
                bal[y] = sum(-p * (1 + rt * (datetime(d.year + 1, 1, 1) - d).total_seconds() / 86400 / 365)
                             * (1 + rt) ** (y - d.year) for p, d, rt in lines if d.year <= y) * (y < end.year)
            drawn = {y: -sum(p for p, d, _ in lines if d.year == y) * (y < end.year) for y in bal}
            ints = {y: (bal[y] - bal.get(y - 1, 0) - drawn[y]) * (y < end.year) for y in bal}
            yearly = [-ints[2026 + i] * cit * jv for i in range(CIT_YEARS)]
            total = -dash[f"{c}{x['tot']}"].value * cit * jv
            flows = [(dash[f"{dc}{r}"].value.date(), dash[f"{c}{r}"].value) for r in range(x["first"], x["rec"] + 1)]
            flows += [(date(2026 + i, 12, 31), v) for i, v in enumerate(yearly)] + [(sale, total - sum(yearly))]
            pos, neg = sum(v for _, v in flows if v > 0), -sum(v for _, v in flows if v < 0)
            py = {"yearly CIT": yearly, "rest at sale": total - sum(yearly), "total loan CIT": total,
                  "MOIC post-tax": pos / neg, "IRR post-tax": xirr(sorted(flows))}
            model = {"yearly CIT": [dash[f"{c}{r}"].value for r in x["cit_years"]],
                     "rest at sale": dash[f"{c}{x['rest']}"].value, "total loan CIT": dash[f"{c}{x['loan_cit']}"].value,
                     "MOIC post-tax": dash[f"{c}{x['moic']}"].value, "IRR post-tax": dash[f"{c}{x['irr']}"].value}
            diff = max([abs(a - b) for a, b in zip(py["yearly CIT"], model["yearly CIT"])] +
                       [abs(py[t] - model[t]) for t in py if t != "yearly CIT"])
            ok &= diff < 1e-6
            out[(nm, k)] = model
            print(f"   {nm:<11} Case {k}: yearly CIT " + ", ".join(f"{v:,.2f}" for v in model["yearly CIT"]) +
                  f" | rest {model['rest at sale']:,.2f} | total {model['total loan CIT']:,.2f} | MOIC post "
                  f"{model['MOIC post-tax']:.4f} | IRR post {model['IRR post-tax']:.4%}  (python diff {diff:.1e})")
    print("independent recomputation (RtB loan CIT):", "PASS" if ok else "FAIL")
    return ok


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
    bad = [x for x in diffs if not expected(x[0], x[3], x[1])]
    print(f"before / after: {len(diffs)} changed cell(s) in the user's cells, {len(bad)} unexpected")
    for sh, r0, r1, cl, x, y in diffs:
        print(f"   {sh}!{cl}{r0} → {cl}{r1}: {x!r} → {y!r}  [{expected(sh, cl, r0) or 'UNEXPECTED'}]")
    struct = [x for x in cells_check(openpyxl.load_workbook(SRC), openpyxl.load_workbook(OUT))
              if not (x[0] == "Dashboard" and x[1] in FACTOR_ROWS)]
    print(f"   non-empty cells per row (formulas / text) vs the user's file: {len(struct)} row(s) differ", struct[:5])
    want = {(sh, c, remap(sh)(r)) for sh, c, r in array_cells(SRC)} | {
        ("Dashboard", c, R[k]) for c in P for k in ("s_irr", "s_eirr")}
    got = set(array_cells(OUT))
    print(f"array formulas saved as arrays: {len(got)} (user's {len(want) - 6} + 6 new)",
          "OK" if got == want else f"MISMATCH {sorted(got ^ want)}")
    wb = openpyxl.load_workbook(OUT)
    m = remap("Dashboard")
    grp = lambda w, f=lambda r: r: {f(r) for r, x in w["Dashboard"].row_dimensions.items() if x.outlineLevel}  # noqa
    same = grp(wb) == grp(openpyxl.load_workbook(SRC), m)
    print("user's grouped rows kept, none added:", "OK" if same else "MISMATCH")
    print("hidden / grouped rows or columns:", sum(1 for s in wb.worksheets for _, x in
                                                   list(s.row_dimensions.items()) + list(s.column_dimensions.items())
                                                   if x.hidden or x.outlineLevel),
          "| comments:", [(s.title, c.coordinate) for s in wb.worksheets for row in s.iter_rows() for c in row
                          if c.comment])
    ok, res = recompute(a_wb, info)
    ok &= recompute28(a_wb, info)
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
              if k in [t[0] for t in TIMELINE] or k.startswith(("s_", "loan"))}
           | {R["check28"]} | {v for x in info["Dashboard"]["rtb"].values() for k, v in x.items()
                               if k in ("rest", "moic", "irr", "loan_cit", "s_moic", "s_irr_post")}
           | {r for x in info["Dashboard"]["rtb"].values() for r in x["cit_years"]}}
    loan_heads = {b["loanB_cash"] - 1 for b in info["Dashboard"]["blocks"].values()}
    new["Dashboard"] |= loan_heads | {r for x in info["Dashboard"]["rtb"].values()
                                      for r in range(x["first"] + 1, x["rec"])}
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
    return ok and not other and not bad and not struct and got == want and same


if __name__ == "__main__":
    build()
    sys.exit(0 if verify() else 1)
