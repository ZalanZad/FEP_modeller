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
§29 (audit fixes): development-devex loan principal repaid before the 70/30 split in both RtB waterfalls (and in FEP's
             receipts); plain XIRR on the whole row instead of the FILTER array IRRs, portfolio IRR in the Total column
             of the RtB cases; checks "dated cash flows tie to the post-tax profit" (both cases) and "no errors in the
             summary block"; dead "RtB price: 0" rows deleted; JV flag on the loan sub-blocks; debt service, interest
             and principal negative; labels without parameters; exposure rows 73/74 (land returned at cost); small fixes.
The COD exit is not touched; the user's row groups are kept as they are. Then verify(): recalc, errors, Master_Check, before / after against the
recalculated base (out/rows_v3.json), independent Python recomputation (hold case, RtB waterfalls incl. the devex
principal, loan CIT), Low / High case flip and a broken-flow test in a scratch session, lint on the new rows,
out/dump_v3.txt.

    python3 projects/hu_wind_2/build_v3.py
"""
import json
import re
import sys
import zipfile
from datetime import datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parents[1]))
from build_lev import add_name, col, letter, xirr  # noqa: E402
from build_lev2 import Rows  # noqa: E402
from build_lev2 import Sheet as _Sheet  # noqa: E402
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
USER_LAST, N_EXTRA = "AS", 4                                # §31.2: the user's timeline 2026–2061, + 4 years
LAST = letter(col(USER_LAST) + N_EXTRA)                     # AW (2065)
TL = [letter(i) for i in range(col("J"), col(LAST) + 1)]
NAMES = ["Bora", "Iváncsa", "Dunaújváros"]
CIT_LABEL = "Corporate income tax (operating profit; exit gains)"
HOLD_LABEL = "Hold to maturity (as the RtB buyer; EBITDA, CIT, optional debt)"

# name, label, unit, value, template ("pct" | "years"), source note
NEW_INPUTS = [
    ("Tax_Dep_Life", "Tax depreciation life (straight line on the depreciable base, from COD)", "years", 20,
     "years", "HU CIT Act Annex 2, 5 % p.a. power plants – to confirm"),
    ("Gearing", "Gearing (debt ÷ all-in cost to COD)", "%", 0.60, "pct",
     "user, 27-Oct-2026"),
    ("Debt_Rate", "Debt interest rate", "%", 0.05, "pct", "Bora model Inputs G220"),
    ("Debt_Tenor", "Debt tenor (annuity from the first full operating year)", "years", 15, "years",
     "Bora model Inputs G222"),
    ("Debt_Fee", "Debt upfront fee (on the debt size, at the first draw)", "%", 0.0175, "pct",
     "Bora model Inputs E85 / HTM E85"),
    ("IDC_Draw_Share", "Share of the year's draw bearing interest in that year (mid-year draws)", "%", 0.5, "pct",
     "Bora model HTM (IDC on half the year's draw)"),
]
SIGNING = 135                                               # Inputs row 17, all three projects (user, 29-Oct-2026)

# Hold-case rows after "Unlevered pre-tax cash flow" (§30.4–30.6, as the Bora model HTM sheet): key, label, unit,
# template, formula, H column. {x} this column, {w} the column to the left (I for J: empty, so openings are 0), {c} the
# project's value column (E / F / G, as the user's "Value" cells); kind "value" = one cell in {c}, no timeline.
TIMELINE = [
    ("dsh", "Tax depreciation share of the year (Tax_Dep_Life years from COD)", "%", "share",
     "=MAX(0,MIN({x}{mo},Tax_Dep_Life*Months_In_Year)-MAX({x}{mo}-Months_In_Year,0))/Months_In_Year", "sum"),
    ("base", "Depreciable base, unlevered (construction capex, connection fees, municipal)", "EUR k", "value",
     "={c}{capx}+{c}{conn}+{c}{muni}", None),
    ("dep", "Tax depreciation, unlevered", "EUR k", "money", "=-${c}{base}/Tax_Dep_Life*{x}{dsh}", "sum"),
    ("cit", "CIT, unlevered", "EUR k", "money", "=-MAX(0,{x}{eb}+{x}{dep})*CIT_Rate", "sum"),
    ("post", "Unlevered post-tax cash flow", "EUR k", "bold", "={x}{pre}+{x}{cit}", "sum"),
    ("size", "Debt size (gearing × all-in cost to COD)", "EUR k", "value", "={c}{allin}*Gearing", None),
    ("draw", "Debt drawn (pro rata with construction)", "EUR k", "money", "=${c}{size}*{x}{cs}", "sum"),
    # IDC on the draws and IDC to date (rows draw:idc, adjacent) = the opening balance before COD; not on the balance
    # row, which carries the principal of the annuity sized on the IDC total (a circular reference otherwise).
    ("idc", "Interest during construction, capitalised (on draws and IDC to date, mid-year draws)", "EUR k",
     "money", "=(SUM($I{draw}:{w}{idc})+{x}{draw}*IDC_Draw_Share)*Debt_Rate*{x}{cs}*${c}{cm}/Months_In_Year", "sum"),
    ("fee", "Debt upfront fee at the first draw (equity-paid; deductible in the year paid, no loss carry-forward)",
     "EUR k", "money", "=-${c}{size}*Debt_Fee*({x}{cs}>0)*({w}{cs}=0)", "sum"),
    ("int", "Interest paid", "EUR k", "money", "=-({w}{bal}+{x}{draw}+{x}{idc})*Debt_Rate*{x}{op}", "sum"),
    ("ds", "Debt service (interest in the COD year; annuity from the first full operating year)", "EUR k", "money",
     "=PMT(Debt_Rate,Debt_Tenor,$H{draw}+$H{idc})*({x}{t}>=1)*({x}{t}<=Debt_Tenor)+{x}{int}*({x}{t}=0)", "sum"),
    ("prin", "Principal repaid", "EUR k", "money", "={x}{ds}-{x}{int}", "sum"),
    ("bal", "Debt balance, closing", "EUR k", "money", "={w}{bal}+{x}{draw}+{x}{idc}+{x}{prin}", None),
    ("depi", "Tax depreciation of the capitalised interest", "EUR k", "money", "=-$H{idc}/Tax_Dep_Life*{x}{dsh}",
     "sum"),
    ("citl", "CIT, levered (interest deductible)", "EUR k", "money",
     "=-MAX(0,{x}{eb}+{x}{dep}+{x}{depi}+{x}{int}+{x}{fee})*CIT_Rate", "sum"),
    ("eq", "Equity cash flow, post-tax (geared)", "EUR k", "bold", "={x}{pre}+{x}{draw}+{x}{fee}+{x}{ds}+{x}{citl}",
     "sum"),
    ("dscr", "DSCR (EBITDA + CIT) ÷ debt service", "x", "ratio", '=IF({x}{ds}<0,-({x}{eb}+{x}{citl})/{x}{ds},"")',
     "min"),
]
SUMMARY = [("s_cit", "CIT over the life, unlevered", "EUR k", "money"),
           ("s_irr", "Project IRR, post-tax", "%", "irr"),
           ("s_debt", "Debt at COD (draws + capitalised interest)", "EUR k", "money"),
           ("s_eirr", "Equity IRR, post-tax (geared)", "%", "irr"),
           ("s_dscr", "Minimum DSCR", "x", "ratio")]
# §28: FEP shareholder loans to the JV (RtB cases), CIT on the interest as capitalised at 31-Dec.
LOAN_HEAD = "FEP shareholder loans to the JV (RtB cases): balance at 31-Dec and interest capitalised, taxed yearly"
LOAN_ROWS = [("cash", "loan balance at 31-Dec, cash lines"), ("dev", "loan balance at 31-Dec, land and devex"),
             ("drawn", "loans drawn in the year"), ("int", "interest capitalised in the year (taxable at FEP)")]
CIT_YEARS = 3                                               # 31-Dec-2026, -2027, -2028: timeline columns J, K, L
LOAN_CIT = "CIT on the interest of FEP's shareholder loans (JV only; paid yearly as capitalised)"
PER_PROJECT = [("post", "Net cash flow, post-tax"), ("eq", "Equity cash flow, post-tax (geared)")]
# §29 (audit fixes).
UNUSED_NAMES = ["Opex_BoP", "Opex_Lease", "Opex_Ins", "Opex_Mgmt", "Opex_PropTax"]
ERR_CHECK = "No errors in the summary block"
JV_LABEL = "JV with a developer (1 = yes; shareholder-loan claims apply)"
SPEND = "Spending (RtB purchase, capex, fees, securities), net"
LAND_NOTE = "Land is in the devex (Inputs rows {}) and bought by the RtB buyer at cost"


class Sheet(_Sheet):
    def copy_format(self, src_row, dst_row, last=None):
        super().copy_format(src_row, dst_row, last or LAST)     # the extended timeline by default


def extend_timeline(doc):
    """§31.2: 4 columns inserted before AS (every $J:$AS range and $AS reference moves to AW), filled with each row's
    formula and format from AR. The timeline is positional (year = First_Year + column), so J:AS keep 2026–2061 and
    AT:AW are 2062–2065: the before / after comparison is column for column."""
    s = doc.Sheets.getByName("Dashboard")
    s.Columns.insertByIndex(col(USER_LAST), N_EXTRA)
    n = s.createCursor()
    n.gotoEndOfUsedArea(False)
    last_row = n.getRangeAddress().EndRow + 1
    src = s.getCellRangeByName(f"AR1:AR{last_row}").getRangeAddress()
    width = s.Columns.getByIndex(col("AR")).Width
    for i in range(N_EXTRA):
        j = col(USER_LAST) + i
        s.copyRange(s.getCellByPosition(j, 0).getCellAddress(), src)
        s.Columns.getByIndex(j).Width = width


class Rows2(Rows):
    """Rows with deletes too: old row → new row, a deleted old row → None (all positions in the user's numbering)."""
    def __init__(self):
        super().__init__()
        self.dels = []

    def delete(self, r):
        self.dels.append(r)

    def __call__(self, r):
        if r in self.dels:
            return None
        return super().__call__(r) - sum(1 for x in self.dels if x < r)


def xirr_row(row, ye):
    return f"=XIRR($J${row}:${LAST}${row},$J${ye}:${LAST}${ye})"


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
    inp.text(f"H{r_ex}", "user, 27-Oct-2026 (share deal)")                       # §29.9: stale notes
    r_sign, r_disc = inp.find("Signing fee and DD"), tmpl["pct"]                  # §30.2, §30.3
    assert [inp.cell(f"{c}{r_sign}").getValue() for c in P] == [SIGNING] * 3       # set before the base recalc
    inp.text(f"H{r_sign}", "user, 29-Oct-2026")
    assert inp.cell(f"C{r_disc}").getString() == "Discount rate (buyer's pre-tax return at COD)"
    inp.text(f"C{r_disc}", "Discount rate (buyer's unlevered post-tax return at COD)")
    r_land = inp.find("Land purchase and easements")
    inp.insert(r_ex + 1, 1 + len(NEW_INPUTS))
    m.add(r_ex + 1, 1 + len(NEW_INPUTS))
    land = [r for r in range(m(r_land), 300) if inp.cell(f"G{r}").getString() == "Land"]   # devex items, class Land
    assert land == [m(r) for r in (80, 81, 111)], land          # the user's rows 80, 81, 111
    inp.text(f"H{m(r_land)}", LAND_NOTE.format(", ".join(map(str, land))))
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
    return m, out, [[r_sign, "H"], [r_disc, "C"]]


def set_signing(doc):
    """§30.2: the user's one input change (signing fee and DD 135 k EUR per project), made before the base recalc so
    that the before / after comparison shows only the model changes."""
    inp = Sheet(doc, "Inputs")
    r = inp.find("Signing fee and DD")
    for c in P:
        inp.number(f"{c}{r}", SIGNING)


def edit_dashboard(doc):
    d = Sheet(doc, "Dashboard")
    m = Rows2()
    f = lambda t, s=1: d.find(t, start=s, exact=False)  # noqa: E731
    a = {"allin": f("All-in cost to COD for a buyer"), "hold": f("Hold to maturity"), "year": d.find("Year", 300),
         "ye": d.find("Year end"), "irr": f("Project IRR, pre-tax"), "moic": f("MOIC, pre-tax", 120),
         "money": f("Construction capex (straight-line"), "chk1": f("Checks (0 = OK)"), "master": d.find("Master check"),
         "s_top": d.find("Returns RTB COD HTM"), "capx": d.find("Construction capex"),
         "conn": f("Connection fees (TSO security"), "muni": d.find("Municipal one-off"), "sale_cod": f("Sale price at COD")}
    a["npv"] = f("NPV at the discount rate", a["hold"])
    a["cod_irr"] = f("IRR, pre-tax", f("COD exit (as the RtB buyer"))
    a["chk_last"] = a["master"] - 1
    s_head = f("Annual cash flow per project")
    blocks = []
    for nm in NAMES:
        h = f(f"{nm}: annual cash flow")
        b = {"t": d.find("Years from the COD year", h), "op": f("Operating share of the year", h),
             "cod": d.find("COD year flag", h), "eb": d.find("EBITDA", h), "pre": d.find("Unlevered pre-tax cash flow", h)}
        b["net"] = d.find("Net cash flow", d.find(nm, s_head))
        b["dead"] = f("RtB price: 0", h)
        b.update(mo=f("Months operated by year end", h), cs=f("Construction share of the year", h),
                 cm=f("Construction months by year end", h), df=f("Discount factor to COD", h),
                 val=f("EBITDA valued at COD", h), inv=f("Investment cash flow", h))
        blocks.append(b)
    expect = {"allin": 99, "hold": 134, "irr": 139, "npv": 140, "year": 328, "ye": 329, "master": 442, "cod_irr": 133,
              "s_top": 108, "capx": 91, "conn": 92, "muni": 96, "sale_cod": 130}
    assert all(a[k] == v for k, v in expect.items()), {k: a[k] for k in expect}
    assert [(b["t"], b["op"], b["cod"], b["eb"], b["pre"], b["net"]) for b in blocks] == [
        (332 + k, 336 + k, 337 + k, 356 + k, 357 + k, 153 + n) for k, n in ((0, 0), (32, 8), (64, 16))], blocks

    q = anchors28(d)
    anchors29(d, q)
    for b, nm in zip(blocks, NAMES):
        b["codx"] = f("COD exit: buyer's cash flow", f(f"{nm}: annual cash flow"))
    assert [b["codx"] for b in blocks] == [361, 393, 425], blocks
    assert [b["dead"] for b in blocks] == [b["codx"] - 1 for b in blocks], blocks
    U = user29(d, a, blocks, q)                                  # §29 edits on the user's cells, user's numbering

    # Inserts (and §29 deletes, n = -1), bottom-up (positions in the user's numbering). §27: hold rows; §28: loan
    # rows, CIT rows, summary; §29: dead rows deleted; §30: the hold rows are the Bora-model set.
    ops = [(a["master"], 3)] + [(b["codx"] + 1, 1 + 2 * len(LOAN_ROWS)) for b in reversed(blocks)] + \
          [(b["pre"] + 1, len(TIMELINE)) for b in reversed(blocks)] + \
          [(q[k]["rec"] + 1, CIT_YEARS + 4) for k in "BA"] + [(q[k]["post"], 1) for k in "BA"] + \
          [(b["net"] + 1, len(PER_PROJECT)) for b in reversed(blocks)] + [(a["npv"] + 1, len(SUMMARY))] + \
          [(q[k]["s_irr"] + 1, 2) for k in "BA"] + [(q["B"]["end"] + 1, N_FACTOR_NEW)] + \
          [(b.pop("dead"), -1) for b in blocks] + [(f("Claim lines paid before the accrual end year"), -1)]
    ops.sort(key=lambda x: -x[0])
    assert len({p for p, _ in ops}) == len(ops), ops
    for p, n in ops:
        if n > 0:
            d.insert(p, n)
            m.add(p, n)
        else:
            d.s.Rows.removeByIndex(p - 1, 1)
            m.delete(p)
    R = {k: m(v) for k, v in a.items()}
    B = [{k: m(v) for k, v in b.items()} for b in blocks]
    factors31(d, m, q)
    tmpl = {"money": R["money"], "bold": B[0]["pre"], "ratio": R["money"], "irr": R["irr"], "moic": R["moic"],
            "npv": R["npv"], "value": R["money"]}
    d.text(f"C{R['hold']}", HOLD_LABEL)
    d.text(f"C{R['sale_cod']}", "Sale price at COD (PV of unlevered post-tax cash flow)")

    # §29.2: the user's FILTER array IRRs (COD exit, hold pre-tax) → plain XIRR on the whole row (no array).
    for r, key in ((R["cod_irr"], "codx"), (R["irr"], "pre")):
        d.cell(f"E{r}:G{r}").clearContents(1 | 2 | 4 | 16)
        for c, b in zip(P, B):
            d.formula(f"{c}{r}", xirr_row(b[key], R["ye"]))

    # Project timeline rows.
    for c, b, blk in zip(P, B, blocks):
        new = m.block(blk["pre"] + 1)
        b.update({key: r for (key, *_), r in zip(TIMELINE, new)})
        refs = b | {"c": c} | {k: R[k] for k in ("capx", "conn", "muni", "allin")}
        for key, lab, unit, kind, fml, tot in TIMELINE:
            r = b[key]
            d.copy_format(b["op"] if kind == "share" else tmpl[kind], r)
            d.text(f"C{r}", lab)
            d.text(f"D{r}", unit)
            fml = fml.replace("{x}", "{{x}}").replace("{w}", "{{w}}").format(**refs)
            if kind == "value":                                    # one value in the project's column
                d.formula(f"{c}{r}", fml)
                d.cell(f"{c}{r}").NumberFormat = d.cell(f"H{R['money']}").NumberFormat
            else:
                set_row(d, r, fml)
            if tot == "sum":
                d.formula(f"H{r}", f"=SUM(J{r}:{LAST}{r})")
            elif tot == "min":
                d.formula(f"H{r}", f"=MIN(J{r}:{LAST}{r})")
            if kind == "ratio":
                for ref in (f"H{r}", f"J{r}:{LAST}{r}"):
                    d.cell(ref).NumberFormat = d.cell(f"E{R['moic']}").NumberFormat
        # §30.3: COD price on the unlevered post-tax cash flow (the CIT row sits above, in the rows just added).
        assert d.cell(f"J{b['val']}").getFormula() == f"=J{b['eb']}*J{b['df']}", d.cell(f"J{b['val']}").getFormula()
        d.text(f"C{b['val']}", "Unlevered post-tax cash flow valued at COD (total = COD price)")
        set_row(d, b["val"], f"=({{x}}{b['eb']}+{{x}}{b['cit']})*{{x}}{b['df']}")

    # Hold summary rows after the NPV row; IRRs as plain XIRR on the whole row.
    for (key, lab, unit, kind), r in zip(SUMMARY, m.block(a["npv"] + 1)):
        d.copy_format({"money": tmpl["npv"], "irr": tmpl["irr"], "ratio": tmpl["moic"]}[kind], r, last="H")
        d.text(f"C{r}", lab)
        d.text(f"D{r}", unit)
        for c, b in zip(P, B):
            if kind == "irr":
                d.formula(f"{c}{r}", xirr_row(b["post"] if key == "s_irr" else b["eq"], R["ye"]))
            elif key == "s_debt":
                d.formula(f"{c}{r}", f"=$H${b['draw']}+$H${b['idc']}")
            else:
                d.formula(f"{c}{r}", "=$H${}".format(b[{"s_cit": "cit", "s_dscr": "dscr"}[key]]))
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
    d.formula(f"H{R['master']}", f"=SUM(H{R['chk1'] + 1}:H{r + 2})")
    R["check28"], R["check_err"] = r + 1, r + 2
    Q = edit28(d, m, R, B, blocks, q)
    edit29(d, m, R, q, Q)
    return m, R, B, Q, U


# §31.1: Case B claim accrual factors in the Case A pattern (a payment may fall in the accrual end year).
FACTOR_LINES = [("bid submission", "Inputs!{c}$21", "Rate_Cash"), ("performance security posting", "Inputs!{c}$22",
                                                                       "Rate_Cash"),
                ("TSO security payment", "Inputs!{c}$23", "Rate_Cash"),
                ("connection agreement (HCSSZ)", "Inputs!{c}$24", "Rate_Cash"),
                ("land payment (amount-weighted)", "$J$56", "Rate_Devex"),
                ("development devex payment (amount-weighted)", "$J$55", "Rate_Devex")]
N_FACTOR_NEW = 1 + 3 * len(FACTOR_LINES) - 14                # 19 rows replace the user's 14


def factors31(d, m, q):
    """Rewrite the Case B factor block (user rows end+1 … end+14, plus the inserted rows) and repoint the claim rows."""
    f = lambda t, s=1: d.find(t, start=s, exact=False)  # noqa: E731
    e, rows = m(q["B"]["end"]), list(range(m(q["B"]["end"]) + 1, m(q["B"]["end"] + 14) + 1))
    assert len(rows) == 1 + 3 * len(FACTOR_LINES), rows
    ta = f("Case A: start of the accrual end year")
    tf = f("Case A, from bid submission: factor in the payment year")
    st = rows[0]
    d.copy_format(ta, st)
    d.text(f"C{st}", "Case B: start of the accrual end year")
    d.text(f"D{st}", "date")
    acc = []
    for i, (name, ref, rate) in enumerate(FACTOR_LINES):
        r1, r2, r3 = rows[1 + 3 * i: 4 + 3 * i]
        for r, src, lab in ((r1, tf, "factor in the payment year (if before the end year)"),
                            (r2, tf + 1, "factor in the end year (to the accrual end)"), (r3, tf + 2, "accrual factor")):
            d.copy_format(src, r)
            d.text(f"C{r}", f"Case B, from {name}: {lab}")
            d.text(f"D{r}", "x")
        for c in P:
            x = ref.format(c=c)
            d.formula(f"{c}{r1}", f"=1+{rate}*(DATE(YEAR({x})+1,1,1)-{x})*({x}<{c}{st})/Days_In_Year")
            d.formula(f"{c}{r2}", f"=1+{rate}*MAX(0,{c}{e}-MAX({x},{c}{st})+1)/Days_In_Year")
            d.formula(f"{c}{r3}", f"={c}{r1}*(1+{rate})^MAX(0,YEAR({c}{st})-YEAR({x})-1)*{c}{r2}")
        acc.append(r3)
    for c in P:
        d.formula(f"{c}{st}", f"=DATE(YEAR({c}{e}),1,1)")
    # Claim rows: the user's factor rows (end+4, +6, …, +14) → the new accrual-factor rows.
    old = {m(q["B"]["end"] + 4 + 2 * i): r for i, r in enumerate(acc)}
    for r in range(m(q["B"]["tot"]) - 8, m(q["B"]["tot"])):
        for c in P:
            cell = d.cell(f"{c}{r}")
            new = re.sub(rf"\b{c}(\d+)\b", lambda h: f"{c}{old.get(int(h.group(1)), h.group(1))}", cell.getFormula())
            assert any(f"{c}{x}" in new for x in acc), (r, new)       # each claim reads a new factor row
            cell.setFormula(new)


def user29(d, a, blocks, q):
    """§29 edits on the user's cells, before any insert (user's numbering). Returns the user's cells that change on
    purpose, by reason, for the before / after comparison."""
    f = lambda t, s=1: d.find(t, start=s, exact=False)  # noqa: E731
    # 4. Dead "RtB price: 0" rows: drop their terms from the COD buyer cash flow and the COD cost-stack check.
    chk_cod = f("COD cost stack = timeline investment flows")
    for c, b in zip(P, blocks):
        rng = d.cell(f"J{b['codx']}:{LAST}{b['codx']}")
        old = rng.getFormulaArray()[0]
        tails = [f"+{x}{b['dead']}" for x in TL]
        assert all(o.endswith(t) for o, t in zip(old, tails)), old[:2]
        rng.setFormulaArray((tuple(o[:-len(t)] for o, t in zip(old, tails)),))
        cell, t = d.cell(f"{c}{chk_cod}"), f"+$H${b['dead']}"
        assert t in cell.getFormula(), cell.getFormula()
        cell.setFormula(cell.getFormula().replace(t, ""))
    # 5., 7.–9. Labels and notes.
    r_jv, r_case = f("JV with a developer"), f("Case: turbine price")
    r_bess, r_sunk = f("BESS (on BESS energy)"), f("Sunk fees and devex")
    r_tot, r_exp = f("Total cash to RtB (= total above)"), f("FEP exposure if the buyer pays")
    assert (r_jv, r_case, r_bess, r_sunk, r_tot, r_exp) == (21, 24, 90, 73, 76, 78)
    r_spend = [f("Spending", d.find(nm, f("Annual cash flow per project"))) for nm in NAMES]
    assert d.cell(f"C{r_spend[0]}").getString() == SPEND, r_spend
    d.text(f"C{r_jv}", JV_LABEL)
    d.text(f"C{r_case}", "Case: turbine price, PPA, O&M")
    d.text(f"C{r_bess}", "BESS (on BESS energy; points only, no BESS revenue or opex modelled)")
    for r in r_spend[1:]:
        d.text(f"C{r}", SPEND)
    # 8. Exposure block: land moves from "sunk" to "returned or reimbursed at cost" (row 76 total unchanged).
    eur = f("Guarantees, securities and connection fees, DEVEX (EUR k)")
    e = {k: f(lab, eur) for k, lab in (("fee", "Participation fee"), ("bs", "Bid security"), ("tso", "TSO connection"),
                                       ("dev", "DEVEX – development"), ("land", "DEVEX – land"),
                                       ("sign", "Signing fee and DD"))}
    assert e == {"fee": 61, "bs": 62, "tso": 64, "dev": 67, "land": 68, "sign": 69}, e
    d.text(f"C{r_sunk}", "Sunk fees and devex (signing and DD, participation fee, development devex)")
    d.text(f"C{r_sunk + 1}", "Securities and land, returned or reimbursed at cost (bid, performance, TSO; land)")
    for c in P:
        d.formula(f"{c}{r_sunk}", f"={c}{e['fee']}+{c}{e['dev']}+{c}{e['sign']}")
        d.formula(f"{c}{r_sunk + 1}", f"=SUM({c}{e['bs']}:{c}{e['tso']})+{c}{e['land']}")
    assert d.cell(f"J{r_tot}").getString().startswith("Case A") and d.cell(f"J{r_exp}").getString().startswith("Case B")
    d.text(f"J{r_tot}", "Case B: FEP pre-funds everything")                 # notes column J ("Who bears it")
    d.text(f"J{r_exp}", "Case A: FEP funds sunk fees, devex and securities only; the Case A detail counts the devex "
                        "paid by sale A")
    # 9. Premium computed once (Dashboard row 98); the profit stacks link to it too (no link to a link).
    prem = f("Developer premium paid for the RtB project")
    for k in "BA":
        for c in P:
            for r in (q[k]["wf_prem"], f("RtB premium paid by the buyer", q[k]["ps"])):
                d.formula(f"{c}{r}", f"={c}{prem}")
    # §30.2: "Exit fees" out of the two profit-stack labels.
    r_sign = [d.find("Less: signing fee and DD, Exit fees", q[k]["ps"]) for k in "BA"]
    for r in r_sign:
        d.text(f"C{r}", "Less: signing fee and DD")
    # The user's cells that change on purpose (user's numbering), by reason; the base is recalculated with the user's
    # signing fee input (135) first, so its effects are not listed here.
    cod = [a["sale_cod"], a["sale_cod"] + 1, a["cod_irr"] - 1, f("Buyer profit, post-tax")]
    cells = lambda rows, cols: [[r, c] for r in rows for c in cols]  # noqa: E731
    return {"Dashboard": {
        "relabel": cells([r_jv, r_case, r_bess, r_sunk, r_sunk + 1, a["hold"], *r_spend[1:], q["A"]["tie"],
                          q["B"]["tie"], q["A"]["s_profit"], q["B"]["s_profit"], *r_sign, a["sale_cod"],
                          *(b["val"] for b in blocks)], "C") + cells([r_tot, r_exp], "J"),
        "exposure rows: land returned at cost": cells([r_sunk, r_sunk + 1], "EFGH"),
        "COD price on the unlevered post-tax cash flow": cells(cod, "EFGH") +
        cells([r for b in blocks for r in (b["val"], b["codx"])], ["H"] + TL),
        "IRR now calculates (FILTER array → XIRR)": cells([a["cod_irr"], a["irr"]], "EFG"),
        "portfolio IRR (Total column)": cells([q[k][r] for k in "AB" for r in ("s_irr", "pre_irr")], "H"),
        "post-tax profit incl. loan CIT (§28)": cells([q[k][r] for k in "AB" for r in ("post", "s_profit")], "EH"),
        "factor cell (§28)": cells(sorted(FACTOR_ROWS), "MNO"),
        "Case B claim factors rebuilt (Case A pattern)": cells(range(q["B"]["end"] + 1, q["B"]["end"] + 15), "CDEFGH"),
        "deleted": cells(sorted(b["dead"] for b in blocks) + [f("Claim lines paid before the accrual end year")], "C")},
        "Inputs": {"relabel": [[39, "C"], [39, "H"], [40, "H"], [47, "H"]]},
        "irr_rows": [a["cod_irr"], a["irr"]], "rtb_user": {k: {"pre": x["pre"]} for k, x in q.items()}}


def edit29(d, m, R, q, Q):
    """§29 rows after the inserts: the tie checks, portfolio IRRs, the summary error check (the devex-principal rows
    of §29.1 are reverted by §30.1)."""
    for k, x in Q.items():
        # 3b. Tie check replaces the tautological "waterfall closes" check.
        r = x["tie"]
        d.text(f"C{r}", f"Case {k} dated cash flows tie to the post-tax profit")
        for c in P:
            d.formula(f"{c}{r}", f"=(ABS(SUM({c}{x['first']}:{c}{x['exit']})-{c}{x['post']})>Tolerance)*1")
        # 2. Portfolio IRR in the Total column (dates common to all projects) and its summary link.
        for r, last in ((x["pre_irr"], x["rec"]), (x["irr"], x["exit"])):
            d.formula(f"H{r}", f"=XIRR(H{x['first']}:H{last},J{x['first']}:J{last})")
        for r, src in ((x["s_irr"], x["pre_irr"]), (x["s_irr_post"], x["irr"])):
            d.formula(f"H{r}", f"=H{src}")
        for r in (x["pre_irr"], x["irr"], x["s_irr"], x["s_irr_post"]):
            d.cell(f"H{r}").NumberFormat = d.cell(f"E{r}").NumberFormat
    # 3a. No errors in the summary block (Total column only: one count for the block).
    r = R["check_err"]
    d.copy_format(R["check"], r, last="H")
    d.text(f"C{r}", ERR_CHECK)
    d.text(f"D{r}", "flag")
    d.formula(f"H{r}", f"=SUMPRODUCT(--ISERROR(E{R['s_top'] + 1}:H{R['s_dscr']}))")


def anchors29(d, q):
    """§29 anchors in the user's numbering, per RtB case: premium, profit stack, receipts, pre-tax returns, tie check."""
    f = lambda t, s=1: d.find(t, start=s, exact=False)  # noqa: E731
    heads = {"B": (f("RtB exit, Case B: waterfall at the sale"), f("FEP profit, Case B (at the sale)"),
                   f("Case B waterfall closes")),
             "A": (f("Waterfall at the sale, Case A"), f("FEP profit, Case A (at sale A)"), f("Case A waterfall closes"))}
    for k, (wf, ps, tie) in heads.items():
        x = q[k]
        x.update(wf_prem=d.find("RtB premium", wf), ps=ps, rs_rec=f("FEP receipts at", ps),
                 pre=d.find("FEP profit, pre-tax", ps), pre_moic=f("MOIC, pre-tax (dated", ps),
                 pre_irr=f("IRR, pre-tax (dated", ps), tie=tie)
        assert x["pre"] == x["post"] - 2, (k, x)
    got = {k: tuple(q[k][i] for i in ("rs_rec", "pre", "tie")) for k in q}
    assert got == {"B": (306, 310, 432), "A": (260, 264, 441)}, got


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
                d.cell(f"{fc}{r}").NumberFormat = d.cell(f"E{Q['B']['end'] + 2}").NumberFormat   # 0.0000
    # 2. Loan sub-blocks after each project's COD-exit row.
    for c, fc, dc, b, blk in zip(P, fcols, dcols, B, blocks):
        rows = m.block(blk["codx"] + 1)
        d.copy_format(R["hold"], rows[0])
        d.text(f"C{rows[0]}", LOAN_HEAD)
        it = iter(rows[1:])
        for k in "BA":
            x = Q[k]
            rr = {key: next(it) for key, _ in LOAN_ROWS}
            fm = {"cash": f"=({{w}}{rr['cash']}*(1+Rate_Cash)-SUMPRODUCT(${c}{x['first'] + 1}:${c}{x['land'] - 1}"
                          f"*${fc}{x['first'] + 1}:${fc}{x['land'] - 1}*(YEAR(${dc}{x['first'] + 1}:${dc}{x['land'] - 1})"
                          f"={{x}}{R['year']})))*({{x}}{R['ye']}<${c}{x['end']})",
                  "dev": f"=({{w}}{rr['dev']}*(1+Rate_Devex)-SUMPRODUCT(${c}{x['land']}:${c}{x['land'] + 1}"
                         f"*${fc}{x['land']}:${fc}{x['land'] + 1}*(YEAR(${dc}{x['land']}:${dc}{x['land'] + 1})"
                         f"={{x}}{R['year']})))*({{x}}{R['ye']}<${c}{x['end']})",
                  "drawn": f"=-SUMPRODUCT(${c}{x['first'] + 1}:${c}{x['land'] + 1}*(YEAR(${dc}{x['first'] + 1}:"
                           f"${dc}{x['land'] + 1})={{x}}{R['year']}))*({{x}}{R['ye']}<${c}{x['end']})",
                  "int": f"=({{x}}{rr['cash']}+{{x}}{rr['dev']}-{{w}}{rr['cash']}-{{w}}{rr['dev']}-{{x}}{rr['drawn']})"
                         f"*({{x}}{R['ye']}<${c}{x['end']})"}
            fm = {key: v + f"*${c}21" for key, v in fm.items()}      # §29.4: no loans without a JV
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
        years, rest, ex, moic, irr = rows[:CIT_YEARS], *rows[CIT_YEARS:]
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
        x_cit = m(q[k]["post"] - 1)                               # §29: the profit stack's exit-CIT row, dated
        assert d.cell(f"C{x_cit}").getString().startswith("CIT on the exit gain"), k
        d.copy_format(x["land"], ex, last="L")
        d.text(f"C{ex}", d.cell(f"C{x_cit}").getString())
        d.text(f"D{ex}", "EUR k")
        for c, dc in zip(P, dcols):
            d.formula(f"{c}{ex}", f"={c}{x_cit}")
            d.formula(f"{dc}{ex}", f"={x['sale']}")
        d.formula(f"H{ex}", f"=SUM(E{ex}:G{ex})")
        rest_ = rest
        rest = ex                                                 # post-tax ranges run through the exit-CIT row
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
        x.update(cit_years=years, rest=rest_, exit=ex, moic=moic, irr=irr)
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


def build():
    unused_names(SRC)
    with Office(SRC) as lo:                                   # Office recalculates on load
        set_signing(lo.doc)                                   # the user's input change, in the base too
        lo.calculate()
        lo.save(BASE_RC)
        extend_timeline(lo.doc)
        mi, ri, inp_cells = edit_inputs(lo.doc)
        md, R, B, Q, U = edit_dashboard(lo.doc)
        irr, rtb_user = set(U.pop("irr_rows")), U.pop("rtb_user")  # the user's FILTER IRRs, now plain XIRR
        for n in UNUSED_NAMES:                                # before keep_arrays: a name change drops the arrays
            lo.doc.NamedRanges.removeByName(n)
        keep_arrays(lo.doc, [x for x in array_cells(SRC) if not (x[0] == "Dashboard" and x[2] in irr)],
                    {"Inputs": mi, "Dashboard": md})
        keep_groups(lo.doc, user_groups(SRC), md)
        lo.calculate()
        lo.save(OUT)
    with Office(OUT) as lo:
        lo.save(RC)
    info = {"Inputs": {"inserts": mi.ins, "rows": ri},
            "Dashboard": {"inserts": md.ins, "deletes": md.dels, "rows": R, "blocks": dict(zip(NAMES, B)), "rtb": Q},
            "expected": U | {"irr_rows": sorted(irr)}, "rtb_user": rtb_user}
    info["expected"]["Inputs"]["relabel"] += inp_cells
    MAPS.write_text(json.dumps(info, indent=1, ensure_ascii=False))
    print(f"wrote {OUT}")


def unused_names(path):
    """§29.9: the names to delete are not used in any formula or other name of the user's file."""
    wb = openpyxl.load_workbook(path)
    text = [d.attr_text or "" for d in wb.defined_names.values()]
    text += [c.value.text if isinstance(c.value, ArrayFormula) else c.value for ws in wb.worksheets
             for row in ws.iter_rows() for c in row if isinstance(c.value, (str, ArrayFormula))]
    used = [n for n in UNUSED_NAMES if any(re.search(rf"\b{n}\b", t) for t in text if t.startswith("="))
            or any(re.search(rf"\b{n}\b", d.attr_text or "") for k, d in wb.defined_names.items() if k != n)]
    assert all(n in wb.defined_names for n in UNUSED_NAMES) and not used, used


# ---------------------------------------------------------------------------------------------------------------------
def remap(sheet):
    m, x = Rows2(), json.loads(MAPS.read_text())[sheet]
    for p, n in x["inserts"]:
        m.add(p, n)
    for r in x.get("deletes", []):
        m.delete(r)
    return m


def compare(b_wb, a_wb):
    """Every cell of the user's two sheets (values after recalc) against its shifted position in the new file."""
    out = []
    for name in ("Dashboard", "Inputs"):
        b, a, m = b_wb[name], a_wb[name], remap(name)
        for r in range(1, b.max_row + 1):
            if m(r) is None:
                continue
            for c in range(1, b.max_column + 1):
                x, y = b.cell(r, c).value, a.cell(m(r), c).value
                if not ((x == y) or (isinstance(x, (int, float)) and isinstance(y, (int, float))
                                     and abs(x - y) <= 1e-9 * max(1, abs(x)))):
                    out.append((name, r, m(r), b.cell(r, c).column_letter, x, y))
    return out


def cells_check(b_wb, a_wb, skip):
    """Formulas: same count of formula cells per old row after the shift (a cheap structure check)."""
    bad = []
    for name in ("Dashboard", "Inputs"):
        b, a, m = b_wb[name], a_wb[name], remap(name)
        for r in range(1, b.max_row + 1):
            if m(r) is None or (name, r) in skip:
                continue
            nb = sum(1 for c in range(1, b.max_column + 1) if b.cell(r, c).value is not None)
            na = sum(1 for c in range(1, b.max_column + 1) if a.cell(m(r), c).value is not None)
            if nb != na:
                bad.append((name, r, m(r), nb, na))
    return bad


def recompute(wb, info):
    """Hold case and COD exit, independently (§30.3–30.6): from the user's recalculated rows (years from COD, months,
    construction and operating shares, EBITDA, pre-tax and investment cash flow) and the capex cells, the tax and debt
    schedule as the Bora model, the COD price (PV of EBITDA + unlevered CIT) and the COD-exit MOIC / IRR."""
    dash, inp, nm_ = wb["Dashboard"], wb["Inputs"], names(wb)
    iv = lambda name: inp[f"E{info['Inputs']['rows'][name]}"].value  # noqa: E731
    life, gear, rate, ten, fee = (iv(k) for k in ("Tax_Dep_Life", "Gearing", "Debt_Rate", "Debt_Tenor", "Debt_Fee"))
    cit_r, dr, mo12 = nm_["CIT_Rate"], nm_["Discount_Rate"], nm_["Months_In_Year"]
    R = info["Dashboard"]["rows"]
    cols = range(col("J") + 1, col(LAST) + 2)
    ye = [dash.cell(R["ye"], c).value for c in cols]
    res, ok = {}, True
    for c, nm in zip(P, NAMES):
        b = info["Dashboard"]["blocks"][nm]
        row = lambda r: [dash.cell(r, k).value for k in cols]  # noqa: E731
        tn, op, eb, pre, mo, cs, inv = (row(b[k]) for k in ("t", "op", "eb", "pre", "mo", "cs", "inv"))
        cm = dash[f"{c}{b['cm']}"].value
        cod = inp[f"{c}{next(r for r in range(1, 60) if inp.cell(r, 3).value == 'COD')}"].value
        base = sum(dash[f"{c}{R[k]}"].value for k in ("capx", "conn", "muni"))
        dsh = [max(0, min(m_, life * mo12) - max(m_ - mo12, 0)) / mo12 for m_ in mo]
        dep = [-base / life * x for x in dsh]
        cit = [-max(0, e + x) * cit_r for e, x in zip(eb, dep)]
        post = [p + t for p, t in zip(pre, cit)]
        size = dash[f"{c}{R['allin']}"].value * gear                     # amendment: on the all-in cost
        bal, prev_cs, draws, idcs, fees, ints, dss, eq, dscr, citl = 0.0, 0.0, [], [], [], [], [], [], [], []
        for n, s_, o in zip(tn, cs, op):
            dw = size * s_
            idc = (bal + dw * iv("IDC_Draw_Share")) * rate * s_ * cm / mo12
            fees.append(-size * fee * (s_ > 0) * (prev_cs == 0))
            i = -(bal + dw + idc) * rate * o
            if n >= 1 and n <= ten:
                p0 = sum(draws) + dw + sum(idcs) + idc               # all draws and IDC fall before COD
                ds = -p0 * rate / (1 - (1 + rate) ** -ten)
            else:
                ds = i if n == 0 else 0.0
            bal += dw + idc + (ds - i)
            draws.append(dw)
            idcs.append(idc)
            ints.append(i)
            dss.append(ds)
            prev_cs = s_
        depi = [-sum(idcs) / life * x for x in dsh]
        for e, x, y, i, p, dw, fe, ds in zip(eb, dep, depi, ints, pre, draws, fees, dss):
            ct = -max(0, e + x + y + i + fe) * cit_r                       # the fee deducted in the year paid
            citl.append(ct)
            eq.append(p + dw + fe + ds + ct)
            if ds < 0:
                dscr.append(-(e + ct) / ds)
        df = [(1 + dr) ** ((ser(cod) - ser(t)) / 365) for t in ye]
        price = sum((e + t) * f_ for e, t, f_ in zip(eb, cit, df))
        codcf = [v * (n <= 0) + (n == 0) * (price + dash[f"{c}62"].value + dash[f"{c}63"].value)
                 for v, n in zip(inv, tn)]
        nz = lambda v: [(t.date(), x) for t, x in zip(ye, v) if x != 0]  # noqa: E731
        moic = sum(v for v in codcf if v > 0) / -sum(v for v in codcf if v < 0)
        py = {"depreciable base": base, "debt size": size, "debt at COD": sum(draws) + sum(idcs), "IDC": sum(idcs), "fee": sum(fees),
              "CIT unlevered total": sum(cit), "project IRR pre-tax": xirr(nz(pre)),
              "project IRR post-tax": xirr(nz(post)), "equity IRR": xirr(nz(eq)), "min DSCR": min(dscr),
              "CIT levered total": sum(citl), "balance at the end": bal, "COD price": price, "COD MOIC": moic,
              "COD IRR": xirr(nz(codcf))}
        model = {"depreciable base": dash[f"{c}{b['base']}"].value, "debt size": dash[f"{c}{b['size']}"].value,
                 "debt at COD": dash[f"{c}{R['s_debt']}"].value,
                 "IDC": dash[f"H{b['idc']}"].value, "fee": dash[f"H{b['fee']}"].value,
                 "CIT unlevered total": dash[f"{c}{R['s_cit']}"].value, "project IRR pre-tax": dash[f"{c}{R['irr']}"].value,
                 "project IRR post-tax": dash[f"{c}{R['s_irr']}"].value, "equity IRR": dash[f"{c}{R['s_eirr']}"].value,
                 "min DSCR": dash[f"{c}{R['s_dscr']}"].value, "CIT levered total": dash[f"H{b['citl']}"].value,
                 "balance at the end": dash[f"{LAST}{b['bal']}"].value, "COD price": dash[f"{c}{R['sale_cod']}"].value,
                 "COD MOIC": dash[f"{c}{R['cod_irr'] - 1}"].value, "COD IRR": dash[f"{c}{R['cod_irr']}"].value}
        for k in py:
            tol = 1e-6 if ("IRR" in k or "DSCR" in k or "MOIC" in k) else 0.01
            ok &= abs(py[k] - model[k]) <= tol
        diff = max(abs(py[k] - model[k]) for k in py)
        res[nm] = model
        print(f"   {nm:<11} " + " | ".join(f"{k} {model[k]:,.4f}" for k in model) + f"   (max python diff {diff:.1e})")
    print("independent recomputation (hold case: tax and levered schedule; COD price, MOIC, IRR):",
          "PASS" if ok else "FAIL")
    return ok, res


def names(wb):
    """Workbook name → value (single cell) or list of values (range)."""
    out = {}
    for n, dn in wb.defined_names.items():
        dest = list(dn.destinations)
        if len(dest) != 1:
            continue
        sh, ref = dest[0]
        v = wb[sh][ref.replace("$", "")]
        out[n] = [c[0].value for c in v] if isinstance(v, tuple) else v.value
    return out


def ser(t):
    return (t - datetime(1899, 12, 30)).total_seconds() / 86400


def dt(x):
    return datetime(1899, 12, 30) + timedelta(days=x)


def factor(pay, end, r):
    """Claim accrual factor (serial dates): simple to the first 31-Dec, compounded yearly, simple in the end year."""
    y0 = ser(datetime(dt(end).year, 1, 1))
    f1 = 1 + r * (ser(datetime(dt(pay).year + 1, 1, 1)) - pay) / 365 if pay < y0 else 1
    f2 = 1 + r * max(0, end - max(pay, y0) + 1) / 365
    return f1 * (1 + r) ** max(0, dt(end).year - dt(pay).year - 1) * f2


def yearfrac30(a, b):
    """YEARFRAC basis 0 (US 30/360)."""
    feb_end = lambda t: t.month == 2 and (t + timedelta(days=1)).month == 3  # noqa: E731
    d1, d2 = a.day, b.day
    if feb_end(a):
        if feb_end(b):
            d2 = 30
        d1 = 30
    if d2 == 31 and d1 >= 30:
        d2 = 30
    d1 = 30 if d1 == 31 else d1
    return ((b.year - a.year) * 360 + (b.month - a.month) * 30 + d2 - d1) / 360


def recompute29(wb, base, info, yearly):
    """Both RtB cases from the inputs: claims (accrual factors), DSA, waterfall (no devex principal: §30.1), FEP receipts, profit pre / post-tax, dated MOIC and IRR (pre / post-tax) and the portfolio IRR (Total column)."""
    dash, inp, nm_ = wb["Dashboard"], wb["Inputs"], names(wb)
    Q, d = info["Dashboard"]["rtb"], lambda c, r: dash[f"{c}{r}"].value
    lab = lambda t: next(r for r in range(1, 200) if str(inp.cell(r, 3).value or "").startswith(t))  # noqa: E731
    rows = {k: lab(t) for k, t in (("sign", "Signing fee and DD paid"), ("bid", "Bid submission"),
                                   ("pb", "Performance security posted"), ("tso", "TSO security paid"),
                                   ("hc", "Connection agreement (HCSSZ)"), ("dsa", "DSA accrual start"),
                                   ("scale", "Devex scale"))}
    items = list(zip(nm_["Devex_HUF"], nm_["Devex_Date"], nm_["Devex_Class"]))
    rc, rd, floor, cit = nm_["Rate_Cash"], nm_["Rate_Devex"], nm_["Devex_Floor"], nm_["CIT_Rate"]
    ok, port, out = True, {}, {}
    for c, nm in zip(P, NAMES):
        iv = lambda k: inp[f"{c}{rows[k]}"].value  # noqa: E731
        dte = {k: ser(iv(k)) for k in ("sign", "bid", "pb", "tso", "hc", "dsa")}
        e = {k: d(c, r) for k, r in (("fee", 61), ("bs", 62), ("pb", 63), ("tso", 64), ("dir", 65), ("ind", 66),
                                    ("sign", 69))}
        jv, share, prem = d(c, 21), d(c, 19), d(c, 18) * d(c, 6)
        for k in "BA":
            x = Q[k]
            sale = ser(nm_["Sale_Date" if k == "B" else "Sale_Date_A"])
            end = min(ser(nm_["Freeze_Date"]), sale)
            dev = {}
            for cls in ("Land", "Devex"):                        # amount-weighted date, × scale ÷ FX
                its = [(h, ser(t)) for h, t, cl in items if cl == cls and (k == "B" or ser(t) <= sale)]
                hs = sum(h for h, _ in its)
                dev[cls] = (hs * iv("scale") / nm_["FX"] / nm_["Thousand"], sum(h * t for h, t in its) / hs)
            cash = [("fee", "bid"), ("bs", "bid"), ("pb", "pb"), ("tso", "tso")] + \
                   ([("dir", "hc"), ("ind", "hc")] if k == "B" else [])
            c12 = jv * sum(e[a] * (factor(dte[t], end, rc) - 1) for a, t in cash)
            cdev = jv * sum(v * max(factor(t, end, rd) - 1, floor - 1) for v, t in dev.values())
            dsa = d(c, 20) * min(nm_["DSA_Rate"] * yearfrac30(iv("dsa"), dt(sale)), nm_["DSA_Cap"]) * d(c, 44)
            reimb = sum(e[a] for a, _ in cash if a != "fee")
            land, devex = dev["Land"][0], dev["Devex"][0]
            av = prem + reimb + land - min(prem + reimb + land, nm_["Sale_Costs"])
            paid = {}
            for key, due in (("sec", reimb), ("land", land), ("c12", c12), ("ret", cdev),
                             ("dsa", dsa)):
                paid[key] = min(av, due)
                av -= paid[key]
            receipts = paid["c12"] + paid["ret"] + av * share
            pre = receipts - e["sign"] - e["fee"] - devex
            post = pre - max(0, pre) * cit * (1 - nm_["Share_Exempt"]) - (c12 + cdev) * cit * jv
            flows = [(dte["sign"], -e["sign"])] + [(dte[t], -e[a]) for a, t in cash] + \
                    [(dev["Land"][1], -land), (dev["Devex"][1], -devex), (sale, receipts + paid["sec"] + paid["land"])]
            flows = [(dt(t).date(), v) for t, v in flows]
            yl = yearly[(nm, k)]
            post_f = flows + [(datetime(2026 + i, 12, 31).date(), v) for i, v in enumerate(yl["yearly"])] + \
                [(dt(sale).date(), yl["rest"]), (dt(sale).date(), post - pre + (c12 + cdev) * cit * jv)]
            moic = lambda fl: sum(v for _, v in fl if v > 0) / -sum(v for _, v in fl if v < 0)  # noqa: E731
            py = {"claims": c12 + cdev, "receipts": receipts, "profit pre-tax": pre, "profit post-tax": post,
                  "MOIC pre-tax": moic(flows), "IRR pre-tax": xirr(sorted(flows)), "MOIC post-tax": moic(post_f),
                  "IRR post-tax": xirr(sorted(post_f))}
            model = {"claims": d(c, x["tot"]), "receipts": d(c, x["rs_rec"]), "profit pre-tax": d(c, x["pre"]),
                     "profit post-tax": d(c, x["post"]), "MOIC pre-tax": d(c, x["pre_moic"]),
                     "IRR pre-tax": d(c, x["pre_irr"]), "MOIC post-tax": d(c, x["moic"]),
                     "IRR post-tax": d(c, x["irr"])}
            diff = max(abs(py[t] - model[t]) for t in py)
            ok &= diff < 1e-6
            # §30.1 revert: the pre-tax profit equals the base's (recalculated with signing / DD 135).
            u = info["rtb_user"][k]["pre"]
            delta = model["profit pre-tax"] - base["Dashboard"][f"{c}{u}"].value
            ok &= abs(delta) < 1e-6
            port.setdefault(k, ([], []))
            port[k][0].extend(flows)
            port[k][1].extend(post_f)
            out[(nm, k)] = model
            print(f"   {nm:<11} Case {k}: " + " | ".join(f"{t} {model[t]:,.4f}" for t in model) +
                  f" | Δ pre-tax profit vs base {delta:,.2f} (python diff {diff:.1e})")
    for k, (fl, fl_post) in port.items():
        x = Q[k]
        py, model = (xirr(sorted(fl)), xirr(sorted(fl_post))), (d("H", x["pre_irr"]), d("H", x["irr"]))
        ok &= max(abs(a - b) for a, b in zip(py, model)) < 1e-6 and model == (d("H", x["s_irr"]), d("H", x["s_irr_post"]))
        print(f"   portfolio Case {k}: IRR pre-tax {model[0]:.4%}, post-tax {model[1]:.4%} (python {py[0]:.4%}, {py[1]:.4%})")
    print("independent recomputation (RtB waterfalls):", "PASS" if ok else "FAIL")
    return ok, out


def scratch(info):
    """Scratch session (not saved): Low / High case flip (Master_Check, error cells), then one broken dated flow."""
    R, Q, out = info["Dashboard"]["rows"], info["Dashboard"]["rtb"], {}
    with Office(OUT) as lo:
        sheets = [lo.doc.Sheets.getByIndex(i) for i in range(lo.doc.Sheets.Count)]

        def state():
            lo.calculate()
            n = sum((x.EndRow - x.StartRow + 1) * (x.EndColumn - x.StartColumn + 1)
                    for s in sheets for x in s.queryFormulaCells(4).getRangeAddresses())   # FormulaResult.ERROR
            return lo.get("Master_Check"), n
        for case in ("Low", "High", "Base"):
            lo.set("Dash_Case", case)
            out[case] = state()
        g = lambda c, r: lo.get(f"Dashboard!{c}{r}")  # noqa: E731
        lo.set("Share_Exempt", 0)                                 # exit gains taxed: the dated exit-CIT rows bite
        ex = state() + tuple(g("H", Q[k]["tie"]) for k in "BA")
        rtb = {nm: {k: (g(c, Q[k]["post"]), g(c, Q[k]["pre_irr"]), g(c, Q[k]["irr"])) for k in "BA"}
               for c, nm in zip(P, NAMES)}
        print(f"   Share_Exempt = 0: Master_Check {ex[0]} | error cells {ex[1]} | tie checks B {ex[2]}, A {ex[3]}")
        for nm, v in rtb.items():
            print(f"      {nm:<11} " + " | ".join(f"Case {k}: profit post-tax {p_:,.2f}, IRR pre {a:.4%} → post {b:.4%}"
                                               for k, (p_, a, b) in v.items()))
        ok_ex = ex == (0, 0, 0, 0) and all(b < a for v in rtb.values() for _, a, b in v.values())
        lo.set("Share_Exempt", 1)
        ok_31 = flow31(lo, info, state)
        cell = lo.doc.Sheets.getByName("Dashboard").getCellRangeByName(f"E{Q['B']['first']}")
        cell.setValue(cell.getValue() - 1)                        # break the first Case B dated flow by 1 k EUR
        out["broken"] = state() + (lo.get(f"Dashboard!H{Q['B']['tie']}"),)
    for k, (mc, n, *t) in out.items():
        print(f"   {k:<6} Master_Check {mc} | error cells {n}" + (f" | Case B tie check {t[0]}" if t else ""))
    return all(out[k] == (0, 0) for k in ("Low", "High", "Base")) and out["broken"][0] == 1 and \
        out["broken"][2] == 1 and ok_ex and ok_31


def claims_lo(lo, info):
    """Case B claims per project recomputed in Python from the live LibreOffice values (act/365, capitalised 31-Dec);
    max difference to the model's claim rows."""
    Q, g = info["Dashboard"]["rtb"]["B"], lo.get
    rc, rd, floor = g("Rate_Cash"), g("Rate_Devex"), g("Devex_Floor")
    end = min(g("Freeze_Date"), g("Sale_Date"))
    diff = 0.0
    for c in P:
        jv, dte = g(f"Dashboard!{c}21"), {k: g(f"Inputs!{c}{r}") for k, r in (("bid", 21), ("pb", 22), ("tso", 23),
                                                                               ("hc", 24))}
        dte.update(land=g("Dashboard!J56"), dev=g("Dashboard!J55"))
        amt = lambda r: g(f"Dashboard!{c}{r}")  # noqa: E731
        py = [amt(r) * (factor(dte[k], end, rc) - 1) for r, k in ((61, "bid"), (62, "bid"), (63, "pb"), (64, "tso"),
                                                                   (65, "hc"), (66, "hc"))]
        py += [amt(r) * max(factor(dte[k], end, rd) - 1, floor - 1) for r, k in ((68, "land"), (67, "dev"))]
        model = [g(f"Dashboard!{c}{r}") for r in range(Q["tot"] - 8, Q["tot"])]
        diff = max([diff] + [abs(jv * a - b) for a, b in zip(py, model)])
    return diff


def flow31(lo, info, state):
    """§31.3 flow-through tests in the scratch session: Iváncsa COD +1 year; Bora HCSSZ in the accrual end year; the
    freeze date at 31-Dec-2028. Each input is set back afterwards."""
    from datetime import date
    sv = lambda t: (date(*t) - date(1899, 12, 30)).days  # noqa: E731
    op = info["Dashboard"]["blocks"]["Iváncsa"]["op"]
    ok = True
    for ref, val, what in (("Inputs!F26", sv((2032, 6, 30)), "Iváncsa COD 30-Jun-2032"),
                           ("Inputs!E24", sv((2029, 3, 1)), "Bora HCSSZ 01-Mar-2029"),
                           ("Freeze_Date", sv((2028, 12, 31)), "Freeze_Date 31-Dec-2028")):
        old = lo.get(ref)
        lo.set(ref, val)
        mc, n = state()
        life, dif = lo.get(f"Dashboard!H{op}"), claims_lo(lo, info)
        print(f"   {what}: Master_Check {mc} | error cells {n} | Iváncsa operating years {life} | Case B claims "
              f"vs Python max diff {dif:.1e}")
        ok &= mc == 0 and n == 0 and abs(life - 30) < 1e-9 and dif < 1e-9
        lo.set(ref, old)
    return ok


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


# User cells that §27 / §28 change on purpose (user's numbering); §29's are listed by build() in rows_v3.json.
RELABELS = {("Dashboard", "C134"), ("Dashboard", "C113"), ("Dashboard", "C119"), ("Inputs", "C39"), ("Inputs", "H39")}
FACTOR_ROWS = {270, *range(272, 278), 315, *range(317, 325)}


def expected(info, sh, cl, r0):
    if (sh, f"{cl}{r0}") in RELABELS:
        return "relabel"
    for why, cells in info["expected"][sh].items():
        if [r0, cl] in cells:
            return why
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
                             * (1 + rt) ** (y - d.year) for p, d, rt in lines if d.year <= y) * (y < end.year) * jv
            drawn = {y: -sum(p for p, d, _ in lines if d.year == y) * (y < end.year) * jv for y in bal}
            ints = {y: (bal[y] - bal.get(y - 1, 0) - drawn[y]) * (y < end.year) for y in bal}
            yearly = [-ints[2026 + i] * cit * jv for i in range(CIT_YEARS)]
            total = -dash[f"{c}{x['tot']}"].value * cit * jv
            flows = [(dash[f"{dc}{r}"].value.date(), dash[f"{c}{r}"].value) for r in range(x["first"], x["rec"] + 1)]
            flows += [(date(2026 + i, 12, 31), v) for i, v in enumerate(yearly)] + [(sale, total - sum(yearly))]
            pre_p = dash[f"{c}{x['pre']}"].value                         # dated exit CIT (0 while exempt)
            flows.append((sale, -max(0, pre_p) * cit * (1 - name("Share_Exempt"))))
            pos, neg = sum(v for _, v in flows if v > 0), -sum(v for _, v in flows if v < 0)
            py = {"yearly CIT": yearly, "rest at sale": total - sum(yearly), "total loan CIT": total,
                  "MOIC post-tax": pos / neg, "IRR post-tax": xirr(sorted(flows))}
            model = {"yearly CIT": [dash[f"{c}{r}"].value for r in x["cit_years"]],
                     "rest at sale": dash[f"{c}{x['rest']}"].value, "total loan CIT": dash[f"{c}{x['loan_cit']}"].value,
                     "MOIC post-tax": dash[f"{c}{x['moic']}"].value, "IRR post-tax": dash[f"{c}{x['irr']}"].value}
            blk = info["Dashboard"]["blocks"][nm]                       # the model's loan sub-block interest row
            loans = [dash[f"{cc}{blk[f'loan{k}_int']}"].value for cc in TL[:CIT_YEARS]]
            diff = max([abs(a - b) for a, b in zip(py["yearly CIT"], model["yearly CIT"])] +
                       [abs(py[t] - model[t]) for t in py if t != "yearly CIT"] +
                       [abs(ints[2026 + i] - v) for i, v in enumerate(loans)])
            ok &= diff < 1e-6
            out[(nm, k)] = {"yearly": yearly, "rest": total - sum(yearly)}
            print(f"   {nm:<11} Case {k}: yearly CIT " + ", ".join(f"{v:,.2f}" for v in model["yearly CIT"]) +
                  f" | rest {model['rest at sale']:,.2f} | total {model['total loan CIT']:,.2f} | MOIC post "
                  f"{model['MOIC post-tax']:.4f} | IRR post {model['IRR post-tax']:.4%}  (python diff {diff:.1e})")
    print("independent recomputation (RtB loan CIT):", "PASS" if ok else "FAIL")
    return ok, out


def verify():
    info = json.loads(MAPS.read_text())
    a_wb = openpyxl.load_workbook(RC, data_only=True)
    b_wb = openpyxl.load_workbook(BASE_RC, data_only=True)
    R = info["Dashboard"]["rows"]
    errs = lint_values(RC)
    print(f"recalc: {len(errs)} error(s) / Master_Check issue(s)", [str(e) for e in errs])
    print("   Master_Check (Dashboard H3):", a_wb["Dashboard"]["H3"].value, "| H", R["master"], "=",
          a_wb["Dashboard"][f"H{R['master']}"].value)
    base_err = [e.where for e in lint_values(BASE_RC) if e.rule == "f-error"]
    print("   base file, same LibreOffice (the user's FILTER IRRs):", base_err)
    with zipfile.ZipFile(OUT) as z:
        filt = [n for n in z.namelist() if n.startswith("xl/") and b"filter(" in z.read(n).lower()]
    print("   FILTER left in the file:", filt or "none")
    diffs = compare(b_wb, a_wb)
    bad = [x for x in diffs if not expected(info, x[0], x[3], x[1])]
    print(f"before / after: {len(diffs)} changed cell(s) in the user's cells, {len(bad)} unexpected")
    for sh, r0, r1, cl, x, y in diffs:
        print(f"   {sh}!{cl}{r0} → {cl}{r1}: {x!r} → {y!r}  [{expected(info, sh, cl, r0) or 'UNEXPECTED'}]")
    print("   deleted user rows (dead 'RtB price: 0' rows; claim-timing check):", [r for r, _ in info["expected"]["Dashboard"]["deleted"]])
    skip = {("Dashboard", r) for r in FACTOR_ROWS} | {("Inputs", 40)} | {("Dashboard", r) for r, _ in
                                                       info["expected"]["Dashboard"]["portfolio IRR (Total column)"]}
    struct = cells_check(openpyxl.load_workbook(SRC), openpyxl.load_workbook(OUT), skip)
    print(f"   non-empty cells per row (formulas / text) vs the user's file: {len(struct)} row(s) differ", struct[:5])
    irr = set(info["expected"]["irr_rows"])
    want = {(sh, c, remap(sh)(r)) for sh, c, r in array_cells(SRC) if not (sh == "Dashboard" and r in irr)}
    got = set(array_cells(OUT))
    print(f"array formulas saved as arrays: {len(got)} (the user's E36:G36 only)",
          "OK" if got == want else f"MISMATCH {sorted(got ^ want)}")
    wb = openpyxl.load_workbook(OUT)
    names_left = [n for n in UNUSED_NAMES if n in wb.defined_names]
    print("unused names removed:", "OK" if not names_left else names_left)
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
    ok28, yearly = recompute28(a_wb, info)
    ok29, rtb = recompute29(a_wb, b_wb, info, yearly)
    print("scratch session: case flip and a broken dated flow")
    ok_s = scratch(info)
    print("   case flip clean and the tie check fires:", "OK" if ok_s else "FAIL")
    # Lint, new rows only.
    Q = info["Dashboard"]["rtb"]
    new = {"Inputs": set(info["Inputs"]["rows"].values()),
           "Dashboard": {R[k] for k in ("s_cit", "s_irr", "s_debt", "s_eirr", "s_dscr", "check", "master",
                                        "check28", "check_err", "irr", "cod_irr")}
           | {v for b in info["Dashboard"]["blocks"].values() for k, v in b.items()
              if k in [t[0] for t in TIMELINE] + ["val"] or k.startswith(("s_", "loan"))}
           | {v for x in Q.values() for k, v in x.items()
              if k in ("rest", "exit", "moic", "irr", "loan_cit", "s_moic", "s_irr_post", "tie", "pre_irr")}
           | {r for x in Q.values() for r in x["cit_years"]}}
    loan_heads = {b["loanB_cash"] - 1 for b in info["Dashboard"]["blocks"].values()}
    new["Dashboard"] |= loan_heads | {r for x in Q.values() for r in range(x["first"] + 1, x["rec"])}
    found = []
    for x in lint(OUT):
        sh, _, where = x.where.partition("!")
        hit = re.search(r"(\d+)$", where)
        if sh in new and hit and int(hit.group(1)) in new[sh]:
            found.append(str(x))
    print("lint (new and §29-changed rows):", "clean" if not found else f"{len(found)} finding(s)")
    for x in found[:20]:
        print("  ", x)
    dump(a_wb, info)
    print("dump:", DUMP)
    return ok and ok28 and ok29 and ok_s and not errs and not bad and not struct and got == want and same \
        and not filt and not names_left


if __name__ == "__main__":
    build()
    sys.exit(0 if verify() else 1)
