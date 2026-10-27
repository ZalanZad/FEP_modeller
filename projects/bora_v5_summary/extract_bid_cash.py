"""Standalone extract: the Bora bid's cash to RtB, all in EUR k, everything paid in cash (no bank guarantees).

Reads the current model values (out/Bora_model_RTB_COD_HTM_recalc.xlsx, Bora selected) into a blue input table and
builds one sheet "Bid_Cash" with formulas below it: securities and fees payable (with the deductions), devex payable
(Devex_FEP lines at the 200 MVA basis × scale), grand total, by-year summary and one check. The model is not changed.

    PYTHONPATH=. python3 projects/bora_v5_summary/extract_bid_cash.py
"""
import sys
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter as L

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))
from fepm.recalc import recalc  # noqa: E402
from fepm.xl.styles import apply, register  # noqa: E402

MODEL = HERE / "out" / "Bora_model_RTB_COD_HTM_recalc.xlsx"
OUT = HERE / "out" / "Bora_bid_cash_extract.xlsx"
BIDS = 2                                                          # BORA 1 and BORA 2 (Bid_Points)
DATE, KEUR = "dd-mmm-yy", "money1"


def model_inputs(wb):
    i, f = wb["Inputs"], wb["Devex_FEP"]
    v = lambda r: i[f"E{r}"].value  # noqa: E731
    mva = wb["Dashboard"]["E20"].value / v(82)                       # Dash_MVA ÷ the model's bids
    inputs = [  # key, label, unit, value, source, format
        ("mva", "Connection capacity per bid", "MVA", mva, "Dashboard E20 ÷ Inputs E82", "#,##0"),
        ("bids", "Number of bids", "#", BIDS, "Bid_Points: BORA 1 and BORA 2", "int"),
        ("ob", "Overbuild (ÖSZ-3)", "%", wb["Dashboard"]["E22"].value, "Dashboard E22", "pct"),
        ("bess", "BESS power, share of MVA (ÖSZ-4)", "%", wb["Dashboard"]["E23"].value, "Dashboard E23", "pct"),
        ("hours", "BESS duration", "h", v(21), "Inputs E21", "int"),
        ("pbm", "Performance security multiple (ÖSZ-6)", "x", wb["Dashboard"]["E27"].value, "Dashboard E27", "mult"),
        ("fx", "FX rate", "HUF/EUR", v(11), "Inputs E11", "#,##0"),
        ("bw", "Security base: benchmark investment, wind", "HUF/MW", v(75) * 1e6, "Inputs E75", "#,##0"),
        ("bb", "Security base: benchmark investment, BESS", "HUF/MWh", v(89) * 1e6, "Inputs E89", "#,##0"),
        ("bsr", "Bid security, share of the base", "%", v(76), "Inputs E76", "0.0%"),
        ("pbr", "Performance security, share of the base", "%", v(77), "Inputs E77", "0.0%"),
        ("pf", "Participation fee per bid", "HUF", v(81) * 1e6, "Inputs E81", "#,##0"),
        ("tso", "TSO connection-request security", "HUF/MVA", v(83) * 1e6, "Inputs E83", "#,##0"),
        ("ind", "Indirect connection fee", "HUF/MVA", v(297), "Inputs E297", "#,##0"),
        ("dir", "Direct connection fee per connection point", "HUF", v(298), "Inputs E298", "#,##0"),
        ("share", "Bids sharing the connection point", "#", BIDS, "ÖSZ-8 shared connection", "int"),
        ("cont", "Devex contingency (on devex excl. other costs + fee)", "%", v(68), "Inputs E68", "pct"),
        ("basis", "Devex file basis (BORA_Devex, 200 MVA project)", "MVA", 200, "Inputs E312", "#,##0"),
        ("days", "Performance security due after award", "days", v(102), "Inputs E102", "int"),
        ("k", "Thousand (HUF → k€ conversion)", "#", 1000, "Unit conversion", "#,##0"),
        ("tol", "Check tolerance", "k€", 1e-6, "House standard", "General"),
        ("d_sign", "Signing", "date", v(27), "Inputs E27", DATE),
        ("d_bid", "Bid submission", "date", v(95), "Inputs E95", DATE),
        ("d_award", "Award", "date", v(96), "Inputs E96", DATE),
        ("d_req", "Connection request (TSO security)", "date", v(97), "Inputs E97", DATE),
        ("d_hcssz", "Connection agreement (HCSSZ): connection fees", "date", v(99), "Inputs E99", DATE),
        ("d_rtb", "RtB", "date", v(56), "Inputs E56", DATE),
    ]
    devex = []
    for r in range(5, 38):
        task, item = f[f"C{r}"].value, f[f"D{r}"].value
        label = f"{task.strip()} – {item}" if task and task.strip() != item.strip() else item
        devex.append((label, f[f"F{r}"].value, f[f"L{r}"].value, r == 36))   # row 36: other costs
    return inputs, devex, (f["F38"].value, f["E39"].value), wb["Deposits"]


def build():
    src = openpyxl.load_workbook(MODEL, data_only=True)
    inputs, devex, (cont_date, model_devex), dep = model_inputs(src)
    wb = openpyxl.Workbook()
    register(wb)
    ws = wb.active
    ws.title = "Bid_Cash"
    ws.sheet_view.showGridLines = False
    for col, width in zip("ABCDEFG", (2, 58, 11, 13, 13, 46, 13)):
        ws.column_dimensions[col].width = width
    R = {}

    def bar(r, text, heads=None):
        for c in range(1, 8):
            apply(ws.cell(r, c), "section")
        ws.cell(r, 2).value = text
        for col, h in (heads or {}).items():
            ws[f"{col}{r}"].value = h

    apply(ws["B1"], "banner")
    for c in range(3, 8):
        apply(ws.cell(1, c), "banner")
    ws["B1"].value = "Bora bid – cash to RtB (EUR k; bid security and performance security as cash deposits)"

    # 1. Inputs (blue) and derived values (black).
    bar(3, "1. Inputs: current Bora bid and tender values (model 27-Oct-2026)",
        {"C": "Unit", "D": "Value", "F": "Source"})
    r = 4
    for key, label, unit, value, source, fmt in inputs:
        R[key] = r
        apply(ws[f"B{r}"], "label")
        apply(ws[f"C{r}"], "unit")
        apply(ws[f"D{r}"], "input", fmt)
        apply(ws[f"F{r}"], "unit")
        ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"], ws[f"F{r}"] = label, unit, value, source
        r += 1
    D = {k: f"$D${v}" for k, v in R.items()}
    derived = [
        ("d_pb", "Performance security posted (award + days)", "date", f"={D['d_award']}+{D['days']}", DATE),
        ("wind", "Wind committed per bid", "MW", f"={D['mva']}*(1+{D['ob']})", "#,##0.0"),
        ("mwh", "BESS energy per bid", "MWh", f"={D['mva']}*{D['bess']}*{D['hours']}", "#,##0.0"),
        ("base", "Security base per bid (wind MW × benchmark + MWh × benchmark)", "k€",
         f"=({{wind}}*{D['bw']}+{{mwh}}*{D['bb']})/{D['fx']}/{D['k']}", KEUR),
        ("pbg", "Performance security per bid, gross", "k€", f"={D['pbr']}*{D['pbm']}*{{base}}", KEUR),
        ("scale", "Devex scale (MVA × bids ÷ devex file basis)", "x", f"={D['mva']}*{D['bids']}/{D['basis']}",
         "0.000"),
    ]
    r += 1
    bar(r, "Derived (formulas)")
    r += 1
    for key, label, unit, f, fmt in derived:
        R[key] = r
        apply(ws[f"B{r}"], "label")
        apply(ws[f"C{r}"], "unit")
        apply(ws[f"D{r}"], "calc", fmt)
        ws[f"B{r}"], ws[f"C{r}"] = label, unit
        ws[f"D{r}"] = f.format(**{k: f"$D${v}" for k, v in R.items()})
        r += 1
    D = {k: f"$D${v}" for k, v in R.items()}

    # 2. Securities and fees payable, cash.
    r += 1
    bar(r, "2. Securities and fees payable, cash", {"C": "Date", "D": "Per bid k€", "E": "All bids k€",
                                                   "F": "Deduction"})
    first = r + 1
    sec = [  # key, label, date, per-bid formula, deduction note
        ("part", "Participation fee", D["d_bid"], f"={D['pf']}/{D['fx']}/{D['k']}", "sunk"),
        ("bs", "Bid security paid at submission", D["d_bid"], f"={D['bsr']}*{D['base']}", None),
        ("bsc", "Bid security credited into the performance security", D["d_pb"], "=-D{bs}",
         "stays posted as part of the PB"),
        ("pb", "Performance security posted at award, gross", D["d_pb"], f"={D['pbg']}",
         "top-up paid = gross − bid security credited"),
        ("tsop", "TSO connection-request security paid", D["d_req"], f"={D['tso']}*{D['mva']}/{D['fx']}/{D['k']}",
         None),
        ("tsoc", "TSO security credited against the indirect fee", D["d_hcssz"], "=-D{tsop}",
         "indirect fee payable = gross − TSO security"),
        ("indg", "Indirect connection fee, gross", D["d_hcssz"], f"={D['ind']}*{D['mva']}/{D['fx']}/{D['k']}", None),
        ("dirp", "Direct connection fee (per point ÷ sharing bids)", D["d_hcssz"],
         f"={D['dir']}/{D['share']}/{D['fx']}/{D['k']}", "shared connection point"),
    ]
    r = first
    for key, label, date, f, note in sec:
        R[key] = r
        apply(ws[f"B{r}"], "label")
        apply(ws[f"C{r}"], "calc", DATE)
        for col in "DE":
            apply(ws[f"{col}{r}"], "calc", KEUR)
        apply(ws[f"F{r}"], "unit")
        ws[f"B{r}"], ws[f"C{r}"] = label, f"={date}"
        ws[f"D{r}"] = f.format(**R)
        ws[f"E{r}"] = f"=D{r}*{D['bids']}"
        ws[f"F{r}"] = note
        r += 1
    R["sec_sub"] = r
    apply(ws[f"B{r}"], "total")
    for col in "DE":
        apply(ws[f"{col}{r}"], "total", KEUR)
        ws[f"{col}{r}"] = f"=SUM({col}{first}:{col}{r - 1})"
    ws[f"B{r}"] = "Securities and fees, cash"

    # 3. Devex payable (Devex_FEP lines at the 200 MVA basis × scale).
    r += 2
    bar(r, "3. Devex payable (dates and amounts as in section 2)", {"G": "200 MVA k€"})
    dfirst = r + 1
    r = dfirst
    for label, date, amount, other in devex:
        if other:
            R["other"] = r
        apply(ws[f"B{r}"], "label")
        apply(ws[f"C{r}"], "input", DATE)
        apply(ws[f"G{r}"], "input", KEUR)
        for col in "DE":
            apply(ws[f"{col}{r}"], "calc", KEUR)
        ws[f"B{r}"], ws[f"C{r}"], ws[f"G{r}"] = label, date, amount
        ws[f"E{r}"] = f"=G{r}*{D['scale']}"
        ws[f"D{r}"] = f"=E{r}/{D['bids']}"
        r += 1
    R["cont"] = r
    apply(ws[f"B{r}"], "label")
    apply(ws[f"C{r}"], "input", DATE)
    for col in "DE":
        apply(ws[f"{col}{r}"], "calc", KEUR)
    apply(ws[f"F{r}"], "unit")
    ws[f"B{r}"], ws[f"C{r}"] = "Contingency (devex excl. other costs + participation fee)", cont_date
    ws[f"E{r}"] = f"={D['cont']}*(SUM(E{dfirst}:E{r - 1})-E{R['other']}+E{R['part']})"
    ws[f"D{r}"] = f"=E{r}/{D['bids']}"
    ws[f"F{r}"] = "other costs outside the base (as the model)"
    r += 1
    R["dev_sub"] = r
    apply(ws[f"B{r}"], "total")
    for col in "DE":
        apply(ws[f"{col}{r}"], "total", KEUR)
        ws[f"{col}{r}"] = f"=SUM({col}{dfirst}:{col}{r - 1})"
    ws[f"B{r}"] = "Devex, land and contingency"
    last = r

    # 4. Grand total, by year, check.
    r += 2
    R["grand"] = r
    apply(ws[f"B{r}"], "total")
    for col in "DE":
        apply(ws[f"{col}{r}"], "total", KEUR)
        ws[f"{col}{r}"] = f"={col}{R['sec_sub']}+{col}{R['dev_sub']}"
    ws[f"B{r}"] = "Grand total cash to RtB"
    r += 2
    bar(r, "4. Cash by year", {"C": "Year", "D": "Per bid k€", "E": "All bids k€", "F": "Cumulative, all bids k€"})
    yfirst = r + 1
    for i in range(4):
        r += 1
        apply(ws[f"B{r}"], "label")
        apply(ws[f"C{r}"], "calc", "0")
        for col in "DEF":
            apply(ws[f"{col}{r}"], "calc", KEUR)
        ws[f"B{r}"] = "Cash paid in the year"
        ws[f"C{r}"] = f"=YEAR({D['d_sign']})" if i == 0 else f"=C{r - 1}+1"
        for col in "DE":
            ws[f"{col}{r}"] = f"=SUMPRODUCT((YEAR($C${first}:$C${last})=$C{r})*{col}${first}:{col}${last})"
        ws[f"F{r}"] = f"=E{r}" if i == 0 else f"=F{r - 1}+E{r}"
    r += 1
    apply(ws[f"B{r}"], "total")
    for col in "DE":
        apply(ws[f"{col}{r}"], "total", KEUR)
        ws[f"{col}{r}"] = f"=SUM({col}{yfirst}:{col}{r - 1})"
    ws[f"B{r}"] = "Total by year"
    r += 1
    apply(ws[f"B{r}"], "label")
    apply(ws[f"D{r}"], "calc")
    ws[f"B{r}"] = "Check: by-year sum = grand total (all bids)"
    ws[f"D{r}"] = f'=IF(ABS(E{r - 1}-E{R["grand"]})<{D["tol"]},"OK","CHECK")'
    ws.freeze_panes = "C4"
    wb.save(OUT)

    # Verification against the model (cash reading of the Deposits k€ values, Devex_FEP at the same scale).
    rc = openpyxl.load_workbook(recalc(OUT, OUT.with_name("Bora_bid_cash_extract_recalc.xlsx")), data_only=True)
    x = rc["Bid_Cash"]
    val = lambda key, col="D": x[f"{col}{R[key]}"].value  # noqa: E731
    model = {"bid security": (val("bs"), dep["M8"].value),
             "PB gross": (val("pb"), dep["M8"].value + dep["M10"].value),
             "PB top-up": (val("pb") + val("bsc"), dep["M10"].value),
             "TSO security": (val("tsop"), dep["M11"].value),
             "indirect net": (val("indg") + val("tsoc"), dep["M18"].value),
             "direct": (val("dirp"), dep["M12"].value),
             "devex per bid (100 MVA)": (val("dev_sub"), model_devex)}
    ok = all(abs(a - b) < 1e-6 for a, b in model.values())
    for k, (a, b) in model.items():
        print(f"{k:<26} extract {a:>12,.1f}  model {b:>12,.1f}")
    years = [(x[f"C{r}"].value, x[f"D{r}"].value, x[f"E{r}"].value, x[f"F{r}"].value) for r in range(yfirst, yfirst + 4)]
    print("grand total per bid / all bids:", round(val("grand"), 1), round(val("grand", "E"), 1))
    print("by year (per bid, all bids, cumulative):", [(y, round(a, 1), round(b, 1), round(c, 1)) for y, a, b, c in years])
    print("check:", x[f"D{yfirst + 5}"].value, "| ties to model:", "PASS" if ok else "FAIL", "| rows:", x.max_row)
    return OUT


if __name__ == "__main__":
    build()
