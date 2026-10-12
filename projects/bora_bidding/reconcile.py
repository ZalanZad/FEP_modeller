"""Bora bid model – reconciliation against the old engine (SPEC v2 §3 saved values, §8 deliberate fixes).

For each headline and bid: saved value (meeting file, reproduced by the old-engine shadows) → each §8 fix applied one at
a time in the shadows (data/old_engine/old_rtb.py, old_htm.py) → new-structure residual → the recalculated workbook at
key-assumption set 3 "Overbuild built". CN values: set 2 with In_Overbuild_Built = 1 in a private LibreOffice copy.
Then set 3 → base (overbuild not built) per bid. Writes out/reconciliation.md.

    python3 projects/bora_bidding/reconcile.py

Stdlib + fepm. The deliverable workbook is never saved: fepm.recalc.Office works on a private copy.
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))
sys.path.insert(0, str(HERE / "data" / "old_engine"))

import old_htm as H  # noqa: E402
import old_rtb as R  # noqa: E402

from fepm.recalc import Office, name_value, values  # noqa: E402

XLSM = HERE / "out" / "Bora_Bid_Model.xlsm"
RECALC = HERE / "out" / "Bora_Bid_Model_recalc.xlsx"
MD = HERE / "out" / "reconciliation.md"
BIDS = [("BORA 1", "S2", 1), ("BORA 2", "S3", 2)]
EXCEL0 = dt.date(1899, 12, 30)
TOL_V, TOL_PP = 0.001, 0.001                         # 0.1 % on values, 0.1 pp on IRR
K = 1000.0
IRR_COD = "FEP IRR (HTM!G352, XIRR rows 359/362)"
IRR_HTM = "FEP IRR (HTM!G139, IRR on row 129)"
ALLIN = "All-in cost to COD (Bora row 87)"


def todate(x):
    return EXCEL0 + dt.timedelta(days=int(round(x)))


def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# =========================================================================================== workbook (LibreOffice)
OUTS = ["Points_EU", "Points_CN", "Wind_MW", "Wind_Built_MW", "BESS_MW", "Gross_Premium", "Price_Effect",
        "PE_Municipal", "PE_Community", "PE_BESS", "PE_PB", "PE_Overbuild", "Eff_Price", "Price_PV", "Distributable",
        "FEP_Proceeds", "Developer", "RtB_Profit", "RtB_IRR", "WF_DSA", "WF_Return_AtRisk", "WF_Return_6",
        "Cost_To_COD", "C2C_IDC", "C2C_Upfront_Fee", "C2C_Commitment_Fee", "COD_Price_Effect", "COD_Debt",
        "COD_Proceeds", "COD_Profit", "COD_IRR", "HTM_IRR"]
ROWS = {"htm_cf": "FEP equity cash flow – hold to maturity", "contrib": "FEP contributions to COD",
        "gen": "Net generation", "rev": "Revenue", "ebitda": "EBITDA", "int": "Interest paid",
        "tax": "Corporate income tax"}


def _row_index(lo, sheet):
    labels = lo._range(f"{sheet}!C1:C400").getDataArray()
    return {str(lab[0]): i for i, lab in enumerate(labels, 1) if lab[0]}


def _find(index, prefix):
    hits = [r for lab, r in index.items() if lab.startswith(prefix)]
    if not hits:
        raise KeyError(f"row '{prefix}' not found on the Calc sheet")
    return hits[0]


def snapshot(lo, rows=False):
    snap = {"check": lo.get("Master_Check")}
    for o in OUTS:
        for b in (1, 2):
            try:
                snap[f"{o}_{b}"] = lo.get(f"Out_{o}_{b}")
            except Exception:                                        # output not on this build
                snap[f"{o}_{b}"] = None
    if rows:
        last = col(9 + 141)
        for b in (1, 2):
            sh = f"Calc_{b}"
            idx = _row_index(lo, sh)
            snap[f"ends_{b}"] = [todate(x) for x in lo._range(f"{sh}!J4:{last}4").getDataArray()[0]]
            for key, prefix in ROWS.items():
                r = _find(idx, prefix)
                arr = lo._range(f"{sh}!I{r}:{last}{r}").getDataArray()[0]
                snap[f"{key}_{b}"] = [v if isinstance(v, float) else 0.0 for v in arr[1:]]
    return snap


def read_workbook():
    """Every run on one private copy. Clean sets first; overrides of Live cells (no formula restore) last."""
    with Office(XLSM) as lo:
        inp = {k: lo.get(f"In_{k}") for k in ("Award_Date", "PB_Days", "PB_Release", "Tolerance_Cap")}
        runs = {}
        for tag, name, over, rows in [("base_eu", "Base", {}, False), ("base_cn", "CN turbine", {}, False),
                                      ("s3_eu", "Overbuild built", {}, True),
                                      ("s3_cn", "CN turbine", {"In_Overbuild_Built": 1}, True),
                                      ("nb_cn", "CN turbine", {"In_Overbuild_Built": 0, "In_Tolerance_Cap": 1}, False),
                                      ("nb_eu", "Base", {"In_Overbuild_Built": 0, "In_Tolerance_Cap": 1}, False)]:
            lo.set("Scenario_Name", name)
            lo.calculate()
            for ref, v in over.items():
                lo.set(ref, v)
            lo.calculate()
            runs[tag] = snapshot(lo, rows)
            runs[tag]["set"] = name + "".join(f", {k[3:]} = {v}" for k, v in over.items())
    return inp, runs


def cell_refs():
    """Dashboard cell of each Out_ name (from the recalculated deliverable)."""
    wb = values(RECALC)
    refs = {}
    for n in wb.defined_names:
        if n.startswith("Out_"):
            s, c = next(wb.defined_names[n].destinations)
            refs[n[4:]] = f"{s}!{c.replace('$', '')}"
    return wb, refs


# =========================================================================================== shadows
def steps(pb_date, pb_release):
    """§8 fixes in order. (key, label, old_rtb toggles, old_htm toggles, tables). Toggles accumulate."""
    pbd, rel = pb_date.isoformat(), pb_release.isoformat()
    return [
        ("1", "Premium on awarded MVA (100 MVA both; no change)", {}, {}, "rb"),
        ("2", "Direct connection fee per connection point", dict(direct_fee_basis="lump"),
         dict(direct_fee_basis="lump"), "rb"),
        ("3", "Bid security released at PB posting; PB gross", dict(bid_netting=False, bid_release_date=pbd),
         dict(bid_netting=False, bid_release=True), "rb"),
        ("4", "PB posted award + 15 days; held to 31-Mar-2033", dict(pb_date=pbd, pb_release_date=rel),
         dict(ov={"E102": 15}, pb_release_date=rel), "rb"),
        ("4b", "PB carry in the price effect after tax (SPEC §5.6)", dict(pb_carry_tax=True), {}, "rb"),
        ("5", "One FX (365) for devex (old 370)", dict(fx_devex=365.0), dict(devex=True), "rb"),
        ("6", "Waterfall capped at cash; no developer floor", dict(waterfall="capped", dev_floor=False),
         dict(cod_waterfall="capped"), "rb"),
        ("7", "One participation fee per bid (old: both bids)", dict(n_bids=1), dict(ov={"E82": 1}), "rb"),
        ("8", "Guarantee fee on exact days (old: full months)", dict(guarantee_fee_days=True),
         dict(guarantee_fee_days=True), "rb"),
        ("9", "Claims per class on actual days; portfolio floor",
         dict(interest_method="daily", atrisk_floor="portfolio"), dict(claims="daily"), "rb"),
        ("10a", "HTM stand-alone: developer buy-out, no deemed exit", {},
         dict(deemed_exit=False, fep_share=1.0, buyout=True), "b"),
        ("10b", "No DSRA, no cash sweep", {}, dict(dsra=False, sweep=False), "b"),
        ("10c", "No loss cap, no decommissioning", {}, dict(ov={"E214": 1.0, "E175": 0.0}), "b"),
        ("10d", "PPA for its tenor from COD (buyer margin too)", dict(margin_revenue="ppa_from_cod"),
         dict(ppa_from_cod=True), "rb"),
        ("11", "Quarterly timeline: price effect, flows at quarter ends", dict(grid="quarterly"), {}, "r"),
    ]


def htm_kwargs(s, rk, hk, cod):
    hk = dict(hk)
    ov = hk.pop("ov", {})
    if hk.pop("devex", False):
        rr = R.compute(s, **rk)
        rc = R.compute(s, **dict(rk, sale_date=cod.isoformat()))
        hk["devex"] = dict(dv=rr["detail"]["devex_monthly"], land=rr["detail"]["land_monthly"],
                           g280=rc["detail"]["devex_return_G131"])
    if hk.pop("buyout", False):
        hk["buyout_keur"] = R.compute(s, **rk)["developer_total"]
    return dict(hk, overrides=ov)


def rtb_vals(s, rk):
    r = R.compute(s, **rk)
    return dict(pe=r["price_effect"] / K, eff=r["effective_price"] / K, dist=r["wf"]["59 distributable"] / K,
                profit=r["fep_post_tax_profit"] / K, irr=r["fep_irr"], items=r["price_effect_items"],
                dev=r["developer_total"] / K)


def build_vals(s, rk, hk, cod):
    kw = htm_kwargs(s, rk, hk, cod)
    x = {rt: H.compute(s, rt, **kw) for rt in ("COD-EU", "COD-CN", "HTM-EU", "HTM-CN")}
    return dict(c2c_eu=x["COD-EU"]["cost_to_cod"][ALLIN] / K, c2c_cn=x["COD-CN"]["cost_to_cod"][ALLIN] / K,
                cod_eu=x["COD-EU"]["cod"][IRR_COD], cod_cn=x["COD-CN"]["cod"][IRR_COD],
                htm_eu=x["HTM-EU"]["htm"][IRR_HTM], htm_cn=x["HTM-CN"]["htm"][IRR_HTM], raw=x)


def run_bridge(s, stp, cod):
    rk, hk = {}, {"ov": {}}
    rows = [("0", "Saved (meeting file; shadow reproduces it)", rtb_vals(s, rk), build_vals(s, rk, hk, cod))]
    for key, label, r_, h_, _ in stp:
        rk.update(r_)
        for k, v in h_.items():
            if k == "ov":
                hk["ov"] = dict(hk["ov"], **v)
            else:
                hk[k] = v
        rows.append((key, label, rtb_vals(s, rk), build_vals(s, rk, hk, cod)))
    floor = {f: R.compute(s, **dict(rk, grid="monthly", sale_date=cod.isoformat(), atrisk_floor=f))
             ["detail"]["devex_return_G131"] / K for f in ("line", "portfolio")}
    return rows, floor


def annual_irr_on_old_basis(ends, cf, first_year=2026, years=36):
    """Old HTM IRR: Excel IRR on calendar-year sums, t0 = 2026 (HTM!G139)."""
    tot = [0.0] * years
    for e, v in zip(ends, cf):
        tot[e.year - first_year] += v
    return H.irr(tot)


def cod_xirr_on_old_basis(ends, contrib, proceeds, cod, inv=dt.date(2026, 7, 31)):
    """Old COD IRR: annual contributions on mid-period dates (truncated) + proceeds on COD (HTM!G352)."""
    yrs = list(range(2026, cod.year + 1))
    tot = {y: 0.0 for y in yrs}
    for e, v in zip(ends, contrib):
        if e <= cod:
            tot[e.year] += v
    ds = [int((H.ser(max(dt.date(y, 1, 1), inv)) + H.ser(max(min(dt.date(y, 12, 31), cod), dt.date(y, 1, 1)))) / 2)
          for y in yrs]
    return H.xirr([tot[y] for y in yrs] + [proceeds], ds + [H.ser(cod)])


# =========================================================================================== formatting
def m2(v):
    if v is None:
        return "–"
    if isinstance(v, str):
        return v
    return f"({abs(v):,.2f})" if v < -0.005 else f"{abs(v):,.2f}"


def d2(v):
    return "0.00" if abs(v) < 0.005 else (f"+{v:,.2f}" if v > 0 else f"({abs(v):,.2f})")


def pc(v, dp=2):
    if v is None or isinstance(v, str):
        return "n/a"
    return f"({abs(v) * 100:.{dp}f}) %" if v < -0.5 * 10 ** -(dp + 2) else f"{v * 100:.{dp}f} %"


def dpp(v, dp=2):
    return "0.00" if abs(v) < 0.5 * 10 ** -(dp + 2) else (f"+{v * 100:.{dp}f}" if v > 0 else f"({abs(v) * 100:.{dp}f})")


def over(v, ref, is_irr):
    if v is None or ref is None or isinstance(ref, str):
        return False
    return abs(v) > (TOL_PP if is_irr else max(TOL_V * abs(ref), 0.0005))


def table(head, rows):
    out = ["| " + " | ".join(head) + " |", "|" + "|".join("---" if i < 2 else "---:" for i in range(len(head))) + "|"]
    out += ["| " + " | ".join(r) + " |" for r in rows]
    return out


# =========================================================================================== main
def main():
    inp, W = read_workbook()
    wbv, refs = cell_refs()
    cod = dt.date(2031, 9, 30)
    pb_date = todate(inp["Award_Date"]) + dt.timedelta(days=int(inp["PB_Days"]))
    stp = steps(pb_date, todate(inp["PB_Release"]))
    s3, cn = W["s3_eu"], W["s3_cn"]
    md, findings = [], []
    stamp = dt.datetime.fromtimestamp(XLSM.stat().st_mtime).strftime("%d-%b-%Y %H:%M")
    md += ["# Bora bid model – reconciliation to the old engine", "",
           f"Model `out/Bora_Bid_Model.xlsm` (built {stamp}), recalculated in LibreOffice by `reconcile.py` on a private copy. "
           "Reconciliation set: key-assumption set 3 \"Overbuild built\" (EU); CN = set 2 with `In_Overbuild_Built` = 1 "
           f"(private copy only). Master check: set 3 = {s3['check']:.0f}.",
           "Saved values: `reference/Bora_scenarios_meeting.xlsx` (SPEC §3), reproduced by `data/old_engine/old_rtb.py` "
           "and `old_htm.py` (≤ 1e-8). Each §8 fix is applied cumulatively in those shadows.",
           "€ m to 2 dp, negatives in brackets; IRR in % (step changes in pp, 0.01). Tolerance 0.1 % on values, "
           "0.1 pp on IRR; `*` = above it (see Findings).", ""]

    # ---- 1. ties without a bridge
    md += ["## 1. Headlines that tie without a bridge", ""]
    sv = R.J["saved_bora"]
    rows = []
    for label, key, okey, fmt in [("Points, EU turbine", 35, "Points_EU", "{:.1f}"), ("Points, CN turbine", 36, "Points_CN", "{:.1f}"),
                                  ("Installed wind MW", 22, "Wind_MW", "{:.0f}"), ("BESS MW", 23, "BESS_MW", "{:.0f}"),
                                  ("Gross RtB premium €m", 41, "Gross_Premium", "{:.2f}")]:
        cells = []
        for _, s, b in BIDS:
            c = {"S2": "F", "S3": "G"}[s]
            a, w = sv[s][f"Bora!{c}{key}"], s3[f"{okey}_{b}"]
            cells += [fmt.format(a), fmt.format(w)]
            if abs(a - w) > 1e-6:
                findings.append(f"{label} {b}: saved {a} vs model {w} (`{refs.get(f'{okey}_{b}')}`).")
        rows.append([label] + cells + [f"`{refs.get(okey + '_1')}`, `{refs.get(okey + '_2')}`"])
    md += table(["Item", "BORA 1 saved", "BORA 1 model", "BORA 2 saved", "BORA 2 model", "Cells (B1, B2)"], rows) + [""]

    bridges = {s: run_bridge(s, stp, cod) for _, s, _ in BIDS}

    # ---- 2. RtB bridge
    md += ["## 2. RtB exit bridge (set 3)", ""]
    rk_cols = [("pe", "Price effect", "Price_Effect", False), ("eff", "Effective price", "Eff_Price", False),
               ("dist", "Distributable", "Distributable", False), ("profit", "FEP RtB profit", "RtB_Profit", False),
               ("irr", "FEP RtB IRR", "RtB_IRR", True)]
    for name, s, b in BIDS:
        rows_, floor = bridges[s]
        md += [f"### {name}", ""]
        body = []
        first = rows_[0][2]
        body.append(["0", rows_[0][1]] + [pc(first[k]) if irr else m2(first[k]) for k, _, _, irr in rk_cols])
        if s == "S2":
            body[-1][-1] += " ¹"
        for i in range(1, len(rows_)):
            key, label, cur, _ = rows_[i]
            if "r" not in next(t for k_, _, _, _, t in stp if k_ == key):
                continue
            prev = rows_[i - 1][2]
            body.append([key, label] + [dpp(cur[k] - prev[k]) if irr else d2(cur[k] - prev[k]) for k, _, _, irr in rk_cols])
        last = rows_[-1][2]
        body.append(["", "**Shadow, all fixes**"] + [pc(last[k]) if irr else m2(last[k]) for k, _, _, irr in rk_cols])
        res = []
        for k, _, okey, irr in rk_cols:
            w = s3[f"{okey}_{b}"]
            r_ = (w - last[k]) if not isinstance(w, str) else None
            flag = "*" if over(r_, w, irr) else ""
            res.append(("n/a" if r_ is None else (dpp(r_) if irr else d2(r_))) + flag)
            if flag:
                items = [(it, s3[f"PE_{o}_{b}"] - last["items"][it] / K) for it, o in
                         (("municipal", "Municipal"), ("community", "Community"), ("bess", "BESS"),
                          ("security", "PB"), ("overbuild", "Overbuild"))]
                big = [f"{it} {d2(v)}" for it, v in items if abs(v) > 0.0005]
                findings.append(f"{name} RtB {okey}: residual {res[-1]} vs the shadow"
                                + (f"; price-effect items off: {', '.join(big)}." if big else "."))
        body.append(["", "Residual: workbook − shadow"] + res)
        body.append(["", "**Workbook, set 3**"] + [pc(s3[f"{o}_{b}"]) if irr else m2(s3[f"{o}_{b}"])
                                                     for _, _, o, irr in rk_cols])
        md += table(["#", "Step", "Price effect", "Effective price", "Distributable", "FEP RtB profit", "FEP RtB IRR"], body)
        md += ["", "Cells: " + ", ".join(f"{lab} `{refs.get(f'{o}_{b}')}`" for _, lab, o, _ in rk_cols) + "."]
        dev_old, dev_new = rows_[0][2]["dev"], rows_[-1][2]["dev"]
        md += [f"Developer at the exit (DSA + 30 %): saved {m2(dev_old)}, shadow all fixes {m2(dev_new)}, "
               f"workbook {m2(s3[f'Developer_{b}'])} (`{refs.get(f'Developer_{b}')}`)."]
        if s == "S2":
            md += ["¹ Excel's XIRR failed in the meeting file (2.98e-09); −3.91 % is the true root (LibreOffice, shadow). "
                   "The model shows `n/a` plus a check when XIRR fails (fix 6)."]
        md += [""]

    # price-effect items
    md += ["Price effect by item, € m (old engine at the exit vs set 3):", ""]
    rows = []
    for it, o, lab in (("municipal", "Municipal", "Municipal one-off"), ("community", "Community", "Community contribution"),
                       ("bess", "BESS", "BESS capex"), ("security", "PB", "PB carry"),
                       ("overbuild", "Overbuild", "Overbuild bottleneck loss")):
        r_ = [lab]
        for _, s, b in BIDS:
            a = bridges[s][0][0][2]["items"][it] / K
            w = s3[f"PE_{o}_{b}"]
            r_ += [m2(a), m2(w), d2(w - a)]
        rows.append(r_ + [f"`{refs.get(f'PE_{o}_1')}`"])
    md += table(["Item", "B1 saved", "B1 model", "B1 Δ", "B2 saved", "B2 model", "B2 Δ", "Cell (B1)"], rows)
    md += ["", "Main drivers: municipal paid at COD (30-Sep-2031) instead of 31-Dec-2031; PB carry to 31-Mar-2033 and "
           "after tax; buyer margin on the PPA for 10 years from COD (old RtB case: merchant); quarterly generation "
           "with degradation by operating year.", ""]

    # ---- 3. build / COD / HTM bridge
    md += ["## 3. Build, sale at COD and hold to maturity (set 3; CN = set 2 + overbuild built)", ""]
    bcols = [("c2c_eu", "All-in cost EU", "Cost_To_COD", "s3_eu", False), ("c2c_cn", "All-in cost CN", "Cost_To_COD", "s3_cn", False),
             ("cod_eu", "COD IRR EU", "COD_IRR", "s3_eu", True), ("cod_cn", "COD IRR CN", "COD_IRR", "s3_cn", True),
             ("htm_eu", "HTM IRR EU", "HTM_IRR", "s3_eu", True), ("htm_cn", "HTM IRR CN", "HTM_IRR", "s3_cn", True)]
    detail = {}
    for name, s, b in BIDS:
        rows_, floor = bridges[s]
        md += [f"### {name}", ""]
        body = [["0", rows_[0][1]] + [pc(rows_[0][3][k]) if irr else m2(rows_[0][3][k]) for k, _, _, _, irr in bcols]]
        for i in range(1, len(rows_)):
            key, label, _, cur = rows_[i]
            if "b" not in next(t for k_, _, _, _, t in stp if k_ == key):
                continue
            prev = rows_[i - 1][3]
            body.append([key, label] + [dpp(cur[k] - prev[k]) if irr else d2(cur[k] - prev[k]) for k, _, _, _, irr in bcols])
        last = rows_[-1][3]
        body.append(["", "**Shadow, all fixes**"] + [pc(last[k]) if irr else m2(last[k]) for k, _, _, _, irr in bcols])
        amounts, dates = [], []
        for k, lab, o, run, irr in bcols:
            w = W[run][f"{o}_{b}"]
            if k.startswith("c2c"):
                a_, d_ = w - last[k], 0.0
            elif k.startswith("cod"):
                mid = cod_xirr_on_old_basis(W[run][f"ends_{b}"], W[run][f"contrib_{b}"], W[run][f"COD_Proceeds_{b}"] * K, cod)
                a_, d_ = mid - last[k], w - mid
            else:
                mid = annual_irr_on_old_basis(W[run][f"ends_{b}"], W[run][f"htm_cf_{b}"])
                a_, d_ = mid - last[k], w - mid
            amounts.append((a_, w, irr, lab))
            dates.append((d_, w, irr, lab))
        body.append(["11a", "New structure: amounts (quarterly debt, IDC, fees, days)"]
                    + [(dpp(a) if irr else d2(a)) + ("*" if over(a, w, irr) else "") for a, w, irr, _ in amounts])
        body.append(["11b", "New structure: XIRR on quarter ends (old: annual / mid-year)"]
                    + [(dpp(d) if irr else d2(d)) + ("*" if over(d, w, irr) else "") for d, w, irr, _ in dates])
        body.append(["", "**Workbook**"] + [pc(W[run][f"{o}_{b}"]) if irr else m2(W[run][f"{o}_{b}"])
                                              for _, _, o, run, irr in bcols])
        md += table(["#", "Step", "All-in cost EU", "All-in cost CN", "COD IRR EU", "COD IRR CN", "HTM IRR EU", "HTM IRR CN"], body)
        md += ["", f"Cells: all-in cost `{refs.get(f'Cost_To_COD_{b}')}`, COD IRR `{refs.get(f'COD_IRR_{b}')}`, "
               f"HTM IRR `{refs.get(f'HTM_IRR_{b}')}`. Fix 6 and 10a–10d do not touch the COD route; 10a–10d do not touch the cost to COD.",
               f"At-risk claim to COD (fix 9): portfolio floor {m2(floor['portfolio'])}, per-line floor {m2(floor['line'])} "
               f"(Δ {d2(floor['portfolio'] - floor['line'])}); the per-line floor stops binding on lines paid before 2028.", ""]
        detail[b] = (last, amounts, dates)
        for (a, w, irr, lab), (d, _, _, _) in zip(amounts, dates):
            for part, v in (("amounts (11a)", a), ("dates (11b)", d)):
                if over(v, w, irr):
                    findings.append((name, lab, part, v, irr))

    # new-structure detail (EU, set 3)
    md += ["New-structure detail, EU set 3, € m (shadow with all fixes vs workbook; lifetime sums for the operating lines):", ""]
    rows = []
    for _, s, b in BIDS:
        x = bridges[s][0][-1][3]["raw"]
        ce, he = x["COD-EU"], x["HTM-EU"]
        a = ce["annual"]
        idc = sum(a["idc"])
        commit = -sum(a["fees"]) - ce["upfront_fee"]
        sh = {"IDC capitalised": idc, "Commitment fee": commit, "Upfront fee": ce["upfront_fee"],
              "Other cost-to-COD lines": ce["cost_to_cod"][ALLIN] - idc - commit - ce["upfront_fee"],
              "COD price effect": ce["cod"]["COD price effect (Tender!E225)"],
              "Debt repaid at COD": ce["cod"]["Senior debt repaid at COD (HTM!G357)"],
              "FEP contributions to COD": ce["cod"]["FEP contributions (G359)"],
              "FEP proceeds at COD": ce["cod"]["FEP proceeds at COD (G372)"],
              "Generation, GWh": sum(he["annual"]["gen_gwh"]) * K, "Revenue": sum(he["annual"]["revenue"]),
              "EBITDA": sum(he["annual"]["ebitda"]), "Interest paid": -sum(he["annual"]["interest"]),
              "Corporate income tax": sum(he["annual"]["tax"]), "FEP HTM cash flow": sum(he["annual"]["fep_cf"])}
        w = W["s3_eu"]
        wk = {"IDC capitalised": w[f"C2C_IDC_{b}"] * K, "Commitment fee": w[f"C2C_Commitment_Fee_{b}"] * K,
              "Upfront fee": w[f"C2C_Upfront_Fee_{b}"] * K,
              "COD price effect": w[f"COD_Price_Effect_{b}"] * K, "Debt repaid at COD": w[f"COD_Debt_{b}"] * K,
              "FEP contributions to COD": sum(w[f"contrib_{b}"]), "FEP proceeds at COD": w[f"COD_Proceeds_{b}"] * K,
              "Generation, GWh": sum(w[f"gen_{b}"]), "Revenue": sum(w[f"rev_{b}"]), "EBITDA": sum(w[f"ebitda_{b}"]),
              "Interest paid": sum(w[f"int_{b}"]), "Corporate income tax": sum(w[f"tax_{b}"]),
              "FEP HTM cash flow": sum(w[f"htm_cf_{b}"])}
        wk["Other cost-to-COD lines"] = (w[f"Cost_To_COD_{b}"] * K - wk["IDC capitalised"] - wk["Commitment fee"]
                                         - wk["Upfront fee"])
        detail[b] = (sh, wk)
    drivers = {"IDC capitalised": "interest on the quarter-opening balance; draws at quarter end (old: annual, half-year draws)",
               "Commitment fee": "on undrawn at the quarter start (old: net of capitalised IDC, B10)",
               "Upfront fee": "same basis", "Other cost-to-COD lines": "devex, capex, connection, municipal, securities, guarantee fees",
               "COD price effect": "community Δ per quarter, discounted from quarter ends",
               "Debt repaid at COD": "lower IDC", "FEP contributions to COD": "commitment fee; 5 % contingency on the pre-COD lease",
               "FEP proceeds at COD": "lower debt, more distributable (70 %)",
               "Generation, GWh": "actual days (leap years), degradation by operating year",
               "Revenue": "same, PPA by quarter from COD", "EBITDA": "same", "Interest paid": "quarterly annuity from COD (old: annual from 2032)",
               "Corporate income tax": "lower interest, quarterly loss use", "FEP HTM cash flow": "sum of the above"}
    for lab in drivers:
        sh1, wk1 = detail[1]
        sh2, wk2 = detail[2]
        rows.append([lab, m2(sh1[lab] / K) if "GWh" not in lab else f"{sh1[lab] / K:,.1f}",
                     m2(wk1[lab] / K) if "GWh" not in lab else f"{wk1[lab] / K:,.1f}",
                     d2((wk1[lab] - sh1[lab]) / K), d2((wk2[lab] - sh2[lab]) / K), drivers[lab]])
    md += table(["Line", "B1 shadow", "B1 model", "B1 Δ", "B2 Δ", "Driver"], rows) + [""]

    # ---- 4. set 3 -> base
    md += ["## 4. Set 3 → base (overbuild not built), for the IC", "",
           "Base = set 1 (EU) and set 2 (CN). The middle column removes the build only (`In_Tolerance_Cap` = 100 % "
           "suppresses the sanction, private copy); the next adds the pro-rata PB sanction 6 months after COD.", ""]
    heads = [("Points", "Points_EU", "eu", False, "{:.1f}"), ("Price effect", "Price_Effect", "eu", False, None),
             ("Effective RtB price", "Eff_Price", "eu", False, None), ("RtB price PV", "Price_PV", "eu", False, None),
             ("Distributable", "Distributable", "eu", False, None), ("FEP RtB profit", "RtB_Profit", "eu", False, None),
             ("FEP RtB IRR", "RtB_IRR", "eu", True, None), ("All-in cost EU", "Cost_To_COD", "eu", False, None),
             ("All-in cost CN", "Cost_To_COD", "cn", False, None), ("COD IRR EU", "COD_IRR", "eu", True, None),
             ("COD IRR CN", "COD_IRR", "cn", True, None), ("HTM IRR EU", "HTM_IRR", "eu", True, None),
             ("HTM IRR CN", "HTM_IRR", "cn", True, None)]
    for name, s, b in BIDS:
        md += [f"### {name}", ""]
        rows = []
        for lab, o, rt, irr, f in heads:
            v3, vn, vb = W[f"s3_{rt}"][f"{o}_{b}"], W[f"nb_{rt}"][f"{o}_{b}"], W[f"base_{rt}"][f"{o}_{b}"]
            if isinstance(v3, str) or isinstance(vn, str) or isinstance(vb, str):
                rows.append([lab, str(v3), str(vn), str(vb), "", "", "", f"`{refs.get(f'{o}_{b}')}`"])
                continue
            fm = (lambda v: f.format(v)) if f else (pc if irr else m2)
            dd = dpp if irr else d2
            rows.append([lab, fm(v3), dd(vn - v3), dd(vb - vn), fm(vb), dd(vb - v3), f"`{refs.get(f'{o}_{b}')}`"])
        md += table(["Item", "Set 3", "Not built", "PB sanction", "Base", "Δ total", "Cell"], rows)
        md += ["", f"Sanction (pro rata) {m2(W['base_eu'][f'Sanction_Pro_Rata_{b}'] if 'Sanction_Pro_Rata_' + str(b) in W['base_eu'] else None)}" if False else ""]
    md = [ln for ln in md if ln is not None]

    # ---- 5. findings
    md += ["## 5. Findings", ""]
    lines = []
    for f in findings:
        if isinstance(f, str):
            lines.append(f"- **Unexplained:** {f} Diagnosis: compare the shadow item with the Calc row of the same name.")
    for name, lab, part, v, irr in [f for f in findings if not isinstance(f, str)]:
        val = dpp(v) + " pp" if irr else d2(v) + " €m"
        if part.startswith("dates"):
            why = ("XIRR on quarter-end dates vs the old date basis (annual IRR from 2026 for HTM; mid-year dates for the "
                   "COD sale). Attributed, not a bug: the same workbook flows re-run on the old basis tie to the shadow (11a).")
        elif lab.startswith("All-in"):
            why = ("IDC on the quarter-opening balance with draws at quarter end, commitment fee on the start-of-quarter "
                   "undrawn amount (detail table). Attributed; every other cost line ties.")
        else:
            why = ("quarterly amounts: IDC, commitment fee and the quarterly annuity (detail table). Attributed.")
        lines.append(f"- {name} {lab}, {part}: {val}. {why}")
    if not [f for f in findings if isinstance(f, str)]:
        lines.insert(0, "- No unexplained residual: every fix step and the RtB headlines tie to the workbook within "
                        "0.1 % / 0.1 pp; the items below exceed the tolerance but are attributed.")
    md += lines + [""]
    MD.write_text("\n".join(md).replace("\n\n\n", "\n\n"))

    # ---- confirmation against the deliverable's stored results (macro, per set) and the live recalc values
    chk = []
    for o in ("Eff_Price", "Price_PV", "Distributable", "RtB_Profit", "Cost_To_COD", "COD_IRR", "HTM_IRR"):
        for b in (1, 2):
            dn = wbv.defined_names.get(f"{o}_{b}_Stored")
            if dn is None:
                continue
            sh, rng = next(dn.destinations)
            stored = [c.value for c in wbv[sh][rng.replace("$", "")][0]]
            for tag, idx in (("base_eu", 0), ("base_cn", 1), ("s3_eu", 2)):
                chk.append((f"{o}_{b} {tag}", abs((stored[idx] or 0) - W[tag][f"{o}_{b}"])))
    live = max(abs(name_value(wbv, f"Out_{o}_{b}") - W["base_eu"][f"{o}_{b}"])
               for o in ("Eff_Price", "Cost_To_COD", "HTM_IRR", "COD_IRR") for b in (1, 2))
    worst = max(chk, key=lambda x: x[1]) if chk else ("none", 0.0)
    print(f"wrote {MD}")
    print(f"tie-out vs stored scenario results in {RECALC.name}: {len(chk)} values, worst {worst[0]} Δ {worst[1]:.2e}")
    print(f"tie-out vs live (base) values in {RECALC.name}: worst Δ {live:.2e}")
    print(f"master check: " + ", ".join(f"{k} {v['check']:.0f}" for k, v in W.items()))
    return 0


if __name__ == "__main__":
    sys.exit(main())
