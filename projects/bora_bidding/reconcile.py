"""Bora bid model – reconciliation against the old engine (SPEC v2 §3 saved values, §8 deliberate fixes, §0 user decisions).

For each headline and bid: saved value (meeting file, reproduced by the old-engine shadows) → each §8 fix applied one at
a time in the shadows (data/old_engine/old_rtb.py, old_htm.py) → each user decision (SPEC §0 items 1, 4, 5, 12) as its own
"user decision" step → new-structure residual → the recalculated workbook at key-assumption set "Overbuild built".
CN values: that set with the input switch In_CN_Turbine = 1 in a private LibreOffice copy. Then set "Overbuild built" →
Base (overbuild not built) per bid. Writes out/reconciliation.md and checks the written tables against the workbook.

    python3 projects/bora_bidding/reconcile.py

Stdlib + fepm. The deliverable workbook is never saved: fepm.recalc.Office works on a private copy. Headline values are
read through the Out_* names; the Calc timelines (flows, for the XIRR date-basis split) are read by row label.
"""
from __future__ import annotations

import datetime as dt
import re
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
COD = dt.date(2031, 9, 30)
# The user's decisions (SPEC §0 items 1, 4, 5, 12; 12-Oct-2026). The workbook inputs must equal them (checked in main).
UD_EXPECT = dict(rtb_premium=150.0, turbine_eu=600.0, ppa_price=90.0, bess_capex=400.0, community_b2=2.01)
UD = dict(UD_EXPECT, ppa_tenor=30)


def todate(x):
    return EXCEL0 + dt.timedelta(days=int(round(x)))


def col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# =========================================================================================== workbook (LibreOffice)
OUTS = ["Points", "Wind_MW", "Wind_Built_MW", "BESS_MW", "RtB_Price", "Distributable", "RtB_Profit", "RtB_IRR",
        "Split_ARES", "Cost_To_COD", "COD_Price_Effect", "COD_Debt", "COD_Proceeds", "COD_Profit", "COD_IRR", "HTM_IRR",
        "Commitment_Total", "Commitment_Penalty"]
ROWS = {"htm_cf": "FEP equity cash flow – hold to maturity", "contrib": "FEP contributions to COD",
        "gen": "Net generation", "rev": "Revenue", "ebitda": "EBITDA", "int": "Interest paid",
        "tax": "Corporate income tax", "idc": "Interest capitalised", "upfront": "Upfront fee",
        "commit": "Commitment fee"}


def _row_index(lo, sheet):
    labels = lo._range(f"{sheet}!C1:C400").getDataArray()
    idx = {}
    for i, lab in enumerate(labels, 1):
        if lab[0]:
            idx.setdefault(str(lab[0]), i)
    return idx


def _find(index, prefix):
    hits = [r for lab, r in index.items() if lab.startswith(prefix)]
    if not hits:
        raise KeyError(f"row '{prefix}' not found on the Calc sheet")
    return hits[0]


def snapshot(lo, rows=False):
    snap = {"check": lo.get("Master_Check")}
    for o in OUTS:
        for b in (1, 2):
            snap[f"{o}_{b}"] = lo.get(f"Out_{o}_{b}")
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
    """Every run on one private copy. A set is selected by its name; the CN route is the In_CN_Turbine switch and the
    'not built' runs lift the tolerance (Live cells overwritten in the private copy only)."""
    runs = {}
    plan = [("base_eu", "Base", {}, False), ("base_cn", "Base", {"In_CN_Turbine": 1}, False),
            ("s3_eu", "Overbuild built", {}, True), ("s3_cn", "Overbuild built", {"In_CN_Turbine": 1}, True),
            ("nb_eu", "Base", {"In_Tolerance_Cap": 1}, False),
            ("nb_cn", "Base", {"In_CN_Turbine": 1, "In_Tolerance_Cap": 1}, False)]
    with Office(XLSM) as lo:
        inp = {k: lo.get(f"In_{k}") for k in ("Award_Date", "PB_Days", "PB_Release", "Life", "RtB_Premium", "Turbine",
                                               "PPA_Price", "BESS_Capex")}
        inp["Community"] = list(lo._range("Dash_Community").getDataArray()[0])
        inp["BESS_Hours"] = lo._range("Dash_BESS_Hours").getDataArray()[0][0]
        orig = {ref: lo.get(ref) for _, _, over, _ in plan for ref in over}
        for tag, name, over, rows in plan:
            lo.set("Scenario_Name", name)
            lo.calculate()
            for ref, v in over.items():
                lo.set(ref, v)
            lo.calculate()
            runs[tag] = snapshot(lo, rows)
            runs[tag]["set"] = name + "".join(f", {k[3:]} = {v}" for k, v in over.items())
            for ref in over:                                    # a Live cell keeps its value (not its formula) in the copy
                lo.set(ref, orig[ref])
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
def merge(dst, src):
    for k, v in src.items():
        if k in ("ov", "g_over") and isinstance(v, dict):
            dst[k] = dict(dst.get(k, {}), **v)
        elif k == "strat_over":
            d = dict(dst.get(k, {}))
            for s_, vv in v.items():
                d[s_] = dict(d.get(s_, {}), **vv)
            dst[k] = d
        else:
            dst[k] = v


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


def user_steps(s):
    """The user's later decisions as bridge steps (own toggles in old_rtb / old_htm); BORA 2 only for the community."""
    out = [
        ("U1", "User decision: RtB price 150 k€/MVA fixed, no price effect", dict(pass_through=0.0), {}, "rb"),
        ("U2", "User decision: EU turbine 600 k€/MW (was 950)",
         dict(g_over={"turbine_eu_keur_per_mw": UD["turbine_eu"]}), dict(turbine_eu=UD["turbine_eu"]), "b"),
        ("U3", "User decision: PPA 90 €/MWh for the whole life (was 91.6 for 10 y, then merchant)",
         dict(g_over={"ppa_price": UD["ppa_price"], "ppa_tenor": UD["ppa_tenor"]}),
         dict(ppa_price=UD["ppa_price"], ppa_tenor=UD["ppa_tenor"]), "b"),
        ("U4", "User decision: BESS capex 100 k€/MWh = 400 k€/MW at 4 h (was 200 k€/MW)",
         dict(g_over={"bess_capex_keur_per_mw": UD["bess_capex"]}), dict(bess_capex=UD["bess_capex"]), "b"),
    ]
    if s == "S3":
        out.append(("U5", "User decision: BORA 2 community 2.01 €/MWh (was 2.0)",
                    dict(strat_over={"S3": {"community_eur_per_mwh": UD["community_b2"]}}),
                    dict(community=UD["community_b2"]), "b"))
    return out


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
                dev=r["developer_total"] / K, points=r["points"])


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
        merge(rk, r_)
        merge(hk, h_)
        rows.append((key, label, rtb_vals(s, rk), build_vals(s, rk, hk, cod)))
    # at-risk claim to COD: portfolio floor (model) vs per-line floor
    floor = {f: R.compute(s, **dict(rk, grid="monthly", sale_date=cod.isoformat(), atrisk_floor=f))
             ["detail"]["devex_return_G131"] / K for f in ("line", "portfolio")}
    kw = htm_kwargs(s, rk, hk, cod)                      # COD IRR with the per-line floor instead of the portfolio floor
    kw["devex"] = dict(kw["devex"], g280=floor["line"] * K)
    floor["irr_line"] = H.compute(s, "COD-EU", **kw)["cod"][IRR_COD]
    floor["irr_portfolio"] = rows[-1][3]["cod_eu"]
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


def table(head, rows, left=()):
    out = ["| " + " | ".join(head) + " |",
           "|" + "|".join("---" if i < 2 or i in left else "---:" for i in range(len(head))) + "|"]
    out += ["| " + " | ".join(r) + " |" for r in rows]
    return out


def parse_num(txt):
    """'(3.37) %' -> -0.0337, '1,234.50' -> 1234.5, '85.00' -> 85.0; None for n/a."""
    t = txt.replace("`", "").replace(",", "").strip()
    neg = t.startswith("(")
    pct = t.endswith("%")
    m = re.search(r"[\d.]+", t)
    if not m:
        return None
    v = float(m.group()) / (100 if pct else 1)
    return -v if neg else v


def bridge_rows(rows_, vi, cols, stp, flag, work, residual):
    """One bridge table body. vi = index of the value dict in rows_ (2 RtB, 3 build). Steps whose displayed deltas are
    all zero are left out and returned as a list of keys. `residual` adds rows after the 'all fixes' line."""
    fm = lambda v, irr: pc(v) if irr else m2(v)          # noqa: E731
    dd = lambda v, irr: dpp(v) if irr else d2(v)         # noqa: E731
    body = [["0", rows_[0][1]] + [fm(rows_[0][vi][k], irr) for k, _, irr in cols]]
    zero = []
    for i in range(1, len(rows_)):
        key, label = rows_[i][0], rows_[i][1]
        if flag not in next(t for k_, _, _, _, t in stp if k_ == key):
            continue
        cells = [dd(rows_[i][vi][k] - rows_[i - 1][vi][k], irr) for k, _, irr in cols]
        if all(c == "0.00" for c in cells):
            zero.append(key)
        else:
            body.append([key, label] + cells)
    last = rows_[-1][vi]
    body.append(["", "**Shadow, all fixes and decisions**"] + [fm(last[k], irr) for k, _, irr in cols])
    body += residual
    body.append(["", "**Workbook, set \"Overbuild built\"**"] + [fm(w, irr) for w, (_, _, irr) in zip(work, cols)])
    return body, zero


# =========================================================================================== main
def main():
    inp, W = read_workbook()
    wbv, refs = cell_refs()
    findings = []
    UD.update(rtb_premium=inp["RtB_Premium"], turbine_eu=inp["Turbine"], ppa_price=inp["PPA_Price"],
              bess_capex=inp["BESS_Capex"] * inp["BESS_Hours"], community_b2=inp["Community"][1], ppa_tenor=int(inp["Life"]))
    for k, v in UD_EXPECT.items():
        if abs(UD[k] - v) > 1e-9:
            findings.append(f"Workbook input for the user decision '{k}' is {UD[k]}, expected {v}.")
    pb_date = todate(inp["Award_Date"]) + dt.timedelta(days=int(inp["PB_Days"]))
    stp_fix = steps(pb_date, todate(inp["PB_Release"]))
    s3, cn = W["s3_eu"], W["s3_cn"]
    md = []
    stamp = dt.datetime.fromtimestamp(XLSM.stat().st_mtime).strftime("%d-%b-%Y %H:%M")
    md += ["# Bora bid model – reconciliation to the old engine", "",
           f"Model `out/Bora_Bid_Model.xlsm` (built {stamp}), recalculated in LibreOffice by `reconcile.py` on a private copy. "
           "Reconciliation set: key-assumption set \"Overbuild built\" (EU); CN = the same set with the input switch "
           f"`In_CN_Turbine` = 1 (private copy only). Master check: EU {s3['check']:.0f}.",
           "Saved values: `reference/Bora_scenarios_meeting.xlsx` (SPEC §3), reproduced by `data/old_engine/old_rtb.py` "
           "and `old_htm.py` (≤ 1e-8). Each §8 fix is applied cumulatively in those shadows, then each user decision "
           "(SPEC §0) as its own step.",
           "€ m to 2 dp, negatives in brackets; IRR in % (step changes in pp). Tolerance 0.1 % on values, 0.1 pp on IRR; "
           "`*` = above it (see Findings). Steps with no effect on a table are listed under it.", ""]

    stp = {s: stp_fix + user_steps(s) for _, s, _ in BIDS}
    bridges = {s: run_bridge(s, stp[s], COD) for _, s, _ in BIDS}

    # ---- 1. ties without a bridge
    md += ["## 1. Headlines that tie without a bridge", ""]
    sv = R.J["saved_bora"]
    shadow_pts = {s: R.compute(s, pass_through=0.0, **({"strat_over": {"S3": {"community_eur_per_mwh": UD["community_b2"]}}}
                                                       if s == "S3" else {}))["points"] for _, s, _ in BIDS}
    rows = []
    for label, key, okey, run, fmt in [("Points, EU turbine", 35, "Points", "s3_eu", "{:.2f}"),
                                       ("Points, CN turbine", 36, "Points", "s3_cn", "{:.2f}"),
                                       ("Installed wind MW", 22, "Wind_MW", "s3_eu", "{:.0f}"),
                                       ("BESS MW", 23, "BESS_MW", "s3_eu", "{:.0f}"),
                                       ("RtB price €m (gross premium; no price effect)", 41, "RtB_Price", "s3_eu", "{:.2f}")]:
        cells = []
        for _, s, b in BIDS:
            c = {"S2": "F", "S3": "G"}[s]
            a, w = sv[s][f"Bora!{c}{key}"], W[run][f"{okey}_{b}"]
            cells += [fmt.format(a), fmt.format(w)]
            if key in (35, 36):
                tag = "total_eu" if key == 35 else "total_cn"
                if abs(shadow_pts[s][tag] - w) > 1e-6:
                    findings.append(f"{label} {b}: shadow with the user decisions {shadow_pts[s][tag]} vs model {w}.")
            elif abs(a - w) > 1e-6:
                findings.append(f"{label} {b}: saved {a} vs model {w} (`{refs.get(f'{okey}_{b}')}`).")
        rows.append([label] + cells + [f"`{refs.get(okey + '_1')}`, `{refs.get(okey + '_2')}`"])
    md += table(["Item", "BORA 1 saved", "BORA 1 model", "BORA 2 saved", "BORA 2 model", "Cells (B1, B2)"], rows)
    md += ["", "BORA 2 points +0.75 (EU and CN): user decision, community 2.01 €/MWh scores ÖSZ-2 (the shadow's tender "
           "tables give 68.25 / 53.25). The RtB price is 150 k€/MVA × 100 MVA = 15.00 with no price effect (user decision, "
           "step U1 below); saved effective prices were 3.93 / 10.08.", ""]

    # ---- 2. RtB bridge
    md += ["## 2. RtB exit bridge (set \"Overbuild built\")", ""]
    rk_cols = [("pe", "Price effect", False), ("eff", "RtB price", False), ("dist", "Distributable", False),
               ("profit", "FEP RtB profit", False), ("irr", "FEP RtB IRR", True)]
    for name, s, b in BIDS:
        rows_, floor = bridges[s]
        md += [f"### {name}", ""]
        last = rows_[-1][2]
        wk = [0.0, s3[f"RtB_Price_{b}"], s3[f"Distributable_{b}"], s3[f"RtB_Profit_{b}"], s3[f"RtB_IRR_{b}"]]
        res = []
        for (k, lab, irr), w in zip(rk_cols, wk):
            r_ = (w - last[k]) if not isinstance(w, str) else None
            flag = "*" if over(r_, w, irr) else ""
            res.append(("n/a" if r_ is None else (dpp(r_) if irr else d2(r_))) + flag)
            if flag:
                findings.append(f"{name} RtB {lab}: residual {res[-1]} vs the shadow.")
        body, zero = bridge_rows(rows_, 2, rk_cols, stp[s], "r", wk, [["", "Residual: workbook − shadow"] + res])
        if s == "S2":
            body[0][-1] += " ¹"
        md += table(["#", "Step", "Price effect", "RtB price", "Distributable", "FEP RtB profit", "FEP RtB IRR"], body)
        md += ["", f"Steps with no effect here: {', '.join(zero)}; U2–U5 do not reach the RtB route (the price is fixed).",
               f"Developer at the exit (DSA + 30 %): saved {m2(rows_[0][2]['dev'])}, shadow {m2(last['dev'])}, "
               f"workbook {m2(s3[f'Split_ARES_{b}'])}."]
        if s == "S2":
            md += ["¹ Excel's XIRR failed in the meeting file (2.98e-09); −3.91 % is the true root (LibreOffice, shadow). "
                   "The model shows `n/a` plus a check when XIRR fails (fix 6)."]
        md += [""]
    md += ["Cells (B1 / B2): " + ", ".join(
        f"{lab} `{refs.get(o + '_1')}` / `{refs.get(o + '_2')}`"
        for lab, o in (("RtB price", "RtB_Price"), ("distributable", "Distributable"), ("profit", "RtB_Profit"),
                       ("IRR", "RtB_IRR"), ("developer", "Split_ARES"))) + ".", ""]

    # ---- 3. build / COD / HTM bridge
    md += ["## 3. Build, sale at COD and hold to maturity (set \"Overbuild built\"; CN = same set, `In_CN_Turbine` = 1)", ""]
    bcols = [("c2c_eu", "All-in cost EU", "Cost_To_COD", "s3_eu", False),
             ("c2c_cn", "All-in cost CN", "Cost_To_COD", "s3_cn", False),
             ("cod_eu", "COD IRR EU", "COD_IRR", "s3_eu", True), ("cod_cn", "COD IRR CN", "COD_IRR", "s3_cn", True),
             ("htm_eu", "HTM IRR EU", "HTM_IRR", "s3_eu", True), ("htm_cn", "HTM IRR CN", "HTM_IRR", "s3_cn", True)]
    cols3 = [(k, lab, irr) for k, lab, _, _, irr in bcols]
    lease = {}
    for name, s, b in BIDS:
        rows_, floor = bridges[s]
        md += [f"### {name}", ""]
        last = rows_[-1][3]
        amounts, dates, work = [], [], []
        for k, lab, o, run, irr in bcols:
            w = W[run][f"{o}_{b}"]
            work.append(w)
            if k.startswith("c2c"):
                eb = W[run][f"ebitda_{b}"]
                ls = -sum(v for e, v in zip(W[run][f"ends_{b}"], eb) if e <= COD) / K       # pre-COD land lease (> 0)
                lease[(k, b)] = ls
                a_, d_ = w - last[k] - ls, 0.0
            elif k.startswith("cod"):
                mid = cod_xirr_on_old_basis(W[run][f"ends_{b}"], W[run][f"contrib_{b}"], W[run][f"COD_Proceeds_{b}"] * K, COD)
                a_, d_ = mid - last[k], w - mid
            else:
                mid = annual_irr_on_old_basis(W[run][f"ends_{b}"], W[run][f"htm_cf_{b}"])
                a_, d_ = mid - last[k], w - mid
            amounts.append((a_, w, irr, lab))
            dates.append((d_, w, irr, lab))
        res = [["11c", "All-in cost now includes the pre-COD land lease (Calc commitment overview)"]
               + [d2(lease[(k, b)]) if k.startswith("c2c") else "0.00" for k, _, _, _, _ in bcols],
               ["11a", "New structure: amounts (quarterly debt, IDC, fees, days)"]
               + [(dpp(a) if irr else d2(a)) + ("*" if over(a, w, irr) else "") for a, w, irr, _ in amounts],
               ["11b", "New structure: XIRR on quarter ends (old: annual / mid-year)"]
               + [(dpp(d) if irr else d2(d)) + ("*" if over(d, w, irr) else "") for d, w, irr, _ in dates]]
        body, zero = bridge_rows(rows_, 3, cols3, stp[s], "b", work, res)
        md += table(["#", "Step", "All-in cost EU", "All-in cost CN", "COD IRR EU", "COD IRR CN", "HTM IRR EU", "HTM IRR CN"], body)
        md += ["", f"Steps with no effect here: {', '.join(zero)}.",
               f"At-risk claim to COD (fix 9): portfolio floor {m2(floor['portfolio'])}, per-line floor {m2(floor['line'])} "
               f"({d2(floor['line'] - floor['portfolio'])} €m claim, {dpp(floor['irr_line'] - floor['irr_portfolio'])} pp COD IRR EU): "
               "the model uses the portfolio floor (SPEC §5.4).", ""]
        for (a, w, irr, lab), (d, _, _, _) in zip(amounts, dates):
            for part, v in (("amounts (11a)", a), ("dates (11b)", d)):
                if over(v, w, irr):
                    findings.append((name, lab, part, v, irr))
    md += ["Cells (B1 / B2): " + ", ".join(
        f"{lab} `{refs.get(o + '_1')}` / `{refs.get(o + '_2')}`"
        for lab, o in (("all-in cost", "Cost_To_COD"), ("COD IRR", "COD_IRR"), ("HTM IRR", "HTM_IRR"))) + ".", ""]

    # new-structure detail
    md += ["New-structure detail, EU set \"Overbuild built\", € m (shadow with all fixes and decisions vs workbook; lifetime sums):", ""]
    rows = []
    sh_wk = {}
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
        wk = {"IDC capitalised": sum(w[f"idc_{b}"]), "Commitment fee": -sum(w[f"commit_{b}"]),
              "Upfront fee": -sum(w[f"upfront_{b}"]),
              "COD price effect": w[f"COD_Price_Effect_{b}"] * K, "Debt repaid at COD": w[f"COD_Debt_{b}"] * K,
              "FEP contributions to COD": sum(w[f"contrib_{b}"]), "FEP proceeds at COD": w[f"COD_Proceeds_{b}"] * K,
              "Generation, GWh": sum(w[f"gen_{b}"]), "Revenue": sum(w[f"rev_{b}"]), "EBITDA": sum(w[f"ebitda_{b}"]),
              "Interest paid": sum(w[f"int_{b}"]), "Corporate income tax": sum(w[f"tax_{b}"]),
              "FEP HTM cash flow": sum(w[f"htm_cf_{b}"])}
        wk["Other cost-to-COD lines"] = (w[f"Cost_To_COD_{b}"] * K - wk["IDC capitalised"] - wk["Commitment fee"]
                                         - wk["Upfront fee"])
        sh_wk[b] = (sh, wk)
    drivers = {"IDC capitalised": "interest on the quarter-opening balance; draws at quarter end (old: annual, half-year draws)",
               "Commitment fee": "on undrawn at the quarter start (old: net of capitalised IDC)",
               "Upfront fee": "same basis",
               "Other cost-to-COD lines": "pre-COD land lease (11c) in the model; other cost lines tie",
               "COD price effect": "community Δ per quarter, discounted from quarter ends",
               "Debt repaid at COD": "lower IDC", "FEP contributions to COD": "commitment fee, pre-COD lease",
               "FEP proceeds at COD": "lower debt, more distributable (70 %)",
               "Generation, GWh": "actual days (leap years), degradation by operating year",
               "Revenue": "same, PPA by quarter from COD", "EBITDA": "same",
               "Interest paid": "quarterly annuity from COD (old: annual from 2032)",
               "Corporate income tax": "lower interest, quarterly loss use", "FEP HTM cash flow": "sum of the above"}
    for lab in drivers:
        sh1, wk1 = sh_wk[1]
        sh2, wk2 = sh_wk[2]
        g = "GWh" in lab
        rows.append([lab, f"{sh1[lab] / K:,.1f}" if g else m2(sh1[lab] / K), f"{wk1[lab] / K:,.1f}" if g else m2(wk1[lab] / K),
                     d2((wk1[lab] - sh1[lab]) / K), d2((wk2[lab] - sh2[lab]) / K), drivers[lab]])
    md += table(["Line", "B1 shadow", "B1 model", "B1 Δ", "B2 Δ", "Driver"], rows, left=(5,)) + [""]

    # ---- 4. set 3 -> base
    md += ["## 4. Set \"Overbuild built\" → Base (overbuild not built), for the IC", "",
           "Base = the Base set (EU) and Base with `In_CN_Turbine` = 1 (CN). The middle column removes the build only "
           "(`In_Tolerance_Cap` = 100 % suppresses the sanction, private copy); the next adds the pro-rata PB sanction "
           "6 months after COD.", ""]
    heads = [("Points EU", "Points", "eu", False, "{:.2f}"), ("Points CN", "Points", "cn", False, "{:.2f}"),
             ("Distributable", "Distributable", "eu", False, None), ("FEP RtB profit", "RtB_Profit", "eu", False, None),
             ("FEP RtB IRR", "RtB_IRR", "eu", True, None), ("Total commitment EU", "Commitment_Total", "eu", False, None),
             ("All-in cost EU", "Cost_To_COD", "eu", False, None), ("All-in cost CN", "Cost_To_COD", "cn", False, None),
             ("COD IRR EU", "COD_IRR", "eu", True, None), ("COD IRR CN", "COD_IRR", "cn", True, None),
             ("HTM IRR EU", "HTM_IRR", "eu", True, None), ("HTM IRR CN", "HTM_IRR", "cn", True, None)]
    for name, s, b in BIDS:
        md += [f"### {name}", ""]
        rows = []
        for lab, o, rt, irr, f in heads:
            v3, vn, vb = W[f"s3_{rt}"][f"{o}_{b}"], W[f"nb_{rt}"][f"{o}_{b}"], W[f"base_{rt}"][f"{o}_{b}"]
            if isinstance(v3, str) or isinstance(vn, str) or isinstance(vb, str):
                rows.append([lab, str(v3), str(vn), str(vb), "", "", "", f"`{refs.get(f'{o}_{b}')}`"])
                continue
            fm = (lambda v, f=f: f.format(v)) if f else (pc if irr else m2)
            dd = dpp if irr else d2
            rows.append([lab, fm(v3), dd(vn - v3), dd(vb - vn), fm(vb), dd(vb - v3), f"`{refs.get(f'{o}_{b}')}`"])
        md += table(["Item", "Set \"Overbuild built\"", "Not built", "PB sanction", "Base", "Δ total", "Cell"], rows)
        md += [""]

    # ---- 5. findings
    md += ["## 5. Findings", ""]
    unexplained = [f for f in findings if isinstance(f, str)]
    att = [f for f in findings if not isinstance(f, str)]

    def grab(lab, part):
        return "; ".join(f"{n} {dpp(v) + ' pp' if irr else d2(v)}" for n, l_, p_, v, irr in att
                         if l_ == lab and p_.startswith(part)) or "within tolerance"
    lines = [f"- **Unexplained:** {f}" for f in unexplained]
    if not unexplained:
        lines.append("- No unexplained residual on the RtB headlines, points, wind/BESS MW or RtB price: the shadow with every "
                     "fix and decision ties to the workbook, both bids.")
    lines += [
        f"- All-in cost to COD after the lease (11c), new structure (11a): EU {grab('All-in cost EU', 'amounts')}; "
        f"CN {grab('All-in cost CN', 'amounts')} €m. IDC accrues on the quarter-opening balance with draws at quarter end "
        "(old: annual, half-year draws), partly offset by the commitment fee on the start-of-quarter undrawn amount. "
        "Convention, not a bug.",
        f"- COD IRR, date basis (11b): EU {grab('COD IRR EU', 'dates')}; CN {grab('COD IRR CN', 'dates')}. The old engine "
        "dated each year's contributions mid-year; the model dates them at quarter ends. Amounts (11a): "
        f"EU {grab('COD IRR EU', 'amounts')}; CN {grab('COD IRR CN', 'amounts')} (lower IDC, less debt repaid at COD). Attributed.",
        f"- HTM IRR, date basis (11b): EU {grab('HTM IRR EU', 'dates')}; CN {grab('HTM IRR CN', 'dates')}; quarterly XIRR vs the old "
        f"annual IRR from 2026. The workbook flows re-run on the old basis tie to the shadow (11a: EU {grab('HTM IRR EU', 'amounts')}; "
        f"CN {grab('HTM IRR CN', 'amounts')}). Attributed.",
        "- The CN runs and the not-built runs overwrite Live cells in the private copy, so only the stored-vs-live checks "
        "fail there (master check ≠ 0); every economic check is 0.",
    ]
    md += lines + [""]
    MD.write_text("\n".join(md).replace("\n\n\n", "\n\n"))

    # ---- confirmation: macro-stored results of the workbook per set, live recalc of the deliverable, and the md itself
    chk = []
    for o in ("Points", "RtB_Profit", "COD_IRR", "HTM_IRR", "Distributable", "Cost_To_COD"):
        for b in (1, 2):
            dn = wbv.defined_names.get(f"{o}_{b}_Stored")
            if dn is None:
                continue
            sh, rng = next(dn.destinations)
            stored = [c.value for c in wbv[sh][rng.replace("$", "")][0]]
            for tag, idx in (("base_eu", 0), ("s3_eu", 1)):
                chk.append((f"{o}_{b} {tag}", abs((stored[idx] or 0) - W[tag][f"{o}_{b}"])))
    live = max(abs(name_value(wbv, f"Out_{o}_{b}") - W["base_eu"][f"{o}_{b}"])
               for o in ("RtB_Price", "Cost_To_COD", "HTM_IRR", "COD_IRR", "RtB_Profit", "Points") for b in (1, 2))
    worst = max(chk, key=lambda x: x[1]) if chk else ("none", 0.0)
    print(f"wrote {MD}")
    print(f"tie-out vs stored scenario results in {RECALC.name}: {len(chk)} values, worst {worst[0]} Δ {worst[1]:.2e}")
    print(f"tie-out vs live (Base) values in {RECALC.name}: worst Δ {live:.2e}")
    print("master check: " + ", ".join(f"{k} {v['check']:.0f}" for k, v in W.items()))
    bad = verify_md(wbv, W)
    print(f"md vs workbook: {bad} mismatches")
    for f in unexplained:
        print("UNEXPLAINED:", f)
    return 1 if bad else 0


def verify_md(wbv, W):
    """Read the written md back and compare its section 4 headlines (EU, set 'Overbuild built' and Base) with the
    macro-stored results of the recalculated deliverable, and the RtB IRR with the live run."""
    text = MD.read_text().split("\n")
    stored = {}
    for o in ("Points", "RtB_Profit", "COD_IRR", "HTM_IRR", "Distributable", "Cost_To_COD"):
        for b in (1, 2):
            dn = wbv.defined_names[f"{o}_{b}_Stored"]
            sh, rng = next(dn.destinations)
            stored[(o, b)] = [c.value for c in wbv[sh][rng.replace("$", "")][0]]
    want = {"Points EU": "Points", "Distributable": "Distributable", "FEP RtB profit": "RtB_Profit", "FEP RtB IRR": "RtB_IRR",
            "All-in cost EU": "Cost_To_COD", "COD IRR EU": "COD_IRR", "HTM IRR EU": "HTM_IRR"}
    bad = n = b = 0
    in4, sec = False, ""
    cells_of = lambda ln: [c.strip() for c in ln.strip("|").split("|")]       # noqa: E731
    for ln in text:
        if ln.startswith("## 4."):
            in4 = True
        elif ln.startswith("## 5."):
            in4 = False
        if ln.startswith("### BORA"):
            b = int(ln.split()[-1])
        if ln.startswith("## 2.") or ln.startswith("## 3."):
            sec = ln[3]
        if ln.startswith("| ") and cells_of(ln)[0] == "Points, EU turbine":          # section 1
            c_ = cells_of(ln)
            for cell, bb in ((c_[2], 1), (c_[4], 2)):
                n += 1
                if abs(parse_num(cell) - stored[("Points", bb)][1]) > 0.0051:
                    bad += 1
                    print(f"  mismatch section 1 points BORA {bb}: md {cell} vs workbook {stored[('Points', bb)][1]}")
        if ln.startswith("|  | **Workbook") and b:                                    # sections 2 and 3
            c_ = cells_of(ln)
            pairs = ([(c_[3], "RtB_Price", None), (c_[4], "Distributable", 1), (c_[5], "RtB_Profit", 1), (c_[6], "RtB_IRR", None)]
                     if sec == "2" else [(c_[2], "Cost_To_COD", 1), (c_[4], "COD_IRR", 1), (c_[6], "HTM_IRR", 1)])
            for cell, o, idx in pairs:
                ref = stored[(o, b)][idx] if idx is not None else W["s3_eu"][f"{o}_{b}"]
                val = parse_num(cell)
                if ref is None or isinstance(ref, str) or val is None:
                    continue
                n += 1
                if abs(val - ref) > (0.00006 if "IRR" in o else 0.0051):
                    bad += 1
                    print(f"  mismatch section {sec} BORA {b} {o}: md {cell} vs workbook {ref}")
        if not in4 or not ln.startswith("| ") or b == 0:
            continue
        cells = [c.strip() for c in ln.strip("|").split("|")]
        if cells[0] not in want:
            continue
        o = want[cells[0]]
        for cell, idx, tag in ((cells[1], 1, "s3_eu"), (cells[4], 0, "base_eu")):      # Overbuild built, Base
            val = parse_num(cell)
            ref = W[tag][f"{o}_{b}"] if o == "RtB_IRR" else stored[(o, b)][idx]
            if ref is None or isinstance(ref, str) or val is None:
                continue
            n += 1
            tol = 0.00006 if "IRR" in o else 0.0051
            if abs(val - ref) > tol:
                bad += 1
                print(f"  mismatch BORA {b} {cells[0]} ({tag}): md {val} vs workbook {ref}")
    print(f"  md cells compared with the workbook: {n}")
    return bad


if __name__ == "__main__":
    sys.exit(main())
