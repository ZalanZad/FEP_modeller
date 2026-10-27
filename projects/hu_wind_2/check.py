"""Verification of projects/hu_wind_2/out/HU_Wind_bid_deposits_fees.xlsx (SPEC bora_v5_summary §21).

1. LibreOffice recalc (fepm.recalc.Office) -> out/HU_Wind_bid_deposits_fees_recalc.xlsx; no error values in the new
   rows, Master_Check 0. (The user's E22/E31 use _xlfn.SINGLE, which LibreOffice cannot evaluate; reported.)
2. Lint (fepm.audit.lint) on the new blocks (rows from 81): workbook lint filtered to the new rows, plus the
   calc-sheet rules the sheet name does not trigger: one formula per timeline row (X:BG) and direction
   (up/left; the summary rows at the top of the HTM block are a summary panel and exempt).
3. Independent Python recomputation from the input cells only: Iváncsa RtB claims (closed form per line), waterfall,
   FEP profit / MOIC / XIRR; Iváncsa HTM project IRR, NPV, COD price, COD-exit FEP profit / IRR. Must match to 1e-6.
4. out/dump.txt: label | unit | Iváncsa | Dunaújváros for the new inputs, the RtB exit and the HTM results rows.
5. §22 bids workbook (out/HU_Wind_bids.xlsx): recalc, lint (+ Dashboard timeline and direction rules), independent
   recomputation of all three projects (points, RtB waterfall with DSA and FEP share, HTM IRR, COD price and exit),
   Iváncsa / Dunaújváros equal to the §21 build, out/dump_bids.txt.

    python3 projects/hu_wind_2/check.py
"""
import json
import re
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import column_index_from_string as ci, range_boundaries

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parents[1]))
from fepm.audit.lint import _direction, _name_targets, lint, lint_values, to_r1c1  # noqa: E402
from fepm.recalc import Office  # noqa: E402

OUT = HERE / "out" / "HU_Wind_bid_deposits_fees.xlsx"
RC = HERE / "out" / "HU_Wind_bid_deposits_fees_recalc.xlsx"
DUMP = HERE / "out" / "dump.txt"
R = json.loads((HERE / "out" / "rows.json").read_text())
FIRST = 81
TL = range(ci("X"), ci("BG") + 1)


def recalc():
    with Office(OUT) as lo:
        lo.calculate()
        lo.save(RC)


def lint_new():
    """Findings on the new rows; pre-existing findings in the user's rows are reported separately."""
    found = lint(OUT)
    new = [f for f in found if f.where.startswith("Bid_Cash!") and _row(f.where) >= FIRST]
    old = [f for f in found if f not in new]
    wb = openpyxl.load_workbook(OUT)
    ws, names = wb["Bid_Cash"], _name_targets(wb)
    lo, hi = R["summary"]
    for row in ws.iter_rows(min_row=FIRST):
        forms = {}
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                if not lo <= c.row <= hi:
                    new += _direction("Bid_Cash", c.row, c.column, c.value, names)
                if c.column in TL and c.row >= R["t_yr"]:
                    forms[c.column] = to_r1c1(c.value, c.row, c.column)
        if forms and (len(set(forms.values())) > 1 or len(forms) != len(TL)):
            new.append(f"[c-row] Bid_Cash!row {row[0].row}: timeline row not uniform or not filled")
    return new, old


def _row(where):
    m = re.search(r"(\d+)$", where)
    return int(m.group(1)) if m else 0


# ---------------------------------------------------------------------------------------------------------------------
def d(x):
    return x.date() if isinstance(x, datetime) else x


def xnpv(rate, flows):
    d0 = flows[0][0]
    return sum(v / (1 + rate) ** ((t - d0).days / 365) for t, v in flows)


def xirr(flows):
    lo, hi = -0.99, 10.0
    for _ in range(300):
        mid = (lo + hi) / 2
        if xnpv(mid, flows) > 0:
            lo = mid
        else:
            hi = mid
    return (lo + hi) / 2


def recompute(v, col="E"):
    """Independent recomputation for one project from the input cells (values) only."""
    g = lambda r: v[f"{col}{r}"].value  # noqa: E731
    e = lambda r: v[f"E{r}"].value      # common inputs and the user's key dates (column E)  # noqa: E731
    i = {k: g(R[k]) for k in ("prem", "sc", "share", "dsa", "dr", "sf", "turb", "inst", "fdn", "civ", "elec", "own",
                              "land", "subs", "basis", "ccont", "bess", "hrs", "curt", "clip", "deg", "ppa", "esc",
                              "om", "bop", "lease", "ins", "mgmt", "ptax", "adm", "bal", "grid", "mekh", "hipa",
                              "ocont", "oesc")}
    c = {k: e(R[k]) for k in ("rtb", "sale", "freeze", "devd", "r12", "r20", "floor", "life")}
    # The user's bid inputs, in HUF -> EUR k (recomputed here, not read from Bid_Cash rows 60-78).
    mva, ob, bshare, hours, muni, comm, pbm = g(5), g(6), g(7), g(8), g(9), g(10), g(12)
    fx_, bw, bb, pfee, bsr, pbr, tso, ind, dirf, sdd, devex = (g(r) for r in (45, 46, 47, 48, 49, 50, 51, 52, 53, 55, 56))
    wind, bess_mwh = mva * (1 + ob), mva * bshare * hours
    base = wind * bw + bess_mwh * bb
    k = lambda huf: huf / fx_ / 1000  # noqa: E731
    dates = {n: d(e(r)) for n, r in (("sign", 34), ("bid", 35), ("pb", 36), ("tso", 37), ("hc", 39), ("ntp", 40),
                                     ("cod", 41), ("rel", 42))}
    lines = [("fee", k(pfee), dates["bid"], c["r12"]), ("bs", k(bsr * base), dates["bid"], c["r12"]),
             ("pb", k(pbr * pbm * base - bsr * base), dates["pb"], c["r12"]),
             ("tso", k(tso * mva), dates["tso"], c["r12"]), ("dir", k(dirf), dates["hc"], c["r12"]),
             ("ind", k(ind * mva - tso * mva), dates["hc"], c["r12"]), ("dev", devex, d(c["devd"]), c["r20"])]
    end = min(d(c["freeze"]), d(c["sale"]))

    def factor(p, r):                       # closed form: act/365 both ends, simple within a year, 31-Dec capitalised
        if p.year == end.year:
            return 1 + r * ((end - p).days + 1) / 365
        first = (date(p.year, 12, 31) - p).days + 1
        last = (end - date(end.year, 1, 1)).days + 1
        return (1 + r * first / 365) * (1 + r) ** (end.year - p.year - 1) * (1 + r * last / 365)

    claims = {}
    for n, amt, p, r in lines:
        acc = amt * (factor(p, r) - 1)
        claims[n] = max(acc, amt * (c["floor"] - 1)) if n == "dev" else acc
    out = {"claim_" + n: x for n, x in claims.items()}
    amt = {n: a for n, a, _, _ in lines}
    premium = i["prem"] * mva
    reimb = sum(amt[n] for n in ("bs", "pb", "tso", "dir", "ind"))
    cash = premium + reimb
    paid = {}
    for n, due in (("sc", i["sc"]), ("rp", reimb), ("c12", sum(claims[n] for n in claims if n != "dev")),
                   ("dev", claims["dev"]), ("dsa", i["dsa"])):
        paid[n] = min(cash, due)
        cash -= paid[n]
    dist = cash
    receipts = paid["rp"] + paid["c12"] + paid["dev"] + dist * i["share"]
    flows = [(d(dates["sign"]), -sdd)] + [(p, -a) for _, a, p, _ in lines] + [(d(c["sale"]), receipts)]
    flows.sort(key=lambda t: t[0])
    out.update(rtb_premium=premium, rtb_reimb=reimb, rtb_dist=dist, rtb_receipts=receipts,
               rtb_profit=sum(x for _, x in flows),
               rtb_moic=sum(x for _, x in flows if x > 0) / -sum(x for _, x in flows if x < 0), rtb_irr=xirr(flows))

    # HTM, annual 2026-2061, flows at 31-Dec.
    built = wind * (1 - i["sf"])
    per_mw = sum(i[n] for n in ("turb", "inst", "fdn", "civ", "elec", "own", "land")) * built
    subs = i["subs"] * mva / i["basis"]
    capex = (per_mw + subs) * (1 + i["ccont"]) + i["bess"] * bess_mwh
    years = range(dates["sign"].year, dates["sign"].year + 36)

    def months(start, n):                 # the n month-ends after `start`, by calendar year
        cnt, y, m = {}, start.year, start.month
        for _ in range(n):
            m += 1
            y, m = (y + 1, 1) if m > 12 else (y, m)
            cnt[y] = cnt.get(y, 0) + 1
        return cnt
    nc = (dates["cod"].year - dates["ntp"].year) * 12 + dates["cod"].month - dates["ntp"].month
    cons, oper = months(dates["ntp"], nc), months(dates["cod"], int(c["life"]) * 12)
    out_flows = [(p, -a) for _, a, p, _ in lines] + [(dates["sign"], -sdd)]
    cf, ebitda_y, inv_y = {}, {}, {}
    for y in years:
        inv = sum(x for t, x in out_flows if t.year == y)
        inv += (amt["bs"] + amt["pb"]) * (y == dates["rel"].year)
        inv -= capex * cons.get(y, 0) / nc
        inv -= muni * mva * (y == dates["cod"].year)
        op = oper.get(y, 0) / 12
        n = y - dates["cod"].year
        gen = built * i["hrs"] * (1 - i["curt"]) * (1 - i["clip"]) * op * (1 - i["deg"]) ** max(n - 1, 0) / 1000
        rev = gen * i["ppa"] * (1 + i["esc"]) ** n
        ix = (1 + i["oesc"]) ** n
        base_opex = (sum(i[x] for x in ("om", "bop", "lease", "ins", "mgmt", "ptax")) * built * op
                     + i["adm"] * mva / i["basis"] * op + gen * (i["bal"] + i["grid"])) * ix + rev * i["mekh"]
        opex = base_opex * (1 + i["ocont"]) + gen * comm + rev * i["hipa"]
        ebitda_y[y], inv_y[y] = rev - opex, inv
        cf[y] = inv + rev - opex
    ye = {y: date(y, 12, 31) for y in years}
    price = sum(ebitda_y[y] * (1 + i["dr"]) ** ((dates["cod"] - ye[y]).days / 365) for y in years)
    cod_flows = [(ye[y], inv_y[y] * (y <= dates["cod"].year)
                  + (price + amt["bs"] + amt["pb"]) * (y == dates["cod"].year)) for y in years]
    htm = [(ye[y], cf[y]) for y in years]
    out.update(htm_irr=xirr(htm), htm_npv=xnpv(i["dr"], htm), cod_price=price,
               cod_profit=sum(x for _, x in cod_flows), cod_irr=xirr(cod_flows),
               allin=capex + amt["tso"] + amt["dir"] + amt["ind"] + muni * mva + amt["fee"] + sdd + devex,
               ebitda1=ebitda_y[dates["cod"].year + 1])
    return out


def compare(v, x):
    s = lambda key, col="E": v[f"{col}{R[key]}"].value  # noqa: E731
    model = {f"claim_{n}": s(f"c_{n if n != 'fee' else 'pf'}") for n in ("fee", "bs", "pb", "tso", "dir", "ind", "dev")}
    model.update(rtb_premium=s("w_prem"), rtb_reimb=s("w_reimb"), rtb_dist=s("w_dist"), rtb_receipts=s("f_rec"),
                 rtb_profit=s("r_profit"), rtb_moic=s("r_moic"), rtb_irr=s("r_irr"), htm_irr=s("h_irr"),
                 htm_npv=s("h_npv"), cod_price=s("x_price"), cod_profit=s("x_profit"), cod_irr=s("x_irr"),
                 allin=s("k_allin"), ebitda1=s("h_ebitda"))
    worst = 0.0
    lines = []
    for key, m in model.items():
        diff = abs(m - x[key])
        worst = max(worst, diff)
        lines.append(f"  {key:<14} model {m:>16,.6f}  python {x[key]:>16,.6f}  diff {diff:.2e}")
    return worst, lines


def dump(v):
    lo = R["h_irr"]
    hi = R["k_allin"]
    out = ["label | unit | Iváncsa | Dunaújváros"]
    ws = v["Bid_Cash"]

    def fmt(x, unit):
        if isinstance(x, datetime):
            return x.strftime("%d-%b-%Y")
        if isinstance(x, (int, float)) and not isinstance(x, bool):
            if unit and unit.startswith("%"):
                return f"{x:.2%}"
            return f"{x:,.4f}" if abs(x) < 10 and x != int(x) else f"{x:,.1f}"
        return "" if x is None else str(x)
    for r in list(range(FIRST, R["r_mva"] + 1)) + list(range(lo, hi + 1)):
        lab = ws[f"C{r}"].value
        if lab is None:
            continue
        cells = [ws[f"{c}{r}"].value for c in "DEF"]
        if ws[f"C{r}"].fill.fgColor.rgb in ("FF00563C", "FFE6F4EC") or all(x is None for x in cells):
            out.append(f"## {lab}")
            continue
        out.append(" | ".join([lab, cells[0] or ""] + [fmt(x, cells[0]) for x in cells[1:]]))
    DUMP.write_text("\n".join(out) + "\n")

# =====================================================================================================================
# §22: out/HU_Wind_bids.xlsx (Dashboard + Inputs)
BIDS = HERE / "out" / "HU_Wind_bids.xlsx"
BIDS_RC = HERE / "out" / "HU_Wind_bids_recalc.xlsx"
BIDS_DUMP = HERE / "out" / "dump_bids.txt"
BID_CASH = HERE.parents[1] / "reference" / "HU_Wind_bid_deposits_fees_20261027.xlsx"


def yf360(a, b):
    """Excel YEARFRAC basis 0 (US 30/360)."""
    feb_end = lambda t: t.month == 2 and (t + timedelta(days=1)).month == 3  # noqa: E731
    d1, d2 = a.day, b.day
    if feb_end(a) and feb_end(b):
        d2 = 30
    if feb_end(a):
        d1 = 30
    if d2 == 31 and d1 >= 30:
        d2 = 30
    if d1 == 31:
        d1 = 30
    return ((b.year - a.year) * 360 + (b.month - a.month) * 30 + d2 - d1) / 360


def bands():
    """Band tables straight from the user's Bid_Cash J3:V28 (not from the bids workbook)."""
    b = openpyxl.load_workbook(BID_CASH)["Bid_Cash"]
    out = {}
    for r in range(3, 29):
        vals = [b.cell(r, c).value for c in range(11, 23) if b.cell(r, c).value is not None]
        out[b[f"J{r}"].value] = vals if len(vals) > 1 else vals[0]
    return out


def band_pts(v, t, n):
    idx = sum((x < v) + (x == v) * inc for x, inc in zip(t[f"{n}_LB"], t[f"{n}_Incl"]))
    return t[f"{n}_Pts"][idx]


def recompute_bid(dash, inp, RB, IB, col):
    k = lambda key: dash[f"{col}{RB[key]}"].value      # key inputs (Dashboard)  # noqa: E731
    pi = lambda key: inp[f"{col}{IB[key]}"].value      # per-project inputs  # noqa: E731
    c = lambda key: inp[f"E{IB[key]}"].value           # common inputs  # noqa: E731
    t = bands()
    mva, ob, bsh, hrs, sf = (k(x) for x in ("Dash_MVA", "Dash_Overbuild", "Dash_BESS_Share", "Dash_Hours",
                                            "Dash_Shortfall"))
    route = k("Dash_Turbine")
    pts = (band_pts(k("Dash_Municipal"), t, "Osz1") * t["Osz1_W"] + band_pts(k("Dash_Community"), t, "Osz2")
           * t["Osz2_W"] + band_pts(ob, t, "Osz3") * t["Osz3_W"] + band_pts(bsh, t, "Osz3") * t["Osz4_W"]
           + k("Dash_Partial") * t["Osz5_Pts"] * t["Osz5_W"] + band_pts(k("Dash_PB_Mult"), t, "Osz6") * t["Osz6_W"]
           + k("Dash_Brownfield") * t["Osz7_Pts"] * t["Osz7_W"] + k("Dash_Shared") * t["Osz8_Pts"] * t["Osz8_W"]
           + (route == "EU") * t["Osz9a_Pts"] * t["Osz9_W"] + pi("trafo") * t["Osz9b_Pts"] * t["Osz9_W"])
    fxr = c("FX")
    wind, mwh = mva * (1 + ob), mva * bsh * hrs
    base = (wind * c("Bench_Wind") + mwh * c("Bench_BESS")) / fxr / 1000
    amt = {"fee": c("Part_Fee") / fxr / 1000, "bs": c("Bid_Sec_Rate") * base}
    amt["pb"] = c("PB_Rate") * k("Dash_PB_Mult") * base - amt["bs"]
    amt["tso"] = c("TSO_Rate") * mva / fxr / 1000
    amt["dir"] = pi("dir") / pi("share_pt") / fxr / 1000
    amt["ind"] = pi("ind") * mva / fxr / 1000 - amt["tso"]
    amt["dev"], sdd = pi("devex"), pi("sdd")
    dt_ = {n: d(pi(x)) for n, x in (("fee", "d_bid"), ("bs", "d_bid"), ("pb", "d_pb"), ("tso", "d_tso"),
                                     ("dir", "d_hc"), ("ind", "d_hc"), ("dev", "d_dev"))}
    end = min(d(c("Freeze_Date")), d(c("Sale_Date")))
    sale = d(c("Sale_Date"))

    def factor(p, r):                   # simple interest within a year, capitalised 31-Dec, act/365 both ends
        bal, y = 1.0, p.year
        start = p
        while True:
            stop = min(date(y, 12, 31), end)
            full = start == date(y, 1, 1) and stop == date(y, 12, 31)      # a whole calendar year: (1 + r)
            bal *= 1 + r * (1 if full else ((stop - start).days + 1) / 365)
            if stop == end:
                return bal
            y += 1
            start = date(y, 1, 1)
    claims = {}
    for n in amt:
        r = c("Rate_Devex") if n == "dev" else c("Rate_Cash")
        acc = amt[n] * (factor(dt_[n], r) - 1)
        claims[n] = max(acc, amt[n] * (c("Devex_Floor") - 1)) if n == "dev" else acc
    dsa = k("Dash_DSA") * min(c("DSA_Rate") * yf360(d(pi("d_dsa")), sale), c("DSA_Cap")) * wind
    reimb = sum(amt[n] for n in ("bs", "pb", "tso", "dir", "ind"))
    cash = k("Dash_Premium") * mva + reimb
    paid = {}
    for n, due in (("sc", c("Sale_Costs")), ("rp", reimb), ("c12", sum(v for x, v in claims.items() if x != "dev")),
                   ("dev", claims["dev"]), ("dsa", dsa)):
        paid[n] = min(cash, due)
        cash -= paid[n]
    receipts = paid["rp"] + paid["c12"] + paid["dev"] + cash * k("Dash_FEP_Share")
    flows = [(d(pi("d_sign")), -sdd)] + [(dt_[n], -a) for n, a in amt.items()] + [(sale, receipts)]
    profit = sum(x for _, x in flows)
    out = dict(claims=sum(claims.values()), dsa=dsa, points=pts, cash=sum(amt.values()) + sdd, rtb_profit=profit,
               rtb_post=profit - max(0, profit) * c("CIT_Rate"), rtb_irr=xirr(sorted(flows)),
               rtb_moic=sum(x for _, x in flows if x > 0) / -sum(x for _, x in flows if x < 0))
    # Build and hold, annual flows at 31-Dec.
    built = wind * (1 - sf)
    turb = c("Turbine_EU") if route == "EU" else c("Turbine_CN")
    per_mw = (turb + sum(c(x) for x in ("Capex_Install", "Capex_Fdn", "Capex_Civil", "Capex_Elec", "Capex_Owner",
                                         "Capex_Land"))) * built
    capex = (per_mw + c("Subs_Lump") * mva / c("Lump_Basis")) * (1 + c("Capex_Cont")) + c("BESS_Capex") * mwh
    ntp, cod, rel = d(pi("d_ntp")), d(pi("d_cod")), d(pi("d_rel"))
    nc = (cod.year - ntp.year) * 12 + cod.month - ntp.month

    def months(start, n):
        cnt, y, m = {}, start.year, start.month
        for _ in range(n):
            m += 1
            y, m = (y + 1, 1) if m > 12 else (y, m)
            cnt[y] = cnt.get(y, 0) + 1
        return cnt
    cons, oper = months(ntp, nc), months(cod, int(c("Life")) * 12)
    outl = flows[:-1]
    fixed = sum(c(x) for x in ("Opex_OM", "Opex_BoP", "Opex_Lease", "Opex_Ins", "Opex_Mgmt", "Opex_PropTax"))
    years = range(c("First_Year"), c("First_Year") + 36)
    inv_y, eb_y, gen_y = {}, {}, {}
    for y in years:
        n = y - cod.year
        op = oper.get(y, 0) / 12
        inv = sum(x for t_, x in outl if t_.year == y) + (amt["bs"] + amt["pb"]) * (y == rel.year)
        inv += -capex * cons.get(y, 0) / nc - k("Dash_Municipal") * mva * (y == cod.year)
        gen = built * c("Hours_P50") * (1 - c("Curtailment")) * (1 - c("Clipping")) * op \
            * (1 - c("Degradation")) ** max(n - 1, 0) / 1000
        rev = gen * c("PPA_Price") * (1 + c("PPA_Esc")) ** n
        ix = (1 + c("Opex_Esc")) ** n
        base_opex = (fixed * built * op + c("Opex_Admin") * mva / c("Lump_Basis") * op
                     + gen * (c("Balancing") + c("Grid_Charges"))) * ix + rev * c("MEKH_Fee")
        opex = base_opex * (1 + c("Opex_Cont")) + gen * k("Dash_Community") + rev * c("HIPA_Rate")
        inv_y[y], eb_y[y], gen_y[y] = inv, rev - opex, gen
    ye = {y: date(y, 12, 31) for y in years}
    htm = [(ye[y], inv_y[y] + eb_y[y]) for y in years]
    price = sum(eb_y[y] * (1 + c("Discount_Rate")) ** ((cod - ye[y]).days / 365) for y in years)
    codf = [(ye[y], inv_y[y] * (y <= cod.year) + (price + amt["bs"] + amt["pb"]) * (y == cod.year)) for y in years]
    cprofit = sum(x for _, x in codf)
    allin = capex + amt["tso"] + amt["dir"] + amt["ind"] + k("Dash_Municipal") * mva + amt["fee"] + amt["dev"] + sdd
    out.update(built=built, allin=allin / 1000, ebitda1=eb_y[cod.year + 1], gwh1=gen_y[cod.year + 1],
               htm_irr=xirr(htm), htm_npv=xnpv(c("Discount_Rate"), htm), price=price / 1000,
               cod_profit=cprofit, cod_post=cprofit - max(0, cprofit) * c("CIT_Rate"), cod_irr=xirr(codf),
               cod_moic=sum(x for _, x in codf if x > 0) / -sum(x for _, x in codf if x < 0))
    return out


EX = {"claims": "c_tot", "dsa": "c_dsa", "points": "ex_pts", "cash": "ex_cash", "rtb_profit": "ex_rp", "rtb_post": "ex_rpost", "rtb_irr": "ex_rirr",
      "rtb_moic": "ex_rmoic", "built": "ex_built", "allin": "ex_allin", "ebitda1": "ex_eb", "gwh1": "ex_gwh",
      "htm_irr": "ex_irr", "htm_npv": "ex_npv", "price": "ex_price", "cod_profit": "ex_xp", "cod_post": "ex_xpost",
      "cod_irr": "ex_xirr", "cod_moic": "ex_xmoic"}
S21 = {"claims": "c_tot", "rtb_profit": "r_profit", "rtb_irr": "r_irr", "rtb_moic": "r_moic", "built": "k_built", "ebitda1": "h_ebitda",
       "gwh1": "h_gwh", "htm_irr": "h_irr", "htm_npv": "h_npv", "cod_profit": "x_profit", "cod_irr": "x_irr",
       "cod_moic": "x_moic"}


def lint_bids(RB):
    """Workbook lint, plus the calc-sheet rules on the Dashboard: one formula per timeline row (J:AS) and direction
    (up/left; within the timeline block also no reference to a later column). Exempt: the executive values and the
    master-check link at the top (summary panel)."""
    found = [str(f) for f in lint(BIDS)]
    wb = openpyxl.load_workbook(BIDS)
    ws, names = wb["Dashboard"], _name_targets(wb)
    lo, hi = RB["exec"]
    t0 = RB["t_yr"]
    tl = range(10, 46)
    for row in ws.iter_rows():
        forms = {}
        for cell in row:
            v = cell.value
            if not (isinstance(v, str) and v.startswith("=")):
                continue
            if cell.row >= t0 and cell.column in tl:
                forms[cell.column] = to_r1c1(v, cell.row, cell.column)
            if lo <= cell.row <= hi or cell.row == 3:
                continue
            for op in [t.value for t in Tokenizer(v).items if t.type == "OPERAND" and t.subtype == "RANGE"]:
                sh, addr = (op.rsplit("!", 1) if "!" in op else (None, op))
                if sh is None and op in names:
                    sh, addr = names[op]
                if (sh or "Dashboard").strip("'") != "Dashboard":
                    continue
                try:
                    c1, r1, c2, r2 = range_boundaries(addr.replace("$", ""))
                except ValueError:
                    continue
                in_tl = cell.row >= t0 and r2 >= t0 and cell.column in tl
                if r2 > cell.row or (in_tl and c2 > cell.column) or (r2 == cell.row and cell.column in tl
                                                                     and c2 >= cell.column):
                    found.append(f"[d-direction] Dashboard!{cell.coordinate}: refers right/down to {op}")
        if forms and (len(set(forms.values())) > 1 or len(forms) != len(tl)):
            found.append(f"[c-row] Dashboard!row {row[0].row}: timeline row not uniform or not filled")
    return found


def bids():
    with Office(BIDS) as lo:
        lo.calculate()
        lo.save(BIDS_RC)
    rb = json.loads((HERE / "out" / "rows_bids.json").read_text())
    RB, IB = rb["R"], rb["I"]
    vals = openpyxl.load_workbook(BIDS_RC, data_only=True)
    dash, inp = vals["Dashboard"], vals["Inputs"]
    errs = lint_values(BIDS_RC)
    print("\n§22 bids workbook")
    print("recalc:", "no errors, Master_Check", dash[f"H{RB['master']}"].value if not errs else errs[:10])
    lf = lint_bids(RB)
    print("lint:", "clean" if not lf else f"{len(lf)} finding(s)")
    for f in lf[:20]:
        print("  ", f)
    worst = 0.0
    for col, nm in zip("EFG", ("Bora", "Iváncsa", "Dunaújváros")):
        x = recompute_bid(dash, inp, RB, IB, col)
        for key, ex in EX.items():
            diff = abs(dash[f"{col}{RB[ex]}"].value - x[key])
            worst = max(worst, diff)
            if diff > 1e-6:
                print(f"   MISMATCH {nm} {key}: model {dash[f'{col}{RB[ex]}'].value} python {x[key]}")
        print(f"  {nm:<12} points {x['points']:.1f} | RtB profit {x['rtb_profit']:,.1f} post {x['rtb_post']:,.1f} "
              f"IRR {x['rtb_irr']:.2%} | HTM IRR {x['htm_irr']:.2%} | COD price {x['price']:,.3f} m")
    print(f"independent recomputation (3 projects): max abs diff {worst:.2e} ->", "PASS" if worst < 1e-6 else "FAIL")
    s21 = openpyxl.load_workbook(RC, data_only=True)["Bid_Cash"]
    tie = 0.0
    for col, c21 in (("F", "E"), ("G", "F")):
        for key, k21 in S21.items():
            tie = max(tie, abs(dash[f"{col}{RB[EX[key]]}"].value - s21[f"{c21}{R[k21]}"].value))
        tie = max(tie, abs(dash[f"{col}{RB['ex_price']}"].value * 1000 - s21[f"{c21}{R['x_price']}"].value),
                  abs(dash[f"{col}{RB['ex_allin']}"].value * 1000 - s21[f"{c21}{R['k_allin']}"].value))
    print(f"Iváncsa / Dunaújváros vs the §21 build: max abs diff {tie:.2e} ->", "PASS" if tie < 1e-6 else "FAIL")
    dump_bids(vals, RB)
    print("dump:", BIDS_DUMP)
    return not errs and not lf and worst < 1e-6 and tie < 1e-6


def dump_bids(vals, RB):
    def fmt(x, unit):
        if isinstance(x, datetime):
            return x.strftime("%d-%b-%Y")
        if isinstance(x, (int, float)) and not isinstance(x, bool):
            if unit and unit.startswith("%"):
                return f"{x:.2%}"
            return f"{x:,.4f}" if abs(x) < 10 and x != int(x) else f"{x:,.1f}"
        return "" if x is None else str(x)
    out = []
    for sheet, cols, last in (("Dashboard", "DEFGH", RB["t_yr"] - 1), ("Inputs", "DEFG", None)):
        ws = vals[sheet]
        out.append(f"#### {sheet}: label | unit | Bora | Iváncsa | Dunaújváros" + (" | Total" if sheet == "Dashboard"
                                                                                  else ""))
        rows = list(range(1, (last or ws.max_row) + 1))
        if sheet == "Dashboard":
            rows += list(range(RB["x_cit"], RB["master"] + 1))
        for r in rows:
            lab = ws[f"C{r}"].value
            if lab is None or lab == "***END***":
                continue
            cells = [ws[f"{c}{r}"].value for c in cols]
            if ws[f"C{r}"].fill.fgColor.rgb in ("FF00563C", "FFE6F4EC"):
                out.append(f"## {lab}")
                continue
            out.append(" | ".join([lab, cells[0] or ""] + [fmt(x, cells[0]) for x in cells[1:]]))
    BIDS_DUMP.write_text("\n".join(out) + "\n")



def main():
    recalc()
    vals = openpyxl.load_workbook(RC, data_only=True)
    v = vals["Bid_Cash"]
    allerr = lint_values(RC)
    errs = [f for f in allerr if f.rule != "f-error" or _row(f.where) >= FIRST]
    print("recalc (new blocks):", "no errors, Master_Check", v[f"G{R['master']}"].value if not errs else errs[:10])
    print("recalc (user's rows, untouched; LibreOffice has no _xlfn.SINGLE, Excel computes them):",
          [f"{f.where} {f.detail}" for f in allerr if f not in errs])
    new, old = lint_new()
    print("lint (new blocks):", "clean" if not new else f"{len(new)} finding(s)")
    for f in new[:30]:
        print("  ", f)
    print("lint (user's rows, untouched):", [str(f) for f in old])
    x = recompute(v)
    worst, lines = compare(v, x)
    print("\n".join(lines))
    print(f"independent recomputation (Iváncsa): max abs diff {worst:.2e} ->", "PASS" if worst < 1e-6 else "FAIL")
    dump(vals)
    print("dump:", DUMP)
    ok = bids()
    return 0 if (not errs and not new and worst < 1e-6 and ok) else 1


if __name__ == "__main__":
    sys.exit(main())
