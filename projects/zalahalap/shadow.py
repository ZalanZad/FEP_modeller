"""Independent Python shadow of the Zalahaláp workbook (SPEC v2): one case, one sensitivity case, annual 2025–2064.

Coded from the SPEC (§4–§6) and the input values, not from the workbook's formulas; reconcile.py compares it with the
recalculated workbook (machine precision) and uses its switches for the last steps of the v3 → v2 bridge.

    run("PV + 4h BESS", "Base")                  → dict of headline results
    run("PV + 4h BESS", "Base", ppa_years=0)     → overrides (any key of params())
Stdlib only.
"""
from __future__ import annotations

import calendar
import csv
import datetime as dt
from pathlib import Path

HERE = Path(__file__).resolve().parent
D = dt.date
YEARS = list(range(2025, 2065))
YE = [D(y, 12, 31) for y in YEARS]
SENS = {"Base": dict(cpv=0.0, cb=0.0, cy=0.0, cc=0.0, delay=0, buyer=0.08),
        "Low": dict(cpv=-0.2, cb=-0.2, cy=-0.02, cc=0.10, delay=12, buyer=0.09),
        "High": dict(cpv=0.2, cb=0.2, cy=0.02, cc=-0.05, delay=0, buyer=0.07)}
CASES = {"PV only": dict(built=0, mw=0, hours=0), "PV + 2h BESS": dict(built=1, mw=50, hours=2),
         "PV + 4h BESS": dict(built=1, mw=50, hours=4)}
BLOCK = {2: "2027 solar DC 1.0cy 2h", 4: "2027 solar DC 1.0cy 4h"}


def params(**over):
    p = dict(mwp=63, mva=49.9, yld=1265, ppa_price=45, ppa_index=0.0, ppa_years=5, basis="Merchant", rtb_price=90,
             cod_price=0, acq=D(2025, 11, 30), rtb=D(2026, 12, 31), exit=D(2027, 3, 31), ntp=D(2027, 6, 30),
             cod=D(2028, 3, 31), life=35, avail=0.995, deg=0.003, clip=0.0, goo=1.75, pv_only_hours=2, rte=0.86,
             haircut=0.0, afrr=0.0, bess_capex=110, repower=1, bess_opex=30.8, balancing=5.0, toll_start=D(2028, 1, 1),
             toll_years=3, toll_fee=49.7, pv_capex=350, cable=3700, subst=9500, grid_fee=0, land=620, ew_ntp=725,
             om=14.96, cable_om=5, lease=96, hipa=0.02, innov=0.003, spv=549, trans=200, dsa=15,
             dsa_first=D(2025, 12, 31), pre_total=3182, pre_first=D(2026, 10, 31), pre_last=D(2027, 3, 31), fx=400,
             bg=0.0187 + 0.0175, rate=0.12, share_low=0.6, share_high=0.5, threshold=90, psh_land=150, psh_devex=15,
             ret=dict(spv=0, trans=0, dep=1, dsa=1, ew=1, pre=0, fee=0), reimb=dict(pre=1, ew=1, dep=1),
             cit=0.09, rh=0, rh_rate=0.31, dep_years=30, dep_bess=10, loss_cap=0.5, loss_years=5, gain_taxed=0,
             disc=0.08, sale_rtb=0.0, sale_cod=0.0,
             early_works=[(60, "HUF", D(2025, 11, 30)), (60, "HUF", D(2026, 2, 28)), (25, "EUR", D(2026, 2, 28)),
                          (460, "EUR", D(2026, 4, 30))],
             deposits=[(45, D(2025, 11, 30), None), (180, D(2025, 11, 30), None),
                       (485.55, D(2025, 11, 30), D(2026, 1, 31))],
             guarantees=[(242.775, D(2026, 1, 31), D(2026, 3, 31)), (242.775, D(2026, 1, 31), None),
                         (80.925, D(2026, 4, 30), None)],
             dsa_lines=13, pre_lines=6)
    p.update(over)
    return p


# ------------------------------------------------------------------------------------------- helpers
def eomonth(d, n):
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    return D(y, m, calendar.monthrange(y, m)[1])


def months_between(a, b):
    """DATEDIF(a, b, "m")."""
    n = (b.year - a.year) * 12 + b.month - a.month
    return n - (1 if b.day < a.day else 0)


def xnpv(rate, flows, dates):
    t0 = dates[0]
    return sum(f / (1 + rate) ** ((t - t0).days / 365.0) for f, t in zip(flows, dates))


def xirr(flows, dates, lo=-0.99, hi=10.0):
    f = lambda r: xnpv(r, flows, dates)        # noqa: E731
    flo, fhi = f(lo), f(hi)
    if flo * fhi > 0:
        return None
    for _ in range(300):
        mid = (lo + hi) / 2
        fm = f(mid)
        if (fm > 0) == (flo > 0):
            lo, flo = mid, fm
        else:
            hi = mid
        if hi - lo < 1e-15:
            break
    return (lo + hi) / 2


def aurora():
    out = {}
    for r in csv.DictReader(open(HERE / "data" / "aurora_curves.csv")):
        vals = [float(r[str(y)]) for y in range(2026, 2061)]
        out[(r["series"], r["block"])] = dict(zip(range(2026, 2061), vals))
    return out


AU = aurora()


def series(name, block, sign=1):
    s = AU[(name, block)]
    return {y: (0.0 if y == 2025 else sign * s[min(y, 2060)]) for y in YEARS}


def euro_index():
    s = AU[("euro_index", "Central")]
    out = {2025: 1.0}
    for y in YEARS[1:]:
        out[y] = s[y] if y <= 2060 else out[y - 1] * 1.02
    return out


IDX = euro_index()


def factor(d, end, r):
    """Closed-form act/365 accrual with 31-Dec capitalisation (wind E169)."""
    days1 = (D(d.year + 1, 1, 1) - d).days
    stub = 1 + r * ((end - D(end.year, 1, 1)).days + 1) / 365
    return (1 + r * days1 / 365) * (1 + r) ** (end.year - d.year - 1) * stub


# ------------------------------------------------------------------------------------------- model
def run(case="PV + 4h BESS", sens="Base", **over):
    p = params(**over)
    c, s = CASES[case], SENS[sens]
    built, mw = c["built"], c["mw"] * c["built"]
    hours = c["hours"]
    mwh = built * c["mw"] * hours
    mwp = p["mwp"]
    codl = eomonth(p["cod"], s["delay"])
    blk = BLOCK[hours] if built else BLOCK[p["pv_only_hours"]]
    rep_cnt = AU[("repower_count", blk)]
    rep_year = next(y for y in range(2026, 2061) if rep_cnt[y] >= 1)
    pxd_m, pxs_m = series("direct_price", blk), series("transfer_price", blk)
    st, pvo, dcs = series("storage_total", blk), series("pv_only_factor", blk), series("down_cap_share", blk)
    chg, dmw = series("charge_res", blk, -1), series("discharge", blk)
    L12 = p["life"] * 12
    e_c = months_between(p["ntp"], codl)
    R = {k: [] for k in ("gen", "sent", "direct", "dis", "rev", "rpv", "rbess", "rgoo", "opex", "ebitda", "capt",
                         "capex", "ewntp", "aug", "tax", "lapse", "oshare", "yrs", "ppa", "pxd", "bal")}
    first_full = codl.year + 1
    for y in YEARS:
        yrs = y - codl.year
        cmon = (y + 1 - p["ntp"].year) * 12 - p["ntp"].month
        cshare = max(0, min(cmon, e_c) - max(cmon - 12, 0)) / e_c
        omon = (y + 1 - codl.year) * 12 - codl.month
        osh = max(0, min(omon, L12) - max(omon - 12, 0)) / 12
        degf = (1 - p["deg"]) ** max(yrs - 1, 0)
        live = 1 - (1 - p["repower"]) * (y >= rep_year)
        repf = p["repower"] * (y == rep_year)
        idx = IDX[y]
        gen = mwp * p["yld"] * (1 + s["cy"]) * (1 - p["clip"]) * p["avail"] * degf * osh
        sent = min(gen, chg[y] * mw * osh * live)
        direct = gen - sent
        dis = min(dmw[y] * mw * osh * live, sent * p["rte"])
        # prices (Revenue bar 3)
        m0 = (y - codl.year) * 12 - codl.month
        opm = max(0, min(12, L12 - m0) - max(0, -m0))
        t0 = (y - p["toll_start"].year) * 12 - p["toll_start"].month + 1
        md = pxd_m[y] * idx * (1 + s["cpv"])
        ms = pxs_m[y] * idx * (1 + s["cpv"])
        ppa = max(0, min(12, p["ppa_years"] * 12 - m0) - max(0, -m0)) / max(1, opm)
        ppx = p["ppa_price"] * (1 + p["ppa_index"]) ** (y - codl.year)
        pxd, pxs = md + ppa * (ppx - md), ms + ppa * (ppx - ms)
        upl = built * (1 - p["haircut"]) * (1 + s["cb"]) * st[y] * (pvo[y] + dcs[y] * p["afrr"]) * idx
        tol = (p["basis"] == "Tolling") * max(0, min(12, p["toll_years"] * 12 - t0) - max(0, -t0)) / 12
        bess_px = upl + tol * (p["toll_fee"] - upl)
        rpv = (direct * pxd + sent * pxs) / 1000
        rgoo = gen * p["goo"] * idx / 1000
        rbess = mw * bess_px * osh * live
        rev = rpv + rgoo + rbess
        opex = (-(p["om"] + p["cable_om"]) * mwp * idx * osh - gen * p["balancing"] * idx * (1 - ppa) / 1000
                - p["lease"] * idx * osh - p["bess_opex"] * mw * idx * osh * live - rev * (p["hipa"] + p["innov"]))
        for k, v in (("gen", gen), ("sent", sent), ("direct", direct), ("dis", dis), ("rev", rev), ("rpv", rpv),
                     ("rbess", rbess), ("rgoo", rgoo), ("opex", opex), ("ebitda", rev + opex), ("oshare", osh),
                     ("yrs", yrs), ("ppa", ppa), ("pxd", pxd)):
            R[k].append(v)
        R.setdefault("cshare", []).append(cshare)
        R.setdefault("omon", []).append(omon)
        R.setdefault("repf", []).append(repf)
        R.setdefault("idx", []).append(idx)
    # capex to COD
    cc = 1 + s["cc"]
    k_bess = p["bess_capex"] * mwh * cc
    k_land = p["land"] * cc
    k_ew = p["ew_ntp"] * cc
    k_tot = (p["pv_capex"] * mwp + p["cable"] + p["subst"] + p["grid_fee"] + p["land"]) * cc + k_bess + k_ew
    for i, y in enumerate(YEARS):
        capex = -(k_tot - k_ew) * R["cshare"][i]
        ewn = -k_ew * (y == p["ntp"].year)
        aug = -mwh * p["bess_capex"] * R["idx"][i] * R["repf"][i]
        R["capex"].append(capex)
        R["ewntp"].append(ewn)
        R["aug"].append(aug)
        R["capt"].append(capex + ewn + aug)
    # dated list (RtB route) and the pre-COD outlays
    lines = []                                                   # (amount, date, earns)
    fxk = 1000 / p["fx"]
    lines.append((-p["spv"], p["acq"], p["ret"]["spv"]))
    lines.append((-p["trans"], p["acq"], p["ret"]["trans"]))
    for huf, paid, rel in p["deposits"]:
        lines.append((-huf * fxk, paid, p["ret"]["dep"]))
    for huf, paid, rel in p["deposits"]:
        lines.append((huf * fxk if rel else 0.0, rel or paid, p["ret"]["dep"]))
    n_dsa = (p["rtb"].year - p["dsa_first"].year) * 12 + p["rtb"].month - p["dsa_first"].month + 1
    d = p["dsa_first"]
    for k in range(p["dsa_lines"]):
        if k:
            d = eomonth(d, 1)
        lines.append((-p["dsa"] * mwp / n_dsa * (d <= p["rtb"]), d, p["ret"]["dsa"]))
    for amt, cur, date in p["early_works"]:
        lines.append((-(amt * fxk if cur == "HUF" else amt), date, p["ret"]["ew"]))
    n_pre = (p["pre_last"].year - p["pre_first"].year) * 12 + p["pre_last"].month - p["pre_first"].month + 1
    d = p["pre_first"]
    for k in range(p["pre_lines"]):
        if k:
            d = eomonth(d, 1)
        lines.append((-p["pre_total"] / n_pre * (d <= p["pre_last"]), d, p["ret"]["pre"]))
    for huf, issued, rel in p["guarantees"]:
        end = rel or p["exit"]
        lines.append((-huf * fxk * p["bg"] * (end - issued).days / 365, end, p["ret"]["fee"]))
    dsa_total = -sum(a for a, d_, _ in lines[2 + 2 * len(p["deposits"]):2 + 2 * len(p["deposits"]) + p["dsa_lines"]])
    ew_total = sum(amt * fxk if cur == "HUF" else amt for amt, cur, _ in p["early_works"])
    pre_spent = p["pre_total"] / n_pre * min(n_pre, p["pre_lines"])
    dep_out = sum(huf for huf, _, rel in p["deposits"] if rel is None) * fxk
    bg_out = sum(huf for huf, _, rel in p["guarantees"] if rel is None) * fxk
    hold_at_cod = dep_out - bg_out * p["bg"] * (codl - p["exit"]).days / 365
    buyer_reimb = p["reimb"]["pre"] * pre_spent + p["reimb"]["ew"] * ew_total + p["reimb"]["dep"] * dep_out
    ret_x = sum(-a * (factor(d_, p["exit"], p["rate"]) - 1) * f for a, d_, f in lines)
    ret_c = sum(-a * (factor(d_, codl, p["rate"]) - 1) * f for a, d_, f in lines)
    codf = [1 if y == codl.year else 0 for y in YEARS]
    pre = [sum(a for a, d_, _ in lines if d_.year == y) + cf * hold_at_cod for y, cf in zip(YEARS, codf)]
    # tax
    dep_base_b = k_bess
    dep_base = k_tot - k_bess - k_land + ew_total + pre_spent
    aug_total = sum(R["aug"])
    lo = 0.0
    lapse_year = p["ntp"].year + p["loss_years"]
    for i, y in enumerate(YEARS):
        om = R["omon"][i]
        db = -dep_base_b * max(0, min(om, p["dep_bess"] * 12) - max(om - 12, 0)) / (p["dep_bess"] * 12)
        dp = -dep_base * max(0, min(om, p["dep_years"] * 12) - max(om - 12, 0)) / (p["dep_years"] * 12)
        da = aug_total / p["dep_bess"] * (y >= rep_year) * (y < rep_year + p["dep_bess"])
        tb = R["ebitda"][i] + db + dp + da
        add = max(0, -tb)
        use = -min(lo, p["loss_cap"] * max(0, tb))
        lapse = -(lo + add + use) * (y == lapse_year)
        lo = lo + add + use + lapse
        R["lapse"].append(lapse)
        R["tax"].append(-p["cit"] * (max(0, tb) + use) - p["rh"] * p["rh_rate"] * max(0, tb))
    n = len(YEARS)
    pcf = [R["ebitda"][i] + R["capt"][i] + pre[i] for i in range(n)]
    pcft = [pcf[i] + R["tax"][i] for i in range(n)]
    # buyer at RtB: implied premium and IRR at our price
    bx = [R["ebitda"][i] + R["capt"][i] + R["tax"][i] + codf[i] * hold_at_cod for i in range(n)]
    ev_x = sum(f * (1 + s["buyer"]) ** ((p["exit"] - t).days / 365) for f, t in zip(bx, YE))
    implied = max(0.0, (ev_x - buyer_reimb) / mwp)
    price = p["rtb_price"] if p["rtb_price"] > 0 else implied
    allin = price + buyer_reimb / mwp
    bxd = [max(t, p["exit"]) for t in YE]
    bxp = [bx[i] - allin * mwp * (YEARS[i] == 2025) for i in range(n)]
    buyer_rtb = xirr(bxp, bxd)
    buyer_rtb_pre = xirr([bxp[i] - R["tax"][i] for i in range(n)], bxd)
    # JV split at the RtB exit
    cons = price * mwp
    avail = cons + buyer_reimb
    sale = -min(avail, p["sale_rtb"] * cons)
    prin = -min(dsa_total + buyer_reimb, avail + sale)
    ret = -min(ret_x, avail + sale + prin)
    spv = -min(p["spv"], avail + sale + prin + ret)
    dist = avail + sale + prin + ret + spv
    share = p["share_low"] if price <= p["threshold"] else p["share_high"]
    fl = dist * share
    pl = dist - fl
    fep_x = -(prin + ret + spv) + fl
    dl = [a for a, _, _ in lines] + [fep_x]
    dd = [d_ for _, d_, _ in lines] + [p["exit"]]
    rtb_pre = sum(dl)
    rtb_tax = -p["cit"] * (-ret + p["gain_taxed"] * max(0, rtb_pre + ret))
    rtb = dict(profit_pre=rtb_pre, profit=rtb_pre + rtb_tax,
               moic=sum(x for x in dl if x > 0) / -sum(x for x in dl if x < 0),
               irr_pre=xirr(dl, dd), irr=xirr(dl + [rtb_tax], dd + [p["exit"]]),
               npv=xnpv(p["disc"], dl + [rtb_tax], dd + [p["exit"]]), proceeds=fep_x,
               psh_profit=pl + (p["dsa"] - p["psh_devex"]) * mwp - p["psh_land"], ret=-ret, tax=rtb_tax)
    # COD sale
    codcol = YEARS.index(codl.year)
    bc = [(R["ebitda"][i] + R["tax"][i] + R["aug"][i]) * (R["yrs"][i] >= 0) for i in range(n)]
    ev_c = sum(f * (1 + s["buyer"]) ** ((YE[codcol] - t).days / 365) for f, t in zip(bc, YE))
    implied_c = max(0.0, ev_c / mwp)
    price_c = p["cod_price"] if p["cod_price"] > 0 else implied_c
    ev = price_c * mwp
    contrib = -(sum(pre) + sum(x for x, yr in zip(R["capt"], R["yrs"]) if yr <= 0))
    c_sale = -min(ev, p["sale_cod"] * ev)
    c_fund = -min(contrib, ev + c_sale)
    c_ret = -min(ret_c, ev + c_sale + c_fund)
    c_dist = ev + c_sale + c_fund + c_ret
    c_fl = c_dist * share
    c_pl = c_dist - c_fl
    c_fep = -(c_fund + c_ret) + c_fl
    codr = [R["capt"][i] * (R["yrs"][i] <= 0) + pre[i] + c_fep * codf[i] for i in range(n)]
    c_tax = -p["cit"] * (-c_ret + p["gain_taxed"] * max(0, sum(codr) + c_ret))
    codt = [codr[i] + c_tax * codf[i] for i in range(n)]
    bcp = [bc[i] - ev * codf[i] for i in range(n)]
    buyer_cod = xirr(bcp[codcol:], YE[codcol:])
    buyer_cod_pre = xirr([bcp[i] - R["tax"][i] * (R["yrs"][i] >= 0) for i in range(n)][codcol:], YE[codcol:])
    # hold
    hold = [pcft[i] - pl * (YEARS[i] == p["exit"].year) for i in range(n)]
    holdp = [hold[i] - R["tax"][i] for i in range(n)]
    dfa = 1 / (1 + p["disc"]) ** ((YE[0] - p["acq"]).days / 365)
    i_ffy = YEARS.index(first_full)
    return dict(
        case=case, sens=sens,
        gen_ffy=R["gen"][i_ffy], sent_ffy=R["sent"][i_ffy], dis_ffy=R["dis"][i_ffy], rev_ffy=R["rev"][i_ffy],
        ebitda_ffy=R["ebitda"][i_ffy], rev_life=sum(R["rev"]), ebitda_life=sum(R["ebitda"]), tax_life=sum(R["tax"]),
        capex_total=k_tot, cash_to_rtb=-sum(a for a, _, _ in lines), buyer_reimb=buyer_reimb,
        implied_rtb=implied, implied_raw=(ev_x - buyer_reimb) / mwp, rtb_price=price, allin=allin, buyer_rtb=buyer_rtb, buyer_rtb_pre=buyer_rtb_pre,
        rtb_profit_pre=rtb["profit_pre"], rtb_profit=rtb["profit"], rtb_moic=rtb["moic"], rtb_irr_pre=rtb["irr_pre"],
        rtb_irr=rtb["irr"], rtb_npv=rtb["npv"], rtb_proceeds=rtb["proceeds"], psh_rtb=rtb["psh_profit"],
        share=share, psh_lot=pl, ret_x=ret_x,
        implied_cod=implied_c, cod_price=price_c, ev=ev, buyer_cod=buyer_cod, buyer_cod_pre=buyer_cod_pre,
        cod_contrib=contrib, cod_proceeds=c_fep, cod_profit_pre=sum(codr), cod_profit=sum(codt),
        cod_moic=sum(x for x in codr if x > 0) / -sum(x for x in codr if x < 0),
        cod_irr=xirr(codt, YE), cod_irr_pre=xirr(codr, YE), cod_npv=xnpv(p["disc"], codt, YE) * dfa,
        psh_cod=c_pl + (p["dsa"] - p["psh_devex"]) * mwp - p["psh_land"],
        hold_equity=-sum(x for x in hold if x < 0), hold_dist=sum(x for x in hold if x > 0),
        hold_profit=sum(hold), hold_profit_pre=sum(holdp), hold_irr=xirr(hold, YE), hold_irr_pre=xirr(holdp, YE),
        hold_npv=xnpv(p["disc"], hold, YE) * dfa, proj_irr=xirr(pcft, YE), proj_irr_pre=xirr(pcf, YE),
        rows=R, pcft=pcft, hold=hold)


if __name__ == "__main__":
    for case in CASES:
        r = run(case)
        print(case, {k: round(v, 6) for k, v in r.items() if isinstance(v, float)})
