"""Python shadow of the original Zalahaláp model (Zalahalap_model_10112025_v3.xlsb), Scenario A / Negotiated case.

Reproduces the reconciliation targets of old_targets.csv from old_inputs.json and old_curves.csv:
  rtb()  Calculation rows 51-180 (monthly, acquisition 30-Nov-25 → sale month 11) + RTB_Model rows 41-78
  htm()  HTM_Model (quarterly I:FL, 01-Oct-25 → 30-Sep-65): project, buyer and Futureal hold-to-maturity flows

Switches (all off = the original model; reconcile.py turns them on one by one, SPEC §13):
  fix2_reimb      pre-RtB costs reimbursed at the later of sale and spend (RTB_Model rows 68/69)
  fix3_bg         no bank-guarantee phantom (Calculation!H118:H120 one row high)
  fix5_clip       PV export clipped: clip_loss of DC energy lost
  fix6_bess       BESS = Aurora block "Storage total cashflow" (bess_block, real k€/MW-yr by year) instead of
                  the v3 stack (0.591 scaling, 25 % network retention, PV energy at the 65 % capture)
  fix7_aug        augmentation indexed and dated at the Aurora repower year (aug_year)
  fix8_interest   simple act/365 interest (31-Dec capitalisation) instead of monthly compounding
  share_step      Futureal share by the user's step rule (60 % at or below 90 k€/MW, 50 % above)
  scope4          early works = the four Stat_inp lines (E44:E47); E48/E49 paid at construction start (NTP)
Stdlib only.
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import math
from pathlib import Path

HERE = Path(__file__).resolve().parent
J = json.loads((HERE / "old_inputs.json").read_text())
EXCEL0 = dt.date(1899, 12, 30)


def v(group, key):
    return J[group][key]["value"]


def d(s):
    return dt.date.fromisoformat(s[:10])


def eomonth(x: dt.date, n: int) -> dt.date:
    m = x.month - 1 + n
    y, m = x.year + m // 12, m % 12 + 1
    nxt = dt.date(y + (m == 12), m % 12 + 1, 1)
    return nxt - dt.timedelta(days=1)


def _num(s):
    try:
        return float(s)
    except ValueError:
        return 0.0                                   # 2025 row: '#VALUE!' (Aurora_Curve!C8 = "no "), patched to 0


CURVES = {int(r["year"]): {k: _num(x) for k, x in r.items() if k not in ("year", "source_cells")}
          for r in csv.DictReader(open(HERE / "old_curves.csv"))}


# =========================================================================================== XIRR / XNPV
def xnpv(rate, flows, dates):
    t0 = dates[0]
    return sum(f / (1 + rate) ** ((t - t0).days / 365.0) for f, t in zip(flows, dates))


def xirr(flows, dates, guess=0.1):
    """Excel XIRR: first non-zero flow is the base date; Newton from `guess` as Excel, else grid bracketing."""
    pairs = [(f, t) for f, t in zip(flows, dates)]
    while pairs and pairs[0][0] == 0:
        pairs.pop(0)
    f = [x[0] for x in pairs]
    t = [x[1] for x in pairs]
    yrs = [(x - t[0]).days / 365.0 for x in t]

    def npv(r):
        return sum(a / (1 + r) ** y for a, y in zip(f, yrs))

    def dnpv(r):
        return sum(-y * a / (1 + r) ** (y + 1) for a, y in zip(f, yrs))
    r = guess
    for _ in range(200):
        dv = dnpv(r)
        if dv == 0:
            break
        step = npv(r) / dv
        r -= step
        if r <= -0.999:
            break
        if abs(step) < 1e-12:
            return r
    grid = [-0.99 + 0.01 * k for k in range(1, 100)] + [x / 10 for x in range(1, 1000)]
    for a, b in zip(grid, grid[1:]):
        if npv(a) * npv(b) <= 0:
            lo, hi = a, b
            for _ in range(200):
                mid = (lo + hi) / 2
                if npv(lo) * npv(mid) <= 0:
                    hi = mid
                else:
                    lo = mid
            return (lo + hi) / 2
    return float("nan")


# =========================================================================================== defaults
def defaults(**over):
    p = dict(
        mw=v("capacity", "pv_mwp"), mva=v("capacity", "grid_mva"), price=v("rtb_deal", "rtb_premium"),
        upfront=v("rtb_deal", "connection_rights"), transaction=v("rtb_deal", "transaction_costs"),
        land_apport=v("rtb_deal", "psh_land_apport"), rate=v("rtb_deal", "pref_rate"),
        bg_margin=v("rtb_deal", "bg_margin"), acq=d(v("rtb_deal", "investment_date")), sale_m=11,
        euribor=J["series"]["euribor"]["values"],
        ew=[v("milestones", f"m{i}") for i in range(1, 7)], ew_m=[v("milestones", f"m{i}_month") for i in range(1, 7)],
        dev_total=v("rtb_deal", "developer_total"), pre_rtb=v("capex", "pre_rtb_total"),
        pre_first=d(v("capex", "pre_rtb_first_month")), pre_last=d(v("capex", "pre_rtb_last_month")),
        # HTM
        yield_=v("yield", "base_yield"), avail=v("yield", "availability"), deg=v("yield", "degradation"),
        cons_start=d(v("dates", "construction_start")), cons_months=v("dates", "construction_months"),
        cod=d(v("dates", "cod")), ops_end=d(v("dates", "operations_end")),
        om=v("opex", "om_per_mw"), bal=v("opex", "balancing_per_mwh"), lease=v("opex", "land_lease"),
        cable=v("opex", "mv_cable_om_per_mw"), rtm=v("opex", "route_to_market"), days_rec=v("opex", "days_receivable"),
        days_pay=v("opex", "days_payable"), bess_mw=v("bess", "mw"), bess_mwh=v("bess", "mwh"),
        bess_opex=v("bess", "fixed_opex"), aug_cost=v("bess", "augment_cost_per_mwh"), aug_year=v("bess", "augment_year"),
        capex_pv=v("capex", "pv_total_per_mwp"), cable_capex=v("capex", "mv_cable"), hv=v("capex", "hv_substation"),
        land=v("capex", "land_to_cod"), bess_capex=v("bess", "capex_per_mwh"),
        cit=v("tax", "cit"), dep_years=v("tax", "depreciation_years"), loss_cap=v("tax", "loss_cap"),
        npv_rate=v("valuation", "npv_rate"), phasing=J["series"]["capex_phasing"]["values"],
        bg_rate_after=v("htm", "bg_rate_after_sale"),
        # switches (SPEC §13)
        fix2_reimb=False, fix3_bg=False, fix5_clip=False, clip_loss=0.02, fix6_bess=False, bess_block=None,
        fix7_aug=False, aug_year_new=2048, fix8_interest=False, share_step=False, scope4=False,
    )
    p.update(over)
    return p


def share(p):
    x = p["price"]
    if p["share_step"]:
        return 0.6 if x <= 90 else 0.5
    if x < 90:
        return 0.6
    return 0.5 + (100 - x) / 100 if x <= 100 else 0.5


# =========================================================================================== RtB (monthly)
DEP_CASH = [[1, 0, -1, 0, 0, 0, 0, 0, 0, 0, 0, 0], [1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, -1]]
DEPOSITS = [112.5, 450.0, 1213.875, 1213.875, 202.3125]        # Stat_inp!E92:E96
PROFILE = {0: {0: 1, 11: -1}, 1: {0: 1, 11: -1}, 2: {0: 1, 2: -1}, 3: {2: 1, 4: -0.5, 11: -0.5}, 4: {5: 1, 11: -1}}
CASH = [1, 1, 1, 0, 0]                                         # Stat_inp!E102:E106


def rtb(p):
    n, S = 19, p["sale_m"]
    dates = [eomonth(p["acq"], m) for m in range(n)]
    dep = [[DEPOSITS[i] * PROFILE[i].get(m, 0) for m in range(n)] for i in range(5)]       # Dyn_Inp rows 12-16
    cash_dep = [sum(dep[i][m] * CASH[i] for i in range(5)) for m in range(n)]           # row 111
    # bank guarantees (rows 115-122); column H points one row high (fix 3)
    bg_add = []
    for m in range(n):
        if m == 0 and not p["fix3_bg"]:
            bg_add.append(sum((1 - CASH[i]) * dep[i - 1][0] for i in range(1, 5)))      # E118*G13 ... E120*G15
        else:
            bg_add.append(sum((1 - CASH[i]) * dep[i][m] for i in range(5)))
    bg_bal, bg_int, open_ = [], [], 0.0
    for m in range(n):
        pre = open_ + bg_add[m]
        rate = (p["euribor"][m] + p["bg_margin"]) / 12
        bg_int.append(pre * rate if m >= 1 else 0.0)                                      # H129 empty
        open_ = pre - (pre if m == S else 0.0)
        bg_bal.append(pre)
    # milestones (rows 54-60); scope4 moves lines 5-6 out of the development period
    ew_lines = [(a * p["mw"], mm) for a, mm in zip(p["ew"], p["ew_m"])]
    if p["scope4"]:
        ew_lines = ew_lines[:4]
    ew = [sum(a for a, mm in ew_lines if mm == m) for m in range(n)]
    ew_total = sum(ew)
    # Futureal interest (rows 73-87)
    upfront = [p["upfront"] if m == 0 else 0 for m in range(n)]
    base = [ew[m] + cash_dep[m] + bg_int[m] for m in range(n)]                          # rows 75-78
    r = p["rate"] / 12
    interest, cum_up, cum_b, cum_i = [], 0.0, 0.0, 0.0
    for m in range(n):
        cum_up += upfront[m]
        cum_b += base[m]
        if 1 <= m <= S:
            if p["fix8_interest"]:              # simple 12 % act/365 on principal, capitalised at 31-Dec
                days = (dates[m] - dates[m - 1]).days
                i = p["rate"] * days / 365 * (cum_up + cum_b + _capitalised(interest, dates, m))
            else:
                i = r * cum_up + r * cum_b + r * cum_i
        else:
            i = 0.0
        interest.append(i)
        cum_i += i
    F79, F129 = sum(interest), sum(bg_int)
    # waterfall (rows 150-160)
    gross = p["mw"] * p["price"]
    F151 = -min(p["upfront"], gross)
    F154 = -min(F79 + F129, gross + F151)
    psh_r = r                                   # PSH interest (rows 94-100): monthly 1 % on the land apport
    psh_i, cum = [], 0.0
    for m in range(n):
        if 1 <= m <= S:
            if p["fix8_interest"]:
                days = (dates[m] - dates[m - 1]).days
                i = p["rate"] * days / 365 * (p["land_apport"] + _capitalised(psh_i, dates, m))
            else:
                i = psh_r * p["land_apport"] + psh_r * cum
        else:
            i = 0.0
        psh_i.append(i)
        cum += i
    F100 = sum(psh_i)
    net = gross + F151 + F154 - p["land_apport"] - F100
    sh = share(p)
    fut, dev = sh * net, (1 - sh) * net
    # Futureal cash flows (rows 132-143)
    cf = [0.0] * n
    for m in range(n):
        cf[m] = -upfront[m] - (p["transaction"] if m == 0 else 0) - cash_dep[m] - ew[m] - bg_int[m]
        if m == S:
            cf[m] += -F154 - F151 + ew_total + fut
    peak_cash = max(_cum(cash_dep))
    # RTB_Model: pre-RtB development costs and their reimbursement (rows 68-69)
    pre_n = round((p["pre_last"] - p["pre_first"]).days / 30.4) + 1
    pre = [-p["pre_rtb"] / pre_n if p["pre_first"] <= dates[m] <= p["pre_last"] else 0.0 for m in range(n)]
    if p["fix2_reimb"]:
        reimb = [(-sum(pre[:S + 1]) if m == S else (-pre[m] if m > S else 0.0)) for m in range(n)]
    else:
        reimb = [p["pre_rtb"] if m == S else 0.0 for m in range(n)]
    cf_all = [a + b + c for a, b, c in zip(cf, pre, reimb)]
    invested = p["upfront"] + p["transaction"] + peak_cash + ew_total + F129 + p["pre_rtb"]
    proceeds = -F151 + ew_total - F154 + peak_cash + fut + p["pre_rtb"]
    D76 = proceeds - fut
    D77 = gross - net
    dev_inv = p["land_apport"] + p["dev_total"] * p["mw"]
    dev_proceeds = ew_total + dev + F100                       # F225: the land apport refund is not credited
    return dict(
        dates=dates, cf=cf, cf_all=cf_all, ew=ew, ew_total=ew_total, interest=F79, bg_cost=F129, bg_bal=bg_bal,
        net=net, share=sh, fut=fut, dev=dev, psh_interest=F100, gross=gross, cash_dep=cash_dep,
        rtb_xirr_incl=xirr(cf_all, dates), rtb_xirr_excl=xirr(cf, dates),
        rtb_moic_incl_dep=proceeds / invested, rtb_moic_excl_dep=(proceeds - peak_cash) / (invested - peak_cash),
        calc_moic_incl_dep=(proceeds - p["pre_rtb"]) / (invested - p["pre_rtb"]),
        calc_moic_excl_dep=(-F151 + ew_total - F154 + fut) / (p["upfront"] + p["transaction"] + ew_total + F129),
        rtb_net_profit=(proceeds - invested) / 1000, rtb_capital=invested / 1000, rtb_proceeds=proceeds / 1000,
        developer_profit=dev_proceeds - dev_inv,
        rtb_breakeven=((invested - D76) / sh + D77) / p["mw"],
        psh_buyout=dev + p["land_apport"] + F100, bg_after=bg_bal[S - 1],
        cash_at_sale=_cum(cash_dep)[S - 1] if S >= 1 else 0.0,
    )


def _cum(xs):
    out, s = [], 0.0
    for x in xs:
        s += x
        out.append(s)
    return out


def _capitalised(series, dates, m):
    """Interest capitalised at 31-Dec before month m (simple-interest fix)."""
    return sum(series[k] for k in range(m) if dates[k].year < dates[m].year)


# =========================================================================================== HTM (quarterly)
def quarters(start=dt.date(2025, 10, 1), n=160):
    out, s = [], start
    for _ in range(n):
        e = eomonth(s, 2)
        out.append((s, e))
        s = e + dt.timedelta(days=1)
    return out


def _share_in(s, e, a, b):
    return max(0, ((min(b, e) - max(a, s)).days + 1) / ((e - s).days + 1))


def htm(p, rt=None):
    rt = rt or rtb(p)
    Q = quarters()
    n = len(Q)
    ends = [e for _, e in Q]
    mw = p["mw"]
    sale = rt["dates"][p["sale_m"]]
    cons_end = eomonth(p["cons_start"], p["cons_months"] - 1)
    # construction phasing (rows 13, 58)
    def cmonth(e):
        if e < p["cons_start"]:
            return 0
        return min(p["cons_months"], (e.year - p["cons_start"].year) * 12 + e.month - p["cons_start"].month + 1)
    ph = p["phasing"]
    den = sum(x for k, x in enumerate(ph, 1) if k <= p["cons_months"])
    phase, prev = [], 0
    for s, e in Q:
        cm = cmonth(e)
        phase.append(sum(x for k, x in enumerate(ph, 1) if k <= p["cons_months"] and prev < k <= cm) / den)
        prev = cm
    ops = [_share_in(s, e, p["cod"], p["ops_end"]) for s, e in Q]
    yf = [((e - s).days + 1) / (366 if _leap(e.year) else 365) for s, e in Q]
    idx = [CURVES[e.year]["inflation_index_selected"] for e in ends]
    price = [CURVES[e.year]["pv_capture_selected_nominal"] for e in ends]
    goo = [CURVES[e.year]["goo_nominal"] for e in ends]
    cum_years, gen = 0.0, []
    for k in range(n):
        dg = (1 - p["deg"]) ** cum_years if ops[k] > 0 else 0
        g = mw * (1 if ops[k] > 0 else 0) * p["yield_"] * ops[k] * yf[k]
        g = (g - g * (1 - p["avail"])) * dg
        if p["fix5_clip"]:
            g *= 1 - p["clip_loss"]
        gen.append(g)
        cum_years += ops[k] * yf[k]
    pv_rev = [g * x / 1000 for g, x in zip(gen, price)]
    goo_rev = [g * x / 1000 for g, x in zip(gen, goo)]
    bmw = [p["bess_mw"] * o for o in ops]
    if p["fix6_bess"]:
        bess = [p["bess_block"].get(e.year, 0.0) * i * m * y for e, i, m, y in zip(ends, idx, bmw, yf)]
    else:
        bess = [sum(CURVES[e.year][c] for c in ("bess_gross_discharge_eur_kw", "bess_ancillary_eur_kw",
                                                 "bess_grid_charging_eur_kw", "bess_network_charges_eur_kw",
                                                 "bess_cost_pv_stored_eur_kw")) * i * m * y
                for e, i, m, y in zip(ends, idx, bmw, yf)]
    rev = [a + b + c for a, b, c in zip(pv_rev, goo_rev, bess)]
    opex = [-(p["om"] * mw * o * y * i) - p["bal"] * g / 1000 * i - p["lease"] * y * o * i
            - p["cable"] * mw * o * y * i - p["rtm"] * pr - p["bess_opex"] * m * y * i
            for o, y, i, g, pr, m in zip(ops, yf, idx, gen, pv_rev, bmw)]
    ebitda = [a + b for a, b in zip(rev, opex)]
    days = [(e - s).days + 1 for s, e in Q]
    nwc = [r * p["days_rec"] / dd + o * p["days_pay"] / dd for r, o, dd in zip(rev, opex, days)]
    dnwc = [(nwc[k - 1] if k else 0) - nwc[k] for k in range(n)]
    capex_total = p["capex_pv"] * mw + p["cable_capex"] + p["hv"] + p["land"] + p["bess_capex"] * p["bess_mwh"]
    ew_dev = rt["ew_total"]
    aug = []
    for (s, e), i in zip(Q, idx):
        if p["fix7_aug"]:
            aug.append(-p["aug_cost"] * p["bess_mwh"] * i if (e.year == p["aug_year_new"] and e.month == 3) else 0.0)
        else:
            aug.append(-p["aug_cost"] * p["bess_mwh"] if (e.year == p["aug_year"] and e.month == 3) else 0.0)
    # tax (rows 64-70)
    dep_base = capex_total - p["land"]
    dep, cum = [], 0.0
    for k in range(n):
        x = -min(dep_base / p["dep_years"] * ops[k] * yf[k], dep_base + cum)
        dep.append(x)
        cum += x
    tax, lcf = [], 0.0
    for k in range(n):
        ebt = ebitda[k] + dep[k]
        used = min(lcf, max(0, ebt) * p["loss_cap"])
        lcf = lcf + max(0, -ebt) - used
        tax.append(-(max(0, ebt) - used) * p["cit"])
    # monthly RtB rows → quarters
    def q_sum(series, upto=None):
        out = []
        for s, e in Q:
            hi = e if upto is None else min(e, upto)
            out.append(sum(x for x, t in zip(series, rt["dates"]) if s <= t <= hi))
        return out
    m_up = [p["upfront"] if m == 0 else 0 for m in range(19)]
    upfront = q_sum([-x for x in m_up])
    trans = q_sum([-p["transaction"] if m == 0 else 0 for m in range(19)])
    pre = [(-p["pre_rtb"] / 6 if p["pre_first"] <= t <= p["pre_last"] else 0.0) for t in rt["dates"]]
    proj = [ebitda[k] + dnwc[k] - capex_total * phase[k] + aug[k] + tax[k] + upfront[k] + trans[k] for k in range(n)]
    sale_q = [1 if s <= sale <= e else 0 for s, e in Q]
    buyer = [-p["price"] * mw * sale_q[k] - ew_dev * sale_q[k] - (capex_total - ew_dev) * phase[k] + aug[k]
             + (ebitda[k] + dnwc[k] + tax[k]) * (1 if ends[k] > sale else 0) for k in range(n)]
    dep_m = q_sum([-x for x in rt["cash_dep"]], upto=sale - dt.timedelta(days=1))
    release = [rt["cash_at_sale"] * (1 if s <= p["cod"] <= e else 0) for s, e in Q]
    ewpre = q_sum([-x for x in rt["ew"]])
    ewpre = [a + b for a, b in zip(ewpre, q_sum(pre))]
    bgm = q_sum([-x for x in _bg_int(rt, p)], upto=sale - dt.timedelta(days=1))
    bg_after = [-rt["bg_after"] * p["bg_rate_after"] * y * _share_in(s, e, sale, p["cod"]) for (s, e), y in zip(Q, yf)]
    buyout = [-rt["psh_buyout"] * x for x in sale_q]
    fut = [upfront[k] + trans[k] + dep_m[k] + release[k] + ewpre[k] + bgm[k] + bg_after[k] + buyout[k]
           - (capex_total - ew_dev) * phase[k] + aug[k] + ebitda[k] + dnwc[k] + tax[k] for k in range(n)]
    cumf, pay = 0.0, None
    for k in range(n):
        cumf += fut[k]
        if pay is None and cumf >= 0 and k > 0:
            pay = ends[k]
    first = 2029
    return dict(
        ends=ends, ebitda=ebitda, rev=rev, gen=gen, tax=tax, proj=proj, buyer=buyer, fut=fut, capex_total=capex_total,
        project_irr=xirr(proj, ends), buyer_irr=xirr(buyer, ends),
        buyer_moic=sum(x for x in buyer if x > 0) / -sum(x for x in buyer if x < 0),
        buyer_npv=xnpv(p["npv_rate"], buyer, ends) / 1000,
        htm_xirr=xirr(fut, ends), htm_moic=sum(x for x in fut if x > 0) / -sum(x for x in fut if x < 0),
        htm_npv=xnpv(p["npv_rate"], fut, ends) / 1000, htm_capital=-sum(x for x in fut if x < 0) / 1000,
        htm_payback=(pay - p["acq"]).days / 365.25 if pay else float("nan"),
        ffy_gwh=sum(g for g, e in zip(gen, ends) if e.year == first) / 1000,
        ffy_price=CURVES[first]["pv_capture_selected_nominal"],
        ffy_revenue=sum(r for r, e in zip(rev, ends) if e.year == first) / 1000,
        life_pv_revenue=sum(pv_rev) / 1000, life_goo=sum(goo_rev) / 1000, life_revenue=sum(rev) / 1000,
        life_ebitda=sum(ebitda) / 1000, life_margin=sum(ebitda) / sum(rev), life_capex=capex_total / 1000,
        life_opex=-sum(opex) / 1000, life_bess_net=sum(bess) / 1000,
    )


def _bg_int(rt, p):
    """Monthly bank-guarantee interest (Calculation row 129), recomputed from the balances."""
    out = []
    for m, bal in enumerate(rt["bg_bal"]):
        out.append(bal * (p["euribor"][m] + p["bg_margin"]) / 12 if m >= 1 else 0.0)
    return out


def _leap(y):
    return y % 4 == 0 and (y % 100 != 0 or y % 400 == 0)


def run(**over):
    p = defaults(**over)
    rt = rtb(p)
    out = {k: x for k, x in rt.items() if isinstance(x, float)}
    out.update({k: x for k, x in htm(p, rt).items() if isinstance(x, float)})
    return out


def targets():
    return {r["key"]: float(r["target"]) for r in csv.DictReader(open(HERE / "old_targets.csv"))}


if __name__ == "__main__":
    res, tg = run(), targets()
    for k, t in tg.items():
        x = res.get(k, float("nan"))
        ok = abs(x - t) <= 1e-6 * max(1, abs(t))
        print(f"{k:<22} target {t:>16.6f}  shadow {x:>16.6f}  {'ties' if ok else 'DIFF'}")
