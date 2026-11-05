"""Zalahaláp – reconciliation of the v2 workbook (SPEC v2 §9): the v3 shadow, a bridge from v3 to the PV + 4 h column
(Base, merchant BESS basis) and the tie-out of an independent Python shadow to the workbook.

    python3 projects/zalahalap/reconcile.py            → out/reconciliation.md (needs LibreOffice for Low / High)
    python3 projects/zalahalap/reconcile.py --targets  → data/recon_targets.csv only

Three independent calculations:
  1. data/old_engine/old_zhlp.py – Python shadow of the original v3 model; ties to every target of old_targets.csv.
  2. run() below – the v1 structure (quarterly, SPEC v1) with switches for the §13 fixes; its v3-rules set is
     stored in data/recon_targets.csv. It carries the bridge from v3 to the v2 inputs on the v1 grid.
  3. shadow.py – the v2 structure (annual, SPEC v2), coded from the SPEC; it carries the last bridge steps and ties to
     the recalculated workbook (out/Zalahalap_Model_recalc.xlsx; Low / High recalculated here in LibreOffice).
Stdlib + fepm.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(HERE / "data" / "old_engine"))
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE / "data"))

import old_zhlp as V3  # noqa: E402
import shadow as V2  # noqa: E402
import shadow_v20 as V20  # noqa: E402  (frozen v2.0 shadow, before the v2.1 timeline rebase)

DATA = HERE / "data"
OUT = HERE / "out"
XLSX = OUT / "Zalahalap_Model.xlsx"
RECALC = OUT / "Zalahalap_Model_recalc.xlsx"
EXCEL0 = dt.date(1899, 12, 30)
N = 155
START = dt.date(2025, 10, 1)


# =========================================================================================== shadow engine
def eom(x, n=0):
    return V3.eomonth(x, n)


def edate(x: dt.date, months: int) -> dt.date:
    m = x.month - 1 + months
    y, m = x.year + m // 12, m % 12 + 1
    last = eom(dt.date(y, m, 1))
    return dt.date(y, m, min(x.day, last.day))


def curves():
    """Case A curves (v3) and the Aurora blocks, with model years 2025-2064."""
    au = {}
    for r in csv.DictReader(open(DATA / "aurora_curves.csv")):
        vals = [float(r[str(y)]) for y in range(2026, 2061)]
        au[(r["series"], r["block"])] = vals

    def ext(vals, first=0.0, grow=None):
        out = [first] + list(vals)
        while len(out) < 40:
            out.append(out[-1] * (1 + grow) if grow else out[-1])
        return dict(zip(range(2025, 2065), out))
    old = V3.CURVES
    c = {"index": ext(au[("euro_index", "Central")], 1.0, 0.02),
         "v3_capture": {y: old[y]["pv_capture_selected_real2025"] for y in range(2025, 2065)},
         "v3_bess": {y: old[y]["bess_net_real2025_eur_kw"] for y in range(2025, 2065)},
         "v3_dis": {y: old[y]["bess_discharge_mwh_per_mw"] for y in range(2025, 2065)},
         "v3_res": {y: -old[y]["bess_charge_from_pv_mwh_per_mw"] for y in range(2025, 2065)}}
    for (s, b), vals in au.items():
        c[(s, b)] = ext(vals)
    return c


C = None


def params(**over):
    """Base set (1) for case A; overrides give the reconciliation set and the bridge steps."""
    p = dict(
        mwp=61, mva=50, yld=1250, rtb=dt.date(2026, 7, 31), exit=dt.date(2026, 10, 31), ntp=dt.date(2027, 3, 31),
        cod=dt.date(2028, 3, 31), acq=dt.date(2025, 11, 30), life=35, pre_first=dt.date(2026, 10, 31),
        pre_last=dt.date(2027, 3, 31), pv_curve="v3_capture", block="2027 solar DC 1.0cy 4h",
        pv_avail=0.995, deg=0.003, clip=0.02, bess_on=1, bess_mw=50, bess_h=4, afrr_down=0.0,
        haircut=0.0, bess_capex=110, aug_cost=110, aug_year=0, aug_year_v3=2048, bess_opex=30.8, goo=1.75,
        pv_capex=350, cable=3700, subst=9500, land=620, pre_rtb=3182, pre_reimb=1,
        scope=4, late_capex=1, ew_reimb=1, ew_int=1, ew_psh=0,
        om=14.96, cable_om=5, bal=5, lease=96, hipa=0.02, innov=0.003, fx=400, bg=0.0187 + 0.0175,
        price=90, step=1, low=0.6, high=0.5, thr=90, slide_end=100, trans=200, spv=549, spv_int=0, dsa=15,
        rate=0.12, psh_land=150, psh_step=0, psh_devex=15,
        fix2=1, fix3=1, fix5=1, fix6=1, fix7=1, fix8=1,
        disc=0.08, buyer_irr=0.08,
        gearing=0.55, dscr_c=1.20, dscr_m=1.35, tenor=15, debt_rate=0.05, upfront=0.0175, commit=0.009, dsra_m=6,
        cit=0.09, dep_years=30, dep_bess=10, loss_cap=0.5, loss_years=5,
        # fix round 1 (False/legacy values reproduce the build before it)
        pv_tr_curve=None, grid_import=0, repower=1, dep_sh=False, unlev_tax=True, int_tax=True,
    )
    p.update(over)
    return p


def params_b(**over):
    """Base set, case B (Aurora merchant, Calc_2)."""
    blk = "2027 solar DC 1.0cy 4h"
    p = params(mwp=63, mva=49.9, yld=1265, rtb=dt.date(2026, 12, 31), exit=dt.date(2027, 3, 31),
               ntp=dt.date(2027, 6, 30), cod=dt.date(2028, 3, 31), pv_curve=("direct_price", blk),
               pv_tr_curve=("transfer_price", blk), clip=0.0, block=blk)
    p.update(over)
    return p


LEGACY = dict(pv_curve=("implied_capture", "2027 solar DC 1.0cy 4h"), pv_tr_curve=None, grid_import=50, clip=0.02,
              dep_sh=True, unlev_tax=False, int_tax=False)     # case B as built before fix round 1


SET2 = dict(scope=6, ew_psh=1, hipa=0, innov=0, step=0, spv=150, spv_int=1, dsa=0, psh_step=1,
            fix2=0, fix3=0, fix5=0, fix6=0, fix7=0, fix8=0, gearing=0, dep_bess=30, loss_years=40)

EW = [(60 / 400 * 1000, 0, dt.date(2025, 11, 30), 1), (60 / 400 * 1000, 0, dt.date(2026, 2, 28), 2),
      (0, 25, dt.date(2026, 2, 28), 3), (0, 460, dt.date(2026, 4, 30), 4), (0, 650, dt.date(2026, 7, 31), 5),
      (30 / 400 * 1000, 0, dt.date(2026, 7, 31), 6)]        # HUF k at FX 400 → EUR k (recomputed with fx), EUR k
DEPOSITS = [(45, 0, dt.date(2025, 11, 30), None, 1), (180, 0, dt.date(2025, 11, 30), None, 1),
            (485.55, 0, dt.date(2025, 11, 30), dt.date(2026, 1, 31), 1),
            (242.775, 0, dt.date(2026, 1, 31), dt.date(2026, 3, 31), 0), (242.775, 0, dt.date(2026, 1, 31), None, 0),
            (80.925, 0, dt.date(2026, 4, 30), None, 0), (485.55, 1, dt.date(2025, 11, 30), None, 0)]


def run(p):
    """One Calc sheet for case A. Returns headline outputs and the timeline rows used by the reconciliation."""
    global C
    C = C or curves()
    Q = []
    s = START
    for _ in range(N):
        e = eom(s, 2)
        Q.append((s, e))
        s = e + dt.timedelta(days=1)
    starts, ends = [a for a, _ in Q], [b for _, b in Q]
    days = [(b - a).days + 1 for a, b in Q]
    yf = [x / 365 for x in days]
    year = [b.year for b in ends]

    def event(d):
        return [1 if a <= d <= b else 0 for a, b in Q]

    def spread(a, b):                                        # amount spread over a → b by days
        n = (b - a).days + 1
        return [max(0, (min(b, e) - max(a, s0)).days + 1) / n for s0, e in Q]

    def window(a, b):
        return [max(0, (min(b, e) - max(a, s0)).days + 1) / d for (s0, e), d in zip(Q, days)]
    cod = eom(p["cod"], 0)
    one = dt.timedelta(days=1)
    acq_ev, exit_ev, ntp_ev, cod_ev = event(p["acq"]), event(p["exit"]), event(p["ntp"]), event(cod)
    post_x = [1 if a > p["exit"] else 0 for a in starts]
    post_c = [1 if b > cod else 0 for b in ends]
    constr = spread(p["ntp"] + one, cod)
    ops = window(cod + one, edate(cod, p["life"] * 12))
    elapsed = [max(0, (a - one - cod).days) / 365 for a in starts]
    dsa_w = spread(p["acq"] + one, p["rtb"])
    pre_w = spread(eom(p["pre_first"], -1) + one, p["pre_last"])
    tenor = window(cod + one, edate(cod, p["tenor"] * 12))
    tenor_end = event(edate(cod, p["tenor"] * 12))
    dep = spread(cod + one, edate(cod, p["dep_years"] * 12))
    dep_b = spread(cod + one, edate(cod, p["dep_bess"] * 12))
    rq = p["debt_rate"] / 4
    q4 = [1 if b.month == 12 else 0 for b in ends]
    q1 = [1 if b.month <= 3 else 0 for b in ends]
    infl = [C["index"][y] for y in year]
    pv_real = [C[p["pv_curve"]][y] for y in year]
    pv_mkt = [x * i for x, i in zip(pv_real, infl)]
    pv_tr = [C[p["pv_tr_curve"] or p["pv_curve"]][y] * i for y, i in zip(year, infl)]
    goo = [p["goo"] * i for i in infl]
    bkey = p["block"] if p["fix6"] else "v3"
    stk_real = [C["v3_bess"][y] if bkey == "v3" else C[("storage_total", bkey)][y] for y in year]
    bmw = p["bess_mw"] * p["bess_on"]
    bmwh = bmw * p["bess_h"]
    if bkey == "v3":
        res_sh = [1.0] * N
        v_res = [C["v3_res"][y] for y in year]
    else:                                                     # PV-only factor by component (fix round 2)
        pvo, dcap = C[("pv_only_factor", bkey)], C[("down_cap_share", bkey)]
        res_sh = [min(1, (pvo[y] if y > 2025 else 1.0) + p["afrr_down"] * dcap[y]) for y in year]
        v_res = [C[("charge_res", bkey)][y] for y in year]
    gcov = min(1, p["grid_import"] / max(1, bmw))
    if p["fix7"]:
        rc = C[("repower_count", p["block"])] if p["fix6"] else None
        aug_y = p["aug_year"] or (next((y for y in range(2026, 2061) if rc[y] >= 1), 0) if rc else 2048)
    else:
        aug_y = p["aug_year_v3"]
    live = [1 - (1 - p["repower"]) * (1 if y >= aug_y else 0) for y in year]
    stack = [x * i * (1 - p["haircut"]) * (r + (1 - r) * gcov)
             for x, i, r in zip(stk_real, infl, res_sh)]
    gen = [p["mwp"] * p["yld"] * p["pv_avail"] * (1 - p["deg"]) ** el * o * y for el, o, y in zip(elapsed, ops, yf)]
    clip_sh = p["clip"] * (p["fix5"] if p["pv_curve"] == "v3_capture" else 1)
    clip = [-g * clip_sh for g in gen]
    net = [g + c for g, c in zip(gen, clip)]
    to_bess = [min(n, -v * bmw * o * y * lv) for n, v, o, y, lv in zip(net, v_res, ops, yf, live)]
    pv_exp = [n - t for n, t in zip(net, to_bess)]
    pv_rev = [(a * x + b * tr) / 1000 for a, x, b, tr in zip(pv_exp, pv_mkt, to_bess, pv_tr)]
    goo_rev = [n * x / 1000 for n, x in zip(net, goo)]
    bess_rev = [x * bmw * o * y * lv for x, o, y, lv in zip(stack, ops, yf, live)]
    rev = [a + b + c for a, b, c in zip(pv_rev, goo_rev, bess_rev)]
    opex = [-(p["om"] + p["cable_om"]) * p["mwp"] * i * o * y - p["bal"] * n / 1000 * i - p["lease"] * i * o * y
            - p["bess_opex"] * bmw * i * o * y * lv - r * (p["hipa"] + p["innov"])
            for i, o, y, n, r, lv in zip(infl, ops, yf, net, rev, live)]
    ebitda = [r + x for r, x in zip(rev, opex)]
    # development
    spv = [-p["spv"] * a for a in acq_ev]
    trans = [-p["trans"] * a for a in acq_ev]
    dsa = [-p["dsa"] * p["mwp"] * w for w in dsa_w]
    ew_lines = [(h * 400 / p["fx"] + e, d, rank) for h, e, d, rank in EW]
    ew_dev = [-sum(a for a, d, rank in ew_lines if s0 <= d <= e and rank <= p["scope"]) for s0, e in Q]
    ew_ntp = [-p["late_capex"] * x * sum(a for a, _, rank in ew_lines if rank > p["scope"]) for x in ntp_ev]
    pre_dev = [-p["pre_rtb"] * w for w in pre_w]
    dep_l = [(h * 1000 / p["fx"] * (1 - ph * p["fix3"]), paid, rel, cash) for h, ph, paid, rel, cash in DEPOSITS]
    dep_cash = [-sum(a for a, paid, _, cash in dep_l if cash and s0 <= paid <= e) for s0, e in Q]
    dep_back = [sum(a for a, _, rel, cash in dep_l if cash and rel and s0 <= rel <= e) for s0, e in Q]
    dep_out = sum(a for a, _, rel, cash in dep_l if cash and rel is None)
    dep_cod = [dep_out * x for x in cod_ev]
    bg_d = [sum(a for a, paid, rel, cash in dep_l if not cash and rel and paid <= e < rel) for e in ends]
    bg_x = [sum(a for a, paid, rel, cash in dep_l if not cash and rel is None and paid <= e) for e in ends]
    bg_hold = [-(a + b * (1 - pc)) * p["bg"] * d / 365 for a, b, pc, d in zip(bg_d, bg_x, post_c, days)]
    bg_rtb = [b * (1 - px) for b, px in zip(bg_hold, post_x)]
    # construction
    capex_b = p["bess_capex"] * bmwh
    capex_tot = p["pv_capex"] * p["mwp"] + p["cable"] + p["subst"] + p["land"] + capex_b
    capex = [-(capex_tot + sum(ew_dev) + sum(ew_ntp)) * c for c in constr]
    aug = [-p["aug_cost"] * p["repower"] * bmwh * (i if p["fix7"] else 1) * (1 if y == aug_y else 0) * a
           for i, y, a in zip(infl, year, q1)]
    cost = [a + b for a, b in zip(capex, ew_ntp)]
    idx = [(1 + rq) ** (k + 1) for k in range(N)]
    idx_cod = sum(x for x, e in zip(idx, ends) if e == cod)
    cost_e = sum(c / i for c, i in zip(cost, idx)) * idx_cod
    contr = [0.0] * N                                         # case A: no offer
    w_dscr = [(c / p["dscr_c"] + (r - c) / p["dscr_m"]) / r if r > 0 else 0 for c, r in zip(contr, rev)]
    nq = p["tenor"] * 4

    def debt(tax):
        cfads = [(e + t) * pc for e, t, pc in zip(ebitda, tax, post_c)]
        caps = [cf * w for cf, w, tn in zip(cfads, w_dscr, tenor) if tn == 1]
        cap_min = min(caps) if caps else 0
        ds_cap = cap_min * (1 - (1 + rq) ** -nq) / rq
        k = max(0, min(p["gearing"], ds_cap / -cost_e if cost_e else 0, 1e99))
        d_cod = k * -cost_e
        ann = d_cod * rq / (1 - (1 + rq) ** -nq) if d_cod else 0.0
        out = {key: [0.0] * N for key in ("int", "sched", "draw", "idc", "close", "up", "cf", "dsra")}
        close = 0.0
        draws = [-k * c for c in cost]
        tot_draw = sum(draws)
        cum = 0.0
        for t in range(N):
            open_ = close
            d_int = -open_ * rq * post_c[t]
            d_sched = -min(open_, ann + d_int) * tenor[t]
            idc = open_ * rq * (1 - post_c[t])
            close = open_ + draws[t] + idc + d_sched
            cum += draws[t]
            out["int"][t], out["sched"][t], out["draw"][t], out["idc"][t], out["close"][t] = d_int, d_sched, \
                draws[t], idc, close
        d_close_cod = sum(x for x, e in zip(out["close"], ends) if e == cod)
        cum = 0.0
        prev_open = 0.0
        for t in range(N):
            open_ = out["close"][t - 1] if t else 0.0
            out["up"][t] = -p["upfront"] * d_close_cod * (1 if open_ == 0 else 0) * (1 if draws[t] > 0 else 0)
            cum += draws[t]
            out["cf"][t] = -p["commit"] * (tot_draw - cum + draws[t]) * yf[t] * (1 if constr[t] > 0 else 0)
        service = [a + b for a, b in zip(out["int"], out["sched"])]
        dsra_e = -sum(sv for sv, e in zip(service, ends) if cod < e <= edate(cod, p["dsra_m"]))
        out["dsra"] = [dsra_e * (te - ce) for te, ce in zip(tenor_end, cod_ev)]
        out["service"], out["cfads"], out["k"], out["d_close_cod"] = service, cfads, k, d_close_cod
        return out

    def losses(taxable):
        lapse_d = edate(p["ntp"], p["loss_years"] * 12)
        lcf, out = 0.0, []
        for t in range(N):
            add = max(0, -taxable[t])
            use = -min(lcf, p["loss_cap"] * max(0, taxable[t]))
            lapse = -(lcf + add + use) * (1 if starts[t] <= lapse_d <= ends[t] else 0)
            lcf = lcf + add + use + lapse
            out.append(-p["cit"] * (max(0, taxable[t]) + use))
        return out

    tax = [0.0] * N
    for _ in range(100):                                       # the tax circularity, as Solver.bas
        db = debt(tax)
        dep_base = capex_tot - capex_b - p["land"] - sum(pre_dev) - (sum(dsa) + sum(spv)) * p["dep_sh"]
        tdep = [-dep_base * x for x in dep]
        tdep_i = [-sum(db["idc"]) * x for x in dep]
        tdep_b = [-capex_b * x for x in dep_b]
        aug_i = sum(aug)
        a0 = next((e for a, e in zip(aug, ends) if a < 0), EXCEL0) + one
        tdep_a = [aug_i * x for x in (spread(a0, edate(a0 - one, p["dep_bess"] * 12)) if aug_i else [0] * N)]
        taxable = [e + a + ii + b + c + i + u + f + g for e, a, ii, b, c, i, u, f, g in
                   zip(ebitda, tdep, tdep_i, tdep_b, tdep_a, db["int"], db["up"], db["cf"], bg_hold)]
        taxable_u = [e + a + b + c + g for e, a, b, c, g in zip(ebitda, tdep, tdep_b, tdep_a, bg_hold)]
        new = losses(taxable)
        if max(abs(a - b) for a, b in zip(new, tax)) < 1e-9:
            tax = new
            break
        tax = new
    db = debt(tax)
    dev_cf = [sum(x) for x in zip(spv, trans, dsa, ew_dev, pre_dev, dep_cash, dep_back, dep_cod, bg_hold)]
    proj_pre = [a + b + c + d + e for a, b, c, d, e in zip(dev_cf, capex, ew_ntp, aug, ebitda)]
    tax_u = losses(taxable_u) if p["unlev_tax"] else tax
    proj_cf = [a + t for a, t in zip(proj_pre, tax_u)]
    debt_cf = [a + b + c + d + e for a, b, c, d, e in zip(db["draw"], db["up"], db["cf"], db["service"], db["dsra"])]
    eq_cf = [a + t + b for a, t, b in zip(proj_pre, tax, debt_cf)]
    # buyer at RtB
    first_exit = next(t for t in range(N) if exit_ev[t])
    ew_to_exit = -sum(x for x, a in zip(ew_dev, starts) if a <= p["exit"])
    reimb_x = [x * (p["ew_reimb"] * ew_to_exit + dep_out) for x in exit_ev]
    pre_e = -sum(x for x, a in zip(pre_dev, starts) if a <= p["exit"]) if p["fix2"] else -sum(pre_dev)
    pre_reimb = [p["pre_reimb"] * (x * pre_e - p["fix2"] * px * d) for x, px, d in zip(exit_ev, post_x, pre_dev)]
    bu_rtb = [px * (pc - dc + bg + dcod + ew) - rx - pr for px, pc, dc, bg, dcod, ew, rx, pr in
              zip(post_x, proj_cf, dev_cf, bg_hold, dep_cod, ew_dev, reimb_x, pre_reimb)]
    imp_r = V3.xnpv(p["buyer_irr"], bu_rtb[first_exit:], ends[first_exit:]) / p["mwp"]
    price = p["price"] if p["price"] > 0 else max(0, imp_r)
    consid = price * p["mwp"]
    buy_r = [b - consid * x for b, x in zip(bu_rtb, exit_ev)]
    # FEP return (12 %)
    f_c = (1 + p["rate"] / 12) ** 3 - 1
    cl_f = [p["rate"] * d / 365 if p["fix8"] else f_c for d in days]
    cl_add = [-(a + p["ew_int"] * b + c + d + e + p["spv_int"] * f) for a, b, c, d, e, f in
              zip(dsa, ew_dev, dep_cash, dep_back, bg_hold, spv)]

    def accrue(adds):
        close, rets, caps = 0.0, [], []
        for t in range(N):
            ret = close * cl_f[t] * (1 - post_c[t])
            rets.append(ret)
            if p["fix8"]:
                cap = q4[t] * sum(r for r, y in zip(rets, year) if y == year[t])
            else:
                cap = ret
            close = close + adds[t] + cap
        return rets
    cl_ret = accrue(cl_add)
    ps_ret = accrue([p["psh_land"] * a for a in acq_ev])
    to_exit = [1 if a <= p["exit"] else 0 for a in starts]
    cl_e = sum(r * x for r, x in zip(cl_ret, to_exit))
    ps_e = sum(r * x for r, x in zip(ps_ret, to_exit))
    # share and waterfall
    if p["step"]:
        sh = p["low"] if price <= p["thr"] else p["high"]
    else:
        sh = p["low"] - (p["low"] - p["high"]) * min(1, max(0, (price - p["thr"]) / (p["slide_end"] - p["thr"])))
    dues = [0.0, p["spv"], -sum(dsa), cl_e, -sum(bg_rtb), p["psh_step"] * (p["psh_land"] + ps_e)]
    paid, avail = [], consid
    for dd in dues:
        x = min(dd, avail - sum(paid))
        paid.append(x)
    dist = max(0, avail - sum(dues))
    w_fep = sum(paid[1:5]) + sh * dist
    w_psh = paid[5] + (1 - sh) * dist
    rtb_out = [(a + b + c + d + e) * (1 - px) + bg + pd for a, b, c, d, e, px, bg, pd in
               zip(spv, trans, dsa, ew_dev, dep_cash, post_x, bg_rtb, pre_dev)]
    rtb_in = [db_ * (1 - px) + rx + pr + w_fep * x for db_, px, rx, pr, x in
              zip(dep_back, post_x, reimb_x, pre_reimb, exit_ev)]
    rtb_pre = [a + b for a, b in zip(rtb_out, rtb_in)]
    rtb_cf = [a - p["cit"] * paid[3] * x * p["int_tax"] for a, x in zip(rtb_pre, exit_ev)]
    # hold to maturity
    htm_w = [1 - px * 0 for px in post_x]
    htm_cf = [e * w - w_psh * x for e, w, x in zip(eq_cf, htm_w, exit_ev)]
    psh_r = [-p["psh_land"] * a - p["psh_devex"] * p["mwp"] * w - d - p["ew_psh"] * e + w_psh * x
             for a, w, d, e, x in zip(acq_ev, dsa_w, dsa, ew_dev, exit_ev)]
    return dict(
        rtb_profit=sum(rtb_cf), rtb_irr=V3.xirr(rtb_cf, ends), rtb_irr_pre=V3.xirr(rtb_pre, ends),
        rtb_moic=(sum(rtb_in) + sum(rtb_cf) - sum(rtb_pre)) / -sum(rtb_out),
        htm_irr_pre=V3.xirr([h - t for h, t in zip(htm_cf, tax)], ends), proj_irr_pre=V3.xirr(proj_pre, ends),
        first_rev=sum(r for r, y in zip(rev, year) if y == cod.year + 1),
        htm_irr=V3.xirr(htm_cf, ends), htm_moic=sum(x for x in htm_cf if x > 0) / -sum(x for x in htm_cf if x < 0),
        htm_npv=V3.xnpv(p["disc"], htm_cf, ends), proj_irr=V3.xirr(proj_cf, ends),
        buyer_irr=V3.xirr(buy_r[first_exit:], ends[first_exit:]), implied_rtb=max(0, imp_r), implied_raw=imp_r,
        life_revenue=sum(rev), life_ebitda=sum(ebitda), share=sh, w_fep=w_fep, w_psh=w_psh, dist=dist,
        psh_profit=sum(psh_r), debt_cod=db["d_close_cod"], tax=sum(tax), ends=ends, rtb_cf=rtb_cf, htm_cf=htm_cf,
    )


# =========================================================================================== stored targets
RECON_KEYS = ["rtb_profit", "rtb_irr", "htm_irr", "proj_irr", "buyer_irr", "life_revenue", "life_ebitda"]
V3_KEY = {"rtb_profit": ("rtb_net_profit", 1000), "rtb_irr": ("rtb_xirr_incl", 1), "htm_irr": ("htm_xirr", 1),
          "proj_irr": ("project_irr", 1), "buyer_irr": ("buyer_irr", 1), "life_revenue": ("life_revenue", 1000),
          "life_ebitda": ("life_ebitda", 1000)}


def write_targets():
    sh = run(params(**SET2))
    tg = V3.targets()
    path = DATA / "recon_targets.csv"
    with open(path, "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["key", "shadow", "v3"])
        for k in RECON_KEYS:
            key, scale = V3_KEY[k]
            w.writerow([k, repr(sh[k]), repr(tg[key] * scale)])
    return path, sh


OLD_STEPS = [  # label, old_zhlp switches (cumulative)
    ("Fix 3: no bank-guarantee phantom", dict(fix3_bg=True)),
    ("Fix 2: pre-RtB costs reimbursed as spent", dict(fix2_reimb=True)),
    ("Fix 8: simple act/365 return, 31-Dec capitalisation", dict(fix8_interest=True)),
    ("User: share step (60 % at 90)", dict(share_step=True)),
    ("User: early works, four lines (E48/E49 out of the RtB period)", dict(scope4=True)),
    ("Fix 5: PV export clipped (2 %)", dict(fix5_clip=True)),
    ("Fix 6: Aurora 4 h block total for the BESS", dict(fix6_bess=True)),
    ("Fix 7: augmentation indexed", dict(fix7_aug=True)),
]
OLD_OUTS = [("rtb_net_profit", "FEP profit – RtB", "m"), ("rtb_xirr_incl", "FEP IRR – RtB", "%"),
            ("rtb_moic_incl_dep", "FEP MOIC – RtB", "x"), ("htm_xirr", "FEP IRR – HTM", "%"),
            ("htm_npv", "FEP NPV – HTM", "m"), ("project_irr", "Project IRR", "%"), ("buyer_irr", "Buyer IRR", "%"),
            ("life_revenue", "Lifetime revenue", "m"), ("life_ebitda", "Lifetime EBITDA", "m")]



def old_bridge():
    block = {y: v for y, v in curves()[("storage_total", "2027 solar DC 1.0cy 4h")].items()}
    over, out = {}, [("v3 original (shadow = targets)", V3.run())]
    for label, sw in OLD_STEPS:
        over.update(sw)
        out.append((label, V3.run(**over, bess_block=block, aug_year_new=2048)))
    return out


# =========================================================================================== bridge v3 → v2
BLK = "2027 solar DC 1.0cy 4h"
D = dt.date
V1_STEPS = [  # label, SPEC reference, switches on the v1 quarterly shadow (cumulative from set 2)
    ("§13 fixes 2, 3, 5–8 (reimbursement timing, guarantee phantom, clipping, Aurora 4 h block, augmentation, "
     "act/365 return)", "v1 §13.2–13.8", dict(fix2=1, fix3=1, fix5=1, fix6=1, fix7=1, fix8=1)),
    ("§13.14 share step (60 % at ≤ 90)", "v1 §13.14", dict(step=1)),
    ("§13.16–13.17 Aurora 4 h direct and transfer prices, DC clipping 0", "v1 §13.16–17",
     dict(pv_curve=("direct_price", BLK), pv_tr_curve=("transfer_price", BLK), clip=0.0)),
    ("63 MWp, 1,265 MWh/MWp, 49.9 MVA", "v2 §3.1", dict(mwp=63, yld=1265, mva=49.9)),
    ("v2 dates: RtB 31-Dec-26, exit 31-Mar-27, NTP 30-Jun-27", "v2 §3.2",
     dict(rtb=D(2026, 12, 31), exit=D(2027, 3, 31), ntp=D(2027, 6, 30))),
    ("Early works four lines; E48/E49 to capex at NTP", "v1 §13.15", dict(scope=4, ew_psh=0)),
    ("SPV price 549 (no return), DSA 945, no PSH land step", "v2 §6", dict(spv=549, spv_int=0, dsa=15, psh_step=0)),
    ("Tax defaults: BESS 10 years, losses lapse NTP + 5 years", "v2 §3.2", dict(dep_bess=10, loss_years=5)),
]
V2_STEPS = [  # label, SPEC reference, shadow overrides (v2.0 shadow; the v2.1 steps follow in bridge())
    ("Annual timeline, year-end dates (the RtB route keeps exact dates), unlevered, v2 JV split", "v2 §4, §6",
     dict(ppa_years=0, hipa=0, innov=0, opex_mw=30.8, opex_mwh=0)),
    ("BESS opex by duration: 18.70 k€/MW + 6.07 k€/MWh (Aurora; 4 h = 42.96)", "fix round M2",
     dict(ppa_years=0, hipa=0, innov=0)),
    ("HIPA 2 % and innovation contribution 0.3 % on revenue", "v2 §3.2", dict(ppa_years=0)),
    ("PPA 5 years at 45 €/MWh on all PV output = v2.0 base (fe5a4fe)", "v2 §5.2", {}),
]
V21_STEPS = [  # label, SPEC reference, shadow.py overrides (v2.1)
    ("Timeline rebase: spend to 31-Dec-26 in one column, all capex in 2027, COD 01-Mar-28 (unlevered)", "v2.1 §12.1",
     dict(gearing=0, structure="Merchant", ppa_price=45, ppa_years=5, om=14.96, cable_om=5)),
    ("PPA 80 €/MWh for 8 years from COD, then merchant (user 05-Nov-2026; unlevered)", "v2.2",
     dict(gearing=0, structure="Merchant", om=14.96, cable_om=5)),
    ("O&M 10 k€/MWp/yr all-in, MV cable O&M inside it (user 05-Nov-2026; unlevered)", "v2.3",
     dict(gearing=0, structure="Merchant")),
    ("Side row: PPA uplift (55/80/85, no BESS value; vs Merchant 80)", "v2.4 §13", dict(gearing=0, structure="PPA uplift")),
    ("Fixed + share (PV 45, BESS 70/55 on 1 cycle, 70 % above; vs Merchant 80)", "v2.4 §13", dict(gearing=0)),
    ("Levered case: 60 % gearing, 5 %, 20-year annuity (hold IRR levered; base)", "v2.1 §12.2", {}),
]
BRIDGE = [("hold_irr", "FEP hold IRR", "%"), ("proj_irr", "Project IRR", "%"), ("rtb_profit", "RtB profit (k€)", "k"),
          ("rtb_irr", "RtB XIRR", "%"), ("implied_raw", "Implied RtB premium before the 0 floor (k€/MWp)", "k")]


def bridge():
    tg = V3.targets()
    rows = [("v3 original (targets; the v3 shadow ties)", "v1 §3.2",
             dict(hold_irr=tg["htm_xirr"], proj_irr=tg["project_irr"], rtb_profit=tg["rtb_net_profit"] * 1000,
                  rtb_irr=tg["rtb_xirr_incl"], implied_raw=None))]
    over = dict(SET2)
    v1 = lambda r: dict(r, hold_irr=r["htm_irr"])           # noqa: E731  (v1 names the hold IRR htm_irr)
    rows.append(("v1 structure (quarterly), v3 rules: set 2 of the v1 workbook", "v1 §11", v1(run(params(**over)))))
    for label, ref, sw in V1_STEPS:
        over.update(sw)
        rows.append((label, ref, v1(run(params(**over)))))
    for label, ref, kw in V2_STEPS:
        rows.append((label, ref, V20.run("PV + 4h BESS", "Base", **kw)))
    for label, ref, kw in V21_STEPS:
        r = V2.run("PV + 4h BESS", "Base", **kw)
        rows.append((label, ref, dict(r, hold_irr=r["hold_irr_l"]) if not kw else r))
    return rows


SIDE = "Side row"      # a side row shows its effect against the row before it; the next row starts from there too


# =========================================================================================== tie-out
TIE = [  # shadow key, Dashboard label, occurrence, scale (workbook = shadow / scale)
    ("gen_ffy", "PV generation", 0, 1000), ("rev_ffy", "Revenue", 0, 1), ("ebitda_ffy", "EBITDA", 0, 1),
    ("rev_life", "Revenue", 1, 1), ("ebitda_life", "EBITDA", 1, 1), ("tax_life", "Lifetime corporate income tax, "
                                                                     "hold route (SPV)", 0, -1),
    ("capex_total", "Total uses to COD (= equity, unlevered)", 0, 1),
    ("implied_rtb", "Implied premium at the buyer target IRR (floored at 0)", 0, 1),
    ("buyer_rtb", "Buyer IRR at our price, unlevered after tax", 0, 1),
    ("buyer_rtb_pre", "Buyer IRR at our price, unlevered before tax", 0, 1),
    ("rtb_profit_pre", "FEP profit, pre-tax", 0, 1), ("rtb_profit", "FEP profit, post-tax", 0, 1),
    ("rtb_moic", "MOIC, pre-tax", 0, 1), ("rtb_irr_pre", "IRR, pre-tax", 0, 1), ("rtb_irr", "IRR, post-tax", 0, 1),
    ("rtb_npv", "NPV at the discount rate, post-tax (at acquisition)", 0, 1), ("psh_rtb", "PSH profit", 0, 1),
    ("implied_cod", "Implied COD price at the buyer target IRR (floored at 0)", 0, 1),
    ("ev", "Enterprise value used", 0, 1), ("buyer_cod", "Buyer IRR at our price, unlevered after tax", 1, 1),
    ("buyer_cod_pre", "Buyer IRR at our price, unlevered before tax", 1, 1),
    ("cod_contrib", "FEP contributions to COD", 0, 1), ("cod_proceeds", "FEP proceeds at COD", 0, 1),
    ("cod_profit_pre", "FEP profit, pre-tax", 1, 1), ("cod_profit", "FEP profit, post-tax", 1, 1),
    ("cod_moic", "MOIC, pre-tax", 1, 1), ("cod_irr_pre", "IRR, pre-tax", 1, 1), ("cod_irr", "IRR, post-tax", 1, 1),
    ("cod_npv", "NPV at the discount rate, post-tax (at acquisition)", 1, 1), ("psh_cod", "PSH profit", 1, 1),
    ("hold_equity", "FEP equity invested (incl. the 2048 repower)", 0, 1), ("hold_dist", "FEP distributions", 0, 1),
    ("hold_profit_pre", "FEP profit, pre-tax", 2, 1), ("hold_profit", "FEP profit, post-tax", 2, 1),
    ("hold_irr", "IRR, after tax", 0, 1), ("hold_irr_pre", "IRR, before tax", 0, 1),
    ("hold_npv", "NPV at the discount rate, after tax (at acquisition)", 0, 1),
    ("proj_irr", "Memo: project IRR, unlevered after tax (no buy-out)", 0, 1),
    ("proj_irr_pre", "Memo: project IRR, unlevered before tax", 0, 1),
    ("debt_drawn", "Debt drawn in the capex year (gearing × total uses)", 0, 1),
    ("idc", "Interest during construction, capitalised", 0, 1), ("debt_cod", "Debt at COD", 0, 1),
    ("annuity", "Annual debt service (level annuity)", 0, 1), ("dscr_min", "Minimum DSCR (information)", 0, 1),
    ("dscr_avg", "Average DSCR (information)", 0, 1), ("tax_life_l", "Lifetime corporate income tax, levered", 0, -1),
    ("buyer_rtb_l", "Buyer IRR at our price, levered after tax (raises debt)", 0, 1),
    ("buyer_cod_l", "Buyer IRR at our price, levered after tax (takes debt)", 0, 1),
    ("cod_profit_l", "Levered: FEP profit, post-tax (debt repaid from price)", 0, 1),
    ("cod_moic_l", "Levered: MOIC, pre-tax", 0, 1), ("cod_irr_lp", "Levered: IRR, pre-tax", 0, 1),
    ("cod_irr_l", "Levered: IRR, post-tax", 0, 1), ("psh_cod_l", "Levered: PSH profit", 0, 1),
    ("hold_equity_l", "Levered: FEP equity invested (uses less debt)", 0, 1),
    ("hold_profit_l", "Levered: FEP profit, post-tax", 0, 1), ("hold_irr_l", "Levered: IRR, after tax", 0, 1),
    ("hold_irr_lp", "Levered: IRR, before tax", 0, 1), ("hold_moic_l", "Levered: MOIC, after tax", 0, 1),
    ("cycles_day", "BESS cycles per day", 0, 1), ("dabove_ffy", "BESS discharge above the fixed cycles", 0, 1000),
    ("rbess_ffy", "BESS revenue", 0, 1), ("rbfix_ffy", "of which fixed price on the fixed cycles", 0, 1)]
RUNS = [("Base", "Fixed + share"), ("Low", "Fixed + share"), ("High", "Fixed + share"), ("Base", "PPA uplift"),
        ("Base", "Merchant")]
CASES = ["PV only", "PV + 2h BESS", "PV + 4h BESS"]


def dashboard_cells():
    """(shadow key, scale) -> Dashboard row, by label and occurrence in column C (visible part only)."""
    import openpyxl
    ws = openpyxl.load_workbook(XLSX, read_only=False)["Dashboard"]
    seen, rows = {}, {}
    for r in range(1, ws.max_row + 1):
        lab = ws.cell(r, 3).value
        if isinstance(lab, str) and not ws.row_dimensions[r].hidden:
            seen.setdefault(lab, []).append(r)
    for key, lab, nth, scale in TIE:
        rows[key] = (seen[lab][nth], scale)
    return rows


def tie_out(runs=RUNS):
    """Workbook (LibreOffice) vs shadow for every case, per (sensitivity, structure): {(run, case, key): (wb, sh)}."""
    from fepm.recalc import Office
    import openpyxl
    cells = dashboard_cells()
    ws = openpyxl.load_workbook(XLSX)["Dashboard"]
    srow = next(r for r in range(1, 40) if str(ws.cell(r, 3).value).startswith("Revenue structure"))
    res = {}
    with Office(XLSX) as lo:
        for sens, struct in runs:
            lo.set("Dash_Case", sens)
            for col in "EFG":
                lo.set(f"Dashboard!{col}{srow}", struct)
            lo.calculate()
            mc = lo.get("Master_Check")
            run = sens if struct == "Fixed + share" else f"{sens}, {struct}"
            for i, case in enumerate(CASES):
                sh = V2.run(case, sens, structure=struct)
                for key, (r, scale) in cells.items():
                    w = lo.get(f"Dashboard!{'EFG'[i]}{r}")
                    if sh[key] is None:                      # IRR undefined (no sign change): workbook shows n/a
                        assert w == "n/a", (sens, case, key, w)
                        continue
                    res[(run, case, key)] = (w, sh[key] / scale)
            res[(run, "master", "check")] = (mc, 0)
    return res


# =========================================================================================== report
def fmt(v, unit):
    if v is None:
        return "–"
    if unit == "%":
        return f"{v * 100:.2f} %"
    if unit == "x":
        return f"{v:.3f}x"
    if unit == "m":
        return f"{v:,.3f}"
    return f"{v:,.1f}"


def delta(a, b, unit):
    if a is None or b is None:
        return "–"
    d = b - a
    if unit == "%":
        return f"{d * 100:+.2f} pp"
    if unit == "x":
        return f"{d:+.3f}x"
    return f"{d:+,.3f}" if unit == "m" else f"{d:+,.1f}"


def report(tie):
    tg, rows3 = V3.targets(), V3.run()
    ties3 = sum(1 for k, t in tg.items() if rows3.get(k) is not None and abs(rows3[k] - t) <= 1e-6 * max(1, abs(t)))
    L = ["# Zalahaláp – reconciliation (SPEC v2 §9)", "",
         "Original: `reference/zalahalap/Zalahalap_model_10112025_v3.xlsb` (Negotiated case). New: "
         "`out/Zalahalap_Model.xlsx` (SPEC v2: PV only | PV + 2 h | PV + 4 h, unlevered, 5-year PV PPA). Written by "
         "`reconcile.py`.", "",
         "Three independent calculations: the v3 shadow `data/old_engine/old_zhlp.py`, the v1-structure shadow in "
         "`reconcile.py` (quarterly, SPEC v1, with a switch per §13 fix) and the v2 shadow `shadow.py` (annual, coded "
         "from SPEC v2). The workbook is recalculated in LibreOffice.", "",
         "## 1. The v3 shadow ties to the original", "",
         f"`old_zhlp.py` reproduces {ties3} of {len(tg)} targets of `data/old_engine/old_targets.csv` to 1e-6 "
         "(HTM targets from the copy with Aurora_Curve!C8 = 0, SPEC v1 §3.2).", "",
         "## 2. Bridge: v3 → PV + 4 h column, Base, merchant BESS basis", "",
         "Rows 1–2 are values; each further row switches one change on (cumulative) and shows its effect. Rows 2–10 run on the v1 quarterly shadow (unlevered: gearing 0 "
         "throughout), rows 11–14 on the frozen v2.0 annual shadow (`data/shadow_v20.py`), rows 15–20 on the v2.1+ shadow (the PPA uplift row is a side "
         "row against Merchant 80); the hold IRR is levered in the levered row and the last row; the last row is the workbook's PV + 4 h column (Base). The "
         "implied premium is shown before the 0 floor; the workbook floors it at 0.", "",
         "| Step | SPEC | " + " | ".join(lbl for _, lbl, _ in BRIDGE) + " |", "|---|---|" + "---:|" * len(BRIDGE)]
    prev = None
    for k, (label, ref, vals) in enumerate(bridge()):
        if k < 2:
            L.append(f"| {label} | {ref} | " + " | ".join(fmt(vals.get(x), u) for x, _, u in BRIDGE) + " |")
        else:
            L.append(f"| + {label} | {ref} | " + " | ".join(delta(prev.get(x), vals.get(x), u) for x, _, u in BRIDGE)
                     + " |")
        if not label.startswith(SIDE):                     # a side row does not move the base of the next step
            prev = vals
    key = {"implied_raw": "implied_rtb", "hold_irr": "hold_irr_l"}
    wb = {x: tie[("Base", "PV + 4h BESS", key.get(x, x))][0] for x, _, _ in BRIDGE}
    L.append("| **PV + 4 h, Base: the workbook (implied premium floored at 0)** | | "
             + " | ".join(fmt(wb[x], u) for x, _, u in BRIDGE) + " |")
    L += ["", "Notes on the steps:", "",
          "- Row 2 vs row 1: the v1 quarterly structure (period-end flows, all project cash in the project IRR, FEP "
          "taxed on the 12 % return, PSH land paid back); documented in SPEC v1 §13 and the v1 reconciliation.",
          "- §13.6 (Aurora's 4 h block total instead of v3's 0.591-scaled stack) lifts the hold IRR; §13.16–17 then "
          "price the PV at Aurora's direct and transfer prices (v3: 65 % of baseload on all output) and cut it back.",
          "- The v2 dates move the exit five months later and the RtB XIRR falls accordingly; the SPV price 549 (to "
          "FEP before the split, no return) and the DSA loan lower the RtB profit through the split.",
          "- Annual timeline step: year-end dates on the grid, construction by calendar month share, the pre-COD cash "
          "dated (RtB route) and mapped by year, the buyer reimbursements inside the JV split, guarantee fees paid at "
          "release, deposits returned at COD on the hold route, no debt, one SPV tax.",
          "- BESS opex step (fix round M2): v3's 30.8 k€/MW is Aurora's 2 h figure (Control!D96 note). Aurora's own "
          "table (Inputs rows 37–39, 2028 entry) gives 2 h 30.83 and 4 h 42.96 €/kW, i.e. 18.70 k€/MW + 6.07 k€/MWh; "
          "the 4 h battery now carries 42.96 instead of 30.8.",
          "- Timeline rebase (v2.1 §12.1): spend to 31-Dec-2026 in one column at its year end, all construction capex "
          "in 2027 (was spread NTP 30-Jun-27 → COD 31-Mar-28), COD 01-Mar-2028 (10 months of production in 2028), "
          "tolling from the later of 01-Jan-28 and COD.",
          "- Levered row: 60 % of the capex to COD drawn in 2027, IDC on mid-year draws capitalised, 20 level "
          "annuities from 31-Dec-2028, interest deductible; the hold IRR is the levered FEP IRR.",
          "- PPA 80/8 step (v2.2): 80 €/MWh nominal flat on all PV output from COD (Mar-2028) to Feb-2036, then "
          "Aurora merchant; no balancing in the PPA years.",
          "- O&M 10 step (v2.3): O&M 10 k€/MWp/yr all-in, real 2025 × the euro index (v3: 14.96 + MV cable 5).",
          "- PPA step: 45 €/MWh nominal on all PV output for 5 years (Aurora's direct price is 98–108 €/MWh nominal in "
          "2028–29 for the 4 h block), no balancing in those years.", "",
          "## 3. The v2 shadow ties to the workbook", "",
          "Workbook recalculated in LibreOffice: Base, Low, High on the default structure (Fixed + share) and Base on "
          "the PPA-uplift and Merchant structures; every Dashboard result below compared with `shadow.py`.", "",
          "| Sensitivity | Master check | Results compared | Largest absolute difference | Largest relative difference |",
          "|---|---:|---:|---:|---:|"]
    for sens in sorted({k[0] for k in tie}, key=lambda x: (len(x), x)):
        diffs = [(abs(w - s), abs(w - s) / max(1e-9, abs(s))) for (se, c, k), (w, s) in tie.items()
                 if se == sens and c != "master"]
        L.append(f"| {sens} | {tie[(sens, 'master', 'check')][0]:g} | {len(diffs)} | {max(d for d, _ in diffs):.2e} | "
                 f"{max(r for _, r in diffs):.2e} |")
    L += ["", "Base results (workbook):", "", "| Result | " + " | ".join(CASES) + " |", "|---|---:|---:|---:|"]
    units = {k: ("%" if "irr" in k or k.startswith("buyer") else "x" if "moic" in k else "k") for k, *_ in TIE}
    route = lambda k: ("Sale at RtB: " if k.startswith(("rtb_", "psh_rtb", "buyer_rtb", "implied_rtb")) else  # noqa
                       "Sale at COD: " if k.startswith(("cod_", "psh_cod", "buyer_cod", "implied_cod", "ev")) else
                       "Hold: " if k.startswith(("hold_", "proj_")) else "")
    for key, lab, nth, scale in TIE:
        unit = " (GWh)" if scale == 1000 else ""
        L.append(f"| {route(key)}{lab}{unit} | " + " | ".join(fmt(tie[("Base", c, key)][0], units[key]) for c in CASES)
                 + " |")
    L += ["", "## 4. Each §13 change inside the v3 structure (v3 shadow, cumulative; appendix)", "",
          "Values in € m (profit, NPV, lifetime totals).", "",
          "| Step | " + " | ".join(lbl for _, lbl, _ in OLD_OUTS) + " |", "|---|" + "---:|" * len(OLD_OUTS)]
    prev = None
    for label, vals in old_bridge():
        if prev is None:
            L.append(f"| {label} | " + " | ".join(fmt(vals[k], u) for k, _, u in OLD_OUTS) + " |")
        else:
            L.append(f"| + {label} | " + " | ".join(delta(prev[k], vals[k], u) for k, _, u in OLD_OUTS) + " |")
        prev = vals
    L.append("")
    (OUT / "reconciliation.md").write_text("\n".join(L))
    return OUT / "reconciliation.md"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--targets", action="store_true", help="write data/recon_targets.csv only")
    a = ap.parse_args()
    path, _ = write_targets()
    print(f"wrote {path}")
    if a.targets:
        return
    tie = tie_out()
    worst = max(abs(w - s) / max(1e-9, abs(s)) for (se, c, k), (w, s) in tie.items() if c != "master")
    print(f"tie-out: largest relative difference {worst:.2e}")
    print(report(tie))


if __name__ == "__main__":
    main()
