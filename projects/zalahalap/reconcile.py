"""Zalahaláp – reconciliation of Scenario A (Calc_1) to the original model (SPEC §3.2, §13).

    python3 projects/zalahalap/reconcile.py            → data/recon_targets.csv and out/reconciliation.md
    python3 projects/zalahalap/reconcile.py --targets  → data/recon_targets.csv only (no workbook needed)

Three independent pieces:
  1. data/old_engine/old_zhlp.py – Python shadow of the original v3 model (monthly RtB route, quarterly HTM). It ties
     to every target of data/old_engine/old_targets.csv; its switches apply the §13 fixes inside the v3 structure.
  2. Shadow – this file: an independent quarterly Python engine of the new structure (same rules as the workbook,
     coded from the SPEC and the v3 inputs, never read from the workbook). Its Scenario A values in the reconciliation
     set (set 2, every fix off) are stored in data/recon_targets.csv; build.py writes them to Curves and the Checks
     sheet tests that the workbook ties to them (0.1 %, 0.1 pp).
  3. The workbook itself (out/Zalahalap_Model.xlsm, private LibreOffice copy): set 2, then each fix and user decision
     switched on in turn (cumulative), with the solvers re-run, ending at the base set.
Stdlib + fepm.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import shutil
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(HERE / "data" / "old_engine"))

import old_zhlp as V3  # noqa: E402

DATA = HERE / "data"
OUT = HERE / "out"
XLSM = OUT / "Zalahalap_Model.xlsm"
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


# =========================================================================================== workbook bridge
OUTS = [  # key, label, unit, (column, Calc_1 row label)
    ("rtb_profit", "FEP profit – RtB", "k", ("I", "FEP cash flow – sale at RtB")),
    ("rtb_irr", "FEP IRR – RtB", "%", ("E", "FEP IRR – sale at RtB, after tax")),
    ("rtb_moic", "FEP MOIC – RtB", "x", ("E", "FEP MOIC – sale at RtB (receipts after tax ÷ outlays)")),
    ("htm_irr", "FEP IRR – HTM", "%", ("E", "FEP IRR – hold to maturity, levered, after tax")),
    ("htm_npv", "FEP NPV – HTM", "k", ("E", "FEP NPV – hold to maturity")),
    ("proj_irr", "Project IRR", "%", ("E", "Project IRR, unlevered, after tax")),
    ("buyer_irr", "Buyer IRR at 90", "%", ("E", "Buyer IRR at RtB, unlevered, after tax")),
    ("life_revenue", "Lifetime revenue", "k", ("I", "Revenue (E: first full year)")),
    ("life_ebitda", "Lifetime EBITDA", "k", ("I", "EBITDA (E: first full year)")),
]
TARGET_MAP = {  # v3 target key -> (column, Calc_1 row label, scale to the target unit), or None (no new-model row)
    "rtb_xirr_incl": ("E", "FEP IRR – sale at RtB, after tax", 1),
    "rtb_xirr_excl": None,
    "rtb_moic_incl_dep": ("E", "FEP MOIC – sale at RtB (receipts after tax ÷ outlays)", 1),
    "rtb_moic_excl_dep": None, "calc_moic_incl_dep": None, "calc_moic_excl_dep": None,
    "rtb_net_profit": ("I", "FEP cash flow – sale at RtB", 1000),
    "rtb_capital": ("I", "FEP outlays – RtB route", -1000),
    "rtb_proceeds": ("I", "FEP receipts – RtB route", 1000),
    "developer_profit": ("I", "PSH cash flow – sale at RtB", 1),
    "rtb_breakeven": None,
    "htm_xirr": ("E", "FEP IRR – hold to maturity, levered, after tax", 1),
    "htm_moic": ("E", "FEP MOIC – hold to maturity", 1),
    "htm_npv": ("E", "FEP NPV – hold to maturity", 1000),
    "htm_capital": ("E", "FEP capital invested – hold to maturity (outflows)", 1000),
    "htm_payback": None,
    "project_irr": ("E", "Project IRR, unlevered, after tax", 1),
    "buyer_irr": ("E", "Buyer IRR at RtB, unlevered, after tax", 1),
    "buyer_moic": None, "buyer_npv": None,
    "ffy_gwh": ("E", "PV generation after availability and degradation (E: first full year)", 1000),
    "ffy_price": None,
    "ffy_revenue": ("E", "Revenue (E: first full year)", 1000),
    "life_pv_revenue": ("I", "PV revenue: direct export and energy sent to the BESS (E: first full year)", 1000),
    "life_goo": ("I", "Guarantee of origin revenue", 1000),
    "life_revenue": ("I", "Revenue (E: first full year)", 1000),
    "life_ebitda": ("I", "EBITDA (E: first full year)", 1000),
    "life_margin": None,
    "life_capex": ("E", "Remaining construction capex (E: total incl. early works, F: of which BESS)", 1000),
    "life_opex": None,
    "life_bess_net": ("I", "BESS revenue (E: first full year)", 1000),
}
WHY = {  # target key -> why the new structure differs in the reconciliation set (SPEC §13)
    "rtb_xirr_incl": "quarterly grid (§13.12): the acquisition flows move from 30-Nov-25 to 31-Dec-25 and the sale "
                     "from 31-Oct-26 to 31-Dec-26, the pre-RtB costs to quarter ends; the profit ties",
    "rtb_moic_incl_dep": "gross receipts ÷ gross outlays as v3; deposits counted when paid and refunded (v3: peak "
                         "balance), guarantee fees and the return on the quarterly grid",
    "rtb_net_profit": "FEP now pays 9 % CIT on the 12 % return it receives (fix round 1, D7); v3 had no tax at FEP",
    "rtb_capital": "guarantee fees run to the end of the exit quarter (31-Dec-26; v3: the sale month) on "
                   "quarter-end balances; deposits equal (1,776)",
    "rtb_proceeds": "mirror of capital and profit (proceeds before FEP tax)",
    "developer_profit": "v3 deducts the PSH land apport (150) from the price but never credits it to PSH "
                        "(Calculation F225); the new model pays it back to PSH",
    "htm_xirr": "no working capital, linear construction spread (v3: buyer-model monthly curve), deposits and "
                "guarantees by line, PSH buy-out at the exit, tax with development costs capitalised (§13.10)",
    "htm_moic": "as the HTM IRR",
    "htm_npv": "as the HTM IRR; discounted to 31-Dec-25 in both",
    "htm_capital": "as the HTM IRR",
    "project_irr": "§13.4: includes all project cash (deposits, guarantees, pre-RtB costs, early works when paid); "
                   "no working capital; tax base incl. development costs",
    "buyer_irr": "buyer pays the price, early works, cash deposits and pre-RtB costs at the exit (v3: price and early "
                 "works only), bears guarantee fees to COD and gets the deposits back at COD",
    "ffy_gwh": "degradation on elapsed years at the period start with a 365-day year (v3: 366 in leap years)",
    "ffy_revenue": "as generation; no working-capital timing",
    "life_pv_revenue": "operations end 31-Mar-2063 = COD + 35 years (v3: 30-Apr-2063, one month more)",
    "life_goo": "as PV revenue", "life_revenue": "as PV revenue", "life_ebitda": "as PV revenue (opex rules identical)",
    "life_capex": "same total (57,170)", "life_bess_net": "same v3 stack (fix 6 off); one month shorter life",
}
STEPS = [  # label, Inputs names whose set-2 override is removed (cumulative), SPEC reference
    ("Fix 2: pre-RtB costs reimbursed as spent", ["In_Fix_Reimb"], "§13.2"),
    ("Fix 3: no bank-guarantee phantom", ["In_Fix_BG"], "§13.3"),
    ("Fix 5: PV export clipped (2 %)", ["In_Fix_Clip"], "§13.5"),
    ("Fix 6: Aurora 4 h block total for the BESS", ["In_Fix_BESS"], "§13.6"),
    ("Fix 7: augmentation indexed, Aurora repower year", ["In_Fix_Augment"], "§13.7"),
    ("Fix 8: dated act/365 return, 31-Dec capitalisation", ["In_Fix_Interest"], "§13.8"),
    ("User: share step (60 % at 90)", ["In_Share_Rule"], "§13.14, §0.4"),
    ("User: early works, four lines; E48/E49 at NTP", ["In_Milestones_Scope"], "§13.15, §0.3, §0.5"),
    ("Tax: HIPA and innovation contribution", ["In_HIPA", "In_Innovation"], "§13.10"),
    ("Tax defaults: BESS 10 years, losses lapse after 5 years", ["In_Dep_BESS_Years", "In_Loss_Years"], "§15.9"),
    ("JV sheet 221025: SPV 549, DSA loan, no PSH land step", ["In_SPV_Price", "In_SPV_Interest", "In_DSA_Per_MWp",
                                                               "In_PSH_Land_Step", "In_EW_To_PSH"], "§7, §15.10"),
    ("Debt: 55 % gearing (Bana terms)", ["In_Gearing"], "§13.9, §15.1"),
]
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


def fmt(v, unit):
    if v is None:
        return "–"
    if isinstance(v, str):
        return v
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
    return f"{d:+,.1f}" if unit == "k" else f"{d:+,.3f}"


def row_index(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True)["Calc_1"]
    idx = {}
    for r, row in enumerate(ws.iter_rows(min_col=3, max_col=3, values_only=True), 1):
        if row[0] and row[0] not in idx:
            idx[row[0]] = r
    return idx


def input_rows(path, names):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    return {n: int(wb.defined_names[n].attr_text.split("$")[-1]) for n in names}


def workbook_bridge():
    from fepm.recalc import Office
    idx = row_index(XLSM)
    names = sorted({n for _, ns, _ in STEPS for n in ns})
    rows = input_rows(XLSM, names)
    snap = lambda lo: {k: lo.get(f"Calc_1!{c}{idx[lab]}") for k, _, _, (c, lab) in OUTS}  # noqa: E731
    tmap = lambda lo: {k: (lo.get(f"Calc_1!{c}{idx[lab]}") / s if m else None)  # noqa: E731
                       for k, m in TARGET_MAP.items() for c, lab, s in [m or ("E", None, 1)]}
    res = []
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp) / XLSM.name
        shutil.copy(XLSM, work)
        with Office(work) as lo:
            lo.set("Scenario_Name", "Reconciliation (v3 rules)")
            lo.run_macro("Scenarios", "RunSolvers")
            set2 = tmap(lo)
            checks2 = lo.get("Model_Checks")
            res.append(("Workbook, set 2 (Reconciliation, v3 rules)", snap(lo)))
            for label, ns, _ in STEPS:
                for n in ns:
                    lo._range(f"Inputs!I{rows[n]}").getCellByPosition(0, 0).setFormula("")
                lo.run_macro("Scenarios", "RunSolvers")
                res.append((label, snap(lo)))
            lo.set("Scenario_Name", "Base")
            lo.run_macro("Scenarios", "RunSolvers")
            base = snap(lo)
            base_b = {k: lo.get(f"Calc_2!{c}{idx[lab]}") for k, _, _, (c, lab) in OUTS}
            base_b["implied_rtb"] = lo.get(f"Calc_2!E{idx[IMPLIED]}")
            checks1 = lo.get("Model_Checks")
    return set2, res, base, checks2, checks1, base_b


IMPLIED = "Implied RtB price at the target IRR, floored at 0 (E), closed form (F)"
BLK = "2027 solar DC 1.0cy 4h"
FIX1 = [  # case B, cumulative from the build before fix round 1 (shadow switches)
    ("v1 build, 5b343d2 (shadow)", dict(LEGACY)),
    ("M1: no grid charging, PV-only factor by component", dict(grid_import=0)),
    ("M2: two PV prices (direct export, transfer to the BESS)",
     dict(pv_curve=("direct_price", BLK), pv_tr_curve=("transfer_price", BLK))),
    ("Clipping 0 % for the DC-coupled case", dict(clip=0.0)),
    ("D6: DSA and SPV price out of the SPV tax base", dict(dep_sh=False)),
    ("Unlevered tax in project, buyer and implied price", dict(unlev_tax=True)),
    ("D7: FEP taxed on the 12 % return", dict(int_tax=True)),
]
FIX1_OUTS = [("htm_irr", "FEP IRR – hold", "%"), ("htm_irr_pre", "FEP IRR – hold, pre-tax", "%"),
             ("proj_irr", "Project IRR", "%"), ("buyer_irr", "Buyer IRR at 90", "%"),
             ("implied_raw", "Implied RtB price before the 0 floor (k€/MWp)", "k"),
             ("rtb_profit", "FEP profit – RtB (k€)", "k"), ("first_rev", "Revenue, first full year (k€)", "k")]


def fix1_bridge():
    over, out = {}, []
    for label, sw in FIX1:
        over.update(sw)
        out.append((label, run(params_b(**over))))
    return out


def old_bridge():
    block = {y: v for y, v in curves()[("storage_total", "2027 solar DC 1.0cy 4h")].items()}
    over, out = {}, [("v3 original (shadow = targets)", V3.run())]
    for label, sw in OLD_STEPS:
        over.update(sw)
        out.append((label, V3.run(**over, bess_block=block, aug_year_new=2048)))
    return out


def report(set2, res, base, checks2, checks1, base_b):
    tg, rows3 = V3.targets(), V3.run()
    tol = lambda t, x, unit: abs(x - t) <= (0.001 if unit == "%" else 0.001 * max(1, abs(t)))  # noqa: E731
    units = {r["key"]: r["unit"] for r in csv.DictReader(open(DATA / "old_engine" / "old_targets.csv"))}
    sh2 = run(params(**SET2))
    L = ["# Zalahaláp – reconciliation of Scenario A to the original model", "",
         "Original: `reference/zalahalap/Zalahalap_model_10112025_v3.xlsb` (Negotiated case), targets in "
         "`data/old_engine/old_targets.csv` ([R] = from the copy with Aurora_Curve!C8 = 0, SPEC §3.2). New: "
         "`out/Zalahalap_Model.xlsm`, Calc_1 = case A. Written by `reconcile.py`.", "",
         "Three independent calculations: the v3 shadow `data/old_engine/old_zhlp.py` (ties to every target), the "
         "new-structure shadow in `reconcile.py` (quarterly, coded from the SPEC; stored in `data/recon_targets.csv` "
         "and tested on the Checks sheet), and the workbook recalculated in LibreOffice with the solvers re-run.", "",
         f"Model checks: set 2 = {checks2:g}, base = {checks1:g}. Tolerance 0.1 % on values, 0.1 pp on rates.", "",
         "## 1. The original targets", "",
         "| Target | Unit | v3 cell | v3 target | v3 shadow | Shadow ties | New model, set 2 | Ties | Why it differs |",
         "|---|---|---|---:|---:|---|---:|---|---|"]
    cells = {r["key"]: r["cell"] + (" [R]" if r["basis"] == "R" else "")
             for r in csv.DictReader(open(DATA / "old_engine" / "old_targets.csv"))}
    for k, t in tg.items():
        u = {"%": "%", "x": "x"}.get(units[k], "m" if units[k] == "EUR m" else "k")
        x3 = rows3.get(k)
        ok3 = "yes" if x3 is not None and abs(x3 - t) <= 1e-6 * max(1, abs(t)) else "NO"
        xn = set2.get(k)
        if xn is None:
            L.append(f"| {k} | {units[k]} | {cells[k]} | {fmt(t, u)} | {fmt(x3, u)} | {ok3} | – | – | "
                     "no equivalent row (by design, §13.13: one RtB return, one break-even = implied price) |")
            continue
        ok = "yes" if tol(t, xn, u) else "no"
        L.append(f"| {k} | {units[k]} | {cells[k]} | {fmt(t, u)} | {fmt(x3, u)} | {ok3} | {fmt(xn, u)} | {ok} | "
                 f"{WHY.get(k, '') if ok == 'no' else ''} |")
    L += ["", "## 2. Each §13 change inside the v3 structure (v3 shadow, cumulative)", "",
          "| Step | " + " | ".join(lbl for _, lbl, _ in OLD_OUTS) + " |", "|---|" + "---:|" * len(OLD_OUTS)]
    prev = None
    for label, vals in old_bridge():
        if prev is None:
            L.append(f"| {label} | " + " | ".join(fmt(vals[k], u) for k, _, u in OLD_OUTS) + " |")
        else:
            L.append(f"| + {label} | " + " | ".join(delta(prev[k], vals[k], u) for k, _, u in OLD_OUTS) + " |")
        prev = vals
    L.append(f"| **v3 structure, all switchable fixes** | " + " | ".join(fmt(prev[k], u) for k, _, u in OLD_OUTS)
             + " |")
    L += ["", "Values in € m (profit, NPV, lifetime totals). At the price of 90 the share step changes nothing "
          "(both rules give 60 %, user 22-Oct-2026).", "",
          "## 3. New structure: reconciliation set → base set (workbook, cumulative)", "",
          "Values in € k (profit, NPV, lifetime totals). Each step removes the set-2 override of the inputs named "
          "(the set then inherits the base value) and re-runs Solver.bas.", "",
          "| Step | SPEC | " + " | ".join(lbl for _, lbl, _, _ in OUTS) + " |", "|---|---|" + "---:|" * len(OUTS)]
    prev = None
    spec = {lbl: s for lbl, _, s in STEPS}
    for label, vals in res:
        if prev is None:
            L.append(f"| {label} | §10 | " + " | ".join(fmt(vals[k], u) for k, _, u, _ in OUTS) + " |")
        else:
            L.append(f"| + {label} | {spec[label]} | " + " | ".join(delta(prev[k], vals[k], u) for k, _, u, _ in OUTS)
                     + " |")
        prev = vals
    L.append("| **Workbook, base set** | | " + " | ".join(fmt(base[k], u) for k, _, u, _ in OUTS) + " |")
    resid = max(abs((base[k] or 0) - (prev[k] or 0)) for k, _, _, _ in OUTS)
    L += ["", f"Last step vs the base set: largest difference {resid:.2e} (the steps restore every set-2 override).",
          "", "## 4. Independent shadow vs the workbook (set 2, Calc_1)", "",
          "| Headline | Shadow | Workbook | Difference |", "|---|---:|---:|---:|"]
    for k, lbl, u, _ in OUTS:
        if k in sh2:
            L.append(f"| {lbl} | {fmt(sh2[k], u)} | {fmt(res[0][1][k], u)} | {sh2[k] - res[0][1][k]:.2e} |")
    L += ["", "## 5. Fix rounds 1–2 (22-Oct-2026): case B bridge, v1 → new base (shadow, cumulative)", "",
          "Each step switches one change on in the independent shadow; the first line equals the v1 workbook (commit "
          "5b343d2: hold 10.72 %, implied RtB 165.4 k€/MWp, buyer IRR 8.83 %), the last line is checked against the "
          "new workbook (Calc_2, base set). The DSRA fix (COD sale only), D1 and D2 do not move these columns.", "",
          "| Step | " + " | ".join(lbl for _, lbl, _ in FIX1_OUTS) + " |", "|---|" + "---:|" * len(FIX1_OUTS)]
    prev = None
    for label, vals in fix1_bridge():
        if prev is None:
            L.append(f"| {label} | " + " | ".join(fmt(vals[k], u) for k, _, u in FIX1_OUTS) + " |")
        else:
            L.append(f"| + {label} | " + " | ".join(delta(prev[k], vals[k], u) for k, _, u in FIX1_OUTS) + " |")
        prev = vals
    L.append("| **Shadow, new base** | " + " | ".join(fmt(prev[k], u) for k, _, u in FIX1_OUTS) + " |")
    wbk = {"htm_irr": base_b["htm_irr"], "proj_irr": base_b["proj_irr"], "buyer_irr": base_b["buyer_irr"],
           "rtb_profit": base_b["rtb_profit"]}
    L.append("| **Workbook, Calc_2 base** | " + " | ".join(fmt(wbk.get(k), u) for k, _, u in FIX1_OUTS) + " |")
    L += ["", f"Implied RtB price in the workbook (floored at 0): {base_b['implied_rtb']:.1f} k€/MWp.",
          "", "## 6. Structural changes (§13 items without a switch)", "",
          "- §13.1: Aurora_Curve!C8 = 'no ' breaks the v3 HTM chain; the [R] targets come from the patched copy.",
          "- §13.4: the project IRR includes all project cash (v3 HTM_Model!D79 left out deposits, guarantees and "
          "pre-RtB costs).",
          "- §13.9: v3 has no debt; set 2 runs at 0 % gearing, the base at 55 % (Bana terms, SPEC §15.1).",
          "- §13.10: HIPA and the innovation contribution are off in set 2; tax on the sale gain per deal form "
          "(share deal, exempt; asset deal taxed); FEP pays CIT on the 12 % return as interest income. The SPV tax "
          "base also carries IDC and pre-RtB costs (v3: construction capex only); the DSA and the SPV price are "
          "shareholder-level and stay out. Unlevered views use an unlevered tax (no interest or fees).",
          "- §13.11: one FX (400) for the HUF deposits and land payments.",
          "- §13.12: one quarterly timeline from 01-Oct-2025 (v3: monthly RtB route, quarterly HTM); flows at "
          "period ends; the exit quarter belongs to the seller; the 12 % return on a payment starts in the next "
          "quarter (v3: in the month paid).",
          "- §13.13: no typed values on calc sheets, no data tables, one break-even (the implied price at the buyer "
          "target IRR), `n/a` plus a check instead of IFERROR(…,0).",
          "- Not in §13, by design of the new structure: no working capital (v3: 30 days receivable and payable), "
          "construction capex spread linearly over NTP → COD (v3: buyer-model monthly curve), case A COD = "
          "31-Mar-2028 (v3 01-Apr-2028 = operations from the same day), degradation on a 365-day year.",
          "", "## 7. v3 bugs not reproduced (SPEC §13)", "",
          "Summary!D28 → RTB_Model!D75 instead of D78; Sensitivities!E39:E42 use D73/D74; Calculation!G225 omits "
          "G221; Dyn_Inp!G24/H26:Y26 row offsets; Control!D26 overrides the construction-start default; stale notes. "
          "Also found: the developer's proceeds (Calculation F225) omit the 150 land apport deducted at F155.", ""]
    (OUT / "reconciliation.md").write_text("\n".join(L))
    return OUT / "reconciliation.md"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--targets", action="store_true", help="write data/recon_targets.csv only")
    a = ap.parse_args()
    path, sh = write_targets()
    print(f"wrote {path}")
    if a.targets:
        return
    print(report(*workbook_bridge()))


if __name__ == "__main__":
    main()
