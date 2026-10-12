"""Pure-Python shadow of the old Bora RtB engine (Bora_scenarios_meeting.xlsx: Tender §1-4, Deposits, Devex_FEP, MOIC).

Stdlib only. Inputs come from old_inputs_rtb.json, securities.csv and devex_lines.csv in this folder (extracted from the
workbook's cached values, case 1 = RtB sale). `compute(strategy, **overrides)` returns every Bora RtB output for
S0..S4 (S1 = current bid = price-effect reference, S2 = BORA 1, S3 = BORA 2). Defaults reproduce the old engine.
`python3 old_rtb.py` prints the comparison with the saved Bora values (Excel cache) and the LibreOffice live values
(lo_runs_rtb.json).

Units: k€ unless stated (Bora sheet shows €m = k€/1000), mHUF = HUF million. Sign convention of the waterfall rows as on
Bora rows 49-61: proceeds +, deductions -.

Toggles for the reconciliation bridge (keyword overrides of compute()); defaults = old engine:
  fx                  365   engine FX for securities / fees (Inputs E11)
  fx_devex            370   FX of the HUF-native devex lines (the pasted BORA_Devex sheet divides by its own AF6 = 370)
  devex_source        'engine_pasted_20260925' | 'devex_file_20261002' (line amounts/dates of the new devex method file)
  premium_basis       'net' (Inputs E149 = 0: premium x net MVA) | 'installed' (x installed wind MW, as the devex file V34)
  direct_fee_basis    'per_mw' (Inputs E84 = lump / net MVA, x successful net MW) | 'lump' (one sum per connection point)
  indirect_fee_basis  'net' (73.07 mHUF x net MVA) | 'installed' (x installed wind MW, as devex file I18)
  bid_netting         True  bid bond credited into the performance bond (Inputs E107)
  bid_release_date    None  None = bid guarantee stays outstanding until the sale (engine); a date releases it then
  pb_date             None  None = award + 17 days (Inputs E96 + E102 = 30-Dec-2026); else a date
  pb_hold_months_after_cod 6  holding of the PB after COD in the ÖSZ-6 price-effect item (Tender E31)
  interest_method     'line' (waterfall claims per line, Act/365 stubs, 31-Dec compounding, pause excluded: Deposits Y,
                      Devex_FEP S) | 'monthly' (MOIC memo schedules rows 111 / 82 / 126+130: monthly effective rate,
                      accrual flag row 10, flat 2x on the at-risk loan incl. the participation fee)
  same_year_stub_fix  False True = an item paid in the year the accrual ends gets a simple stub (devex-file quirk fixed)
  pause               True  interest pause 30-Jun-2029 -> 01-Jan-2030 (Inputs E64/E73); False = accrue to the sale date
  sale_date / cod / net_mva / rtb_premium / tender_mw  timing and size overrides
  dev_floor           True  developer share floored at -DSA (MOIC G154); False = negative distributable shared 70/30
  guarantee_fee       True  FEP pays the 6 % bank-guarantee fee monthly (MOIC row 79/163)
  signing_dd          True  signing fee 250 + DD 150 at month 0
  tax                 True  9 % CIT at sale with the participation exemption on the JV share (MOIC row 174)
  osz4_basis_installed None  None = strategy input (Bora row 10, all 0 = BESS ratio on net MVA); True = GYIK 20 rule
  pass_through        1.0   share of the promise PV passed into the RtB price (Tender E30)
Toggles added for the new-model reconciliation (SPEC v2 §8; defaults = old engine):
  pb_release_date     None  a date overrides COD + pb_hold_months_after_cod as the end of the PB carry (fix 4: 31-Mar-2033)
  pb_carry_tax        False True = the ÖSZ-6 PB carry in the price effect is after tax, x (1 - CIT) (SPEC v2 §5.6)
  n_bids              None  None = Inputs E82 (2 bids' participation fees charged to each bid); 1 = one fee per bid (fix 7)
  guarantee_fee_days  False True = guarantee fee on exact days outstanding, [posting, release), act/365 (fix 8)
  interest_method     'daily' = the new model's per-class accrual: simple act/365 within each calendar year on actual
                      days (leap years at /365), capitalised at 31-Dec, freeze excluded (fix 9)
  atrisk_floor        'line' (old: 2x per devex line) | 'portfolio' (MAX(accrued 20 %, 1 x principal), fix 9)
  waterfall           'full' (old: every claim deducted in full) | 'capped' (each step MIN(due, cash left), no
                      negative distributable, no developer floor; fix 6)
  margin_revenue      'merchant' (old case 1) | 'ppa_from_cod' (buyer margin on the PPA for its tenor from COD, then
                      merchant: the new model's single revenue line; fix 10)
  g_over              None  {general-input key: value} patched into G for this run (user decisions: turbine_eu_keur_per_mw,
                      bess_capex_keur_per_mw, ppa_price, ppa_tenor); they only reach the RtB route through the price effect
  strat_over          None  {strategy: {strategy-input key: value}} patched into J['strategies'] for this run
                      (user decision: BORA 2 community_eur_per_mwh 2.0 -> 2.01; also scores ÖSZ-2)
  grid                'monthly' (old MOIC grid) | 'quarterly' (new timeline: flows at quarter ends from 30-Sep-2026,
                      XIRR / NPV on those dates, price effect per quarter as the new Calc rows; fix 11)
"""
from __future__ import annotations

import csv
import datetime as dt
import json
import math
from pathlib import Path

HERE = Path(__file__).resolve().parent
STRATS = ['S0', 'S1', 'S2', 'S3', 'S4']
CODE = {'S1': 0, 'S2': 1, 'S3': 2, 'S4': 3, 'S0': 4}           # Tender E5 strategy code
BORA_COL = {'S0': 'D', 'S1': 'E', 'S2': 'F', 'S3': 'G', 'S4': 'H'}
BORA_YCOL = {'S0': 'Z', 'S1': 'AA', 'S2': 'AB', 'S3': 'AC', 'S4': 'AD'}


# ----------------------------------------------------------------------------------------------- date helpers
def D(x) -> dt.date:
    if isinstance(x, dt.date):
        return x
    return dt.date.fromisoformat(str(x)[:10])


def eomonth(d: dt.date, n: int) -> dt.date:
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    nxt = dt.date(y + (m == 12), m % 12 + 1, 1)
    return nxt - dt.timedelta(days=1)


def edate(d: dt.date, n: int) -> dt.date:
    m = d.month - 1 + n
    y, m = d.year + m // 12, m % 12 + 1
    last = eomonth(dt.date(y, m, 1), 0).day
    return dt.date(y, m, min(d.day, last))


def months_between(d: dt.date, d0: dt.date) -> int:          # (YEAR(d)-YEAR(d0))*12 + MONTH(d)-MONTH(d0)
    return (d.year - d0.year) * 12 + d.month - d0.month


def yearfrac_30360(d1: dt.date, d2: dt.date) -> float:       # Excel YEARFRAC basis 0 (US NASD 30/360)
    if d1 > d2:
        d1, d2 = d2, d1
    y1, m1, a = d1.year, d1.month, d1.day
    y2, m2, b = d2.year, d2.month, d2.day
    feb_end = lambda d: d.month == 2 and (d + dt.timedelta(days=1)).month == 3
    if feb_end(d1) and feb_end(d2):
        b = 30
    if feb_end(d1):
        a = 30
    if b == 31 and a >= 30:
        b = 30
    if a == 31:
        a = 30
    return ((y2 - y1) * 360 + (m2 - m1) * 30 + (b - a)) / 360.0


def fep_factor(r: float, pay: dt.date, end: dt.date, fix_same_year=False) -> float:
    """BORA_Devex / Deposits Z convention: simple Act/365 stubs counting both ends, compounding on 31-Dec.
    (1 + r*(31Dec(pay)-pay+1)/365) * (1+r)^(Y(end)-Y(pay)-1) * (1 + r*(end-1Jan(end)+1)/365); 1 if pay > end."""
    if pay > end:
        return 1.0
    if fix_same_year and pay.year == end.year:
        return 1 + r * ((end - pay).days + 1) / 365
    a = (dt.date(pay.year, 12, 31) - pay).days + 1
    b = (end - dt.date(end.year, 1, 1)).days + 1
    return (1 + r * a / 365) * (1 + r) ** (end.year - pay.year - 1) * (1 + r * b / 365)


def daily_factor(r: float, pay: dt.date, end: dt.date) -> float:
    """New-model convention: simple act/365 within each calendar year on the actual days (pay and end both counted,
    366 days in a leap year), capitalised on 31-Dec. 1 if pay > end."""
    if pay > end:
        return 1.0
    f = 1.0
    for y in range(pay.year, end.year + 1):
        a, b = max(pay, dt.date(y, 1, 1)), min(end, dt.date(y, 12, 31))
        f *= 1 + r * ((b - a).days + 1) / 365
    return f


def two_segment(r, pay, sale, pause_start, pause_end, pause=True, fix=False, daily=False):
    """Deposits X = Z x AB / Devex_FEP N x P: accrual pay -> MIN(sale, pause start), then MAX(pay, pause end) -> sale."""
    fac = (lambda r_, a, b, _f: daily_factor(r_, a, b)) if daily else fep_factor
    if not pause:
        return fac(r, pay, sale, fix)
    z = fac(r, pay, min(sale, pause_start), fix)
    start2 = max(pay, pause_end)
    ab = 1.0 if sale < start2 else fac(r, start2, sale, fix)
    return z * ab


def xnpv(rate, cfs, dates):
    d0 = dates[0]
    return sum(c / (1 + rate) ** ((d - d0).days / 365) for c, d in zip(cfs, dates))


def xirr(cfs, dates):
    """Exact root of XNPV = 0 (bisection on a bracket, then Newton polish). Excel's XIRR stops at ~1e-8."""
    f = lambda r: xnpv(r, cfs, dates)
    lo, hi = -0.99, 10.0
    flo, fhi = f(lo), f(hi)
    if flo * fhi > 0:
        return float('nan')
    for _ in range(300):
        mid = (lo + hi) / 2
        fm = f(mid)
        if flo * fm <= 0:
            hi, fhi = mid, fm
        else:
            lo, flo = mid, fm
        if hi - lo < 1e-15:
            break
    return (lo + hi) / 2


# ----------------------------------------------------------------------------------------------- inputs
def load_inputs():
    j = json.loads((HERE / 'old_inputs_rtb.json').read_text())
    g = {k: v['value'] for k, v in j['general'].items()}
    sec = list(csv.DictReader(open(HERE / 'securities.csv', newline='')))
    dev = list(csv.DictReader(open(HERE / 'devex_lines.csv', newline='')))
    return j, g, sec, dev


J, G, SEC, DEV = load_inputs()


def band_points(x, table, eligible=1):
    """Tender rows 53/54: band = Σ(lower < x) + (lower = x) x included; inner = INDEX(points incl. 'below', band+1)."""
    bands = table['bands']
    n = sum((lo < x) + (lo == x) * inc for lo, inc, _ in bands)
    pts = [table['below_points']] + [p for _, _, p in bands]
    return pts[int(n)] * eligible


def strategy_physical(s, net_mva):
    st = J['strategies'][s]
    ob = st['overbuild']
    gross = round(net_mva * (1 + ob), 6)                                  # Tender row 40
    bess = round(st['bess_ratio'] * (gross if st['bess_basis_installed'] == 1 else net_mva), 6)   # row 41
    clip = max(0, gross / net_mva - 1) / (1 + max(0, gross / net_mva - 1)) * G['bottleneck_loss_rate']  # rows 43, 45
    return dict(gross=gross, bess=bess, mwh=bess * st['bess_hours'], hours=st['bess_hours'], clip=clip,
                overbuild=gross / net_mva - 1, uplift=st['perf_security_multiple'], municipal=st['municipal_keur_per_mw'],
                community=st['community_eur_per_mwh'], eu=st['eu_turbine'])


def tender_points(s, net_mva, osz4_basis_installed=None):
    st = J['strategies'][s]
    tt = J['tender_tables']
    w, p = tt['weights'], tt['params']
    ph = strategy_physical(s, net_mva)
    out = {}
    x1 = round(st['municipal_keur_per_mw'] * 1000, 6)
    out['osz1'] = band_points(x1, tt['osz1'], int(p['osz1_min'] <= x1 <= p['osz1_max'])) * w['osz1'] / 100
    x2 = round(st['community_eur_per_mwh'], 6)
    out['osz2'] = band_points(x2, tt['osz2'], int(p['osz2_min'] <= x2 <= p['osz2_max'])) * w['osz2'] / 100
    x3 = round(ph['overbuild'] * 100, 6)
    out['osz3'] = band_points(x3, tt['osz3'], int(x3 >= 0)) * w['osz3'] / 100
    basis_inst = st['bess_basis_installed'] == 1 if osz4_basis_installed is None else osz4_basis_installed
    ratio = ph['bess'] / ph['gross'] if basis_inst else ph['bess'] / net_mva
    x4 = round(ratio * 100, 6)
    el4 = int(x4 >= p['osz4_min_ratio'] and st['bess_hours'] >= p['osz4_min_hours'] and round(ph['bess'], 6) <= round(ph['gross'], 6))
    out['osz4'] = band_points(x4, tt['osz4'], el4) * w['osz4'] / 100
    out['osz5'] = st['partial_award'] * p['osz5_points'] * w['osz5'] / 100
    out['osz6'] = band_points(round(st['perf_security_multiple'] * 100, 6), tt['osz6']) * w['osz6'] / 100
    out['osz7'] = st['brownfield'] * p['osz7_points'] * w['osz7'] / 100
    out['osz8'] = st['shared_connection'] * p['osz8_points'] * w['osz8'] / 100
    out['osz9'] = (p['osz9a_points'] * st['eu_turbine'] + p['osz9b_points'] * st['eu_transformer']) * w['osz9'] / 100
    sigma = sum(out[f'osz{i}'] for i in range(1, 9))
    out['total_eu'] = round(sigma + out['osz9'], 2)
    out['total_cn'] = round(sigma, 2)
    return out


# ----------------------------------------------------------------------------------------------- HTM generation (for the price effect)
def ppa_share_from_cod(y, cod, tenor):
    """Share of year y's operating months inside [COD, COD + tenor years): PPA from COD (new model)."""
    end = edate(cod, int(tenor) * 12)
    op_m = [(dt.date(y, m, 1) > cod) and dt.date(y, m, 1) <= edate(cod, 360) for m in range(1, 13)]
    ppa_m = [op and dt.date(y, m, 1) < end for op, m in zip(op_m, range(1, 13))]
    return sum(ppa_m) / sum(op_m) if sum(op_m) else 0.0


def htm_years(o, ph, cod, ntp):
    """HTM rows 5, 7, 9, 10, 38, 45, 47, 48, 55, 90 for 2026..2061 (columns H:AQ)."""
    y0 = o['investment_date'].year
    years = list(range(y0, y0 + 36))
    first_full, last = cod.year + 1, cod.year + int(G['operating_life'])
    aur = J['aurora_wind_capture']
    real = {y: v for y, v in zip(aur['years'], aur[['low_real', 'central_real', 'high_real'][int(G['aurora_scenario']) - 1]])}
    infl, base = G['aurora_inflation'], G['aurora_base_year']
    last_aur = max(real)
    hours = G['p50_hours'] * (1 - G['availability_haircut']) * (1 - G['curtailment']) * G['yield_factor'] * (1 - ph['clip'])
    rows = []
    for y in years:
        op = ((cod.year == y) * (12 - cod.month) / 12 + (first_full <= y < last) * 1 + (y == last) * cod.month / 12)
        idx = (y - first_full + 1) if (op > 0 and y >= first_full) else 0
        cons = max(0, (min(dt.date(y, 12, 31), cod) - max(dt.date(y, 1, 1) - dt.timedelta(days=1), ntp)).days) / (cod - ntp).days
        degr = (1 - G['degradation']) ** max(idx - 1, 0) if op > 0 else 0
        gen = ph['gross'] * (op > 0) * hours * degr * op / 1000                       # GWh
        if y in real:
            cap = real[y] * (1 + infl) ** (y - base) * G['price_factor']
        elif y > last_aur:
            cap = real[last_aur] * (1 + infl) ** (last_aur - base) * (1 + infl) ** (y - last_aur) * G['price_factor']
        else:
            cap = None
        est = G['est_baseload'] * (1 + G['est_baseload_esc']) ** (y - cod.year) * G['est_capture_rate']
        price = cap if (G['price_source'] == 1 and cap is not None) else est
        ppa = G['ppa_price'] * (1 + G['ppa_esc']) ** (y - cod.year)
        ppa_share = (1 - G['merchant_share']) if 1 <= idx <= G['ppa_tenor'] else 0
        if o.get('margin_revenue') == 'ppa_from_cod':
            ppa_share = ppa_share_from_cod(y, cod, G['ppa_tenor'])
        esc = (1 + G['opex_escalation']) ** (y - cod.year) if op > 0 else 0
        rows.append(dict(year=y, op=op, idx=idx, cons=cons, gen=gen, price=price, ppa=ppa, ppa_share=ppa_share, esc=esc))
    cod_year_op = sum(r['op'] for r in rows if r['idx'] == 0)
    n = G['tax_depreciation_years']
    for r in rows:
        if r['op'] == 0:
            r['dep'] = 0
        elif r['idx'] == 0:
            r['dep'] = r['op']
        elif r['idx'] < n:
            r['dep'] = 1
        elif r['idx'] == n:
            r['dep'] = 1 - cod_year_op
        else:
            r['dep'] = 0
    return rows


def price_effect(s, o, ph, ref, E18, E29, E20, E23, cod, ntp, sale):
    """Tender section 4 (rows 176-222): PV at the RtB sale date of the buyer's extra promise cash flows vs S1."""
    tax = G['cit'] + G['robin_hood_switch'] * G['robin_hood_rate']                   # Tender E180
    r_b = G['buyer_discount_rate']                                                     # E176
    pb_release = D(o['pb_release_date']) if o['pb_release_date'] else edate(cod, int(o['pb_hold_months_after_cod']))  # E179
    d_mun = (ph['municipal'] - ref['municipal']) * E18                                 # E190
    d_bess = (E20 - ref['bess']) * G['bess_capex_keur_per_mw']                         # E191
    pb_live = G['perf_bond_rate'] * ph['uplift'] * (E29 * G['bench_wind_mhuf_per_mw'] + E23 * G['bench_bess_mhuf_per_mwh'])     # E192 = Deposits H10
    pb_ref = G['perf_bond_rate'] * ref['uplift'] * (E29 * G['bench_wind_mhuf_per_mw'] + ref['bess'] * G['conversion_rate'] * ref['hours'] * G['bench_bess_mhuf_per_mwh'])  # E193
    d_pb = (pb_live - pb_ref) * 1000 / o['fx']                                         # E194
    d_comm = ph['community'] - ref['community']                                        # G186
    e205 = 0 if 1 - ph['clip'] == 0 else (ph['clip'] - ref['clip']) / (1 - ph['clip'])
    d_tur = (G['eu_turbine_price_switch'] * ref['gross'] * (ph['eu'] - ref['eu'])
             * (G['turbine_eu_keur_per_mw'] - G['turbine_cn_keur_per_mw']) * (1 + G['capex_contingency']))   # E213
    rows = htm_years(o, ph, cod, ntp)
    items = {k: 0.0 for k in ('municipal', 'community', 'bess', 'security', 'eu_turbine', 'overbuild')}
    flows = []
    for r in rows:
        y = r['year']
        df = 1 / (1 + r_b) ** (max(0, (dt.date(y, 12, 31) - sale).days) / 365)            # row 197
        depsh = r['dep'] / G['tax_depreciation_years']                                     # row 200
        yrs_pb = max(0, (min(dt.date(y, 12, 31), pb_release) - max(dt.date(y, 1, 1) - dt.timedelta(days=1), sale)).days) / 365   # row 203
        if r['op'] > 0:                                                                    # row 206
            margin = ((r['price'] * (1 - r['ppa_share']) + r['ppa'] * r['ppa_share'])
                      * (1 - G['mekh_fee_pct'] * (1 + G['opex_contingency']) - G['hipa_pct'])
                      - (G['balancing_eur_mwh'] + G['grid_eur_mwh']) * r['esc'] * (1 + G['opex_contingency']) - ph['community'])
        else:
            margin = 0
        f = {'municipal': -d_mun * (y == cod.year) + d_mun * depsh * tax,                 # row 209
             'community': -d_comm * r['gen'] * (1 - tax),                                  # row 210
             'bess': -d_bess * r['cons'] + d_bess * depsh * tax,                           # row 211
             'security': -d_pb * G['guarantee_rate'] * yrs_pb * (1 - tax * o['pb_carry_tax']),  # row 212
             'eu_turbine': -d_tur * r['cons'] + d_tur * depsh * tax,                       # row 213
             'overbuild': -(r['gen'] * e205) * margin * (1 - tax)}                         # rows 205, 214
        for k, v in f.items():
            items[k] += v * df
        flows.append(dict(year=y, df=df, **f))
    total = sum(items.values())                                                            # E218
    return dict(items=items, sum_pv=total, adjustment=o['pass_through'] * total, d_pb_keur=d_pb, flows=flows)


def quarters(start=dt.date(2026, 7, 1), n=141):
    """New-model timeline: (start, end) of each quarter from 01-Jul-2026 (141 periods to 30-Sep-2061)."""
    return [(edate(start, 3 * k), eomonth(start, 3 * k + 2)) for k in range(n)]


def price_effect_quarterly(o, ph, ref, E18, E29, E20, cod, ntp, sale):
    """Price effect on the new quarterly timeline (Calc rows pe_mun..pe_ob): flows at quarter ends, discounted to the exit
    at the buyer rate; generation degrades by operating year from COD; PPA for its tenor from COD, then Aurora."""
    tax, r_b = G['cit'], G['buyer_discount_rate']
    pb_release = D(o['pb_release_date']) if o['pb_release_date'] else edate(cod, int(o['pb_hold_months_after_cod']))
    aur = J['aurora_wind_capture']
    real = dict(zip(aur['years'], aur[['low_real', 'central_real', 'high_real'][int(G['aurora_scenario']) - 1]]))
    life_end, dep_end = edate(cod, 12 * int(G['operating_life'])), edate(cod, 12 * int(G['tax_depreciation_years']))
    c0, dep0 = ntp + dt.timedelta(days=1), cod + dt.timedelta(days=1)
    share = lambda s_, e_, a, b: max(0, ((min(b, e_) - max(a, s_)).days + 1)) / ((b - a).days + 1)  # noqa: E731
    loss_ref = ref['overbuild'] / (1 + ref['overbuild']) * G['bottleneck_loss_rate']
    loss = ph['clip']
    d_mun = (ph['municipal'] - ref['municipal']) * E18
    d_com = ph['community'] - ref['community']
    d_bess = (E20 - ref['bess']) * G['bess_capex_keur_per_mw']
    pb = G['perf_bond_rate'] * ph['uplift'] * (E29 * G['bench_wind_mhuf_per_mw'] + ph['mwh'] * G['bench_bess_mhuf_per_mwh'])
    pb_ref = G['perf_bond_rate'] * ref['uplift'] * (E29 * G['bench_wind_mhuf_per_mw'] + ref['bess'] * ref['hours'] * G['bench_bess_mhuf_per_mwh'])
    d_pb = (pb - pb_ref) * 1000 / o['fx']
    e_ob = (loss - loss_ref) / (1 - loss)
    items = {k: 0.0 for k in ('municipal', 'community', 'bess', 'security', 'eu_turbine', 'overbuild')}
    for s_, e_ in quarters():
        days = (e_ - s_).days + 1
        df = 1 / (1 + r_b) ** (max(0, (e_ - sale).days) / 365)
        constr = share(s_, e_, c0, cod)
        dep = share(s_, e_, dep0, dep_end)
        ops = max(0, (min(life_end, e_) - max(dep0, s_)).days + 1) / days
        cod_f = 1 if s_ <= cod <= e_ else 0
        opyr = math.ceil(yearfrac_30360(cod, e_) - 1e-12) if e_ > cod else 0
        y = e_.year
        cap = real.get(min(y, max(real)), 0.0) * (1 + G['aurora_inflation']) ** (y - G['aurora_base_year'])
        price = G['ppa_price'] * (1 + G['ppa_esc']) ** (y - cod.year) if opyr <= G['ppa_tenor'] else cap
        esc = (1 + G['opex_escalation']) ** (y - cod.year)
        gen = (ph['gross'] * G['p50_hours'] * (1 - G['curtailment']) * (1 - loss)
               * (1 - G['degradation']) ** (opyr - 1) * ops * days / 365) if ops else 0.0          # MWh
        rev = gen * price / 1000
        var = -(G['balancing_eur_mwh'] + G['grid_eur_mwh']) * gen * esc / 1000
        margin = (rev * (1 - G['mekh_fee_pct'] * (1 + G['opex_contingency']) - G['hipa_pct'])
                  + var * (1 + G['opex_contingency']) - ph['community'] * gen / 1000)
        pb_days = max(0, (min(pb_release, e_) - max(sale + dt.timedelta(days=1), s_)).days + 1)
        f = {'municipal': d_mun * (dep * tax - cod_f), 'community': -d_com * gen * (1 - tax) / 1000,
             'bess': d_bess * (dep * tax - constr),
             'security': -d_pb * G['guarantee_rate'] * pb_days / 365 * (1 - tax * o['pb_carry_tax']),
             'eu_turbine': 0.0, 'overbuild': -margin * e_ob * (1 - tax)}
        for k, v in f.items():
            items[k] += v * df
    total = sum(items.values())
    return dict(items=items, sum_pv=total, adjustment=o['pass_through'] * total, d_pb_keur=d_pb, flows=[])


# ----------------------------------------------------------------------------------------------- main engine
DEFAULTS = dict(
    fx=None, fx_devex=370.0, devex_source='engine_pasted_20260925', premium_basis='net', direct_fee_basis='per_mw',
    indirect_fee_basis='net', bid_netting=None, bid_release_date=None, pb_date=None, pb_hold_months_after_cod=None,
    interest_method='line', same_year_stub_fix=False, pause=True, sale_date=None, cod=None, net_mva=None, rtb_premium=None,
    tender_mw=None, dev_floor=True, guarantee_fee=True, signing_dd=True, tax=True, osz4_basis_installed=None, pass_through=None,
    n_bids=None, pb_release_date=None, pb_carry_tax=False, guarantee_fee_days=False, atrisk_floor='line', waterfall='full', margin_revenue='merchant', grid='monthly',
    g_over=None, strat_over=None,
)


def compute(strategy: str, **kw) -> dict:
    """Old RtB engine for one strategy; `g_over` / `strat_over` patch the module inputs for this call only."""
    g_over, strat_over = kw.pop('g_over', None), kw.pop('strat_over', None)
    saved_g = {k: G[k] for k in (g_over or {})}
    saved_s = {s_: {k: J['strategies'][s_][k] for k in v} for s_, v in (strat_over or {}).items()}
    G.update(g_over or {})
    for s_, v in (strat_over or {}).items():
        J['strategies'][s_].update(v)
    try:
        return _compute(strategy, **kw)
    finally:
        G.update(saved_g)
        for s_, v in saved_s.items():
            J['strategies'][s_].update(v)


def _compute(strategy: str, **kw) -> dict:
    o = dict(DEFAULTS)
    unknown = set(kw) - set(o)
    if unknown:
        raise TypeError(f'unknown override(s): {sorted(unknown)}')
    o.update(kw)
    o['fx'] = o['fx'] or G['fx']
    o['bid_netting'] = G['bid_bond_netting'] == 1 if o['bid_netting'] is None else o['bid_netting']
    o['pb_hold_months_after_cod'] = o['pb_hold_months_after_cod'] if o['pb_hold_months_after_cod'] is not None else G['pb_hold_months_after_cod']
    o['pass_through'] = G['pass_through'] if o['pass_through'] is None else o['pass_through']
    o['investment_date'] = inv = D(G['investment_date'])
    net = o['net_mva'] or G['net_mva']
    fx = o['fx']

    # ---- timing (Inputs rows 27, 28, 55-58, 152-155)
    cod_in = D(o['cod'] or G['cod'])
    m_rtb = months_between(cod_in, inv) - int(G['ntp_lag_months']) - int(G['construction_months'])    # E55
    rtb = eomonth(inv, m_rtb)                                                                          # E56
    sale_in = D(o['sale_date'] or G['rtb_sale_date'])
    lag = max(0, months_between(sale_in, rtb))                                                         # E28
    m_sale = m_rtb + lag                                                                               # E57
    sale = eomonth(inv, m_sale)                                                                        # E58
    ntp = eomonth(rtb, int(G['ntp_lag_months']))                                                       # E154
    cod = eomonth(ntp, int(G['construction_months']))                                                  # E155
    p_start, p_end = D(G['pause_start']), D(G['pause_end'])
    NM = 61                                                                                            # MOIC H:BP
    mdates = [eomonth(inv, m) for m in range(NM)]
    gm = lambda d: months_between(d, inv)                                                             # grid month

    # ---- strategy -> MW (Tender §2; Inputs E16-E29)
    ph = strategy_physical(strategy, net)
    ref = strategy_physical('S1', net)
    E16 = o['tender_mw'] or G['tender_mw']
    E18 = E16 * G['conversion_rate']
    E19 = E16 - E18
    E20 = ph['bess']
    E22 = E20 * ph['hours']
    E23 = E22 * G['conversion_rate']
    E29 = E18 if net == 0 else E18 * ph['gross'] / net
    E150 = E29 if o['premium_basis'] == 'installed' else E18

    # ---- Deposits schedule (rows 7-18; mHUF -> k€ at fx)
    award = D(G['award_date'])
    pb_date = D(o['pb_date']) if o['pb_date'] else award + dt.timedelta(days=int(G['pb_days_after_award']))
    sw = {r['key']: r for r in SEC}
    lines = []

    def add(key, mhuf, date, typ, cash=None, rec=None, share=None):
        s = sw.get(key, sw['bid_bond'])
        lines.append(dict(key=key, mhuf=mhuf, date=date, type=typ,
                          cash=float(s['cash_1_guarantee_0']) if cash is None else cash,
                          rec=float(s['recovered_by_buyer_pct']) if rec is None else rec,
                          share=float(s['recoverable_loan_share']) if share is None else share))
    fee = G['participation_fee_mhuf_per_bid'] * (o['n_bids'] or G['n_bids'])
    add('fee', fee, D(G['bid_date']), 'Fee')
    base_bid = E16 * (1 if E18 == 0 else E29 / E18) * G['bench_wind_mhuf_per_mw'] + E22 * G['bench_bess_mhuf_per_mwh']   # Inputs E91
    bid = G['bid_bond_rate'] * base_bid                                                                                  # H8
    add('bid_bond', bid, D(G['bid_date']), 'Security')
    add('bid_bond_refund_failed', -G['bid_bond_rate'] * E19 * G['bench_wind_mhuf_per_mw'] * (1 if E18 == 0 else E29 / E18)
        * G['failed_bid_bond_refund'], award, 'Refund')
    base_pb = E29 * G['bench_wind_mhuf_per_mw'] + E23 * G['bench_bess_mhuf_per_mwh']                                    # Inputs E92
    pb_gross = G['perf_bond_rate'] * ph['uplift'] * base_pb                                                              # H10
    add('perf_bond', pb_gross - (bid if o['bid_netting'] else 0), pb_date, 'Security')
    if o['bid_release_date']:
        add('bid_bond_release', -bid, D(o['bid_release_date']), 'Refund',
            cash=float(sw['bid_bond']['cash_1_guarantee_0']), rec=float(sw['bid_bond']['recovered_by_buyer_pct']),
            share=float(sw['bid_bond']['recoverable_loan_share']))
    tso = G['tso_security_mhuf_per_mw'] * E18
    add('tso_security', tso, D(G['connection_request_date']), 'Security')
    direct_lump = G['direct_fee_lump_huf'] / 1e6
    direct = direct_lump if o['direct_fee_basis'] == 'lump' else direct_lump / net * E18
    add('direct_fee', direct, D(G['hcssz_date']), 'Connection fee')
    mgt = D(G['mgt_date'])
    apr = dt.date(mgt.year, 4, 23)
    e104 = apr if apr >= mgt else dt.date(mgt.year + 1, 4, 23)
    permit = D(G['permit_date'])
    supp = []
    for key, share, months in (('supp_1', 0.4, 0), ('supp_2', 0.2, 12), ('supp_3', 0.2, 24), ('supp_4', 0.2, 36)):
        h = G['bench_wind_mhuf_per_mw'] * G['supp_security_rate'] * share * E29
        d = edate(e104, months)
        supp.append((key, h, d))
    refund = -G['supp_refund_ratio'] * sum(h for _, h, d in supp if d <= permit)
    for key, h, d in supp[:2]:
        add(key, h - h * G['supp_refund_ratio'] * (d > permit), d, 'Security')
    add('supp_refund', refund, permit, 'Refund', cash=float(sw['supp_1']['cash_1_guarantee_0']))
    for key, h, d in supp[2:]:
        add(key, h - h * G['supp_refund_ratio'] * (d > permit), d, 'Security')
    ind_mw = E29 if o['indirect_fee_basis'] == 'installed' else E18
    add('indirect_fee', G['indirect_fee_mhuf_per_mw'] * ind_mw - tso, D(G['hcssz_date']), 'Connection fee')
    for ln in lines:
        ln['keur'] = ln['mhuf'] * 1000 / fx                                                        # M
        ln['month'] = gm(ln['date'])                                                                # L
        after = ln['month'] > m_sale                                                                # N
        ln['incl'] = ln['keur'] * (G['include_after_sale'] if after else 1)                         # O
        ln['rate'] = G['r_recoverable'] if ln['cash'] == 1 else G['r_guarantee_interest']           # W
        ln['factor'] = two_segment(ln['rate'], ln['date'], sale, p_start, p_end, o['pause'], o['same_year_stub_fix'],
                                   o['interest_method'] == 'daily')                                  # X
        ln['interest'] = ln['incl'] * (ln['factor'] - 1)                                            # Y
    Y22 = sum(l['interest'] for l in lines if l['cash'] == 0)                                       # 6 % guarantees
    Y23 = sum(l['interest'] for l in lines if l['cash'] == 1 and l['type'] != 'Fee')                # 12 % cash
    Y24 = sum(l['interest'] for l in lines if l['type'] == 'Fee')                                    # fee 12 %
    fee_keur = next(l['incl'] for l in lines if l['key'] == 'fee')

    # ---- Devex_FEP (devex lines, land, contingency; 20 % with the 2x floor per line)
    src = o['devex_source']
    dl = []
    for r in DEV:
        if r['source'] != src or r['category'] not in ('Devex', 'Land', 'Excluded'):
            continue
        if r['native'] == 'MODEL' or (src == 'devex_file_20261002' and r['item'] == 'Contingency'):
            continue                                                                                  # contingency rebuilt below
        keur = (float(r['huf_native'] or 0) / o['fx_devex'] + float(r['eur_native'] or 0)) / 1000
        pd = r['pay_date']
        pd = dt.date(2027, 12, 31) if pd == 'Continuous' else D(pd)                                  # engine: continuous items 31-Dec-2027
        fep = 1 if r['category'] == 'Devex' else 0
        buyer = 1 if r['category'] == 'Land' else 0
        dl.append(dict(item=r['item'], row=r['devex_row'], keur=keur, date=pd, H=keur * fep, K=keur * buyer * (1 - fep)))
    cont = G['devex_contingency'] * (sum(x['H'] for x in dl) + sum(x['K'] for x in dl) + fee_keur)   # Devex_FEP E38
    dl.append(dict(item='Contingency', row='49', keur=cont, date=dt.date(2028, 12, 31), H=cont, K=0.0))
    for x in dl:
        x['M'] = x['H'] + x['K'] * (G['devex_basis_switch'] == 1)
        x['month'] = gm(x['date'])
        x['S_raw'] = two_segment(G['r_at_risk'], x['date'], sale, p_start, p_end, o['pause'], o['same_year_stub_fix'],
                                 o['interest_method'] == 'daily')
        x['S'] = max(x['S_raw'], G['min_multiple_at_risk']) if o['atrisk_floor'] == 'line' else x['S_raw']
    G18 = sum(x['H'] for x in dl if 0 <= x['month'] < NM)
    G20 = sum(x['K'] for x in dl if 0 <= x['month'] < NM) if G['devex_basis_switch'] == 1 else 0
    G131 = sum(x['M'] * (x['S'] - 1) for x in dl if x['month'] <= m_sale and 0 <= x['month'] < NM)
    if o['atrisk_floor'] == 'portfolio':
        G131 = max(G131, (G['min_multiple_at_risk'] - 1) * sum(x['M'] for x in dl if x['month'] <= m_sale and 0 <= x['month'] < NM))

    # ---- DSA (MOIC rows 22-25)
    dsa_start = D(G['dsa_start'])
    dsa_mw = min(G['dsa_cap_keur_per_mw'], G['dsa_keur_per_mw_yr'] * (sale >= dsa_start) * yearfrac_30360(dsa_start, sale))
    G25 = G['jv_live'] * dsa_mw * E29

    # ---- monthly flows (MOIC blocks 3-8)
    fep_sh = G['fep_share'] if G['developer_funds_pro_rata'] == 1 else 1.0                        # E106
    cash_m = [0.0] * NM; rec_m = [0.0] * NM; guar_m = [0.0] * NM; dev_m = [0.0] * NM; land_m = [0.0] * NM
    recovered = 0.0
    for l in lines:
        if 0 <= l['month'] < NM:
            if l['cash'] == 1:
                cash_m[l['month']] += l['incl']
                rec_m[l['month']] += l['incl'] * l['share']
                if l['month'] <= m_sale:
                    recovered += l['incl'] * l['rec']                                                # G99
            else:
                guar_m[l['month']] += l['incl']
    for x in dl:
        if 0 <= x['month'] < NM:
            dev_m[x['month']] += x['H']; land_m[x['month']] += x['K'] * (G['devex_basis_switch'] == 1)
    unrec_m = [cash_m[m] - rec_m[m] for m in range(NM)]
    flag_acc = [1 if (d <= p_start or d > p_end) or not o['pause'] else 0 for d in mdates]            # row 10
    eff = lambda r: (1 + r) ** (1 / 12) - 1 if G['compounding_effective_switch'] == 1 else r / 12
    gbal = 0.0; gfee = [0.0] * NM; gint_acc = 0.0; gint_total = 0.0
    rb = 0.0; rint = 0.0; ab = 0.0; aint = 0.0; acum = 0.0; topup = 0.0; arepay = 0.0
    g_rate = G['guarantee_rate'] if G['guarantee_rate_flat_switch'] == 1 else None
    if g_rate is None:
        raise NotImplementedError('EURIBOR + margin guarantee pricing (Inputs E114 = 0) not shadowed')
    for m in range(NM):
        sale_f = 1 if m == m_sale else 0
        issued = guar_m[m]
        gfee[m] = (gbal + issued) * g_rate / 12                                                      # row 79
        gi = eff(G['r_guarantee_interest']) * (gbal + issued + gint_acc) * flag_acc[m]               # row 82
        gint_total += gi
        gint_acc = (gint_acc + gi) * (1 - sale_f)
        gbal = (gbal + issued) * (1 - sale_f)                                                        # rows 73-76
        draw = fep_sh * rec_m[m]                                                                     # rows 109-113
        i = eff(G['r_recoverable']) * (rb + draw) * flag_acc[m]
        rint += i
        rb = (rb + draw + i) * (1 - sale_f)
        adraw = fep_sh * (unrec_m[m] + dev_m[m] + land_m[m])                                         # rows 123-130
        acum += adraw
        ai = eff(G['r_at_risk']) * (ab + adraw) * flag_acc[m] * (1 - G['at_risk_flat_multiple_switch'])
        aint += ai
        if sale_f:
            topup = max(G['min_multiple_at_risk'] * acum - (ab + adraw + ai), 0)
            arepay = ab + adraw + ai
        ab = (ab + adraw + ai) * (1 - sale_f)
    G110 = sum(fep_sh * rec_m[m] for m in range(NM))
    G117 = sum((1 - fep_sh) * rec_m[m] for m in range(NM))
    G124 = sum(fep_sh * unrec_m[m] for m in range(NM))
    G125 = sum(fep_sh * (dev_m[m] + land_m[m]) for m in range(NM))
    if o['guarantee_fee_days']:                                                       # fix 8: [posting, release), act/365
        gfee = [0.0] * NM
        for l in lines:
            if l['cash'] == 0 and l['incl'] and l['date'] <= sale:
                for m in range(NM):
                    a, b = max(l['date'], eomonth(inv, m - 1) + dt.timedelta(days=1)), min(sale, mdates[m])
                    if b >= a:
                        gfee[m] += l['incl'] * g_rate * ((b - a).days + 1) / 365
    G79 = sum(gfee) if o['guarantee_fee'] else 0.0

    # ---- price effect and premium (Tender §4; MOIC G143)
    pe = (price_effect_quarterly(o, ph, ref, E18, E29, E20, cod, ntp, sale) if o['grid'] == 'quarterly'
          else price_effect(strategy, o, ph, ref, E18, E29, E20, E23, cod, ntp, sale))
    premium = o['rtb_premium'] if o['rtb_premium'] is not None else G['rtb_premium_keur_per_mw']
    gross_premium = premium * E150 + G['bess_value_keur_per_mw'] * E20                              # Tender E220
    G143 = gross_premium + pe['adjustment']

    # ---- waterfall (MOIC G143-G154)
    G144 = recovered + G20
    G145 = G143 + G144
    G146 = -G['sale_costs_keur']
    if o['interest_method'] == 'monthly':
        rec_int, atrisk_ret, guar_int = rint, aint + topup, gint_total
    else:
        rec_int, atrisk_ret, guar_int = Y23, G131 + Y24, Y22
    G147 = -((G110 + G117) + rec_int)
    G148 = -(G125 if G['at_risk_principal_repaid_switch'] == 1 else G20)
    G149 = -atrisk_ret
    G150 = -guar_int
    G151 = -G25
    G152 = G145 + G146 + G147 + G148 + G149 + G150 + G151
    dev_share = 1 - (G['fep_share'] if G['jv_live'] == 1 else 1)
    if o['waterfall'] == 'capped':                                                    # fix 6: MIN(due, cash left)
        left = max(0.0, G145)
        paid = []
        for due in (-G146, G110 + G117, rec_int, -G148, atrisk_ret, guar_int, -G151):
            p_ = min(due, left)
            paid.append(p_)
            left -= p_
        G146, G147, G148, G149, G150, G151 = (-paid[0], -(paid[1] + paid[2]), -paid[3], -paid[4], -paid[5], -paid[6])
        G152 = left
        G154 = G152 * dev_share
    else:
        G154 = max(G151, G152 * dev_share) if o['dev_floor'] else G152 * dev_share
    G153 = G152 - G154

    # ---- FEP monthly cash flows (MOIC block 10) and returns (block 11)
    out_m = [0.0] * NM
    for m in range(NM):
        out_m[m] -= fep_sh * rec_m[m] + fep_sh * unrec_m[m] + fep_sh * (dev_m[m] + land_m[m])
        out_m[m] -= fep_sh * gfee[m] if o['guarantee_fee'] else 0
    if o['signing_dd']:
        out_m[0] -= G['signing_fee_keur'] + G['dd_costs_keur']
    G164 = sum(out_m)
    infl_sale = (-G147 - G148 - G149 - G150) * fep_sh + G153                                         # rows 166-170
    pre = [out_m[m] + (infl_sale if m == m_sale else 0) for m in range(NM)]
    cum = 0.0; taxm = [0.0] * NM
    for m in range(NM):
        cum += pre[m]
        if m == m_sale and o['tax']:
            taxm[m] = -G['cit'] * max(cum - G['participation_exemption'] * G153, 0)                   # row 174
    post = [pre[m] + taxm[m] for m in range(NM)]
    G172 = sum(pre); G174 = sum(taxm); G182 = sum(post)
    G171 = infl_sale
    moic = (G171 + G174) / -G164 if G164 else float('nan')
    cf, cfd = post, mdates
    if o['grid'] == 'quarterly':                                     # month m (0 = Jul-2026) -> quarter m // 3, dated at its end
        qd = [e for _, e in quarters()][:(NM + 2) // 3]
        cf = [sum(post[m] for m in range(NM) if m // 3 == q) for q in range(len(qd))]
        cfd = qd
    irr = xirr(cf, cfd)
    dr = G['discount_rate']
    npv = sum(c / (1 + dr) ** ((d - cfd[0]).days / 365) for c, d in zip(cf, cfd))
    pv_eq = -sum(c / (1 + dr) ** ((d - cfd[0]).days / 365) for c, d in zip(cf, cfd) if c < 0)
    cpost, peak, econ, prev = 0.0, 0.0, 0.0, 0.0
    for m in range(NM):
        charge = 0.0 if m == 0 else min(prev, 0) * ((1 + G['coc_econ_profit']) ** (1 / 12) - 1)
        cpost += post[m]; prev = cpost
        peak = min(peak, cpost)
        econ += post[m] + charge

    # ---- developer (MOIC block 12)
    dev_in = -G151 * 1 + G154                                                                        # rows 212-213 (developer funds nothing)
    dev_tax = -G['cit'] * max(dev_in - G['participation_exemption'] * G154, 0) if o['tax'] else 0.0
    pts = tender_points(strategy, net, o['osz4_basis_installed'])

    k = 1000.0
    res = {
        'strategy': strategy, 'name': J['strategies'][strategy]['name'],
        'dates': dict(rtb=rtb.isoformat(), sale=sale.isoformat(), ntp=ntp.isoformat(), cod=cod.isoformat(), sale_month=m_sale, pb_date=pb_date.isoformat()),
        'installed_wind_mw': ph['gross'], 'bess_mw': ph['bess'], 'bess_mwh': ph['mwh'], 'net_mw_E18': E18, 'installed_E29': E29,
        'premium_basis_mw': E150, 'clip': ph['clip'],
        'points': pts,
        'gross_premium': gross_premium, 'price_effect': pe['adjustment'], 'price_effect_sum_pv': pe['sum_pv'],
        'price_effect_items': pe['items'], 'effective_price': G143,
        'effective_price_per_basis_mw': (gross_premium + pe['adjustment']) / E150 if E150 else 0,
        'wf': {'49 premium proceeds (effective)': G143, '50 buyer reimbursement': G144, '51 total proceeds': G145,
               '52 sale costs': G146, '53 recoverable loans - principal': -(G110 + G117),
               '54 recoverable loans - 12% interest': G147 + (G110 + G117),
               '55 unrecoverable loans - principal (land)': G148, '56 at-risk return MAX(20%,2x) + fee 12%': G149,
               '57 guarantee interest 6%': G150, '58 DSA': G151, '59 distributable': G152, '60 to FEP': G153, '61 to developer': G154},
        'fep_pre_tax_profit': G172, 'fep_tax_at_sale': G174, 'fep_post_tax_profit': G182, 'fep_irr': irr, 'fep_moic': moic,
        'fep_npv': npv, 'fep_npv_to_equity': npv / pv_eq if pv_eq else float('nan'), 'fep_peak_exposure': -peak,
        'economic_profit': econ, 'fep_cash_invested': -G164, 'fep_inflows_at_sale': G171,
        'developer_post_tax_profit': dev_in + dev_tax,
        'detail': {'securities': [{k2: (v.isoformat() if isinstance(v, dt.date) else v) for k2, v in l.items()} for l in lines],
                   'devex_fep_principal_G18': G18, 'land_G20': G20, 'devex_return_G131': G131, 'Y22': Y22, 'Y23': Y23, 'Y24': Y24,
                   'dsa_keur_per_mw': dsa_mw, 'guarantee_fee_G79': G79, 'monthly_rec_interest_G111': rint, 'monthly_guar_interest_G82': gint_total,
                   'monthly_atrisk_topup_G130': topup, 'G124_unrec_cash': G124, 'G125_devex_land': G125, 'contingency': cont,
                   'pb_gross_mhuf': pb_gross, 'bid_bond_mhuf': bid, 'd_pb_keur': pe['d_pb_keur'],
                   'post_tax_monthly': post, 'month_dates': [d.isoformat() for d in mdates],
                   'devex_monthly': dev_m, 'land_monthly': land_m, 'devex_lines': dl, 'fee_keur': fee_keur,
                   'cash_flows': cf, 'cash_flow_dates': [d.isoformat() for d in cfd]},
        'developer_total': -G151 + G154,
    }
    return res


# ----------------------------------------------------------------------------------------------- comparison
def bora_view(r):
    """Outputs in Bora-sheet units (€m, %, x, pts) keyed by Bora cell row, plus the Y22:AD29 item block."""
    k = 1000.0
    p = r['points']
    v = {22: r['installed_wind_mw'], 23: r['bess_mw'], 26: p['osz1'], 27: p['osz2'], 28: p['osz3'], 29: p['osz4'], 30: p['osz5'],
         31: p['osz6'], 32: p['osz7'], 33: p['osz8'], 34: p['osz9'], 35: p['total_eu'], 36: p['total_cn'],
         41: r['gross_premium'] / k, 42: r['price_effect'] / k, 43: r['effective_price'] / k, 44: r['fep_post_tax_profit'] / k,
         45: r['fep_irr'], 46: r['fep_moic']}
    for row, key in zip(range(49, 62), r['wf']):
        v[row] = r['wf'][key] / k
    v[51] = r['wf']['51 total proceeds'] / k
    it = r['price_effect_items']
    y = {23: it['municipal'], 24: it['community'], 25: it['bess'], 26: it['security'], 27: it['eu_turbine'], 28: it['overbuild']}
    y = {rr: vv / k for rr, vv in y.items()}
    y[29] = sum(y.values())
    return v, y


LABEL = {22: 'Installed wind MW', 23: 'BESS MW', 26: 'pts OSZ-1 municipal', 27: 'pts OSZ-2 community', 28: 'pts OSZ-3 overbuild',
         29: 'pts OSZ-4 BESS', 30: 'pts OSZ-5 partial', 31: 'pts OSZ-6 security', 32: 'pts OSZ-7 brownfield', 33: 'pts OSZ-8 shared',
         34: 'pts OSZ-9 EU equipment', 35: 'Total score EU', 36: 'Total score CN', 41: 'Gross RtB premium €m', 42: 'Price effect €m',
         43: 'Effective RtB price €m', 44: 'FEP net profit post-tax €m', 45: 'FEP IRR', 46: 'FEP MOIC', 49: 'WF premium proceeds',
         50: 'WF buyer reimbursement', 51: 'WF total proceeds', 52: 'WF sale costs', 53: 'WF recoverable principal', 54: 'WF recoverable 12% int',
         55: 'WF land principal', 56: 'WF at-risk return', 57: 'WF guarantee int 6%', 58: 'WF DSA', 59: 'WF distributable', 60: 'WF to FEP',
         61: 'WF to developer'}
LO_MAP = {41: ('Calc!BQ305', 1e-3), 42: ('Calc!BS305', 1e-3), 43: ('Calc!BT305', 1e-3), 44: ('Calc!H305', 1e-3), 45: ('Calc!I305', 1),
          46: ('Calc!L305', 1), 50: ('Calc!CF305', 1e-3), 52: ('Calc!CG305', 1e-3), 53: ('Calc!CH305', 1e-3), 54: ('Calc!CI305', 1e-3),
          55: ('Calc!CJ305', 1e-3), 56: ('Calc!CK305', 1e-3), 57: ('Calc!CL305', 1e-3), 58: ('Calc!CM305', 1e-3), 59: ('Calc!CN305', 1e-3),
          60: ('Calc!CO305', 1e-3), 61: ('Calc!CP305', 1e-3), 49: ('Calc!BT305', 1e-3), 22: ('Calc!CD305', 1), 23: ('Calc!CE305', 1)}
LO_Y = {23: 'Calc!DC305', 24: 'Calc!DD305', 25: 'Calc!DE305', 26: 'Calc!DF305', 27: 'Calc!DG305', 28: 'Calc!DH305'}
EXTRA = [('developer_post_tax_profit', None, 'MOIC!G220', 'Developer post-tax profit k€'),
         ('fep_tax_at_sale', None, 'MOIC!G174', 'FEP tax at sale k€'), ('fep_cash_invested', None, 'MOIC!G164', 'FEP cash invested k€ (LO = -G164)'),
         ('fep_pre_tax_profit', 'G', 'Calc!G305', 'FEP pre-tax profit k€'), ('fep_npv', 'J', 'Calc!J305', 'FEP NPV @8% k€'),
         ('fep_npv_to_equity', 'K', 'Calc!K305', 'FEP NPV/Equity'), ('fep_peak_exposure', 'M', 'Calc!M305', 'Peak exposure k€'),
         ('economic_profit', 'N', 'Calc!N305', 'Economic profit k€'), ('fep_inflows_at_sale', 'Q', 'Calc!Q305', 'FEP inflows at sale k€'),
         ('price_effect_sum_pv', 'BR', 'Calc!BR305', 'Σ PV promises k€')]


def rel(a, b):
    if a is None or b is None or isinstance(a, str) or isinstance(b, str):
        return None
    return abs(a - b) / max(abs(b), 1e-9) if abs(b) > 1e-9 else abs(a - b)


def compare(strategies=('S1', 'S2', 'S3', 'S0', 'S4')):
    lo_path = HERE / 'lo_runs_rtb.json'
    lo = json.loads(lo_path.read_text())['runs'] if lo_path.exists() else {}
    worst = {}
    print(f"{'item':34s} {'strategy':8s} {'shadow':>16s} {'saved(Excel)':>16s} {'Δrel saved':>11s} {'LO live':>16s} {'Δrel LO':>10s}")
    for s in strategies:
        r = compute(s)
        v, y = bora_view(r)
        sv = J['saved_bora'][s]
        run = str(CODE[s] * 5 + 1)
        L = lo.get(run, {})
        col, yc = BORA_COL[s], BORA_YCOL[s]
        rows = [(f'{LABEL[k]}', v[k], sv.get(f'Bora!{col}{k}'), (L.get(LO_MAP[k][0]) * LO_MAP[k][1]) if k in LO_MAP and L.get(LO_MAP[k][0]) is not None else None, k == 45)
                for k in sorted(v)]
        rows += [(f'PE item Y{k} {["", "", "", "municipal", "community", "BESS", "security", "EU turbine", "overbuild", "total"][k - 20]}', y[k],
                  sv.get(f'Bora!{yc}{k}'), (L.get(LO_Y[k]) / 1000 if k in LO_Y and L.get(LO_Y[k]) is not None else None), False) for k in sorted(y)]
        dts = J['saved_data_table'][s]
        rr = {'S1': 306, 'S2': 311, 'S3': 316, 'S4': 321, 'S0': 326}[s]
        rows += [(lab, r[key], dts.get(f'Calc!{c}{rr}') if c else None,
                  (-L[lc] if key == 'fep_cash_invested' else L[lc]) if lc in L else None, False) for key, c, lc, lab in EXTRA]
        tcol = {'S1': 'E', 'S2': 'F', 'S3': 'G', 'S4': 'H', 'S0': 'I'}[s]                     # Tender points block (live)
        for k, tr in zip(range(26, 37), (55, 60, 65, 70, 72, 76, 78, 80, 82, 84, 85)):
            rows = [(lab, a, b, (L.get(f'Tender!{tcol}{tr}') if lab == LABEL[k] else c), i) for lab, a, b, c, i in rows]
        for lab, a, b, c, is_irr in rows:
            d1, d2 = rel(a, b), rel(a, c)
            if is_irr:
                d1 = abs(a - b) if b is not None else None
                d2 = abs(a - c) if c is not None else None
            f = lambda x: '' if x is None else (f'{x:16.8f}' if abs(x) < 10 else f'{x:16.6f}')
            g = lambda x: '' if x is None else f'{x:10.1e}'
            flag = ''
            tol = 1e-8 if is_irr else 1e-6
            if d2 is not None and d2 > tol:
                flag = '  <-- LO'
            if d1 is not None and d1 > tol and not is_irr:
                flag += '  <-- saved'
            if is_irr and b is not None and abs(b) < 1e-6 and abs(a) > 1e-4:
                flag += '  <-- saved = Excel XIRR failure (2.98e-09)'
            print(f'{lab:34s} {s:8s} {f(a)} {f(b)} {g(d1):>11s} {f(c)} {g(d2):>10s}{flag}')
            if d2 is not None:
                worst[(s, 'LO')] = max(worst.get((s, 'LO'), 0), d2)
            if d1 is not None and not (is_irr and b is not None and abs(b) < 1e-6):
                worst[(s, 'saved')] = max(worst.get((s, 'saved'), 0), d1)
        print()
    print('Worst difference per strategy (relative; IRR absolute):')
    for (s, w), x in sorted(worst.items()):
        print(f'  {s} vs {w:5s}: {x:.2e}')
    print('Note: saved Excel IRR for S2 (BORA 1) and S4 is 2.98e-09 = an Excel XIRR failure; LibreOffice and this shadow give the true root.')


if __name__ == '__main__':
    compare()
