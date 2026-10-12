#!/usr/bin/env python3
"""Pure-Python shadow of the OLD Bora engine – build / sale-at-COD / hold-to-maturity side.

Source: reference/Bora_scenarios_meeting.xlsx, sheets HTM (annual engine, blocks 1–12 and 17), Deposits (securities
and connection-fee schedule), Tender (strategy → MW, grid-bottleneck loss, COD price effect) and Inputs.
Logic map with cell references: htm_logic.md (same folder).

Stdlib only. Inputs: old_inputs_htm.json (Inputs E-column values, case values, strategies, RtB-engine interface),
curves.csv (Aurora wind capture, nominal). Comparison data: lo_runs_htm.json (LibreOffice live values + saved values).

    compute(strategy, route, **params) -> dict of outputs (k€, %, x)
        strategy: 'S1' (current bid), 'S2' (BORA 1), 'S3' (BORA 2)
        route:    'HTM-EU', 'HTM-CN', 'COD-EU', 'COD-CN'
    python3 old_htm.py              -> comparison table vs saved workbook values and LibreOffice live values
    python3 old_htm.py --toggles    -> materiality of the mechanics (HTM IRR / COD IRR deltas)

Sign convention as in the workbook: + = inflow to the party, − = outflow. Units k€ unless stated.
"""
import csv
import datetime as dt
import json
import math
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
YEARS = list(range(2026, 2062))            # HTM!H5:AQ5
N = len(YEARS)
EXCEL0 = dt.date(1899, 12, 30)


# ----------------------------------------------------------------------------------------------- helpers
def ser(d):
    """Excel serial number of a date."""
    return (d - EXCEL0).days


def todate(x):
    return dt.date.fromisoformat(x) if isinstance(x, str) else x


def eomonth(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    y, m = d.year + y, m + 1
    nxt = dt.date(y + (m == 12), m % 12 + 1, 1)
    return nxt - dt.timedelta(days=1)


def edate(d, n):
    y, m = divmod(d.month - 1 + n, 12)
    y, m = d.year + y, m + 1
    last = eomonth(dt.date(y, m, 1), 0).day
    return dt.date(y, m, min(d.day, last))


def yearfrac30360(a, b):
    """Excel YEARFRAC basis 0 (US NASD 30/360)."""
    d1, d2 = a.day, b.day
    feb_end = lambda d: d.month == 2 and (d + dt.timedelta(days=1)).month == 3
    if feb_end(a) and feb_end(b):
        d2 = 30
    if feb_end(a):
        d1 = 30
    if d2 == 31 and d1 >= 30:
        d2 = 30
    if d1 == 31:
        d1 = 30
    return ((b.year - a.year) * 360 + (b.month - a.month) * 30 + (d2 - d1)) / 360


def irr(flows, lo=-0.99, hi=10.0):
    """Annual IRR (Excel IRR on equally spaced flows, t0 = first flow), solved to machine precision."""
    f = lambda r: sum(v / (1 + r) ** t for t, v in enumerate(flows))
    return _root(f, lo, hi)


def xirr(flows, dates, lo=-0.99, hi=10.0):
    """XIRR: Σ v_i / (1 + r)^((d_i − d_0) / 365), d_0 = first date of the array (as Excel / LibreOffice)."""
    d0 = dates[0]
    f = lambda r: sum(v / (1 + r) ** ((d - d0) / 365.0) for v, d in zip(flows, dates))
    return _root(f, lo, hi)


def _root(f, lo, hi):
    flo, fhi = f(lo), f(hi)
    if flo * fhi > 0:
        return float('nan')
    for _ in range(300):
        mid = (lo + hi) / 2
        fm = f(mid)
        if fm == 0 or hi - lo < 1e-16:
            return mid
        if (fm > 0) == (flo > 0):
            lo, flo = mid, fm
        else:
            hi = mid
    return (lo + hi) / 2


def days_between(y, start, end):
    """MAX(0, MIN(DATE(y,12,31), end) − MAX(DATE(y,1,1) − 1, start)) in days (Excel day arithmetic)."""
    a = max(dt.date(y, 1, 1) - dt.timedelta(days=1), start)
    b = min(dt.date(y, 12, 31), end)
    return max(0, (b - a).days)


# ----------------------------------------------------------------------------------------------- data
def load_inputs():
    with open(os.path.join(HERE, 'old_inputs_htm.json'), encoding='utf-8') as f:
        data = json.load(f)
    curve = {}
    with open(os.path.join(HERE, 'curves.csv'), encoding='utf-8') as f:
        for r in csv.DictReader(f):
            if r['year'] and r['year'][0].isdigit():
                curve[int(r['year'])] = float(r['wind_capture_selected_nominal_used_by_HTM'])
    return data, curve


DATA, CURVE = load_inputs()
ROUTES = {'HTM-EU': ('2', 'EU', 'HTM'), 'HTM-CN': ('3', 'CN', 'HTM'), 'COD-EU': ('7', 'EU', 'COD'), 'COD-CN': ('8', 'CN', 'COD')}


def _E(row):
    return DATA['inputs'][f'Inputs!E{row}']['value']


# ----------------------------------------------------------------------------------------------- engine
def compute(strategy='S2', route='HTM-EU', *, fx=None, built_overbuild=True, bess_built=True, deemed_exit=None,
            deemed_exit_basis='htm_case', fep_share=None, dsra=True, sweep=True, pb_hold_months=0, bid_netting=True,
            direct_fee_basis='per_mw', direct_fee_share=1.0, pb_forfeit=False, overrides=None, detail=False,
            bid_release=False, pb_release_date=None, devex=None, guarantee_fee_days=False, claims='line',
            buyout_keur=0.0, ppa_from_cod=False, cod_waterfall='full'):
    """Recompute one strategy × route of the old engine. Defaults reproduce the workbook.

    fx               HUF/EUR for the Deposits schedule (Inputs E11 = 365). Devex and the deemed-exit interface stay
                     at their LibreOffice values (they come from the RtB engine).
    built_overbuild  False: build only the net MVA (no overbuild) – physical MW for capex, generation, opex, decom,
                     bottleneck loss, COD price and DSA; securities stay on the committed MW.
    bess_built       False: no BESS capex (securities still on the committed MWh).
    deemed_exit      None = old (HTM on, COD off: Inputs E228); True / False forces rows 21–22.
    deemed_exit_basis 'htm_case' (old: MOIC re-run under case 2/3 inputs) or 'rtb_route' (the RtB route result).
    fep_share        None = old rule HTM E128 (100 % with the deemed exit, else 70 %); a number forces it.
    dsra, sweep      False switches off the DSRA (Inputs E268 = 0) / cash sweep (E269 = 0).
    pb_hold_months   HTM only: guarantees outstanding at the RtB-sale date run this many months past COD (old: 0).
    bid_netting      False: the bid bond is not credited into the performance bond (Deposits I10 = 0).
    direct_fee_basis 'per_mw' (old: Inputs E84 × net MW) or 'lump' (Annex 1: 508.345804 mHUF × direct_fee_share).
    pb_forfeit       True: proportional performance-security call at COD when built < committed by > 10 % (wind
                     and BESS tested separately; tender rules §2.3) – not in the old engine.
    overrides        {'E214': 1.0, ...} replaces Inputs E-values (row numbers of the Inputs sheet).

    Toggles added for the new-model reconciliation (SPEC v2 §8; defaults = old engine):
    bid_release      True: the bid guarantee is released when the PB is posted (fix 3; use with bid_netting=False).
    pb_release_date  date the guarantees outstanding after the RtB-sale date end on (HTM; fix 4: 31-Mar-2033);
                     overrides pb_hold_months.
    devex            {'dv': [...], 'land': [...], 'g280': float}: monthly devex / land vectors (MOIC rows 18 / 20) and the
                     at-risk claim to COD (MOIC G280) recomputed by old_rtb.py under the same fixes (fix 5, 7, 9).
    guarantee_fee_days True: fee on exact days outstanding [posting, release), act/365, by calendar year (fix 8).
    claims           'line' (old Deposits AF) | 'daily': COD claims on the new per-class accrual (fix 9).
    buyout_keur      HTM only: developer buy-out (DSA + 30 %) paid by FEP in the RtB-sale year (fix 10).
    ppa_from_cod     True: PPA for its tenor from COD (COD-year stub on PPA, 2041 part-year), then merchant (fix 10).
    cod_waterfall    'full' (old: developer floored at -DSA) | 'capped' (MIN steps, no floor; fix 6).
    """
    case, turb, kind = ROUTES[route]
    S = DATA['strategies'][strategy]
    I = {int(k.split('E')[1]): v['value'] for k, v in DATA['inputs'].items()}
    for k, v in DATA['case_values'][case].items():
        I[int(k.split('E')[1])] = v
    I[233], I[234], I[90] = S['municipal_keur_per_mw'], S['community_eur_per_mwh'], S['pb_uplift']
    I[21] = S['bess_hours']
    if fx is not None:
        I[11] = fx
    if not dsra:
        I[268] = 0.0
    if not sweep:
        I[269] = 0.0
    if deemed_exit is not None:
        I[228] = 1.0 if deemed_exit else 0.0
    for k, v in (overrides or {}).items():
        I[int(str(k).lstrip('E'))] = v
    T = {k: v['value'] for k, v in DATA['tender_and_calc'].items()}
    cod_sale = kind == 'COD'

    # --- capacities (Tender rows 37–45, Inputs rows 16–29) ---------------------------------------------------
    mva = T['Calc!E74']                                                    # Tender E37 = Calc E74
    E18 = I[16] * I[17]                                                    # successful net MW
    mw_commit = round(mva * (1 + S['overbuild']), 6)                       # Tender row 40 → Inputs E297
    E29 = E18 * mw_commit / mva                                            # Inputs E29 (E281 = net MVA)
    bess_c = round(S['bess_ratio'] * (mw_commit if S['bess_basis_installed'] else mva), 6)   # Tender row 41
    E20 = bess_c
    E22 = E20 * I[21]
    E23m = E22 * I[17]
    mw = E29 if built_overbuild else E18                                   # physical wind MW (HTM E35)
    bess = E20 if bess_built else 0.0
    ob = max(0.0, mw / mva - 1)                                            # Tender row 43 (on the built MW)
    clip = ob / (1 + ob) * T['Tender!E27']                                 # Tender row 45: grid-bottleneck loss

    # --- timeline (Inputs rows 27, 55–58, 152–158) -----------------------------------------------------------
    inv = todate(I[27])
    cod_in = todate(I[298])
    E55 = (cod_in.year - inv.year) * 12 + cod_in.month - inv.month - I[152] - I[153]
    rtb = eomonth(inv, int(E55))
    E57 = int(E55 + I[28])
    sale = eomonth(inv, E57)                                               # RtB sale date (Inputs E58)
    ntp = eomonth(rtb, int(I[152]))
    cod = eomonth(ntp, int(I[153]))
    ycod, yfirst, ylast = cod.year, cod.year + 1, cod.year + int(I[157])
    assert ylast <= YEARS[-1]

    # --- Deposits schedule (Deposits rows 7–18) ----------------------------------------------------------------
    FX = I[11]
    E91 = I[16] * (E29 / E18) * I[75] + E22 * I[89]                        # bid-stage base (committed MW)
    E92 = E29 * I[75] + E23m * I[89]                                       # post-award base
    lines = []

    def line(key, H, I_adj, K, P, U, R):
        J = H + I_adj
        L = (K.year - inv.year) * 12 + K.month - inv.month
        M = J * 1000 / FX
        Nf = 1 if L > E57 else 0
        O = M * (I[109] if Nf else 1)
        lines.append(dict(key=key, H=H, J=J, K=K, L=L, M=M, O=O, P=P, U=U, R=R))

    award = todate(I[96])
    line('fee', I[81] * I[82], 0, todate(I[95]), I[117], I[256], 'Fee')
    H8 = I[76] * E91
    line('bid', H8, 0, todate(I[95]), I[118], I[257], 'Security')
    line('bid_refund', -I[76] * ((I[16] - E18) * I[75] * (E29 / E18)) * I[110], 0, award, I[118], I[257], 'Refund')
    H10 = I[77] * I[90] * E92
    line('pb', H10, -H8 * I[107] * (1 if bid_netting else 0), award + dt.timedelta(days=int(I[102])), I[119], I[258], 'Security')
    if bid_release:
        line('bid_release', -H8, 0, award + dt.timedelta(days=int(I[102])), I[118], I[257], 'Refund')
    H11 = I[83] * E18
    line('tso', H11, 0, todate(I[97]), I[120], I[259], 'Security')
    H12 = I[84] * E18 if direct_fee_basis == 'per_mw' else 508.345804 * direct_fee_share
    line('direct', H12, 0, todate(I[99]), I[121], I[260], 'Connection fee')
    assert I[78] == 0, 'supplementary security (Deposits rows 13–17) is 0 in this procedure – not shadowed'
    line('indirect', I[85] * E18, -H11, todate(I[99]), I[126], I[265], 'Connection fee')
    tso = next(l for l in lines if l['key'] == 'tso')

    # --- guarantee fees (MOIC rows 72–79 → HTM row 23) ---------------------------------------------------------
    rate = I[113]
    assert I[114] == 1, 'EURIBOR-based guarantee rate not shadowed'
    gfee_month = {}
    bal = 0.0
    for m in range(0, 61):
        issued = sum(l['O'] * (1 - l['P']) for l in lines if l['L'] == m)
        fee = (bal + issued) * rate / 12
        y = eomonth(inv, m).year
        gfee_month[y] = gfee_month.get(y, 0.0) + fee
        bal = 0.0 if m == E57 else bal + issued
    G_sale = sum(l['O'] * (1 - l['P']) for l in lines if l['L'] <= E57)    # −MOIC!G75
    g_end = edate(cod, int(pb_hold_months)) if (kind == 'HTM' and pb_hold_months) else cod
    if kind == 'HTM' and pb_release_date:
        g_end = todate(pb_release_date)

    # --- annual grid ---------------------------------------------------------------------------------------------
    E15 = I[226]
    Z = lambda: [0.0] * N
    yr = YEARS
    t = [y - inv.year for y in yr]                                         # HTM row 4
    span = (cod - ntp).days
    r7 = [days_between(y, ntp, cod) / span for y in yr]                    # construction share
    r14 = [days_between(y, ntp, cod) / ((dt.date(y, 12, 31) - dt.date(y, 1, 1)).days + 1) for y in yr]
    r10 = [((12 - cod.month) / 12 if y == ycod else 0) + (1 if yfirst <= y < ylast else 0) + (cod.month / 12 if y == ylast else 0)
           for y in yr]
    r8 = [1 if x >= 1 else 0 for x in r10]
    r9 = [y - yfirst + 1 if (r10[i] > 0 and y >= yfirst) else 0 for i, y in enumerate(yr)]
    r11 = [1 if y == ylast else 0 for y in yr]

    # 1. development phase (rows 16–24) – devex interface from MOIC rows 18 + 20
    itf = DATA['interface']
    mdates = [todate(d) for d in itf['MOIC!H5:BP5 (month-end dates)']]
    dv = itf['MOIC!H18:BP18 devex (k€)'] if devex is None else devex['dv']
    land = itf['MOIC!H20:BP20 land paid by FEP (k€)'] if devex is None else devex['land']
    r16 = [-sum(dv[k] + land[k] for k in range(len(mdates)) if mdates[k].year == y) * E15 for y in yr]
    r17 = [-(I[24] + I[25]) * E15 if y == inv.year else 0.0 for y in yr]
    r18 = [-sum(l['P'] * l['M'] for l in lines if l['K'].year == y and l['R'] != 'Connection fee') * E15 for y in yr]
    r19 = [(sum(l['P'] * l['M'] for l in lines if l['R'] in ('Security', 'Refund')) - tso['P'] * tso['M']) * E15
           if y == ycod else 0.0 for y in yr]
    r20 = Z()                                                              # Inputs E227 = 0 (superseded)
    assert I[227] == 0
    de = itf['deemed_exit'].get(strategy)
    E21 = I[228] * T['Tender!J48']
    if E21 and de is None:
        raise ValueError(f'no deemed-exit interface values for {strategy}')
    dsrc = (de or {}).get('as_used_in_HTM (MOIC under case 2/3)' if deemed_exit_basis == 'htm_case'
                          else 'RtB_route (MOIC under case 1)', {'G182': 0, 'G154': 0, 'G151': 0})
    r21 = [-E21 * dsrc['G182'] if y == sale.year else 0.0 for y in yr]
    r22 = [-E21 * (dsrc['G154'] - dsrc['G151']) if y == sale.year else 0.0 for y in yr]
    r23 = [-(gfee_month.get(y, 0.0) + rate * G_sale * days_between(y, sale, g_end) / 365) * E15 for y in yr]
    if guarantee_fee_days:                                           # [posting, release), release = day after g_end
        r23 = [-sum(l['O'] * rate * days_between(y, l['K'] - dt.timedelta(days=1), g_end) / 365
                    for l in lines if l['P'] == 0 and l['K'] <= g_end) * E15 for y in yr]
    if kind == 'HTM' and buyout_keur:
        r22 = [r22[i] - (buyout_keur if y == sale.year else 0.0) for i, y in enumerate(yr)]
    r24 = [sum(v) for v in zip(r16, r17, r18, r19, r20, r21, r22, r23)]

    # 2. capex (rows 27–30)
    E173 = ((I[162] + I[163] + I[164] + I[165] + I[166] + I[169] + I[170]) * (1 + I[171]) * mw / E18
            + I[167] * (1 + I[171]) / E18 + I[168] * bess / E18)
    E27c = E173 * E18
    r27 = [-E27c * x for x in r7]
    r28 = [-sum(l['M'] for l in lines if l['K'].year == y and l['R'] == 'Connection fee') for y in yr]
    r29 = [-I[233] * E18 if y == ycod else 0.0 for y in yr]
    r30 = [a + b + c for a, b, c in zip(r27, r28, r29)]
    G30 = sum(r30)

    # PB forfeiture (option, not in the old engine)
    forfeit = 0.0
    if pb_forfeit:
        pb_w = I[77] * I[90] * E29 * I[75] * 1000 / FX
        pb_b = I[77] * I[90] * E23m * I[89] * 1000 / FX
        s_w = max(0.0, 1 - mw / E29) if E29 else 0.0
        s_b = max(0.0, 1 - bess / E20) if E20 else 0.0
        forfeit = (s_w * pb_w if s_w > 0.1 else 0.0) + (s_b * pb_b if s_b > 0.1 else 0.0)
        r24 = [v - (forfeit if y == ycod else 0.0) for v, y in zip(r24, yr)]

    # 3. generation (rows 35–39)
    E36 = I[179] * (1 - I[180]) * (1 - I[181]) * I[183] * (1 - clip)
    r37 = [(1 - I[182]) ** max(r9[i] - 1, 0) if r10[i] > 0 else 0.0 for i in range(N)]
    r38 = [(mw if r10[i] > 0 else 0) * E36 * r37[i] * r10[i] / 1000 for i in range(N)]   # GWh

    # 4. prices (rows 42–49)
    last_y = max(CURVE)
    r42 = []
    for y in yr:
        if y in CURVE:
            r42.append(CURVE[y] * I[184])
        elif y > last_y:
            r42.append(CURVE[last_y] * (1 + 0.02) ** (y - last_y) * I[184])   # Curves E19 = 2 %
        else:
            r42.append(None)
    r43 = [I[187] * (1 + I[188]) ** (y - ycod) for y in yr]
    r44 = [x * I[189] for x in r43]
    r45 = [r42[i] if (I[186] == 1 and r42[i] is not None) else r44[i] for i in range(N)]
    r47 = [I[191] * (1 + I[192]) ** (y - ycod) for y in yr]
    r48 = [(1 - I[190]) if 1 <= r9[i] <= I[193] else 0.0 for i in range(N)]
    if ppa_from_cod:
        ppa_end = edate(cod, int(I[193]) * 12)
        r48 = [(sum(1 for mth in range(1, 13) if cod < dt.date(y, mth, 1) < ppa_end) / (r10[i] * 12)) if r10[i] > 0 else 0.0
               for i, y in enumerate(yr)]

    # 5. revenue, opex, EBITDA (rows 52–70)
    esc = I[209]
    r52 = [r38[i] * r45[i] * (1 - r48[i]) for i in range(N)]
    r53 = [r38[i] * r47[i] * r48[i] for i in range(N)]
    r54 = [a + b for a, b in zip(r52, r53)]
    r55 = [(1 + esc) ** (y - ycod) if r10[i] > 0 else 0.0 for i, y in enumerate(yr)]
    r56 = [-r38[i] * I[194] * r55[i] for i in range(N)]
    r57 = [-r38[i] * I[195] * r55[i] for i in range(N)]
    r58 = [-r54[i] * I[196] for i in range(N)]
    r59 = [-r38[i] * I[234] for i in range(N)]
    perw = lambda rate_: [-rate_ * mw * r55[i] * r10[i] for i in range(N)]
    r60, r61 = perw(I[201]), perw(I[202])
    r62 = [-I[203] * mw * (1 + esc) ** (y - ycod) * (r10[i] + r14[i]) for i, y in enumerate(yr)]
    r63, r64, r65 = perw(I[204]), perw(I[205]), perw(I[206])
    r66 = [-I[207] * r55[i] * r10[i] for i in range(N)]
    r67 = [(r56[i] + r57[i] + r58[i] + r60[i] + r61[i] + r62[i] + r63[i] + r64[i] + r65[i] + r66[i]) * I[208] for i in range(N)]
    r68 = [-r54[i] * I[210] for i in range(N)]
    r69 = [r56[i] + r57[i] + r58[i] + r59[i] + r60[i] + r61[i] + r62[i] + r63[i] + r64[i] + r65[i] + r66[i] + r67[i] + r68[i]
           for i in range(N)]
    r70 = [a + b for a, b in zip(r54, r69)]

    # 6. senior debt (rows 75–89) and 7. tax (rows 90–101) – sequential (sweep needs this year's tax)
    rd = I[218]
    E75 = I[217] * -G30
    r77 = [E75 * x for x in r7]
    tenor = int(I[220])
    r76, r78, r79, r80, r81, r83, r84, r85, r86, r87, r88, r89 = (Z() for _ in range(12))
    r90 = Z()
    sum_opfrac_idx0 = sum(r10[i] for i in range(N) if r9[i] == 0)
    E92d = I[213]
    for i in range(N):
        r90[i] = 0.0 if r10[i] == 0 else (r10[i] if r9[i] == 0 else (1.0 if r9[i] < E92d else
                                          (1 - sum_opfrac_idx0 if r9[i] == E92d else 0.0)))
    # IDC and drawdowns do not depend on the sweep → G77 + G79 known before the annuity
    prev_close = 0.0
    idc = Z()
    for i in range(N):                                                    # first pass: construction only
        open_ = prev_close
        idc[i] = (open_ + r77[i] / 2) * rd * (r14[i] / (r14[i] + r10[i]) if r10[i] > 0 else 1) if r7[i] > 0 else 0.0
        prev_close = open_ + r77[i] + idc[i] * I[221]
        if r10[i] > 0 and r9[i] >= 1:
            break
    G77 = sum(r77)
    G79 = sum(idc) * I[221]
    E82 = (G77 + G79) * rd / (1 - (1 + rd) ** -tenor) if (G77 + G79) else 0.0
    G92base = (-G30 + G79 + tso['P'] * tso['M'] * E15) / E92d
    r92 = [G92base * x for x in r90]
    r93, r94, r95, r96, r97, r98, r99, r100, r101 = (Z() for _ in range(9))
    loss_open = 0.0
    prev_close = 0.0
    prev_dsra = 0.0
    for i in range(N):
        r76[i] = prev_close
        r78[i] = (r76[i] + r77[i] / 2) * rd * (r14[i] / (r14[i] + r10[i]) if r10[i] > 0 else 1) if r7[i] > 0 else 0.0
        r79[i] = r78[i] * I[221]
        r80[i] = r78[i] * (1 - I[221])
        r81[i] = ((r76[i] + r77[i] + r79[i]) * rd * r10[i] if r9[i] == 0 else r76[i] * rd) if r10[i] > 0 else 0.0
        r83[i] = min(max(E82 - r81[i], 0), r76[i]) if 1 <= r9[i] <= tenor else 0.0
        r85[i] = -(E75 * I[219] if (r77[i] > 0 and r76[i] == 0) else 0.0) - max(0.0, E75 - r76[i] - r77[i] / 2) * I[222] * r14[i]
        r86[i] = -(r81[i] + r83[i])
        r88[i] = (I[268] / 12 * min(E82, (r76[i] + r77[i] + r79[i] - r83[i]) * (1 + rd))
                  if (r10[i] > 0 and r9[i] + 1 <= tenor) else 0.0)
        r89[i] = -(r88[i] - prev_dsra)
        prev_dsra = r88[i]
        # tax
        r93[i] = r70[i] - r92[i] - r81[i] + r85[i] + r16[i]
        r94[i] = loss_open
        r95[i] = min(r94[i], max(r93[i], 0) * I[214])
        r96[i] = r94[i] + max(-r93[i], 0) - r95[i]
        loss_open = r96[i]
        r97[i] = max(r93[i], 0) - r95[i]
        r98[i] = -r97[i] * I[12]
        E99 = I[197] * I[198]
        r99[i] = -max(r93[i], 0) * E99
        r100[i] = r98[i] + r99[i]
        r101[i] = -max(r70[i] - r92[i], 0) * (I[12] + E99)
        # sweep
        r87[i] = (min(max(r70[i] + r100[i] + r86[i] + r89[i], 0) * I[269], r76[i] + r77[i] + r79[i] - r83[i])
                  if 1 <= r9[i] <= I[270] else 0.0)
        r84[i] = r76[i] + r77[i] + r79[i] - r83[i] - r87[i]
        prev_close = r84[i]

    # 8.–10. cash flows (rows 104–131)
    r107 = [-I[175] * mw * r11[i] for i in range(N)]
    r108 = [I[176] * mw * r11[i] for i in range(N)]
    r109 = [r24[i] + r30[i] + r70[i] + r107[i] + r108[i] for i in range(N)]
    r110 = [r109[i] + r101[i] for i in range(N)]
    r117 = [r85[i] - r80[i] for i in range(N)]
    r119 = [r86[i] - r87[i] for i in range(N)]
    r123 = [r24[i] + r30[i] + r77[i] + r117[i] + r70[i] + r119[i] + r89[i] + r100[i] + r107[i] + r108[i] for i in range(N)]
    E128 = (1.0 if I[228] == 1 else I[33]) if fep_share is None else fep_share
    assert T['Tender!E370'] == 0, 'community purchase right (HTM rows 383–387) is off in the workbook – not shadowed'
    r129 = [x * E128 for x in r123]

    # 11. results (column G)
    npv = lambda fl, r: sum(v / (1 + r) ** t[i] for i, v in enumerate(fl))
    cum = []
    c = 0.0
    for v in r129:
        c += v
        cum.append(c)
    cum123 = []
    c = 0.0
    for v in r123:
        c += v
        cum123.append(c)
    pos, neg = sum(v for v in r129 if v > 0), -sum(v for v in r129 if v < 0)
    pv_inv = -sum(v / (1 + I[253]) ** t[i] for i, v in enumerate(r129) if v < 0)
    dscr = [(r70[i] + r100[i]) / -r86[i] for i in range(N) if r9[i] >= 1 and r86[i] < -0.01]
    df12 = [1 / (1 + I[224]) ** x for x in t]
    lcoe = -(sum(a * b for a, b in zip(r24, df12)) + sum(a * b for a, b in zip(r30, df12)) + sum(a * b for a, b in zip(r69, df12))
             + sum(a * b for a, b in zip(r107, df12))) / sum(a * b for a, b in zip(r38, df12))
    htm = {
        'FEP IRR (HTM!G139, IRR on row 129)': irr(r129),
        'FEP NPV (HTM!G140)': npv(r129, I[253]),
        'FEP NPV/Equity (HTM!G142)': npv(r129, I[253]) / pv_inv,
        'FEP MOIC (HTM!G147)': pos / neg,
        'FEP net profit post-tax (HTM!G144)': sum(r129),
        'FEP net profit pre-tax (Run_Map!D22)': sum(r129) - sum(r100) * E128,
        'FEP peak exposure (HTM!G146)': -min(cum),
        'Payback year (HTM!G148)': next((yr[i] for i in range(N) if cum123[i] > 0), None),
        'Project IRR pre-tax (HTM!G134)': irr(r109),
        'Project IRR post-tax (HTM!G135)': irr(r110),
        'Project NPV (HTM!G136)': npv(r110, I[224]),
        'Equity IRR 100% (HTM!G137)': irr(r123),
        'Min DSCR (HTM!G126)': min(dscr) if dscr else None,
        'LCOE (HTM!G156)': lcoe,
        'Debt at COD (HTM!G157)': G77 + G79,
        'Lifetime generation GWh (HTM!G151)': sum(r38),
        'Average achieved price (HTM!G153)': sum(r54) / sum(r38),
    }
    costs = {
        'Devex incl. land (Calc!CQ = −HTM!G16)': -sum(r16),
        'Signing fee and DD (CR = −G17)': -sum(r17),
        'Turbine supply (CS)': I[162] * (1 + I[171]) * mw,
        'Other wind capex (CT)': (I[163] + I[164] + I[165] + I[166] + I[169] + I[170]) * (1 + I[171]) * mw,
        'Substation (CU)': I[167] * (1 + I[171]),
        'BESS capex (CV)': I[168] * bess,
        'Grid connection fee (CW = −G28)': -sum(r28),
        'Municipal one-off (CX = −G29)': -sum(r29),
        'Total construction capex (CY = −G30)': -G30,
        'Tender securities in cash (CZ = −(G18+G19))': -(sum(r18) + sum(r19)),
        'Guarantee fees to COD (DA = −G23)': -sum(r23),
        'IDC and financing fees (DB = G78 − G85)': sum(r78) - sum(r85),
    }
    costs['All-in cost to COD (Bora row 87)'] = (costs['Devex incl. land (Calc!CQ = −HTM!G16)'] + costs['Signing fee and DD (CR = −G17)']
                                                 + costs['Total construction capex (CY = −G30)'] + costs['Tender securities in cash (CZ = −(G18+G19))']
                                                 + costs['Guarantee fees to COD (DA = −G23)'] + costs['IDC and financing fees (DB = G78 − G85)'])
    costs['Installed wind MW (CD)'] = mw
    costs['All-in cost per installed MW (Bora row 89)'] = costs['All-in cost to COD (Bora row 87)'] / mw

    out = {'strategy': strategy, 'route': route, 'installed_mw': mw, 'committed_mw': E29, 'bess_mw': bess,
           'debt_size': E75, 'upfront_fee': E75 * I[219],
           'annual': {'years': yr, 'fep_cf': r129, 'dev': r24, 'capex': r27, 'conn': r28, 'municipal': r29, 'draw': r77,
                      'fees': r85, 'idc': r78, 'interest': r81, 'principal': r83, 'sweep': r87, 'dsra': r89,
                      'revenue': r54, 'opex': r69, 'ebitda': r70, 'tax': r100, 'decom': r107, 'gen_gwh': r38,
                      'dep': r92, 'gfee': r23},
           'grid_bottleneck_loss': clip, 'cost_to_cod': costs, 'pb_forfeit': forfeit}

    # 17. sale at COD (rows 333–380)
    tax = I[12] + I[197] * I[198]
    ref_comm = DATA['strategies']['S1']['community_eur_per_mwh']           # Tender E13 (S1 reference)
    df_cod = [1 / (1 + I[224]) ** (max(0, (dt.date(y, 12, 31) - cod).days) / 365) for y in yr]   # Tender row 198
    E225 = T['Tender!E30'] * sum(-(I[234] - ref_comm) * r38[i] * (1 - tax) * df_cod[i] for i in range(N))
    if cod_sale:
        price_mw = T['Calc!E133' if turb == 'EU' else 'Calc!E134'] * (1 - clip)   # Tender J46 / J47 → Inputs E230
        pre = [1 if y <= ycod else 0 for y in yr]
        flag = [1 if y == ycod else 0 for y in yr]
        r335 = [r24[i] * pre[i] for i in range(N)]
        r336 = [r30[i] * pre[i] for i in range(N)]
        r337 = [r77[i] * pre[i] for i in range(N)]
        r338 = [r117[i] * pre[i] for i in range(N)]
        r339 = [-I[203] * mw * (1 + esc) ** (y - ycod) * r14[i] * pre[i] for i, y in enumerate(yr)]
        r340 = [flag[i] * (price_mw * mw + E225) for i in range(N)]
        r341 = [-flag[i] * (r76[i] + r77[i] + r79[i]) for i in range(N)]
        r342 = [-flag[i] * I[26] for i in range(N)]
        r343 = [r335[i] + r336[i] + r337[i] + r338[i] + r339[i] + r340[i] + r341[i] + r342[i] for i in range(N)]
        G343 = sum(r343)
        r344 = [-flag[i] * I[231] * I[12] * max(0, G343) for i in range(N)]
        G340, G341, G342, G344 = sum(r340), sum(r341), sum(r342), sum(r344)
        G365 = G340 + G341 + G342 + G344
        moic_G20 = sum(land)
        moic_G124 = sum(l['O'] * l['P'] * (1 - l['U']) for l in lines)       # MOIC rows 56 − 102 (E106 = 1)
        G366 = -(-sum(r335) - sum(r336) - sum(r337) - sum(r338) - sum(r339) + sum(r17) + sum(r23)
                 + (1 - I[71]) * (sum(r16) + (moic_G20 - moic_G124) * E15))
        # Deposits AC:AF – FEP-convention interest per line to the COD sale, interest pause excluded
        pause_s, pause_e = todate(I[64]), todate(I[73])

        def acc(w, k, end):
            if k > end:
                return 1.0
            if claims == 'daily':
                f = 1.0
                for yy in range(k.year, end.year + 1):
                    a, b = max(k, dt.date(yy, 1, 1)), min(end, dt.date(yy, 12, 31))
                    f *= 1 + w * ((b - a).days + 1) / 365
                return f
            return ((1 + w * ((dt.date(k.year, 12, 31) - k).days + 1) / 365) * (1 + w) ** (end.year - k.year - 1)
                    * (1 + w * ((end - dt.date(end.year, 1, 1)).days + 1) / 365))
        AF = []
        for l in lines:
            w = I[61] if l['P'] == 1 else I[65]
            ac = acc(w, l['K'], min(cod, pause_s))
            aa = max(l['K'], pause_e)
            ad = 1.0 if cod < aa else acc(w, aa, cod)
            AF.append(l['O'] * (ac * ad - 1))
        G367 = -(sum(AF) + (itf['MOIC!G280 devex/land return to the COD sale (k€)'] if devex is None else devex['g280']))
        G375 = I[38] * min(I[37], I[36] * yearfrac30360(todate(I[70]), max(todate(I[70]), cod))) * mw
        G368 = -G375 * (1 - I[227])
        G369 = G365 + G366 + G367 + G368
        G371 = max(G368, G369 * I[34])
        if cod_waterfall == 'capped':                                   # MIN(due, cash left) per step, no floor
            left = max(0.0, G340 + G341 + G344)
            paid = []
            for due in (-G342, -G366, -G367, -G368):
                paid.append(min(due, left))
                left -= paid[-1]
            G366, G367, G368, G369 = -paid[1], -paid[2], -paid[3], left
            G371 = G369 * I[34]
        G370 = G369 - G371
        G372 = -G366 - G367 + G370
        r359 = [r335[i] + r336[i] + r337[i] + r338[i] + r339[i] for i in range(N)]
        r360 = [flag[i] * G372 for i in range(N)]
        r346 = [a + b for a, b in zip(r359, r360)]
        d362 = [(ser(max(dt.date(y, 1, 1), inv)) + ser(max(min(dt.date(y, 12, 31), cod), dt.date(y, 1, 1)))) / 2 for y in yr]
        fl, ds = r359 + [G372], d362 + [ser(cod)]
        G359, G360 = sum(r359), sum(r360)
        c, cmin = 0.0, 0.0
        for v in r359:
            c += v
            cmin = min(cmin, c)
        npv354 = npv(r346, I[253])
        out['cod'] = {
            'Sale price per gross MW after bottleneck loss (Inputs!E230)': price_mw,
            'COD price effect (Tender!E225)': E225,
            'Sale price at COD (HTM!G340)': G340,
            'Senior debt repaid at COD (HTM!G357)': -G341,
            'Equity value at COD (HTM!G356)': G340 + G341 + G342,
            'Net equity proceeds (G365)': G365,
            'FEP funding returned (G366)': G366,
            'FEP Stage-1 returns (G367)': G367,
            'Developer success fee (G368)': G368,
            'Distributable (G369)': G369,
            'To developer (G371)': G371,
            'FEP proceeds at COD (G372)': G372,
            'FEP contributions (G359)': G359,
            # Excel and LibreOffice truncate XIRR dates to whole days (row 362 holds x.5 mid-points)
            'FEP IRR (HTM!G352, XIRR rows 359/362)': xirr(fl, [math.floor(x) for x in ds]),
            'Memo: XIRR on the untruncated mid-point dates': xirr(fl, ds),
            'FEP MOIC (HTM!G353)': G360 / -G359 if G359 < 0 else None,
            'FEP NPV (HTM!G354)': npv354,
            'FEP NPV/Equity (Run_Map!F26)': npv354 / -npv(r359, I[253]),
            'FEP net profit post-tax (HTM!G351)': sum(r346),
            'FEP peak exposure (HTM!G355)': -cmin,
        }
    out['htm'] = htm
    if detail:
        out['rows'] = {k: v for k, v in locals().items() if k.startswith('r') and k[1:].isdigit() and isinstance(v, list)}
        out['lines'] = lines
    return out


# ----------------------------------------------------------------------------------------------- comparison
RUNS = {('S1', 'COD-EU'): 2, ('S1', 'COD-CN'): 3, ('S1', 'HTM-EU'): 4, ('S1', 'HTM-CN'): 5,
        ('S2', 'COD-EU'): 7, ('S2', 'COD-CN'): 8, ('S2', 'HTM-EU'): 9, ('S2', 'HTM-CN'): 10,
        ('S3', 'COD-EU'): 12, ('S3', 'COD-CN'): 13, ('S3', 'HTM-EU'): 14, ('S3', 'HTM-CN'): 15}
CMP_HTM = [('H', 'FEP net profit post-tax (HTM!G144)'), ('G', 'FEP net profit pre-tax (Run_Map!D22)'), ('I', 'FEP IRR (HTM!G139, IRR on row 129)'),
           ('J', 'FEP NPV (HTM!G140)'), ('K', 'FEP NPV/Equity (HTM!G142)'), ('L', 'FEP MOIC (HTM!G147)'), ('M', 'FEP peak exposure (HTM!G146)'),
           ('N', 'Project NPV (HTM!G136)'), ('AF', 'Project IRR post-tax (HTM!G135)'), ('AG', 'Equity IRR 100% (HTM!G137)'), ('AI', 'Min DSCR (HTM!G126)')]
CMP_COD = [('H', 'FEP net profit post-tax (HTM!G351)'), ('I', 'FEP IRR (HTM!G352, XIRR rows 359/362)'), ('J', 'FEP NPV (HTM!G354)'),
           ('K', 'FEP NPV/Equity (Run_Map!F26)'), ('L', 'FEP MOIC (HTM!G353)'), ('M', 'FEP peak exposure (HTM!G355)'),
           ('BW', 'Sale price at COD (HTM!G340)'), ('BX', 'COD price effect (Tender!E225)')]
CMP_COST = [('CQ', 'Devex incl. land (Calc!CQ = −HTM!G16)'), ('CR', 'Signing fee and DD (CR = −G17)'), ('CS', 'Turbine supply (CS)'),
            ('CT', 'Other wind capex (CT)'), ('CU', 'Substation (CU)'), ('CV', 'BESS capex (CV)'), ('CW', 'Grid connection fee (CW = −G28)'),
            ('CX', 'Municipal one-off (CX = −G29)'), ('CY', 'Total construction capex (CY = −G30)'),
            ('CZ', 'Tender securities in cash (CZ = −(G18+G19))'), ('DA', 'Guarantee fees to COD (DA = −G23)'),
            ('DB', 'IDC and financing fees (DB = G78 − G85)'), ('CD', 'Installed wind MW (CD)')]


def _rel(a, b, is_rate):
    if a is None or b is None or isinstance(b, str):
        return float('nan')
    return abs(a - b) if is_rate else abs(a - b) / max(abs(b), 1e-9)


def compare(verbose=True):
    with open(os.path.join(HERE, 'lo_runs_htm.json'), encoding='utf-8') as f:
        lo = json.load(f)
    worst = {'value': 0.0, 'rate': 0.0}
    lines = []
    hdr = f"{'run':>3} {'strategy':8} {'route':7} {'metric':48} {'shadow':>18} {'saved (Excel)':>18} {'LO live':>18} {'Δ saved':>9} {'Δ LO':>9}"
    lines.append(hdr)
    lines.append('-' * len(hdr))
    for (s, route), run in RUNS.items():
        res = compute(s, route)
        sv, lv = lo['saved'][str(run)], lo[str(run)]['Calc!305']
        lv = {k.split(' ')[0]: v for k, v in lv.items()}
        sel = (CMP_HTM if route.startswith('HTM') else CMP_COD) + (CMP_COST if route.startswith('HTM') else [])
        for col, key in sel:
            mine = res['htm'].get(key) if key in res['htm'] else res.get('cod', {}).get(key, res['cost_to_cod'].get(key))
            a, b = sv.get(f'Calc!{col}{305 + run}'), lv.get(f'Calc!{col}305')
            rate = col in ('I', 'AF', 'AG')
            da, db = _rel(mine, a, rate), _rel(mine, b, rate)
            for d in (da, db):
                if d == d:
                    k = 'rate' if rate else 'value'
                    worst[k] = max(worst[k], d)
            name = f'{key[:46]}'
            fmt = (lambda x: f'{x:18.10f}') if rate else (lambda x: f'{x:18.4f}' if isinstance(x, (int, float)) else f'{str(x):>18}')
            lines.append(f'{run:>3} {s:8} {route:7} {name:48} {fmt(mine)} {fmt(a)} {fmt(b)} {da:9.1e} {db:9.1e}')
    if verbose:
        print('\n'.join(lines))
        print(f"\nWorst difference: values (relative) {worst['value']:.2e}; IRRs (absolute) {worst['rate']:.2e}")
        print('Δ saved = vs the cached Excel data table (Calc!G306:DH330 = Bora rows 63–106); Δ LO = vs LibreOffice live '
              '(Calc!E10 = run, data table cleared). Relative for values, absolute for IRRs. The COD XIRR matches LibreOffice '
              'exactly and the saved Excel value to Excel\'s own XIRR tolerance (~1e-8).')
    return worst


def toggles():
    """Materiality of the mechanics: Δ FEP IRR (pp) and Δ FEP NPV (k€) vs the old behaviour."""
    tests = [
        ('deemed exit charge off (FEP share kept 100 %)', dict(deemed_exit=False, fep_share=1.0)),
        ('deemed exit off, old rule (FEP share → 70 %)', dict(deemed_exit=False)),
        ('deemed exit on the RtB-route MOIC result', dict(deemed_exit_basis='rtb_route')),
        ('DSRA off', dict(dsra=False)),
        ('cash sweep off', dict(sweep=False)),
        ('DSRA and sweep off', dict(dsra=False, sweep=False)),
        ('loss cap off (100 % usable)', dict(overrides={'E214': 1.0})),
        ('commitment fee off', dict(overrides={'E222': 0.0})),
        ('upfront fee off', dict(overrides={'E219': 0.0})),
        ('IDC paid by equity, not capitalised', dict(overrides={'E221': 0.0})),
        ('degradation off', dict(overrides={'E182': 0.0})),
        ('opex contingency off', dict(overrides={'E208': 0.0})),
        ('HIPA off', dict(overrides={'E210': 0.0})),
        ('decommissioning off', dict(overrides={'E175': 0.0})),
        ('devex etc. excluded (Inputs E226 = 0)', dict(overrides={'E226': 0.0})),
        ('PB held 6 m after COD (HTM)', dict(pb_hold_months=6)),
        ('no bid-bond netting', dict(bid_netting=False)),
        ('FX 400 HUF/EUR', dict(fx=400.0)),
        ('direct fee as Annex 1 lump sum', dict(direct_fee_basis='lump')),
        ('overbuild not built', dict(built_overbuild=False)),
        ('overbuild not built + PB call', dict(built_overbuild=False, pb_forfeit=True)),
        ('BESS not built', dict(bess_built=False)),
        ('BESS not built + PB call', dict(bess_built=False, pb_forfeit=True)),
        ('community €/MWh = 0', dict(overrides={'E234': 0.0})),
        ('CIT 0', dict(overrides={'E12': 0.0})),
    ]
    cases = [('S2', 'HTM-EU'), ('S3', 'HTM-EU'), ('S2', 'HTM-CN'), ('S2', 'COD-EU')]
    base = {c: compute(*c) for c in cases}
    key = lambda r, c: (r['htm']['FEP IRR (HTM!G139, IRR on row 129)'], r['htm']['FEP NPV (HTM!G140)']) if c[1].startswith('HTM') \
        else (r['cod']['FEP IRR (HTM!G352, XIRR rows 359/362)'], r['cod']['FEP NPV (HTM!G354)'])
    print(f"{'toggle':48}" + ''.join(f"{c[0] + ' ' + c[1]:>26}" for c in cases))
    print(f"{'base FEP IRR / NPV k€':48}" + ''.join(f"{key(base[c], c)[0] * 100:12.3f}% {key(base[c], c)[1]:12.0f}" for c in cases))
    for name, kw in tests:
        row = f'{name:48}'
        for c in cases:
            try:
                b0, n0 = key(base[c], c)
                b1, n1 = key(compute(*c, **kw), c)
                row += f'{(b1 - b0) * 100:+11.3f}pp {n1 - n0:+12.0f}'
            except Exception as e:  # noqa: BLE001
                row += f'{"n/a":>26}'
        print(row)


if __name__ == '__main__':
    if '--toggles' in sys.argv:
        toggles()
    else:
        compare()
