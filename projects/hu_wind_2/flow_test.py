"""Flow-through test of HU_Wind_bids_v3 (SPEC bora_v5_summary §31).

Shocks one input at a time in LibreOffice, recalculates, reads the headline outputs of each project, the Master check
and the count of error cells on both sheets, then restores the input. Every shock must give 0 errors and Master
check 0, except "HCSSZ before sale A", where the check must fire. Prints which outputs moved and by how much.

    python3 projects/hu_wind_2/flow_test.py [out.json]
"""
import sys, json
from datetime import date
sys.path.insert(0, "/home/user/FEP_modeller")
from fepm.recalc import Office

WB = "/home/user/FEP_modeller/projects/hu_wind_2/out/HU_Wind_bids_v3.xlsx"
OUT = {  # label: Dashboard row
    "points": 43, "MW built": 46, "all-in cost": 99,
    "A profit post": 113, "A IRR post": 117, "B profit post": 121, "B IRR post": 125,
    "COD price": 134, "COD IRR": 137, "revenue": 140, "EBITDA": 142, "proj IRR pre": 143,
    "CIT unlev": 145, "proj IRR post": 146, "debt": 147, "equity IRR": 148, "min DSCR": 149}
COLS = "EFG"
serial = lambda d: (d - date(1899, 12, 30)).days
SHOCKS = [
    ("Bora P50 2500→2750 h", "Dashboard!E22", 2750),
    ("Base-case PPA 90→100 €/MWh", "Dashboard!E28", 100),
    ("Case selector → Low", "Dashboard!H25", "Low"),
    ("Case selector → High", "Dashboard!H25", "High"),
    ("Base CN turbine 500→550 k€/MW", "Dashboard!E27", 550),
    ("Base O&M 12→14 k€/MW/yr", "Dashboard!E30", 14),
    ("Bora turbine route CN→EU", "Dashboard!E17", "EU"),
    ("Iváncsa MVA 50→60", "Dashboard!F6", 60),
    ("Bora overbuild 20→51 %", "Dashboard!E7", 0.51),
    ("Bora BESS 10→51 %", "Dashboard!E8", 0.51),
    ("Dunaújváros build shortfall 10→0 %", "Dashboard!G10", 0),
    ("Bora municipal 20→30 k€/MVA", "Dashboard!E11", 30),
    ("Iváncsa community 3→4 €/MWh", "Dashboard!F12", 4),
    ("Bora RtB premium 150→200 k€/MVA", "Dashboard!E18", 200),
    ("Bora FEP share 70→60 %", "Dashboard!E19", 0.6),
    ("Bora JV off", "Dashboard!E21", 0),
    ("Signing+DD Bora 135→200 k€", "Inputs!E17", 200),
    ("CIT 15→9 %", "CIT_Rate", 0.09),
    ("Share exemption off", "Share_Exempt", 0),
    ("Gearing 60→70 %", "Gearing", 0.7),
    ("Debt rate 5→6 %", "Debt_Rate", 0.06),
    ("Debt tenor 15→18 y", "Debt_Tenor", 18),
    ("Tax depreciation 20→15 y", "Tax_Dep_Life", 15),
    ("Buyer discount rate 8→9 %", "Discount_Rate", 0.09),
    ("FX 365→400 HUF/€", "FX", 400),
    ("Curtailment 2→5 %", "Curtailment", 0.05),
    ("Degradation 0.3→0.5 %", "Degradation", 0.005),
    ("PPA escalation 1→2 %", "PPA_Esc", 0.02),
    ("Opex escalation 2.5→3.5 %", "Opex_Esc", 0.035),
    ("Operating life 30→25 y", "Life", 25),
    ("Sale date B 30-Sep-29→31-Mar-29", "Sale_Date", serial(date(2029, 3, 31))),
    ("Iváncsa COD +1 y (30-Jun-2032)", "Inputs!F26", serial(date(2032, 6, 30))),
    ("Bora NTP 30-Jun-29→31-Dec-29", "Inputs!E25", serial(date(2029, 12, 31))),
    ("Bora HCSSZ 01-Oct-28→01-Mar-29", "Inputs!E24", serial(date(2029, 3, 1))),
    ("Bora HCSSZ 01-Oct-28→01-Sep-28 (before sale A: check must fire)", "Inputs!E24", serial(date(2028, 9, 1))),
    ("Dunaújváros bid date 30-Oct-26→30-Nov-26", "Inputs!G21", serial(date(2026, 11, 30))),
    ("Bora devex scale 0.25→0.50", "Inputs!E18", 0.5),
    ("Iváncsa devex scale 0.25→0.50", "Inputs!F18", 0.5),
    ("Sale date A 30-Sep-28→30-Jun-28", "Sale_Date_A", serial(date(2028, 6, 30))),
    ("Freeze date 30-Jun-29→31-Dec-28", "Freeze_Date", serial(date(2028, 12, 31))),
    ("Bora indirect fee +10 %", "Inputs!E14", 73071306 * 1.1),
    ("Bora perf. security released 2033→2032", "Inputs!E27", serial(date(2032, 3, 31))),
]

def read(o):
    r = {f"{k}|{c}": o.get(f"Dashboard!{c}{row}") for k, row in OUT.items() for c in COLS}
    r["master"] = o.get("Dashboard!H3")
    errs = 0
    for name in ("Dashboard", "Inputs"):
        sh = o.doc.Sheets.getByName(name)
        q = sh.queryFormulaCells(4)  # FormulaResult.ERROR
        for i in range(q.Count):
            a = q.getByIndex(i).RangeAddress
            errs += (a.EndRow - a.StartRow + 1) * (a.EndColumn - a.StartColumn + 1)
    r["errors"] = errs
    return r

res = {}
with Office(WB) as o:
    o.calculate(); base = read(o); res["BASE"] = base
    for label, ref, v in SHOCKS:
        old = o.get(ref)
        o.set(ref, v); o.calculate(); res[label] = read(o)
        o.set(ref, old); o.calculate()
    back = read(o)
assert all(abs((back[k] or 0) - (base[k] or 0)) < 1e-9 if isinstance(base[k], float) else back[k] == base[k] for k in base), "restore failed"
if len(sys.argv) > 1:
    json.dump(res, open(sys.argv[1], "w"), indent=1, default=str)
bad = []
for s, v in res.items():
    must_fire = "check must fire" in s
    ok = v["errors"] == 0 and ((v["master"] != 0) if must_fire else (v["master"] == 0))
    moved = [k for k in base if "|" in k and isinstance(base[k], float) and isinstance(v[k], float)
             and abs(v[k] - base[k]) > 1e-9 * max(1, abs(base[k]))]
    print(f"{'OK ' if ok else 'BAD'} {s:62s} master {v['master']:g} errors {v['errors']} outputs moved {len(moved)}")
    bad += [] if ok else [s]
print("FAILED:", bad) if bad else print(f"all {len(res)} runs pass")
sys.exit(1 if bad else 0)
