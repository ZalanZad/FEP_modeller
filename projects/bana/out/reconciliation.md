# Bana – reconciliation to the old engine

Old engine: `reference/Bana_IC_model_v3.xlsm`, figures as shown in `reference/Bana_IC_20260109_vDSz.pptx` (slides 5–6). New: `out/Bana_Model.xlsm`, base set. Each step below is cumulative: one old input put back into the new model, tax solver re-run (`reconcile.py`).

## European turbine

| Step | Total uses to COD | FEP equity to COD | FEP profit – sale at COD | FEP IRR – sale at COD | FEP MOIC – sale at COD | FEP IRR – hold to maturity | FEP MOIC – hold to maturity | Minimum DSCR |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| New model, base | 83.42 | 22.55 | 38.58 | 61.0% | 2.71x | 12.2% | 6.27x | 1.58x |
| + EU turbine 950 k€/MW | 102.87 | 26.64 | 19.13 | 34.9% | 1.72x | 8.3% | 4.55x | 1.28x |
| + Debt share cap 60 % EU / 50 % CN | 101.50 | 44.33 | 20.50 | 29.7% | 1.46x | 7.7% | 3.43x | 1.70x |
| + Tenor 10 years | 101.22 | 48.03 | 20.78 | 28.8% | 1.43x | 7.1% | 3.51x | 1.20x |
| **Old engine (deck 09-Oct-2026)** | 101.75 | 44.23 | 20.25 | 26.4% | 1.46x | 9.2% | 5.64x | 1.09x |

## Chinese turbine

| Step | Total uses to COD | FEP equity to COD | FEP profit – sale at COD | FEP IRR – sale at COD | FEP MOIC – sale at COD | FEP IRR – hold to maturity | FEP MOIC – hold to maturity | Minimum DSCR |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| New model, base | 76.03 | 27.53 | 28.47 | 46.8% | 2.03x | 12.8% | 5.84x | 1.97x |
| + EU turbine 950 k€/MW | 76.03 | 27.53 | 28.47 | 46.8% | 2.03x | 12.8% | 5.84x | 1.97x |
| + Debt share cap 60 % EU / 50 % CN | 75.02 | 40.38 | 29.48 | 41.6% | 1.73x | 11.4% | 4.54x | 2.75x |
| + Tenor 10 years | 75.02 | 40.38 | 29.48 | 41.6% | 1.73x | 10.5% | 4.70x | 1.81x |
| **Old engine (deck 09-Oct-2026)** | 75.15 | 40.32 | 29.35 | 37.7% | 1.73x | 12.1% | 7.09x | 1.77x |

## Step sources

- EU turbine 950 k€/MW: Bana v3 Inputs E162; new: Bora user decision 600.
- Debt share cap 60 % EU / 50 % CN: Bana v3 Inputs E217.
- Tenor 10 years: Bana v3 Inputs E220.

## What the steps do not reproduce (change of method, by design)

- Revenue: the old hold case had the 90 €/MWh PPA for 10 years and Aurora Central merchant prices after (average 134.7 €/MWh over the life); the new model has the PPA for the whole life (user, 13-Oct-2026).
- Debt: old = gearing-sized annuity from the first full operating year with a 50 % cash sweep for 5 years; new = debt share of cost capped by the DSCR capacity, sculpted repayment, no sweep (Rock / Alpha).
- Timing: old purchase price and DD on 31-Jul-2026 (monthly grid); new price at closing 31-Dec-2026 after the award and DD at the bid date, quarterly grid.
- Tax: old 50 % loss cap with devex expensed and decommissioning at 30 k€/MW; new capitalises the purchase price, DD and devex into the 20-year depreciation base and has no decommissioning (immaterial at 8 % over 30 years).
- Opex: the new model adds the 0.3 % innovation contribution on the HIPA base.

The turbine step only moves the European route. Tax losses still unused 5 years after the NTP lapse in every step (Tao. 17. §), so the steps carry no tax shield the rule would not allow.
