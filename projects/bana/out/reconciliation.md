# Bana – reconciliation to the old engine

Old engine: `reference/Bana_IC_model_v3.xlsm`, figures as shown in `reference/Bana_IC_20260109_vDSz.pptx` (slides 5–6). New: `out/Bana_Model.xlsm`, base set. Each step below is cumulative: one old input put back into the new model, tax solver re-run (`reconcile.py`).

## European turbine

| Step | Total uses to COD | FEP equity to COD | FEP profit – sale at COD | FEP IRR – sale at COD | FEP MOIC – sale at COD | FEP IRR – hold to maturity | FEP MOIC – hold to maturity | Minimum DSCR |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| New model, base | 83.38 | 22.51 | 38.62 | 60.6% | 2.72x | 12.2% | 6.28x | 1.59x |
| + EU turbine 950 k€/MW | 102.82 | 26.59 | 19.18 | 34.8% | 1.72x | 8.3% | 4.57x | 1.29x |
| + Debt share cap 60 % EU / 50 % CN | 101.47 | 44.30 | 20.53 | 29.6% | 1.46x | 7.7% | 3.44x | 1.71x |
| + Tenor 10 years | 101.20 | 47.80 | 20.80 | 28.8% | 1.44x | 7.1% | 3.52x | 1.20x |
| **Old engine (deck 09-Oct-2026)** | 101.75 | 44.23 | 20.25 | 26.4% | 1.46x | 9.2% | 5.64x | 1.09x |

## Chinese turbine

| Step | Total uses to COD | FEP equity to COD | FEP profit – sale at COD | FEP IRR – sale at COD | FEP MOIC – sale at COD | FEP IRR – hold to maturity | FEP MOIC – hold to maturity | Minimum DSCR |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| New model, base | 76.00 | 27.50 | 28.50 | 46.6% | 2.04x | 12.8% | 5.84x | 1.98x |
| + EU turbine 950 k€/MW | 76.00 | 27.50 | 28.50 | 46.6% | 2.04x | 12.8% | 5.84x | 1.98x |
| + Debt share cap 60 % EU / 50 % CN | 75.01 | 40.36 | 29.49 | 41.4% | 1.73x | 11.3% | 4.54x | 2.75x |
| + Tenor 10 years | 75.01 | 40.36 | 29.49 | 41.4% | 1.73x | 10.5% | 4.70x | 1.81x |
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

The turbine step only moves the European route. At the old inputs the European route breaches the new model's 5-year tax-loss check, which is a finding for that case, not for the base.
