# Bana – reconciliation to the old engine

Old engine: `reference/Bana_IC_model_v3.xlsm`, figures as shown in `reference/Bana_IC_20260109_vDSz.pptx` (slides 5–6). New: `out/Bana_Model.xlsm`, base set. Each step below is cumulative: one old input put back into the new model, tax solver re-run (`reconcile.py`).

## European turbine

| Step | Total uses to COD | FEP equity to COD | FEP profit – sale at COD | FEP IRR – sale at COD | FEP MOIC – sale at COD | FEP IRR – hold to maturity | FEP MOIC – hold to maturity | Minimum DSCR |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| New model, base | 88.82 | 23.76 | 44.82 | 66.4% | 2.89x | 12.9% | 6.60x | 1.62x |
| + EU turbine 950 k€/MW | 110.21 | 28.26 | 23.42 | 39.9% | 1.83x | 8.8% | 4.77x | 1.31x |
| + Debt share cap 60 % EU / 50 % CN | 108.75 | 47.28 | 24.89 | 33.8% | 1.53x | 8.0% | 3.56x | 1.73x |
| + Tenor 10 years | 108.52 | 50.29 | 25.12 | 33.0% | 1.50x | 7.4% | 3.67x | 1.20x |
| **Old engine (deck 09-Oct-2026)** | 101.75 | 44.23 | 20.25 | 26.4% | 1.46x | 9.2% | 5.64x | 1.09x |

## Chinese turbine

| Step | Total uses to COD | FEP equity to COD | FEP profit – sale at COD | FEP IRR – sale at COD | FEP MOIC – sale at COD | FEP IRR – hold to maturity | FEP MOIC – hold to maturity | Minimum DSCR |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| New model, base | 80.84 | 29.05 | 33.64 | 51.8% | 2.16x | 13.5% | 6.12x | 2.02x |
| + EU turbine 950 k€/MW | 80.84 | 29.05 | 33.64 | 51.8% | 2.16x | 13.5% | 6.12x | 2.02x |
| + Debt share cap 60 % EU / 50 % CN | 79.77 | 42.78 | 34.70 | 46.1% | 1.81x | 11.9% | 4.73x | 2.81x |
| + Tenor 10 years | 79.77 | 42.78 | 34.70 | 46.1% | 1.81x | 11.0% | 4.88x | 1.85x |
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
- Overbuild: the old bid had none (50 MW); the new bid commits 20 % and builds 10 % (55 MW, user 13-Oct-2026), which adds capex, generation, the COD price on the extra MW and a larger security base.

The turbine step only moves the European route. Tax losses still unused 5 years after the NTP lapse in every step (Tao. 17. §), so the steps carry no tax shield the rule would not allow.
