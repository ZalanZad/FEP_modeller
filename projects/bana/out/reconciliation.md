# Bana – reconciliation to the old engine

Old engine: `reference/Bana_IC_model_v3.xlsm`, figures as shown in `reference/Bana_IC_20260109_vDSz.pptx` (slides 5–6). New: `out/Bana_Model.xlsm`, base set. Each step below is cumulative: one old input put back into the new model, tax solver re-run (`reconcile.py`).

## European turbine

| Step | Total uses to COD | FEP equity to COD | FEP profit – sale at COD | FEP IRR – sale at COD | FEP MOIC – sale at COD | FEP IRR – hold to maturity | FEP MOIC – hold to maturity | Minimum DSCR |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| New model, base | 112.75 | 63.25 | 20.89 | 24.8% | 1.33x | 7.4% | 3.13x | 1.94x |
| + CIT 9 % | 112.75 | 63.25 | 20.89 | 24.8% | 1.33x | 7.7% | 3.28x | 1.96x |
| + P50 2,460 h | 112.75 | 63.25 | 20.89 | 24.8% | 1.33x | 6.9% | 3.00x | 1.84x |
| + Debt share cap 60 % (both routes; old CN 50 %) | 112.75 | 63.25 | 20.89 | 24.8% | 1.33x | 6.9% | 3.00x | 1.84x |
| + Tenor 10 years | 112.75 | 63.25 | 20.89 | 24.8% | 1.33x | 6.7% | 3.09x | 1.37x |
| **Old engine (deck 09-Oct-2026)** | 101.75 | 44.23 | 20.25 | 26.4% | 1.46x | 9.2% | 5.64x | 1.09x |

## Chinese turbine

| Step | Total uses to COD | FEP equity to COD | FEP profit – sale at COD | FEP IRR – sale at COD | FEP MOIC – sale at COD | FEP IRR – hold to maturity | FEP MOIC – hold to maturity | Minimum DSCR |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| New model, base | 85.52 | 36.02 | 28.95 | 43.1% | 1.80x | 11.8% | 5.31x | 1.89x |
| + CIT 9 % | 85.52 | 36.02 | 28.95 | 43.1% | 1.80x | 12.3% | 5.62x | 1.93x |
| + P50 2,460 h | 85.52 | 36.02 | 28.95 | 43.1% | 1.80x | 11.2% | 5.13x | 1.81x |
| + Debt share cap 60 % (both routes; old CN 50 %) | 85.37 | 38.02 | 29.11 | 42.3% | 1.77x | 11.1% | 4.95x | 1.89x |
| + Tenor 10 years | 85.37 | 38.02 | 29.11 | 42.3% | 1.77x | 10.4% | 5.08x | 1.41x |
| **Old engine (deck 09-Oct-2026)** | 75.15 | 40.32 | 29.35 | 37.7% | 1.73x | 12.1% | 7.09x | 1.77x |

## Step sources

- CIT 9 %: Bana v3 Inputs E12; new: 15 % Pillar Two minimum (user, 14-Oct-2026).
- P50 2,460 h: Bana v3 Inputs E179; new: 2,600 h (user, 13-Oct-2026).
- Debt share cap 60 % (both routes; old CN 50 %): Bana v3 Inputs E217.
- Tenor 10 years: Bana v3 Inputs E220.

## What the steps do not reproduce (change of method, by design)

- Revenue: the old hold case had the 90 €/MWh PPA for 10 years and Aurora Central merchant prices after (average 134.7 €/MWh over the life); the new model has the PPA for the whole life (user, 13-Oct-2026).
- Debt: old = gearing-sized annuity from the first full operating year with a 50 % cash sweep for 5 years; new = debt share of cost capped by the DSCR capacity, sculpted repayment, no sweep (Rock / Alpha).
- Timing: old purchase price and DD on 31-Jul-2026 (monthly grid); new price at closing 31-Dec-2026 after the award and DD at the bid date, quarterly grid.
- Tax: old 50 % loss cap with devex expensed and decommissioning at 30 k€/MW; new capitalises the purchase price, DD and devex into the 20-year depreciation base and has no decommissioning (immaterial at 8 % over 30 years).
- Opex: the new model adds the 0.3 % innovation contribution on the HIPA base.
- Overbuild: the old bid had none (50 MW); the new bid commits 20 % and builds 10 % (55 MW, user 13-Oct-2026), which adds capex, generation, the COD price on the extra MW and a larger security base.
- Bid and BESS: the old bid was municipal 35 k€/MW and community 2.5 €/MWh with no BESS; the new bid is 50 and 3 with a 10 MW / 4 h BESS at capex only (user, 13-Oct-2026).

The old engine geared the Chinese route at 50 %; the new model has one cap for both. Tax losses still unused 5 years after the NTP lapse in every step (Tao. 17. §), so the steps carry no tax shield the rule would not allow.
