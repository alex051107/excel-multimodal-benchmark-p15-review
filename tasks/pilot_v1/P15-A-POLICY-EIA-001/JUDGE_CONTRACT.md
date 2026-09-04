# Judge V3 contract

## Version boundary

The archived pre-v2 campaign used a prompt whose provenance, GWh-to-MWh conversion, and base-gas/Other-generation relationship were not stated consistently. The frozen campaign replay assigns those archived attempts `TASK_INVALID` before evaluator dispatch.

This directory is now task version 2.0.0. Its instruction, inputs, oracle, reference, and Judge all implement the corrected contract below. `tests/evaluate.py` always scores that current contract; it has no environment-controlled or CLI-controlled historical mode. This prevents the same evaluator file from carrying two incompatible meanings.

## Source and scenario boundary

The public context is EIA Electric Power Annual table 3.1.A, whose published unit is thousand megawatthours (numerically equal to GWh). The supplied values are benchmark-rounded/adapted scenario inputs, not a verbatim EIA table extract. The policy assumptions and emissions factors are benchmark-authored.

## What the task explicitly asks for

- Preserve the supplied 2023 coal, natural-gas, wind, solar, demand, and emissions-factor inputs.
- Use historical generation as the base case. Other generation is the residual needed to reconcile the supplied demand and is held constant in the policy case.
- In the policy case, displace 8% of coal, increase wind 5%, increase solar 12%, grow demand 1%, and use natural gas to balance the remaining demand.
- Calculate emissions and intensity and report the emissions reduction, intensity reduction, and gas change.
- Keep scenario calculations formula-linked to the supplied inputs and assumptions.

## Unit rule

Generation is in GWh and emissions factors are in metric tonnes CO2 per MWh. Absolute emissions in metric tonnes therefore equal `GWh × 1,000 × tCO2/MWh`. Intensity equals metric tonnes divided by `GWh × 1,000` and is reported in tCO2/MWh.

## Accepted layouts and equivalent implementations

- A combined base/policy scenario table.
- Separate `Base` and `Policy` sheets.
- Renamed source, assumption, result, and check sheets.
- An explicit `Other generation` row or an algebraically equivalent gas-balancing formula that holds the same residual generation constant.
- Positive reduction (`base - policy`) or a clearly labeled signed change (`policy - base`).

The Judge locates the scenario by labels and relationships, not fixed sheet names or row numbers. Wrong units, changed source values, a broken demand balance, or a result that does not respond to the assumptions remains a substantive failure.
