---
id: R15-A-FORECAST-QC-001
track: A
status: DESIGN_FROZEN
package_not_built: true
harbor_package_built: false
primary_eligible: false
split_group: R15-A-FORECAST-QC-001
---

# Frozen-vintage demand-forecast quality control

## Professional work and user scenario

A supply-chain planner must audit weekly forecasts across SKUs, regions, vintages, and horizons. The central risk is look-ahead leakage: the workbook must select the forecast actually available at the required lead time, not the newest row in the file. It must quantify WAPE, bias, naive-baseline accuracy, forecast value added, data-quality failures, and exception priorities.

## Instruction draft

“Build forecast QC from the supplied forecast vintages, actuals, fiscal calendar, and thresholds. For each SKU-region-target week and requested horizon, select the latest forecast published on or before the permitted cutoff. Detect duplicate or missing keys before calculating accuracy. Calculate aggregate WAPE, normalized bias, naive-baseline WAPE, and FVA by SKU, region, horizon, and period. Flag breaches and zero-demand groups under the stated policy and preserve source keys and chosen vintage. Do not average vintages or use information published after cutoff.”

## Packaged-input design and authenticity boundary

- Planned inputs: forecast_vintages.csv, actuals.csv, fiscal_calendar.csv, qc_thresholds.csv, forecast_policy.md, and starting_workbook.xlsx.
- Forecast rows contain SKU, region, published week, target week, horizon, and quantity.
- Actuals contain unique SKU-region-target-week values, including one zero-demand group and one missing eligible vintage.
- The visible policy defines cutoff selection, naive forecast, WAPE, bias, FVA, zero-demand handling, and exception priority.
- Data is synthetic and is not a real company forecast, service level, sales history, or planning decision.

## Target workbook and professional operations

Required sheets are Forecast_Vintages, Actuals, Aligned_Forecast, Accuracy, Bias_FVA, DQ_Exceptions, Management_View, and Checks. Work includes deterministic vintage selection, key validation, horizon alignment, aggregate metrics, naive baseline, FVA, thresholds, and exception ranking.

## Independent oracle design

A pandas oracle will enforce unique keys, select the maximum publication date not later than each cutoff, and compute error as forecast minus actual; WAPE as sum absolute error divided by sum actual; normalized bias as sum error divided by sum actual; the visible seasonal-lag naive baseline; and FVA as naive WAPE minus submitted-forecast WAPE. Zero denominators become explicit flags. The oracle emits row keys, chosen vintage, metrics, and flags without cell-address truth.

## Atomic rubric design

- R001: Required sheets exist and source populations are retained.
- R002: Duplicate and missing actual/forecast keys are detected before metrics.
- R003: The chosen forecast is the latest vintage available by cutoff.
- R004: Horizon and fiscal-period alignment is correct.
- R005: Absolute error and aggregate WAPE are formula-linked and correctly weighted.
- R006: Bias uses the stated sign and aggregation.
- R007: The naive baseline follows the visible seasonal-lag policy.
- R008: FVA is baseline WAPE minus submitted-forecast WAPE.
- R009: Zero-demand groups follow the explicit policy without suppressed division errors.
- R010: Threshold breaches and priorities identify the correct groups.
- R011: Each exception retains SKU, region, target week, horizon, and vintage.
- P001: Penalize look-ahead leakage, averaged vintages, mean row APE substitution, bias reversal, hidden imputation, or pasted management metrics.

## Acceptable materially different equivalent

One solution may use a long aligned ledger with helper keys and aggregate tables. A materially different equivalent may use horizon matrices with a separate vintage-selection table. Cutoff selection, metrics, zero-demand behavior, exceptions, and traceability must match.

## Negative fixtures

- No-op: source tables present and alignment, metrics, and exceptions blank.
- Malformed: unreadable workbook or missing required sheets.
- M1: Uses a forecast published after the cutoff.
- M2: Averages eligible vintages instead of selecting the latest eligible one.
- M3: Computes mean APE instead of aggregate WAPE.
- M4: Reverses the bias sign.
- M5: Treats zero demand as zero error without a flag.
- M6: Computes FVA against actuals rather than naive error.

## Two meaningful input perturbations

- P1: Change one eligible four-week-ahead forecast quantity from 410 to 470. Aligned forecast, errors, group WAPE, bias, FVA, status, and priority must propagate; actuals, chosen vintage, other horizons, and naive forecast remain invariant.
- P2: Tighten the WAPE threshold from 0.25 to 0.20. Metric values and vintages remain invariant; only newly crossing groups and exception counts/priorities change.

## Verifier-private CONFIRM sibling

The CONFIRM sibling uses different SKUs, regions, weeks, horizons, thresholds, and publication patterns, with a distinct missing-vintage and zero-demand case. It preserves cutoff and metric semantics. Confirm truth and reference remain verifier-private; policy and solvable data remain visible.

## Difficulty-change policy

Allowed difficulty changes: more legitimate horizons, a second region hierarchy, visible holiday-naive rules, rolling limits, or additional vintages.

Forbidden difficulty changes: secret cutoffs, undefined duplicate handling, impossible joins, hidden zero-demand policy, irrelevant high-volume rows, stale caches, or textual traps.

## Activation rule

Activate only if a Track A PRIMARY is invalid or remains VALID_BUT_EASY after three professional revisions, followed by full package and held-out-sibling validation.

## Build state

**package_not_built: true — no task directory, workbook, fixture, Judge, or Harbor package exists.**
