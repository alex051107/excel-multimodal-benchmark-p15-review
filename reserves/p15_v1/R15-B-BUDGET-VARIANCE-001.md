---
id: R15-B-BUDGET-VARIANCE-001
track: B
status: DESIGN_FROZEN
package_not_built: true
harbor_package_built: false
primary_eligible: false
split_group: R15-B-BUDGET-VARIANCE-001
---

# Outpatient service-line flexible-budget variance analysis

## Professional work and user scenario

A hospital finance business partner must explain a monthly outpatient service-line result to operational leaders. Static budget, actual encounters, staffing hours, payer-adjusted revenue, supply units, payroll detail, and general-ledger activity must be converted into a flexible budget and a controlled rate, volume, price, mix, and efficiency variance bridge.

## Instruction draft

“Build the outpatient flexible-budget variance analysis in the supplied workbook. Map every GL and payroll row to the visible department and variance-driver rules; retain unmapped activity as an exception. Flex the approved budget to actual visit and procedure volumes, separate revenue volume, payer-mix, and yield effects, and separate labor rate, overtime, and efficiency effects without overlap. Produce department and consolidated bridges from static budget to actual, classify favorable and unfavorable results consistently, rank material explanations, and reconcile every bridge to the GL and payroll controls. Keep formulas, driver provenance, and exception evidence visible.”

## Packaged-input design and authenticity boundary

- Planned inputs: static_budget.csv, actual_activity.csv, payroll_detail.csv, gl_actuals.csv, payer_mix.csv, mapping_rules.csv, variance_conventions.md, and starting_workbook.xlsx.
- The budget will include volume drivers, unit revenue, productive-hour standards, wage assumptions, supply standards, and fixed or variable behavior by department.
- The visible convention will define sign display, bridge order, payer-mix benchmark, overtime allocation, favorable and unfavorable labels, and materiality.
- GL data will contain intentionally unmapped and cross-department rows that must be disclosed, not discarded.
- All facilities, payers, departments, personnel identifiers, and values will be synthetic and will not be presented as real patient, payroll, or financial data.
- No input file has been generated at DESIGN_FROZEN status.

## Target workbook and professional operations

Required sheets are Conventions, Budget_Raw, Activity_Raw, Payroll_Raw, GL_Raw, Mapping, Flexible_Budget, Revenue_Bridge, Labor_Bridge, Department_Summary, Exceptions, and Checks. The workbook must flex variable standards to actual activity, preserve fixed budget amounts, map actuals completely, compute nonoverlapping revenue and labor decompositions, classify signs, identify material drivers, and reconcile departmental plus consolidated bridges to the recorded result.

## Independent oracle design

A task-local Python oracle will independently map source rows, derive expected activity at actual mix, flex each budget line, calculate revenue volume, mix, and yield variances, calculate labor volume, rate, overtime, and efficiency variances, aggregate by department, and build control equations from static budget through flexible budget to actual. It will report row-keyed facts and explicit unmapped or duplicate mappings rather than rely on workbook cell positions.

## Atomic rubric design

- R001: Required sheets exist and all supplied source rows are retained.
- R002: Every GL and payroll row is mapped once or appears explicitly in Exceptions.
- R003: Variable budget lines flex to the correct actual driver while fixed lines remain fixed.
- R004: Flexible-budget department and consolidated totals are correct and formula-linked.
- R005: Revenue volume variance uses actual activity against budget volume at budget economics.
- R006: Payer-mix and yield variances are separated without double counting and bridge to actual revenue.
- R007: Labor volume, base-rate, overtime-premium, and efficiency variances use the visible standards and do not overlap.
- R008: Supply and other variable-cost variances use their stated operational drivers.
- R009: Favorable and unfavorable labels follow the visible sign convention for revenue and expense rows.
- R010: Department bridges and the consolidated bridge reconcile exactly from static budget to actual.
- R011: Material explanations are ranked by the stated threshold and retain source-driver provenance.
- P001: Penalize static-budget-only comparison, discarded unmapped GL, overlapping variance components, sign inversions, or pasted bridge totals.

## Acceptable materially different equivalent

One solution may create a long-form driver ledger and produce department summaries through structured aggregations. A materially different equivalent may construct separate revenue, labor, and supply bridges by department and combine them through a control sheet. Sheet order and helper strategy may differ; flexible-budget values, nonoverlapping components, signs, controls, and provenance must agree.

## Negative fixtures

- No-op: source data copied into the workbook with no flexible budget or variance decomposition.
- Malformed: missing bridge sheets, unreadable workbook, or text values replacing required numeric formulas.
- M1: Actual is compared only with static budget, so activity volume is mislabeled as operating performance.
- M2: Revenue variance direction is reversed while expense variance direction is left unchanged.
- M3: Payer-mix variance is calculated after applying actual yield and is counted again in yield variance.
- M4: Overtime premium is included in both rate and efficiency variances.
- M5: Unmapped GL rows are dropped from department and consolidated controls.
- M6: Labor standard hours use total visits instead of the procedure-weighted workload driver.

## Two meaningful input perturbations

- P1: Increase actual high-complexity procedures in one department by 25 while leaving rates and recorded actual cost unchanged. The flexible revenue and labor standards, revenue volume variance, labor volume and efficiency components, and department bridge must propagate; fixed budget, payer yields, wage rates, and other departments remain invariant.
- P2: Increase the overtime hourly rate for one payroll group by USD 4.00 with hours unchanged. Actual labor cost, overtime or rate variance, affected department result, consolidated bridge, and materiality ranking must change by the exact overtime hours times USD 4.00; volume, standard hours, efficiency variance, and revenue components remain invariant.

## Verifier-private CONFIRM sibling

The verifier-private CONFIRM sibling will use a different outpatient network, departments, procedure weights, payer mix, staffing groups, GL accounts, thresholds, and actual-versus-budget pattern. It preserves the flexible-budget and nonoverlapping variance-bridge capability contract. Confirm inputs, truth, oracle, perturbations, and reference remain verifier-private and outside Agent-visible input.

## Difficulty-change policy

Allowed difficulty changes: one additional legitimate activity driver, a visible fixed-versus-step-variable rule, a payer-mix category, an overtime allocation rule, or more departments with the same contract.

Forbidden difficulty changes: concealing the bridge order or sign convention, withholding mappings required for a unique result, introducing patient-identifying data, padding irrelevant GL rows, corrupting workbook interfaces, or tuning against one model.

## Activation rule

Activate only if a Track B PRIMARY is invalid beyond repair, remains VALID_BUT_EASY after three professional-complexity revisions, or reviewers require a distinct flexible-budget capability. Activation requires a newly built, independently validated package and does not inherit evidence from this design card.

## Build state

package_not_built: true. No task directory, workbook, fixture, Judge, Harbor wrapper, pass-at-k sample, Windows Excel receipt, or human acceptance exists for this reserve.
