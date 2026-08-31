---
id: R15-B-GRANT-MONITORING-001
track: B
status: DESIGN_FROZEN
package_not_built: true
harbor_package_built: false
primary_eligible: false
split_group: R15-B-GRANT-MONITORING-001
---

# Sponsored-research award monitoring and indirect-cost forecast

## Professional work and user scenario

A university research administrator must prepare a sponsor-compliant award monitoring workbook from the notice of award, approved budget, general-ledger transactions, payroll distribution, open commitments, indirect-cost agreement, and project milestones. The workbook must distinguish cash and encumbrance status from allowability, calculate MTDC-based indirect costs, forecast completion, and surface action items before a sponsor report.

## Instruction draft

“Complete the award-monitoring workbook for the supplied grant. Map every transaction, payroll row, and commitment to the approved budget categories; disclose unmapped or potentially unallowable items. Apply the award period, direct-cost allowability, MTDC exclusions, indirect-cost rate, cost-share, and commitment rules in the visible policy. Calculate actual, committed, forecast, available, burn-rate, and end-of-award balances by category; show the direct-cost and indirect-cost roll-forwards; and identify budget, period, allowability, payroll, reporting, and closeout risks. Reconcile source controls and retain document, transaction, employee-token, PO, category, and period provenance.”

## Packaged-input design and authenticity boundary

- Planned inputs: award_terms.pdf, approved_budget.csv, gl_transactions.csv, payroll_distribution.csv, open_commitments.csv, milestones.csv, idc_agreement.md, monitoring_policy.md, and starting_workbook.xlsx.
- The visible terms will state award dates, sponsor categories, total direct and indirect authorization, cost-share, reporting dates, and rebudgeting thresholds.
- The policy will define allowability flags, payroll treatment, commitment status, MTDC exclusions such as equipment and qualifying subaward amounts, indirect-cost timing, forecast assumptions, and sign conventions.
- Transactions will include realistic corrections, one out-of-period row, an unallowable candidate, open and partially received POs, and effective-dated payroll distributions.
- All sponsor names, project identifiers, people tokens, documents, and values will be synthetic. The package will not represent a real award, sponsor decision, or institutional ledger.
- No input file has been generated at DESIGN_FROZEN status.

## Target workbook and professional operations

Required sheets are Terms, Budget, GL_Raw, Payroll_Raw, Commitments_Raw, Mapping, Allowability, IDC_Calculation, Forecast, Award_Summary, Risks, and Checks. The workbook must retain every source record, apply period and allowability rules, calculate the MTDC base and indirect costs without excluded costs, avoid double counting commitments and actuals, forecast payroll and other known costs through award end, and reconcile authorized budget to actual, committed, forecast, and remaining balances.

## Independent oracle design

A task-local Python oracle will independently parse the frozen award facts and policy, map source records, apply award-period and allowability decisions, calculate MTDC exclusions and indirect cost, roll commitments from open to received status, project recurring payroll and milestone costs, determine remaining balance and reporting risks, and verify all direct, indirect, and total control equations. Outputs will be keyed by immutable transaction, payroll, PO, category, and period identifiers.

## Atomic rubric design

- R001: Required sheets exist and every supplied GL, payroll, and commitment row is retained.
- R002: Source records map once to approved categories or appear explicitly as unresolved exceptions.
- R003: Award-period, transaction-status, and allowability rules are applied correctly with evidence.
- R004: Actual direct costs reconcile to the GL and payroll controls without duplicated payroll posting.
- R005: Open, partial, and closed commitments are valued correctly and are not double counted with actual receipts.
- R006: MTDC base includes eligible direct costs and excludes equipment, specified subaward amounts, and other visible exclusions.
- R007: Indirect cost is calculated at the correct rate and reconciles by period and in total.
- R008: Forecast payroll and milestone costs use the visible remaining-period assumptions.
- R009: Available balance and estimate-at-completion reconcile by category and for the award total.
- R010: Cost-share, rebudgeting, burn-rate, reporting, period, and closeout risks use supplied thresholds and dates.
- R011: Each summary and risk retains transaction, payroll, PO, category, and period provenance.
- P001: Penalize IDC on excluded bases, commitment double counting, ignored award dates, unsupported allowability assumptions, or pasted forecast totals.

## Acceptable materially different equivalent

One solution may use a unified cost-and-commitment ledger with status and MTDC attributes. A materially different equivalent may keep GL, payroll, and purchase orders in separate schedules, then feed an explicit category bridge and IDC worksheet. Layout, row order, and helper formulas may differ; allowability, MTDC, indirect cost, forecast, balance, risks, and provenance must agree.

## Negative fixtures

- No-op: supplied sources imported but monitoring schedules, IDC calculation, and risk outputs left blank.
- Malformed: non-XLSX payload, missing award-summary or IDC sheets, or invalid dates and numbers in required outputs.
- M1: Indirect cost is charged on equipment and the full subaward instead of the visible MTDC base.
- M2: A partially received PO is counted at both its original commitment and posted receipt amounts.
- M3: An out-of-period transaction is included as allowable award cost without a risk flag.
- M4: A visibly unallowable entertainment charge is included in available-budget consumption.
- M5: Actual payroll is mapped to the wrong sponsor category and the category control is silently plugged.
- M6: Forecast stops at the report date and omits recurring payroll through the award end date.

## Two meaningful input perturbations

- P1: Add one in-period, allowable monthly payroll distribution of USD 8,000 to a salary category that is inside the MTDC base. Actual direct cost increases by USD 8,000, indirect cost increases by the visible rate times USD 8,000, remaining balance and EAC change by the combined amount, and related risks propagate; equipment exclusion, cost-share, award authorization, and unrelated categories remain invariant.
- P2: Change one equipment PO from fully open to fully received on the close date while holding its total amount constant. Committed equipment decreases and actual equipment increases by the same amount; equipment EAC, total EAC, MTDC base, indirect cost, and overall remaining balance remain invariant, while status and closeout risks update.

## Verifier-private CONFIRM sibling

The verifier-private CONFIRM sibling will use a different sponsor type, award dates, budget categories, MTDC exclusion mix, indirect-cost rate, payroll cadence, commitments, reporting calendar, and risk pattern. It preserves the allowability, commitment, MTDC, indirect-cost, forecast, and control contract. Confirm inputs, truth, oracle, perturbations, and reference remain outside Agent-visible input.

## Difficulty-change policy

Allowed difficulty changes: a disclosed rate change by period, one qualifying subaward threshold, one cost-share category, more reporting milestones, or a legitimate partial-receipt sequence.

Forbidden difficulty changes: withholding award terms, inventing ambiguous allowability decisions, using real employee or sponsor-sensitive data, flooding the ledger with noise, breaking expected workbook access, or exploiting a target model quirk.

## Activation rule

Activate only when a Track B PRIMARY cannot meet validity gates, remains VALID_BUT_EASY after three genuine professional revisions, or senior review prefers sponsored-research monitoring as a more distinct capability. Activation begins a new package construction and validation cycle; this frozen card supplies no benchmark or difficulty evidence by itself.

## Build state

package_not_built: true. No task directory, workbook, fixture, Judge, Harbor wrapper, model run, external Excel evidence, or reviewer decision exists for this reserve.
