---
name: track-b-construction
description: Construct or revise P15-style Track B Excel benchmark tasks for source selection, cleaning and joins, reconciliation, native PivotTables, and traceable analytical reporting; use for benchmark authoring, not for solving agent-visible workbooks.
---

> Validated against P15 checkpoint: `8df20f98366633a4813cf6b98ee7e78270e3bdb7`
> Last evidence refresh: `2026-08-30`

# Track B Construction

## Purpose

Build a distinct, realistic Excel work product whose correctness can be derived from supplied professional records. Preserve source lineage, make requested outputs respond to source changes, and keep verifier-private truth outside the Agent-visible package.

## Required inputs

- A concrete professional decision and user role.
- The minimum source records, reference tables, exception evidence, and business rules needed to complete it.
- An independently computed DEV oracle and a planned materially distinct CONFIRM source instance.
- The intended workbook interface, protected source areas, declared perturbations, and available Excel runtime.
- The task-local instruction, metadata, rubric, Judge, and packaging conventions already frozen for the release.

## Exact procedure

1. Choose one real workflow, not a renamed copy: approved-release selection, normalized cleaning and joins, transaction reconciliation, native public-program Pivot reporting, or traceable health reporting.
2. Put only necessary source files in the Agent-visible input bundle and synchronize that bundle with the environment input. Remove stale or unrelated visible files.
3. Preserve supplied source sheets and ranges. Put transformations, exception handling, calculations, and presentation in separate output areas.
4. Recompute truth independently from record-level sources. Do not derive the oracle from the reference workbook.
5. Link counts, coverage, exceptions, adjustments, reconciliations, claims, and charts to source records or derived formulas. Require final closure where the workflow has a control total.
6. Declare perturbations on actual source cells used by the workbook and Judge. Test the resulting downstream value, not merely whether a formula exists.
7. Create a verifier-private CONFIRM sibling with different source values and changed drivers but the same capability invariant. Freeze DEV, CONFIRM, Judge, and model configuration before difficulty testing. Never copy CONFIRM truth or reference files into Agent-visible inputs.
8. For native Pivot work, require a Windows Microsoft Excel build and readback of the source Table, PivotCache, PivotTable fields, filter, SUM measures, refresh configuration, and PivotChart. Keep the task invalid until that native artifact exists.
9. Produce only the task-local assets needed by the benchmark. Do not add a generic platform, runner, registry, generator, or dashboard.
10. Record observed validation state and external blockers without promoting a design, starter workbook, or local formula summary into delivery evidence.

## Required outputs

- A non-template instruction, synchronized Agent-visible input bundle, Golden Solution, task metadata, and standard Harbor task package.
- An independent DEV oracle, protected-source contract, perturbation contract, and task-local evaluator fixtures.
- A private CONFIRM contract, distinct input overlay, independent truth/oracle, and reference workbook when the required runtime exists.
- A validation receipt that separates local semantic evidence, native Excel evidence, Harbor packaging evidence, target-agent evidence, and external blockers.

## Validation checks

- Confirm every visible file is required and environment input matches the declared bundle.
- Replay the oracle from source records and compare record counts, totals, tolerances, and closures.
- Apply each declared source perturbation and verify the expected downstream change.
- Reject source edits outside allowed cells and reject hardcoded counts, claims, and totals that do not respond to source changes.
- Inspect chart category and value-series bindings, not just chart existence.
- For Pivot tasks, inspect native objects and exact field identities/indexes; a formula summary or matching text is insufficient.
- Verify CONFIRM has distinct values and drivers and remains under `tests/confirm`, outside Agent-visible inputs.

## Allowed difficulty changes

- Add genuine release-selection ambiguity that must be resolved from authoritative metadata.
- Add professional normalization, deduplication, exception retention, multi-key joins, or record-level traceability.
- Add transaction-date rules, documented adjustments, tolerances, unmatched-item handling, and final reconciliation closure.
- Add meaningful period comparisons, formula-linked narrative claims, or correctly bound charts.
- Require native Excel object behavior when that behavior is the professional deliverable and the required runtime can validate it.

## Forbidden difficulty changes

- Irrelevant files, extreme noise, missing necessary information, broken interfaces, hidden assumptions, token or time throttling, or traps tailored to one model.
- Cosmetic renaming or number substitution presented as a new task.
- Hardcoded outputs, output-cell perturbations, or a reference-derived oracle.
- A formula summary, static chart, or OOXML substring match used as a substitute for required native Pivot behavior.
- Any generic platform, runner, registry, provenance system, or learned verifier.

## Common failure modes

- Stale inputs remain visible or task metadata and environment input disagree.
- Counts and closures are typed values rather than formulas linked to records.
- A perturbation changes a Judge-only cell instead of a source cell used by the solution.
- The main total passes while a dropped exception, wrong join, wrong FX date, or stale chart remains undetected.
- A chart exists but points to the wrong categories or hardcoded values.
- Pivot XML contains plausible labels but no genuine PivotCache, PivotTable, selected filter, SUM measures, refresh behavior, or pivot-bound chart.
- CONFIRM reuses DEV values or leaks verifier truth into the Agent-visible package.

## Stop conditions

- Stop construction when task-local assets and focused checks are complete and only a real external dependency remains.
- Keep a native Pivot task `TASK_INVALID` / `PENDING_EXTERNAL_WINDOWS_EXCEL` until Windows Excel builds, reopens, recalculates, and validates the workbook.
- Stop after three professional-complexity revision rounds; replace with a reserve or label `VALID_BUT_EASY` if it remains easy.
- Never fill an evidence gap with a document, wrapper, synthetic receipt, or unsupported claim.

## Examples from verified P15 work

- `P15-B-SALES-DISCOVERY-001` required selecting the approved corrected 2024Q2 production release, preserving its registry, and linking six-record coverage and revenue to source data.
- `P15-B-OPS-CLEAN-JOIN-001` normalized orders, removed one exact normalized duplicate, retained exceptions, joined valid product/location records, and linked counts and closure to records.
- `P15-B-FIN-RECON-001` used transaction-date FX, a documented adjustment, a one-dollar tolerance, unmatched-item handling, and final reconciliation closure.
- `P15-B-HEALTH-REPORT-001` compared 2012–2014 with 2015–2017, linked claims to formulas, and verified exact chart categories and series.
- `P15-B-PUBLIC-PIVOT-001` specifies `ProgramEventsTable`, `ProgramDeliveryPivot`, Region rows, Program columns, the `2024Q2` filter, SUM Participants/Spend, refresh, and a pivot-bound clustered chart. It remains invalid locally because no Windows-native workbook has been produced.
