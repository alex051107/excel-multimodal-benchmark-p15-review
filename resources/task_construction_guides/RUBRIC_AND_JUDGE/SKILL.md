---
name: excel-benchmark-rubric-judge
description: Design and audit atomic deterministic rubrics, Judges, fixture gates, and DEV-CONFIRM contracts for P15-style Excel benchmark tasks; use when scoring task semantics, not when grading by visual similarity alone.
---

> Validated against P15 checkpoint: `8df20f98366633a4813cf6b98ee7e78270e3bdb7`
> Last evidence refresh: `2026-08-30`

# Rubric and Judge

## Purpose

Turn each professional requirement into an independently testable scoring criterion. Make critical semantic failures fall below the pass threshold, while accepting materially different correct workbook implementations.

## Required inputs

- The frozen instruction, source bundle, workbook contract, and independently recomputed oracle.
- At least six atomic rubric criteria with explicit weights and a criterion-to-check coverage map.
- Protected-source ranges, source perturbations, required formulas or native objects, and tolerance rules.
- A reference, a materially different acceptable equivalent, a no-op, a malformed file, and at least four task-specific semantic mutants.
- Distinct DEV and CONFIRM contracts stored only in verifier-private test assets.

## Exact procedure

1. Split bundled rubric prose into atomic criteria: source selection, record inclusion, formula lineage, transformation semantics, closure, protected sources, and presentation bindings should be independently observable.
2. Implement every high-weight criterion in the Judge and map each criterion identifier to its concrete check. Do not award points for an unimplemented requirement.
3. Recompute expected values from verifier-private source snapshots. For dynamic checks, perturb the actual source cells and evaluate the affected formulas, records, chart bindings, or native objects.
4. Use tolerances that match the professional rule. Enforce final reconciliation and control-total closure when applicable.
5. Apply a justified critical cap or dependency only when failure of a required source, join, reconciliation, dynamic formula, or chart binding makes the deliverable professionally unusable. The resulting score must be below `0.70`.
6. Return structured JSON with normalized score, per-criterion scores, evidence, failure codes, split, and validation status. Malformed workbooks must fail closed rather than raise an unclassified exception.
7. Support DEV by default and CONFIRM only through explicit `--split confirm` or `P15_EVAL_SPLIT=confirm`. Resolve both repository-local paths and Harbor's `/tests` runtime without reading task-root verifier assets.
8. Run the reference, equivalent, no-op, malformed, and semantic-mutant fixture categories. Run every deterministic fixture five times on the same version and require identical results.
9. Freeze Judge, DEV/CONFIRM contracts, fixture workbooks, task version, and model configuration before target-agent attempts.

## Required outputs

- Atomic rubric criteria with weights, acceptance conditions, and implemented check identifiers.
- A deterministic task-local evaluator and verifier-private DEV/CONFIRM oracle/source snapshots.
- Reference, materially different equivalent, no-op, malformed, and at least four task-specific semantic-mutant fixtures.
- Structured result JSON and a receipt containing five-run scores, standard deviations, failure codes, and fixture-gate status.

## Validation checks

- Reference score equals `1.0`; acceptable equivalent is at least `0.95`.
- No-op is below `0.30`; malformed is exactly `0` with a structured failure.
- Every task-specific semantic mutant is below `0.70`, and the mutants do not all receive the same score.
- The same deterministic workbook produces the same score in five runs. For any LLM/VLM criterion, measured standard deviation must be below `0.05`.
- A critical dynamic/source/join/reconciliation/chart failure cannot retain a normalized score of `0.70` or above.
- Protected-source edits are detected; declared perturbations alter the expected downstream result.
- The equivalent changes material implementation choices rather than rewriting one `SUM` cell.
- Harbor verifier assets are self-contained under `tests/`; the Agent image cannot see oracle, reference, mutant, or CONFIRM truth.

## Allowed difficulty changes

- Add atomic checks for genuine professional semantics that the original rubric omitted.
- Strengthen lineage, closure, tolerance, protected-source, chart-binding, or native-object validation.
- Add a materially different correct equivalent and semantic mutants representing plausible professional errors.
- Use a critical cap whose trigger corresponds to an unusable deliverable and whose evidence is reported explicitly.

## Forbidden difficulty changes

- Score shaping solely to manufacture mutant diversity or a desired difficulty label.
- One monolithic rubric item that bundles several requirements while the Judge checks only one.
- Visual similarity, filename, worksheet order, formula-string identity, or OOXML substring matching as the sole correctness test.
- Hidden requirements, exact-format traps, output-side perturbations, or model-specific failure triggers.
- Verifier truth, CONFIRM sources, fixture workbooks, or oracle logic in Agent-visible inputs.
- A generic Judge platform, runner, registry, learned verifier, or trajectory-diagnosis system.

## Common failure modes

- High-weight rubric statements have no executable check, creating normalized-score holes.
- The aggregate total is correct despite a wrong source, dropped exception, wrong join, wrong FX date, omitted adjustment, or stale chart.
- All mutants collapse to zero because criteria are bundled or a global failure erases useful semantic coverage.
- An equivalent passes only because it changes one formula without changing the implementation approach.
- Local evaluation succeeds by reaching task-root files that are absent when Harbor builds the verifier from `tests/`.
- A malformed workbook crashes before writing a zero reward and structured failure.
- A Pivot Judge accepts labels or formulas without object-level PivotCache/PivotTable/chart relationships.

## Stop conditions

- Stop when all fixture gates pass on the frozen Judge and repeated deterministic scores are stable.
- If a critical requirement lacks a reproducible check, keep the task out of release until it is implemented or explicitly invalidated.
- If native Excel evidence is required but unavailable, retain `TASK_INVALID` rather than inventing a reference score.
- Rerun only the fixture category affected by a repair; run the final focused set once after the batch.

## Examples from verified P15 work

- Sales mutants scored differently: dropped coverage `0.08`, hardcoded revenue `0.36`, registry collateral edit `0.48`, and wrong source `0.0`.
- Ops mutants exposed separate semantics: dropped exception `0.269231`, included duplicate `0.230769`, hardcoded extended cost `0.384615`, and wrong product join `0.038462`.
- Fin mutants avoided an all-zero result: dropped unmatched `0.50`, omitted adjustment `0.038462`, wrong FX date `0.192308`, and wrong transaction match `0.461538`.
- Health mutants distinguished hardcoded claim `0.333333`, stale chart `0.296296`, wrong geography `0.296296`, and wrong time window `0.185185`.
- For Sales, Ops, Fin, and Health, reference and equivalent scored `1.0`, no-op and malformed scored `0`, and five repeated runs were identical. Pivot remained `TASK_INVALID` because the required native objects were absent.
