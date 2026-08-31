---
name: excel-benchmark-oracle-and-gold
description: Build and audit independent Excel benchmark oracles, formula-linked gold workbooks, equivalent solutions, private splits, fixtures, and verifier runtime boundaries.
---

> Validated against P15 checkpoint: `8df20f98366633a4813cf6b98ee7e78270e3bdb7`
> Last evidence refresh: `2026-08-30`

# Oracle and Gold Construction

## Purpose

Create truth and evaluation assets that prove workbook semantics without reading answers back from the reference workbook or leaking private truth. Separate deterministic semantic evidence, Harbor packaging evidence, Windows Excel evidence, target-agent difficulty, and human review.

This skill is verifier-side only. Never include it, oracle code/values, fixture deltas, CONFIRM assets, or Judge criteria in Agent-visible inputs.

## Required inputs

- The frozen professional task contract, source/input files, units, formulas, tolerances, protected cells, and critical-failure policy.
- The DEV input instance and a verifier-private CONFIRM sibling plan with a shared capability invariant but different values/source instance.
- The required workbook layout and any chart, Pivot, external-link, or recalculation requirement.
- The task-local Harbor layout, including the separate Agent environment and `tests/` verifier Docker context.

## Exact procedure

1. Implement an independent oracle from source facts and declared rules. It must not open or derive truth from `solution/reference.xlsx`, an equivalent, a candidate, or cached workbook results.
2. Return every critical semantic component, not only the final answer: intermediate periods/rows, bridges, constraints, selected item, checks, chart source values, and perturbation results.
3. Build the gold workbook with `@oai/artifact-tool`; keep critical outputs formula-linked to visible inputs. Inspect formulas/values, render relevant sheets, and scan formula errors before export.
4. Build a materially different acceptable equivalent. Change a real implementation choice—for example a helper layout, algebraically distinct formula family, or equally valid summary organization—while preserving the same professional semantics. Record the material delta; a byte-identical or identity-copy workbook is not an equivalent.
5. Build a no-op, a malformed file, and at least four task-specific semantic mutants. Each mutant represents a different plausible professional error and changes only the intended failure.
6. Map 8–12 atomic rubric criteria to separate evaluator checks. Apply a critical cap/veto only to failures that truly invalidate the professional result or protected-source integrity.
7. Create DEV and CONFIRM contracts. CONFIRM must use a distinct source/professional instance and changed drivers, preserve the capability invariant, and stay under `tests/confirm`; do not tune it after observing model failures.
8. Make split selection explicit with `--split dev|confirm`. Default local/release smoke to DEV unless a CONFIRM invocation is explicitly requested and recorded.
9. Honor Harbor 0.22 verifier isolation: `tests/Dockerfile` builds with `tests/` as its context and copies it to `/tests`. Put every verifier-private runtime dependency needed in the container under `tests/`, and make evaluator path resolution work both from the task checkout and from `/tests`. Do not solve this by copying verifier assets into the Agent environment.
10. Run reference, equivalent, no-op, malformed, and mutants repeatedly. Record score, pass, failure codes, criterion scores, split, task ID, repetition count, and variability.
11. Require real Microsoft Excel for canonical open/save/recalculate, async-query barriers, formula-error/external-link/protection checks, and native Pivot readback. Artifact-tool, Open XML inspection, LibreOffice, or Harbor cannot substitute for that receipt.

## Required outputs

- Independent DEV oracle and private contract with atomic criteria, formula regions, two exact perturbations, and protected invariants.
- Formula-linked reference workbook plus a documented materially different acceptable equivalent.
- No-op, malformed, and at least four task-specific semantic mutants.
- `tests/confirm/contract.json`, distinct input overlay, independent oracle, reference workbook, and explicit evaluator invocation.
- A tests-context-complete verifier image and a separate Agent environment containing no truth assets.
- Deterministic validation receipt and separately labeled Windows, Harbor, target-agent, and human evidence states.

## Validation checks

- Oracle source code never reads reference/equivalent/candidate workbooks and can recompute truth from declared inputs.
- Reference `=1.0`; materially different equivalent `>=0.95`; no-op `<0.30`; malformed `=0`; every mutant `<0.70`; semantic-mutant scores are not all the same.
- Five repeated deterministic runs produce identical scores/signatures. Any LLM/VLM criterion has score standard deviation `<0.05`.
- Two meaningful DEV perturbations use exact expected propagation and protected invariants; changed-only checks fail review.
- DEV and CONFIRM contracts share only the capability invariant, not values or hidden truth.
- Repository-local and `/tests` evaluator invocations resolve the same split-specific verifier assets.
- `EXCEL_VALIDATED`, Harbor pass, target-agent hardness, and human acceptance are each backed by their own receipt; none is inferred from another.

## Allowed difficulty changes

- Add independently computable professional dependencies, constraints, cross-sheet bridges, sensitivity coordinates, or decision checks.
- Add a second legitimate source or revision when its precedence and join keys are supplied.
- Refine rubric atomicity or a critical cap when a real mutant shows that distinct professional failures are being conflated.
- Replace a persistently easy task with its designated reserve after at most three legitimate revisions.

## Forbidden difficulty changes

- Reading truth from the reference workbook, comparing only to workbook bytes, or using the gold as the oracle.
- Identity-copy equivalents, all-zero/all-equal mutant scoring, changed-only perturbation checks, or one broad criterion standing in for several high-weight claims.
- Leaking DEV/CONFIRM truth into instructions, task inputs, environment layers, logs visible to the Agent, or this skill into the package.
- Modifying CONFIRM in response to target-model failures.
- Calling artifact-tool/LibreOffice/Open XML evaluation Windows Excel proof, or calling Harbor smoke a difficulty test.
- Adding generic registries, routers, generators, learned verifiers, or provenance platforms for a task-local contract.

## Common failure modes

- The oracle imports cached values from the reference, making reference and Judge fail together.
- Only the headline answer is checked, so broken intermediate logic receives full credit.
- The acceptable equivalent is just a copied reference and therefore does not test semantic tolerance.
- A broad critical veto sends every mutant to the same score instead of preserving atomic diagnostic coverage.
- The evaluator works in the repository but fails in Harbor because `/tests/evaluate.py` resolves task-root metadata outside the Docker context.
- CONFIRM is merely DEV with renamed labels or is copied into `environment/input`.
- A formula-error scan failure is serialized as an empty list and incorrectly permits `EXCEL_VALIDATED`.

## Stop conditions

- Stop if two independent implementations cannot reproduce the intended truth from supplied facts and rules.
- Stop if a material equivalent cannot be stated without copying the reference implementation.
- Stop and fail closed if the verifier image cannot resolve its private contract/oracle in both supported locations.
- Stop at the relevant pending status when Windows Excel, Harbor service, model credentials/quota, or human review is unavailable.
- Stop if closing a gate would require leaking truth or manufacturing difficulty.

## Current verified examples

- The inspected Track A private contracts use task-specific oracle entry points, protected inputs, formula-cell sets, and two DEV perturbations. DCF recomputes growth/WACC propagation; ENG recomputes hydraulic selection; DEBUG recomputes downstream model ties; STAT recomputes paired inference; EIA recomputes scenario balance and emissions.
- Track A local receipts recorded five stable deterministic repetitions with zero score variance. DCF, for example, scored reference/equivalent `1.0`, no-op/malformed `0`, and distinct semantic mutants from `0.0` to `0.592593`.
- The five CONFIRM contracts use different source instances and changed drivers with `dev_values_reused=false`; each explicitly invokes `tests/evaluate.py ... --split confirm`.
- The inspected verifier Dockerfiles use `COPY . /tests`, demonstrating why task-root metadata cannot be assumed available in the separate verifier image.
- [`results/ATTEMPTS.csv`](../../../results/ATTEMPTS.csv) records 85 real Agent attempts: 56 Codex runs, 14 Claude Code + Opus 4.8 runs, 14 Qwen Code + Qwen3.8-max runs, and one supplemental Opus 5 run. These runs do not replace Windows or human evidence.

## Historical status note

- Task-local validation receipts were created before Agent screening. Their older `agent_screen=NOT_STARTED` field is superseded by the release-level attempt ledger and must not be used as current model-run status.
