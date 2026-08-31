---
name: excel-benchmark-task-selection
description: Select or revise a fixed Excel benchmark portfolio when exact track quotas, professional diversity, source feasibility, and evidence-bounded status must be audited before package construction.
---

> Validated against P15 checkpoint: `8df20f98366633a4813cf6b98ee7e78270e3bdb7`
> Last evidence refresh: `2026-08-30`

# Excel Benchmark Task Selection

## Purpose

Select a coherent, non-template benchmark portfolio before building packages. For P15, freeze exactly 15 PRIMARY tasks—five each in Tracks A, B, and C—while keeping nine RESERVE designs outside the packaged primary set. Selection is a design decision, not evidence of validity, difficulty, Windows Excel success, Harbor success, or human acceptance.

Keep this skill repository-side. Never copy it, its examples, private selection rationale, or reserve mapping into task `data/input_files`, `environment/input`, or other Agent-visible inputs.

## Required inputs

- The current requirements, authoritative portfolio manifest, and track definitions.
- Candidate task briefs with professional user, decision, inputs, workbook deliverable, source/license status, and native-Excel needs.
- Current task-local status and evidence, separated into designed, locally checked, externally checked, and blocked.
- The nine reserve design cards and the replacement relationship, if any.

## Exact procedure

1. Read the authoritative manifest and count PRIMARY candidates by track. Stop unless the proposed result is exactly `A=5`, `B=5`, `C=5`, `total=15`.
2. Describe each candidate as a real professional job: who needs the workbook, what decision it supports, what source files are supplied, and what must remain editable or formula-linked.
3. Compare candidates on profession, source/formats, transformation, decision output, Excel capability, and failure semantics. Reject title/role/number swaps that preserve the same underlying work.
4. Admit a candidate only when required information can be supplied without hidden conventions, its truth can be independently reconstructed, and its evaluator can distinguish semantic failure from file failure.
5. Keep all nine RESERVE tasks as complete design cards, not bulk-generated Harbor packages. Identify a reserve only as a potential replacement until the replacement decision is recorded.
6. Record the narrowest honest status. A candidate may be selected while still pending Windows Excel, Harbor runtime, target-agent, or human evidence.
7. Recheck that selection material and verifier-private truth are absent from Agent-visible inputs.

## Required outputs

- One manifest containing exactly 15 unique PRIMARY task IDs and five tasks per track.
- A compact diversity matrix covering professional job, inputs, deliverable, capability, source family, oracle feasibility, and current blocker.
- Nine separate RESERVE design cards with no unnecessary packages.
- Explicit rejection/replacement notes for candidates that are duplicative, invalid, infeasible, or still externally blocked.

## Validation checks

- Exact counts: `15 PRIMARY`, `A=5`, `B=5`, `C=5`; exactly nine RESERVE designs remain outside the primary package set.
- Every PRIMARY instruction names a materially different professional job and decision.
- No pair differs only by title, persona, labels, dates, or numeric constants.
- Every selected candidate has a plausible independent oracle and at least one materially different acceptable-equivalent strategy.
- Status fields do not promote local artifacts, smoke tests, or plans into Windows, human, difficulty, or acceptance evidence.
- No selection skill, reserve truth, oracle value, or reviewer note appears in Agent-visible inputs.

## Allowed difficulty changes

- Replace an invalid or persistently easy task with its designated reserve while preserving the exact track quota.
- Add disclosed, professionally natural dependencies, constraints, revisions, or decision checks that belong to the same job.
- Increase diversity by selecting a different real workflow, source family, or workbook capability.

## Forbidden difficulty changes

- Renaming the user, company, sheet, or task while retaining the same semantic work.
- Inflating row/page count, adding irrelevant sheets, hiding necessary facts, corrupting interfaces, or throttling time/tokens.
- Choosing a task because one target model has a known quirk.
- Treating candidate count as an acceptance quota or calling an unverified task hard.
- Packaging reserves merely to increase artifact count.

## Common failure modes

- A visually different workbook masks a duplicate computation and decision.
- A source-rich task lacks a unique or independently reconstructable answer.
- The portfolio is balanced by track label but concentrated in one profession or transformation.
- A selected native-Excel task is reported as ready before its required Excel objects are created and read back.
- RESERVE designs are silently promoted, or PRIMARY tasks silently removed, so the manifest no longer has exact counts.

## Stop conditions

- Stop selection if exact quotas can only be met with a duplicate, under-specified, unlicensed, or oracle-infeasible task.
- Stop replacement after three legitimate professional-complexity revisions; record `VALID_BUT_EASY` or use the designated reserve.
- Stop and record an external blocker when only Windows Excel, Harbor service, model credentials/quota, or human review can close the gate.
- Stop if the next step would expose private truth to the target Agent.

## Examples from verified P15 work

- `manifests/pilot_15_v1.yaml` freezes 15 candidates with five per track and explicitly states that the target count is not an acceptance quota.
- Track A spans five different jobs: acquisition DCF, hidden-root-cause FP&A repair, cooling-water pump sizing, paired crossover inference, and an electricity-policy scenario model.
- `P15-B-PUBLIC-PIVOT-001` remained Windows-native gated in the inspected manifest. Its presence in the 15-task selection did not establish PivotCache/PivotTable/PivotChart validation.
- The nine `reserves/p15_v1/*.md` cards record allowed and forbidden professional changes without generating replacement packages.
