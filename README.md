# P15 Excel Agent Benchmark: Review Release

This public development-review repository is the technical supplement for the first 15-task P15 pilot. It contains reviewable task packages, scoring logic, sanitized model-run summaries, construction guidance, and release evidence. It is not a sealed benchmark: Gold, evaluator, and CONFIRM/private-labeled materials committed here must be treated as disclosed.

## Current status

P15 contains exactly 15 primary task candidates: five Track A tasks, five Track B tasks, and five Track C tasks. All 15 pass the existing local answer-and-evaluator checks. The Pivot task uses a genuine Excel template and has been read back in Microsoft Excel for Mac with one PivotCache, one PivotTable, two SUM measures, a Q2 filter, refresh behavior, and one PivotChart. Windows Excel compatibility remains pending.

Nine additional reserve tasks are fully designed but are not packaged. They are available under [`reserves/p15_v1/`](reserves/p15_v1/).

The current n=8 development campaign targets 360 task-system attempts. At the fixed 2026-09-03 12:34:49 UTC evidence cutoff, strict score-admissible, artifact-backed coverage was 305/360. An audit removed Guard-forced and infrastructure-invalid runs from the coverage counter even when a workbook was recoverable. Forty-three completed workbooks from abnormal Trials remain available for diagnostic regrade, but they are not called clean end-to-end samples. One mid-Agent workbook and four distinct Pivot workbooks with inconsistent native-Excel receipt chains are explicitly `N/A`. Missing cells are replaced only through new attempt IDs; completed normal attempts are not rerun.

| Current system | Strict coverage | Remaining to 120 | Tasks at 8/8 | Current-contract task-balanced mean, provisional |
| --- | ---: | ---: | ---: | ---: |
| Codex CLI + GPT-5.6 sol high | 108 | 12 | 6 | 0.582 |
| Claude Code + Opus 5 | 99 | 21 | 2 | 0.228 |
| Qwen Code + Qwen3.8-max | 98 | 22 | 3 | 0.530 |

These means are diagnostic previews, not rankings. Judge v2 is not frozen or uniformly applied to the preserved workbooks. Forty-one historical baseline attempts are reused to avoid duplicate paid work, but they are not pooled with the current execution contract for formal pass@8.

The exact current snapshot is in [`results/N8_SNAPSHOT.json`](results/N8_SNAPSHOT.json), with task and system tables in [`results/N8_TASK_SUMMARY.csv`](results/N8_TASK_SUMMARY.csv) and [`results/N8_SYSTEM_SUMMARY.csv`](results/N8_SYSTEM_SUMMARY.csv). The earlier development comparison and 41-attempt group-gateway checkpoint remain preserved as separate historical evidence in [`results/`](results/README.md).

A score of 0.70 or above counts as one successful attempt. Empirical success is `c/n`. Standard pass@8 uses `1 - C(n-c, 8) / C(n, 8)`, but with exactly `n=8` it collapses to 0 when `c=0` and 1 whenever `c>=1`. The original requirement is still ambiguous between this estimator and the empirical rate over eight runs. Both the mixed execution contracts and the unresolved definition block formal pass@8 reporting.

## Development cost

| Record | Amount | Interpretation |
| --- | ---: | --- |
| Codex CLI framework estimate | USD 38.79 | Estimated by the run framework for 56 valid Codex development attempts; it is not a verified card charge. |
| OpenRouter account usage delta | USD 15.81 | Account-usage increase during 29 paid runs covering Opus 4.8, Qwen3.8-max, and one supplemental Opus 5 run. |
| Current dedicated ZCloud token | 34.508372 displayed units | Authoritative token-total readback aligned to 418 terminal runtime receipts at the fixed cutoff. Per-run deltas overlap under parallel lanes and are not summed. The release does not assert that one displayed unit equals one US dollar. |
| Earlier group-gateway token log | 834.003584 displayed units | Historical 3,944-request checkpoint; kept separate from the current dedicated-token campaign. |

These records use different accounting bases and are not added together.

## Reading paths

For the current design-and-difficulty review, read these files in order:

1. [Result analysis and Semantic Harness plan](docs/RESULT_ANALYSIS_AND_HARNESS_PLAN.md)
2. [Current sanitized result snapshot](results/README.md)
3. [Results, limitations, and the 15-task revision plan](docs/RESULTS_AND_LIMITATIONS.md)
4. [Task selection and construction logic](docs/TASK_DESIGN.md)
5. [The 15-task index](tasks/INDEX.md)

**10 minutes**

1. [Project overview](docs/PROJECT_OVERVIEW.md)
2. [Task index](tasks/INDEX.md)
3. [Current 15-task summary](results/N8_TASK_SUMMARY.csv)

**30 minutes**

1. [Results and limitations](docs/RESULTS_AND_LIMITATIONS.md)
2. [Task design](docs/TASK_DESIGN.md)
3. [Result analysis and Semantic Harness plan](docs/RESULT_ANALYSIS_AND_HARNESS_PLAN.md)
4. [Project construction and execution](docs/REUSE_GUIDE.md#project-construction)
5. [Evaluation and Judge](docs/EVALUATION_AND_JUDGE.md)
6. One representative package: [monthly ledger reconciliation](tasks/pilot_v1/P15-B-FIN-RECON-001/)

**Technical review**

1. All [15 task packages](tasks/pilot_v1/)
2. [Evaluation configuration](results/EVALUATION_CONFIG.yaml)
3. [Current task-system metrics](results/N8_TASK_SYSTEM_METRICS.csv) and [sanitized attempt records](results/N8_ATTEMPTS.csv)
4. [Human review guide](review/REVIEW_GUIDE.md)
5. [Reusable construction resources](resources/task_construction_guides/INDEX.md)
6. [Release manifest](release/MANIFEST.json) and [checksums](release/CHECKSUMS.sha256)

## Evidence boundaries

- `15/15 LOCAL_READY` means the reference, acceptable equivalent, no-op, malformed workbook, task-specific mutants, and deterministic evaluator checks meet the local contract. Pivot native objects are also verified on Microsoft Excel for Mac. This status is not Windows Excel or human validation.
- Harbor reference/no-op/malformed runs show that a package can load and its Judge can execute. They are not task-difficulty evidence.
- Judge v2 freeze and uniform artifact regrade, Windows Excel validation, completion of the corrected 360-attempt development coverage, formal pass@8, and external human review remain incomplete.
- No task is labeled `ACCEPTED_HARD`, and no human-review outcome is prefilled.
- Most instructions are benchmark-authored professional scenarios. They are not transcripts of observed customer requests.
- Gold, solver, evaluator, and CONFIRM/private-labeled assets already committed here are public development material and cannot serve as future private holdout evidence.

## Repository scope

This public review repository contains only allowlisted P15 release materials. Development worktrees, account credentials, provider responses, approvals, token fingerprints, Agent sessions, raw trajectories, and runtime directories are excluded. See [`DATA_AND_LICENSE_NOTICES.md`](DATA_AND_LICENSE_NOTICES.md) for the disclosure boundary.
