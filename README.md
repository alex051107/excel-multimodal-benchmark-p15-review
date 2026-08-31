# P15 Excel Agent Benchmark: Review Release

This repository is the fixed technical supplement to the project Feishu page. Feishu carries the project narrative, weekly updates, open decisions, and next steps. This repository contains the reviewable task packages, scoring logic, model-run summaries, construction guidance, and release evidence for the first 15-task pilot.

## Current status

P15 contains exactly 15 primary task candidates: five Track A tasks, five Track B tasks, and five Track C tasks. Fourteen tasks pass the current local answer-and-evaluator checks. `P15-B-PUBLIC-PIVOT-001` remains invalid until a genuine PivotCache, PivotTable, filter, SUM measures, refresh behavior, and PivotChart are created and read back in Microsoft Excel on Windows.

Nine additional reserve tasks are fully designed but are not packaged. They are available under [`reserves/p15_v1/`](reserves/p15_v1/).

The current Agent results are development evidence:

| System | Valid attempts | Passes at score >= 0.70 | Mean score | Current per-task sample size |
| --- | ---: | ---: | ---: | ---: |
| Codex CLI + GPT-5.6 sol | 56 | 26 | 0.634 | 4 |
| Claude Code + Opus 4.8 | 14 | 8 | 0.678 | 1 |
| Qwen Code + Qwen3.8-max | 14 | 3 | 0.488 | 1 |

One additional Claude Code + Opus 5 run was completed on the DCF task and scored 1.000. It is supplemental and is not pooled with the three-system comparison.

A score of 0.70 or above counts as one successful attempt. The empirical pass@1 estimate for a system with `n` independent attempts and `c` successes is `c / n`. Standard pass@8 uses `1 - C(n-c, 8) / C(n, 8)` and is not reported when `n < 8`. The current results therefore support task-development decisions, not stable system rankings or accepted-hard claims.

## Development cost

| Record | Amount | Interpretation |
| --- | ---: | --- |
| Codex CLI framework estimate | USD 38.79 | Estimated by the run framework for 56 valid Codex development attempts; it is not a verified card charge. |
| OpenRouter account usage delta | USD 15.81 | Account-usage increase during 29 paid runs covering Opus 4.8, Qwen3.8-max, and one supplemental Opus 5 run. |

The two amounts use different accounting bases and are not added together because invoice-level deduplication evidence is unavailable.

## Reading paths

For the current design-and-difficulty review, read these four files in order:

1. [Results, limitations, and the 15-task revision plan](docs/RESULTS_AND_LIMITATIONS.md)
2. [Task selection and construction logic](docs/TASK_DESIGN.md)
3. [The 15-task index](tasks/INDEX.md)
4. [Reusable construction and execution resources](docs/REUSE_GUIDE.md)

**10 minutes**

1. [Project overview](docs/PROJECT_OVERVIEW.md)
2. [Task index](tasks/INDEX.md)
3. [Model summary](results/MODEL_SUMMARY.csv)

**30 minutes**

1. [Results and limitations](docs/RESULTS_AND_LIMITATIONS.md)
2. [Task design](docs/TASK_DESIGN.md)
3. [Project construction and execution](docs/REUSE_GUIDE.md#project-construction)
4. [Evaluation and Judge](docs/EVALUATION_AND_JUDGE.md)
5. One representative package: [monthly ledger reconciliation](tasks/pilot_v1/P15-B-FIN-RECON-001/)

**Technical review**

1. All [15 task packages](tasks/pilot_v1/)
2. [Evaluation configuration](results/EVALUATION_CONFIG.yaml)
3. [Aggregated run results](results/RUN_RESULTS.csv) and [attempt-level records](results/ATTEMPTS.csv)
4. [Human review guide](review/REVIEW_GUIDE.md)
5. [Reusable construction resources](resources/task_construction_guides/INDEX.md)
6. [Release manifest](release/MANIFEST.json) and [checksums](release/CHECKSUMS.sha256)

## Evidence boundaries

- `14/15 LOCAL_READY` means the reference, acceptable equivalent, no-op, malformed workbook, task-specific mutants, and deterministic evaluator checks meet the local contract. It is not Windows Excel or human validation.
- Harbor reference/no-op/malformed runs show that a package can load and its Judge can execute. They are not task-difficulty evidence.
- Windows Excel validation, held-out Agent runs on CONFIRM siblings, formal per-task pass@8, and external human review remain incomplete.
- No task is labeled `ACCEPTED_HARD`, and no human-review outcome is prefilled.
- Most instructions are benchmark-authored professional scenarios. They are not transcripts of observed customer requests.

## Repository scope

This review repository contains only the P15 release materials. Development branches, internal discussions, intermediate reports, account credentials, provider responses, and unrelated historical packages are not included.
