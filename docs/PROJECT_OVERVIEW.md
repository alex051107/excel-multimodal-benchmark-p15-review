# P15 Project Overview

## Research objective

P15 evaluates whether an Excel Agent can deliver a professional workbook that another person can continue to edit, recalculate, and audit. A correct final number is insufficient when formulas are hard-coded, dependencies are broken, source records are dropped, or workbook-native objects are missing.

P15 is the first 15-task pilot used to test both this research question and the task-production method. It is organized by work type rather than industry.

| Track | Work covered | Five task forms |
| --- | --- | --- |
| A: Professional models | Build, calculate, or repair a recalculable professional workbook | Model construction; model repair; constrained engineering decision; statistical analysis; policy scenario |
| B: Multi-file analysis | Select, clean, join, reconcile, and report from several files | Source/version selection; cleaning and joining; ledger reconciliation; native Pivot; evidence-based report |
| C: Documents to workbook | Convert PDF, image, or scan inputs into an editable and traceable workbook | Multipage invoice; quote scope; receipt batch; document amendment; continuation statement |

## Current release state

- 15 primary candidates: five per track.
- 9 design-only reserves: three per track.
- All 15 tasks pass the current local answer-and-evaluator checks.
- The native Pivot task has been created and read back in Microsoft Excel for Mac; Windows Excel compatibility remains pending.
- Three target systems have produced real development answers.
- A group-gateway n=8 batch has 41 valid scored attempts and is preserved for resumption after the token-level quota is restored.
- Standard pass@8, Windows Excel validation, and external human review are not complete.

The per-task status is recorded in [`results/TASK_STATUS.csv`](../results/TASK_STATUS.csv). The task descriptions and direct package links are in [`tasks/INDEX.md`](../tasks/INDEX.md).

## Task package contents

Each primary package includes:

- a professional instruction;
- all Agent-visible input files;
- a reference workbook;
- an oracle that recomputes correctness independently from the reference workbook;
- an atomic rubric with at least six scoring criteria;
- a deterministic Judge;
- task metadata and a Harbor package;
- an acceptable equivalent, no-op, malformed workbook, and task-specific semantic mutants;
- local, CONFIRM-reference, and Harbor smoke receipts.

The reference workbook demonstrates one complete delivery. The oracle establishes what is correct. The rubric states what will be checked. The Judge opens a candidate workbook and performs those checks. Keeping these components separate reduces the risk that scoring merely reproduces the reference workbook's layout.

## Development result

The earlier comparison has four Codex attempts per task and one Claude and Qwen attempt per task. A separate group-gateway batch uses GPT-5.6 sol, Opus 5, and Qwen3.8-max, but stopped before any task-system pair reached eight valid samples. These results can identify obvious easy cases, evaluator problems, and cross-system failure signals. They do not support stable pass@8 or a system ranking.

Three tasks currently appear too easy: DCF, financial-model repair, and pump selection. Several Track B and Track C tasks show low cross-system scores. Low scores require human review before they are interpreted as professional difficulty, because an ambiguous instruction or a false-negative Judge can produce the same pattern.

## Decisions required before formal expansion

1. Confirmation of the A/B/C coverage for the intended forms of Excel work.
2. Confirmation that benchmark-authored professional scenarios satisfy the project's “real query” requirement.
3. Review of the Golden/Oracle/Rubric/Judge boundaries.
4. Selection of low-score tasks for human review before additional samples are purchased.
5. Selection of construction methods for reuse in the next batch.
6. Confirmation that the group gateway can keep the Opus 5 route, upstream provider, and token budget fixed for the resumed batch.

## Where to inspect the evidence

- Task portfolio: [`tasks/INDEX.md`](../tasks/INDEX.md)
- Task design: [`docs/TASK_DESIGN.md`](TASK_DESIGN.md)
- Evaluation: [`docs/EVALUATION_AND_JUDGE.md`](EVALUATION_AND_JUDGE.md)
- Results: [`docs/RESULTS_AND_LIMITATIONS.md`](RESULTS_AND_LIMITATIONS.md)
- Construction resources: [`resources/task_construction_guides/`](../resources/task_construction_guides/)
- Human review: [`review/`](../review/)
