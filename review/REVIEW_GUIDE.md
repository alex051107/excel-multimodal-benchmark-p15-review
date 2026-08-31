# Review Guide

## Ten-minute review

1. Read the [project overview](../docs/PROJECT_OVERVIEW.md).
2. Scan the [15-task index](../tasks/INDEX.md).
3. Check the [model summary](../results/MODEL_SUMMARY.csv).
4. Confirm the external evidence still listed as pending.

## Thirty-minute review

1. Read the [task-design method](../docs/TASK_DESIGN.md).
2. Read the [evaluation and Judge contract](../docs/EVALUATION_AND_JUDGE.md).
3. Inspect one representative task from each track.
4. Compare the development signals with the proposed next action in [`TASK_STATUS.csv`](../results/TASK_STATUS.csv).

Suggested representatives:

- Track A: [`P15-A-FIN-DCF-001`](../tasks/pilot_v1/P15-A-FIN-DCF-001/)
- Track B: [`P15-B-FIN-RECON-001`](../tasks/pilot_v1/P15-B-FIN-RECON-001/)
- Track C: [`P15-C-INVOICE-001`](../tasks/pilot_v1/P15-C-INVOICE-001/)

## Deep technical review

1. Open the instruction and Agent-visible inputs without viewing the solution.
2. Solve or inspect the workbook as a professional deliverable.
3. Lock the candidate workbook and record its checksum.
4. Unblind to the Golden, Oracle, rubric, Judge, and fixtures.
5. Decide `ACCEPT`, `REVISE`, or `REJECT` using the [human review form](HUMAN_REVIEW_FORM.md).

The Pivot task may be reviewed for instruction and input sufficiency, but it must not be accepted as a valid task until the Windows Excel native-object receipt exists.

## Review criteria

1. The task resembles professional work and avoids benchmark-puzzle mechanics.
2. The instruction is clear and the supplied information is sufficient.
3. The reference workbook is professionally usable.
4. The oracle establishes correctness independently.
5. The rubric covers the material requirements.
6. The Judge accepts reasonable equivalents and rejects professional errors.
7. Observed model failures are separated from ambiguity and environment failures.
