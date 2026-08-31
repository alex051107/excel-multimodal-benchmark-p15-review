# Evaluation and Judge

## Four separate correctness components

| Component | Purpose | Reason for separation |
| --- | --- | --- |
| Golden Solution | Demonstrates one complete, editable delivery | Makes the target workbook understandable without defining all acceptable layouts |
| Oracle | Recomputes correct results from inputs and professional rules | Prevents the Judge from merely copying the Golden layout |
| Rubric | Divides the task into atomic, weighted requirements | Makes a score explainable at the criterion level |
| Judge | Opens the candidate workbook and executes the rubric checks | Produces deterministic, repeatable scores |

## Local acceptance contract

For each locally valid task:

- reference score = 1.0;
- materially different acceptable equivalent >= 0.95;
- no-op < 0.30;
- malformed workbook = 0;
- every task-specific semantic mutant < 0.70;
- five repeated scores for the same deterministic output are identical.

The current portfolio has 164 rubric criteria. Task-specific negative fixtures cover 128 criteria; 26 still need a criterion-linked mutant, and 10 require Windows Excel evidence.

## Score interpretation

Scores are normalized to the range 0 to 1. A score of 0.70 or above counts as one successful attempt. This threshold is fixed before difficulty interpretation; it does not mean that every workbook above 0.70 is ready for publication.

For `n` independent attempts and `c` successes, empirical pass@1 is `c / n`. Standard pass@8 is:

```text
1 - C(n-c, 8) / C(n, 8)
```

The formula is not reported when `n < 8`. Formal hard-candidate evaluation preferentially targets `n = 16` per task and system.

## Regrade policy

Real Agent workbooks exposed three false-negative Judge boundaries:

| Task | Earlier false negative | Corrected boundary |
| --- | --- | --- |
| Paired statistics | Accepted `STDEV.S` but rejected Excel's valid `STDEV` alias | Accepts both aliases and equivalent decision formulas |
| Public-health report | Could not evaluate report text assembled with `TEXT(...)` and `&` | Accepts formula-generated numeric and percentage text while retaining metric dependencies |
| DCF | Rejected a valid one-array `SUMPRODUCT(array)` formula | Evaluates the one-array form as an array sum |

A regrade reuses the same `answer.xlsx` and does not increase the attempt count. Original and effective scores remain distinguishable in [`results/ATTEMPTS.csv`](../results/ATTEMPTS.csv).

## Harbor boundary

Harbor supplies the instruction and inputs, receives `answer.xlsx`, runs the independent Judge, and records score, duration, and run metadata. The 45 reference/no-op/malformed smoke runs confirm package loading and Judge execution. They do not measure task difficulty.

## Minimal package check

From a task directory, the local shell entry point evaluates the reference workbook:

```bash
cd tasks/pilot_v1/P15-B-FIN-RECON-001
sh tests/test.sh
```

Running the Judge on another candidate requires the dependencies declared by the task's verifier image:

```bash
python3 tests/evaluate.py /path/to/answer.xlsx --split dev
```

The Harbor package contract is defined in `task.toml`, `environment/Dockerfile`, and `tests/Dockerfile` within each task directory.
