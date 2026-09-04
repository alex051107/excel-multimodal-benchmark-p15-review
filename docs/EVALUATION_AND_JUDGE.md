# Evaluation and Judge

## Current trust status

The scores in the current n=8 snapshot are provisional. The active runtime branch does not yet include the reviewed Judge v2 changes, and the preserved Agent workbooks have not been uniformly regraded under one frozen evaluator version. Current means can guide audit sampling; they cannot support a formal ranking or accepted-hard decision.

Judge v2 must preserve the original scoring philosophy while enforcing four implementation rules:

1. every independently assessable rubric criterion continues to run after a local failure;
2. a penalty applies only to additional harmful behavior such as deletion, overwrite, out-of-scope modification, or fabrication, not to the same omission already scored by a positive criterion;
3. a true delivery-critical hurdle may block pass but must not erase independently earned continuous score;
4. evaluator or infrastructure failure produces `N/A`, never a model score of zero.

The task-local Judge may accept exact names, predefined aliases, or deterministic role recognition. It must not use a hidden canonical layout as a global zero condition when the instruction permits an equivalent workbook structure.

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

Those three earlier corrections are not the same as the pending portfolio-wide Judge v2 freeze. The current campaign includes direct Guard verdicts, evidence-bound raw-Judge attestations, offline evaluator attestations, and macOS Excel canonicalized Pivot verdicts. A formal analysis requires every preserved candidate to be bound to the exact evaluator version used for its effective score, then regraded under the one frozen Judge version.

Seven old attestations first exposed that countability had been accepted without a bound candidate workbook. The corrected controller now excludes all Guard-forced and infrastructure-invalid Trials from strict coverage. Recovered workbooks remain available for diagnostic regrade, with Agent completion, runtime health, native-Excel consistency, and score admissibility reported independently in [`N8_ATTEMPTS.csv`](../results/N8_ATTEMPTS.csv). New countable attestations must bind the actual candidate and its hash.

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
> 历史说明：本页记录旧版评分设计。当前飞书汇报和新版结果请看 [FEISHU_SUBMISSION.md](../outputs/p15_final/FEISHU_SUBMISSION.md) 与 [新版评分初步结果](../results/judge_v3_initial/README.md)。
