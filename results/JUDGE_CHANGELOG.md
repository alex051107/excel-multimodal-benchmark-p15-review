# Judge Changelog

Original attempt scores and effective regrades remain separate in [`ATTEMPTS.csv`](ATTEMPTS.csv). Regrading an existing `answer.xlsx` does not create a new attempt.

| Task / system | Raw score | Effective score | Change |
| --- | ---: | ---: | --- |
| `P15-A-STAT-EXPERIMENT-001` / Opus 4.8 | 0.166667 | 1.000000 | The earlier Judge accepted `STDEV.S` but rejected Excel's valid `STDEV` alias. The current Judge accepts both and accepts equivalent decision formulas that share the same p-value and alpha dependencies. |
| `P15-B-HEALTH-REPORT-001` / Opus 4.8 | 0.444444 | 0.925926 | The earlier Judge could not evaluate report text assembled with `TEXT(...)` and `&`. The current Judge accepts formula-generated numeric and percentage text while retaining the required dependency on the period metrics. |
| `P15-A-FIN-DCF-001` / supplemental Opus 5 | 0.600000 | 1.000000 | The earlier formula engine incorrectly required at least two arrays in `SUMPRODUCT`. The current evaluator treats the valid one-array form as an array sum. |

The other 26 paid attempts have identical raw and effective scores. After these changes, the task-local deterministic fixtures and the final 45 reference/no-op/malformed Harbor smoke runs remained within their expected score boundaries.
