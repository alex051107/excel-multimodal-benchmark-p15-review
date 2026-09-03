# Results Directory

This directory contains two separate evidence generations.

## Historical development evidence

- `ATTEMPTS.csv`, `RUN_RESULTS.csv`, `MODEL_SUMMARY.csv`, and `TASK_STATUS.csv` describe the earlier development comparison.
- `GROUP_GATEWAY_CHECKPOINT.csv` preserves the first interrupted 41-attempt group-gateway checkpoint.
- These files are historical and are not overwritten by the current campaign.

## Current n=8 development campaign

- [`N8_SNAPSHOT.json`](N8_SNAPSHOT.json): current target, coverage correction, cost basis, and claim boundaries.
- [`N8_ATTEMPTS.csv`](N8_ATTEMPTS.csv): sanitized terminal-receipt table. It excludes prompts, local paths, credentials, provider responses, trajectories, and workbooks.
- [`N8_TASK_SYSTEM_METRICS.csv`](N8_TASK_SYSTEM_METRICS.csv): one row per task and system, with frozen-baseline and current-contract results kept separate.
- [`N8_TASK_SUMMARY.csv`](N8_TASK_SUMMARY.csv): compact 15-task view.
- [`N8_SYSTEM_SUMMARY.csv`](N8_SYSTEM_SUMMARY.csv): compact three-system view.

The `controller_coverage_n` field is a development coverage counter, not a formal same-contract sample size. Forty-one frozen baseline attempts are reused to avoid paying for duplicate work, but the current campaign uses a different execution contract. Standard pass@8 is therefore not reported.

All current means are provisional because Judge v2 has not been frozen and applied uniformly to the preserved artifacts. `model_artifact_outcome_valid`, `runtime_clean`, and `score_admissible` are independent fields. Recovered workbooks from abnormal Trials remain available for diagnosis, while only strict score-admissible rows increment coverage. A replacement always receives a new attempt ID.

Displayed ZCloud units are not labeled as US dollars without an independently verified unit-to-USD conversion.

Regenerate the public-safe snapshot from a local development checkout with:

```bash
python3 scripts/export_current_n8_snapshot.py \
  --source-repo /path/to/ExcelBench-P15-Goal \
  --campaign-config path/relative/to/source/repo/campaign.json \
  --evidence-cutoff 2026-09-03T12:34:49.183770+00:00 \
  --usage-receipt path/relative/to/source/repo/latest_usage.json \
  --source-controller-head 7011da968a5d3de28354b04c78aa4d684fb230ec \
  --expected-runtime-receipts 418 \
  --output-dir results
```

The exporter reads local runtime evidence but writes only the allowlisted fields documented above.
