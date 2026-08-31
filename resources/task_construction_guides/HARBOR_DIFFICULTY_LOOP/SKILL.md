---
name: harbor-difficulty-loop
description: Package P15 Excel tasks for Harbor separate verification and run evidence-bounded target-agent difficulty loops with DEV-CONFIRM freeze, valid-attempt accounting, and standard pass-at-k; use only after semantic fixture gates pass.
---

> Validated against P15 checkpoint: `8df20f98366633a4813cf6b98ee7e78270e3bdb7`
> Last evidence refresh: `2026-08-30`

# Harbor Difficulty Loop

## Purpose

Use Harbor to prove packaging and Judge execution, then measure target-agent capability with valid independent attempts. Keep infrastructure failures, quota limits, and incomplete jobs out of ability statistics.

## Required inputs

- A semantically valid task whose reference/equivalent/no-op/malformed/mutant gates already pass.
- A frozen DEV contract, materially distinct private CONFIRM sibling, Judge, task version, and target-model configuration.
- A Harbor 0.22 task package with separate Agent and verifier images.
- Three actually available target systems, attempt identifiers, score threshold `0.70`, and a failure-classification ledger.
- Current credential, quota, runtime, and Windows Excel availability status.

## Exact procedure

1. Build the verifier with `tests/` as its Docker context. Keep DEV/CONFIRM oracle, required source snapshots, perturbations, reference, and evaluator under `tests/`; never copy them into the Agent image.
2. Make the verifier write `/logs/verifier/reward.txt`. Initialize it to `0`, emit structured Judge output, validate that normalized score is finite and within `[0,1]`, and atomically promote the parsed reward. A verifier error must leave reward `0` and an inspectable failure.
3. Run Harbor Oracle, no-op, and malformed smoke once on the final package. Treat these as package/Judge execution evidence only, never as difficulty evidence.
4. Freeze DEV, CONFIRM, Judge, task version, target system/model settings, prompts, tool configuration, and pass threshold before collecting attempts.
5. Count an attempt only when the target system actually ran, produced a candidate workbook, Harbor completed verification, and a finite score was recorded. Classify quota, authentication, setup, environment, Windows, Harbor, reward-file, and incomplete-job failures separately; exclude them from `n` and `c`.
6. Run one valid pass@1 attempt on DEV for each of the three available target systems. Record the score and structured failure cause.
7. For still-valid, non-obviously-easy tasks, expand each system to four valid independent DEV attempts. At `n=4`, report raw scores and observed success fraction only; do not call it standard pass@8.
8. If a task remains a hard candidate, aim for `n=16` valid independent samples per target system. Let `c` be attempts with normalized score at least `0.70` and compute `pass@8 = 1 - C(n-c, 8) / C(n, 8)` only when `n >= 8`.
9. Require at least two of three systems to have `pass@8 < 0.70`, mean score across valid attempts below `0.60`, capability-caused failures, and the same trend on the held-out CONFIRM sibling. Without human review, label only `AGENT_HARD_CANDIDATE`.
10. If a task is too easy, make at most three revision rounds using genuine professional complexity. After each round, rerun only affected semantic fixtures, then the four-attempt screen. After round three, use the matching reserve or label `VALID_BUT_EASY`.
11. Save a checkpoint with task status, valid/invalid attempt accounting, failure classification, model configuration, quota state, and the next executable command.

## Required outputs

- Harbor Oracle/no-op/malformed smoke receipts and reward/Judge artifacts.
- A per-attempt ledger with target system, frozen version, split, candidate path, score, pass flag, validity, and structured failure class.
- A difficulty matrix with pass@1, four-attempt raw results, eligible `n=16` pass@8 calculations, mean scores, and CONFIRM trend.
- A revision log capped at three rounds and a truthful final label: invalid, valid, `VALID_BUT_EASY`, or `AGENT_HARD_CANDIDATE` pending human review.
- Machine, model, credential/quota, and external-blocker state sufficient for the next operator to continue.

## Validation checks

- Confirm the verifier runs from `/tests` without reading task-root metadata or data paths.
- Confirm reward is always present, numeric, finite, bounded, and consistent with Judge output.
- Oracle receives the expected valid score; no-op and malformed receive their expected failing scores and structured codes.
- Confirm no verifier-private assets appear in the Agent image or Agent-visible input bundle.
- Audit every excluded attempt and verify it was excluded for infrastructure, quota, authentication, setup, environment, or incomplete execution rather than low capability.
- Do not compute standard pass@8 for `n < 8`; for eligible tasks, recompute the combination formula from recorded `n` and `c`.
- Confirm DEV and CONFIRM were frozen before attempts and the held-out split was selected explicitly.

## Allowed difficulty changes

- Add source ambiguity, joins, exceptions, tolerance/closure, dynamic formulas, analytical comparisons, or native-object constraints that belong to the professional work.
- Replace an easy task with its predesigned reserve after the three-round limit.
- Refine a Judge only to close a demonstrated semantic coverage hole, followed by a new version freeze.

## Forbidden difficulty changes

- Counting quota, authentication, environment, setup, Windows, Harbor, reward-file, timeout, or incomplete-job failures as model failures.
- Describing Oracle/no-op/malformed smoke as target-agent difficulty.
- Computing or reporting pass@8 from four attempts, pooling incompatible model versions, or silently discarding valid low scores.
- Noise, missing necessary inputs, broken interfaces, token/time throttling, single-model traps, or arbitrary Judge score shaping.
- More than three revision rounds or the label `ACCEPTED_HARD` without real human evidence.
- A generic orchestration platform, runner, registry, dashboard, or trajectory framework.

## Common failure modes

- Local evaluation reaches task-root files, but Harbor's separate verifier contains only the `tests/` context.
- The Judge writes JSON but no reward file, causing a Harbor infrastructure failure.
- A smoke job with `model_name=null` is cited as a target-agent result.
- A real target process starts but produces no candidate or completed evaluation; the attempt is incorrectly counted as a zero.
- Four DEV attempts are used to claim pass@8, or CONFIRM reuses DEV values and is not held out.
- Revision adds irrelevant complexity instead of professional semantics.
- A Pivot task without a Windows-native reference is included in difficulty statistics.

## Stop conditions

- Stop a measurement batch when the planned valid-attempt budget is reached or all available target resources are genuinely exhausted and checkpointed.
- Stop and exclude a run when credentials, quota, environment, Harbor, reward generation, or candidate production fails.
- Stop revisions after three rounds; replace with a reserve or mark `VALID_BUT_EASY`.
- Keep Pivot invalid until Windows Excel completes native build, reopen, recalculation, validation, and receipt generation.
- Stop at `AGENT_HARD_CANDIDATE` until an actual human reviewer supplies acceptance evidence.

## Current verified examples

- Final Harbor smoke gave Oracle `1.0` for Sales, Ops, Fin, and Health; Pivot scored `0` with `TASK_INVALID` because native objects were missing. All five no-op runs scored `0` with `OUTPUT_MISSING`, and all malformed runs scored `0` with `MALFORMED_XLSX:BadZipFile`.
- These smoke jobs used no target model and therefore established packaging/Judge behavior, not difficulty.
- The current release ledger records 85 valid real Agent attempts: Codex has four DEV attempts per valid task, Claude Code + Opus 4.8 and Qwen Code + Qwen3.8-max have one each, and one Opus 5 DCF run is supplemental. The recorded paid runs have no Harbor environment exception.
- No task has standard pass@8 evidence because Codex has `n=4` and Claude/Qwen have `n=1` per valid task. The native Pivot task remains excluded.
- Sales, Ops, Fin, and Health have validated distinct private CONFIRM references. Pivot has a distinct CONFIRM design but remains `PENDING_EXTERNAL_WINDOWS_EXCEL` without a fabricated reference.

## Historical failure examples

- Before the current runner fixes, one real `gpt-5.6-sol` Sales attempt read the supplied inputs but its Harbor job ended without a candidate or score. This is a historical pre-fix infrastructure example and contributes neither `n` nor `c`.
