# Result Analysis and Semantic Harness Plan

## Decision

P15 should be analyzed as a capability-diagnosis and measurement-calibration set, not as a model leaderboard. The correct evidence chain is:

```text
result admissibility
→ Judge validity
→ task × system behavior
→ capability-level failures
→ task difficulty and discrimination
→ B0/B1/E2 measurement ablation
→ witness-guided repair
→ family-disjoint generalization
```

The current 15 tasks can generate hypotheses about Agent weaknesses. They cannot establish broad generalization because repeated attempts on one task instance measure stochastic stability, not coverage across materially distinct task families.

## 1. Research questions

### RQ0 — Is each result interpretable?

Every attempt must keep six states separate. In particular, do not compress artifact delivery, runtime health, and score admissibility into one `countable` flag:

| State | Question | Effect on ability analysis |
| --- | --- | --- |
| Provider | Was the request accepted by the intended provider and model route? | A rejected request is not a model outcome. |
| Runtime | Did the Agent finish naturally, and did the surrounding Trial reach a clean terminal state? | An interruption before the Agent completes makes the model outcome `N/A`. A failure after an immutable workbook was delivered may preserve that artifact outcome for offline regrade, but it is not a clean end-to-end Trial. |
| Artifact | Was a readable `answer.xlsx` delivered? | Under the current P15 counting contract, no workbook means no countable sample. |
| Native Excel | Were recalculation and native objects checked in the required Excel backend? | A pending native check blocks the relevant conclusion. |
| Judge | Did the evaluator return a finite, trustworthy result? | `JUDGE_ERROR` is `N/A`, not zero. |
| Billing | What did the run cost and how long did it take? | Used for efficiency analysis, never to decide correctness. |

The analysis unit is a **model + Agent harness + tools and permissions + provider route + runtime configuration**. Results should not be attributed to the base model alone.

### RQ1 — What can each Agent configuration do reliably?

For each system, answer:

1. What is the empirical probability of success on each frozen task?
2. How variable are eight repeated attempts on the same task?
3. Which tracks and capability dimensions are consistently strong or weak?
4. Which failures are shared across systems, and which are system-specific?
5. What time and cost are required for comparable artifact quality?

### RQ2 — Which tasks are easy, hard, diagnostic, unstable, or invalid?

Low score alone does not establish task difficulty. Each task should receive one reviewed disposition:

| Disposition | Required pattern |
| --- | --- |
| `VALID_BUT_EASY` | Valid task and Judge; systems repeatedly succeed. |
| `INFORMATIVE_HARD_CANDIDATE` | Human-solvable and valid; at least two target systems fail consistently for capability reasons. |
| `MODEL_SPECIFIC_DIAGNOSTIC` | One system fails while the others reliably succeed. |
| `NOISY_OR_UNSTABLE` | Large within-system variation or divergent tool paths. |
| `INVALID_OR_AMBIGUOUS` | Necessary information, professional intent, or acceptable solution space is not well defined. |
| `JUDGE_OR_BACKEND_BLOCKED` | Scores remain uninterpretable because the Judge or native Excel requirement is unresolved. |

An accepted-hard claim still requires the frozen numerical gate, human failure attribution, and the same trend on a family-disjoint held-out task.

### RQ3 — Does evidence-relation evaluation add information?

The measurement comparison is:

| Layer | What it checks |
| --- | --- |
| `B0_STATIC_NATIVE` | Final values, fields, one-state workbook structure, and required native objects. |
| `B1_DYNAMIC_NATIVE` | Formulas, recalculation, perturbation propagation, dependency preservation, and change locality. |
| `E2_EVIDENCE_RELATION` | Source identity, authority, support, and source-to-artifact relationships. |

E2 is useful only when it finds a human-confirmed, professionally meaningful failure that a strong B1 evaluator misses, without increasing rejection of materially different valid equivalents. Dynamic correctness is already established prior art in spreadsheet evaluation; it is a baseline, not this project's novelty.

### RQ4 — Does a clause-level witness improve repair?

Use frozen natural failures and compare:

1. generic retry;
2. scalar score;
3. violated clause only;
4. clause plus evidence witness;
5. clause, witness, and a targeted probe.

Primary outcomes are independent-Judge repair success, recovery of the original violated clause, newly introduced violations, and protected-region damage. Secondary outcomes are edit size, full-workbook regeneration, latency, token use, and cost. Include a length-matched control so that extra information is not mistaken for a witness-specific effect.

### RQ5 — Does the result generalize across families?

Formal generalization requires materially distinct families, family-blocked splits, and private tasks that are opened only after the evaluator and repair method are frozen. Layout, source files, templates, and task truth must not cross the split.

## 2. Statistical reporting hierarchy

### Attempt level

Report countability, artifact delivery, score, pass/fail, Judge status and version, native Excel status, runtime, cost, and failure class. Attempts are repeated observations; they are not independent tasks.

### Task × system level

Report valid `n`, score distribution, successes `c`, empirical success `c/n`, interval estimates, within-cell variance, cost, time, and criterion-level loss. Preserve the raw eight scores so that `0/8`, `1/8`, and `8/8` remain distinguishable.

The standard estimator

```text
pass@8 = 1 - C(n-c, 8) / C(n, 8)
```

is degenerate when `n = 8`: it is zero when `c = 0` and one whenever `c >= 1`. It must be reported alongside `c/n`, mean score, and the raw distribution. The original requirement remains ambiguous between this estimator and the empirical success rate over eight runs; no formal pass@8 conclusion is allowed until that definition is frozen.

### Task level

Report a task-balanced mean across systems, system differences, within-system variance, missing/Judge-error rates, human completion evidence, and the reviewed task disposition. Do not fit item-response models to 15 tasks and three systems; the parameters would not be stable.

### Track and capability level

Use task-balanced macro averages. For each capability, divide failed weight by the rubric weight that actually exposed that capability. Raw failure counts are misleading when tasks contain different numbers and weights of criteria. Report how many criteria, tasks, and families support each capability estimate.

### Program level

Report the distribution across tracks, worst-quartile performance, artifact completion, cost/time/quality trade-offs, and counts of easy, hard-candidate, model-specific, noisy, invalid, and blocked tasks. A single pooled score is not the main result.

## 3. Capability and failure coding

Code each failure on two axes.

### Execution stage

1. `DISCOVERY_AND_INSPECTION`
2. `SOURCE_SELECTION`
3. `EVIDENCE_ACQUISITION`
4. `TRANSFORMATION_OR_COMPUTATION`
5. `NATIVE_ARTIFACT_CONSTRUCTION`
6. `VERIFICATION_AND_DELIVERY`

### Semantic failure type

- `INSPECTION_INCOMPLETE`
- `SOURCE_AUTHORITY_ERROR`
- `IDENTITY_OR_JOIN_ERROR`
- `DOMAIN_RULE_ERROR`
- `DEPENDENCY_BREAK`
- `HARDCODED_OUTPUT`
- `EXCEPTION_OR_COVERAGE_LOSS`
- `NATIVE_OBJECT_SUBSTITUTION`
- `PROVENANCE_ERROR`
- `COLLATERAL_EDIT`
- `VERIFICATION_SKIPPED`
- `UNDERSPECIFICATION_NOT_QUERIED`
- `DELIVERY_FAILURE`

Each label also carries an evidence strength:

- `ARTIFACT_CONFIRMED`: directly supported by the workbook and deterministic checks;
- `TRAJECTORY_CONSISTENT`: consistent with the trajectory and reported only as the first observable violation;
- `INTERVENTION_SUPPORTED`: supported by replacing a candidate step and replaying downstream execution.

Without intervention and replay, the project must not call the first logged anomaly a causal root cause.

## 4. Judge trust gate

The current scores are provisional. The active runtime branch does not contain the reviewed Judge v2 changes, and historical Judges have produced global zeroes, duplicate penalties, and `OUTPUT_MISSING` results that were confused with infrastructure outcomes.

Before model ranking or difficulty interpretation:

1. freeze one Judge v2 commit, task-local rubric, Oracle, and role/alias rules;
2. re-evaluate every preserved `answer.xlsx` without rerunning the Agent or increasing `n`;
3. represent evaluator and infrastructure failures as `N/A`;
4. verify reference, two materially different equivalents, no-op, malformed, and semantic mutants;
5. blind-review artifacts with the largest v1/v2 difference, zero scores, threshold crossings, nonstandard layouts, and Judge errors;
6. report criterion sensitivity, equivalent false-rejection, and five-run repeat stability;
7. test whether Judge missingness is concentrated in a model or task before comparing only scoreable rows.

Rubric-level ability analysis begins only after this gate. Until then, current task and system means are diagnostic previews.

## 5. B0/B1/E2 measurement ablation

Use the same frozen artifacts and blind human labels for all three evaluator layers. Include:

- reference and materially different equivalents;
- `M0`: a defect visible to B0;
- `M1`: B0 passes and B1 fails;
- `M2`: B1 passes and E2 fails for a meaningful source/authority/support reason;
- `M3`: a valid equivalent that all layers should accept.

Report human-unacceptable failure recall, equivalent specificity, B1-missed/E2-caught rate, witness precision, evidence-locator accuracy, applicability coverage, deterministic repeatability, and review-time reduction. Use paired comparisons on the same artifact and task- or family-clustered uncertainty intervals. Mutants do not increase the number of independent task families.

## 6. Turning stable failures into harder tasks

Do not modify a task merely because a target model failed it. Use this sequence:

1. confirm that the failure is not caused by the Judge, permissions, runtime, or information ambiguity;
2. reproduce the capability gap across multiple tasks and preferably distinct families;
3. map the gap to a real professional work event;
4. create a new task version or family in which that capability is necessary;
5. increase only one or two natural complexity axes while keeping two to four primary stress mechanisms;
6. validate with an independent Oracle, multiple valid equivalents, and a blind human solve;
7. test on a family-disjoint validation/private family without changing its truth after results are seen.

Natural difficulty axes are source/evidence complexity, workbook/dependency complexity, workflow horizon, domain semantics, diagnosis/change locality, and deliverable integration. Valid examples include deeper cross-sheet propagation, authoritative-version conflicts, join cardinality and exception closure, native-object refresh, cross-page hierarchy, and revision precedence. Unreadable inputs, hidden necessary rules, irrelevant file volume, arbitrary time limits, and numeric template swaps are not acceptable difficulty mechanisms.

Keep easy tasks as calibration anchors. A useful benchmark needs a difficulty spectrum, not 150 uniformly failing tasks.

## 7. Novelty claim ladder

Novelty is an empirical sequence, not a feature list:

1. **Benchmark contribution:** establish valid family-disjoint native Excel tasks across computational, analytical, and document-grounded work.
2. **Measurement contribution:** show that E2 catches meaningful human-rejected failures missed by a strong dynamic-native B1 baseline.
3. **Improvement contribution:** show that clause-plus-evidence witness improves local repair over generic retry, scalar feedback, and clause-only feedback.
4. **General Semantic Harness:** claim this only after the effect holds in at least two Excel tracks, on private families, and in a non-Excel transfer setting without reference-layout or hidden-gold dependence.

Until the final step is supported, the accurate name is **relation-aware semantic evaluator for Excel agents**.

## 8. Current P15 evidence boundary

- The live development campaign snapshot and its corrected counting boundary are in [`results/N8_SNAPSHOT.json`](../results/N8_SNAPSHOT.json).
- Task-level and system-level provisional means are in [`results/N8_TASK_SUMMARY.csv`](../results/N8_TASK_SUMMARY.csv) and [`results/N8_SYSTEM_SUMMARY.csv`](../results/N8_SYSTEM_SUMMARY.csv).
- Seven legacy attestation rows without a candidate workbook have been excluded from current coverage and must be resubmitted with new attempt IDs; their old receipts remain preserved.
- Abnormal runs are re-submitted as new attempts. A completed normal attempt is never overwritten or repeated merely because a later reporting table changes.
- Forty-three runs in which the Agent had already completed and delivered a hash-bound workbook before a verifier or cleanup interruption are retained as recoverable artifact outcomes at the frozen cutoff. They remain runtime-unhealthy and are not described as clean end-to-end trials. One mid-Agent interruption and four distinct Pivot workbooks with inconsistent native-Excel receipt chains are `N/A` pending replacement or revalidation.
- Frozen baseline attempts and current ZCloud attempts remain separate in [`results/N8_ATTEMPTS.csv`](../results/N8_ATTEMPTS.csv). They are combined only to avoid duplicate work, not to produce formal same-contract pass@8.
- Judge v2, external human review, and Windows Excel compatibility are not frozen or complete.
- This repository is public and its committed development Gold and CONFIRM assets must be treated as disclosed. They cannot support future unseen-test or anti-contamination claims.

## 9. Required analysis outputs

The full study should produce six reviewable artifacts:

1. `RESULT_ADMISSIBILITY_TABLE`: orthogonal provider/runtime/artifact/native/Judge status for every attempt;
2. `ABILITY_MATRIX`: system by track, archetype, and exposure-normalized capability;
3. `TASK_DIFFICULTY_DISPOSITION`: reviewed easy/hard-candidate/model-specific/noisy/invalid/blocked labels;
4. `FAILURE_ATLAS`: artifact evidence, first observable violation, and B0/B1/E2 layer;
5. `HARNESS_ABLATION_REPORT`: incremental detection, equivalent false rejection, and repair comparisons;
6. `NEXT_FAMILY_DESIGN_BRIEFS`: two or three family-disjoint task routes for each stable capability deficit.

## 10. Primary-source comparators

- [SpreadsheetBench 2](https://arxiv.org/abs/2606.29955) motivates workflow-level analysis and inspection/target-cell failure coding.
- [BlueFin](https://arxiv.org/abs/2605.30907) provides a strong comparator for granular rubric and dynamic spreadsheet correctness.
- [Harness-Bench](https://arxiv.org/abs/2605.27922) supports reporting the model-harness configuration rather than the base model alone.
- [WorkSurface-Bench](https://arxiv.org/abs/2607.25765) motivates separating routing, evidence use, answer quality, and efficiency.
- [ContractBench](https://arxiv.org/abs/2605.17281) places intermediate artifact contracts and failure labels in prior art.
- [DataFlow-Harness](https://arxiv.org/abs/2607.16617) is a comparator for typed workflow mutation and structural execution validity.
- [CausalFlow](https://arxiv.org/abs/2605.25338) sets the intervention-and-replay boundary for causal failure attribution.

Search scope: targeted primary-source review, updated 2026-09-03. This is a design-oriented comparator scan, not a systematic literature review.
