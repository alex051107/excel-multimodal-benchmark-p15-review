# Reuse Guide

This guide routes a new contributor to the smallest relevant resource. Reusing a package structure does not remove the need to re-establish professional truth, source rights, natural difficulty, and human validity for each new task family.

## Project construction

The repository has four operational layers.

| Layer | Responsibility | Main repository entry | Completion evidence |
| --- | --- | --- | --- |
| Task construction | Build typed truth, inputs, reference workbook, oracle, rubric, Judge, and negative fixtures | [`tasks/pilot_v1/`](../tasks/pilot_v1/) | Reference/equivalent/no-op/malformed/mutant score boundaries and five-run determinism |
| Harbor execution | Present the instruction and inputs to an Agent, collect `answer.xlsx`, and execute the Judge | Per-task `task.toml`, `environment/`, and `tests/` | Finite score, no runtime exception, and preserved run metadata |
| Agent evaluation | Freeze system, model, provider, tools, sample count, and budget | [`results/judge_v3_initial/`](../results/judge_v3_initial/) | Attempt-level rows and task-system score summaries |
| External validation | Check native Excel behavior and professional validity | [`review/`](../review/) | Windows Excel receipts and completed human review forms |

### Task package anatomy

| Path | Function |
| --- | --- |
| `instruction.md` | Agent-visible professional request and output contract |
| `data/input_files/` | All Agent-visible workbooks, documents, tables, and policies |
| `solution/reference.xlsx` | One complete and editable delivery |
| `metadata/oracle_recompute.py` | Independent recomputation from inputs and professional rules |
| `rubric.json` | Atomic criteria, weights, methods, and score aggregation |
| `tests/evaluate.py` | Deterministic Judge implementation |
| `fixtures/equivalent/` | Materially different acceptable solution |
| `fixtures/noop/` | Unmodified or effectively blank solution |
| `fixtures/malformed/` | Unreadable or structurally invalid workbook |
| `fixtures/mutants/` | Task-specific professional errors |
| `task.toml` | Harbor task metadata, artifact path, environment, and verifier contract |
| `receipts/` | Local, CONFIRM-reference, and Harbor smoke evidence |

### Configuration and execution flow

1. Freeze the task version, Judge version, and task checksum.
2. Record the system, model, provider, tools, sample count, and budget in the campaign configuration.
3. Set the Agent version, exact model identifier, provider/gateway, tool environment, sample count, and budget.
4. Launch the task through Harbor using the task's `task.toml` package contract.
5. Require the Agent to write `/app/output/answer.xlsx`.
6. Run the separate verifier image and save the normalized score, duration, token fields, cost fields, and exception status.
7. Keep environment failures outside capability statistics.
8. Aggregate valid attempts into the [score-only result tables](../results/judge_v3_initial/README.md); keep missing or temporarily unreadable files outside the score mean.

The clean repository publishes the configuration and result fields needed to understand the experiment. Credentials, account identifiers, raw provider responses, and internal machine paths are excluded.

### Scaling to the next batch

The practical scaling unit is:

```text
professional task family
  x source or template family
  x independent task instance
```

Add a small batch of genuinely different instances, validate them, and then decide which families merit further expansion. Do not create dozens of instances before the professional objective, source rights, oracle, Judge, and human-review boundary have been demonstrated for the family.

### Reuse checklist for a new task family

- [ ] The professional user, decision, and final handoff are specific.
- [ ] Source identity, usage rights, and required context are recorded.
- [ ] The supplied information is sufficient for one defensible professional result.
- [ ] Structured truth is established before the reference workbook is built or the document is rendered.
- [ ] The reference workbook is complete, editable, and recalculable.
- [ ] The Oracle recomputes the result from inputs and professional rules rather than copying the reference workbook.
- [ ] The Rubric contains at least six atomic requirements and covers the relationships that make the workbook usable.
- [ ] A materially different acceptable workbook and task-specific professional errors test the Judge boundary.
- [ ] Reference, equivalent, no-op, malformed, and mutant scores meet the release contract and repeat deterministically.
- [ ] Task, Judge, system version, provider, tools, budget, and sample count are frozen before formal Agent sampling.
- [ ] Windows Excel and independent human review are completed before native validity or accepted difficulty is claimed.

The checklist records common release evidence. Domain review still decides whether the work is professionally valid. Construction stops when the task's professional truth, information sufficiency, or human validity cannot be established.

| Goal | Read | Required input | Output | Stop when |
| --- | --- | --- | --- | --- |
| Select a new task | [Task Selection](../resources/task_construction_guides/TASK_SELECTION/SKILL.md) | A defined professional user, decision, source, and deliverable | Admitted task-family card or design-only reserve | Necessary information, source rights, or independent truth cannot be established |
| Build Track A | [Track A Construction](../resources/task_construction_guides/TRACK_A_CONSTRUCTION/SKILL.md) | Recalculable model, assumptions, or one controlled defect | Formula/model task package | The target can only be recovered by guessing hidden conventions |
| Build Track B | [Track B Construction](../resources/task_construction_guides/TRACK_B_CONSTRUCTION/SKILL.md) | Multi-file rules, versions, joins, exceptions, and control totals | Multi-file analytical task package | Difficulty depends on irrelevant clutter or missing rules |
| Build Track C | [Track C Construction](../resources/task_construction_guides/TRACK_C_CONSTRUCTION/SKILL.md) | Typed truth and a controlled document template | Document-to-workbook task package | The answer must be inferred from a lossy rendering rather than typed truth |
| Build Golden and Oracle | [Oracle and Gold](../resources/task_construction_guides/ORACLE_AND_GOLD/SKILL.md) | Frozen inputs and professional rules | Reference workbook, independent oracle, equivalent, and perturbations | Oracle output depends on copying the reference workbook |
| Build Rubric and Judge | [Rubric and Judge](../resources/task_construction_guides/RUBRIC_AND_JUDGE/SKILL.md) | Frozen truth, acceptable equivalents, and professional errors | Atomic rubric, deterministic evaluator, fixtures | Reasonable equivalents cannot be distinguished from errors |
| Run Agents | [Harbor Difficulty Loop](../resources/task_construction_guides/HARBOR_DIFFICULTY_LOOP/SKILL.md) | Frozen task, Judge, models, provider, and budget | Attempt records and difficulty signals | Failures are environmental or the task/Judge is still changing |
| Review and release | [Release and Scale](../resources/task_construction_guides/RELEASE_AND_SCALE/SKILL.md) | Frozen tasks and honest evidence status | Review package and release manifest | A report would promote local checks into Windows, Agent, or human evidence |

## Repository examples

- Task design: [`docs/TASK_DESIGN.md`](TASK_DESIGN.md)
- Complete package: [`tasks/pilot_v1/P15-B-FIN-RECON-001/`](../tasks/pilot_v1/P15-B-FIN-RECON-001/)
- Evaluation principles: [`docs/JUDGE_V3_ADJUDICATION_CONTRACT.md`](JUDGE_V3_ADJUDICATION_CONTRACT.md)
- Current score-only results: [`results/judge_v3_initial/README.md`](../results/judge_v3_initial/README.md)
- Human review: [`review/REVIEW_GUIDE.md`](../review/REVIEW_GUIDE.md)
