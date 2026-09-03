# Literature-to-Design Map

This file records the research mechanisms used in P15. The project adapted methods; it did not import the cited benchmarks' tasks, workbooks, annotations, rubrics, or leaderboard results.

## 1. Rewind a completed workbook into a task

**Primary sources:** BlueFin, arXiv:2605.30907v1, sections 3.1, 4, 4.2, and Appendix E.1; WTM-Bench, arXiv:2608.07873v1, sections 4.1–4.3; SpreadsheetBench 2, arXiv:2606.29955v1, sections 2.1–2.4.

**P15 adaptation:** Establish a professionally correct completion first, then delete target regions or inject one controlled root-cause defect. Preserve required assumptions and protected regions.

**Current examples:** DCF and financial-model repair.

**Construction resources:** [Track A Construction](../resources/task_construction_guides/TRACK_A_CONSTRUCTION/SKILL.md) and [Oracle and Gold](../resources/task_construction_guides/ORACLE_AND_GOLD/SKILL.md).

**Stop condition:** The task can only be solved by guessing a hidden convention, or the completed workbook cannot be independently recalculated.

## 2. Test dynamic correctness and hard-coding

**Primary sources:** BlueFin, arXiv:2605.30907v1; Segura et al., *A Survey on Metamorphic Testing*, DOI:10.1109/TSE.2016.2532875; CausalFlow, arXiv:2605.25338v1.

**P15 adaptation:** Change a declared professional input, recompute expected downstream values independently, and verify both result propagation and protected invariants.

**Current examples:** DCF assumptions, reconciliation FX rates and approved adjustments, policy scenario inputs.

**Construction resources:** [Oracle and Gold](../resources/task_construction_guides/ORACLE_AND_GOLD/SKILL.md) and [Rubric and Judge](../resources/task_construction_guides/RUBRIC_AND_JUDGE/SKILL.md).

**Stop condition:** The check only proves that a cell changed, or the perturbation changes the meaning of the task.

## 3. Preserve versions, joins, exceptions, and control totals

**Primary sources:** FINCH, arXiv:2512.13168v5, sections 2.1–2.3; DataGovBench, arXiv:2512.04416v2, sections 3.1–3.5; FinBalance, arXiv:2606.15949v1, pages 1 and 4–6.

**P15 adaptation:** Encode authoritative versions, join keys, tolerances, exception queues, approved adjustments, and control totals in typed truth and evaluator logic.

**Current examples:** Q2 source selection, order cleaning and joining, monthly ledger reconciliation.

**Construction resources:** [Track B Construction](../resources/task_construction_guides/TRACK_B_CONSTRUCTION/SKILL.md) and [Rubric and Judge](../resources/task_construction_guides/RUBRIC_AND_JUDGE/SKILL.md).

**Stop condition:** Failure depends on irrelevant files, obscure naming, or missing join rules rather than the professional workflow.

## 4. Render documents from structured truth

**Primary sources:** DocILE, arXiv:2302.05658v2, sections 1–3; ReceiptBench, ACL 2026 paper 2135; VAREX, arXiv:2603.15118v2, sections 1–3.

**P15 adaptation:** Freeze document, page, line-item, and amount relationships in typed data before rendering PDF or image inputs. Score against the pre-render truth.

**Current examples:** Invoice, quote, receipt batch, purchase-order amendment, and bank statement.

**Construction resources:** [Track C Construction](../resources/task_construction_guides/TRACK_C_CONSTRUCTION/SKILL.md), [Oracle and Gold](../resources/task_construction_guides/ORACLE_AND_GOLD/SKILL.md), and [Rubric and Judge](../resources/task_construction_guides/RUBRIC_AND_JUDGE/SKILL.md).

**Stop condition:** The answer must be reconstructed from a lossy rendered artifact because the typed truth was not preserved.

## 5. Check information sufficiency before construction

**Primary source:** QuestBench, arXiv:2503.22674v2.

**P15 adaptation:** List the variables, rules, source identities, and output relationships required to complete the task. Stop construction if necessary information is missing or multiple answers remain equally defensible.

**Current examples:** All 15 tasks.

**Construction resource:** [Task Selection](../resources/task_construction_guides/TASK_SELECTION/SKILL.md).

**Stop condition:** Missing information is being used as difficulty.

Exact source links are listed in [`SOURCE_CATALOG.md`](SOURCE_CATALOG.md).

## 6. Diagnose the harness and test evidence-relation value

**Primary sources:** Harness-Bench, arXiv:2605.27922; WorkSurface-Bench, arXiv:2607.25765; ContractBench, arXiv:2605.17281; DataFlow-Harness, arXiv:2607.16617; CausalFlow, arXiv:2605.25338.

**P15 adaptation:** Attribute observed capability to the model-harness-tool configuration; separate discovery, routing, evidence use, artifact construction, and verification; compare static-native, dynamic-native, and source-to-artifact relation checks on the same frozen artifacts; describe the first observable violation unless intervention and downstream replay support a causal statement.

**Novelty boundary:** Harness diagnostics, intermediate artifact contracts, typed workflow mutation, failure taxonomies, and causal localization are prior art. P15's narrower empirical question is whether source identity, authority, and support relations catch professionally meaningful native-workbook failures missed by a strong dynamic-native baseline, and whether the resulting witness improves local repair on family-disjoint tasks.

**Stop condition:** E2 adds no human-confirmed recall over B1, rejects valid equivalents, depends on hidden Gold or a fixed reference layout, or fails to transfer across private families.

The complete measurement and repair protocol is in [`docs/RESULT_ANALYSIS_AND_HARNESS_PLAN.md`](../docs/RESULT_ANALYSIS_AND_HARNESS_PLAN.md).

## Difficulty evidence and interpretation

These benchmarks use different units, tasks, Agents, and metrics. Their absolute scores are not comparable with one another or with P15. The table records the failure mechanism that informs P15 design.

| Source | Fixed location | Reported difficulty evidence | Design lesson used in P15 |
| --- | --- | --- | --- |
| BlueFin, arXiv:2605.30907v1 | Sections 3.1, 4.1, and 4.2; PDF pages 5–8 | The strongest evaluated model remained below 50% overall. Formula correctness was much higher than dynamic correctness after input changes. | Test cross-sheet propagation, sign and date conventions, period logic, hard-coding, and workbook integration. |
| WTM-Bench, arXiv:2608.07873v1 | Sections 4.1–4.3, 5, and 8; PDF pages 4–10 | The best reported system reached 33.7% hard score and 54.0% soft score. More object transformations reduced performance; Pivot soft scores were 0%–10%. | Restore mutually dependent spreadsheet objects from a professionally correct completed state. |
| SpreadsheetBench 2, arXiv:2606.29955v1 | Sections 2.1–2.4 and 3.2–3.3; PDF pages 3–9 | Opus 4.6 reached 34.89% complete-task accuracy. Target cells could be correct while the whole workbook failed. | Check the intended output, downstream dependencies, and preservation of non-target regions. |
| FINCH, arXiv:2512.13168v5 | Sections 2.1–2.2 and 3.2–3.3; PDF pages 3–13 | Reported human pass rate fell from 48.6% for one-stage work to 5.6% for workflows with four or more stages. | Combine stages only when they form one professional workflow over shared business materials. |
| DataGovBench, arXiv:2512.04416v2 | Sections 3.1–3.5 and 6.1; PDF pages 3–7 | Code-running rates were substantially higher than task-success rates. | Separate file or pipeline validity from business correctness; inject task-relevant defects only. |
| FinBalance, arXiv:2606.15949v1 | Sections 3, 5–7 and Appendix difficulty rubric; PDF pages 3–11 | The best reported final-statement accuracy was 46%; the highest difficulty tier reached 0% final-statement accuracy in the reported evaluation. | Preserve evidence identity, accounting treatment, provenance, ledger replay, and report closure. |
| VAREX, arXiv:2603.15118v2 | Sections 3.1 and 4.1–4.4; PDF pages 3–7 | Frontier models exceeded 94% on the reported field-extraction metric. | Reuse its traceable typed-truth construction, but do not treat OCR degradation as a professional difficulty mechanism. |
| DocILE, arXiv:2302.05658v2 | Sections 3.3–3.5, 4.1–4.2, and 5.7; PDF pages 7–15 | Line-item recognition requires fields to be grouped into the correct document rows, not merely extracted. | Score field value, document identity, page provenance, and line-item grouping separately. |

Across the cited studies, low performance follows from several professional relationships that must remain correct at the same time. None of the sources supports missing necessary information, unreadable images, irrelevant file volume, interface defects, or model-specific traps as valid difficulty.
