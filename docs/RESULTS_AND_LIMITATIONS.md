# Results and Limitations

## Current result

P15 contains 15 primary task candidates and nine reserve designs. All 15 primary tasks pass the current local answer-and-evaluator checks. The Pivot task was rebuilt as an update to a genuine Excel template and read back in Microsoft Excel for Mac with one PivotCache, one PivotTable, two SUM measures, a Q2 filter, refresh behavior, and one PivotChart. Windows Excel compatibility remains pending.

The three target systems have completed real development runs on every locally valid task. The current sample sizes differ: Codex has four runs per task, while Claude and Qwen each have one. The results are sufficient to identify obvious easy cases and set review priorities. They are too small to establish stable model rankings or formal task difficulty.

| System | Attempts | Successful attempts | Mean score | Evidence scope |
| --- | ---: | ---: | ---: | --- |
| Codex CLI + GPT-5.6 sol | 56 | 26 | 0.634 | Four DEV attempts per valid task |
| Claude Code + Opus 4.8 | 14 | 8 | 0.678 | One paid DEV attempt per valid task |
| Qwen Code + Qwen3.8-max | 14 | 3 | 0.488 | One paid DEV attempt per valid task |
| Claude Code + Opus 5 | 1 | 1 | 1.000 | Supplemental DCF run only; excluded from the three-system comparison |

Detailed task-by-system results are in [RUN_RESULTS.csv](../results/RUN_RESULTS.csv). Attempt-level scores, durations, token counts, cost fields, and regrade deltas are in [ATTEMPTS.csv](../results/ATTEMPTS.csv).

## Group-gateway n=8 checkpoint

A second development batch uses GPT-5.6 sol, Opus 5, and Qwen3.8-max through the group gateway. It was configured for eight independent attempts for every task-system pair, including the revised Pivot task. The gateway token reached its own quota before the batch finished.

| System | Terminal receipts | Valid attempts | Successful attempts | Mean score | Valid task coverage |
| --- | ---: | ---: | ---: | ---: | ---: |
| Codex CLI + GPT-5.6 sol | 19 | 12 | 6 | 0.646 | 12 of 14 non-Pivot tasks |
| Claude Code + Opus 5 | 19 | 15 | 2 | 0.201 | 14 of 14 non-Pivot tasks |
| Qwen Code + Qwen3.8-max | 18 | 14 | 5 | 0.548 | 14 of 14 non-Pivot tasks |
| Revised Pivot task, all systems | 8 | 0 | 0 | Not available | 0 of 1 task |

The checkpoint contains 64 terminal receipts, 41 valid scored attempts, and 23 environment errors or cancellations. No task-system pair has eight valid attempts, so these results do not produce a standard pass@8 estimate. They are not pooled with the earlier batch because the Claude model version and gateway contract differ. The resumable summary is in [GROUP_GATEWAY_CHECKPOINT.csv](../results/GROUP_GATEWAY_CHECKPOINT.csv).

## Score interpretation

Each task has a weighted set of atomic requirements. The Judge returns a normalized score from 0 to 1 after checking the submitted workbook. A score of 0.70 or above counts as one successful attempt. It does not mean that 70% of cells are correct; the score combines task-specific requirements such as source selection, formulas, reconciliation, provenance, native objects, and preservation of unaffected regions.

For one system and one task, the empirical pass@1 estimate is `c / n`, where `c` is the number of successful attempts and `n` is the number of independent attempts. Standard pass@8 estimates the probability that at least one of eight independent attempts succeeds:

```text
pass@8 = 1 - C(n-c, 8) / C(n, 8)
```

The formula cannot be evaluated when `n < 8`. Codex currently has `n=4` per valid task; Claude and Qwen each have `n=1`. No standard per-task pass@8 is reported in this release.

## Track-level pattern

The track summary averages the current per-task mean scores. It describes the development runs and does not provide a statistical comparison among systems.

| Track | Valid tasks | Codex mean | Opus 4.8 mean | Qwen3.8-max mean | Current interpretation |
| --- | ---: | ---: | ---: | ---: | --- |
| A: professional models | 5 | 0.833 | 1.000 | 0.845 | The first three tasks are clearly easy in the current version; the statistics task needs validity review before more runs. |
| B: multi-file analysis | 4 of 5 represented | 0.277 | 0.281 | 0.161 | Current scores are lowest, but human review must separate genuine workflow difficulty from ambiguity or narrow Judge assumptions. The revised Pivot task was not part of this earlier batch. |
| C: documents to workbooks | 5 | 0.719 | 0.673 | 0.394 | Difficulty varies: the invoice is easy, receipts show a low-score signal, and the remaining tasks are mixed. |

## Remaining evidence for formal difficulty

The release requirement defines a difficult task candidate using four conditions: at least two of three target systems must have `pass@8 < 0.70`; valid attempts must average below 0.60; failures must reflect Agent capability rather than environment faults; and an independent CONFIRM sibling must show the same trend. A task still requires external human evidence before it can be described as accepted hard.

P15 has not completed that chain for the following reasons:

1. The per-task sample sizes are below the minimum needed to compute pass@8. The group-gateway batch stopped at 41 valid attempts; every task-system pair remains below eight valid samples. Difficult candidates were intended to receive up to 16 independent samples per system after the task and Judge were frozen.
2. The paid Claude and Qwen runs used gateways that did not prove a single locked upstream provider. They are genuine development runs, but the formal experiment still needs a fixed provider, model version, tool environment, and budget.
3. CONFIRM references and Oracles have passed self-consistency checks, but the target Agents have not yet run on the held-out CONFIRM siblings.
4. Six tasks show low-score signals, but no independent human reviewer has yet confirmed that the instruction, supplied information, and Judge admit reasonable professional solutions.
5. Four tasks are already too easy. Their revised versions must be treated as new task versions and tested again; scores from the current version cannot be inherited.
6. The Pivot task now has valid native objects on Microsoft Excel for Mac, but its group-gateway trials produced no valid scored attempt before the quota interruption. Windows Excel compatibility is also pending.
7. Real Agent outputs exposed legitimate Excel equivalents that early Judges rejected. The artifacts were regraded without increasing the attempt count, but formal sampling should begin only after the final Judge is frozen.
8. Windows Excel open/save/recalculate checks and external human review remain incomplete.

The current development labels therefore mean only:

- **Too easy:** the available evidence is already sufficient to justify a professional revision.
- **Preliminary difficult signal:** several systems scored poorly, but validity and failure attribution still require human review.
- **Mixed:** the current runs disagree or remain too sparse to justify changing the task.
Every locally valid task remains `INSUFFICIENT_EVIDENCE` for final difficulty.

## Task-by-task revision plan

Scores are shown as `Codex / Opus 4.8 / Qwen3.8-max`. Codex values are means over four attempts; the other two values come from one attempt each.

### Track A: professional models

| Task | Current scores | Development signal | Next professional revision | Decision rule |
| --- | ---: | --- | --- | --- |
| [DCF valuation](../tasks/pilot_v1/P15-A-FIN-DCF-001/) | 1.000 / 1.000 / 1.000 | Too easy | Add two operating segments with distinct revenue, margin, capital-expenditure, and working-capital drivers. Connect debt, cash, lease liabilities, scenarios, and valuation across sheets. | Revise at most three times. If it remains easy, activate the Project Finance reserve. |
| [Financial-model repair](../tasks/pilot_v1/P15-A-FIN-DEBUG-001/) | 1.000 / 1.000 / 1.000 | Too easy | Place one unique upstream defect in a monthly operating plan and propagate it through working capital, debt, and a covenant. Preserve a strict minimum-edit boundary. | After three unsuccessful revisions, activate the Forecast QC reserve. |
| [Cooling-water pump selection](../tasks/pilot_v1/P15-A-ENG-SIZING-001/) | 1.000 / 1.000 / 1.000 | Too easy | Add normal and peak duty cases, unit conversions, NPSH or temperature derating, and a parallel-pump constraint. The selected equipment must satisfy every declared condition. | After three unsuccessful revisions, activate the Mass Balance reserve. |
| [Paired sleep experiment](../tasks/pilot_v1/P15-A-STAT-EXPERIMENT-001/) | 0.385 / 1.000 / 0.667 | Preliminary difficult signal | Do not add difficulty yet. Review source preservation, QC exclusions, equivalent statistical formulas, and criterion weights. A later revision may add a prespecified exclusion rule, period effect, or missing-pair sensitivity analysis. | Freeze and expand only after a human reviewer confirms the task and Judge boundaries. |
| [Electricity-policy scenario](../tasks/pilot_v1/P15-A-POLICY-EIA-001/) | 0.780 / 1.000 / 0.560 | Mixed | Keep the current version for review. If later evidence shows it is easy, add a multiyear policy path, capacity limits, non-negative balancing generation, and scenario sensitivity. | Collect more frozen samples before changing the task. |

### Track B: multi-file analysis

| Task | Current scores | Development signal | Next professional revision | Decision rule |
| --- | ---: | --- | --- | --- |
| [Q2 data-version selection](../tasks/pilot_v1/P15-B-SALES-DISCOVERY-001/) | 0.580 / 0.160 / 0.160 | Preliminary difficult signal | First verify that the source-authority rules are clear. A later version may add partial restatements, effective dates, schema migration, and a coverage proof tied to control totals. | If the current task is valid, freeze and expand it before adding complexity. |
| [Order cleaning and joining](../tasks/pilot_v1/P15-B-OPS-CLEAN-JOIN-001/) | 0.096 / 0.038 / 0.038 | Preliminary difficult signal | Review whether the Judge accepts reasonable output layouts. A later version may add effective-dated master data, unit conversion, duplicate precedence, and several explicit exception classes. | Attribute current failures before spending on additional samples. |
| [Monthly ledger reconciliation](../tasks/pilot_v1/P15-B-FIN-RECON-001/) | 0.173 / 0.000 / 0.000 | Preliminary difficult signal | Do not make the current task harder. A later version may add split payments, credit notes, one-to-many matches, period cut-off, and a documented FX-source hierarchy. | Human review must first establish that zero scores reflect missing professional work rather than an overly narrow Judge. |
| [Native Pivot reporting](../tasks/pilot_v1/P15-B-PUBLIC-PIVOT-001/) | No valid Agent result in the interrupted group batch | Locally ready on Mac; Windows pending | Keep the native template-update task unchanged, complete Windows Excel compatibility readback, and resume the same frozen three-system batch. Canonicalize Agent workbooks through Excel before the final Judge readback. | Do not count quota, installation, or Excel-session failures as task failures. Start difficulty interpretation only after valid task-system samples exist. |
| [Heart-disease mortality brief](../tasks/pilot_v1/P15-B-HEALTH-REPORT-001/) | 0.259 / 0.926 / 0.444 | Preliminary difficult signal | Review whether alternative report layouts and evidence-linked prose are scored fairly. A later version may add demographic strata, suppression rules, confidence intervals, and cross-period comparability. | Complete human validity review before adding data volume. |

### Track C: documents to editable workbooks

| Task | Current scores | Development signal | Next professional revision | Decision rule |
| --- | ---: | --- | --- | --- |
| [Multipage invoice](../tasks/pilot_v1/P15-C-INVOICE-001/) | 1.000 / 1.000 / 0.621 | Too easy | Add true cross-page line continuations, tax and discount allocation, PO-line matching, duplicate-invoice control, and full amount closure. Keep every page readable. | After three unsuccessful revisions, activate the EOB reserve. |
| [Scanned quote](../tasks/pilot_v1/P15-C-QUOTE-001/) | 0.638 / 1.000 / 0.414 | Mixed | Keep the current version for review. A later version may add allowances, exclusions, several alternates, different tax bases, and revision-page precedence. | Expand the frozen task before revising it. |
| [Receipt batch](../tasks/pilot_v1/P15-C-RECEIPTS-001/) | 0.383 / 0.333 / 0.333 | Preliminary difficult signal | Do not reduce image quality. After validity review, a later version may add refunds, split tender, company-card matching, or a clearly specified multicurrency reimbursement rule. | Confirm image legibility and acceptable workbook structures before more runs. |
| [Purchase-order amendment](../tasks/pilot_v1/P15-C-PO-ADDENDUM-001/) | 0.758 / 0.633 / 0.400 | Mixed | Keep the current version. A later version may add two sequential addenda, effective dates, partial overrides, and delivery-schedule changes while preserving unaffected lines. | Expand the frozen task first; use the Change Order reserve only if a later version remains easy. |
| [Bank-statement reconciliation](../tasks/pilot_v1/P15-C-STATEMENT-001/) | 0.817 / 0.400 / 0.200 | Mixed | Keep the current version. A later version may add reversals, pending-versus-posted activity, fees, interest, running-balance checks, and continuation across more pages. | Expand the frozen task before revising it. |

## Difficulty mechanisms from existing benchmarks

The sources below use different tasks and metrics. Their absolute scores are not comparable with one another or with P15. The useful evidence is the mechanism behind the failures.

| Benchmark | Reported evidence | Main source of difficulty | P15 use |
| --- | --- | --- | --- |
| BlueFin | The strongest evaluated model remained below 50% overall. Formula correctness was substantially higher than correctness after input changes. | Hard-coded values, sign and date conventions, percentage units, beginning/end-period logic, and broken workbook integration. | Check formula relationships, model connections, protected regions, and downstream behavior after declared inputs change. |
| WTM-Bench | The best reported system reached 33.7% on its hard score and 54.0% on its soft score; Pivot soft scores were 0%–10%. | Multiple spreadsheet objects must be restored in the right locations and remain mutually consistent. | Rewind a correct workbook into a partial state and make formulas, charts, formatting, or Pivot objects depend on one another. |
| SpreadsheetBench 2 | Opus 4.6 reached 34.89% complete-task accuracy. Financial-model target cells were often correct while complete workbooks were not; debugging full-task accuracy was much lower than target-cell accuracy. | Long workbooks, hidden but unique root causes, incomplete inspection, and unintended edits outside the target region. | Increase dependency depth and check both the intended repair and preservation of unaffected regions. |
| FINCH | Reported human pass rate fell from 48.6% for one-stage work to 5.6% for workflows with four or more stages. | Several professional steps share the same business materials and must all be completed correctly. | Connect source discovery, cleaning, exception handling, reconciliation, and reporting when they form one real workflow. |
| DataGovBench | Executable code rates were much higher than task-success rates. | A pipeline can run while source selection, joins, transformations, or business outputs remain wrong. | Separate file validity from professional correctness and inject only task-relevant data defects. |
| FinBalance | The best reported final-statement accuracy was 46%, and the highest difficulty tier reached 0% final-statement accuracy in the reported evaluation. | Evidence selection, accounting treatment, provenance, ledger replay, and final reporting must agree. | Preserve document identity, cross-period rules, supporting evidence, and control totals; allow explicit unresolved exceptions. |
| VAREX | Frontier models exceeded 94% on the reported field-extraction metric. | Its main contribution is traceable construction from structured truth, not a demonstrated frontier-Agent difficulty ceiling. | Use typed truth to generate readable documents and preserve layout variation without relying on OCR degradation. |
| DocILE | Its line-item task scores fields and their grouping into the correct row and document structure. | A field can be individually correct but attached to the wrong document, page, or line item. | Score values, document identity, page provenance, and line-item grouping separately. |

Source versions, sections, project adaptations, and reuse boundaries are recorded in [LITERATURE_TO_DESIGN_MAP.md](../references/LITERATURE_TO_DESIGN_MAP.md) and [SOURCE_CATALOG.md](../references/SOURCE_CATALOG.md).

Across these studies, low performance is associated with several professional relationships that must remain correct at the same time. File count, unreadable images, hidden rules, irrelevant noise, tool faults, and model-specific traps do not provide valid difficulty.

## Revision and validation sequence

1. **Validate the low-score tasks.** Independent reviewers inspect the instruction, supplied information, acceptable equivalents, and Judge behavior before any claim that low scores reflect Agent capability.
2. **Revise the clearly easy tasks.** DCF, financial-model repair, pump selection, and invoice receive one professional revision at a time, with a maximum of three. A revised task receives a new version and new model evidence.
3. **Complete the Pivot compatibility evidence.** The native objects are verified on Microsoft Excel for Mac. Windows Excel must open, refresh, recalculate, save, and read back the same object chain before cross-platform validity is claimed.
4. **Freeze the experiment.** Task, Judge, system version, model identifier, provider, tools, budget, and sample count are fixed before formal sampling.
5. **Collect formal samples.** A valid task needs at least eight samples per system to compute pass@8; difficult candidates should preferentially receive 16.
6. **Run held-out CONFIRM siblings.** The same capability must show a similar pattern on independent data or templates.
7. **Complete Windows and human review.** Only these external checks can support native Excel validity and accepted professional difficulty.

## Review priorities

The next review should resolve the following decisions:

- whether the 15 professional work forms are distinct and useful enough to keep;
- whether the benchmark-authored instructions provide a credible professional scenario and sufficient information;
- whether the six low-score tasks fail for substantive capability reasons or because the task or Judge needs repair;
- whether each proposed revision adds real professional dependency rather than artificial friction;
- whether the formal Claude contract should retain Opus 4.8 or move to Opus 5 before additional paid sampling;
- which task families are strong enough to expand into independent DEV and CONFIRM instances.

## Cost record

The run framework estimates USD 38.79 for 56 valid Codex attempts. The OpenRouter account-usage increase for 29 paid runs is USD 15.81. These are separate accounting records and are not summed without invoice-level deduplication evidence.

## Claims supported now

- Fifteen primary task candidates and nine reserve designs have been assembled.
- All 15 tasks meet the current local answer-and-evaluator contract; the Pivot task also has native-object evidence from Microsoft Excel for Mac.
- Three target systems have produced real development workbooks.
- The group-gateway n=8 batch has 41 valid scored attempts and a resumable checkpoint; its incomplete sample is not presented as pass@8.
- The package, Judge, and result-collection path has been exercised.
- Construction practices that were exercised in P15 have been recorded for reuse.

## Claims not supported now

- All 15 tasks have completed Windows Excel and professional human validation.
- The 15 tasks are accepted as hard.
- Standard per-task pass@8 has been estimated.
- Windows Excel validation is complete.
- External human review is complete.
- DEV results reproduce on held-out CONFIRM siblings.
- The portfolio represents the full distribution of professional Excel work.
