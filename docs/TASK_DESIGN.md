# Task Design

## Portfolio logic

P15 first separates Excel work into three tracks and then selects five distinct work forms within each track. The five-form structure prevents the portfolio from becoming a set of renamed templates.

Track A covers model creation, model repair, constrained engineering choice, statistical analysis, and policy scenarios. Track B covers source/version selection, cleaning and joining, reconciliation, native Pivot reporting, and evidence-based narrative reporting. Track C covers multipage invoices, quote scope, receipt batches, document amendments, and continuation statements.

Each task must differ in its professional action, input structure, delivery object, independent oracle, and likely failure modes.

The selection review covers seven requirements:

1. A defined professional decision or handoff supported by the workbook.
2. A defined set of information that the Agent must discover, calculate, reconcile, or preserve.
3. A workbook-level deliverable beyond a final scalar answer.
4. Independent recomputation from professional rules.
5. At least one materially different acceptable solution.
6. Task-specific professional errors that the Judge must reject.
7. A work form that adds portfolio coverage instead of repeating an existing template.

The 15-task portfolio is therefore a capability sample, not a balanced industry survey. Industry labels provide context; the unit of diversity is the professional work form.

## Selection rationale for the 15 tasks

The portfolio uses five different work forms in each track. Each row changes the professional action, the information relationship that must be preserved, and the workbook that another person would receive.

| Track | Task | Distinct professional action | Relationship the workbook must preserve |
| --- | --- | --- | --- |
| A | [DCF valuation](../tasks/pilot_v1/P15-A-FIN-DCF-001/) | Build a forward-looking valuation model from assumptions and historical facts | Forecast drivers, free cash flow, terminal value, equity bridge, and sensitivity must recalculate together. |
| A | [Financial-model repair](../tasks/pilot_v1/P15-A-FIN-DEBUG-001/) | Find and repair one upstream root cause | The smallest valid edit must restore downstream balances without changing unaffected regions. |
| A | [Cooling-water pump selection](../tasks/pilot_v1/P15-A-ENG-SIZING-001/) | Convert engineering conditions into an equipment choice | Units, hydraulic calculations, safety margin, and catalog eligibility must support the selected device. |
| A | [Paired sleep experiment](../tasks/pilot_v1/P15-A-STAT-EXPERIMENT-001/) | Analyze a paired crossover experiment | Participant pairing, QC exclusions, estimates, uncertainty, power, and the chart must use the same analysis population. |
| A | [Electricity-policy scenario](../tasks/pilot_v1/P15-A-POLICY-EIA-001/) | Translate a policy assumption into a balanced scenario | Public source data, generation balance, policy changes, and emissions intensity must remain internally consistent. |
| B | [Q2 data-version selection](../tasks/pilot_v1/P15-B-SALES-DISCOVERY-001/) | Identify the authoritative release before analysis | Version status, schema, period coverage, metrics, and the coverage proof must point to the same release. |
| B | [Order cleaning and joining](../tasks/pilot_v1/P15-B-OPS-CLEAN-JOIN-001/) | Normalize records and join them to master data | Duplicate handling, normalized keys, valid matches, exceptions, and control totals must reconcile. |
| B | [Monthly ledger reconciliation](../tasks/pilot_v1/P15-B-FIN-RECON-001/) | Match accounting records and explain the residual difference | Ledger lines, subledger records, FX policy, approved adjustments, exceptions, and the variance bridge must close. |
| B | [Native Pivot reporting](../tasks/pilot_v1/P15-B-PUBLIC-PIVOT-001/) | Build a refreshable native Excel analysis object | Source range, PivotCache, PivotTable, filters, SUM measures, refresh behavior, and PivotChart must stay connected. |
| B | [Heart-disease mortality brief](../tasks/pilot_v1/P15-B-HEALTH-REPORT-001/) | Turn source data into an evidence-linked management brief | Calculations, chart, narrative findings, and source references must describe the same population and period. |
| C | [Multipage invoice](../tasks/pilot_v1/P15-C-INVOICE-001/) | Convert a continued invoice into an editable workpaper | Document identity, page continuation, line items, taxes, totals, and provenance must close. |
| C | [Scanned quote](../tasks/pilot_v1/P15-C-QUOTE-001/) | Normalize scope and commercial terms from a quote | Base scope, discounts, tax, alternates, exclusions, and source locations must remain distinct. |
| C | [Receipt batch](../tasks/pilot_v1/P15-C-RECEIPTS-001/) | Consolidate several document layouts into one batch | Receipt identity, item hierarchy, categories, tender totals, and the batch summary must reconcile. |
| C | [Purchase-order amendment](../tasks/pilot_v1/P15-C-PO-ADDENDUM-001/) | Apply an addendum to an existing order | Amendment precedence, changed lines, new lines, quantities, prices, and preserved original terms must agree. |
| C | [Bank-statement reconciliation](../tasks/pilot_v1/P15-C-STATEMENT-001/) | Reconstruct activity across statement pages | Account identity, transaction order, debit/credit direction, running balance, and page provenance must close. |

This selection logic prevents three common forms of accidental duplication: changing only an industry label, changing only the source layout, or changing only the numbers while the Agent performs the same operation.

## Construction sequence

1. Select a real professional work product and decision.
2. Confirm source identity, usage rights, and information sufficiency.
3. Establish structured truth from professional rules.
4. Build one complete reference workbook.
5. Rewind, render, or inject a controlled defect to create the Agent-visible input.
6. Write the instruction in a realistic work context with a precise output contract.
7. Implement an independent oracle, atomic rubric, and deterministic Judge.
8. Test the Judge with a reference, materially different equivalent, no-op, malformed workbook, and task-specific professional errors.
9. Run target Agents, then obtain Windows Excel and human review evidence.

## Track-specific construction

### Track A

Start from a recalculable model, a controlled professional case, or a workbook with one unique root-cause defect. Preserve the assumptions and protected regions needed to solve the task. Recompute key outputs independently and perturb declared inputs to check downstream propagation.

### Track B

Start from multiple files with meaningful versions, join keys, tolerances, exceptions, and control totals. The oracle replays source selection, normalization, joining, aggregation, and reconciliation from record-level truth. Irrelevant clutter and missing rules are not valid difficulty mechanisms.

### Track C

Start from typed structured truth, render the truth into controlled PDFs or images, and preserve document/page/line-item identity. The oracle checks values and relationships against the pre-render truth rather than attempting to infer the answer from the rendered document.

## Task families and instances

A **task family** is one professional work pattern plus its shared source, template, and business rules. “Foreign-currency ledger reconciliation” is a task family. One concrete set of ledgers, FX rates, adjustments, and delivery requirements is a task instance.

A DEV instance supports construction and debugging. A CONFIRM sibling uses independent data or templates to test the same capability. Directory structure and validation patterns may be reused across a family; source rights, truth, natural difficulty, and human validity must be established for every new instance.

## Difficulty policy

Difficulty may come from longer dependency chains, multiple valid operating conditions, legitimate cross-file relationships, visible version hierarchies, exception handling, reconciliation, and professional constraints. It may not come from missing necessary information, unreadable images, unrelated-file volume, interface defects, or model-specific traps.

An easy task may receive at most three professional revisions. If it remains easy, activate a reserve from the same capability area or retain the task as valid-but-easy.

## Worked package

The [monthly ledger reconciliation package](../tasks/pilot_v1/P15-B-FIN-RECON-001/) shows the complete chain:

1. [`instruction.md`](../tasks/pilot_v1/P15-B-FIN-RECON-001/instruction.md) defines the professional request.
2. [`data/input_files/`](../tasks/pilot_v1/P15-B-FIN-RECON-001/data/input_files/) contains the Agent-visible ledgers, FX policy, and adjustment evidence.
3. [`metadata/oracle_recompute.py`](../tasks/pilot_v1/P15-B-FIN-RECON-001/metadata/oracle_recompute.py) recomputes the correct reconciliation.
4. [`solution/reference.xlsx`](../tasks/pilot_v1/P15-B-FIN-RECON-001/solution/reference.xlsx) is one complete delivery.
5. [`rubric.json`](../tasks/pilot_v1/P15-B-FIN-RECON-001/rubric.json) defines atomic requirements.
6. [`tests/evaluate.py`](../tasks/pilot_v1/P15-B-FIN-RECON-001/tests/evaluate.py) scores a candidate workbook.
7. [`fixtures/`](../tasks/pilot_v1/P15-B-FIN-RECON-001/fixtures/) contains the equivalent, no-op, malformed, and semantic-error cases.

## Reuse boundary

The repository can reuse package layout, metadata fields, Oracle/Golden separation, rubric schema, Judge entry points, negative-fixture patterns, Harbor packaging, result collection, and review forms.

Every new task family must re-establish the professional objective, source identity and rights, information sufficiency, typed truth, domain oracle, acceptable equivalents, natural difficulty, and human validity. Reuse is a construction shortcut, not evidence that a new task is correct.
