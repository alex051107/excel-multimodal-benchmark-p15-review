---
name: excel-benchmark-track-c-construction
description: Construct and repair document-to-Excel benchmark tasks that require grounded extraction, provenance, reconciliation, and formula-linked outputs from PDFs or images.
---

> Validated against P15 checkpoint: `8df20f98366633a4813cf6b98ee7e78270e3bdb7`
> Last evidence refresh: `2026-08-30`

# Purpose

Build a Track C package whose difficulty comes from authentic document interpretation and spreadsheet semantics. Keep document truth verifier-private and make every visible input sufficient for a competent worker.

# Required inputs

- A rights-clear PDF or image family and a benchmark-authored professional request.
- An Agent-visible starter workbook, context note, and document inventory.
- A verifier-private typed truth model with document, page, row identity, and expected calculations.
- A distinct CONFIRM sibling from the same capability family.

# Exact procedure

1. Choose one professional workflow: invoice, quote, receipts, purchase-order revision, or statement reconciliation. Do not obtain diversity by relabeling one template.
2. Render and visually inspect every source page. Resolve cross-page continuation, revision precedence, grouped lines, signs, currencies, and totals before creating the workbook.
3. Give the Agent only the source documents, starter, context, and output path. Copy that exact public input set to `/app/input/` in the Agent image.
4. Define typed private truth by content identity rather than fixed output row number. Require a readable locator containing the correct file and page; allow materially different non-empty row IDs that join output rows to provenance.
5. Author a formula-linked reference and an equivalent that differs materially in formula or row-ID implementation while preserving semantics.
6. Create no-op, malformed, and at least four single-defect semantic mutants tied to the workflow, such as wrong revision precedence, false page provenance, dropped receipt, or incorrect debit sign.
7. Declare at least two source perturbations and exact propagation plus protected-source expectations. Freeze a distinct CONFIRM sibling before Agent difficulty testing.
8. Keep DEV truth, rubric, oracle, and perturbation assets under the separate verifier context. The Agent Dockerfile must not copy `tests/`, `private/`, `solution/`, or fixtures.

# Required outputs

- Complete Harbor task package with instruction, public inputs, reference, equivalent, negative fixtures, atomic rubric, deterministic Judge, task metadata, DEV oracle, and CONFIRM sibling.
- Render/source inspection evidence and local 5-run Judge receipts.
- Explicit Windows Excel status; no inferred canonical evidence.

# Validation checks

- Source pages and all reference sheets were visually inspected.
- Reference is 1.0; equivalent is at least 0.95; no-op is below 0.30; malformed is 0; every semantic mutant is below 0.70 for five identical runs.
- Two DEV perturbations propagate exactly without changing declared protected source fields.
- CONFIRM input and reference hashes differ from DEV and its reference scores 1.0 five times.
- A real Harbor Oracle/NOP/malformed smoke produces the expected rewards without verifier exceptions.
- Agent image inspection finds no private truth or answer-bearing file.

# Allowed difficulty changes

- Add realistic cross-page continuation, cross-document identity, revision or addendum precedence, grouped lines, totals reconciliation, or an original scan/photo variation.
- Strengthen document provenance and formula-linked reconciliation when the user-visible evidence remains sufficient.

# Forbidden difficulty changes

- Unreadable noise, missing necessary pages, arbitrary distractor files, hostile filenames, unsupported interfaces, token throttling, or model-specific traps.
- Exact private locator-string matching when file, page, and readable content identity are already correct.
- Copying truth, rubric, oracle, CONFIRM, or construction skills into Agent-visible inputs.

# Common failure modes

- The starter already contains answer rows or totals.
- The Judge matches output row numbers instead of normalized content identity.
- Provenance accepts the correct value with a false page, or rejects a valid equivalent locator.
- The Agent image copies the task root and leaks `private/`.
- A document mutant changes many unrelated workbook formulas and collapses to an uninformative zero.

# Stop conditions

Stop construction when rights or source sufficiency are unclear, the private truth cannot be independently replayed, the starter leaks the answer, or a native spreadsheet requirement cannot be verified. Mark the task invalid rather than filling the gap with documentation.

# Examples from verified P15 work

- `P15-C-INVOICE-001` matched four line items by normalized description, accepted candidate-defined joined row IDs, and required filename/page provenance.
- `P15-C-PO-ADDENDUM-001` tested revised quantities and prices without treating the earlier PO as the final authority.
- All five P15 Track C Agent images were rebuilt to copy only `input/`; Docker inspection found public inputs under `/app/input` and no private truth.
- Invoice, quote, receipts, PO addendum, and statement references passed DEV and distinct CONFIRM Judges at 1.0; task-specific mutants ranged below 0.70.
