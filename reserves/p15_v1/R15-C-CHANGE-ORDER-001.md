---
id: R15-C-CHANGE-ORDER-001
track: C
status: DESIGN_FROZEN
package_not_built: true
harbor_package_built: false
primary_eligible: false
split_group: R15-C-CHANGE-ORDER-001
---

# Construction change-order commercial review

## Professional work and user scenario

An owner-side project controls analyst must review a synthetic contractor change-order package and update the contract-control workbook. The source combines a change-order form, labor detail, material quotations, equipment logs, subcontractor proposals, tax, overhead and profit, bond and insurance, retention, prior approved changes, and a requested schedule extension.

## Instruction draft

“Review the supplied change-order package and complete the owner control workbook. Capture the change identifier, scope, reason, status, contract reference, cost lines, schedule request, and page provenance. Recalculate labor, material, equipment, subcontract, tax, overhead and profit, bond and insurance, retention, and net payable under the visible commercial rules; distinguish proposed, eligible, excluded, and unsupported amounts. Reconcile contractor request to the independently evaluated amount, update original contract, prior changes, pending change, revised contract, committed and payable views, and identify cost, documentation, duplication, and schedule exceptions. Preserve every source line and do not treat approval fields as approved unless printed.”

## Packaged-input design and authenticity boundary

- Planned inputs: change_order_packet.pdf, contract_control.csv, commercial_rules.md, cost_code_lookup.csv, and starting_workbook.xlsx.
- The multipage synthetic packet will contain a cover form, labor tickets, vendor quote, equipment log, subcontractor proposal, prior-change reference, markup summary, signatures with mixed statuses, and a schedule narrative.
- The visible rules will define eligible hours and rates, material taxability, equipment caps, subcontract treatment, markup sequence and bases, bond and insurance base, retention treatment, rounding, change status, and revised-contract presentation.
- The contract control will state original contract, approved changes, pending changes, paid-to-date, retention-to-date, and schedule baseline.
- All parties, project names, signatures, contracts, costs, and dates will be synthetic. Nothing will be described as a real approval, commitment, payment, or legal determination.
- No PDF, workbook, or fixture has been generated at DESIGN_FROZEN status.

## Target workbook and professional operations

Required sheets are Change_Header, Cost_Detail, Commercial_Rules, Evaluation, Contract_Control, Schedule_Impact, Exceptions, and Checks. The workbook must retain each source cost once, validate quantity times rate, apply eligibility and exclusions, reconstruct tax and nested markups on the stated bases, treat retention as a withholding rather than project cost, compare requested with evaluated amount, update contract views without promoting pending value to approved, assess requested schedule days, and reconcile all controls with source provenance.

## Independent oracle design

A task-local oracle will use verifier-private document truth and the visible commercial rules to independently calculate line extensions, rate caps, exclusions, taxable base and tax, overhead and profit layers, bond and insurance, evaluated gross change, retention withholding, net payable, request-to-evaluation variance, contract roll-forward, and schedule effects. It will compare by change ID, source document, cost line, cost code, and page rather than cell position.

## Atomic rubric design

- R001: Required sheets exist and header, status, contract, scope, reason, and schedule fields match the source.
- R002: Every labor, material, equipment, and subcontract line is captured exactly once with source provenance.
- R003: Quantities, hours, units, rates, and line extensions are arithmetically correct.
- R004: Eligible, excluded, capped, duplicated, and unsupported amounts follow the visible commercial rules.
- R005: Material tax and taxable base are calculated correctly without taxing excluded or nontaxable lines.
- R006: Overhead and profit use the correct category bases, rates, sequence, and rounding.
- R007: Bond and insurance are applied to the stated base without markup duplication.
- R008: Retention is calculated as a withholding and net payable reconciles to evaluated gross change.
- R009: Contractor request, independent evaluation, and variance reconcile by category and in total.
- R010: Original contract, approved changes, pending change, revised-contract presentation, and payable views preserve status semantics.
- R011: Requested schedule days, owner-evaluated days, and documentation exceptions match the printed evidence.
- R012: All cost, contract, payable, and schedule controls reconcile and remain formula-linked.
- P001: Penalize silent approval, duplicate cost capture, incorrect nested markup, retention added as cost, or unsupported schedule conclusions.

## Acceptable materially different equivalent

One solution may use a normalized source-cost ledger with rule attributes and a formula-driven evaluation bridge. A materially different equivalent may keep labor, materials, equipment, and subcontracts in separate calculation blocks, then consolidate them into markup, contract, and payable schedules. Row order, helper columns, and sheet order may differ; eligible bases, evaluated amount, status, contract roll-forward, retention, schedule result, and provenance must agree.

## Negative fixtures

- No-op: starting workbook returned with contract controls unchanged and evaluation sheets blank.
- Malformed: unreadable workbook, missing cost or contract-control sheets, or invalid numeric results.
- M1: Overhead and profit are both applied to the fully marked-up subtotal, producing unauthorized compounding.
- M2: Tax is applied to labor, excluded costs, and a second time after markup.
- M3: Retention is added to the change-order cost instead of withheld from gross payable.
- M4: Pending change value overwrites original contract or is presented as already approved.
- M5: Labor tickets and the contractor summary are both captured as independent cost lines, duplicating labor.
- M6: Requested schedule days are accepted without evaluating the printed concurrent-delay evidence or preserving source page.

## Two meaningful input perturbations

- P1: Increase one eligible labor ticket by 12 hours at its unchanged approved hourly rate. Labor direct cost, applicable overhead and profit, bond and insurance, evaluated gross change, retention, net payable, request variance, and contract pending view must propagate exactly; material tax, material quantities, equipment, original contract, prior approved changes, and schedule days remain invariant.
- P2: Increase one taxable material quotation by USD 5,000 with quantity and all other categories unchanged. Material direct cost, taxable base, tax, applicable markups, bond and insurance, evaluated amount, retention, net payable, and request variance must change by the visible nested rules; labor, equipment, subcontract, original contract, prior changes, approval status, and schedule-day result remain invariant.

## Verifier-private CONFIRM sibling

The verifier-private CONFIRM sibling will use a different synthetic project, contract, change ID, scope, parties, cost-code mix, markup rates, tax treatment, prior changes, status evidence, source layout, and schedule narrative. It preserves the line capture, commercial-rule, contract-control, payable, schedule, and provenance capability contract. Confirm source, truth, oracle, perturbations, and reference remain outside Agent-visible input.

## Difficulty-change policy

Allowed difficulty changes: a disclosed category-specific markup, an equipment rate cap, a taxable versus nontaxable material split, one prior-change duplication check, or concurrent schedule evidence.

Forbidden difficulty changes: hidden commercial rules, illegible signatures or numbers, legal conclusions not supported by the packet, real project data, irrelevant page volume, corrupt interfaces, or adversarial layout tuned to one model.

## Activation rule

Activate only if a Track C PRIMARY cannot reach validity, remains VALID_BUT_EASY after three real professional revisions, or senior reviewers choose change-order commercial review as the more distinct capability. Activation requires a complete newly constructed and validated package; DESIGN_FROZEN does not imply benchmark validity, difficulty, or acceptance.

## Build state

package_not_built: true. No task directory, PDF, workbook, fixture, Judge, Harbor wrapper, Agent attempt, Windows Excel receipt, or human-review decision exists for this reserve.
