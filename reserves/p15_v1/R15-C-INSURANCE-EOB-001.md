---
id: R15-C-INSURANCE-EOB-001
track: C
status: DESIGN_FROZEN
package_not_built: true
harbor_package_built: false
primary_eligible: false
split_group: R15-C-INSURANCE-EOB-001
---

# Health-insurance explanation-of-benefits reconciliation

## Professional work and user scenario

A benefits-accounting analyst must convert a synthetic multipage explanation of benefits into a service-line workbook that supports member inquiry and claim-payment reconciliation. The document mixes claim headers, service dates, procedure descriptions, billed and allowed amounts, contractual adjustments, deductible, copay, coinsurance, plan payment, member responsibility, and denial messages across pages.

## Instruction draft

“Read the supplied explanation of benefits and complete the workbook at service-line level. Capture each claim and service line once, preserve page and section provenance, distinguish billed, noncovered or adjusted, allowed, plan-paid, deductible, copay, coinsurance, other member responsibility, and total member responsibility, and retain denial or remark codes. Build claim-level and statement-level reconciliations, flag any line that does not satisfy the visible arithmetic conventions, and produce a concise member-inquiry view. Treat the document as an explanation, not an invoice, and do not infer facts that are not printed.”

## Packaged-input design and authenticity boundary

- Planned inputs: eob_packet.pdf, eob_field_conventions.md, claim_status_lookup.csv, and starting_workbook.xlsx.
- The PDF will contain three synthetic claims, multiple service lines, a continuation page, repeated claim context, remark codes, a not-a-bill notice, and one denied line.
- The visible convention will define service-line identity, sign handling, printed rounding, allowed-amount reconciliation, member-responsibility components, claim subtotal treatment, and how continuation context is inherited.
- The lookup will translate only printed claim and remark codes; it will not provide hidden monetary truth.
- All member, provider, plan, claim, and service information will be synthetic and clearly labeled as a benchmark fixture, with no real protected health information or insurer decision.
- No PDF, workbook, or fixture has been generated at DESIGN_FROZEN status.

## Target workbook and professional operations

Required sheets are Instructions, Service_Lines, Claims, Member_Inquiry, Exceptions, and Checks. The workbook must transcribe every service line once, link service lines to the correct claim even across pages, separate monetary fields and responsibility components, preserve denial and remark evidence, reconcile billed less adjustments to allowed where the printed convention applies, reconcile allowed to plan payment plus member responsibility, and tie service lines to claim and packet totals.

## Independent oracle design

A task-local oracle will use a verifier-private structured representation independently created from the frozen source-document specification. It will compare claim and service-line identities by printed content, calculate all visible arithmetic equations with stated rounding, aggregate claim and packet totals, verify status and code interpretations, and score provenance by printed page and section. The oracle will not depend on workbook coordinates or OCR output from the candidate.

## Atomic rubric design

- R001: Required sheets exist and every printed service line is represented exactly once.
- R002: Claim numbers, service dates, provider labels, procedure descriptions, and statuses match the source document.
- R003: Billed, adjustment or noncovered, allowed, and plan-paid amounts are transcribed into the correct fields.
- R004: Deductible, copay, coinsurance, and other member-responsibility amounts remain distinct.
- R005: Each service line satisfies the printed allowed-amount and responsibility equations within the stated rounding rule.
- R006: Denied and partially paid lines retain the correct printed remark or denial codes and explanations.
- R007: Continuation-page service lines are assigned to the correct claim without repeating claim totals.
- R008: Claim-level totals agree with their service lines and printed claim summaries.
- R009: Packet-level totals reconcile across all claims without counting headers or subtotals as service lines.
- R010: Member_Inquiry accurately summarizes what the plan paid, what the member may owe, and which items need follow-up.
- R011: Each captured fact includes correct page and source-section provenance.
- P001: Penalize treating billed as allowed, conflating member and plan payment, duplicate subtotal capture, unsupported medical inference, or fabricated values.

## Acceptable materially different equivalent

One solution may use a normalized service-line table and formula-driven claim summaries. A materially different equivalent may create claim-specific extraction blocks with a separate reconciliation ledger and consolidated member view. Row order, helper columns, and code-description placement may differ; claim identity, money fields, arithmetic, statuses, totals, and provenance must agree.

## Negative fixtures

- No-op: starting workbook returned with the source packet untouched and all extraction tables blank.
- Malformed: missing service-line sheet, unreadable workbook, or amounts stored as unusable text without reconciliations.
- M1: Billed amount is copied into the allowed field for every paid line.
- M2: Member responsibility is recorded as plan payment, reversing payer and member shares.
- M3: Contractual adjustment is omitted, so billed-to-allowed reconciliation fails.
- M4: Printed claim subtotals are entered as additional service lines and double counted.
- M5: Deductible and coinsurance are swapped while their total is preserved.
- M6: A continuation-page line is assigned to the neighboring claim and given the wrong page provenance.

## Two meaningful input perturbations

- P1: Increase the printed allowed amount for one paid service line by USD 60.00 and increase plan payment by USD 48.00 and coinsurance by USD 12.00, with billed amount and other fields unchanged. The line, claim, packet, plan-paid, member-responsibility, and coinsurance totals must propagate exactly; billed, deductible, copay, other claims, status, and provenance remain invariant.
- P2: Reallocate USD 35.00 on one service line from plan payment to deductible while keeping allowed amount and total member-plus-plan allocation unchanged. Plan-paid decreases and deductible and member responsibility increase by USD 35.00 at line, claim, and packet levels; allowed amount, billed amount, adjustment, claim identity, and packet allowed total remain invariant.

## Verifier-private CONFIRM sibling

The verifier-private CONFIRM sibling will use a newly rendered EOB from a different synthetic plan and member, with different claims, page layout, continuation pattern, amounts, codes, responsibility mix, and one different denial case. It preserves the service-line extraction, arithmetic, aggregation, and provenance contract. Confirm source, truth, oracle, perturbations, and reference remain verifier-private and are not present in Agent-visible DEV input.

## Difficulty-change policy

Allowed difficulty changes: a legitimate continuation page, an additional responsibility component, one partial denial, a second claim-summary format, or a visible rounding convention.

Forbidden difficulty changes: illegible rendering, hidden monetary values, real health information, ambiguous line identity, gratuitous page noise, broken file access, medical-necessity inference, or a single-model visual trap.

## Activation rule

Activate only if a Track C PRIMARY is invalid, remains VALID_BUT_EASY after three professional revisions, or reviewers identify EOB reconciliation as the stronger distinct document-to-workbook capability. Activation requires source generation, private truth, reference, Judge, package, perturbation, Excel, Harbor, model, and review validation from scratch.

## Build state

package_not_built: true. No task directory, PDF, workbook, fixture, Judge, Harbor wrapper, model sample, Excel receipt, or human-review evidence exists for this reserve.
