---
id: R15-C-UTILITY-BILL-001
track: C
status: DESIGN_FROZEN
package_not_built: true
harbor_package_built: false
primary_eligible: false
split_group: R15-C-UTILITY-BILL-001
---

# Commercial utility bill validation and charge reconstruction

## Professional work and user scenario

An energy manager must validate a synthetic commercial electric bill against printed meter reads and a visible tariff summary. The bill contains time-of-use consumption, billing demand, meter multipliers, demand ratchets, energy and demand charges, riders, taxes, prior balance and payment activity, and a current amount due. The professional goal is a transparent charge reconstruction plus actionable exceptions.

## Instruction draft

“Extract the supplied commercial utility bill into the workbook and reconstruct the current-period charges. Preserve account, service period, meter, rate class, page, and line provenance. Calculate usage from printed reads and multipliers, separate on-peak, off-peak, and total energy, determine billing demand under the visible ratchet rule, apply the printed energy, demand, customer, rider, and tax rates to their stated bases, and reconcile reconstructed current charges to the bill. Keep prior balance and payment activity separate from current charges, flag discrepancies and unusual drivers, and leave all calculations formula-linked.”

## Packaged-input design and authenticity boundary

- Planned inputs: utility_bill.pdf, tariff_summary.md, charge_code_lookup.csv, and starting_workbook.xlsx.
- The two-page synthetic bill will contain two meters, current and previous reads, multipliers, time-of-use quantities, current peak demand, a prior-demand ratchet reference, detailed current charges, payments, and total amount due.
- The visible tariff will state rate units, billing-demand rule, rider bases, tax base, rounding sequence, customer charge, and the treatment of prior balance and payments.
- The lookup will expand only charge codes printed on the bill; it will not contain hidden expected totals.
- All utility, account, premises, meter, rate, dates, and values will be authored synthetic fixtures and will not be presented as a real bill or customer record.
- No PDF, workbook, or fixture has been generated at DESIGN_FROZEN status.

## Target workbook and professional operations

Required sheets are Bill_Fields, Meter_Usage, Tariff, Charge_Reconstruction, Account_Activity, Variance_Review, and Checks. The workbook must extract printed fields, calculate meter usage, aggregate time-of-use energy, apply the billing-demand rule, reconstruct each charge on the correct base and rounding step, separate current charges from balance-forward activity, reconcile current charges and amount due, and provide page and line provenance plus variance flags.

## Independent oracle design

A task-local oracle will consume verifier-private structured bill truth and the visible tariff to independently calculate read differences times multipliers, time-of-use totals, billing demand, energy charges, demand charges, fixed charges, rider bases, taxes, reconstructed current charges, and account-balance roll-forward. It will compare facts by account, meter, period, charge code, and printed line rather than workbook cell address.

## Atomic rubric design

- R001: Required sheets exist and all printed account, service-period, meter, and rate-class identifiers are captured.
- R002: Previous and current meter reads and multipliers match the document.
- R003: Meter usage equals read difference times multiplier and reconciles to printed time-of-use quantities.
- R004: On-peak, off-peak, and total kWh are classified and aggregated correctly.
- R005: Billing demand uses the maximum required by the visible current-demand and ratchet rule.
- R006: Energy, demand, and fixed customer charges use the correct units, rates, and bases.
- R007: Riders and taxes apply to the stated bases and follow the printed rounding convention.
- R008: Reconstructed current charges reconcile to the printed current-charge total within tolerance.
- R009: Prior balance, payments, adjustments, current charges, and amount due form a correct account roll-forward.
- R010: Variance and unusual-driver flags identify genuine discrepancies or material charge drivers without altering totals.
- R011: Every extracted and reconstructed line retains correct page and charge-line provenance.
- P001: Penalize stale totals, omitted demand charges, wrong meter multipliers, taxes on the wrong base, or mixing balance-forward activity into usage cost.

## Acceptable materially different equivalent

One solution may normalize every meter and charge line into tables that feed a reconstruction summary. A materially different equivalent may keep a meter-read schedule and tariff calculation blocks separate, then bridge printed to reconstructed charges by charge code. Layout, row order, and helper formulas may differ; usage, billing demand, charges, account roll-forward, exceptions, and provenance must agree.

## Negative fixtures

- No-op: starting workbook returned without source extraction or charge reconstruction.
- Malformed: unreadable workbook, missing Meter_Usage or Charge_Reconstruction sheets, or invalid numeric outputs.
- M1: Usage is taken as current read alone rather than current minus previous read times multiplier.
- M2: All kWh are charged at the on-peak rate and the time-of-use split is ignored.
- M3: Demand charge is omitted or calculated from kWh instead of billing kW.
- M4: Billing demand uses only current peak and ignores the visible ratchet minimum.
- M5: A rider is applied to total amount due rather than its stated energy-and-demand base.
- M6: Tax is calculated before exclusions or with prior balance included, while the printed total is hard-coded to hide the error.

## Two meaningful input perturbations

- P1: Increase one meter on-peak current read so that on-peak usage rises by exactly 1,000 kWh after its multiplier, with demand unchanged. On-peak and total usage, the matching energy charge, applicable riders and taxes, current charges, and amount due must propagate; off-peak usage, billing demand, demand charge, prior balance, payments, and meter identity remain invariant.
- P2: Increase the printed current peak from 420 kW to 465 kW where the ratchet floor is 440 kW. Billing demand becomes 465 kW, and demand charge, demand-based riders, tax, current charges, and amount due change by exact tariff rules; meter kWh, energy charges, customer charge, prior balance, and payments remain invariant.

## Verifier-private CONFIRM sibling

The verifier-private CONFIRM sibling will use a different synthetic utility, account, premises, two new meters, rate class, service dates, ratchet history, time-of-use mix, rates, riders, taxes, page layout, and account activity. It preserves the extraction, usage, tariff, reconstruction, roll-forward, and provenance contract. Confirm source, truth, oracle, perturbations, and reference stay outside Agent-visible input.

## Difficulty-change policy

Allowed difficulty changes: a disclosed second meter multiplier, one demand ratchet, one legitimate rider base, seasonal tariff rates, or an explicit printed rounding sequence.

Forbidden difficulty changes: unreadable scans, missing tariff facts, unit ambiguity, hidden charge bases, excessive decorative noise, corrupted PDFs, unrelated pages, or adversarial formatting aimed at one system.

## Activation rule

Activate only if a Track C PRIMARY cannot be repaired to validity, remains VALID_BUT_EASY after three authentic complexity revisions, or senior review selects utility-bill reconstruction as a more distinct professional document task. Activation starts a full build and evidence cycle; no evidence transfers from the design-only card.

## Build state

package_not_built: true. No task directory, PDF, workbook, fixture, Judge, Harbor wrapper, Agent result, Excel validation, or reviewer decision exists for this reserve.
