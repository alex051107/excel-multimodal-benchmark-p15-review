---
id: R15-B-INVENTORY-RECON-001
track: B
status: DESIGN_FROZEN
package_not_built: true
harbor_package_built: false
primary_eligible: false
split_group: R15-B-INVENTORY-RECON-001
---

# Distribution-center inventory reconciliation

## Professional work and user scenario

A distribution-center controller must reconcile a month-end warehouse-management snapshot to the ERP inventory subledger while incorporating blind cycle counts, receipts and shipments around cutoff, in-transit ownership, quarantine stock, and standard cost. The deliverable supports a close adjustment and an operations exception queue, not merely a two-column lookup.

## Instruction draft

“Complete the month-end inventory reconciliation in the supplied workbook. Match stock by item, location, lot, and ownership where those dimensions apply; retain unmatched records from either system. Apply the visible cutoff and in-transit rules, compare book quantities with validated physical counts, value each explained and unexplained difference at the effective standard cost, and produce a posting-ready adjustment view plus a prioritized exception queue. Reconcile units and value from source totals through adjustments to the final difference, preserve source keys and reason evidence, and leave all derived results formula-linked.”

## Packaged-input design and authenticity boundary

- Planned inputs: wms_on_hand.csv, erp_inventory.csv, cycle_counts.csv, cutoff_movements.csv, in_transit.csv, standard_costs.csv, reconciliation_policy.md, and starting_workbook.xlsx.
- WMS rows will carry item, site, bin, lot, status, ownership, quantity, and snapshot timestamp; ERP rows will carry item, site, lot, ownership, account, quantity, and value.
- The visible policy will define bin-to-location mapping, owned versus consigned treatment, count precedence, cutoff timestamps, in-transit title passage, reason codes, and materiality thresholds.
- Standard costs will be effective-dated; unmatched and duplicate keys will be intentional professional exceptions rather than corrupt filler.
- All entities, identifiers, values, and movement records will be authored synthetic fixtures. They will not be described as records from a real company or audit.
- No input file has been generated at DESIGN_FROZEN status.

## Target workbook and professional operations

Required sheets are Policy, WMS_Raw, ERP_Raw, Counts, Cutoff, Reconciliation, Exceptions, Posting_View, and Checks. The workbook must normalize keys without suppressing unmatched rows, identify valid count overrides, apply owned/in-transit/cutoff logic, calculate book-to-physical and system-to-system unit differences, assign evidence-backed reason codes, value differences with effective costs, and balance the posting view to the reconciled unexplained difference.

## Independent oracle design

A task-local Python oracle will parse the visible policy and independently construct normalized composite keys, full outer joins, count precedence, cutoff ownership adjustments, effective-cost lookups, explained and unexplained unit differences, value impacts, materiality flags, and posting totals. It will emit facts keyed by item, location, lot, and owner; reject duplicate authoritative keys; and verify source-unit, adjusted-unit, and value control equations without reading workbook formulas.

## Atomic rubric design

- R001: All required sheets exist and preserve the supplied raw records.
- R002: Item, location, lot, and ownership keys are normalized and joined without dropping unmatched records.
- R003: WMS and ERP source quantities reconcile to their respective control totals.
- R004: Valid cycle counts override the correct stock keys once, while stale or duplicate counts are flagged.
- R005: Cutoff receipts, shipments, and in-transit ownership are adjusted on the correct side and with the correct sign.
- R006: Consigned and quarantined inventory follow the visible inclusion and valuation rules.
- R007: Explained and unexplained unit differences are calculated and reason-coded from source evidence.
- R008: Effective standard costs are selected by item and date, with missing or duplicate costs exposed.
- R009: Adjustment value, posting direction, and materiality priority are correct and formula-linked.
- R010: Posting_View totals equal the unexplained reconciled difference in both units and value.
- R011: Every exception retains unique source provenance and can be traced to all contributing records.
- P001: Penalize inner-join loss, silent duplicate aggregation, unsupported plugs, stale costs, or hard-coded posting totals.

## Acceptable materially different equivalent

One solution may use a normalized full-outer-join ledger with one row per reconciliation key. A materially different equivalent may keep WMS, ERP, counts, and cutoff records in separate blocks and use an explicit exception bridge before a posting summary. Row order, helper fields, and summary layout may differ; record retention, ownership/cutoff semantics, final units, value, and provenance must agree.

## Negative fixtures

- No-op: starting workbook returned with raw data present but reconciliation, exceptions, and posting results blank.
- Malformed: unreadable workbook, missing required sheets, or nonnumeric inventory outputs where calculations are required.
- M1: Join uses item only and merges distinct locations or lots.
- M2: Full outer reconciliation is replaced by an inner join, dropping unmatched ERP or WMS stock.
- M3: A post-cutoff receipt is added to both systems or applied with the wrong sign.
- M4: Blind count rows are summed instead of selecting the valid final count, double-counting physical stock.
- M5: Consigned inventory is treated as owned and posted to the company inventory account.
- M6: Current standard cost is used instead of the cost effective on the close date.

## Two meaningful input perturbations

- P1: Move one owned inbound receipt timestamp from thirty minutes after cutoff to thirty minutes before cutoff. The receipt classification, cutoff adjustment, affected item quantity difference, reason code, and posting amount must propagate; source snapshot totals, unrelated items, physical count, and standard cost remain invariant.
- P2: Change the close-date standard cost for one material unmatched item from USD 18.40 to USD 19.15 without changing quantity. Unit differences and reason codes remain invariant; the affected value difference, posting amount, materiality flag, and value controls must change by the exact unit difference times USD 0.75.

## Verifier-private CONFIRM sibling

The verifier-private CONFIRM sibling will use a different distribution center, item and lot identifiers, ownership mix, cutoff date, count hierarchy, in-transit events, costs, and exception pattern. It preserves the composite-key reconciliation, adjustment, valuation, and provenance capability contract. Confirm inputs, truth, oracle, perturbations, and reference remain outside Agent-visible input.

## Difficulty-change policy

Allowed difficulty changes: a second ownership rule, an effective-dated cost boundary, one legitimate duplicate-count resolution case, more locations, or an additional visible cutoff event type.

Forbidden difficulty changes: hiding the join keys or cutoff convention, creating ambiguous duplicate authoritative rows, adding irrelevant bulk data, corrupting interfaces, withholding a required cost, or using a model-specific trap.

## Activation rule

Activate only if a Track B PRIMARY remains VALID_BUT_EASY after three evidence-based professional revisions, fails validity and cannot be repaired without changing its capability, or a senior reviewer identifies inventory reconciliation as the stronger distinct work sample. Before activation, build and validate the full package from this frozen design; reserve status alone is not evidence of validity or difficulty.

## Build state

package_not_built: true. No task directory, workbook, fixture, Judge, Harbor wrapper, model result, Excel receipt, or human-review evidence exists for this reserve.
