---
id: R15-A-MASS-BALANCE-001
track: A
status: DESIGN_FROZEN
package_not_built: true
harbor_package_built: false
primary_eligible: false
split_group: R15-A-MASS-BALANCE-001
---

# Batch-process mass and active-component balance

## Professional work and user scenario

A process engineer must reconcile wet-mass measurements and laboratory assays for multiple production batches. Feed, product, waste, wash, and internal recycle streams have different moisture and active-component assays. Management needs defensible dry-mass closure, component yield, unaccounted loss, and exceptions—not a sum of wet weights.

## Instruction draft

“Build a batch-level mass-balance workbook from the supplied measurements and assays. Normalize every stream to dry mass and active-component mass, distinguish external inputs and outputs from internal recycle, calculate overall and component closure, product yield, waste loss, and the uncertainty-based exception for each batch. Retain stream and sample identifiers and show which assay supports each normalization. Reconcile with formulas and do not suppress negative or out-of-tolerance balances.”

## Packaged-input design and authenticity boundary

- Planned inputs: stream_measurements.csv, laboratory_assays.csv, instrument_uncertainty.csv, balance_rules.md, and starting_workbook.xlsx.
- Measurements will include batch ID, stream ID, role, direction, wet mass, timestamp, and internal-recycle flag.
- Assays will include sample ID, stream ID, moisture fraction, active-on-dry-basis fraction, laboratory status, and effective timestamp.
- The visible rules will define dry mass, active mass, external closure, recycle treatment, and uncertainty thresholds.
- The fixture is synthetic and is not a real plant batch, validated analytical method, regulatory record, or released product.
- No CSV or workbook exists yet.

## Target workbook and professional operations

Required sheets are Measurements, Assays, Normalized_Streams, Batch_Balance, Component_Balance, Yield_Loss, Exceptions, and Checks. Operations include one-to-one effective assay lookup, wet-to-dry and active conversion, external input/output aggregation, separate recycle accounting, yield/loss, uncertainty evaluation, and row provenance.

## Independent oracle design

The oracle will join streams to the effective assay by IDs and timestamps, reject ambiguous joins, and compute with Decimal: dry mass equals wet mass times one minus moisture; active mass equals dry mass times dry-basis assay; external closure equals external output divided by external input; unaccounted mass equals input minus output; product yield equals product active mass divided by external input active mass. Internal recycle is reported but excluded from external closure. The Judge matches batch/stream content rather than row number.

## Atomic rubric design

- R001: Required sheets exist and all supplied batches and streams are retained.
- R002: Each measurement joins to exactly one valid assay.
- R003: Dry mass is formula-linked to wet mass and moisture in the correct direction.
- R004: Active mass is formula-linked to dry mass and dry-basis assay.
- R005: External inputs and outputs are classified correctly.
- R006: Internal recycle is reported separately and excluded from external closure.
- R007: Dry-mass closure and unaccounted dry mass are correct per batch.
- R008: Active closure, product yield, and active loss are correct per batch.
- R009: Tolerance and exception flags use the batch-appropriate uncertainty rule.
- R010: Stream/sample provenance is complete, unique, and readable.
- R011: Batch and grand-total checks reconcile without hidden plugs.
- P001: Penalize wet/dry confusion, recycle double counting, unweighted assay averaging, removed negative balances, or pasted results.

## Acceptable materially different equivalent

One solution may use a normalized stream ledger with helper columns and batch summaries. A materially different equivalent may use batch-specific blocks fed from a validated assay lookup. Structured aggregation or explicit formulas are both acceptable when classification, effective assay, normalized masses, closure, yield, flags, and provenance agree.

## Negative fixtures

- No-op: source imports present and normalized/balance outputs blank.
- Malformed: unreadable workbook or missing required calculation sheets.
- M1: Moisture is multiplied rather than subtracted from one.
- M2: Active component is calculated from wet mass rather than dry mass.
- M3: Internal recycle is added to both external feed and output.
- M4: A simple mean assay replaces the stream-specific effective assay.
- M5: Negative unaccounted mass is converted to an absolute value.
- M6: Exception thresholds ignore supplied uncertainty.

## Two meaningful input perturbations

- P1: Change one product-stream moisture fraction from 0.120 to 0.100. Product dry and active mass, both closures, yield, losses, and the batch flag must propagate; wet masses, other streams, and other batches remain invariant.
- P2: Increase an internal recycle wet mass from 240 kg to 300 kg. Recycle reporting and internal throughput change, while external inputs, outputs, closure, and product yield remain invariant.

## Verifier-private CONFIRM sibling

The CONFIRM sibling will use a different process, batch/stream/sample IDs, assay effective-date change, moisture/assay values, and recycle topology. It preserves the dry/active and external-versus-internal contract. Confirm truth, oracle, perturbations, and reference remain verifier-private; all operational rules remain visible.

## Difficulty-change policy

Allowed difficulty changes: more batches, legitimate stream roles, visible assay effective dates, uncertainty propagation, or a second active component.

Forbidden difficulty changes: missing units, hidden classification, inconsistent sample IDs, deliberately ambiguous timestamps, irrelevant noise, private truth leakage, or parser traps.

## Activation rule

Activate only as a substantive Track A replacement after a PRIMARY fails validity or remains VALID_BUT_EASY after three allowed revisions, then require full construction and every release gate.

## Build state

**package_not_built: true — no task directory, workbook, fixture, Judge, or Harbor package exists.**
