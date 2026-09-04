# Judge V3 contract

## Business result that must be correct

- Select the approved, complete, schema-compatible corrected Q2 release.
- Preserve exactly its six records and compute the four operating KPIs and four-region coverage correctly.
- Explain the selection, keep the chosen release and file visible, and retain the registry unchanged.

## Explicit implementation requirements retained

Source selection, records, KPIs, coverage checks, rationale, provenance, and source preservation remain mandatory. The task does not require formulas or hidden source-cell edits.

The selection explanation (R005) is a delivery hurdle. It may use ordinary business language; the Judge verifies a substantive explanation attached to the one chosen-release record rather than requiring a fixed phrase or technical keyword.

The unchanged registry is only the list of candidates. It cannot itself satisfy source selection, rationale, or provenance. The workbook must contain one distinct chosen-release record, with a real explanation tied to that record; the release identifier and filename used for provenance must appear there rather than being harvested from registry-wide text.

## Preferences removed from scoring

- Fixed sheet names, column order, cell addresses, and exact reference wording.
- Formula-only KPI and coverage credit.
- Hidden bookings and revenue perturbations.
- One exact label for the selected release, metric, or coverage fields when the business identity remains unambiguous.

## New positive and negative examples

- Positive: `PFjoLEa` correctly selects the release, records, KPIs, coverage, rationale, provenance, and registry without using formulas.
- Positive: `hardcoded_recognized_revenue.xlsx` is reclassified as equivalent because 377,000 is correct and formula use was not requested.
- Negative: `wrong_source_release.xlsx`, `dropped_south_coverage.xlsx`, and `registry_collateral_edit.xlsx` still fail.
