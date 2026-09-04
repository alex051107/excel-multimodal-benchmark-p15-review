# Judge V3 contract

## Business result that must be correct

- Normalize all seven source orders, keep the exact duplicate from inflating totals, and classify invalid units and missing master-data joins correctly.
- Join exactly the three valid records with correct product, region, units, unit cost, and extended cost.
- Report the $3,940 accepted-order total, retain all four exceptions, and provide a reviewable weekly order-cost view and checks.
- Preserve the raw extract and both masters.

## Explicit implementation requirements retained

Correct normalization, joining, duplicate handling, exceptions, total, weekly view, checks, and source preservation remain mandatory. No live perturbation or formula behavior is stated.

## Preferences removed from scoring

- Fixed output columns, fixed row numbers, and exact sheet-name punctuation.
- Particular disposition wording when the same accepted/exception meaning is clear.
- Formula-only extended costs, summary values, and check values.
- Hidden unit-cost and unit-count perturbations.

## New positive and negative examples

- Positive: `gK8tQgf` has a wider professional joined table, all four exception types, the weekly view, and the correct $3,940 total.
- Positive: `hardcoded_extended_cost.xlsx` is reclassified as equivalent because its business values are correct and formulas were never requested.
- Negative: `wrong_product_join.xlsx`, `duplicate_included.xlsx`, and `dropped_exception.xlsx` still fail.
