# Judge V3 contract

## Business result that must be correct

- Use the transaction-date FX policy for every retained ledger and subledger row.
- Apply only the documented adjustment, match invoices within the stated $1 tolerance, and retain ledger-only and subledger-only exceptions separately.
- Close the June variance bridge and final review decision to those records.
- Preserve the ledgers, FX table, and adjustment evidence.

## Explicit implementation requirements retained

The task asks for normalized, matched, unmatched, bridge, decision, and integrity-check outputs. It does not require live source-cell perturbation behavior or a particular Excel formula syntax.

The close period is the requested month represented by the eight in-period source rows. The later two-row invoice pair remains in normalized evidence but is excluded from matching, bridge totals, and the final decision. Matching and exception tables are exact multisets: blank, duplicate, extra, or wrong-amount records do not pass. The bridge, final decision, and checks must each contain their own verified amounts; words found elsewhere in the workbook cannot substitute for those values. Reconciliation closure (R004) is a delivery hurdle.

## Preferences removed from scoring

- Exact sheet names, row numbers, column order, and reference-workbook layout.
- A fixed four-match interpretation that silently includes the July `INV-102` pair in a June close. A documented June-only close may retain those rows as out of period.
- Formula-only credit for normalization, bridge, decision, exception, or check values.
- Fixed check labels such as one exact wording for the closure residual.

## New positive and negative examples

- Positive: `F3asQhw` is a static, reviewable June close with correct FX, three June matches, two exceptions, and the July pair retained out of period.
- Negative: `wrong_fx_date.xlsx` and `dropped_unmatched_item.xlsx` still fail because they change business facts.
