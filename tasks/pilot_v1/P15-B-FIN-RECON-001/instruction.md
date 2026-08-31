# June ledger-to-subledger reconciliation

Accounting needs the June reconciliation closed for review. Normalize both ledgers using the transaction-date FX policy, apply only the documented adjustment, match valid invoices within the stated tolerance, and retain all unmatched items.

Deliver `/app/output/answer.xlsx` with normalized records, matched and unmatched items, the adjustment evidence, a variance bridge, a final decision sheet, and integrity checks. Do not overwrite the input ledgers, FX table, or adjustment evidence.
