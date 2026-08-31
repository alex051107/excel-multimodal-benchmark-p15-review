# Shipment-order master-data cleanup

The operations team needs a reliable weekly order-cost view. Normalize the order extract, join only valid records to the product and location masters, and retain every data-quality exception for follow-up. Exact normalized duplicates must not inflate totals.

Deliver `/app/output/answer.xlsx` with clean records, joined valid records, an exceptions queue, a summary, and checks. Preserve the raw order extract and both masters.
