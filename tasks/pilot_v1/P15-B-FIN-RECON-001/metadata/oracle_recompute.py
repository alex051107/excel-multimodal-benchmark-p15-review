#!/usr/bin/env python3
"""Independent invoice-level FX reconciliation replay."""
import csv, json
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1] / "data" / "input_files"
def recompute():
    fx = {(r["date"], r["currency"]): float(r["usd_per_unit"]) for r in csv.DictReader((ROOT / "fx_rates.csv").open())}
    adjustments = {r["invoice"]: float(r["amount_usd"]) for r in csv.DictReader((ROOT / "adjustments.csv").open())}
    sources = {"Ledger": list(csv.DictReader((ROOT / "ledger.csv").open())), "Subledger": list(csv.DictReader((ROOT / "subledger.csv").open()))}
    normalized, by_source = [], {"Ledger": {}, "Subledger": {}}
    for source, records in sources.items():
        for r in records:
            rate = 1.0 if r["currency"] == "USD" else fx[(r["date"], r["currency"])]
            usd = float(r["original_amount"]) * rate
            by_source[source][r["invoice"]] = usd
            normalized.append((source, r["record_id"], r["invoice"], r["date"], r["currency"], float(r["original_amount"]), rate, usd))
    matched, unmatched = [], []
    for invoice in sorted(set(by_source["Ledger"]) | set(by_source["Subledger"])):
        ledger, sub = by_source["Ledger"].get(invoice), by_source["Subledger"].get(invoice)
        if ledger is None: unmatched.append((invoice, "Subledger", sub)); continue
        if sub is None: unmatched.append((invoice, "Ledger", ledger)); continue
        adjustment = adjustments.get(invoice, 0.0); adjusted = sub + adjustment
        matched.append((invoice, ledger, sub, adjustment, adjusted, ledger - adjusted))
    ledger_total = sum(by_source["Ledger"].values())
    matched_total = sum(row[4] for row in matched)
    ledger_only = sum(row[2] for row in unmatched if row[1] == "Ledger")
    subledger_only = sum(row[2] for row in unmatched if row[1] == "Subledger")
    return {"normalized": normalized, "matched": matched, "unmatched": unmatched, "ledger_total": ledger_total, "matched_total": matched_total, "ledger_only": ledger_only, "subledger_only": subledger_only, "investigation": subledger_only - ledger_only}
if __name__ == "__main__": print(json.dumps(recompute(), sort_keys=True))
