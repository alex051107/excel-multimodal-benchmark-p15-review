#!/usr/bin/env python3
"""Independent held-out invoice-level FX reconciliation replay."""
import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent / "input_files"


def recompute():
    fx = {(row["date"], row["currency"]): float(row["usd_per_unit"]) for row in csv.DictReader((ROOT / "fx_rates.csv").open())}
    adjustments = {row["invoice"]: float(row["amount_usd"]) for row in csv.DictReader((ROOT / "adjustments.csv").open())}
    sources = {
        "Ledger": list(csv.DictReader((ROOT / "ledger.csv").open())),
        "Subledger": list(csv.DictReader((ROOT / "subledger.csv").open())),
    }
    normalized, by_source = [], {"Ledger": {}, "Subledger": {}}
    for source, records in sources.items():
        for row in records:
            rate = 1.0 if row["currency"] == "USD" else fx[(row["date"], row["currency"])]
            usd = float(row["original_amount"]) * rate
            by_source[source][row["invoice"]] = usd
            normalized.append((source, row["record_id"], row["invoice"], row["date"], row["currency"], float(row["original_amount"]), rate, usd))
    matched, unmatched = [], []
    for invoice in sorted(set(by_source["Ledger"]) | set(by_source["Subledger"])):
        ledger, subledger = by_source["Ledger"].get(invoice), by_source["Subledger"].get(invoice)
        if ledger is None:
            unmatched.append((invoice, "Subledger", subledger))
            continue
        if subledger is None:
            unmatched.append((invoice, "Ledger", ledger))
            continue
        adjustment = adjustments.get(invoice, 0.0)
        adjusted = subledger + adjustment
        matched.append((invoice, ledger, subledger, adjustment, adjusted, ledger - adjusted))
    ledger_total = sum(by_source["Ledger"].values())
    matched_total = sum(row[4] for row in matched)
    ledger_only = sum(row[2] for row in unmatched if row[1] == "Ledger")
    subledger_only = sum(row[2] for row in unmatched if row[1] == "Subledger")
    return {
        "normalized": normalized,
        "matched": matched,
        "unmatched": unmatched,
        "ledger_total": ledger_total,
        "matched_total": matched_total,
        "ledger_only": ledger_only,
        "subledger_only": subledger_only,
        "investigation": subledger_only - ledger_only,
    }


if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
