#!/usr/bin/env python3
"""Independent semantic oracle for P15-C-STATEMENT-001; reads verifier-private typed truth."""
from __future__ import annotations
import json
from pathlib import Path

TRUTH = Path(__file__).resolve().parents[1] / "private" / "truth.json"

def load_truth():
    return json.loads(TRUTH.read_text())

def recompute():
    t = load_truth(); transactions = []
    for row in t["transactions"]: transactions.append({**row, "net": round(row["credit"] - row["debit"], 2)})
    credits = round(sum(row["credit"] for row in transactions), 2); debits = round(sum(row["debit"] for row in transactions), 2)
    closing = round(t["headers"]["opening_balance"] + credits - debits, 2)
    categories = {"Customer receipts": round(sum(row["net"] for row in transactions if row["category"] == "Customer receipts"), 2), "Equipment": round(sum(row["net"] for row in transactions if row["category"] == "Equipment"), 2), "Travel and fees": round(sum(row["net"] for row in transactions if row["category"] in ("Travel", "Bank fees")), 2)}
    return {"headers": t["headers"], "transactions": transactions, "credits": credits, "debits": debits, "closing": closing, "categories": categories, "document": t["document"]}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
