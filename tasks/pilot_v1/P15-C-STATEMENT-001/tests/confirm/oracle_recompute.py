#!/usr/bin/env python3
"""Verifier-private confirm oracle for the distinct Summit Cooperative statement."""
from __future__ import annotations
import json
from pathlib import Path

TRUTH = Path(__file__).with_name("truth.json")

def recompute():
    truth = json.loads(TRUTH.read_text(encoding="utf-8"))
    transactions = [{**row, "net": round(row["credit"] - row["debit"], 2)} for row in truth["transactions"]]
    credits = round(sum(row["credit"] for row in transactions), 2)
    debits = round(sum(row["debit"] for row in transactions), 2)
    closing = round(truth["headers"]["opening_balance"] + credits - debits, 2)
    categories = {
        "Customer receipts": round(sum(row["net"] for row in transactions if row["category"] == "Customer receipts"), 2),
        "Equipment": round(sum(row["net"] for row in transactions if row["category"] == "Equipment"), 2),
        "Travel and fees": round(sum(row["net"] for row in transactions if row["category"] in ("Travel", "Bank fees")), 2),
    }
    return {"headers": truth["headers"], "transactions": transactions, "credits": credits, "debits": debits, "closing": closing, "categories": categories, "document": truth["document"]}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
