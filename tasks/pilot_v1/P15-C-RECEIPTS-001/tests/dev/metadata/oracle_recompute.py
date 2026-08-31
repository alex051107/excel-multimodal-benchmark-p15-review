#!/usr/bin/env python3
"""Independent semantic oracle for P15-C-RECEIPTS-001; reads verifier-private typed truth."""
from __future__ import annotations
import json
from pathlib import Path

TRUTH = Path(__file__).resolve().parents[1] / "private" / "truth.json"

def load_truth():
    return json.loads(TRUTH.read_text())

def recompute():
    t = load_truth(); documents = []
    for doc in t["documents"]:
        items = [row for row in t["items"] if row["document_id"] == doc["document_id"]]
        subtotal = round(sum(row["amount"] for row in items), 2); tax = round(subtotal * doc["tax_rate"], 2); total = round(subtotal + tax + doc["tip"], 2)
        documents.append({**doc, "subtotal": subtotal, "tax": tax, "total": total})
    categories = {category: round(sum(doc["total"] for doc in documents if any(item["document_id"] == doc["document_id"] and item["category"] == category for item in t["items"])), 2) for category in ("Meals", "Travel", "Office")}
    return {"documents": documents, "items": t["items"], "categories": categories, "batch_total": round(sum(doc["total"] for doc in documents), 2)}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
