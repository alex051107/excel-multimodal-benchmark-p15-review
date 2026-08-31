#!/usr/bin/env python3
"""Verifier-private confirm oracle for the distinct July receipt batch."""
from __future__ import annotations
import json
from pathlib import Path

TRUTH = Path(__file__).with_name("truth.json")

def recompute():
    truth = json.loads(TRUTH.read_text(encoding="utf-8"))
    documents = []
    for document in truth["documents"]:
        rows = [row for row in truth["items"] if row["document_id"] == document["document_id"]]
        subtotal = round(sum(row["amount"] for row in rows), 2)
        tax = round(subtotal * document["tax_rate"], 2)
        documents.append({**document, "subtotal": subtotal, "tax": tax, "total": round(subtotal + tax + document["tip"], 2)})
    categories = {
        category: round(sum(document["total"] for document in documents if any(row["document_id"] == document["document_id"] and row["category"] == category for row in truth["items"])), 2)
        for category in ("Meals", "Travel", "Office")
    }
    return {"documents": documents, "items": truth["items"], "categories": categories, "batch_total": round(sum(document["total"] for document in documents), 2)}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
