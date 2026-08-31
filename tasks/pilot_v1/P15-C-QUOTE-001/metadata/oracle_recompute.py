#!/usr/bin/env python3
"""Independent semantic oracle for P15-C-QUOTE-001; reads verifier-private typed truth."""
from __future__ import annotations
import json
from pathlib import Path

TRUTH = Path(__file__).resolve().parents[1] / "private" / "truth.json"

def load_truth():
    return json.loads(TRUTH.read_text())

def recompute():
    t = load_truth(); base = [row for row in t["line_items"] if not row["optional"]]; alternate = [row for row in t["line_items"] if row["optional"]]
    subtotal = round(sum(row["amount"] for row in base), 2); discount = round(subtotal * t["terms"]["discount_rate"], 2); tax = round((subtotal - discount) * t["terms"]["tax_rate"], 2)
    return {"headers": t["headers"], "items": t["line_items"], "document": t["document"], "base": subtotal, "discount": discount, "tax": tax, "total": round(subtotal - discount + tax, 2), "alternate": round(sum(row["amount"] for row in alternate), 2)}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
