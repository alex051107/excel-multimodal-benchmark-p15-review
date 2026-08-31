#!/usr/bin/env python3
"""Independent semantic oracle for P15-C-INVOICE-001; reads verifier-private typed truth."""
from __future__ import annotations
import json
from pathlib import Path

TRUTH = Path(__file__).resolve().parents[1] / "private" / "truth.json"

def load_truth():
    return json.loads(TRUTH.read_text())

def recompute():
    t = load_truth(); items = []
    for row in t["line_items"]:
        items.append({**row, "line_total": round(row["quantity"] * row["unit_price"], 2)})
    subtotal = round(sum(row["line_total"] for row in items), 2)
    discount = round(subtotal * t["terms"]["discount_rate"], 2)
    taxable = round(subtotal - discount, 2)
    tax = round(taxable * t["terms"]["tax_rate"], 2)
    return {"headers": t["headers"], "items": items, "subtotal": subtotal, "discount": discount, "taxable": taxable, "tax": tax, "freight": t["terms"]["freight"], "total": round(taxable + tax + t["terms"]["freight"], 2), "document": t["document"]}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
