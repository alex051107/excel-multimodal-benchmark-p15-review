#!/usr/bin/env python3
"""Verifier-private confirm oracle for the distinct Copperline invoice instance."""
from __future__ import annotations
import json
from pathlib import Path

TRUTH = Path(__file__).with_name("truth.json")

def recompute():
    truth = json.loads(TRUTH.read_text(encoding="utf-8"))
    items = [{**row, "line_total": round(row["quantity"] * row["unit_price"], 2)} for row in truth["line_items"]]
    subtotal = round(sum(row["line_total"] for row in items), 2)
    discount = round(subtotal * truth["terms"]["discount_rate"], 2)
    taxable = round(subtotal - discount, 2)
    tax = round(taxable * truth["terms"]["tax_rate"], 2)
    total = round(taxable + tax + truth["terms"]["freight"], 2)
    return {"headers": truth["headers"], "items": items, "subtotal": subtotal, "discount": discount, "taxable": taxable, "tax": tax, "freight": truth["terms"]["freight"], "total": total, "document": truth["document"]}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
