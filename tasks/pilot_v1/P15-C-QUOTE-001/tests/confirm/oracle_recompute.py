#!/usr/bin/env python3
"""Verifier-private confirm oracle for the distinct Granite Peak quote instance."""
from __future__ import annotations
import json
from pathlib import Path

TRUTH = Path(__file__).with_name("truth.json")

def recompute():
    truth = json.loads(TRUTH.read_text(encoding="utf-8"))
    base_rows = [row for row in truth["line_items"] if not row["optional"]]
    alternate_rows = [row for row in truth["line_items"] if row["optional"]]
    base = round(sum(row["amount"] for row in base_rows), 2)
    discount = round(base * truth["terms"]["discount_rate"], 2)
    tax = round((base - discount) * truth["terms"]["tax_rate"], 2)
    return {"headers": truth["headers"], "items": truth["line_items"], "document": truth["document"], "base": base, "discount": discount, "tax": tax, "total": round(base - discount + tax, 2), "alternate": round(sum(row["amount"] for row in alternate_rows), 2)}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
