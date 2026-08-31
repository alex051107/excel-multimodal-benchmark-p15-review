#!/usr/bin/env python3
"""Verifier-private confirm oracle for the distinct PO-8820 addendum instance."""
from __future__ import annotations
import json
from pathlib import Path

TRUTH = Path(__file__).with_name("truth.json")

def recompute():
    truth = json.loads(TRUTH.read_text(encoding="utf-8"))
    base = [dict(row) for row in truth["base_lines"]]
    revised = {row["line_id"]: dict(row) for row in base}
    for change in truth["changes"]:
        if change["change_type"] == "inserted":
            revised[change["line_id"]] = {"line_id": change["line_id"], "description": change["description"], "quantity": change["quantity"], "unit_price": change["unit_price"], "unit": change["unit"], "protected": False, "status": "inserted"}
        else:
            revised[change["line_id"]].update({"quantity": change["quantity"], "unit_price": change["unit_price"], "status": change["change_type"]})
    for row in revised.values():
        row.setdefault("status", "unchanged")
        row["extended"] = round(row["quantity"] * row["unit_price"], 2)
    ordered = [revised[line_id] for line_id in truth["revised_order"]]
    return {"headers": truth["headers"], "base": base, "changes": truth["changes"], "revised": ordered, "total": round(sum(row["extended"] for row in ordered), 2), "document": truth["document"]}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
