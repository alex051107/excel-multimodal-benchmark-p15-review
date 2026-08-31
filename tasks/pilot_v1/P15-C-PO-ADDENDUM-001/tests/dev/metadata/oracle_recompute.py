#!/usr/bin/env python3
"""Independent semantic oracle for P15-C-PO-ADDENDUM-001; reads verifier-private typed truth."""
from __future__ import annotations
import json
from pathlib import Path

TRUTH = Path(__file__).resolve().parents[1] / "private" / "truth.json"

def load_truth():
    return json.loads(TRUTH.read_text())

def recompute():
    t = load_truth(); base = [dict(row) for row in t["base_lines"]]; revised = {row["line_id"]: dict(row) for row in base}
    for change in t["changes"]:
        if change["change_type"] == "inserted": revised[change["line_id"]] = {"line_id": change["line_id"], "description": change["description"], "quantity": change["quantity"], "unit_price": change["unit_price"], "unit": change["unit"], "protected": False, "status": change["change_type"]}
        else:
            revised[change["line_id"]].update({"quantity": change["quantity"], "unit_price": change["unit_price"], "status": change["change_type"]})
    for row in revised.values(): row.setdefault("status", "unchanged"); row["extended"] = round(row["quantity"] * row["unit_price"], 2)
    ordered = [revised[key] for key in ("PO-01", "PO-02", "PO-03", "PO-04")]
    return {"headers": t["headers"], "base": base, "changes": t["changes"], "revised": ordered, "total": round(sum(row["extended"] for row in ordered), 2), "document": t["document"]}

if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
