#!/usr/bin/env python3
"""Independent release-selection and KPI replay; it does not read the workbook."""
import csv, json
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1] / "data" / "input_files"
def recompute():
    registry = list(csv.DictReader((ROOT / "release_registry.csv").open()))
    candidates = [r for r in registry if r["period"] == "2024Q2" and r["coverage"] == "complete" and r["schema"] == "1.1"]
    selected = max(candidates, key=lambda r: r["published"])
    rows = list(csv.DictReader((ROOT / selected["file"]).open()))
    normalized = [(r["record_id"], r["region"], r["product"], r["stage"], int(r["bookings"]), int(r["recognized_revenue"]), int(r["closed_won_flag"])) for r in rows]
    return {"selection": (selected["release_id"], selected["file"], selected["period"], selected["coverage"]), "rows": normalized, "bookings": sum(r[4] for r in normalized), "revenue": sum(r[5] for r in normalized), "closed_won": sum(r[6] for r in normalized), "south_revenue": sum(r[5] for r in normalized if r[1] == "South")}
if __name__ == "__main__": print(json.dumps(recompute(), sort_keys=True))
