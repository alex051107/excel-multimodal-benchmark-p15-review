#!/usr/bin/env python3
"""Independent held-out release-selection and KPI replay."""
import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent / "input_files"


def recompute():
    registry = list(csv.DictReader((ROOT / "release_registry.csv").open()))
    candidates = [row for row in registry if row["period"] == "2024Q2" and row["coverage"] == "complete" and row["schema"] == "1.1"]
    selected = max(candidates, key=lambda row: row["published"])
    rows = list(csv.DictReader((ROOT / selected["file"]).open()))
    normalized = [
        (row["record_id"], row["region"], row["product"], row["stage"], int(row["bookings"]), int(row["recognized_revenue"]), int(row["closed_won_flag"]))
        for row in rows
    ]
    return {
        "selection": (selected["release_id"], selected["file"], selected["period"], selected["coverage"]),
        "rows": normalized,
        "bookings": sum(row[4] for row in normalized),
        "revenue": sum(row[5] for row in normalized),
        "closed_won": sum(row[6] for row in normalized),
        "south_revenue": sum(row[5] for row in normalized if row[1] == "South"),
    }


if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
