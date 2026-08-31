#!/usr/bin/env python3
"""Independent held-out public-health time-series replay."""
import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent / "input_files"


def recompute():
    rows = []
    for filename in ("vermont_heart_disease.csv", "new_hampshire_heart_disease.csv"):
        rows.extend((row["state"], int(row["year"]), int(row["deaths"]), float(row["aadr"])) for row in csv.DictReader((ROOT / filename).open()))
    metrics = {}
    for state in ("Vermont", "New Hampshire"):
        records = sorted([row for row in rows if row[0] == state], key=lambda row: row[1])
        baseline = sum(row[3] for row in records[:3]) / 3
        current = sum(row[3] for row in records[3:]) / 3
        metrics[state] = (baseline, current, current - baseline, (current - baseline) / baseline)
    years = sorted({row[1] for row in rows})
    return {
        "rows": rows,
        "metrics": metrics,
        "highest_current": max(metrics, key=lambda state: metrics[state][1]),
        "baseline_label": f"Baseline {years[0]}-{years[2]}",
        "current_label": f"Current {years[3]}-{years[5]}",
    }


if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
