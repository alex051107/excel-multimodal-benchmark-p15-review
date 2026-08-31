#!/usr/bin/env python3
"""Independent public-health time-series replay from the two supplied CDC extracts."""
import csv, json
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1] / "data" / "input_files"
def recompute():
    rows = []
    for filename in ("vermont_heart_disease.csv", "new_hampshire_heart_disease.csv"):
        rows.extend((r["state"], int(r["year"]), int(r["deaths"]), float(r["aadr"])) for r in csv.DictReader((ROOT / filename).open()))
    metrics = {}
    for state in ("Vermont", "New Hampshire"):
        records = sorted([r for r in rows if r[0] == state], key=lambda r: r[1])
        baseline = sum(r[3] for r in records[:3]) / 3
        current = sum(r[3] for r in records[3:]) / 3
        metrics[state] = (baseline, current, current - baseline, (current - baseline) / baseline)
    return {"rows": rows, "metrics": metrics, "highest_current": max(metrics, key=lambda state: metrics[state][1])}
if __name__ == "__main__": print(json.dumps(recompute(), sort_keys=True))
