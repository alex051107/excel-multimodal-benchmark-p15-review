#!/usr/bin/env python3
"""Independent report specification for the future native-Pivot build."""
import csv, json
from pathlib import Path
ROOT = Path(__file__).resolve().parents[1] / "data" / "input_files"
def recompute():
    rows = list(csv.DictReader((ROOT / "program_events.csv").open()))
    q2 = [r for r in rows if r["quarter"] == "2024Q2"]
    return {"q2_participants": sum(int(r["participants"]) for r in q2), "q2_spend": sum(int(r["spend"]) for r in q2), "north_outreach": sum(int(r["participants"]) for r in q2 if r["region"] == "North" and r["program"] == "Outreach"), "source_range": "Program_Data!A3:F11"}
if __name__ == "__main__": print(json.dumps(recompute(), sort_keys=True))
