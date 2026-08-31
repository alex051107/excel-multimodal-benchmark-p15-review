#!/usr/bin/env python3
"""Independent held-out report specification for the native-Pivot template update."""
import csv
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent / "input_files"


def recompute():
    rows = list(csv.DictReader((ROOT / "program_events.csv").open()))
    q2 = [row for row in rows if row["quarter"] == "2024Q2"]
    return {
        "q2_participants": sum(int(row["participants"]) for row in q2),
        "q2_spend": sum(int(row["spend"]) for row in q2),
        "focus_participants": sum(int(row["participants"]) for row in q2 if row["region"] == "East" and row["program"] == "Vaccination"),
        "focus_region": "East",
        "focus_program": "Vaccination",
        "source_range": "Program_Data!A3:F11",
    }


if __name__ == "__main__":
    print(json.dumps(recompute(), sort_keys=True))
