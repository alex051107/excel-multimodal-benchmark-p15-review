#!/usr/bin/env python3
"""Execute the declared DEV input perturbations in background Microsoft Excel."""
from __future__ import annotations

import json
import argparse
import subprocess
from datetime import datetime, timezone
from pathlib import Path


TASK_ROOT = Path(__file__).resolve().parents[1]
SCRIPT = Path(__file__).with_name("validate_input_perturbation_step_background.applescript")
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--workbook-name", required=True, help="Background Excel workbook produced by the native builder")
    parser.add_argument("--split", choices=("dev", "confirm"), default="dev")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    def change(cell: str, value: float) -> list[float]:
        completed = subprocess.run(
            ["osascript", str(SCRIPT), args.workbook_name, cell, str(value)],
            check=True,
            capture_output=True,
            text=True,
        )
        values = completed.stdout.strip().split("|")
        if len(values) != 3:
            raise RuntimeError(f"unexpected Excel readback: {completed.stdout!r}")
        return [float(value) for value in values]

    if args.split == "dev":
        participant_cell, participant_value, participant_baseline = "E5", 55, 50
        spend_cell, spend_value, spend_baseline = "F5", 151000, 150000
        expected_participants, expected_focus, expected_spend = 137.0, 55.0, 397000.0
        restored_expected = (132.0, 396000.0, 50.0)
        participant_id = "Q2_NORTH_OUTREACH_PARTICIPANTS_50_TO_55"
        spend_id = "Q2_NORTH_OUTREACH_SPEND_PLUS_1000"
        receipt = TASK_ROOT / "receipts" / "input_perturbation_validation.json"
    else:
        participant_cell, participant_value, participant_baseline = "E5", 39, 38
        spend_cell, spend_value, spend_baseline = "F11", 58500, 57500
        expected_participants, expected_focus, expected_spend = 123.0, 39.0, 342000.0
        restored_expected = (122.0, 341000.0, 38.0)
        participant_id = "confirm_participant_refresh"
        spend_id = "confirm_spend_refresh"
        receipt = TASK_ROOT / "receipts" / "confirm_input_perturbation_validation.json"

    participant_case = change(participant_cell, participant_value)
    participant_restored = change(participant_cell, participant_baseline)
    spend_case = change(spend_cell, spend_value)
    restored = change(spend_cell, spend_baseline)
    observed = {
        "participant_case_q2_total": participant_case[0],
        "participant_case_focus": participant_case[2],
        "participant_restore": participant_restored,
        "spend_case_q2_total": spend_case[1],
        "restored_q2_participants": restored[0],
        "restored_q2_spend": restored[1],
        "restored_focus_participants": restored[2],
    }
    checks = {
        "participant_total_propagated": observed["participant_case_q2_total"] == expected_participants,
        "focus_participants_propagated": observed["participant_case_focus"] == expected_focus,
        "spend_total_propagated": observed["spend_case_q2_total"] == expected_spend,
        "baseline_restored": (
            observed["restored_q2_participants"] == restored_expected[0]
            and observed["restored_q2_spend"] == restored_expected[1]
            and observed["restored_focus_participants"] == restored_expected[2]
        ),
    }
    if not all(checks.values()):
        raise RuntimeError(f"perturbation propagation failed: {observed}")
    payload = {
        "task_id": "P15-B-PUBLIC-PIVOT-001",
        "split": args.split,
        "status": "PASS",
        "validated_at_utc": datetime.now(timezone.utc).isoformat(),
        "screen_control_used": False,
        "cases": [
            {
                "id": participant_id,
                "changed_source_cell": f"Program_Data!{participant_cell}",
                "observed": {"q2_participants": participant_case[0], "focus_participants": participant_case[2]},
            },
            {
                "id": spend_id,
                "changed_source_cell": f"Program_Data!{spend_cell}",
                "observed": {"q2_spend": spend_case[1]},
            },
        ],
        "restoration": {"observed": observed, "checks": checks},
        "background_excel_workbook": args.workbook_name,
    }
    receipt.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n")
    print(json.dumps(payload, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
