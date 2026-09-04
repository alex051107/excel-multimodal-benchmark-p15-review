#!/usr/bin/env python3
"""Focused validation for the current corrected POLICY Judge."""
from __future__ import annotations

import json
import math
import os
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests"))
import evaluate  # noqa: E402
from judge_v2_support import build_result  # noqa: E402


def result(path):
    criteria, failures = evaluate.evaluate(path)
    payload = build_result(task=evaluate.TASK, split="dev", candidate=str(path), criteria=criteria, failures=failures)
    return {
        "score": payload["normalized_score"], "pass": payload["pass"],
        "evaluation_status": payload["evaluation_status"],
        "criterion_scores": payload["criterion_scores"], "failure_codes": payload["failure_codes"],
    }


def main():
    reference = ROOT / "solution/reference.xlsx"
    cli_command = [
        sys.executable,
        str(ROOT / "tests/evaluate.py"),
        str(reference),
        "--split",
        "dev",
    ]
    with tempfile.TemporaryDirectory(prefix="p15-policy-v3-") as log_dir:
        cli_environment = dict(os.environ)
        cli_environment["P15_VERIFIER_LOG_DIR"] = log_dir
        cli_run = subprocess.run(
            cli_command,
            cwd=ROOT,
            env=cli_environment,
            text=True,
            capture_output=True,
            check=False,
        )
    try:
        cli_payload = json.loads(cli_run.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"Current corrected CLI did not emit JSON: {cli_run.stderr.strip()}"
        ) from exc
    current_default_cli = {
        "invocation": [
            "python",
            "tests/evaluate.py",
            "solution/reference.xlsx",
            "--split",
            "dev",
        ],
        "exit_code": cli_run.returncode,
        "score": cli_payload.get("normalized_score"),
        "pass": cli_payload.get("pass"),
        "evaluation_status": cli_payload.get("evaluation_status"),
        "failure_codes": cli_payload.get("failure_codes", []),
        "expected": {"status": "OK", "score": 1.0, "pass": True},
        "expectation_basis": (
            "Task version 2.0.0 is the corrected task, so the ordinary CLI "
            "must score its corrected reference without a mode flag."
        ),
    }
    current_default_cli["meets_expectation"] = (
        cli_run.returncode == 0
        and current_default_cli["evaluation_status"] == "OK"
        and math.isclose(current_default_cli["score"], 1.0, abs_tol=1e-12)
        and current_default_cli["pass"] is True
    )
    case_specs = {
        "standard": (reference, True, "The current v2 corrected reference must pass."),
        "equivalent": (ROOT / "fixtures/equivalent/candidate_equivalent.xlsx", True, "Algebraically equivalent formulas must pass."),
        "alternate_layout": (ROOT / "fixtures/equivalent/candidate_v3_separate_layout.xlsx", True, "Separate Base and Policy sheets with the same relationships must pass."),
        "noop": (ROOT / "fixtures/noop/candidate_noop.xlsx", False, "An unchanged blank scenario block does not complete the current v2 corrected task."),
        "malformed": (ROOT / "fixtures/malformed/candidate_malformed.xlsx", False, "A non-workbook cannot pass."),
    }
    for path in sorted((ROOT / "fixtures/mutants").glob("*.xlsx")):
        case_specs[f"mutant:{path.stem}"] = (path, False, "The seeded policy, unit, source, or linkage defect must prevent a corrected-task pass.")
    cases = {}
    for name, (path, expected_pass, basis) in case_specs.items():
        row = result(path)
        row.update({"expected_pass": expected_pass, "expectation_basis": basis, "meets_expectation": row["pass"] is expected_pass})
        cases[name] = row

    oracle, module = evaluate.load_oracle(evaluate.load_contract())
    source_text = (ROOT / "data/input_files/source_locator.md").read_text()
    instruction_text = (ROOT / "instruction.md").read_text()
    contract_checks = {
        "provenance_boundary": {
            "pass": "not a verbatim" in source_text.casefold() and "benchmark" in instruction_text.casefold(),
            "basis": "The task must describe the values as adapted benchmark inputs, not exact EIA rows.",
        },
        "demand_assumption_disclosed": {
            "pass": "demand value is a benchmark scenario assumption" in instruction_text.casefold(),
            "basis": "Demand=4,000,000 is project-authored and must be disclosed as such.",
        },
        "gwh_to_mwh_conversion": {
            "pass": math.isclose(oracle["policy_emissions"], 1359476000.0, abs_tol=0.5),
            "basis": "GWh × 1,000 × tCO2/MWh must yield metric tonnes.",
        },
        "historical_gas_and_other_baseline": {
            "pass": math.isclose(oracle["base"]["gas"], module.DATA["gas"], abs_tol=1e-9)
            and math.isclose(oracle["base"]["other"], oracle["policy"]["other"], abs_tol=1e-9),
            "basis": "Base gas stays historical and the implicit Other residual stays constant.",
        },
    }
    receipt = {
        "task_id": evaluate.TASK["task_id"], "judge_version": "v3-current-corrected-v2",
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "current_default_cli": current_default_cli,
        "cases": cases,
        "contract_checks": contract_checks,
        "all_expectations_met": current_default_cli["meets_expectation"]
        and all(row["meets_expectation"] for row in cases.values())
        and all(row["pass"] for row in contract_checks.values()),
    }
    destination = ROOT / "receipts/judge_v3_local_validation.json"
    destination.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
    print(json.dumps(receipt, indent=2, sort_keys=True))
    if not receipt["all_expectations_met"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
