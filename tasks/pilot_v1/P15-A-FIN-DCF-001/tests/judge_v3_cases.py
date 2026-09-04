#!/usr/bin/env python3
"""One focused V3 validation pass for the DCF Judge."""
from __future__ import annotations

import json
import sys
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import load_workbook

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


def build_constant_formula_mutant(path):
    workbook = load_workbook(ROOT / "solution/reference.xlsx")
    oracle, _ = evaluate.load_oracle(evaluate.load_contract())
    workbook["Forecast"]["F5"] = f"={oracle['revenue'][-1]}"
    workbook["Valuation"]["B9"] = f"={oracle['enterprise_value']}"
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(path)


def main():
    equivalent = ROOT / "fixtures/equivalent/candidate_equivalent.xlsx"
    constant_formula = ROOT / "fixtures/mutants/constant_formula_dynamics.xlsx"
    build_constant_formula_mutant(constant_formula)
    cases = {
        "standard": (ROOT / "solution/reference.xlsx", True, "The reference valuation must pass."),
        "equivalent": (equivalent, True, "Algebraically equivalent valuation formulas must pass."),
        "alternate_implementation": (equivalent, True, "The task requires the supplied five-sheet structure, so the admissible alternative is formula equivalence within that structure."),
        "noop": (ROOT / "fixtures/noop/candidate_noop.xlsx", False, "An uncompleted DCF does not pass."),
        "malformed": (ROOT / "fixtures/malformed/candidate_malformed.xlsx", False, "A non-workbook cannot pass."),
    }
    for path in sorted((ROOT / "fixtures/mutants").glob("*.xlsx")):
        cases[f"mutant:{path.stem}"] = (path, False, "The seeded forecast or valuation defect must prevent a pass.")
    observed = {}
    for name, (path, expected_pass, basis) in cases.items():
        row = result(path)
        row.update({"expected_pass": expected_pass, "expectation_basis": basis, "meets_expectation": row["pass"] is expected_pass})
        observed[name] = row
    receipt = {
        "task_id": evaluate.TASK["task_id"], "judge_version": "v3-contract-v2-semantics-retained",
        "generated_at": datetime.now(timezone.utc).isoformat(), "cases": observed,
        "all_expectations_met": all(row["meets_expectation"] for row in observed.values()),
    }
    destination = ROOT / "receipts/judge_v3_local_validation.json"
    destination.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
    print(json.dumps(receipt, indent=2, sort_keys=True))
    if not receipt["all_expectations_met"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
