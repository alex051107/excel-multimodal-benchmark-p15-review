#!/usr/bin/env python3
"""Focused Judge V3 regression: equivalent layout plus real business errors."""
from __future__ import annotations

import json
import tempfile
from pathlib import Path

import openpyxl

import evaluate as judge

ROOT = Path(__file__).resolve().parents[1]


def run_case(path: Path, split: str = "dev") -> dict:
    criteria, failures = judge.evaluate(path, split)
    payload = judge.build_result(
        task=judge.TASK, split=split, candidate=str(path), criteria=criteria, failures=failures
    )
    return {
        "score": payload["normalized_score"], "pass": payload["pass"],
        "criteria": payload["criterion_scores"], "failure_codes": payload["failure_codes"],
    }


def alternative_layout(path: Path) -> None:
    workbook = openpyxl.load_workbook(ROOT / "solution/reference.xlsx")
    workbook["Checks"].title = "Data Quality Checks"
    workbook.save(path)


def main() -> None:
    expected_mutants = {"hardcoded_extended_cost": True}
    cases = {
        "reference": (ROOT / "solution/reference.xlsx", "dev", True),
        "confirm_reference": (ROOT / "tests/confirm/reference.xlsx", "confirm", True),
        "equivalent": (ROOT / "fixtures/equivalent/candidate_equivalent.xlsx", "dev", True),
        "noop": (ROOT / "fixtures/noop/candidate_noop.xlsx", "dev", False),
        "malformed": (ROOT / "fixtures/malformed/candidate_malformed.xlsx", "dev", False),
    }
    cases.update({
        f"mutant:{path.stem}": (path, "dev", expected_mutants.get(path.stem, False))
        for path in sorted((ROOT / "fixtures/mutants").glob("*.xlsx"))
    })
    results = {}
    with tempfile.TemporaryDirectory() as directory:
        alternative = Path(directory) / "data_quality_checks_with_spaces.xlsx"
        alternative_layout(alternative)
        cases["alternative_layout:data_quality_checks_with_spaces"] = (alternative, "dev", True)
        for name, (path, split, expected_pass) in cases.items():
            result = run_case(path, split)
            assert result["pass"] is expected_pass, (name, result)
            results[name] = {**result, "expected_pass": expected_pass}
    print(json.dumps(results, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
