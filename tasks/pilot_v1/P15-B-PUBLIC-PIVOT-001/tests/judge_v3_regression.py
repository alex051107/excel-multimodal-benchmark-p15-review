#!/usr/bin/env python3
"""Focused Judge V3 regression for native Pivot object semantics and status."""
from __future__ import annotations

import json
import re
import tempfile
import zipfile
from pathlib import Path

import evaluate as judge

ROOT = Path(__file__).resolve().parents[1]


def run_case(path: Path, split: str = "dev") -> dict:
    criteria, failures, native_status = judge.evaluate(path, split)
    payload = judge.build_result(
        task=judge.TASK,
        split=split,
        candidate=str(path),
        criteria=criteria,
        failures=failures,
        explicit_status=(
            "JUDGE_ERROR"
            if native_status in {"JUDGE_ERROR", "NATIVE_RECALC_REQUIRED"}
            else "OK"
        ),
    )
    return {
        "score": payload["normalized_score"], "pass": payload["pass"],
        "native_status": native_status, "criteria": payload["criterion_scores"],
        "failure_codes": payload["failure_codes"],
    }


def rewrite_package(source: Path, destination: Path, transform) -> None:
    with zipfile.ZipFile(source) as source_zip, zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as destination_zip:
        for item in source_zip.infolist():
            data = source_zip.read(item.filename)
            destination_zip.writestr(item, transform(item.filename, data))


def alternative_object_names(path: Path) -> None:
    def transform(name: str, data: bytes) -> bytes:
        if name.endswith(".xml") or name.endswith(".rels"):
            data = data.replace(b"ProgramEventsTable", b"Q2EventsInput")
            data = data.replace(b"ProgramDeliveryPivot", b"Q2DeliveryAnalysis")
        return data
    rewrite_package(ROOT / "solution/reference.xlsx", path, transform)


def without_native_kpi_cache(path: Path) -> None:
    pattern = re.compile(
        rb'(<c r="B[456]"[^>]*><f>GETPIVOTDATA.*?</f>)<v>[^<]*</v>(</c>)'
    )
    def transform(name: str, data: bytes) -> bytes:
        return pattern.sub(rb"\1<v></v>\2", data) if name.startswith("xl/worksheets/") else data
    rewrite_package(ROOT / "solution/reference.xlsx", path, transform)


def main() -> None:
    cases = {
        "reference": (ROOT / "solution/reference.xlsx", "dev", True, "NATIVE_OBJECT_CHECKED"),
        "confirm_reference": (ROOT / "tests/confirm/reference.xlsx", "confirm", True, "NATIVE_OBJECT_CHECKED"),
        "equivalent": (ROOT / "fixtures/equivalent/candidate_equivalent.xlsx", "dev", True, "NATIVE_OBJECT_CHECKED"),
        "noop": (ROOT / "fixtures/noop/candidate_noop.xlsx", "dev", False, "NATIVE_OBJECT_CHECKED"),
        "malformed": (ROOT / "fixtures/malformed/candidate_malformed.xlsx", "dev", False, "TASK_INVALID"),
    }
    cases.update({
        f"mutant:{path.stem}": (path, "dev", False, "NATIVE_OBJECT_CHECKED")
        for path in sorted((ROOT / "fixtures/mutants").glob("*.xlsx"))
    })
    results = {}
    with tempfile.TemporaryDirectory() as directory:
        alternative = Path(directory) / "renamed_native_objects.xlsx"
        pending = Path(directory) / "native_recalc_pending.xlsx"
        alternative_object_names(alternative)
        without_native_kpi_cache(pending)
        cases["alternative_layout:renamed_native_objects"] = (
            alternative, "dev", True, "NATIVE_OBJECT_CHECKED"
        )
        cases["native_recalc_pending"] = (
            pending, "dev", False, "NATIVE_RECALC_REQUIRED"
        )
        for name, (path, split, expected_pass, expected_status) in cases.items():
            result = run_case(path, split)
            assert result["pass"] is expected_pass, (name, result)
            assert result["native_status"] == expected_status, (name, result)
            results[name] = {
                **result, "expected_pass": expected_pass,
                "expected_native_status": expected_status,
            }
    print(json.dumps(results, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
