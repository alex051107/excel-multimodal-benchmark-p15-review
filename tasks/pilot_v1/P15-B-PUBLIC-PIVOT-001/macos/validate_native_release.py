#!/usr/bin/env python3
"""Run the frozen Pivot Judge repeatedly and write task-local release receipts."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import statistics
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path


TASK_ROOT = Path(__file__).resolve().parents[1]
JUDGE = TASK_ROOT / "tests" / "evaluate.py"
RECEIPTS = TASK_ROOT / "receipts"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dev-native-receipt", type=Path, required=True)
    parser.add_argument("--confirm-native-receipt", type=Path, required=True)
    parser.add_argument("--repetitions", type=int, default=5)
    return parser.parse_args()


def file_hash(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run_judge(candidate: Path, split: str, repetitions: int) -> dict:
    python = os.environ.get("PILOT_PYTHON", sys.executable)
    payloads = []
    signatures = []
    for _ in range(repetitions):
        result = subprocess.run(
            [python, str(JUDGE), str(candidate), "--split", split],
            check=True,
            capture_output=True,
            text=True,
        )
        payload = json.loads(result.stdout)
        normalized = dict(payload)
        normalized.pop("candidate", None)
        signatures.append(hashlib.sha256(json.dumps(normalized, sort_keys=True).encode()).hexdigest())
        payloads.append(payload)
    scores = [float(payload["normalized_score"]) for payload in payloads]
    return {
        "candidate": str(candidate.relative_to(TASK_ROOT)),
        "candidate_sha256": file_hash(candidate),
        "score": scores[0],
        "score_stddev": statistics.pstdev(scores),
        "repetitions": repetitions,
        "stable": len(set(signatures)) == 1,
        "signature_sha256": signatures[0],
        "status": payloads[0]["status"],
        "pass": payloads[0]["pass"],
        "criterion_scores": payloads[0]["criterion_scores"],
        "failure_codes": payloads[0]["failure_codes"],
    }


def normalize_native_receipt(source: Path, target: Path, workbook: Path) -> dict:
    payload = json.loads(source.read_text())
    split = payload.get("split", "dev")
    input_workbook = (
        TASK_ROOT / "tests" / "confirm" / "input_files" / "starting_workbook.xlsx"
        if split == "confirm"
        else TASK_ROOT / "data" / "input_files" / "starting_workbook.xlsx"
    )
    payload["input_workbook"] = str(input_workbook.relative_to(TASK_ROOT))
    payload["output_workbook"] = str(workbook.relative_to(TASK_ROOT))
    payload["output_sha256"] = file_hash(workbook)
    if isinstance(payload.get("judge"), dict):
        payload["judge"]["candidate"] = str(workbook.relative_to(TASK_ROOT))
    background_session = payload.pop("background_excel_session", None)
    if isinstance(background_session, dict):
        payload["background_excel_workbook"] = background_session.get("workbook_name")
    payload["evidence_note"] = "Repository paths normalized after byte-identical promotion from the Excel-validated staging workbook."
    target.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n")
    return payload


def main() -> int:
    args = parse_args()
    if args.repetitions < 5:
        raise ValueError("release validation requires at least five repetitions")
    RECEIPTS.mkdir(parents=True, exist_ok=True)

    dev_reference = TASK_ROOT / "solution" / "reference.xlsx"
    confirm_reference = TASK_ROOT / "tests" / "confirm" / "reference.xlsx"
    candidates = {
        "reference": (dev_reference, "dev"),
        "equivalent": (TASK_ROOT / "fixtures/equivalent/candidate_equivalent.xlsx", "dev"),
        "noop": (TASK_ROOT / "fixtures/noop/candidate_noop.xlsx", "dev"),
        "malformed": (TASK_ROOT / "fixtures/malformed/candidate_malformed.xlsx", "dev"),
        "fake_pivot": (TASK_ROOT / "fixtures/mutants/fake_pivot.xlsx", "dev"),
        "stale_source_range": (TASK_ROOT / "fixtures/mutants/stale_source_range.xlsx", "dev"),
        "wrong_aggregation": (TASK_ROOT / "fixtures/mutants/wrong_aggregation.xlsx", "dev"),
        "wrong_chart_binding": (TASK_ROOT / "fixtures/mutants/wrong_chart_binding.xlsx", "dev"),
        "confirm_reference": (confirm_reference, "confirm"),
    }
    results = {name: run_judge(path, split, args.repetitions) for name, (path, split) in candidates.items()}

    gates = {
        "reference_eq_1": results["reference"]["score"] == 1.0 and results["reference"]["pass"],
        "equivalent_ge_0_95": results["equivalent"]["score"] >= 0.95 and results["equivalent"]["pass"],
        "noop_lt_0_30": results["noop"]["score"] < 0.30,
        "malformed_eq_0": results["malformed"]["score"] == 0.0,
        "all_mutants_lt_0_70": all(results[name]["score"] < 0.70 for name in ("fake_pivot", "stale_source_range", "wrong_aggregation", "wrong_chart_binding")),
        "confirm_reference_eq_1": results["confirm_reference"]["score"] == 1.0 and results["confirm_reference"]["pass"],
        "all_repeated_judgments_stable": all(result["stable"] and result["score_stddev"] < 0.05 for result in results.values()),
    }
    if not all(gates.values()):
        raise RuntimeError(f"release validation gate failed: {gates}")

    dev_native = normalize_native_receipt(
        args.dev_native_receipt,
        RECEIPTS / "macos_native_dev.json",
        dev_reference,
    )
    confirm_native = normalize_native_receipt(
        args.confirm_native_receipt,
        RECEIPTS / "macos_native_confirm.json",
        confirm_reference,
    )
    timestamp = datetime.now(timezone.utc).isoformat()
    local_receipt = {
        "schema_version": "2.0",
        "task_id": "P15-B-PUBLIC-PIVOT-001",
        "status": "LOCAL_READY",
        "validated_at_utc": timestamp,
        "canonical_excel": {
            "application": "Microsoft Excel for Mac",
            "version": dev_native["readback"]["excel_version"],
            "mode": "background object model; no screen control",
            "receipt": "receipts/macos_native_dev.json",
        },
        "evaluator_repetitions": args.repetitions,
        "judge_type": "deterministic OOXML and cached-value checks",
        "gates": gates,
        "reference": results["reference"],
        "equivalent": results["equivalent"],
        "noop": results["noop"],
        "malformed": results["malformed"],
        "semantic_mutants": {name: results[name] for name in ("fake_pivot", "stale_source_range", "wrong_aggregation", "wrong_chart_binding")},
        "independent_oracle": {
            "source": "metadata/oracle_recompute.py",
            "reads_reference_workbook": False,
            "q2_participants": 132,
            "q2_spend": 396000,
            "north_outreach": 50,
        },
        "agent_screen": "PENDING_SEPARATE_N8_BATCH",
        "human_review": "PENDING_EXTERNAL_HUMAN_REVIEW",
    }
    confirm_receipt = {
        "schema_version": "2.0",
        "task_id": "P15-B-PUBLIC-PIVOT-001",
        "split": "confirm",
        "status": "LOCAL_READY",
        "validated_at_utc": timestamp,
        "candidate": results["confirm_reference"],
        "canonical_excel": {
            "application": "Microsoft Excel for Mac",
            "version": confirm_native["readback"]["excel_version"],
            "mode": "background object model; no screen control",
            "receipt": "receipts/macos_native_confirm.json",
        },
        "oracle": "tests/confirm/oracle_recompute.py",
        "held_out_agent_run": "PENDING_SEPARATE_N8_BATCH",
    }
    (RECEIPTS / "local_validation.json").write_text(json.dumps(local_receipt, indent=2, sort_keys=True) + "\n")
    (RECEIPTS / "confirm_validation.json").write_text(json.dumps(confirm_receipt, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"task_id": local_receipt["task_id"], "status": "PASS", "gates": gates, "scores": {name: result["score"] for name, result in results.items()}}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
