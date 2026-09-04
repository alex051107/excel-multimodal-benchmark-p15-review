#!/usr/bin/env python3
"""Replay current P15 Judges against every frozen workbook and task fixture.

This script never calls an answer-producing Agent or Harbor.  It treats the
review packet's workbook manifest as the artifact inventory, evaluates every
archived workbook with the task-local Judge now in the repository, and writes
an audit-only comparison.  It also evaluates every authored fixture once.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import csv
import hashlib
import json
import math
import os
import statistics
import subprocess
import sys
import tempfile
import time
import zipfile
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
TASK_ROOT = ROOT / "tasks" / "pilot_v1"
PASS_THRESHOLD = 0.70
SCOREABLE_STATUSES = {"SCORED", "NATIVE_OBJECT_CHECKED"}
HISTORICAL_TASK_INVALID = {
    "P15-A-POLICY-EIA-001": (
        "The archived answers were produced under a source-conflicted, "
        "unit-inconsistent task contract. The corrected V3 task is for future "
        "runs and cannot be applied retroactively."
    )
}
TASK_ORDER = [
    "P15-A-ENG-SIZING-001",
    "P15-A-FIN-DCF-001",
    "P15-A-FIN-DEBUG-001",
    "P15-A-POLICY-EIA-001",
    "P15-A-STAT-EXPERIMENT-001",
    "P15-B-FIN-RECON-001",
    "P15-B-HEALTH-REPORT-001",
    "P15-B-OPS-CLEAN-JOIN-001",
    "P15-B-PUBLIC-PIVOT-001",
    "P15-B-SALES-DISCOVERY-001",
    "P15-C-INVOICE-001",
    "P15-C-PO-ADDENDUM-001",
    "P15-C-QUOTE-001",
    "P15-C-RECEIPTS-001",
    "P15-C-STATEMENT-001",
]

# Per-fixture exceptions may be recorded as ``expected`` in the frozen V3
# validation receipt.  The default for authored mutants is that the delivery
# must not pass.
MUTANT_EXPECTATIONS: dict[tuple[str, str], str] = {}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet-root", type=Path, default=ROOT / "review_packet")
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--timeout-seconds", type=int, default=90)
    parser.add_argument("--output-label", default="judge_v3")
    parser.add_argument(
        "--fixture-receipt-name",
        default="judge_v3_local_validation.json",
    )
    parser.add_argument(
        "--judge-binding",
        type=Path,
        default=ROOT / "review_packet" / "data" / "judge_v3_manifest.json",
    )
    return parser.parse_args()


def evaluator_python() -> str:
    configured = os.environ.get("PILOT_PYTHON")
    if configured:
        return configured
    return sys.executable


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def write_csv(path: Path, rows: list[dict[str, Any]], fields: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=fields,
            extrasaction="ignore",
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(rows)


def compact_json(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def parse_bool(value: Any) -> bool | None:
    if isinstance(value, bool):
        return value
    normalized = str(value).strip().lower()
    if normalized == "true":
        return True
    if normalized == "false":
        return False
    return None


def parse_score(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        score = float(value)
    except (TypeError, ValueError):
        return None
    return score if math.isfinite(score) else None


def frozen_path(relative: Any, label: str) -> Path:
    """Resolve a manifest path inside the repository, rejecting path escapes."""
    if not isinstance(relative, str) or not relative:
        raise RuntimeError(f"{label}: missing relative path")
    path = Path(relative)
    if path.is_absolute() or ".." in path.parts:
        raise RuntimeError(f"{label}: unsafe frozen path {relative!r}")
    resolved = (ROOT / path).resolve()
    try:
        resolved.relative_to(ROOT.resolve())
    except ValueError as exc:
        raise RuntimeError(f"{label}: frozen path leaves repository: {relative}") from exc
    if not resolved.is_file():
        raise RuntimeError(f"{label}: frozen file is missing: {relative}")
    return resolved


def verify_hash_binding(binding: Any, label: str) -> tuple[str, str]:
    if not isinstance(binding, dict):
        raise RuntimeError(f"{label}: binding must be an object")
    relative = binding.get("path")
    expected = binding.get("sha256")
    if not isinstance(expected, str) or len(expected) != 64:
        raise RuntimeError(f"{label}: invalid sha256 binding")
    path = frozen_path(relative, label)
    current = sha256(path)
    if current != expected:
        raise RuntimeError(
            f"{label}: hash mismatch for {relative}: current={current} frozen={expected}"
        )
    return str(relative), current


def verify_frozen_judge_binding(path: Path) -> dict[str, Any]:
    """Verify every file named by the JSON freeze manifest before regrading."""
    if path.suffix.lower() != ".json":
        raise RuntimeError("Judge V3 binding must be a JSON freeze manifest")
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise RuntimeError("frozen Judge binding must be a JSON object")
    task_rows = payload.get("tasks")
    if not isinstance(task_rows, list):
        raise RuntimeError("frozen Judge binding has no task list")
    task_ids = [row.get("task_id") for row in task_rows if isinstance(row, dict)]
    duplicates = sorted(task_id for task_id, count in Counter(task_ids).items() if count > 1)
    if duplicates:
        raise RuntimeError(f"duplicate tasks in frozen Judge binding: {duplicates}")
    if len(task_rows) != len(TASK_ORDER) or set(task_ids) != set(TASK_ORDER):
        raise RuntimeError("frozen Judge binding does not contain the exact 15-task set")
    if payload.get("task_count") != len(task_rows):
        raise RuntimeError("frozen Judge binding task_count does not match its task list")

    evaluator_hashes: dict[str, str] = {}
    dependency_file_count = 0
    for row in task_rows:
        if not isinstance(row, dict):
            raise RuntimeError("frozen Judge task entry must be an object")
        task_id = str(row["task_id"])
        dependencies = row.get("dependency_files")
        if not isinstance(dependencies, list) or not dependencies:
            raise RuntimeError(f"{task_id}: frozen dependency_files is empty or invalid")
        dependency_by_path: dict[str, str] = {}
        expected_prefix = Path("tasks") / "pilot_v1" / task_id
        for index, dependency in enumerate(dependencies):
            label = f"{task_id}: dependency_files[{index}]"
            if not isinstance(dependency, dict):
                raise RuntimeError(f"{label}: entry must be an object")
            relative = dependency.get("path")
            if not isinstance(relative, str):
                raise RuntimeError(f"{label}: missing path")
            relative_path = Path(relative)
            if expected_prefix not in (relative_path, *relative_path.parents):
                raise RuntimeError(f"{label}: dependency is outside its task: {relative}")
            if relative in dependency_by_path:
                raise RuntimeError(f"{task_id}: duplicate frozen dependency: {relative}")
            verified_relative, current = verify_hash_binding(dependency, label)
            frozen_bytes = dependency.get("bytes")
            if not isinstance(frozen_bytes, int) or frozen_bytes < 0:
                raise RuntimeError(f"{label}: invalid byte count")
            if frozen_path(verified_relative, label).stat().st_size != frozen_bytes:
                raise RuntimeError(f"{label}: byte count differs from frozen manifest")
            dependency_by_path[verified_relative] = current
            dependency_file_count += 1

        files = row.get("files")
        if not isinstance(files, dict) or "evaluator" not in files:
            raise RuntimeError(f"{task_id}: missing named Judge file bindings")
        for file_label, file_binding in files.items():
            label = f"{task_id}: files.{file_label}"
            if not isinstance(file_binding, dict):
                raise RuntimeError(f"{label}: binding must be an object")
            relative = file_binding.get("path")
            expected = file_binding.get("sha256")
            if relative not in dependency_by_path:
                raise RuntimeError(f"{label}: named file is not in dependency_files")
            if expected != dependency_by_path[relative]:
                raise RuntimeError(f"{label}: named hash differs from dependency hash")
        evaluator_binding = files["evaluator"]
        evaluator_hash = evaluator_binding["sha256"]
        if row.get("evaluator_sha256") != evaluator_hash:
            raise RuntimeError(f"{task_id}: evaluator_sha256 disagrees with files.evaluator")
        evaluator_hashes[task_id] = evaluator_hash

    driver_path, driver_hash = verify_hash_binding(
        payload.get("regrade_driver"), "regrade_driver"
    )
    contract_path, contract_hash = verify_hash_binding(
        payload.get("adjudication_contract"), "adjudication_contract"
    )
    return {
        "manifest_sha256": sha256(path),
        "dependency_file_count": dependency_file_count,
        "evaluator_hashes": evaluator_hashes,
        "regrade_driver_path": driver_path,
        "regrade_driver_sha256": driver_hash,
        "adjudication_contract_path": contract_path,
        "adjudication_contract_sha256": contract_hash,
    }


def evaluator_runtime() -> dict[str, str]:
    command = evaluator_python()
    probe = subprocess.run(
        [
            command,
            "-c",
            (
                "import json,openpyxl,platform,sys;"
                "print(json.dumps({'python_command':sys.executable,"
                "'python_version':platform.python_version(),"
                "'openpyxl_version':openpyxl.__version__},sort_keys=True))"
            ),
        ],
        text=True,
        capture_output=True,
        check=False,
        timeout=30,
    )
    if probe.returncode != 0:
        raise RuntimeError(
            f"cannot inspect evaluator runtime ({command}): {probe.stderr[-1000:]}"
        )
    try:
        payload = json.loads(probe.stdout)
    except json.JSONDecodeError as exc:
        raise RuntimeError("evaluator runtime probe returned invalid JSON") from exc
    runtime_fields = ("python_command", "python_version", "openpyxl_version")
    if not all(isinstance(payload.get(key), str) and payload[key] for key in runtime_fields):
        raise RuntimeError("evaluator runtime probe is incomplete")
    return payload


def saved_validation_payload(task_id: str, receipt_name: str) -> dict[str, Any]:
    path = TASK_ROOT / task_id / "receipts" / receipt_name
    payload = json.loads(path.read_text(encoding="utf-8"))
    if payload.get("all_expectations_met") is not True:
        raise RuntimeError(f"{task_id}: saved validation receipt is not fully passing")
    recorded_task_id = payload.get("task_id")
    if recorded_task_id not in (None, task_id):
        raise RuntimeError(
            f"{task_id}: saved validation receipt names task {recorded_task_id!r}"
        )
    if task_id == "P15-A-POLICY-EIA-001":
        contract_checks = payload.get("contract_checks")
        if not isinstance(contract_checks, dict) or not contract_checks:
            raise RuntimeError(f"{task_id}: policy contract checks are missing")
        failed_checks = sorted(
            name
            for name, result in contract_checks.items()
            if not isinstance(result, dict) or result.get("pass") is not True
        )
        if failed_checks:
            raise RuntimeError(
                f"{task_id}: policy contract checks failed: {failed_checks}"
            )
    fixture_cases = payload.get("fixtures", payload.get("cases"))
    if not isinstance(fixture_cases, dict):
        raise RuntimeError(f"{task_id}: missing saved fixture validation receipt")
    new_cases = payload.get("new_cases", {})
    if not isinstance(new_cases, dict):
        raise RuntimeError(f"{task_id}: invalid generated-case validation receipt")
    payload["_fixture_cases"] = fixture_cases
    payload["_all_cases"] = {**fixture_cases, **new_cases}
    return payload


def saved_case_for(
    validation: dict[str, Any], group: str, name: str
) -> tuple[str, dict[str, Any]]:
    cases = validation["_all_cases"]
    candidates = [name, f"{group}:{name}"]
    if group == "reference":
        candidates.extend(["reference", "standard"])
    if group == "equivalent":
        stripped = name.removeprefix("candidate_")
        candidates.append(stripped)
        if stripped == "equivalent":
            candidates.append("equivalent")
        else:
            candidates.extend(["alternate_layout", "alternate_implementation"])
    for candidate in candidates:
        saved = cases.get(candidate)
        if isinstance(saved, dict):
            return candidate, saved
    raise KeyError(name)


def explicit_saved_expectation(
    validation: dict[str, Any], case_key: str, saved: dict[str, Any]
) -> list[str]:
    constraints: list[str] = []
    expected = saved.get("expected")
    if isinstance(expected, dict):
        expected_status = expected.get("native_status", expected.get("status"))
        if isinstance(expected_status, str) and expected_status:
            constraints.append(f"status:{expected_status}")
        expected_pass = parse_bool(expected.get("pass"))
        if expected_pass is not None:
            constraints.append(f"pass:{str(expected_pass).lower()}")
        if "score" in expected:
            expected_score = parse_score(expected.get("score"))
            constraints.append(
                "score:none" if expected_score is None else f"eq:{expected_score}"
            )
    if (
        isinstance(expected, str)
        and expected.partition(":")[0] in {"eq", "ge", "lt", "pass", "status"}
    ):
        constraints.append(expected)
    expected_pass = parse_bool(saved.get("expected_pass"))
    if expected_pass is not None:
        constraints.append(f"pass:{str(expected_pass).lower()}")
    reclassified = validation.get("fixture_reclassification", {})
    stem = case_key.partition(":")[2] if ":" in case_key else case_key
    if isinstance(reclassified, dict) and stem in reclassified and not constraints:
        constraints.append("pass:true")
    return list(dict.fromkeys(constraints))


def default_fixture_expectations(group: str) -> list[str]:
    if group == "reference":
        return ["eq:1.0"]
    if group == "equivalent":
        return ["ge:0.95"]
    if group == "noop":
        return ["pass:false"]
    if group == "malformed":
        return ["eq:0.0"]
    if group == "mutant":
        return ["pass:false"]
    raise RuntimeError(f"unsupported physical fixture group: {group}")


def final_fixture_expectations(
    task_id: str,
    group: str,
    explicit: list[str],
) -> list[str]:
    if task_id in HISTORICAL_TASK_INVALID:
        return ["status:TASK_INVALID", "pass:false", "score:none"]
    defaults = default_fixture_expectations(group)
    explicitly_positive = "pass:true" in explicit
    if group == "mutant" and explicitly_positive:
        defaults = []
    return list(dict.fromkeys([*defaults, *explicit]))


def expectation_class(expectations: list[str]) -> str:
    pass_constraints = {
        item.split(":", 1)[1].lower() == "true"
        for item in expectations
        if item.startswith("pass:")
    }
    if len(pass_constraints) > 1:
        raise RuntimeError(f"conflicting pass expectations: {expectations}")
    if pass_constraints:
        return "positive" if next(iter(pass_constraints)) else "negative"
    for item in expectations:
        operator, raw_value = item.split(":", 1)
        if operator == "eq":
            return "positive" if float(raw_value) >= PASS_THRESHOLD else "negative"
        if operator == "ge" and float(raw_value) >= PASS_THRESHOLD:
            return "positive"
        if operator == "lt" and float(raw_value) <= PASS_THRESHOLD:
            return "negative"
    if any(item.startswith("status:") or item == "score:none" for item in expectations):
        return "non_numeric"
    return "unspecified"


def saved_validation_matches(saved: dict[str, Any], payload: dict[str, Any]) -> bool:
    """Match every status/pass/score field recorded by the frozen receipt."""
    comparisons: list[bool] = []
    if "native_status" in saved:
        comparisons.append(str(payload.get("status", "")) == str(saved["native_status"]))
    elif "status" in saved:
        comparisons.append(str(payload.get("status", "")) == str(saved["status"]))
    if "evaluation_status" in saved:
        comparisons.append(
            str(payload.get("evaluation_status", ""))
            == str(saved["evaluation_status"])
        )
    if "pass" in saved:
        saved_pass = parse_bool(saved.get("pass"))
        comparisons.append(saved_pass is not None and payload.get("pass") is saved_pass)
    if "score" in saved:
        saved_score = parse_score(saved.get("score"))
        current_score = parse_score(payload.get("normalized_score"))
        comparisons.append(
            (saved_score is None and current_score is None)
            or (
                saved_score is not None
                and current_score is not None
                and math.isclose(saved_score, current_score, abs_tol=1e-6)
            )
        )
    return bool(comparisons) and all(comparisons)


def run_evaluator(
    task_id: str,
    candidate: Path,
    log_dir: Path,
    timeout_seconds: int,
    python_command: str,
) -> dict[str, Any]:
    task_dir = TASK_ROOT / task_id
    evaluator = task_dir / "tests" / "evaluate.py"
    started = time.monotonic()
    try:
        evaluator_env = dict(os.environ)
        # Historical Policy workbooks must never enter the opt-in corrected-task
        # branch merely because the caller's shell happened to export this flag.
        evaluator_env.pop("P15_POLICY_CORRECTED_TASK", None)
        evaluator_env.update(
            {
                "P15_VERIFIER_LOG_DIR": str(log_dir),
                "P15_EVAL_SPLIT": "dev",
            }
        )
        process = subprocess.run(
            [python_command, str(evaluator), str(candidate), "--split", "dev"],
            cwd=task_dir,
            text=True,
            capture_output=True,
            env=evaluator_env,
            check=False,
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired as exc:
        return {
            "audit_status": "EVALUATOR_TIMEOUT",
            "error": f"timed out after {timeout_seconds}s",
            "elapsed_seconds": round(time.monotonic() - started, 4),
            "stdout": (exc.stdout or "")[-2000:] if isinstance(exc.stdout, str) else "",
            "stderr": (exc.stderr or "")[-2000:] if isinstance(exc.stderr, str) else "",
        }
    elapsed = round(time.monotonic() - started, 4)
    if process.returncode != 0:
        return {
            "audit_status": "EVALUATOR_ERROR",
            "error": f"exit {process.returncode}",
            "elapsed_seconds": elapsed,
            "stdout": process.stdout[-2000:],
            "stderr": process.stderr[-2000:],
        }
    try:
        payload = json.loads(process.stdout)
    except json.JSONDecodeError as exc:
        return {
            "audit_status": "INVALID_JSON",
            "error": str(exc),
            "elapsed_seconds": elapsed,
            "stdout": process.stdout[-2000:],
            "stderr": process.stderr[-2000:],
        }
    score = payload.get("normalized_score")
    if score is not None and (
        not isinstance(score, (int, float))
        or isinstance(score, bool)
        or not math.isfinite(float(score))
    ):
        return {
            "audit_status": "INVALID_SCORE",
            "error": f"normalized_score={score!r}",
            "elapsed_seconds": elapsed,
            "payload": payload,
        }
    return {
        "audit_status": "EVALUATED",
        "elapsed_seconds": elapsed,
        "payload": payload,
    }


def fixture_candidates(task_id: str) -> list[tuple[str, str, Path, str]]:
    task_dir = TASK_ROOT / task_id
    fixtures = task_dir / "fixtures"
    rows: list[tuple[str, str, Path, str]] = [
        ("reference", "reference", task_dir / "solution" / "reference.xlsx", "eq:1.0"),
    ]
    groups = [
        ("equivalent", "ge:0.95"),
        ("noop", "pass:false"),
        ("malformed", "eq:0.0"),
    ]
    for group, expectation in groups:
        candidates = sorted((fixtures / group).glob("*.xlsx"))
        if not candidates:
            raise RuntimeError(f"{task_id}: missing {group} fixture")
        for candidate in candidates:
            name = group if len(candidates) == 1 else candidate.stem
            rows.append((group, name, candidate, expectation))
    for candidate in sorted((fixtures / "mutants").glob("*.xlsx")):
        rows.append(
            (
                "mutant",
                candidate.stem,
                candidate,
                MUTANT_EXPECTATIONS.get((task_id, candidate.stem), "pass:false"),
            )
        )
    return rows


def expectation_met(payload: dict[str, Any], expectation: str) -> bool:
    operator, raw_value = expectation.split(":", 1)
    if operator == "status":
        return str(payload.get("status", "")) == raw_value
    if operator == "pass":
        expected = raw_value.lower() == "true"
        return payload.get("pass") is expected
    if operator == "score" and raw_value == "none":
        return parse_score(payload.get("normalized_score")) is None
    score = float(payload["normalized_score"])
    value = float(raw_value)
    if operator == "eq":
        return math.isclose(score, value, abs_tol=1e-9)
    if operator == "ge":
        return score >= value
    if operator == "lt":
        return score < value
    raise ValueError(f"unsupported expectation: {expectation}")


def expectations_met(payload: dict[str, Any], expectations: list[str]) -> bool:
    return bool(expectations) and all(
        expectation_met(payload, expectation) for expectation in expectations
    )


def validate_campaign_inventory(data_dir: Path) -> dict[str, Any]:
    receipt_rows = read_csv(data_dir / "current_campaign_receipts.csv")
    manifest_rows = read_csv(data_dir / "current_campaign_workbook_manifest.csv")
    receipt_by_wave = {row["wave_id"]: row for row in receipt_rows}
    if len(receipt_by_wave) != len(receipt_rows):
        raise RuntimeError("duplicate wave_id in current campaign receipts")
    manifest_by_wave = {row["wave_id"]: row for row in manifest_rows}
    if len(manifest_by_wave) != len(manifest_rows):
        raise RuntimeError("duplicate wave_id in workbook manifest")
    if set(receipt_by_wave) != set(manifest_by_wave):
        missing_manifest = sorted(set(receipt_by_wave) - set(manifest_by_wave))
        missing_receipt = sorted(set(manifest_by_wave) - set(receipt_by_wave))
        raise RuntimeError(
            "campaign receipt and workbook manifest wave sets differ: "
            + compact_json(
                {
                    "missing_manifest": missing_manifest,
                    "missing_receipt": missing_receipt,
                }
            )
        )
    allowed_archive_statuses = {"INCLUDED", "NO_WORKBOOK_FOUND"}
    unexpected_archive_statuses = sorted(
        {row.get("archive_status", "") for row in manifest_rows}
        - allowed_archive_statuses
    )
    if unexpected_archive_statuses:
        raise RuntimeError(
            f"unexpected workbook archive statuses: {unexpected_archive_statuses}"
        )
    alignment_failures: list[dict[str, str]] = []
    for wave_id, receipt in receipt_by_wave.items():
        manifest_row = manifest_by_wave[wave_id]
        for field in ("task_id", "system", "judge_status"):
            if receipt.get(field, "") != manifest_row.get(field, ""):
                alignment_failures.append(
                    {
                        "wave_id": wave_id,
                        "field": field,
                        "receipt": receipt.get(field, ""),
                        "manifest": manifest_row.get(field, ""),
                    }
                )
        receipt_valid = parse_bool(receipt.get("valid_attempt"))
        manifest_valid = parse_bool(manifest_row.get("valid_attempt"))
        if (
            receipt_valid is None
            or manifest_valid is None
            or receipt_valid is not manifest_valid
        ):
            alignment_failures.append(
                {
                    "wave_id": wave_id,
                    "field": "valid_attempt",
                    "receipt": receipt.get("valid_attempt", ""),
                    "manifest": manifest_row.get("valid_attempt", ""),
                }
            )
    if alignment_failures:
        raise RuntimeError(
            "campaign receipt and workbook manifest fields differ: "
            + compact_json(alignment_failures[:20])
        )

    included_rows = [
        row for row in manifest_rows if row["archive_status"] == "INCLUDED"
    ]
    no_workbook_rows = [
        row for row in manifest_rows if row["archive_status"] == "NO_WORKBOOK_FOUND"
    ]
    if len({row["archive_member"] for row in included_rows}) != len(included_rows):
        raise RuntimeError("duplicate archive_member in workbook manifest")
    snapshot = json.loads((data_dir / "snapshot_metadata.json").read_text(encoding="utf-8"))
    snapshot_checks = {
        "receipt_n": len(receipt_rows),
        "workbook_archive_included_n": len(included_rows),
        "workbook_archive_missing_n": len(no_workbook_rows),
    }
    snapshot_mismatches = {
        key: {"snapshot": snapshot.get(key), "observed": observed}
        for key, observed in snapshot_checks.items()
        if snapshot.get(key) != observed
    }
    if len(manifest_rows) != len(included_rows) + len(no_workbook_rows):
        raise RuntimeError("workbook manifest does not partition into included and missing")
    if len(receipt_rows) != len(manifest_rows):
        raise RuntimeError("receipt and workbook manifest record counts differ")
    if snapshot_mismatches:
        raise RuntimeError(
            "snapshot metadata does not match campaign inventory: "
            + compact_json(snapshot_mismatches)
        )
    return {
        "receipt_rows": receipt_rows,
        "manifest_rows": manifest_rows,
        "receipt_by_wave": receipt_by_wave,
        "included_rows": included_rows,
        "no_workbook_rows": no_workbook_rows,
        "snapshot": snapshot,
    }


def main() -> None:
    args = parse_args()
    packet_root = args.packet_root.resolve()
    data_dir = packet_root / "data"
    if not args.output_label.replace("_", "").replace("-", "").isalnum():
        raise RuntimeError("output label may contain only letters, digits, hyphens, and underscores")
    output_prefix = data_dir / args.output_label
    inventory = validate_campaign_inventory(data_dir)
    manifest_rows = inventory["manifest_rows"]
    receipt_by_wave = inventory["receipt_by_wave"]
    included_rows = inventory["included_rows"]
    no_workbook_rows = inventory["no_workbook_rows"]
    snapshot = inventory["snapshot"]

    replay_rows: list[dict[str, Any]] = []
    fixture_rows: list[dict[str, Any]] = []
    evaluator_hashes = {
        task_id: sha256(TASK_ROOT / task_id / "tests" / "evaluate.py")
        for task_id in TASK_ORDER
    }
    judge_binding = args.judge_binding.resolve()
    if not judge_binding.is_file():
        raise RuntimeError(f"missing frozen Judge binding: {judge_binding}")
    binding_verification = verify_frozen_judge_binding(judge_binding)
    if evaluator_hashes != binding_verification["evaluator_hashes"]:
        raise RuntimeError("current evaluator hashes disagree with verified freeze manifest")
    runtime = evaluator_runtime()
    runtime_python = runtime["python_command"]

    with tempfile.TemporaryDirectory(prefix="p15-current-judge-audit-") as temporary_name:
        temporary = Path(temporary_name)
        extracted = temporary / "archive"
        archive_path = data_dir / "current_campaign_answer_workbooks.zip"
        archive_sha256 = sha256(archive_path)
        if snapshot.get("workbook_archive_sha256") != archive_sha256:
            raise RuntimeError(
                "workbook archive hash differs from snapshot metadata: "
                f"current={archive_sha256} snapshot={snapshot.get('workbook_archive_sha256')}"
            )
        scheduled: list[tuple[dict[str, str], dict[str, str], Path]] = []
        with zipfile.ZipFile(archive_path) as archive:
            infos = archive.infolist()
            archive_names = [info.filename for info in infos if not info.is_dir()]
            if len(archive_names) != len(set(archive_names)):
                raise RuntimeError("duplicate member name in workbook archive")
            archive_info = {info.filename: info for info in infos if not info.is_dir()}
            manifest_names = {row["archive_member"] for row in included_rows}
            if set(archive_names) != manifest_names:
                raise RuntimeError("workbook archive and manifest do not have the same members")
            bad_member = archive.testzip()
            if bad_member:
                raise RuntimeError(f"CRC failure in workbook archive: {bad_member}")
            for manifest_row in included_rows:
                member = manifest_row["archive_member"]
                member_path = Path(member)
                if (
                    not member.startswith("workbooks/")
                    or member_path.is_absolute()
                    or ".." in member_path.parts
                    or member_path.suffix.lower() != ".xlsx"
                    or member_path.name != "answer.xlsx"
                ):
                    raise RuntimeError(f"unsafe archive member: {member}")
                if archive_info[member].file_size != int(manifest_row["bytes"]):
                    raise RuntimeError(f"archive size differs from manifest: {member}")
                if manifest_row["wave_id"] not in receipt_by_wave:
                    raise RuntimeError(f"manifest wave missing campaign receipt: {manifest_row['wave_id']}")
                archive.extract(member, extracted)
                scheduled.append(
                    (manifest_row, receipt_by_wave[manifest_row["wave_id"]], extracted / member)
                )

        def evaluate_workbook(
            item: tuple[dict[str, str], dict[str, str], Path]
        ) -> dict[str, Any]:
            manifest_row, receipt, candidate = item
            task_id = manifest_row["task_id"]
            original_score = parse_score(receipt.get("normalized_score"))
            invalid_reason = HISTORICAL_TASK_INVALID.get(task_id)
            if invalid_reason:
                result = {
                    "audit_status": "EVALUATED",
                    "elapsed_seconds": 0.0,
                    "payload": {
                        "status": "TASK_INVALID",
                        "evaluation_status": "TASK_INVALID",
                        "normalized_score": None,
                        "pass": False,
                        "criterion_scores": {},
                        "failure_codes": [
                            "HISTORICAL_TASK_CONTRACT_INVALID:SOURCE_UNITS_AND_BASELINE"
                        ],
                        "adjudication_note": invalid_reason,
                    },
                }
            else:
                result = run_evaluator(
                    task_id,
                    candidate,
                    temporary / "logs" / "workbooks" / manifest_row["wave_id"],
                    args.timeout_seconds,
                    runtime_python,
                )
            base: dict[str, Any] = {
                "wave_id": manifest_row["wave_id"],
                "task_id": task_id,
                "system": manifest_row["system"],
                "valid_attempt": manifest_row["valid_attempt"],
                "artifact_role": manifest_row["artifact_role"],
                "archive_member": manifest_row["archive_member"],
                "artifact_sha256": sha256(candidate),
                "original_judge_status": receipt.get("judge_status", ""),
                "original_score": "" if original_score is None else original_score,
                "original_score_available": original_score is not None,
                "original_recorded_pass": (
                    ""
                    if parse_bool(receipt.get("passed_threshold")) is None
                    else parse_bool(receipt.get("passed_threshold"))
                ),
                "audit_status": result["audit_status"],
                "replay_score": "",
                "score_delta": "",
                "score_comparison": "UNVERIFIED",
                "replay_pass": "",
                "replay_zero": "",
                "replay_scoreable": "",
                "criterion_scores": "{}",
                "failure_codes": "[]",
                "evaluator_status": "",
                "evaluator_sha256": evaluator_hashes[task_id],
                "judge_binding_manifest_sha256": binding_verification[
                    "manifest_sha256"
                ],
                "error": result.get("error", ""),
                "elapsed_seconds": result.get("elapsed_seconds", ""),
            }
            if result["audit_status"] != "EVALUATED":
                return base
            payload = result["payload"]
            evaluator_status = str(payload.get("status", ""))
            score_value = payload.get("normalized_score")
            replay_scoreable = (
                evaluator_status in SCOREABLE_STATUSES
                and isinstance(score_value, (int, float))
                and not isinstance(score_value, bool)
                and math.isfinite(float(score_value))
            )
            base.update(
                {
                    "audit_status": (
                        "SCORE_REPLAYED"
                        if replay_scoreable
                        else "NATIVE_RECALC_REQUIRED"
                        if evaluator_status == "NATIVE_RECALC_REQUIRED"
                        else "JUDGE_UNABLE_TO_SCORE"
                        if evaluator_status in {"JUDGE_ERROR", "TASK_INVALID"}
                        else "UNEXPECTED_EVALUATOR_STATUS"
                    ),
                    "replay_scoreable": replay_scoreable,
                    "criterion_scores": compact_json(payload.get("criterion_scores", {})),
                    "failure_codes": compact_json(payload.get("failure_codes", [])),
                    "evaluator_status": evaluator_status,
                    "error": "",
                }
            )
            if replay_scoreable:
                replay = float(score_value)
                delta = replay - original_score if original_score is not None else None
                base.update(
                    {
                        "replay_score": round(replay, 9),
                        "score_delta": "" if delta is None else round(delta, 9),
                        "score_comparison": (
                            "NO_ORIGINAL_SCORE"
                            if delta is None
                            else "MATCH"
                            if math.isclose(delta, 0.0, abs_tol=1e-6)
                            else "CHANGED"
                        ),
                        "replay_pass": bool(payload.get("pass", replay >= PASS_THRESHOLD)),
                        "replay_zero": math.isclose(replay, 0.0, abs_tol=1e-9),
                    }
                )
            else:
                base["score_comparison"] = (
                    "NATIVE_RECALC_REQUIRED"
                    if evaluator_status == "NATIVE_RECALC_REQUIRED"
                    else "UNSCOREABLE"
                )
            return base

        with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
            replay_rows.extend(pool.map(evaluate_workbook, scheduled))

        saved_validation = {
            task_id: saved_validation_payload(task_id, args.fixture_receipt_name)
            for task_id in TASK_ORDER
        }
        fixture_jobs: list[
            tuple[str, str, str, Path, list[str], str, dict[str, Any]]
        ] = []
        matched_saved_case_keys: defaultdict[str, set[str]] = defaultdict(set)
        for task_id in TASK_ORDER:
            for group, name, candidate, _default_expectation in fixture_candidates(task_id):
                validation = saved_validation[task_id]
                try:
                    case_key, saved = saved_case_for(validation, group, name)
                except KeyError:
                    raise RuntimeError(f"{task_id}: no saved validation receipt for {name}")
                matched_saved_case_keys[task_id].add(case_key)
                explicit_expectations = explicit_saved_expectation(
                    validation, case_key, saved
                )
                if task_id in HISTORICAL_TASK_INVALID:
                    gate = validation.get("historical_default_gate", {})
                    if isinstance(gate, dict) and gate:
                        saved = gate
                        explicit_expectations = explicit_saved_expectation(
                            validation, "historical_default_gate", saved
                        )
                expectations = final_fixture_expectations(
                    task_id, group, explicit_expectations
                )
                fixture_jobs.append(
                    (
                        task_id,
                        group,
                        name,
                        candidate,
                        expectations,
                        expectation_class(expectations),
                        saved,
                    )
                )

        def evaluate_fixture(
            item: tuple[str, str, str, Path, list[str], str, dict[str, Any]]
        ) -> dict[str, Any]:
            task_id, group, name, candidate, expectations, expected_class, saved = item
            result = run_evaluator(
                task_id,
                candidate,
                temporary / "logs" / "fixtures" / task_id / group / name,
                args.timeout_seconds,
                runtime_python,
            )
            row: dict[str, Any] = {
                "task_id": task_id,
                "fixture_group": group,
                "fixture_name": name,
                "fixture_evidence_scope": "LIVE_PHYSICAL_REPLAY",
                "live_replayed": True,
                "expected_case_class": expected_class,
                "frozen_expectation": compact_json(expectations),
                "audit_status": result["audit_status"],
                "normalized_score": "",
                "pass": "",
                "frozen_contract_ok": False,
                "saved_validation_score": saved.get("score", ""),
                "saved_validation_pass": saved.get("pass", ""),
                "saved_validation_status": saved.get(
                    "status", saved.get("native_status", saved.get("evaluation_status", ""))
                ),
                "saved_validation_match": False,
                "expectation_basis": saved.get("expectation_basis", "group default"),
                "threshold_risk": False,
                "criterion_scores": "{}",
                "failure_codes": "[]",
                "evaluator_status": "",
                "evaluator_sha256": evaluator_hashes[task_id],
                "judge_binding_manifest_sha256": binding_verification[
                    "manifest_sha256"
                ],
                "error": result.get("error", ""),
                "elapsed_seconds": result.get("elapsed_seconds", ""),
            }
            if result["audit_status"] != "EVALUATED":
                return row
            payload = result["payload"]
            evaluator_status = str(payload.get("status", ""))
            score_value = payload.get("normalized_score")
            fixture_evaluable = (
                evaluator_status in SCOREABLE_STATUSES
                and isinstance(score_value, (int, float))
                and not isinstance(score_value, bool)
                and math.isfinite(float(score_value))
            )
            if not fixture_evaluable:
                contract_ok = expectations_met(payload, expectations)
                expected_non_numeric_status = any(
                    expectation.startswith("status:") for expectation in expectations
                )
                saved_match = saved_validation_matches(saved, payload)
                row.update(
                    {
                        "audit_status": (
                            "EXPECTED_NON_NUMERIC_STATUS"
                            if contract_ok and expected_non_numeric_status
                            else
                            "JUDGE_UNABLE_TO_SCORE"
                            if evaluator_status in {"JUDGE_ERROR", "TASK_INVALID"}
                            else "UNEXPECTED_EVALUATOR_STATUS"
                        ),
                        "frozen_contract_ok": contract_ok,
                        "saved_validation_match": saved_match,
                        "criterion_scores": compact_json(payload.get("criterion_scores", {})),
                        "failure_codes": compact_json(payload.get("failure_codes", [])),
                        "evaluator_status": evaluator_status,
                        "error": "" if contract_ok else "Judge did not return the expected result",
                    }
                )
                return row
            score = float(score_value)
            current_pass = bool(payload.get("pass", score >= PASS_THRESHOLD))
            row.update(
                {
                    "audit_status": "SCORE_REPLAYED",
                    "normalized_score": round(score, 9),
                    "pass": current_pass,
                    "frozen_contract_ok": expectations_met(payload, expectations),
                    "saved_validation_match": saved_validation_matches(saved, payload),
                    "threshold_risk": (
                        expected_class == "negative"
                        and (score >= PASS_THRESHOLD or current_pass)
                    ),
                    "criterion_scores": compact_json(payload.get("criterion_scores", {})),
                    "failure_codes": compact_json(payload.get("failure_codes", [])),
                    "evaluator_status": evaluator_status,
                    "error": "",
                }
            )
            return row

        with concurrent.futures.ThreadPoolExecutor(max_workers=max(1, args.workers)) as pool:
            fixture_rows.extend(pool.map(evaluate_fixture, fixture_jobs))

        # Some V3 equivalence/error probes are deterministic workbooks generated
        # by the task-local validator.  Their exact generator and saved receipt
        # are frozen with the Judge manifest; record those checks separately
        # instead of pretending the temporary workbook is still available.
        for task_id in TASK_ORDER:
            validation = saved_validation[task_id]
            top_basis = validation.get("expectation_basis", {})
            generated_cases = {
                name: saved
                for name, saved in validation.get("new_cases", {}).items()
                if (
                    task_id in HISTORICAL_TASK_INVALID
                    or name not in matched_saved_case_keys[task_id]
                )
            }
            for name, saved in validation["_fixture_cases"].items():
                if name not in matched_saved_case_keys[task_id]:
                    generated_cases.setdefault(name, saved)
            for name, saved in sorted(generated_cases.items()):
                if not isinstance(saved, dict):
                    raise RuntimeError(f"{task_id}: invalid generated case {name}")
                score_value = parse_score(saved.get("score"))
                pass_value = parse_bool(saved.get("pass"))
                expected_text = str(saved.get("expected", "saved task-local expectation"))
                machine_expectations = explicit_saved_expectation(
                    validation, name, saved
                )
                if (
                    machine_expectations
                    and expectation_class(machine_expectations) == "positive"
                    and not any(
                        item.startswith(("eq:", "ge:", "lt:", "score:"))
                        for item in machine_expectations
                    )
                ):
                    machine_expectations.append(f"ge:{PASS_THRESHOLD}")
                saved_status = saved.get(
                    "status",
                    saved.get("native_status", saved.get("evaluation_status", "")),
                )
                claim_payload = {
                    "status": saved_status,
                    "evaluation_status": saved.get("evaluation_status", ""),
                    "normalized_score": score_value,
                    "pass": pass_value,
                }
                expectation_ok = expectations_met(
                    claim_payload, machine_expectations
                )
                expected_class = (
                    expectation_class(machine_expectations)
                    if machine_expectations
                    else "unspecified"
                )
                item_basis = saved.get("expectation_basis")
                if not item_basis and isinstance(top_basis, dict):
                    item_basis = top_basis.get(name)
                if not item_basis and isinstance(top_basis, str):
                    item_basis = top_basis
                fixture_rows.append(
                    {
                        "task_id": task_id,
                        "fixture_group": "generated_v3",
                        "fixture_name": name,
                        "fixture_evidence_scope": "SAVED_ONLY_FROZEN_VALIDATION_CLAIM",
                        "live_replayed": False,
                        "expected_case_class": expected_class,
                        "frozen_expectation": compact_json(machine_expectations),
                        "audit_status": "SAVED_ONLY_VALIDATION_CLAIM",
                        "normalized_score": "" if score_value is None else score_value,
                        "pass": "" if pass_value is None else pass_value,
                        "frozen_contract_ok": expectation_ok,
                        "saved_validation_score": "" if score_value is None else score_value,
                        "saved_validation_pass": "" if pass_value is None else pass_value,
                        "saved_validation_status": saved.get("status", ""),
                        "saved_validation_match": expectation_ok,
                        "expectation_basis": item_basis or "task-local V3 validator",
                        "threshold_risk": bool(
                            expected_class == "negative"
                            and (
                                pass_value is True
                                or (score_value is not None and score_value >= PASS_THRESHOLD)
                            )
                        ),
                        "criterion_scores": compact_json(saved.get("criterion_scores", {})),
                        "failure_codes": compact_json(saved.get("failure_codes", [])),
                        "evaluator_status": saved_status or "SAVED_VALIDATION",
                        "evaluator_sha256": evaluator_hashes[task_id],
                        "judge_binding_manifest_sha256": binding_verification[
                            "manifest_sha256"
                        ],
                        "error": (
                            ""
                            if expectation_ok
                            else "Saved-only validation claim does not satisfy its frozen expectations: "
                            + expected_text
                        ),
                        "elapsed_seconds": "",
                    }
                )

    task_index = {task_id: index for index, task_id in enumerate(TASK_ORDER)}
    replay_rows.sort(key=lambda row: (task_index[row["task_id"]], row["system"], row["wave_id"]))
    fixture_rows.sort(
        key=lambda row: (
            task_index[row["task_id"]],
            ["reference", "equivalent", "noop", "malformed", "mutant", "generated_v3"].index(row["fixture_group"]),
            row["fixture_name"],
        )
    )

    replay_fields = [
        "wave_id", "task_id", "system", "valid_attempt", "artifact_role", "archive_member",
        "artifact_sha256", "original_judge_status", "original_score", "original_score_available",
        "original_recorded_pass", "audit_status", "replay_score", "score_delta", "score_comparison", "replay_pass",
        "replay_zero", "replay_scoreable", "criterion_scores", "failure_codes", "evaluator_status",
        "evaluator_sha256", "judge_binding_manifest_sha256", "error", "elapsed_seconds",
    ]
    fixture_fields = [
        "task_id", "fixture_group", "fixture_name", "fixture_evidence_scope", "live_replayed",
        "expected_case_class", "frozen_expectation", "audit_status",
        "normalized_score", "pass", "frozen_contract_ok", "saved_validation_score",
        "saved_validation_pass", "saved_validation_status", "saved_validation_match",
        "expectation_basis", "threshold_risk", "criterion_scores",
        "failure_codes", "evaluator_status", "evaluator_sha256",
        "judge_binding_manifest_sha256", "error", "elapsed_seconds",
    ]
    write_csv(Path(f"{output_prefix}_full_archive_replay.csv"), replay_rows, replay_fields)
    with Path(f"{output_prefix}_full_archive_replay.jsonl").open("w", encoding="utf-8") as handle:
        for row in replay_rows:
            handle.write(json.dumps(row, ensure_ascii=False, sort_keys=True) + "\n")
    write_csv(Path(f"{output_prefix}_fixture_audit.csv"), fixture_rows, fixture_fields)

    task_summaries: list[dict[str, Any]] = []
    for task_id in TASK_ORDER:
        current = [row for row in replay_rows if row["task_id"] == task_id]
        evaluated = [
            row
            for row in current
            if row["audit_status"]
            in {"SCORE_REPLAYED", "JUDGE_UNABLE_TO_SCORE", "NATIVE_RECALC_REQUIRED"}
        ]
        scoreable_replays = [row for row in evaluated if bool(row["replay_scoreable"])]
        valid_scoreable_replays = [
            row
            for row in scoreable_replays
            if parse_bool(row["valid_attempt"]) is True
        ]
        fixtures = [row for row in fixture_rows if row["task_id"] == task_id]
        original_scores = [
            float(row["original_score"])
            for row in evaluated
            if bool(row["original_score_available"])
        ]
        replay_scores = [float(row["replay_score"]) for row in scoreable_replays]
        task_summaries.append(
            {
                "task_id": task_id,
                "archived_workbook_n": len(current),
                "valid_workbook_n": sum(parse_bool(row["valid_attempt"]) is True for row in current),
                "original_score_n": len(original_scores),
                "evaluated_n": len(evaluated),
                "replay_scoreable_n": len(scoreable_replays),
                "valid_replay_scoreable_n": len(valid_scoreable_replays),
                "replay_judge_error_n": sum(row["evaluator_status"] == "JUDGE_ERROR" for row in evaluated),
                "replay_task_invalid_n": sum(row["evaluator_status"] == "TASK_INVALID" for row in evaluated),
                "native_recalc_required_n": sum(
                    row["evaluator_status"] == "NATIVE_RECALC_REQUIRED" for row in evaluated
                ),
                "artifact_missing_n": sum(row["audit_status"] == "ARTIFACT_MISSING" for row in current),
                "evaluator_error_n": sum(
                    row["audit_status"]
                    not in {
                        "SCORE_REPLAYED",
                        "JUDGE_UNABLE_TO_SCORE",
                        "NATIVE_RECALC_REQUIRED",
                        "ARTIFACT_MISSING",
                    }
                    for row in current
                ),
                "score_match_n": sum(row["score_comparison"] == "MATCH" for row in evaluated),
                "score_changed_n": sum(row["score_comparison"] == "CHANGED" for row in evaluated),
                "original_zero_n": sum(math.isclose(score, 0.0, abs_tol=1e-9) for score in original_scores),
                "replay_zero_n": sum(math.isclose(score, 0.0, abs_tol=1e-9) for score in replay_scores),
                "original_score_ge_threshold_n": sum(
                    score >= PASS_THRESHOLD for score in original_scores
                ),
                "original_recorded_pass_n": sum(
                    parse_bool(row["original_recorded_pass"]) is True for row in current
                ),
                "valid_original_recorded_pass_n": sum(
                    parse_bool(row["valid_attempt"]) is True
                    and parse_bool(row["original_recorded_pass"]) is True
                    for row in current
                ),
                "replay_pass_n": sum(bool(row["replay_pass"]) for row in scoreable_replays),
                "valid_replay_pass_n": sum(
                    bool(row["replay_pass"]) for row in valid_scoreable_replays
                ),
                "original_mean": round(statistics.fmean(original_scores), 9) if original_scores else "",
                "replay_mean": round(statistics.fmean(replay_scores), 9) if replay_scores else "",
                "valid_replay_mean": (
                    round(
                        statistics.fmean(
                            float(row["replay_score"])
                            for row in valid_scoreable_replays
                        ),
                        9,
                    )
                    if valid_scoreable_replays
                    else ""
                ),
                "fixture_n": len(fixtures),
                "fixture_frozen_contract_pass_n": sum(
                    bool(row["frozen_contract_ok"]) for row in fixtures
                ),
                "fixture_frozen_contract_fail_n": sum(
                    not bool(row["frozen_contract_ok"]) for row in fixtures
                ),
                "fixture_saved_validation_match_n": sum(
                    bool(row["saved_validation_match"]) for row in fixtures
                ),
                "fixture_threshold_risk_n": sum(
                    bool(row["threshold_risk"]) for row in fixtures
                ),
                "evaluator_sha256": evaluator_hashes[task_id],
            }
        )
    task_fields = list(task_summaries[0])
    write_csv(Path(f"{output_prefix}_task_audit_summary.csv"), task_summaries, task_fields)

    system_summaries: list[dict[str, Any]] = []
    for system in sorted({row["system"] for row in replay_rows}):
        current = [row for row in replay_rows if row["system"] == system]
        scoreable = [row for row in current if bool(row["replay_scoreable"])]
        valid_scoreable = [
            row for row in scoreable if parse_bool(row["valid_attempt"]) is True
        ]
        system_summaries.append(
            {
                "system": system,
                "archived_workbook_n": len(current),
                "valid_workbook_n": sum(
                    parse_bool(row["valid_attempt"]) is True for row in current
                ),
                "replay_scoreable_n": len(scoreable),
                "valid_replay_scoreable_n": len(valid_scoreable),
                "original_score_ge_threshold_n": sum(
                    parse_score(row["original_score"]) is not None
                    and float(row["original_score"]) >= PASS_THRESHOLD
                    for row in current
                ),
                "original_recorded_pass_n": sum(
                    parse_bool(row["original_recorded_pass"]) is True for row in current
                ),
                "valid_original_recorded_pass_n": sum(
                    parse_bool(row["valid_attempt"]) is True
                    and parse_bool(row["original_recorded_pass"]) is True
                    for row in current
                ),
                "replay_judge_error_n": sum(
                    row["evaluator_status"] == "JUDGE_ERROR" for row in current
                ),
                "replay_task_invalid_n": sum(
                    row["evaluator_status"] == "TASK_INVALID" for row in current
                ),
                "native_recalc_required_n": sum(
                    row["evaluator_status"] == "NATIVE_RECALC_REQUIRED" for row in current
                ),
                "replay_zero_n": sum(bool(row["replay_zero"]) for row in scoreable),
                "replay_pass_n": sum(bool(row["replay_pass"]) for row in scoreable),
                "valid_replay_pass_n": sum(
                    bool(row["replay_pass"]) for row in valid_scoreable
                ),
                "valid_replay_mean": (
                    round(
                        statistics.fmean(
                            float(row["replay_score"]) for row in valid_scoreable
                        ),
                        9,
                    )
                    if valid_scoreable
                    else ""
                ),
            }
        )
    write_csv(
        Path(f"{output_prefix}_system_audit_summary.csv"),
        system_summaries,
        list(system_summaries[0]),
    )

    no_workbook_audit = []
    for manifest_row in no_workbook_rows:
        receipt = receipt_by_wave[manifest_row["wave_id"]]
        no_workbook_audit.append(
            {
                "wave_id": manifest_row["wave_id"],
                "task_id": manifest_row["task_id"],
                "system": manifest_row["system"],
                "valid_attempt": manifest_row["valid_attempt"],
                "judge_status": receipt.get("judge_status", ""),
                "normalized_score": receipt.get("normalized_score", ""),
                "infra_error": receipt.get("infra_error", ""),
                "invalidation_category": receipt.get("invalidation_category", ""),
                "audit_classification": "NO_WORKBOOK_FOUND",
                "score_without_workbook": parse_score(
                    receipt.get("normalized_score")
                )
                is not None,
                "validity_flag_error": parse_bool(manifest_row["valid_attempt"]) is True,
            }
        )
    write_csv(
        Path(f"{output_prefix}_no_workbook.csv"),
        no_workbook_audit,
        list(no_workbook_audit[0]),
    )

    evaluated_all = [
        row
        for row in replay_rows
        if row["audit_status"]
        in {"SCORE_REPLAYED", "JUDGE_UNABLE_TO_SCORE", "NATIVE_RECALC_REQUIRED"}
    ]
    scoreable_replayed_all = [row for row in evaluated_all if bool(row["replay_scoreable"])]
    valid_scoreable_replayed_all = [
        row
        for row in scoreable_replayed_all
        if parse_bool(row["valid_attempt"]) is True
    ]
    original_scored_all = [
        row for row in evaluated_all if bool(row["original_score_available"])
    ]
    no_workbook_validity_errors = [
        row for row in no_workbook_rows if parse_bool(row["valid_attempt"]) is True
    ]
    replay_execution_failures = [
        row
        for row in replay_rows
        if row["audit_status"]
        not in {"SCORE_REPLAYED", "JUDGE_UNABLE_TO_SCORE", "NATIVE_RECALC_REQUIRED"}
    ]
    fixture_contract_failures = [
        row for row in fixture_rows if not bool(row["frozen_contract_ok"])
    ]
    fixture_saved_mismatches = [
        row for row in fixture_rows if not bool(row["saved_validation_match"])
    ]
    live_fixture_rows = [
        row for row in fixture_rows if bool(row["live_replayed"])
    ]
    saved_only_fixture_rows = [
        row for row in fixture_rows if not bool(row["live_replayed"])
    ]
    blocking_failures: list[str] = []
    if replay_execution_failures:
        blocking_failures.append(
            f"{len(replay_execution_failures)} evaluator executions/statuses failed"
        )
    if fixture_contract_failures:
        blocking_failures.append(
            f"{len(fixture_contract_failures)} fixtures violate frozen expectations"
        )
    if fixture_saved_mismatches:
        blocking_failures.append(
            f"{len(fixture_saved_mismatches)} fixtures differ from frozen validation records"
        )
    summary = {
        "schema_version": "p15_current_judge_audit_v3",
        "judge_label": args.output_label,
        "generated_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "agent_or_harbor_invoked": False,
        "audit_gate_passed": not blocking_failures,
        "blocking_failures": blocking_failures,
        "manifest_records": len(manifest_rows),
        "archived_workbooks": len(included_rows),
        "no_workbook_records": len(no_workbook_rows),
        "no_workbook_marked_valid": len(no_workbook_validity_errors),
        "score_without_workbook_n": sum(
            bool(row["score_without_workbook"]) for row in no_workbook_audit
        ),
        "valid_without_workbook_n": len(no_workbook_validity_errors),
        "valid_archived_workbooks": sum(
            parse_bool(row["valid_attempt"]) is True for row in included_rows
        ),
        "original_score_records_on_archived_workbooks": len(original_scored_all),
        "evaluated_workbooks": len(evaluated_all),
        "judge_run_completed_workbooks": len(replay_rows),
        "non_numeric_after_replay": sum(
            not bool(row["replay_scoreable"]) for row in replay_rows
        ),
        "unverified_workbooks": sum(
            not bool(row["replay_scoreable"]) for row in replay_rows
        ),
        "current_scoreable_after_replay": len(scoreable_replayed_all),
        "valid_current_scoreable_after_replay": len(valid_scoreable_replayed_all),
        "current_judge_error_after_replay": sum(row["evaluator_status"] == "JUDGE_ERROR" for row in evaluated_all),
        "current_task_invalid_after_replay": sum(row["evaluator_status"] == "TASK_INVALID" for row in evaluated_all),
        "current_native_recalc_required": sum(
            row["evaluator_status"] == "NATIVE_RECALC_REQUIRED" for row in evaluated_all
        ),
        "score_matches": sum(row["score_comparison"] == "MATCH" for row in original_scored_all),
        "score_changes": sum(row["score_comparison"] == "CHANGED" for row in original_scored_all),
        "original_zero_on_evaluated": sum(
            math.isclose(float(row["original_score"]), 0.0, abs_tol=1e-9)
            for row in original_scored_all
        ),
        "current_zero_on_scoreable_replayed": sum(bool(row["replay_zero"]) for row in scoreable_replayed_all),
        "original_score_ge_threshold_n": sum(
            float(row["original_score"]) >= PASS_THRESHOLD for row in original_scored_all
        ),
        "original_recorded_pass_n": sum(
            parse_bool(row["original_recorded_pass"]) is True for row in replay_rows
        ),
        "valid_original_recorded_pass_n": sum(
            parse_bool(row["valid_attempt"]) is True
            and parse_bool(row["original_recorded_pass"]) is True
            for row in replay_rows
        ),
        "current_pass_on_scoreable_replayed": sum(bool(row["replay_pass"]) for row in scoreable_replayed_all),
        "valid_current_pass_on_scoreable_replayed": sum(
            bool(row["replay_pass"]) for row in valid_scoreable_replayed_all
        ),
        "fixture_cases": len(fixture_rows),
        "live_physical_fixture_replays": len(live_fixture_rows),
        "saved_only_validation_claims": len(saved_only_fixture_rows),
        "saved_only_evidence_limit": (
            "Frozen task-local validation claims were checked for internal consistency "
            "but were not independently replayed by this root audit."
        ),
        "fixture_frozen_contract_passes": sum(
            bool(row["frozen_contract_ok"]) for row in fixture_rows
        ),
        "fixture_frozen_contract_failures": sum(
            not bool(row["frozen_contract_ok"]) for row in fixture_rows
        ),
        "fixture_saved_validation_matches": sum(
            bool(row["saved_validation_match"]) for row in fixture_rows
        ),
        "fixture_saved_validation_mismatches": sum(
            not bool(row["saved_validation_match"]) for row in fixture_rows
        ),
        "fixture_threshold_risks": sum(
            bool(row["threshold_risk"]) for row in fixture_rows
        ),
        "audit_status_counts": dict(Counter(row["audit_status"] for row in replay_rows)),
        "artifact_role_counts": dict(Counter(row["artifact_role"] for row in replay_rows)),
        "judge_binding_source": str(judge_binding.relative_to(ROOT)),
        "judge_binding_manifest_sha256": binding_verification["manifest_sha256"],
        "judge_binding_dependency_files_verified": binding_verification[
            "dependency_file_count"
        ],
        "judge_binding_regrade_driver": {
            "path": binding_verification["regrade_driver_path"],
            "sha256": binding_verification["regrade_driver_sha256"],
        },
        "judge_binding_adjudication_contract": {
            "path": binding_verification["adjudication_contract_path"],
            "sha256": binding_verification["adjudication_contract_sha256"],
        },
        "judge_hash_binding_verified": True,
        "evaluator_sha256_by_task": evaluator_hashes,
        "workbook_archive_sha256": archive_sha256,
        "evaluator_runtime": runtime,
        "policy_corrected_task_environment_removed": True,
    }
    Path(f"{output_prefix}_full_archive_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True))
    if blocking_failures:
        raise RuntimeError("audit gate failed: " + "; ".join(blocking_failures))


if __name__ == "__main__":
    main()
