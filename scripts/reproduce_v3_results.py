#!/usr/bin/env python3
"""Replay the public P15 workbook bundle with the task-local V3 evaluators.

The script does not call any answer-producing model. It evaluates the frozen
answer workbooks already published under ``reproduction/data`` and verifies
that the 286 rows used by the current report reproduce exactly.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import csv
import hashlib
import json
import math
import os
import subprocess
import sys
import tempfile
import time
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
TASK_ROOT = ROOT / "tasks" / "pilot_v1"
DEFAULT_DATA_DIR = ROOT / "reproduction" / "data"
DEFAULT_OUTPUT_DIR = ROOT / "reproduction" / "output"
SCOREABLE_STATUSES = {"SCORED", "NATIVE_OBJECT_CHECKED"}
HISTORICAL_TASK_INVALID = {
    "P15-A-POLICY-EIA-001": (
        "这 24 份答卷来自来源和单位定义互相冲突的旧版题面，"
        "不能套用修订后的题面和 Evaluator 追溯打分。"
    )
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Re-evaluate all published P15 answer workbooks and compare scores."
    )
    parser.add_argument("--data-dir", type=Path, default=DEFAULT_DATA_DIR)
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR)
    parser.add_argument("--workers", type=int, default=4)
    parser.add_argument("--timeout-seconds", type=int, default=90)
    return parser.parse_args()


def read_csv(path: Path) -> list[dict[str, str]]:
    with path.open(newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def parse_score(value: str | None) -> float | None:
    if value in (None, ""):
        return None
    try:
        score = float(value)
    except ValueError:
        return None
    return score if math.isfinite(score) else None


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def safe_archive_member(name: str) -> bool:
    path = Path(name)
    return (
        name.startswith("workbooks/")
        and not path.is_absolute()
        and ".." not in path.parts
        and path.suffix.lower() == ".xlsx"
        and path.name == "answer.xlsx"
    )


def run_evaluator(
    row: dict[str, str],
    candidate: Path,
    log_root: Path,
    timeout_seconds: int,
) -> dict[str, Any]:
    task_id = row["task_id"]
    if task_id in HISTORICAL_TASK_INVALID:
        return {
            "runner_status": "EVALUATED",
            "payload": {
                "status": "TASK_INVALID",
                "evaluation_status": "TASK_INVALID",
                "normalized_score": None,
                "failure_codes": [
                    "HISTORICAL_TASK_CONTRACT_INVALID:SOURCE_UNITS_AND_BASELINE"
                ],
                "adjudication_note": HISTORICAL_TASK_INVALID[task_id],
            },
            "elapsed_seconds": 0.0,
            "error": "",
        }
    evaluator = TASK_ROOT / task_id / "tests" / "evaluate.py"
    if not evaluator.is_file():
        return {"runner_status": "MISSING_EVALUATOR", "error": str(evaluator)}

    env = dict(os.environ)
    env.pop("P15_POLICY_CORRECTED_TASK", None)
    env["P15_EVAL_SPLIT"] = "dev"
    env["P15_VERIFIER_LOG_DIR"] = str(log_root / row["wave_id"])
    started = time.monotonic()
    try:
        process = subprocess.run(
            [sys.executable, str(evaluator), str(candidate), "--split", "dev"],
            cwd=evaluator.parent.parent,
            env=env,
            text=True,
            capture_output=True,
            check=False,
            timeout=timeout_seconds,
        )
    except subprocess.TimeoutExpired:
        return {
            "runner_status": "TIMEOUT",
            "error": f"超过 {timeout_seconds} 秒",
            "elapsed_seconds": round(time.monotonic() - started, 4),
        }

    elapsed = round(time.monotonic() - started, 4)
    if process.returncode != 0:
        return {
            "runner_status": "PROCESS_ERROR",
            "error": f"退出码 {process.returncode}: {process.stderr[-1000:]}",
            "elapsed_seconds": elapsed,
        }
    try:
        payload = json.loads(process.stdout)
    except json.JSONDecodeError:
        return {
            "runner_status": "INVALID_JSON",
            "error": process.stdout[-1000:],
            "elapsed_seconds": elapsed,
        }
    return {
        "runner_status": "EVALUATED",
        "payload": payload,
        "elapsed_seconds": elapsed,
        "error": "",
    }


def compare_result(row: dict[str, str], result: dict[str, Any]) -> dict[str, Any]:
    output: dict[str, Any] = {
        "wave_id": row["wave_id"],
        "task_id": row["task_id"],
        "system": row["system"],
        "runner_status": result["runner_status"],
        "evaluator_status": "",
        "score": "",
        "expected_evaluator_status": row["expected_evaluator_status"],
        "expected_score": row["expected_score"],
        "current_report_included": row["current_report_included"],
        "current_report_score": row["current_report_score"],
        "expected_match": False,
        "current_report_match": False,
        "failure_codes": "[]",
        "error": result.get("error", ""),
        "elapsed_seconds": result.get("elapsed_seconds", ""),
    }
    if result["runner_status"] != "EVALUATED":
        return output

    payload = result["payload"]
    status = str(payload.get("status", payload.get("evaluation_status", "")))
    score = payload.get("normalized_score")
    actual_score = (
        float(score)
        if isinstance(score, (int, float))
        and not isinstance(score, bool)
        and math.isfinite(float(score))
        else None
    )
    expected_score = parse_score(row["expected_score"])
    status_match = status == row["expected_evaluator_status"]
    score_match = (
        (actual_score is None and expected_score is None)
        or (
            actual_score is not None
            and expected_score is not None
            and math.isclose(actual_score, expected_score, abs_tol=1e-6)
        )
    )
    included = row["current_report_included"].lower() == "true"
    report_score = parse_score(row["current_report_score"])
    current_match = (
        not included
        or (
            status in SCOREABLE_STATUSES
            and actual_score is not None
            and report_score is not None
            and math.isclose(actual_score, report_score, abs_tol=1e-6)
        )
    )
    output.update(
        {
            "evaluator_status": status,
            "score": "" if actual_score is None else f"{actual_score:.6f}",
            "expected_match": status_match and score_match,
            "current_report_match": current_match,
            "failure_codes": json.dumps(
                payload.get("failure_codes", []),
                ensure_ascii=False,
                separators=(",", ":"),
            ),
        }
    )
    return output


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    fields = list(rows[0])
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    args = parse_args()
    data_dir = args.data_dir.resolve()
    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    metadata = json.loads((data_dir / "bundle_manifest.json").read_text(encoding="utf-8"))
    archive_path = data_dir / "answer_workbooks.zip"
    observed_archive_sha = sha256(archive_path)
    if observed_archive_sha != metadata["archive_sha256"]:
        raise RuntimeError("答卷归档的 SHA-256 与 bundle_manifest.json 不一致")

    manifest_rows = read_csv(data_dir / "workbook_manifest.csv")
    if len(manifest_rows) != metadata["workbook_count"]:
        raise RuntimeError("workbook_manifest.csv 的行数与 bundle_manifest.json 不一致")
    if len({row["wave_id"] for row in manifest_rows}) != len(manifest_rows):
        raise RuntimeError("workbook_manifest.csv 中出现重复 wave_id")
    if len({row["archive_member"] for row in manifest_rows}) != len(manifest_rows):
        raise RuntimeError("workbook_manifest.csv 中出现重复 archive_member")

    with tempfile.TemporaryDirectory(prefix="p15-public-replay-") as temporary_name:
        temporary = Path(temporary_name)
        extracted_root = temporary / "workbooks"
        candidates: dict[str, Path] = {}
        with zipfile.ZipFile(archive_path) as archive:
            members = [info for info in archive.infolist() if not info.is_dir()]
            names = [info.filename for info in members]
            expected_names = {row["archive_member"] for row in manifest_rows}
            if len(names) != len(set(names)) or set(names) != expected_names:
                raise RuntimeError("答卷归档和 workbook_manifest.csv 的文件清单不一致")
            bad_member = archive.testzip()
            if bad_member:
                raise RuntimeError(f"答卷归档 CRC 检查失败: {bad_member}")
            for info in members:
                if not safe_archive_member(info.filename):
                    raise RuntimeError(f"不安全的答卷路径: {info.filename}")
                archive.extract(info, extracted_root)
                candidates[info.filename] = extracted_root / info.filename

        def evaluate(row: dict[str, str]) -> dict[str, Any]:
            result = run_evaluator(
                row,
                candidates[row["archive_member"]],
                temporary / "logs",
                args.timeout_seconds,
            )
            return compare_result(row, result)

        with concurrent.futures.ThreadPoolExecutor(
            max_workers=max(1, args.workers)
        ) as pool:
            output_rows = list(pool.map(evaluate, manifest_rows))

    output_rows.sort(key=lambda row: (row["task_id"], row["system"], row["wave_id"]))
    write_csv(output_dir / "replayed_scores.csv", output_rows)

    status_counts = Counter(row["evaluator_status"] for row in output_rows)
    runner_errors = sum(row["runner_status"] != "EVALUATED" for row in output_rows)
    expected_mismatches = sum(not row["expected_match"] for row in output_rows)
    current_rows = [
        row for row in output_rows if row["current_report_included"].lower() == "true"
    ]
    current_mismatches = sum(not row["current_report_match"] for row in current_rows)
    summary = {
        "python_version": sys.version.split()[0],
        "workbooks_replayed": len(output_rows),
        "runner_errors": runner_errors,
        "expected_result_mismatches": expected_mismatches,
        "current_report_rows_checked": len(current_rows),
        "current_report_score_mismatches": current_mismatches,
        "evaluator_status_counts": dict(sorted(status_counts.items())),
        "archive_sha256": observed_archive_sha,
        "reproduction_ok": (
            len(output_rows) == metadata["workbook_count"]
            and len(current_rows) == metadata["current_report_workbook_count"]
            and runner_errors == 0
            and expected_mismatches == 0
            and current_mismatches == 0
        ),
    }
    (output_dir / "reproduction_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if not summary["reproduction_ok"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
