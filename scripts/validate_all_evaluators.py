#!/usr/bin/env python3
"""Run every published P15 Evaluator regression in an isolated task copy."""

from __future__ import annotations

import argparse
import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl


REPO_ROOT = Path(__file__).resolve().parents[1]
JOB_CONFIG = REPO_ROOT / "benchmark" / "configs" / "p15_v3_n8.json"
REQUIRED_FILES = (
    "task.toml",
    "instruction.md",
    "rubric.json",
    "JUDGE_CONTRACT.md",
    "solution/reference.xlsx",
    "solution/solve.sh",
    "tests/Dockerfile",
    "tests/test.sh",
    "tests/evaluate.py",
    "tests/validate_judge_v3.py",
)


def configured_tasks() -> list[Path]:
    config = json.loads(JOB_CONFIG.read_text(encoding="utf-8"))
    tasks = [REPO_ROOT / row["path"] for row in config["tasks"]]
    if len(tasks) != 15 or len({path.name for path in tasks}) != 15:
        raise RuntimeError("The frozen benchmark config must contain 15 unique tasks")
    return tasks


def validate_package(task_dir: Path) -> list[str]:
    missing = [relative for relative in REQUIRED_FILES if not (task_dir / relative).is_file()]
    if not task_dir.is_dir():
        missing.insert(0, "task directory")
    return missing


def run_validator(task_dir: Path, staging_root: Path, timeout: int) -> dict:
    staged_task = staging_root / task_dir.name
    shutil.copytree(
        task_dir,
        staged_task,
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
    )
    environment = os.environ.copy()
    environment.pop("P15_POLICY_CORRECTED_TASK", None)
    try:
        process = subprocess.run(
            [sys.executable, "validate_judge_v3.py"],
            cwd=staged_task / "tests",
            env=environment,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=False,
        )
    except subprocess.TimeoutExpired:
        return {
            "task_id": task_dir.name,
            "status": "TIMEOUT",
            "returncode": None,
            "detail": f"validator exceeded {timeout} seconds",
        }

    result = {
        "task_id": task_dir.name,
        "status": "PASS" if process.returncode == 0 else "FAIL",
        "returncode": process.returncode,
        "regression_suite": "PASS" if process.returncode == 0 else "FAIL",
    }
    if process.returncode != 0:
        combined = "\n".join(part.strip() for part in (process.stdout, process.stderr) if part.strip())
        result["detail"] = combined[-4000:] if combined else "validator returned no diagnostic output"
        return result

    reference_checks = {}
    references = {
        "dev": staged_task / "solution" / "reference.xlsx",
        "confirm": staged_task / "tests" / "confirm" / "reference.xlsx",
    }
    for split, candidate in references.items():
        reference_environment = environment.copy()
        reference_environment["P15_VERIFIER_LOG_DIR"] = str(staged_task / "validation-logs" / split)
        try:
            reference_process = subprocess.run(
                [
                    sys.executable,
                    str(staged_task / "tests" / "evaluate.py"),
                    str(candidate),
                    "--split",
                    split,
                ],
                cwd=staged_task,
                env=reference_environment,
                capture_output=True,
                text=True,
                timeout=timeout,
                check=False,
            )
        except subprocess.TimeoutExpired:
            reference_checks[split] = {"status": "TIMEOUT"}
            continue
        try:
            payload = json.loads(reference_process.stdout)
        except json.JSONDecodeError:
            payload = {}
        passed = (
            reference_process.returncode == 0
            and payload.get("pass") is True
            and payload.get("normalized_score") == 1.0
        )
        reference_checks[split] = {
            "status": "PASS" if passed else "FAIL",
            "returncode": reference_process.returncode,
            "score": payload.get("normalized_score"),
            "evaluation_status": payload.get("evaluation_status"),
        }
        if not passed:
            diagnostic = reference_process.stderr.strip() or reference_process.stdout.strip()
            reference_checks[split]["detail"] = diagnostic[-2000:] or "Evaluator returned no result"

    result["reference_checks"] = reference_checks
    failed_references = [split for split, row in reference_checks.items() if row["status"] != "PASS"]
    if failed_references:
        result["status"] = "FAIL"
        result["detail"] = "Reference checks failed: " + ", ".join(failed_references)
    return result


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Check all 15 published Evaluator packages and run their V3 regression suites."
    )
    parser.add_argument("--timeout-per-task", type=int, default=300)
    args = parser.parse_args()
    if args.timeout_per_task < 1:
        parser.error("--timeout-per-task must be at least 1 second")

    tasks = configured_tasks()
    package_failures = []
    for task in tasks:
        missing = validate_package(task)
        if missing:
            package_failures.append({"task_id": task.name, "missing": missing})
    if package_failures:
        print(
            json.dumps(
                {
                    "task_count": len(tasks),
                    "package_failures": package_failures,
                    "all_evaluators_ok": False,
                },
                indent=2,
            )
        )
        raise SystemExit(1)

    with tempfile.TemporaryDirectory(prefix="p15-evaluator-validation-") as directory:
        staging_root = Path(directory)
        results = [run_validator(task, staging_root, args.timeout_per_task) for task in tasks]

    failures = [row for row in results if row["status"] != "PASS"]
    summary = {
        "python_version": ".".join(map(str, sys.version_info[:3])),
        "openpyxl_version": openpyxl.__version__,
        "task_count": len(tasks),
        "passed": len(results) - len(failures),
        "failed": len(failures),
        "all_evaluators_ok": not failures,
        "tasks": results,
    }
    print(json.dumps(summary, indent=2, ensure_ascii=False))
    if failures:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
