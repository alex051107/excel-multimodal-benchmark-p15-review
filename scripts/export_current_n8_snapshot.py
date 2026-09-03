#!/usr/bin/env python3
"""Export a sanitized, analysis-facing snapshot of the P15 n=8 campaign.

The source campaign remains the counting authority.  This exporter deliberately
keeps four different notions separate:

* strict coverage: complete, scored attempts admitted by the current controller;
* recovered artifacts: workbooks captured from interrupted or invalid Trials;
* model non-delivery: a completed model path that did not deliver a workbook;
* infrastructure/Judge failures: Trials with no attributable model outcome.

The generated files contain no prompts, credentials, provider endpoints, local
paths, token fingerprints, raw trajectories, or model-produced workbooks.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import os
import statistics
import subprocess
import sys
import tempfile
import zipfile
from collections import defaultdict
from dataclasses import replace
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


PASS_THRESHOLD = 0.70
EXPORT_SCHEMA_VERSION = 2
CONTROLLER_RELATIVE_PATH = Path("scripts/p15_n8_campaign_control.py")


def read_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise ValueError(f"JSON object required: {path}")
    return value


def finite_score(value: Any) -> bool:
    return (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(float(value))
        and 0 <= float(value) <= 1
    )


def truthy(value: Any) -> bool:
    return str(value).strip().lower() in {"1", "true", "yes"}


def task_track(task_id: str) -> str:
    parts = task_id.split("-")
    if len(parts) < 2 or parts[1] not in {"A", "B", "C"}:
        raise ValueError(f"Cannot infer track from task id: {task_id}")
    return parts[1]


def parse_time(value: str | None) -> datetime | None:
    if not value:
        return None
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    if parsed.tzinfo is None:
        raise ValueError(f"Timestamp must include a timezone: {value!r}")
    return parsed.astimezone(timezone.utc)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def resolve_inside(source_root: Path, raw: str | Path, *, label: str) -> Path:
    value = Path(raw)
    resolved = (value if value.is_absolute() else source_root / value).resolve()
    try:
        resolved.relative_to(source_root)
    except ValueError as exc:
        raise ValueError(f"{label} escapes the source repository: {raw}") from exc
    return resolved


def repository_relative(source_root: Path, path: Path) -> str:
    return str(path.resolve().relative_to(source_root))


def valid_xlsx(path: Path) -> bool:
    if not path.is_file() or not zipfile.is_zipfile(path):
        return False
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            return "[Content_Types].xml" in names and "xl/workbook.xml" in names
    except (OSError, zipfile.BadZipFile):
        return False


def candidate_evidence(
    source_root: Path,
    run_dir: Path,
    final: dict[str, Any],
) -> tuple[Path | None, str]:
    """Find a candidate without ever accepting an out-of-repository path."""

    candidates: list[Path] = []
    trial = final.get("trial_outcome") if isinstance(final.get("trial_outcome"), dict) else {}
    for key in ("canonical_candidate_path", "candidate_path"):
        raw = trial.get(key)
        if isinstance(raw, str) and raw:
            candidates.append(resolve_inside(source_root, raw, label=f"trial_outcome.{key}"))

    attestation_path = run_dir / "judge_attestation.json"
    if attestation_path.is_file():
        attestation = read_json(attestation_path)
        candidate = attestation.get("candidate") if isinstance(attestation.get("candidate"), dict) else {}
        raw = candidate.get("path")
        if isinstance(raw, str) and raw:
            candidates.append(resolve_inside(source_root, raw, label="judge_attestation.candidate.path"))

    job_root = source_root / "outputs/p15_final/harbor_jobs" / f"p15_budgeted_{run_dir.name}"
    candidates.extend(sorted(job_root.glob("P15-*/artifacts/app/output/answer.xlsx")))

    seen: set[Path] = set()
    first_invalid: Path | None = None
    for path in candidates:
        resolved = resolve_inside(source_root, path, label="candidate workbook")
        if resolved in seen:
            continue
        seen.add(resolved)
        if not resolved.is_file():
            continue
        if valid_xlsx(resolved):
            return resolved, "VALID_XLSX"
        if first_invalid is None:
            first_invalid = resolved
    return (first_invalid, "INVALID_XLSX") if first_invalid else (None, "MISSING")


def baseline_candidate_evidence(
    source_root: Path,
    result_path_raw: str,
) -> tuple[Path | None, str]:
    if not result_path_raw:
        return None, "MISSING_RESULT_RECEIPT"
    result_path = resolve_inside(source_root, result_path_raw, label="baseline result_path")
    candidate_path = resolve_inside(
        source_root,
        result_path.parent / "artifacts/app/output/answer.xlsx",
        label="baseline candidate workbook",
    )
    if not candidate_path.is_file():
        return None, "MISSING"
    if not valid_xlsx(candidate_path):
        return candidate_path, "INVALID_XLSX"
    return candidate_path, "VALID_XLSX"


def harbor_trial_directory(source_root: Path, run_dir: Path) -> Path | None:
    job_root = source_root / "outputs/p15_final/harbor_jobs" / f"p15_budgeted_{run_dir.name}"
    trials = sorted(path for path in job_root.glob("P15-*") if path.is_dir())
    if len(trials) > 1:
        raise ValueError(f"Runtime wave has more than one Harbor Trial: {run_dir.name}")
    return trials[0] if trials else None


def agent_completion_evidence(source_root: Path, run_dir: Path) -> tuple[str, str]:
    """Conservatively identify whether the Agent reached its own terminal turn."""

    trial_dir = harbor_trial_directory(source_root, run_dir)
    if trial_dir is None:
        return "unknown", "NO_HARBOR_TRIAL_EVIDENCE"

    trajectory_path = trial_dir / "agent/trajectory.json"
    if trajectory_path.is_file():
        trajectory = read_json(trajectory_path)
        steps = trajectory.get("steps") if isinstance(trajectory.get("steps"), list) else []
        for step in reversed(steps):
            if not isinstance(step, dict) or step.get("source") != "agent":
                continue
            extra = step.get("extra") if isinstance(step.get("extra"), dict) else {}
            stop_reason = str(extra.get("stop_reason") or "").strip().lower()
            if stop_reason == "end_turn":
                return "true", "TRAJECTORY_END_TURN"
            if stop_reason in {"tool_use", "max_tokens"}:
                return "false", f"TRAJECTORY_{stop_reason.upper()}"
            break

    result_path = trial_dir / "result.json"
    if not result_path.is_file():
        return "unknown", "NO_HARBOR_RESULT_EVIDENCE"
    result = read_json(result_path)
    agent_execution = (
        result.get("agent_execution")
        if isinstance(result.get("agent_execution"), dict)
        else {}
    )
    agent_finished = parse_time(str(agent_execution.get("finished_at") or ""))
    exception = result.get("exception_info") if isinstance(result.get("exception_info"), dict) else {}
    traceback = str(exception.get("exception_traceback") or "")
    verifier = result.get("verifier") if isinstance(result.get("verifier"), dict) else {}
    verifier_started = parse_time(str(verifier.get("started_at") or ""))
    if agent_finished is not None and (
        not exception or verifier_started is not None or "_run_verifier" in traceback
    ):
        return "true", "HARBOR_AGENT_PHASE_FINISHED"
    if "_run_agent" in traceback:
        return "false", "HARBOR_AGENT_PHASE_INTERRUPTED"
    return "unknown", "AGENT_TERMINAL_STATE_UNRESOLVED"


def native_receipt_consistency(run_dir: Path, final: dict[str, Any]) -> tuple[str, str]:
    """Keep a failed native-Excel canonicalization from inheriting a later score."""

    receipt_path = run_dir / "native_excel/receipt.json"
    if receipt_path.is_file():
        receipt = read_json(receipt_path)
    else:
        receipt = final.get("native_excel") if isinstance(final.get("native_excel"), dict) else {}
    if not receipt:
        return "unknown", "NO_NATIVE_RECEIPT"
    status = str(receipt.get("status") or "").strip().upper()
    if status == "CANONICAL_EXCEL_ERROR" or receipt.get("error"):
        return "false", status or "NATIVE_RECEIPT_ERROR"
    return "true", status or "NATIVE_RECEIPT_PRESENT"


def git_output(source_root: Path, *args: str) -> str:
    return subprocess.check_output(
        ["git", "-C", str(source_root), *args],
        text=True,
        stderr=subprocess.STDOUT,
    ).strip()


def validate_source_controller(source_root: Path, expected_head: str) -> tuple[str, str]:
    expected = expected_head.strip().lower()
    if len(expected) != 40 or any(char not in "0123456789abcdef" for char in expected):
        raise ValueError("--source-controller-head must be a full 40-character Git SHA")
    actual = git_output(source_root, "rev-parse", "HEAD").lower()
    if actual != expected:
        raise ValueError(f"Source repository HEAD mismatch: expected {expected}, found {actual}")

    controller_path = source_root / CONTROLLER_RELATIVE_PATH
    working_hash = sha256_file(controller_path)
    committed_bytes = subprocess.check_output(
        ["git", "-C", str(source_root), "show", f"{actual}:{CONTROLLER_RELATIVE_PATH.as_posix()}"],
        stderr=subprocess.STDOUT,
    )
    committed_hash = hashlib.sha256(committed_bytes).hexdigest()
    if working_hash != committed_hash:
        raise ValueError(
            "The imported campaign controller differs from the recorded source HEAD; commit it before export"
        )
    return actual, working_hash


def runtime_records_at_cutoff(
    controller: Any,
    record_root: Path,
    config: dict[str, Any],
    source_root: Path,
    cutoff: datetime,
) -> tuple[list[Any], int]:
    """Load only terminal receipts at the cutoff, isolating later receipts."""

    all_final_paths = sorted(record_root.glob("*/final.json"))
    included_paths: list[Path] = []
    for path in all_final_paths:
        payload = read_json(path)
        finished_at = parse_time(str(payload.get("finished_at") or ""))
        if finished_at is None:
            raise ValueError(f"Runtime receipt has no finished_at: {path}")
        if finished_at <= cutoff:
            included_paths.append(path.resolve())

    original_by_wave = {path.parent.name: path for path in included_paths}
    if len(original_by_wave) != len(included_paths):
        raise ValueError("Duplicate runtime wave directory name at evidence cutoff")

    with tempfile.TemporaryDirectory(prefix="p15-export-cutoff-") as temporary:
        isolated_root = Path(temporary)
        for path in included_paths:
            os.symlink(path.parent, isolated_root / path.parent.name, target_is_directory=True)
        isolated_records = controller.read_runtime_records(isolated_root, config, source_root)

    records = [
        replace(record, path=str(original_by_wave[Path(record.path).parent.name]))
        for record in isolated_records
    ]
    return records, len(all_final_paths) - len(included_paths)


def duration_from_run(run_dir: Path, final: dict[str, Any]) -> float | None:
    preflight = run_dir / "preflight.json"
    if not preflight.is_file():
        return None
    start = parse_time(read_json(preflight).get("started_at"))
    end = parse_time(final.get("finished_at"))
    if start is None or end is None:
        return None
    return max(0.0, (end - start).total_seconds())


def runtime_score(controller: Any, source_root: Path, final_path: Path, final: dict[str, Any]) -> tuple[float | None, str]:
    status = final.get("judge_status")
    current_guard = bool(
        final.get("receipt_schema_version") == 2
        and final.get("receipt_source") == "GUARD_RUNTIME"
        and final.get("experiment_valid") is True
        and final.get("infra_error") is not True
        and final.get("harbor_output_valid") is True
        and not controller.guard_forced_stop(final)
        and status in controller.COUNTABLE_JUDGE_STATUSES
        and finite_score(final.get("normalized_score"))
    )
    if current_guard:
        return float(final["normalized_score"]), str(final.get("normalized_score_source") or "GUARD_RUNTIME")

    attestation_path = final_path.parent / "judge_attestation.json"
    if attestation_path.is_file() and controller.countable_judge_attestation(
        final_path, repo_root=source_root, expected_task_id=str(final.get("task_id") or "")
    ):
        attestation = read_json(attestation_path)
        return float(attestation["normalized_score"]), str(attestation.get("evidence_kind") or "JUDGE_ATTESTATION")
    return None, "UNCOUNTABLE"


def diagnostic_score(final_path: Path, final: dict[str, Any]) -> tuple[float | None, str]:
    """Preserve a provisional score for audit without making it countable."""

    if finite_score(final.get("normalized_score")):
        return float(final["normalized_score"]), str(
            final.get("normalized_score_source") or "FINAL_RECEIPT"
        )
    attestation_path = final_path.parent / "judge_attestation.json"
    if attestation_path.is_file():
        attestation = read_json(attestation_path)
        if finite_score(attestation.get("normalized_score")):
            return float(attestation["normalized_score"]), str(
                attestation.get("evidence_kind") or "JUDGE_ATTESTATION"
            )
    return None, "N/A"


def tri_state(value: Any) -> str:
    if value is True:
        return "true"
    if value is False:
        return "false"
    return "unknown"


def runtime_clean_state(controller: Any, final: dict[str, Any], *, record_infra: bool) -> str:
    if record_infra or final.get("infra_error") is True or controller.guard_forced_stop(final):
        return "false"
    if final.get("runtime_health") is True:
        return "true"
    if final.get("runtime_health") is False:
        return "false"
    return "unknown"


def write_csv(path: Path, fieldnames: list[str], rows: Iterable[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=fieldnames,
            extrasaction="ignore",
            lineterminator="\n",
        )
        writer.writeheader()
        writer.writerows(rows)


def fmt(value: float | None) -> str:
    return "" if value is None else f"{value:.6f}"


def mean(values: list[float]) -> float | None:
    return statistics.fmean(values) if values else None


def median(values: list[float]) -> float | None:
    return statistics.median(values) if values else None


def pstdev(values: list[float]) -> float | None:
    return statistics.pstdev(values) if values else None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-repo", type=Path, required=True)
    parser.add_argument("--campaign-config", type=Path, required=True)
    parser.add_argument(
        "--evidence-cutoff",
        required=True,
        help="Inclusive ISO-8601 cutoff for runtime final receipts (timezone required)",
    )
    parser.add_argument(
        "--usage-receipt",
        type=Path,
        required=True,
        help="Immutable settled latest_usage.json receipt used for total token-unit accounting",
    )
    parser.add_argument(
        "--source-controller-head",
        required=True,
        help="Full source-repository HEAD whose committed controller is imported",
    )
    parser.add_argument(
        "--expected-runtime-receipts",
        type=int,
        required=True,
        help="Fail closed unless this many runtime final receipts fall on or before the cutoff",
    )
    parser.add_argument("--output-dir", type=Path, default=Path("results"))
    args = parser.parse_args()

    source_root = args.source_repo.resolve()
    if not source_root.is_dir():
        raise ValueError(f"Source repository is not a directory: {source_root}")
    cutoff = parse_time(args.evidence_cutoff)
    if cutoff is None:
        raise ValueError("--evidence-cutoff is required")
    if args.expected_runtime_receipts < 0:
        raise ValueError("--expected-runtime-receipts must be nonnegative")
    source_head, controller_sha256 = validate_source_controller(
        source_root, args.source_controller_head
    )
    config_path = resolve_inside(source_root, args.campaign_config, label="campaign config")
    usage_receipt_path = resolve_inside(source_root, args.usage_receipt, label="usage receipt")
    if not config_path.is_file():
        raise ValueError(f"Campaign config does not exist: {config_path}")
    if not usage_receipt_path.is_file():
        raise ValueError(f"Usage receipt does not exist: {usage_receipt_path}")
    if usage_receipt_path.name != "latest_usage.json":
        raise ValueError("--usage-receipt must name an immutable per-run latest_usage.json")
    usage_final_path = usage_receipt_path.parent / "final.json"
    if not usage_final_path.is_file():
        raise ValueError("--usage-receipt must be bound to a terminal run with final.json")
    usage_final_finished = parse_time(
        str(read_json(usage_final_path).get("finished_at") or "")
    )
    if usage_final_finished is None or usage_final_finished > cutoff:
        raise ValueError("Usage receipt's terminal run is outside the fixed evidence cutoff")
    output_dir = args.output_dir.resolve()

    sys.path.insert(0, str(source_root / "scripts"))
    import p15_n8_campaign_control as controller  # type: ignore

    config = controller.load_config(config_path, source_root)
    tasks = list(controller.task_index(config))
    systems = controller.system_index(config)
    if systems.get("qwen38max", {}).get("model_name") != "openai/qwen3.8-max":
        raise ValueError("Frozen qwen38max model must be exactly openai/qwen3.8-max")
    baseline_counts = controller.load_baseline_counts(config, source_root)
    record_root = source_root / "outputs/p15_final/gateway_budget_runs"
    runtime_records, ignored_runtime_receipts = runtime_records_at_cutoff(
        controller, record_root, config, source_root, cutoff
    )
    if len(runtime_records) != args.expected_runtime_receipts:
        raise ValueError(
            "Runtime receipt count mismatch at fixed cutoff: "
            f"expected {args.expected_runtime_receipts}, found {len(runtime_records)}"
        )
    controller_counts = controller.aggregate_counts(
        baseline_counts,
        runtime_records,
        target=int(config["target_valid_attempts_per_task_system"]),
    )

    rows: list[dict[str, Any]] = []

    # Preserve every frozen baseline terminal receipt, not only the 41 counted
    # rows.  This makes missingness and infrastructure attrition visible.
    for metric_source in (config.get("reuse_policy") or {}).get("baseline_sources", []):
        attempts_path = (source_root / str(metric_source)).with_name("ATTEMPTS.csv")
        with attempts_path.open(newline="", encoding="utf-8") as handle:
            for raw in csv.DictReader(handle):
                task_id = str(raw.get("task_id") or "")
                system = str(raw.get("system") or "")
                if task_id not in tasks or system not in systems:
                    raise ValueError(f"Unexpected frozen baseline row in {attempts_path}: {(task_id, system)}")
                raw_score = float(raw["score"]) if raw.get("score") else None
                controller_countable = truthy(raw.get("valid")) and finite_score(raw_score)
                candidate_path, candidate_status = baseline_candidate_evidence(
                    source_root, str(raw.get("result_path") or "")
                )
                artifact_valid = candidate_status == "VALID_XLSX"
                score_admissible = controller_countable and artifact_valid
                score = float(raw_score) if score_admissible else None
                if artifact_valid and finite_score(raw_score):
                    model_outcome = "true"
                elif controller_countable:
                    model_outcome = "false"
                else:
                    model_outcome = "unknown"
                rows.append(
                    {
                        "attempt_id": raw.get("trial_name") or "",
                        "source_batch": "FROZEN_BASELINE",
                        "task_id": task_id,
                        "track": task_track(task_id),
                        "system": system,
                        "agent": raw.get("agent") or systems[system].get("agent") or "",
                        "model": systems[system].get("model_name") or "",
                        "observed_model": raw.get("model") or "",
                        "finished_at": "",
                        "controller_countable": str(controller_countable).lower(),
                        "agent_completed": "unknown",
                        "agent_completion_evidence": "LEGACY_NOT_RECORDED",
                        "model_artifact_outcome_valid": model_outcome,
                        "runtime_clean": "unknown",
                        "native_receipt_consistent": "unknown",
                        "native_receipt_evidence": "LEGACY_NOT_RECORDED",
                        "score_admissible": str(score_admissible).lower(),
                        "analysis_disposition": (
                            "COUNTABLE_COMPLETE_ARTIFACT"
                            if score_admissible
                            else "BASELINE_DECLARATION_WITHOUT_VALID_ARTIFACT"
                            if controller_countable
                            else "UNCOUNTABLE_BASELINE_RUNTIME"
                        ),
                        "answer_xlsx_present": str(candidate_path is not None).lower(),
                        "candidate_xlsx_status": candidate_status,
                        "candidate_sha256": sha256_file(candidate_path) if candidate_path else "",
                        "score": fmt(score),
                        "pass_at_0_70": "" if score is None else str(score >= PASS_THRESHOLD).lower(),
                        "diagnostic_score_provisional": fmt(raw_score if finite_score(raw_score) else None),
                        "diagnostic_score_source": "FROZEN_BASELINE_JUDGE" if finite_score(raw_score) else "N/A",
                        "diagnostic_exclusion_reason": "",
                        "judge_status": "SCORED" if score is not None else "N/A",
                        "score_source": "FROZEN_BASELINE_JUDGE" if score is not None else "N/A",
                        "judge_version_status": "PROVISIONAL_PRE_V2_FREEZE",
                        "runtime_health": "unknown",
                        "infra_error": "unknown",
                        "invalidation_category": raw.get("exception_type") or "",
                        "native_excel_status": "NOT_RECORDED",
                        "duration_seconds": raw.get("total_sec") or "",
                        "cost_value": raw.get("harbor_reported_cost_usd") or "",
                        "cost_unit": "USD_FRAMEWORK_FIELD" if raw.get("harbor_reported_cost_usd") else "",
                        "cost_basis": "NOT_POOLED_WITH_ZCLOUD_UNITS" if raw.get("harbor_reported_cost_usd") else "",
                    }
                )

    record_by_path = {Path(record.path).resolve(): record for record in runtime_records}
    for final_path in sorted(record_by_path, key=lambda item: (record_by_path[item].finished_at, str(item))):
        record = record_by_path[final_path]
        final = read_json(final_path)
        run_dir = final_path.parent
        countable = record.valid_attempt
        score, score_source = runtime_score(controller, source_root, final_path, final) if countable else (None, "UNCOUNTABLE")
        provisional_score, provisional_score_source = diagnostic_score(final_path, final)
        candidate_path, candidate_status = candidate_evidence(source_root, run_dir, final)
        artifact_valid = candidate_status == "VALID_XLSX"
        infra = final.get("infra_error") is True or record.infra_error
        guard_stopped = controller.guard_forced_stop(final)
        runtime_clean = runtime_clean_state(controller, final, record_infra=record.infra_error)
        agent_completed, agent_completion_source = agent_completion_evidence(
            source_root, run_dir
        )
        native_consistent, native_consistency_source = native_receipt_consistency(
            run_dir, final
        )
        diagnostic_exclusion_reason = ""
        if agent_completed == "false":
            provisional_score = None
            provisional_score_source = "N/A"
            diagnostic_exclusion_reason = "AGENT_DID_NOT_REACH_TERMINAL_TURN"
        elif native_consistent == "false":
            provisional_score = None
            provisional_score_source = "N/A"
            diagnostic_exclusion_reason = "NATIVE_EXCEL_RECEIPT_INCONSISTENT"
        attestation_path = run_dir / "judge_attestation.json"
        attestation = read_json(attestation_path) if attestation_path.is_file() else {}
        legacy_attestation_claimed_countable = attestation.get("countable") is True
        trial = final.get("trial_outcome") if isinstance(final.get("trial_outcome"), dict) else {}
        if artifact_valid and agent_completed == "true":
            model_artifact_outcome_valid = "true"
        elif artifact_valid and agent_completed == "false":
            model_artifact_outcome_valid = "false"
        elif candidate_status in {"MISSING", "INVALID_XLSX"} and runtime_clean == "true":
            model_artifact_outcome_valid = "false"
        else:
            model_artifact_outcome_valid = "unknown"
        score_admissible = bool(
            countable
            and artifact_valid
            and finite_score(score)
            and agent_completed != "false"
            and native_consistent != "false"
        )
        if score_admissible:
            disposition = "COUNTABLE_COMPLETE_ARTIFACT"
        elif countable:
            disposition = "CONTROLLER_INCONSISTENCY_SCORE_NOT_ADMISSIBLE"
        elif artifact_valid and agent_completed == "false":
            disposition = "INCOMPLETE_AGENT_ARTIFACT_FROM_ABNORMAL_TRIAL"
        elif artifact_valid and (infra or guard_stopped or runtime_clean != "true"):
            disposition = "RECOVERED_ARTIFACT_FROM_ABNORMAL_TRIAL"
        elif artifact_valid:
            disposition = "UNCOUNTABLE_ARTIFACT_PENDING_REVIEW"
        elif legacy_attestation_claimed_countable and not infra and not guard_stopped:
            disposition = "MODEL_DELIVERY_FAILURE"
        elif legacy_attestation_claimed_countable and infra:
            disposition = "EXCLUDED_INFRA_FAILURE_WITHOUT_ARTIFACT"
        elif legacy_attestation_claimed_countable:
            disposition = "EXCLUDED_NON_DELIVERY_WITHOUT_ARTIFACT"
        elif runtime_clean == "true":
            disposition = "MODEL_DELIVERY_FAILURE"
        elif infra or guard_stopped or runtime_clean == "false":
            disposition = "EXCLUDED_INFRA_FAILURE_WITHOUT_ARTIFACT"
        else:
            disposition = "UNCOUNTABLE_RUNTIME_OR_JUDGE"

        native = final.get("native_excel") if isinstance(final.get("native_excel"), dict) else {}
        rows.append(
            {
                "attempt_id": run_dir.name,
                "source_batch": "CURRENT_ZCLOUD_RUNTIME",
                "task_id": record.task_id,
                "track": task_track(record.task_id),
                "system": record.system,
                "agent": record.agent,
                "model": systems[record.system].get("model_name") or "",
                "observed_model": final.get("model_name") or "",
                "finished_at": final.get("finished_at") or "",
                "controller_countable": str(countable).lower(),
                "legacy_attestation_claimed_countable": str(legacy_attestation_claimed_countable).lower(),
                "agent_completed": agent_completed,
                "agent_completion_evidence": agent_completion_source,
                "model_artifact_outcome_valid": model_artifact_outcome_valid,
                "runtime_clean": runtime_clean,
                "native_receipt_consistent": native_consistent,
                "native_receipt_evidence": native_consistency_source,
                "score_admissible": str(score_admissible).lower(),
                "analysis_disposition": disposition,
                "answer_xlsx_present": str(candidate_path is not None).lower(),
                "candidate_xlsx_status": candidate_status,
                "candidate_sha256": sha256_file(candidate_path) if candidate_path else "",
                "score": fmt(score),
                "pass_at_0_70": "" if score is None else str(score >= PASS_THRESHOLD).lower(),
                "diagnostic_score_provisional": fmt(provisional_score),
                "diagnostic_score_source": provisional_score_source,
                "diagnostic_exclusion_reason": diagnostic_exclusion_reason,
                "judge_status": final.get("judge_status") or "N/A",
                "score_source": score_source,
                "judge_version_status": "PROVISIONAL_PRE_V2_FREEZE",
                "runtime_health": tri_state(final.get("runtime_health")),
                "infra_error": str(infra).lower(),
                "invalidation_category": final.get("invalidation_category") or ("GUARD_FORCED_STOP" if guard_stopped else ""),
                "native_excel_status": native.get("status") or ("PENDING_WINDOWS_COMPATIBILITY" if record.task_id == "P15-B-PUBLIC-PIVOT-001" else "NOT_APPLICABLE"),
                "duration_seconds": fmt(duration_from_run(run_dir, final)),
                "cost_value": fmt(record.spent_units),
                "cost_unit": "ZCLOUD_DISPLAYED_UNIT",
                "cost_basis": "RUN_DELTA_MAY_OVERLAP_DO_NOT_SUM",
            }
        )

    attempts_fields = [
        "attempt_id", "source_batch", "task_id", "track", "system", "agent", "model",
        "observed_model", "finished_at", "controller_countable", "legacy_attestation_claimed_countable",
        "agent_completed", "agent_completion_evidence", "model_artifact_outcome_valid",
        "runtime_clean", "native_receipt_consistent", "native_receipt_evidence", "score_admissible",
        "analysis_disposition", "answer_xlsx_present", "candidate_xlsx_status", "candidate_sha256",
        "score", "pass_at_0_70", "diagnostic_score_provisional", "diagnostic_score_source",
        "diagnostic_exclusion_reason",
        "judge_status", "score_source", "judge_version_status",
        "runtime_health", "infra_error", "invalidation_category", "native_excel_status",
        "duration_seconds", "cost_value", "cost_unit", "cost_basis",
    ]
    write_csv(output_dir / "N8_ATTEMPTS.csv", attempts_fields, rows)

    included: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    for source in ("FROZEN_BASELINE", "CURRENT_ZCLOUD_RUNTIME"):
        source_rows = sorted(
            (
                row
                for row in rows
                if row["source_batch"] == source and row["score_admissible"] == "true"
            ),
            key=lambda row: (row["finished_at"], row["attempt_id"]),
        )
        for row in source_rows:
            key = (row["task_id"], row["system"])
            if len(included[key]) < 8:
                included[key].append(row)

    metric_rows: list[dict[str, Any]] = []
    for task_id in tasks:
        for system in systems:
            selected = included[(task_id, system)]
            baseline = [row for row in selected if row["source_batch"] == "FROZEN_BASELINE"]
            current = [row for row in selected if row["source_batch"] == "CURRENT_ZCLOUD_RUNTIME"]
            baseline_scores = [float(row["score"]) for row in baseline if row["score"] != ""]
            current_scores = [float(row["score"]) for row in current if row["score"] != ""]
            complete = [row for row in selected if row["analysis_disposition"] == "COUNTABLE_COMPLETE_ARTIFACT"]
            artifact_backed = [row for row in complete if row["candidate_xlsx_status"] == "VALID_XLSX"]
            all_cell_rows = [
                row for row in rows
                if row["task_id"] == task_id and row["system"] == system
            ]
            recovered_abnormal = [
                row for row in all_cell_rows
                if row["analysis_disposition"] == "RECOVERED_ARTIFACT_FROM_ABNORMAL_TRIAL"
            ]
            delivery_failures = [
                row for row in all_cell_rows
                if row["analysis_disposition"] == "MODEL_DELIVERY_FAILURE"
            ]
            incomplete_agent_artifacts = [
                row
                for row in all_cell_rows
                if row["analysis_disposition"]
                == "INCOMPLETE_AGENT_ARTIFACT_FROM_ABNORMAL_TRIAL"
            ]
            native_receipt_inconsistent_artifacts = [
                row
                for row in all_cell_rows
                if row["model_artifact_outcome_valid"] == "true"
                and row["native_receipt_consistent"] == "false"
            ]
            controller_inconsistencies = [
                row
                for row in all_cell_rows
                if row["controller_countable"] == "true"
                and row["score_admissible"] != "true"
            ]
            excluded_legacy = [
                row for row in all_cell_rows
                if row.get("analysis_disposition", "").startswith("EXCLUDED_")
            ]
            baseline_successes = sum(float(row["score"]) >= PASS_THRESHOLD for row in baseline)
            current_successes = sum(float(row["score"]) >= PASS_THRESHOLD for row in current)
            metric_rows.append(
                {
                    "task_id": task_id,
                    "track": task_track(task_id),
                    "system": system,
                    "controller_coverage_n": controller_counts[(task_id, system)],
                    "baseline_n": len(baseline),
                    "current_contract_n": len(current),
                    "countable_complete_n": len(complete),
                    "artifact_backed_n": len(artifact_backed),
                    "controller_inconsistency_n": len(controller_inconsistencies),
                    "recovered_abnormal_artifact_n": len(recovered_abnormal),
                    "incomplete_agent_artifact_n": len(incomplete_agent_artifacts),
                    "native_receipt_inconsistent_artifact_n": len(
                        native_receipt_inconsistent_artifacts
                    ),
                    "model_delivery_failure_n": len(delivery_failures),
                    "excluded_unscoreable_receipt_n": len(excluded_legacy),
                    "baseline_successes_at_0_70": baseline_successes,
                    "baseline_empirical_success_rate": fmt(
                        baseline_successes / len(baseline) if baseline else None
                    ),
                    "baseline_mean_provisional": fmt(mean(baseline_scores)),
                    "baseline_median_provisional": fmt(median(baseline_scores)),
                    "baseline_stddev_provisional": fmt(pstdev(baseline_scores)),
                    "current_contract_successes_at_0_70": current_successes,
                    "current_contract_empirical_success_rate": fmt(
                        current_successes / len(current) if current else None
                    ),
                    "current_contract_mean_provisional": fmt(mean(current_scores)),
                    "current_contract_median_provisional": fmt(median(current_scores)),
                    "current_contract_stddev_provisional": fmt(pstdev(current_scores)),
                    "standard_pass_at_8": "",
                    "pass_at_8_status": "NOT_REPORTED_MIXED_CONTRACTS_AND_DEFINITION_UNFROZEN",
                    "judge_status": "PROVISIONAL_PRE_V2_FREEZE",
                    "windows_excel_status": "PENDING" if task_id == "P15-B-PUBLIC-PIVOT-001" else "NOT_REQUIRED_OR_UNDETERMINED",
                }
            )

    metric_fields = list(metric_rows[0])
    write_csv(output_dir / "N8_TASK_SYSTEM_METRICS.csv", metric_fields, metric_rows)

    task_rows: list[dict[str, Any]] = []
    for task_id in tasks:
        items = {row["system"]: row for row in metric_rows if row["task_id"] == task_id}
        baseline_task_balanced_values = [
            float(row["baseline_mean_provisional"])
            for row in items.values()
            if row["baseline_mean_provisional"] != ""
        ]
        current_task_balanced_values = [
            float(row["current_contract_mean_provisional"])
            for row in items.values()
            if row["current_contract_mean_provisional"] != ""
        ]
        task_rows.append(
            {
                "task_id": task_id,
                "track": task_track(task_id),
                "controller_coverage_n": sum(int(row["controller_coverage_n"]) for row in items.values()),
                "countable_complete_n": sum(int(row["countable_complete_n"]) for row in items.values()),
                "artifact_backed_n": sum(int(row["artifact_backed_n"]) for row in items.values()),
                "controller_inconsistency_n": sum(int(row["controller_inconsistency_n"]) for row in items.values()),
                "recovered_abnormal_artifact_n": sum(int(row["recovered_abnormal_artifact_n"]) for row in items.values()),
                "incomplete_agent_artifact_n": sum(int(row["incomplete_agent_artifact_n"]) for row in items.values()),
                "native_receipt_inconsistent_artifact_n": sum(
                    int(row["native_receipt_inconsistent_artifact_n"])
                    for row in items.values()
                ),
                "model_delivery_failure_n": sum(int(row["model_delivery_failure_n"]) for row in items.values()),
                "excluded_unscoreable_receipt_n": sum(int(row["excluded_unscoreable_receipt_n"]) for row in items.values()),
                "codex_n": items["codex_gpt56sol"]["controller_coverage_n"],
                "codex_baseline_n": items["codex_gpt56sol"]["baseline_n"],
                "codex_baseline_mean_provisional": items["codex_gpt56sol"]["baseline_mean_provisional"],
                "codex_current_contract_n": items["codex_gpt56sol"]["current_contract_n"],
                "codex_current_contract_mean_provisional": items["codex_gpt56sol"]["current_contract_mean_provisional"],
                "claude_n": items["claude_opus5"]["controller_coverage_n"],
                "claude_baseline_n": items["claude_opus5"]["baseline_n"],
                "claude_baseline_mean_provisional": items["claude_opus5"]["baseline_mean_provisional"],
                "claude_current_contract_n": items["claude_opus5"]["current_contract_n"],
                "claude_current_contract_mean_provisional": items["claude_opus5"]["current_contract_mean_provisional"],
                "qwen_n": items["qwen38max"]["controller_coverage_n"],
                "qwen_baseline_n": items["qwen38max"]["baseline_n"],
                "qwen_baseline_mean_provisional": items["qwen38max"]["baseline_mean_provisional"],
                "qwen_current_contract_n": items["qwen38max"]["current_contract_n"],
                "qwen_current_contract_mean_provisional": items["qwen38max"]["current_contract_mean_provisional"],
                "baseline_system_balanced_mean_provisional": fmt(mean(baseline_task_balanced_values)),
                "current_contract_system_balanced_mean_provisional": fmt(mean(current_task_balanced_values)),
                "numeric_difficulty_status": "NOT_INTERPRETABLE_BEFORE_JUDGE_V2_HUMAN_AND_CONTRACT_FREEZE",
                "windows_excel_status": "PENDING" if task_id == "P15-B-PUBLIC-PIVOT-001" else "NOT_REQUIRED_OR_UNDETERMINED",
            }
        )
    write_csv(output_dir / "N8_TASK_SUMMARY.csv", list(task_rows[0]), task_rows)

    system_rows: list[dict[str, Any]] = []
    for system, metadata in systems.items():
        items = [row for row in metric_rows if row["system"] == system]
        attempt_rows = [row for row in rows if row["system"] == system and row["source_batch"] == "CURRENT_ZCLOUD_RUNTIME"]
        selected_rows = [row for key, values in included.items() if key[1] == system for row in values]
        complete_rows = [row for row in selected_rows if row["analysis_disposition"] == "COUNTABLE_COMPLETE_ARTIFACT"]
        artifact_rows = [
            row for row in complete_rows if row["candidate_xlsx_status"] == "VALID_XLSX"
        ]
        recovered_rows = [
            row for row in attempt_rows
            if row["analysis_disposition"] == "RECOVERED_ARTIFACT_FROM_ABNORMAL_TRIAL"
        ]
        delivery_rows = [
            row for row in attempt_rows
            if row["analysis_disposition"] == "MODEL_DELIVERY_FAILURE"
        ]
        incomplete_agent_rows = [
            row
            for row in attempt_rows
            if row["analysis_disposition"]
            == "INCOMPLETE_AGENT_ARTIFACT_FROM_ABNORMAL_TRIAL"
        ]
        native_receipt_inconsistent_rows = [
            row
            for row in attempt_rows
            if row["model_artifact_outcome_valid"] == "true"
            and row["native_receipt_consistent"] == "false"
        ]
        baseline_task_means = [
            float(row["baseline_mean_provisional"])
            for row in items
            if row["baseline_mean_provisional"] != ""
        ]
        current_task_means = [
            float(row["current_contract_mean_provisional"])
            for row in items
            if row["current_contract_mean_provisional"] != ""
        ]
        system_rows.append(
            {
                "system": system,
                "agent": metadata.get("agent") or "",
                "model": metadata.get("model_name") or "",
                "controller_coverage_n": sum(int(row["controller_coverage_n"]) for row in items),
                "countable_complete_n": len(complete_rows),
                "artifact_backed_n": len(artifact_rows),
                "controller_inconsistency_n": sum(
                    int(row["controller_inconsistency_n"]) for row in items
                ),
                "recovered_abnormal_artifact_n": len(recovered_rows),
                "incomplete_agent_artifact_n": len(incomplete_agent_rows),
                "native_receipt_inconsistent_artifact_n": len(
                    native_receipt_inconsistent_rows
                ),
                "model_delivery_failure_n": len(delivery_rows),
                "excluded_unscoreable_receipt_n": sum(int(row["excluded_unscoreable_receipt_n"]) for row in items),
                "remaining_to_120_by_controller": 120 - sum(int(row["controller_coverage_n"]) for row in items),
                "tasks_at_controller_n8": sum(int(row["controller_coverage_n"]) == 8 for row in items),
                "baseline_task_balanced_mean_provisional": fmt(mean(baseline_task_means)),
                "current_contract_task_balanced_mean_provisional": fmt(mean(current_task_means)),
                "current_runtime_terminal_receipts": len(attempt_rows),
                "model_cost_units": "",
                "model_cost_status": "NOT_ATTRIBUTABLE_FROM_OVERLAPPING_RUN_DELTAS",
                "formal_ranking_status": "NOT_ALLOWED_BEFORE_JUDGE_V2_AND_CONTRACT_FREEZE",
            }
        )
    write_csv(output_dir / "N8_SYSTEM_SUMMARY.csv", list(system_rows[0]), system_rows)

    controller_total = sum(controller_counts.values())
    selected_rows = [row for values in included.values() for row in values]
    controller_inconsistencies = [
        row
        for row in rows
        if row["controller_countable"] == "true" and row["score_admissible"] != "true"
    ]
    recovered_abnormal = [row for row in rows if row["analysis_disposition"] == "RECOVERED_ARTIFACT_FROM_ABNORMAL_TRIAL"]
    incomplete_agent_artifacts = [
        row
        for row in rows
        if row["analysis_disposition"]
        == "INCOMPLETE_AGENT_ARTIFACT_FROM_ABNORMAL_TRIAL"
    ]
    native_receipt_inconsistent_artifacts = [
        row
        for row in rows
        if row["model_artifact_outcome_valid"] == "true"
        and row["native_receipt_consistent"] == "false"
    ]
    model_delivery_failures = [row for row in rows if row["analysis_disposition"] == "MODEL_DELIVERY_FAILURE"]
    excluded_infra = [row for row in rows if row["analysis_disposition"] == "EXCLUDED_INFRA_FAILURE_WITHOUT_ARTIFACT"]
    excluded_non_delivery = [row for row in rows if row["analysis_disposition"] == "EXCLUDED_NON_DELIVERY_WITHOUT_ARTIFACT"]
    complete = [row for row in selected_rows if row["analysis_disposition"] == "COUNTABLE_COMPLETE_ARTIFACT"]
    artifact_backed = [row for row in complete if row["candidate_xlsx_status"] == "VALID_XLSX"]

    usage_receipt = read_json(usage_receipt_path)
    if usage_receipt.get("settled") is not True or usage_receipt.get("phase") != "settlement":
        raise ValueError("--usage-receipt must be an immutable settled settlement receipt")
    usage_snapshot = (
        usage_receipt.get("snapshot")
        if isinstance(usage_receipt.get("snapshot"), dict)
        else {}
    )
    token_total_units = usage_snapshot.get("total_units")
    if (
        not isinstance(token_total_units, (int, float))
        or isinstance(token_total_units, bool)
        or not math.isfinite(float(token_total_units))
        or float(token_total_units) < 0
    ):
        raise ValueError("Dedicated-token usage receipt has no finite nonnegative total_units")
    usage_receipt_observed = parse_time(str(usage_receipt.get("observed_at") or ""))
    usage_snapshot_observed = parse_time(str(usage_snapshot.get("observed_at") or ""))
    if usage_receipt_observed is None or usage_snapshot_observed is None:
        raise ValueError("Usage receipt and snapshot both require timezone-aware observed_at")
    if usage_receipt_observed > cutoff or usage_snapshot_observed > cutoff:
        raise ValueError("Usage receipt is newer than the fixed evidence cutoff")

    snapshot = {
        "schema_version": EXPORT_SCHEMA_VERSION,
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "evidence_cutoff_finished_at": cutoff.isoformat(),
        "provenance": {
            "source_repository_head": source_head,
            "controller_path": CONTROLLER_RELATIVE_PATH.as_posix(),
            "controller_sha256": controller_sha256,
            "campaign_config_path": repository_relative(source_root, config_path),
            "campaign_config_sha256": sha256_file(config_path),
            "usage_receipt_path": repository_relative(source_root, usage_receipt_path),
            "usage_receipt_sha256": sha256_file(usage_receipt_path),
            "expected_runtime_receipts_at_cutoff": args.expected_runtime_receipts,
            "included_runtime_receipts_at_cutoff": len(runtime_records),
            "ignored_runtime_receipts_after_cutoff": ignored_runtime_receipts,
            "baseline_terminal_receipts": sum(
                row["source_batch"] == "FROZEN_BASELINE" for row in rows
            ),
        },
        "campaign_id": config.get("campaign_id"),
        "campaign_status": config.get("status"),
        "target": {"tasks": len(tasks), "systems": len(systems), "attempts_per_task_system": 8, "total": 360},
        "coverage": {
            "controller_effective": controller_total,
            "controller_remaining": 360 - controller_total,
            "countable_complete_scored": len(complete),
            "artifact_backed_scored": len(artifact_backed),
            "controller_inconsistencies": len(controller_inconsistencies),
            "recovered_artifacts_from_abnormal_trials": len(recovered_abnormal),
            "incomplete_agent_artifacts_from_abnormal_trials": len(
                incomplete_agent_artifacts
            ),
            "native_receipt_inconsistent_artifacts": len(
                native_receipt_inconsistent_artifacts
            ),
            "model_delivery_failures": len(model_delivery_failures),
            "excluded_infrastructure_trials_without_admissible_score": len(excluded_infra),
            "excluded_non_delivery_trials_without_admissible_score": len(excluded_non_delivery),
        },
        "runtime": {
            "terminal_receipts": len(runtime_records),
            "baseline_terminal_receipts": sum(
                row["source_batch"] == "FROZEN_BASELINE" for row in rows
            ),
            "attempt_rows_total": len(rows),
            "dedicated_token_total_used_units": round(float(token_total_units), 6),
            "usage_receipt_observed_at": usage_receipt_observed.isoformat(),
            "usage_snapshot_observed_at": usage_snapshot_observed.isoformat(),
            "recent_log_window_requests": usage_snapshot.get("request_count"),
            "per_model_cumulative_cost_available": False,
            "run_delta_fields_may_overlap": True,
            "per_run_deltas_summed_as_campaign_cost": False,
            "unit_to_usd_equivalence": "NOT_ASSERTED",
        },
        "score_contract": {
            "pass_threshold": PASS_THRESHOLD,
            "judge_status": "PROVISIONAL_PRE_V2_FREEZE",
            "formal_pooling_with_frozen_baseline": False,
            "standard_pass_at_8_reported": False,
            "reason": "The baseline and current runs use different execution contracts, the Cowork pass@8 wording remains unresolved, and Judge v2 is not frozen.",
        },
        "public_release_boundary": {
            "attempt_prompts_included": False,
            "trajectories_included": False,
            "model_workbooks_included": False,
            "credentials_or_endpoints_included": False,
            "local_paths_included": False,
            "new_private_truth_included": False,
        },
    }
    (output_dir / "N8_SNAPSHOT.json").write_text(
        json.dumps(snapshot, indent=2, sort_keys=True, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    print(json.dumps(snapshot, indent=2, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
