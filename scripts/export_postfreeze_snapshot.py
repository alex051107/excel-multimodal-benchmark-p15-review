#!/usr/bin/env python3
"""Publish a fail-closed, aggregate-only post-freeze P15 snapshot.

The private analyzer is the authority for joins and aggregation.  This command
verifies that authority's schema, provenance and file hashes, then copies only
five allowlisted aggregate tables into a new public directory.  Attempt-level
identifiers, workbook/trace locations and provider transport data are never
published.
"""

from __future__ import annotations

import argparse
import csv
import ctypes
import errno
import hashlib
import io
import json
import math
import os
import re
import shutil
import stat
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Sequence


PRIVATE_SCHEMA_VERSION = "p15_private_analysis_v1"
PUBLIC_SCHEMA_VERSION = "p15_public_postfreeze_v1"
FROZEN_EVALUATOR_COMMIT = "7628d01d23ac2d82db2b9947a7d1a3b1e7038bd4"
FROZEN_EVALUATOR_DATA_COMMIT = "419dbed0509cbe9eb026806778a88aed847c9916"
FROZEN_JUDGE_VERSION = "p15-judge-v2-evidence-locality-v2"
FROZEN_SCORE_SOURCE = "FROZEN_EVALUATOR_SCORES_CSV"
REQUIRED_PAUSED_TASKS = {
    "P15-B-OPS-CLEAN-JOIN-001",
    "P15-C-RECEIPTS-001",
}
EXPECTED_SYSTEMS = {"codex_gpt56sol", "claude_opus5", "qwen38max"}
EXPECTED_SYSTEM_ORDER = ("codex_gpt56sol", "claude_opus5", "qwen38max")
EXPECTED_COHORTS = {"current_contract", "historical_baseline", "diagnostic_only"}
EXPECTED_TASK_IDS = {
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
}

SYSTEM_FIELDS = [
    "cohort", "system", "ranking_eligible", "score_source", "artifact_count",
    "scored_count", "unscored_count", "score_coverage", "original_scored_count",
    "original_mean_score_scored_only", "frozen_mean_score_scored_only",
    "frozen_mean_score_lower_bound", "frozen_mean_score_upper_bound", "pass_count",
    "pass_rate_scored_only", "judge_error_count", "infra_error_count",
]
TRACK_FIELDS = ["cohort", "track", *SYSTEM_FIELDS[1:]]
TASK_SYSTEM_FIELDS = [
    "cohort", "task_id", "track", "system", "task_status", "ranking_eligible",
    "score_source", "artifact_count", "independent_complete_scored_count",
    "scored_count", "unscored_count", "score_coverage", "original_scored_count",
    "original_mean_score_scored_only", "frozen_mean_score_scored_only",
    "frozen_mean_score_lower_bound", "frozen_mean_score_upper_bound", "pass_count",
    "pass_rate_scored_only", "judge_error_count", "infra_error_count", "pass_at_8",
    "pass_at_8_status",
]
COST_TIME_FIELDS = [
    "cohort", "system", "ranking_eligible", "score_source", "attempt_count",
    "spent_units_reported_count", "spent_units_sum", "spent_units_mean",
    "gateway_spent_units_reported_count", "gateway_spent_units_sum",
    "gateway_spent_units_mean", "token_estimated_units_reported_count",
    "token_estimated_units_sum", "token_estimated_units_mean",
    "harbor_reported_cost_usd_reported_count", "harbor_reported_cost_usd_sum",
    "harbor_reported_cost_usd_mean", "provider_total_units_snapshot_reported_count",
    "provider_total_units_snapshot_latest", "provider_total_units_snapshot_latest_observed_at",
    "run_wall_seconds_reported_count", "run_wall_seconds_sum", "run_wall_seconds_mean",
    "harbor_wall_seconds_reported_count", "harbor_wall_seconds_sum",
    "harbor_wall_seconds_mean", "agent_execution_seconds_reported_count",
    "agent_execution_seconds_sum", "agent_execution_seconds_mean",
]
TRACE_FIELDS = [
    "cohort", "system", "ranking_eligible", "score_source", "attempt_count",
    "trace_present_count", "trace_missing_count", "step_count_reported_count",
    "total_steps", "mean_steps", "prompt_tokens_reported_count", "total_prompt_tokens",
    "cached_tokens_reported_count", "total_cached_tokens",
    "completion_tokens_reported_count", "total_completion_tokens", "total_tool_calls",
    "total_completed_tool_calls",
]

PUBLIC_TABLES = {
    "system_summary.csv": SYSTEM_FIELDS,
    "track_summary.csv": TRACK_FIELDS,
    "task_system_summary.csv": TASK_SYSTEM_FIELDS,
    "cost_time_summary.csv": COST_TIME_FIELDS,
    "trace_summary.csv": TRACE_FIELDS,
}
PRIVATE_HASHED_OUTPUTS = {
    "attempts.csv",
    "criteria_long.csv",
    *PUBLIC_TABLES,
}
REQUIRED_PRIVATE_INPUT_HASHES = {
    "evidence_index.csv",
    "historical_reuse_audit.csv",
    "scores.csv",
    "summary.json",
}

EXPECTED_AGGREGATION_CONTRACT = {
    "group_by_cohort": True,
    "cross_cohort_pooling": False,
    "paused_tasks_excluded_from_rankings": True,
    "diagnostic_only_excluded_from_rankings": True,
    "historical_baseline_descriptive_only": True,
    "ranking_cohort": "current_contract",
    "aggregate_systems": list(EXPECTED_SYSTEM_ORDER),
    "non_target_system_attempts_retained_only_in_private_attempt_tables": True,
}
EXPECTED_PASS_AT_8_CONTRACT = {
    "k": 8,
    "grouping": ["cohort", "task_id", "system"],
    "minimum_independent_complete_scored_attempts": 8,
    "formula": "1-C(n-c,8)/C(n,8)",
    "missing_when_n_lt_8": True,
    "paused_tasks_missing": True,
    "historical_baseline_missing": True,
}
EXPECTED_MISSING_VALUE_CONTRACT = {
    "csv_null_encoding": "empty cell",
    "missing_is_zero": False,
    "judge_error_score_and_pass_are_null": True,
    "infra_error_score_and_pass_are_null": True,
}
EXPECTED_COST_CONTRACT = {
    "spent_units": "LOCAL_LEDGER_UNITS_NOT_USD",
    "gateway_spent_units": "GATEWAY_LEDGER_UNITS_NOT_USD",
    "token_estimated_units": "LOCAL_TOKEN_PRICING_ESTIMATE_NOT_PROVIDER_CHARGE",
    "harbor_reported_cost_usd": "HARBOR_FRAMEWORK_REPORTED_USD",
    "provider_total_units_snapshot": "PROVIDER_CUMULATIVE_ACCOUNT_SNAPSHOT_NOT_PER_RUN",
    "cross_basis_total_prohibited": True,
    "provider_cumulative_snapshots_are_not_summed": True,
}
PRIVATE_INPUT_FILES = {"provenance.json", *PRIVATE_HASHED_OUTPUTS}

ATTEMPT_FIELDS = [
    "attempt_id", "run_id", "wave_id", "artifact", "artifact_sha256", "trial_name",
    "task_id", "track", "system", "model", "agent", "cohort", "campaign_contract_id",
    "run_classification", "evidence_match_status", "evidence_completeness",
    "effective_experiment_valid", "protocol_valid", "experiment_valid",
    "receipt_valid_attempt", "paused_task", "ranking_eligible",
    "ranking_exclusion_reason", "formal_comparison_eligible", "scoreable_independent",
    "score_source", "original_evaluation_status", "original_normalized_score",
    "receipt_normalized_score", "frozen_evaluation_status", "frozen_normalized_score",
    "frozen_pass", "score_delta_from_original", "hurdle_failures", "failure_codes",
    "criterion_scores_json", "spent_units", "gateway_spent_units",
    "token_estimated_units", "harbor_reported_cost_usd", "provider_total_units_snapshot",
    "provider_snapshot_observed_at", "run_wall_seconds", "harbor_wall_seconds",
    "agent_execution_seconds", "trace_status", "atif_schema_version", "atif_step_count",
    "trace_total_steps", "trace_prompt_tokens", "trace_cached_tokens",
    "trace_completion_tokens", "trace_tool_call_count", "trace_completed_tool_call_count",
]

COUNT_FIELDS = {
    "artifact_count", "scored_count", "unscored_count", "original_scored_count",
    "pass_count", "judge_error_count", "infra_error_count",
    "independent_complete_scored_count", "attempt_count", "spent_units_reported_count",
    "gateway_spent_units_reported_count", "token_estimated_units_reported_count",
    "harbor_reported_cost_usd_reported_count",
    "provider_total_units_snapshot_reported_count", "run_wall_seconds_reported_count",
    "harbor_wall_seconds_reported_count", "agent_execution_seconds_reported_count",
    "trace_present_count", "trace_missing_count", "step_count_reported_count",
    "total_steps", "prompt_tokens_reported_count", "total_prompt_tokens",
    "cached_tokens_reported_count", "total_cached_tokens",
    "completion_tokens_reported_count", "total_completion_tokens", "total_tool_calls",
    "total_completed_tool_calls",
}
RATIO_FIELDS = {
    "score_coverage", "original_mean_score_scored_only", "frozen_mean_score_scored_only",
    "frozen_mean_score_lower_bound", "frozen_mean_score_upper_bound",
    "pass_rate_scored_only", "pass_at_8",
}
NONNEGATIVE_FIELDS = {
    "spent_units_sum", "spent_units_mean", "gateway_spent_units_sum",
    "gateway_spent_units_mean", "token_estimated_units_sum",
    "token_estimated_units_mean", "harbor_reported_cost_usd_sum",
    "harbor_reported_cost_usd_mean", "provider_total_units_snapshot_latest",
    "run_wall_seconds_sum", "run_wall_seconds_mean", "harbor_wall_seconds_sum",
    "harbor_wall_seconds_mean", "agent_execution_seconds_sum",
    "agent_execution_seconds_mean", "mean_steps",
}
NULL_TOKENS = {"null", "none"}
SAFE_TOKEN = re.compile(r"[A-Za-z0-9][A-Za-z0-9_.:+-]{0,159}\Z")
TASK_ID = re.compile(r"P15-[A-Z0-9]+(?:-[A-Z0-9]+)+-\d{3}\Z")
HEX40 = re.compile(r"[0-9a-f]{40}\Z")
HEX64 = re.compile(r"[0-9a-f]{64}\Z")
SENSITIVE_PATTERNS = [
    ("absolute POSIX path", re.compile(r"(?<![A-Za-z0-9])/(?:Users|home|private|tmp|var|Volumes|opt|etc)/")),
    ("absolute Windows path", re.compile(r"(?<![A-Za-z0-9])[A-Za-z]:\\\\")),
    ("network endpoint", re.compile(r"(?:https?|wss?)://", re.IGNORECASE)),
    ("authorization value", re.compile(r"\bBearer\s+[A-Za-z0-9._~+/-]{8,}", re.IGNORECASE)),
    ("credential-like value", re.compile(r"\b(?:sk|api)[-_][A-Za-z0-9_-]{12,}\b", re.IGNORECASE)),
    ("provider request id", re.compile(r"\b(?:req|request)[-_][A-Za-z0-9_-]{8,}\b", re.IGNORECASE)),
    ("UUID", re.compile(r"\b[0-9a-f]{8}-[0-9a-f]{4}-[1-5][0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}\b", re.IGNORECASE)),
    ("runtime wave id", re.compile(r"\b20\d{6}T\d{6}Z-P15-")),
]


@dataclass(frozen=True)
class FileFingerprint:
    sha256: str
    device: int
    inode: int
    mode: int
    size_bytes: int
    mtime_ns: int
    ctime_ns: int


def absolute_without_resolving(path: Path) -> Path:
    return Path(os.path.abspath(os.fspath(path.expanduser())))


def reject_symlink_components(path: Path, *, label: str) -> None:
    absolute = absolute_without_resolving(path)
    current = Path(absolute.anchor)
    for component in absolute.parts[1:]:
        current /= component
        try:
            current_state = current.lstat()
        except FileNotFoundError:
            return
        except OSError as exc:
            raise ValueError(f"Cannot inspect {label} path component") from exc
        if stat.S_ISLNK(current_state.st_mode):
            raise ValueError(f"{label} path contains a symlink component")


def fingerprint_from_stat(digest: str, value: os.stat_result) -> FileFingerprint:
    return FileFingerprint(
        sha256=digest,
        device=value.st_dev,
        inode=value.st_ino,
        mode=value.st_mode,
        size_bytes=value.st_size,
        mtime_ns=value.st_mtime_ns,
        ctime_ns=value.st_ctime_ns,
    )


def read_stable_bytes(path: Path, *, label: str) -> tuple[bytes, FileFingerprint]:
    reject_symlink_components(path, label=label)
    flags = os.O_RDONLY
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW
    try:
        descriptor = os.open(path, flags)
    except OSError as exc:
        raise ValueError(f"Cannot open {label}") from exc
    try:
        before = os.fstat(descriptor)
        if not stat.S_ISREG(before.st_mode):
            raise ValueError(f"{label} must be a regular file")
        chunks: list[bytes] = []
        digest = hashlib.sha256()
        while True:
            chunk = os.read(descriptor, 1024 * 1024)
            if not chunk:
                break
            chunks.append(chunk)
            digest.update(chunk)
        after = os.fstat(descriptor)
        before_fingerprint = fingerprint_from_stat(digest.hexdigest(), before)
        after_fingerprint = fingerprint_from_stat(digest.hexdigest(), after)
        if before_fingerprint != after_fingerprint:
            raise ValueError(f"{label} changed while being read")
    finally:
        os.close(descriptor)
    try:
        current = path.lstat()
    except OSError as exc:
        raise ValueError(f"Cannot re-inspect {label}") from exc
    if stat.S_ISLNK(current.st_mode) or fingerprint_from_stat(
        digest.hexdigest(), current
    ) != before_fingerprint:
        raise ValueError(f"{label} changed while being read")
    return b"".join(chunks), before_fingerprint


def capture_private_file_snapshot(private_dir: Path) -> dict[str, FileFingerprint]:
    return {
        filename: read_stable_bytes(
            private_dir / filename, label=f"private analyzer {filename}"
        )[1]
        for filename in sorted(PRIVATE_INPUT_FILES)
    }


def require_private_file_snapshot(
    private_dir: Path,
    expected: dict[str, FileFingerprint],
    *,
    phase: str,
) -> None:
    current = capture_private_file_snapshot(private_dir)
    if current != expected:
        changed = sorted(
            filename
            for filename in set(current) | set(expected)
            if current.get(filename) != expected.get(filename)
        )
        raise ValueError(
            f"Private analyzer input changed during export ({phase}): "
            + ", ".join(changed[:5])
        )


def read_json_object(path: Path) -> dict[str, Any]:
    try:
        payload, _ = read_stable_bytes(path, label=path.name)
        value = json.loads(payload.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"Cannot read JSON object: {path.name}") from exc
    if not isinstance(value, dict):
        raise ValueError(f"JSON object required: {path.name}")
    return value


def sha256_file(path: Path) -> str:
    return read_stable_bytes(path, label=path.name)[1].sha256


def aware_time(value: Any, *, label: str) -> datetime:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{label} must be a non-empty ISO-8601 timestamp")
    try:
        parsed = datetime.fromisoformat(value.strip().replace("Z", "+00:00"))
    except ValueError as exc:
        raise ValueError(f"{label} must be an ISO-8601 timestamp") from exc
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        raise ValueError(f"{label} must include a timezone")
    return parsed.astimezone(timezone.utc)


def full_sha(value: Any, *, label: str) -> str:
    normalized = str(value or "").strip().lower()
    if not HEX40.fullmatch(normalized):
        raise ValueError(f"{label} must be a full 40-character Git SHA")
    return normalized


def validate_hash_map(value: Any, *, label: str, required: set[str]) -> dict[str, str]:
    if not isinstance(value, dict) or not value:
        raise ValueError(f"{label} must be a non-empty object")
    if not required.issubset(value):
        missing = ", ".join(sorted(required - set(value)))
        raise ValueError(f"{label} is missing required entries: {missing}")
    result: dict[str, str] = {}
    for name, digest in value.items():
        if not isinstance(name, str) or Path(name).name != name or name in {".", ".."}:
            raise ValueError(f"{label} keys must be plain filenames")
        normalized = str(digest or "").strip().lower()
        if not HEX64.fullmatch(normalized):
            raise ValueError(f"{label}.{name} must be a 64-character SHA-256")
        result[name] = normalized
    return result


def check_safe_cell(value: str, *, filename: str, row_number: int, field: str) -> None:
    if "\n" in value or "\r" in value or "\x00" in value:
        raise ValueError(f"{filename} row {row_number} field {field} contains control text")
    for label, pattern in SENSITIVE_PATTERNS:
        if pattern.search(value):
            raise ValueError(f"{filename} row {row_number} field {field} contains {label}")


def read_csv_exact(path: Path, fields: list[str]) -> list[dict[str, str]]:
    try:
        payload, _ = read_stable_bytes(path, label=f"private analyzer {path.name}")
        text = payload.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError(f"Cannot decode private analyzer output: {path.name}") from exc
    with io.StringIO(text, newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames != fields:
            raise ValueError(f"{path.name} header does not match {PRIVATE_SCHEMA_VERSION}")
        rows: list[dict[str, str]] = []
        for row_number, raw in enumerate(reader, start=2):
            if None in raw:
                raise ValueError(f"{path.name} row {row_number} has extra cells")
            row: dict[str, str] = {}
            for field in fields:
                value = str(raw.get(field) or "").strip()
                if field in RATIO_FIELDS | NONNEGATIVE_FIELDS and value.lower() in NULL_TOKENS:
                    value = ""
                check_safe_cell(value, filename=path.name, row_number=row_number, field=field)
                row[field] = value
            rows.append(row)
    if not rows:
        raise ValueError(f"{path.name} must contain at least one aggregate row")
    return rows


def as_int(value: str, *, label: str) -> int:
    if not re.fullmatch(r"\d+", value):
        raise ValueError(f"{label} must be a nonnegative integer")
    return int(value)


def as_number(value: str, *, label: str, nullable: bool = False) -> float | None:
    if value == "":
        if nullable:
            return None
        raise ValueError(f"{label} must not be empty")
    try:
        parsed = float(value)
    except ValueError as exc:
        raise ValueError(f"{label} must be numeric") from exc
    if not math.isfinite(parsed) or parsed < 0:
        raise ValueError(f"{label} must be finite and nonnegative")
    return parsed


def as_ratio(value: str, *, label: str, nullable: bool = True) -> float | None:
    parsed = as_number(value, label=label, nullable=nullable)
    if parsed is not None and parsed > 1:
        raise ValueError(f"{label} must be between 0 and 1")
    return parsed


def as_bool(value: str, *, label: str) -> bool:
    normalized = value.lower()
    if normalized not in {"true", "false"}:
        raise ValueError(f"{label} must be true or false")
    return normalized == "true"


def safe_token(value: str, *, label: str) -> str:
    if not SAFE_TOKEN.fullmatch(value):
        raise ValueError(f"{label} must be a non-empty machine-readable token")
    return value


def validate_common_rows(
    filename: str,
    rows: list[dict[str, str]],
    *,
    key_fields: tuple[str, ...],
) -> None:
    seen: set[tuple[str, ...]] = set()
    for index, row in enumerate(rows, start=2):
        label = f"{filename} row {index}"
        safe_token(row["cohort"], label=f"{label} cohort")
        safe_token(row["system"], label=f"{label} system")
        if row["cohort"] not in EXPECTED_COHORTS:
            raise ValueError(f"{label} is not one of the three declared cohorts")
        if row["system"] not in EXPECTED_SYSTEMS:
            raise ValueError(f"{label} is not one of the three frozen agent systems")
        if row["score_source"] != FROZEN_SCORE_SOURCE:
            raise ValueError(f"{label} has an unexpected score_source")
        eligible = as_bool(row["ranking_eligible"], label=f"{label} ranking_eligible")
        row["ranking_eligible"] = str(eligible).lower()
        key = tuple(row[field] for field in key_fields)
        if key in seen:
            raise ValueError(f"{filename} contains duplicate aggregate keys")
        seen.add(key)


def validate_score_rows(filename: str, rows: list[dict[str, str]]) -> None:
    for index, row in enumerate(rows, start=2):
        label = f"{filename} row {index}"
        for field in COUNT_FIELDS & set(row):
            as_int(row[field], label=f"{label} {field}")
        for field in NONNEGATIVE_FIELDS & set(row):
            as_number(row[field], label=f"{label} {field}", nullable=True)
        for field in RATIO_FIELDS & set(row):
            as_ratio(row[field], label=f"{label} {field}", nullable=True)

        artifacts = int(row["artifact_count"])
        scored = int(row["scored_count"])
        unscored = int(row["unscored_count"])
        original_scored = int(row["original_scored_count"])
        passes = int(row["pass_count"])
        if scored + unscored != artifacts:
            raise ValueError(f"{label} scored_count + unscored_count must equal artifact_count")
        if original_scored > artifacts or passes > scored:
            raise ValueError(f"{label} has a count larger than its eligible denominator")
        if int(row["judge_error_count"]) > unscored or int(row["infra_error_count"]) > unscored:
            raise ValueError(f"{label} error counts must not exceed unscored_count")

        coverage = as_ratio(row["score_coverage"], label=f"{label} score_coverage")
        if artifacts and (coverage is None or not math.isclose(coverage, scored / artifacts, abs_tol=1e-6)):
            raise ValueError(f"{label} score_coverage does not match scored/artifact counts")
        if artifacts == 0 and coverage not in {None, 0.0}:
            raise ValueError(f"{label} score_coverage must be empty or zero with no artifacts")

        original_mean = as_ratio(
            row["original_mean_score_scored_only"],
            label=f"{label} original_mean_score_scored_only",
        )
        if (original_scored == 0) != (original_mean is None):
            raise ValueError(f"{label} original mean must be empty exactly when original_scored_count is zero")

        frozen_fields = [
            "frozen_mean_score_scored_only",
            "frozen_mean_score_lower_bound",
            "frozen_mean_score_upper_bound",
            "pass_rate_scored_only",
        ]
        frozen_values = [as_ratio(row[field], label=f"{label} {field}") for field in frozen_fields]
        if scored == 0 and any(value is not None for value in frozen_values):
            raise ValueError(f"{label} frozen score summaries must stay empty when scored_count is zero")
        if scored > 0 and any(value is None for value in frozen_values):
            raise ValueError(f"{label} frozen score summaries are required when scored_count is nonzero")
        if scored > 0:
            mean_value, lower, upper, pass_rate = frozen_values
            assert mean_value is not None and lower is not None and upper is not None and pass_rate is not None
            if not lower <= mean_value <= upper:
                raise ValueError(f"{label} frozen score bounds do not contain the mean")
            if not math.isclose(pass_rate, passes / scored, abs_tol=1e-6):
                raise ValueError(f"{label} pass_rate_scored_only does not match pass/scored counts")


def validate_aggregate_tables(
    tables: dict[str, list[dict[str, str]]],
    paused_task_ids: set[str],
    sparse_task_system_inventory: set[tuple[str, str, str]],
) -> None:
    validate_common_rows("system_summary.csv", tables["system_summary.csv"], key_fields=("cohort", "system"))
    validate_common_rows("track_summary.csv", tables["track_summary.csv"], key_fields=("cohort", "track", "system"))
    validate_common_rows(
        "task_system_summary.csv",
        tables["task_system_summary.csv"],
        key_fields=("cohort", "task_id", "system"),
    )
    validate_common_rows("cost_time_summary.csv", tables["cost_time_summary.csv"], key_fields=("cohort", "system"))
    validate_common_rows("trace_summary.csv", tables["trace_summary.csv"], key_fields=("cohort", "system"))

    for filename in ("system_summary.csv", "track_summary.csv", "task_system_summary.csv"):
        validate_score_rows(filename, tables[filename])

    for filename in ("system_summary.csv", "track_summary.csv", "cost_time_summary.csv", "trace_summary.csv"):
        for index, row in enumerate(tables[filename], start=2):
            expected_eligible = row["cohort"] == "current_contract"
            if (row["ranking_eligible"] == "true") != expected_eligible:
                raise ValueError(
                    f"{filename} row {index} has ranking eligibility inconsistent with its cohort"
                )

    observed_paused: set[str] = set()
    for index, row in enumerate(tables["task_system_summary.csv"], start=2):
        label = f"task_system_summary.csv row {index}"
        if not TASK_ID.fullmatch(row["task_id"]):
            raise ValueError(f"{label} has an invalid task_id")
        if row["track"] not in {"A", "B", "C"} or row["task_id"].split("-")[1] != row["track"]:
            raise ValueError(f"{label} track does not match task_id")
        if row["task_id"] not in EXPECTED_TASK_IDS:
            raise ValueError(f"{label} task_id is outside the declared task inventory")
        paused = row["task_id"] in paused_task_ids
        diagnostic = row["cohort"] == "diagnostic_only"
        historical = row["cohort"] == "historical_baseline"
        expected_status = (
            "PAUSED_TASK_CONTRACT"
            if paused
            else "DIAGNOSTIC_ONLY"
            if diagnostic
            else "HISTORICAL_BASELINE_DESCRIPTIVE"
            if historical
            else "ACTIVE"
        )
        expected_eligible = not paused and row["cohort"] == "current_contract"
        if row["task_status"] != expected_status or (
            (row["ranking_eligible"] == "true") != expected_eligible
        ):
            raise ValueError(f"{label} does not preserve the task pause/ranking marker")
        if paused:
            observed_paused.add(row["task_id"])
            if row["pass_at_8"] or row["pass_at_8_status"] != "PAUSED_TASK":
                raise ValueError(f"{label} must exclude a paused task from pass@8")
        elif diagnostic:
            if row["pass_at_8"] or row["pass_at_8_status"] != "DIAGNOSTIC_ONLY":
                raise ValueError(f"{label} must exclude diagnostic-only rows from pass@8")
        elif historical:
            if (
                row["pass_at_8"]
                or row["pass_at_8_status"] != "HISTORICAL_DESCRIPTIVE_ONLY"
            ):
                raise ValueError(
                    f"{label} must exclude the sparse historical cohort from pass@8"
                )
        else:
            if row["pass_at_8_status"] not in {"N_LT_8", "COMPUTED"}:
                raise ValueError(f"{label} has an invalid pass_at_8_status")
            if (row["pass_at_8_status"] == "COMPUTED") != bool(row["pass_at_8"]):
                raise ValueError(f"{label} pass_at_8 value/status mismatch")
        completed = int(row["independent_complete_scored_count"])
        if completed > int(row["scored_count"]):
            raise ValueError(f"{label} independent complete count exceeds scored count")
        if row["pass_at_8_status"] == "COMPUTED" and completed < 8:
            raise ValueError(f"{label} computes pass@8 with fewer than eight independent complete scores")
    if observed_paused != paused_task_ids:
        raise ValueError("task_system_summary.csv does not preserve every provenance paused task")

    declared_system_keys = {
        (cohort, system) for cohort in EXPECTED_COHORTS for system in EXPECTED_SYSTEMS
    }
    system_keys = {
        (row["cohort"], row["system"]) for row in tables["system_summary.csv"]
    }
    cost_keys = {
        (row["cohort"], row["system"]) for row in tables["cost_time_summary.csv"]
    }
    trace_keys = {
        (row["cohort"], row["system"]) for row in tables["trace_summary.csv"]
    }
    if system_keys != declared_system_keys or cost_keys != system_keys or trace_keys != system_keys:
        raise ValueError(
            "system_summary, cost_time_summary and trace_summary must share the full declared cohort/system keys"
        )
    track_keys = {
        (row["cohort"], row["system"], row["track"])
        for row in tables["track_summary.csv"]
    }
    declared_track_keys = {
        (cohort, system, track)
        for cohort, system in declared_system_keys
        for track in ("A", "B", "C")
    }
    if track_keys != declared_track_keys:
        raise ValueError("track_summary must contain A, B and C for every declared cohort/system")

    task_system_keys = {
        (row["cohort"], row["task_id"], row["system"])
        for row in tables["task_system_summary.csv"]
    }
    declared_task_system_keys = {
        (cohort, task_id, system)
        for cohort in EXPECTED_COHORTS
        for task_id in EXPECTED_TASK_IDS
        for system in EXPECTED_SYSTEMS
    }
    if task_system_keys & sparse_task_system_inventory:
        raise ValueError("A task/system/cohort key cannot be both published and declared sparse")
    if task_system_keys | sparse_task_system_inventory != declared_task_system_keys:
        raise ValueError(
            "task_system_summary has a silent missing or extra task/system/cohort key"
        )

    for filename in ("cost_time_summary.csv", "trace_summary.csv"):
        for index, row in enumerate(tables[filename], start=2):
            label = f"{filename} row {index}"
            for field in COUNT_FIELDS & set(row):
                as_int(row[field], label=f"{label} {field}")
            for field in NONNEGATIVE_FIELDS & set(row):
                as_number(row[field], label=f"{label} {field}", nullable=True)
    for index, row in enumerate(tables["cost_time_summary.csv"], start=2):
        label = f"cost_time_summary.csv row {index}"
        for base in (
            "spent_units",
            "gateway_spent_units",
            "token_estimated_units",
            "harbor_reported_cost_usd",
            "run_wall_seconds",
            "harbor_wall_seconds",
            "agent_execution_seconds",
        ):
            count = int(row[f"{base}_reported_count"])
            total = as_number(row[f"{base}_sum"], label=f"{label} {base}_sum", nullable=True)
            mean = as_number(row[f"{base}_mean"], label=f"{label} {base}_mean", nullable=True)
            if count == 0 and (total is not None or mean is not None):
                raise ValueError(f"{label} {base} sum/mean must stay empty when nothing was reported")
            if count > 0 and (total is None or mean is None):
                raise ValueError(f"{label} {base} sum/mean are required when values were reported")
            if count > 0 and not math.isclose(mean, total / count, abs_tol=1e-6):
                raise ValueError(f"{label} {base}_mean does not match sum/reported_count")
        provider_count = int(row["provider_total_units_snapshot_reported_count"])
        provider_latest = as_number(
            row["provider_total_units_snapshot_latest"],
            label=f"{label} provider_total_units_snapshot_latest",
            nullable=True,
        )
        observed = row["provider_total_units_snapshot_latest_observed_at"]
        if provider_count == 0 and (provider_latest is not None or observed):
            raise ValueError(f"{label} provider snapshot must stay empty when nothing was reported")
        if provider_count > 0 and (provider_latest is None or not observed):
            raise ValueError(f"{label} provider snapshot value/time are required when reported")
        if observed:
            aware_time(observed, label=f"{label} provider timestamp")

    for index, row in enumerate(tables["trace_summary.csv"], start=2):
        label = f"trace_summary.csv row {index}"
        attempts = int(row["attempt_count"])
        present = int(row["trace_present_count"])
        missing = int(row["trace_missing_count"])
        if present + missing != attempts:
            raise ValueError(f"{label} trace present/missing counts do not match attempt_count")
        reported_steps = int(row["step_count_reported_count"])
        total_steps = int(row["total_steps"])
        mean_steps = as_number(row["mean_steps"], label=f"{label} mean_steps", nullable=True)
        if reported_steps > present:
            raise ValueError(f"{label} step_count_reported_count exceeds trace_present_count")
        if reported_steps == 0 and (total_steps != 0 or mean_steps is not None):
            raise ValueError(f"{label} step totals must stay zero/empty when nothing was reported")
        if reported_steps > 0 and (
            mean_steps is None or not math.isclose(mean_steps, total_steps / reported_steps, abs_tol=1e-6)
        ):
            raise ValueError(f"{label} mean_steps does not match total/reported_count")
        if int(row["total_completed_tool_calls"]) > int(row["total_tool_calls"]):
            raise ValueError(f"{label} completed tool calls exceed total tool calls")


def read_and_validate_attempts(path: Path, expected_count: int) -> None:
    try:
        payload, _ = read_stable_bytes(path, label="private analyzer attempts.csv")
        text = payload.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("Cannot decode private analyzer output: attempts.csv") from exc
    with io.StringIO(text, newline="") as handle:
        reader = csv.DictReader(handle)
        if reader.fieldnames != ATTEMPT_FIELDS:
            raise ValueError(f"attempts.csv header does not match {PRIVATE_SCHEMA_VERSION}")
        count = 0
        for row_number, row in enumerate(reader, start=2):
            if None in row:
                raise ValueError(f"attempts.csv row {row_number} has extra cells")
            count += 1
            status = str(row.get("frozen_evaluation_status") or "").strip().upper()
            score = str(row.get("frozen_normalized_score") or "").strip()
            passed = str(row.get("frozen_pass") or "").strip()
            if score.lower() in NULL_TOKENS:
                score = ""
            if passed.lower() in NULL_TOKENS:
                passed = ""
            if status in {"JUDGE_ERROR", "INFRA_ERROR"} and (score or passed):
                raise ValueError(
                    f"attempts.csv row {row_number} coerces {status} into a frozen score/pass"
                )
            if not score and passed:
                raise ValueError(f"attempts.csv row {row_number} has a pass value without a frozen score")
            if score:
                as_ratio(score, label=f"attempts.csv row {row_number} frozen_normalized_score", nullable=False)
                as_bool(passed, label=f"attempts.csv row {row_number} frozen_pass")
        if count != expected_count:
            raise ValueError(f"attempts.csv row count mismatch: expected {expected_count}, found {count}")


def validate_provenance(
    private_dir: Path,
    *,
    source_controller_head: str,
    expected_attempt_records: int,
    evidence_cutoff: datetime,
) -> tuple[dict[str, Any], set[str], set[tuple[str, str, str]]]:
    provenance_path = private_dir / "provenance.json"
    provenance = read_json_object(provenance_path)
    if provenance.get("schema_version") != PRIVATE_SCHEMA_VERSION:
        raise ValueError(f"provenance schema_version must be {PRIVATE_SCHEMA_VERSION}")
    aware_time(provenance.get("generated_at_utc"), label="provenance.generated_at_utc")
    attempt_count = provenance.get("attempt_count")
    if (
        not isinstance(attempt_count, int)
        or isinstance(attempt_count, bool)
        or attempt_count < 0
        or attempt_count != expected_attempt_records
    ):
        raise ValueError("provenance.attempt_count does not match --expected-attempt-records")
    counts: dict[str, int] = {"attempt_count": attempt_count}
    for field in ("evidence_row_count", "score_row_count", "joined_score_count", "evidence_only_count", "score_only_count"):
        value = provenance.get(field)
        if not isinstance(value, int) or isinstance(value, bool) or value < 0:
            raise ValueError(f"provenance.{field} must be a nonnegative integer")
        counts[field] = value
    if counts["joined_score_count"] + counts["evidence_only_count"] != counts["evidence_row_count"]:
        raise ValueError("provenance evidence join counts do not close")
    if counts["joined_score_count"] + counts["score_only_count"] != counts["score_row_count"]:
        raise ValueError("provenance score join counts do not close")
    if (
        counts["joined_score_count"]
        + counts["evidence_only_count"]
        + counts["score_only_count"]
        != counts["attempt_count"]
    ):
        raise ValueError("provenance attempt_count does not equal the joined union")
    if provenance.get("score_source") != FROZEN_SCORE_SOURCE:
        raise ValueError("provenance.score_source does not name the frozen evaluator scores")
    if provenance.get("evaluator_commit") != FROZEN_EVALUATOR_COMMIT:
        raise ValueError("provenance.evaluator_commit is not the frozen evaluator commit")
    if provenance.get("evaluator_data_commit") != FROZEN_EVALUATOR_DATA_COMMIT:
        raise ValueError("provenance.evaluator_data_commit is not the frozen evaluator data commit")
    if provenance.get("judge_version") != FROZEN_JUDGE_VERSION:
        raise ValueError("provenance.judge_version is not the frozen Judge version")
    if full_sha(
        provenance.get("source_controller_head"), label="provenance.source_controller_head"
    ) != source_controller_head:
        raise ValueError("provenance.source_controller_head does not match the CLI value")
    provenance_cutoff = aware_time(
        provenance.get("evidence_cutoff"), label="provenance.evidence_cutoff"
    )
    if provenance_cutoff != evidence_cutoff:
        raise ValueError("provenance.evidence_cutoff does not match the CLI UTC instant")

    declarations = {
        "systems": EXPECTED_SYSTEMS,
        "task_ids": EXPECTED_TASK_IDS,
        "cohorts": EXPECTED_COHORTS,
    }
    for field, expected in declarations.items():
        raw = provenance.get(field)
        if (
            not isinstance(raw, list)
            or any(not isinstance(item, str) for item in raw)
            or len(raw) != len(set(raw))
            or set(raw) != expected
        ):
            raise ValueError(f"provenance.{field} must declare the exact frozen inventory")

    input_hashes = validate_hash_map(
        provenance.get("input_hashes"),
        label="provenance.input_hashes",
        required=REQUIRED_PRIVATE_INPUT_HASHES,
    )
    if set(input_hashes) - (REQUIRED_PRIVATE_INPUT_HASHES | {"evidence_index.json"}):
        raise ValueError("provenance.input_hashes contains an unexpected input name")
    output_hashes = validate_hash_map(
        provenance.get("output_hashes"),
        label="provenance.output_hashes",
        required=PRIVATE_HASHED_OUTPUTS,
    )
    if set(output_hashes) != PRIVATE_HASHED_OUTPUTS:
        raise ValueError("provenance.output_hashes must contain exactly the private analyzer outputs")
    for filename, expected_hash in output_hashes.items():
        path = private_dir / filename
        if not path.is_file() or sha256_file(path) != expected_hash:
            raise ValueError(f"Private analyzer output hash mismatch: {filename}")

    score_sources = provenance.get("score_sources")
    expected_sources = {
        "original": "ORIGINAL_HARBOR_JUDGE",
        "receipt": "GUARD_FINAL_RECEIPT",
        "frozen": FROZEN_SCORE_SOURCE,
    }
    if score_sources != expected_sources:
        raise ValueError("provenance.score_sources does not match the frozen source contract")
    paused_raw = provenance.get("paused_tasks")
    if not isinstance(paused_raw, list):
        raise ValueError("provenance.paused_tasks must be a list")
    paused_ids: set[str] = set()
    for item in paused_raw:
        if not isinstance(item, dict):
            raise ValueError("Each provenance.paused_tasks entry must be an object")
        task_id = str(item.get("task_id") or "")
        if not TASK_ID.fullmatch(task_id) or task_id in paused_ids:
            raise ValueError("provenance.paused_tasks contains an invalid or duplicate task_id")
        if item.get("reason") != "TASK_CONTRACT_PAUSED" or item.get("ranking_eligible") is not False:
            raise ValueError("Every paused task must carry the pause reason and ranking exclusion")
        paused_ids.add(task_id)
    if paused_ids != REQUIRED_PAUSED_TASKS:
        raise ValueError("provenance.paused_tasks must contain exactly OPS and RECEIPTS")

    sparse_raw = provenance.get("sparse_task_system_inventory")
    if not isinstance(sparse_raw, list):
        raise ValueError("provenance.sparse_task_system_inventory must be a list")
    sparse_inventory: set[tuple[str, str, str]] = set()
    for item in sparse_raw:
        if not isinstance(item, dict) or set(item) != {"cohort", "task_id", "system", "reason"}:
            raise ValueError("Each sparse inventory entry must contain cohort/task_id/system/reason")
        cohort = str(item.get("cohort") or "")
        task_id = str(item.get("task_id") or "")
        system = str(item.get("system") or "")
        reason = str(item.get("reason") or "")
        if cohort not in {"historical_baseline", "diagnostic_only"}:
            raise ValueError("Only historical or diagnostic inventory may be declared sparse")
        if task_id not in EXPECTED_TASK_IDS or system not in EXPECTED_SYSTEMS:
            raise ValueError("Sparse inventory contains an undeclared task or system")
        safe_token(reason, label="sparse inventory reason")
        key = (cohort, task_id, system)
        if key in sparse_inventory:
            raise ValueError("provenance.sparse_task_system_inventory contains a duplicate")
        sparse_inventory.add(key)

    harbor_inventory = provenance.get("harbor_trace_inventory")
    harbor_fields = {
        "requested_score_only_trial_count",
        "trace_found_count",
        "trace_missing_count",
        "judge_found_count",
        "judge_missing_count",
        "inventory_sha256",
    }
    if not isinstance(harbor_inventory, dict) or set(harbor_inventory) != harbor_fields:
        raise ValueError("provenance.harbor_trace_inventory has an unexpected schema")
    harbor_counts: dict[str, int] = {}
    for field in harbor_fields - {"inventory_sha256"}:
        value = harbor_inventory.get(field)
        if not isinstance(value, int) or isinstance(value, bool) or value < 0:
            raise ValueError(f"provenance.harbor_trace_inventory.{field} must be a nonnegative integer")
        harbor_counts[field] = value
    requested = harbor_counts["requested_score_only_trial_count"]
    if requested != counts["score_only_count"]:
        raise ValueError("Harbor inventory request count must equal score_only_count")
    if harbor_counts["trace_found_count"] + harbor_counts["trace_missing_count"] != requested:
        raise ValueError("Harbor trace inventory counts do not close")
    if harbor_counts["judge_found_count"] + harbor_counts["judge_missing_count"] != requested:
        raise ValueError("Harbor Judge inventory counts do not close")
    inventory_sha256 = str(harbor_inventory.get("inventory_sha256") or "").strip().lower()
    if not HEX64.fullmatch(inventory_sha256):
        raise ValueError("provenance.harbor_trace_inventory.inventory_sha256 must be SHA-256")

    expected_contracts = {
        "aggregation_contract": EXPECTED_AGGREGATION_CONTRACT,
        "pass_at_8_contract": EXPECTED_PASS_AT_8_CONTRACT,
        "missing_value_contract": EXPECTED_MISSING_VALUE_CONTRACT,
        "cost_contract": EXPECTED_COST_CONTRACT,
    }
    for field, expected in expected_contracts.items():
        if provenance.get(field) != expected:
            raise ValueError(f"provenance.{field} does not match the frozen contract")
    return provenance, paused_ids, sparse_inventory


def write_csv(path: Path, fields: list[str], rows: list[dict[str, str]]) -> None:
    with path.open("x", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, extrasaction="raise", lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
        handle.flush()
        os.fsync(handle.fileno())


def render_readme(
    *,
    cutoff: datetime,
    source_controller_head: str,
    evaluator_data_commit: str,
    expected_attempt_records: int,
    paused_task_ids: set[str],
) -> str:
    paused = "\n".join(f"- `{task_id}`" for task_id in sorted(paused_task_ids))
    return f"""# P15 Post-freeze Aggregate Snapshot

This directory is an aggregate-only export from the frozen P15 evaluation.
It compares three complete agent systems under the recorded campaign contract;
it is not a ranking of base models in isolation.

## Fixed provenance

- Evidence cutoff: `{cutoff.isoformat()}`
- Source controller commit: `{source_controller_head}`
- Frozen evaluator commit: `{FROZEN_EVALUATOR_COMMIT}`
- Evaluator data commit: `{evaluator_data_commit}`
- Frozen Judge version: `{FROZEN_JUDGE_VERSION}`
- Attempt records validated: `{expected_attempt_records}`
- Frozen score source: `{FROZEN_SCORE_SOURCE}`
- Original score source: `ORIGINAL_HARBOR_JUDGE`
- Receipt score source: `GUARD_FINAL_RECEIPT`

## Published tables

- `system_summary.csv`: cohort-separated system aggregates.
- `track_summary.csv`: cohort-, track-, and system-level aggregates.
- `task_system_summary.csv`: task-by-system aggregates and task pause markers.
- `cost_time_summary.csv`: separately labeled cost/unit and elapsed-time aggregates.
- `trace_summary.csv`: redacted counts of steps, tokens, and tool calls.
- `SNAPSHOT.json`: public provenance and release boundaries.

## Ranking exclusions

The following task contracts are paused. Their rows remain visible in the task
table, carry `ranking_eligible=false`, and do not contribute to ranking fields:

{paused}

Cohorts remain separate. Historical scores and frozen scores are labeled by
source and are not automatically pooled. A missing score stays empty. A Judge
or infrastructure error is counted as unscored and is never converted to zero.

Cost fields retain their stated units and bases. Provider cumulative snapshots
are reported only as latest observations; they are not summed as per-run cost.

## Public boundary

The export contains no attempt/run identifiers, workbook or trajectory
locations, provider transport identifiers or addresses, machine-local
locations, authentication material, or model-generated text.
"""


def fsync_directory(path: Path) -> None:
    flags = os.O_RDONLY
    if hasattr(os, "O_DIRECTORY"):
        flags |= os.O_DIRECTORY
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW
    descriptor = os.open(path, flags)
    try:
        os.fsync(descriptor)
    finally:
        os.close(descriptor)


def atomic_rename_no_replace(source: Path, destination: Path) -> None:
    """Atomically publish a directory without replacing an existing target."""

    if source.parent != destination.parent:
        raise ValueError("Atomic publication requires sibling source and destination")
    reject_symlink_components(destination.parent, label="output parent")
    flags = os.O_RDONLY
    if hasattr(os, "O_DIRECTORY"):
        flags |= os.O_DIRECTORY
    if hasattr(os, "O_NOFOLLOW"):
        flags |= os.O_NOFOLLOW
    parent_fd = os.open(destination.parent, flags)
    try:
        libc = ctypes.CDLL(None, use_errno=True)
        if sys.platform == "darwin" and hasattr(libc, "renameatx_np"):
            rename = libc.renameatx_np
            no_replace_flag = 0x00000004  # RENAME_EXCL
        elif sys.platform.startswith("linux") and hasattr(libc, "renameat2"):
            rename = libc.renameat2
            no_replace_flag = 0x00000001  # RENAME_NOREPLACE
        else:
            raise RuntimeError("Atomic no-replace directory publication is unavailable")
        rename.argtypes = [
            ctypes.c_int,
            ctypes.c_char_p,
            ctypes.c_int,
            ctypes.c_char_p,
            ctypes.c_uint,
        ]
        rename.restype = ctypes.c_int
        result = rename(
            parent_fd,
            os.fsencode(source.name),
            parent_fd,
            os.fsencode(destination.name),
            no_replace_flag,
        )
        if result == 0:
            return
        error = ctypes.get_errno()
        if error in {errno.EEXIST, errno.ENOTEMPTY}:
            raise FileExistsError(destination)
        raise OSError(error, os.strerror(error), str(destination))
    finally:
        os.close(parent_fd)


def require_private_directory(path: Path) -> Path:
    absolute = absolute_without_resolving(path)
    reject_symlink_components(absolute, label="private analyzer directory")
    try:
        state = absolute.lstat()
    except OSError as exc:
        raise ValueError("Private analyzer output directory does not exist") from exc
    if not stat.S_ISDIR(state.st_mode):
        raise ValueError("Private analyzer output directory does not exist")
    return absolute


def prepare_output_directory(path: Path, *, private_dir: Path) -> Path:
    output = absolute_without_resolving(path)
    try:
        output.relative_to(private_dir)
    except ValueError:
        pass
    else:
        raise ValueError("Output directory cannot be inside the private analyzer input")
    reject_symlink_components(output, label="output directory")
    try:
        output.lstat()
    except FileNotFoundError:
        pass
    except OSError as exc:
        raise ValueError("Cannot inspect output directory") from exc
    else:
        raise FileExistsError(f"Refusing to overwrite existing output directory: {output}")
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise ValueError("Cannot create output parent directory") from exc
    reject_symlink_components(output.parent, label="output parent")
    try:
        parent_state = output.parent.lstat()
    except OSError as exc:
        raise ValueError("Cannot inspect output parent directory") from exc
    if not stat.S_ISDIR(parent_state.st_mode):
        raise ValueError("Output parent is not a directory")
    try:
        output.lstat()
    except FileNotFoundError:
        return output
    raise FileExistsError(f"Refusing to overwrite existing output directory: {output}")


def export_snapshot(
    private_dir: Path,
    *,
    source_controller_head: str,
    expected_attempt_records: int,
    evidence_cutoff: str,
    output_dir: Path,
) -> dict[str, Any]:
    private_dir = require_private_directory(private_dir)
    source_head = full_sha(source_controller_head, label="--source-controller-head")
    if (
        not isinstance(expected_attempt_records, int)
        or isinstance(expected_attempt_records, bool)
        or expected_attempt_records < 0
    ):
        raise ValueError("--expected-attempt-records must be nonnegative")
    cutoff = aware_time(evidence_cutoff, label="--evidence-cutoff")
    private_file_snapshot = capture_private_file_snapshot(private_dir)

    provenance, paused_task_ids, sparse_inventory = validate_provenance(
        private_dir,
        source_controller_head=source_head,
        expected_attempt_records=expected_attempt_records,
        evidence_cutoff=cutoff,
    )
    read_and_validate_attempts(private_dir / "attempts.csv", expected_attempt_records)
    tables = {
        filename: read_csv_exact(private_dir / filename, fields)
        for filename, fields in PUBLIC_TABLES.items()
    }
    validate_aggregate_tables(tables, paused_task_ids, sparse_inventory)
    require_private_file_snapshot(
        private_dir, private_file_snapshot, phase="after parsing"
    )
    output_dir = prepare_output_directory(output_dir, private_dir=private_dir)

    manifest = {
        "schema_version": PUBLIC_SCHEMA_VERSION,
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "evidence_cutoff": cutoff.isoformat(),
        "source_controller_head": source_head,
        "expected_attempt_records": expected_attempt_records,
        "private_analysis": {
            "schema_version": provenance["schema_version"],
            "judge_version": provenance["judge_version"],
            "evaluator_commit": FROZEN_EVALUATOR_COMMIT,
            "evaluator_data_commit": FROZEN_EVALUATOR_DATA_COMMIT,
            "private_provenance_sha256": private_file_snapshot[
                "provenance.json"
            ].sha256,
            "private_input_sha256": dict(sorted(provenance["input_hashes"].items())),
            "private_output_sha256": dict(sorted(provenance["output_hashes"].items())),
            "harbor_trace_inventory_validated": {
                key: provenance["harbor_trace_inventory"][key]
                for key in (
                    "requested_score_only_trial_count",
                    "trace_found_count",
                    "trace_missing_count",
                    "judge_found_count",
                    "judge_missing_count",
                    "inventory_sha256",
                )
            },
        },
        "score_sources": provenance["score_sources"],
        "declared_inventory": {
            "systems": sorted(EXPECTED_SYSTEMS),
            "task_ids": sorted(EXPECTED_TASK_IDS),
            "cohorts": sorted(EXPECTED_COHORTS),
        },
        "sparse_task_system_inventory": [
            {"cohort": cohort, "task_id": task_id, "system": system}
            for cohort, task_id, system in sorted(sparse_inventory)
        ],
        "paused_tasks": [
            {"task_id": task_id, "reason": "TASK_CONTRACT_PAUSED", "ranking_eligible": False}
            for task_id in sorted(paused_task_ids)
        ],
        "aggregation": {
            "cohorts_kept_separate": True,
            "historical_and_frozen_scores_automatically_pooled": False,
            "paused_tasks_excluded_from_rankings": True,
            "diagnostic_only_excluded_from_rankings": True,
            "historical_baseline_descriptive_only": True,
            "ranking_cohort": "current_contract",
            "missing_scores_coerced_to_zero": False,
        },
        "aggregate_row_counts": {name: len(rows) for name, rows in tables.items()},
        "published_files": [*PUBLIC_TABLES, "SNAPSHOT.json", "README.md"],
        "public_boundary": {
            "attempt_level_rows_included": False,
            "run_or_attempt_ids_included": False,
            "artifact_or_trace_locations_included": False,
            "provider_transport_metadata_included": False,
            "machine_local_locations_included": False,
            "authentication_material_included": False,
            "model_generated_text_included": False,
        },
    }

    reject_symlink_components(output_dir.parent, label="output parent")
    stage = Path(tempfile.mkdtemp(prefix=f".{output_dir.name}.staging-", dir=output_dir.parent))
    try:
        for filename, fields in PUBLIC_TABLES.items():
            write_csv(stage / filename, fields, tables[filename])
        readme_path = stage / "README.md"
        with readme_path.open("x", encoding="utf-8") as handle:
            handle.write(
                render_readme(
                    cutoff=cutoff,
                    source_controller_head=source_head,
                    evaluator_data_commit=manifest["private_analysis"]["evaluator_data_commit"],
                    expected_attempt_records=expected_attempt_records,
                    paused_task_ids=paused_task_ids,
                )
            )
            handle.flush()
            os.fsync(handle.fileno())
        manifest["published_file_sha256"] = {
            filename: sha256_file(stage / filename)
            for filename in [*PUBLIC_TABLES, "README.md"]
        }
        snapshot_path = stage / "SNAPSHOT.json"
        with snapshot_path.open("x", encoding="utf-8") as handle:
            json.dump(manifest, handle, indent=2, sort_keys=True, ensure_ascii=False)
            handle.write("\n")
            handle.flush()
            os.fsync(handle.fileno())
        if set(path.name for path in stage.iterdir()) != set(manifest["published_files"]):
            raise ValueError("Staging directory contains a non-allowlisted file")
        for public_path in stage.iterdir():
            public_text = read_stable_bytes(
                public_path, label=f"staged public file {public_path.name}"
            )[0].decode("utf-8")
            for sensitive_label, pattern in SENSITIVE_PATTERNS:
                if pattern.search(public_text):
                    raise ValueError(
                        f"Staged public output contains {sensitive_label}: "
                        f"{public_path.name}"
                    )
        require_private_file_snapshot(
            private_dir, private_file_snapshot, phase="before atomic publication"
        )
        reject_symlink_components(output_dir, label="output directory")
        fsync_directory(stage)
        atomic_rename_no_replace(stage, output_dir)
        fsync_directory(output_dir.parent)
    except Exception:
        if stage.exists():
            shutil.rmtree(stage)
        raise
    return manifest


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("private_output_dir", nargs="?", type=Path)
    parser.add_argument("--input-dir", dest="private_output_dir_option", type=Path)
    parser.add_argument("--source-controller-head", required=True)
    parser.add_argument("--expected-attempt-records", required=True, type=int)
    parser.add_argument("--evidence-cutoff", required=True)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args(argv)
    choices = [value for value in (args.private_output_dir, args.private_output_dir_option) if value is not None]
    if len(choices) != 1:
        parser.error("provide exactly one private analyzer directory, positionally or with --input-dir")
    args.private_output_dir = choices[0]
    return args


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    manifest = export_snapshot(
        args.private_output_dir,
        source_controller_head=args.source_controller_head,
        expected_attempt_records=args.expected_attempt_records,
        evidence_cutoff=args.evidence_cutoff,
        output_dir=args.output_dir,
    )
    print(
        json.dumps(
            {
                "status": "EXPORTED",
                "schema_version": manifest["schema_version"],
                "aggregate_row_counts": manifest["aggregate_row_counts"],
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
