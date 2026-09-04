#!/usr/bin/env python3
"""Small, self-contained support layer for the frozen P15 Judge v2 contract.

Each Harbor verifier receives a copy of this module inside its own ``tests``
directory.  It deliberately contains no task semantics, ontology, model call,
or reference-layout guessing.
"""
from __future__ import annotations

import math
import re
from typing import Any


def sheet_token(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).casefold())


class RoleMappedWorkbook:
    """Read-only sheet-role view that preserves the candidate workbook."""

    def __init__(self, workbook: Any, role_map: dict[str, str]):
        self._workbook = workbook
        self.role_map = dict(role_map)

    @property
    def sheetnames(self) -> list[str]:
        return list(dict.fromkeys([*self._workbook.sheetnames, *self.role_map]))

    @property
    def worksheets(self) -> list[Any]:
        return self._workbook.worksheets

    def actual_sheet_name(self, role: str) -> str:
        return self.role_map.get(role, role)

    def __getitem__(self, role: str) -> Any:
        return self._workbook[self.actual_sheet_name(role)]

    def __getattr__(self, name: str) -> Any:
        return getattr(self._workbook, name)


def resolve_sheet_roles(
    workbook: Any,
    required_sheets: list[str],
    aliases: dict[str, tuple[str, ...]] | None = None,
) -> tuple[RoleMappedWorkbook, dict[str, str], list[str], list[tuple[str, tuple[str, ...]]]]:
    """Resolve exact names, task-local aliases, then normalized exact names.

    A physical sheet may satisfy at most one role.  A sheet already claimed by
    an earlier role is not available to a later one, including when the later
    role's own name matches it exactly -- otherwise one physical sheet would
    silently satisfy two roles with no failure code.  Unresolved or ambiguous
    roles are returned to the task Judge for explicit review; they are never
    guessed.
    """
    aliases = aliases or {}
    role_map: dict[str, str] = {}
    used: set[str] = set()
    unresolved: list[str] = []
    ambiguous: list[tuple[str, tuple[str, ...]]] = []
    for role in required_sheets:
        if role in workbook.sheetnames and role not in used:
            used.add(role)
            continue
        allowed = set(aliases.get(role, ()))
        allowed_tokens = {sheet_token(value) for value in allowed}
        normalized_role = sheet_token(role)
        candidates = [
            sheet
            for sheet in workbook.sheetnames
            if sheet not in used
            and (
                sheet in allowed
                or sheet_token(sheet) == normalized_role
                or sheet_token(sheet) in allowed_tokens
            )
        ]
        if len(candidates) == 1:
            role_map[role] = candidates[0]
            used.add(candidates[0])
        elif candidates:
            ambiguous.append((role, tuple(candidates)))
        else:
            unresolved.append(role)
    return RoleMappedWorkbook(workbook, role_map), role_map, unresolved, ambiguous


def sheet_resolution_failures(
    role_map: dict[str, str],
    unresolved: list[str],
    ambiguous: list[tuple[str, tuple[str, ...]]],
) -> list[str]:
    # A resolved alias is a supported equivalent layout, not a failure.
    failures: list[str] = []
    failures.extend(f"UNRESOLVED_SHEET_ROLE:{role}" for role in unresolved)
    failures.extend(
        f"AMBIGUOUS_SHEET_ROLE:{role}={','.join(options)}"
        for role, options in ambiguous
    )
    return failures


def evaluation_status(failures: list[str], explicit: str | None = None) -> str:
    if explicit is not None:
        if explicit not in {"OK", "JUDGE_ERROR", "INFRA_ERROR"}:
            raise ValueError(f"INVALID_EVALUATION_STATUS:{explicit}")
        return explicit
    # Ambiguous and unresolved semantic layouts are evaluator limitations until
    # a task-local rule can distinguish an omitted deliverable from a materially
    # equivalent alternative layout.  Do not turn that uncertainty into a model
    # zero.  A plain OUTPUT_MISSING or malformed workbook remains a scorable
    # model-delivery failure.
    judge_prefixes = (
        "AMBIGUOUS_SHEET_ROLE:",
        "UNRESOLVED_SHEET_ROLE:",
        "UNSUPPORTED_LAYOUT:",
        "UNSUPPORTED_FORMULA:",
        "NATIVE_EXCEL_RECALC_REQUIRED:",
        "SEMANTIC_EVALUATION_ERROR:",
        "EVALUATION_ERROR:",
    )
    judge_failure = any(str(code).startswith(judge_prefixes) for code in failures)
    judge_failure = judge_failure or any(
        "UnsupportedFormulaError" in str(code) for code in failures
    )
    return "JUDGE_ERROR" if judge_failure else "OK"


def weighted_score(
    task: dict[str, Any], criteria: dict[str, float], split: str
) -> float:
    definitions = [
        row
        for row in task["criteria"]
        if split in row.get("method_params", {}).get("applies_to", [split])
    ]
    criterion_weight = lambda row: row.get("method_params", {}).get(
        "split_weights", {}
    ).get(split, row["weight"])

    criterion_ids = [str(row.get("id", "")) for row in definitions]
    if any(not criterion_id for criterion_id in criterion_ids):
        raise ValueError("EMPTY_CRITERION_ID")
    if len(set(criterion_ids)) != len(criterion_ids):
        raise ValueError("DUPLICATE_CRITERION_IDS")

    positive_ids = {
        row["id"] for row in definitions if row.get("type") == "positive"
    }
    hurdle_ids = list(task.get("hurdle_criteria", []))
    if len(set(hurdle_ids)) != len(hurdle_ids):
        raise ValueError("DUPLICATE_HURDLE_CRITERIA")
    unknown_hurdles = sorted(set(hurdle_ids) - positive_ids)
    if unknown_hurdles:
        raise ValueError(
            "INVALID_HURDLE_CRITERIA:" + ",".join(unknown_hurdles)
        )

    threshold = task.get("pass_threshold")
    if (
        not isinstance(threshold, (int, float))
        or isinstance(threshold, bool)
        or not math.isfinite(float(threshold))
        or not 0.0 <= float(threshold) <= 1.0
    ):
        raise ValueError("INVALID_PASS_THRESHOLD")

    invalid_values = []
    for criterion_id in criterion_ids:
        if criterion_id not in criteria:
            continue
        value = criteria[criterion_id]
        if (
            not isinstance(value, (int, float))
            or isinstance(value, bool)
            or not math.isfinite(float(value))
            or not 0.0 <= float(value) <= 1.0
        ):
            invalid_values.append(criterion_id)
    if invalid_values:
        raise ValueError(
            "INVALID_CRITERION_SCORES:" + ",".join(sorted(invalid_values))
        )

    invalid_weights = []
    for row in definitions:
        weight = criterion_weight(row)
        if (
            not isinstance(weight, (int, float))
            or isinstance(weight, bool)
            or not math.isfinite(float(weight))
            or (row.get("type") == "positive" and float(weight) <= 0.0)
            or (row.get("type") == "penalty" and float(weight) >= 0.0)
            or row.get("type") not in {"positive", "penalty"}
        ):
            invalid_weights.append(str(row.get("id", "")))
    if invalid_weights:
        raise ValueError(
            "INVALID_CRITERION_WEIGHTS:" + ",".join(sorted(invalid_weights))
        )

    positive_total = sum(
        criterion_weight(row) for row in definitions if row["type"] == "positive"
    )
    if positive_total <= 0:
        raise ValueError("POSITIVE_WEIGHT_TOTAL_MUST_BE_POSITIVE")
    missing = [
        row["id"] for row in definitions if row["id"] not in criteria
    ]
    if missing:
        # "Every independently judgeable positive criterion runs."  Treating an
        # absent criterion as a scored zero is indistinguishable from one that
        # ran and failed, so refuse instead of quietly deflating the score.
        raise ValueError("MISSING_CRITERION_SCORES:" + ",".join(sorted(missing)))
    raw = sum(
        criterion_weight(row) * float(criteria.get(row["id"], 0.0))
        for row in definitions
        if row["type"] == "positive"
    )
    raw += sum(
        criterion_weight(row) * float(criteria.get(row["id"], 0.0))
        for row in definitions
        if row["type"] == "penalty"
    )
    value = raw / positive_total
    if not math.isfinite(value):
        raise ValueError("NONFINITE_NORMALIZED_SCORE")
    return round(max(0.0, min(1.0, value)), 6)


def build_result(
    *,
    task: dict[str, Any],
    split: str,
    candidate: str,
    criteria: dict[str, float],
    failures: list[str],
    explicit_status: str | None = None,
) -> dict[str, Any]:
    status = evaluation_status(failures, explicit_status)
    normalized = None
    if status == "OK":
        try:
            normalized = weighted_score(task, criteria, split)
        except ValueError as exc:
            # An incomplete criterion set is a Judge fault, not a model outcome.
            status = "JUDGE_ERROR"
            failures = [*failures, str(exc)]
    hurdle_failures = (
        [
            criterion
            for criterion in task.get("hurdle_criteria", [])
            if float(criteria.get(criterion, 0.0)) < 1.0
        ]
        if status == "OK"
        else []
    )
    passed = bool(
        status == "OK"
        and normalized is not None
        and normalized >= float(task["pass_threshold"])
        and not hurdle_failures
    )
    return {
        "task_id": task["task_id"],
        "split": split,
        "candidate": candidate,
        # Keep the legacy operational status consumed by Harbor/controller
        # while exposing the orthogonal v2 attribution status separately.
        "status": "SCORED" if status == "OK" else status,
        "evaluation_status": status,
        "normalized_score": normalized,
        "pass": passed,
        "hurdle_failures": hurdle_failures,
        "criterion_scores": criteria,
        "failure_codes": sorted(set(failures)),
        "stderr": [],
    }
