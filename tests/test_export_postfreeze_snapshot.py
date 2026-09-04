from __future__ import annotations

import csv
import hashlib
import importlib.util
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


SCRIPT = Path(__file__).resolve().parents[1] / "scripts/export_postfreeze_snapshot.py"
SPEC = importlib.util.spec_from_file_location("export_postfreeze_snapshot", SCRIPT)
assert SPEC and SPEC.loader
EXPORTER = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = EXPORTER
SPEC.loader.exec_module(EXPORTER)


def write_csv(path: Path, fields: list[str], rows: list[dict[str, str]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def digest(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def score_row(**overrides: str) -> dict[str, str]:
    row = {
        "cohort": "current_contract",
        "system": "codex_gpt56sol",
        "ranking_eligible": "true",
        "score_source": EXPORTER.FROZEN_SCORE_SOURCE,
        "artifact_count": "1",
        "scored_count": "1",
        "unscored_count": "0",
        "score_coverage": "1",
        "original_scored_count": "1",
        "original_mean_score_scored_only": "0.5",
        "frozen_mean_score_scored_only": "0.75",
        "frozen_mean_score_lower_bound": "0.75",
        "frozen_mean_score_upper_bound": "0.75",
        "pass_count": "1",
        "pass_rate_scored_only": "1",
        "judge_error_count": "0",
        "infra_error_count": "0",
    }
    row.update(overrides)
    return row


def paused_task_row(
    cohort: str,
    task_id: str,
    track: str,
    system: str,
) -> dict[str, str]:
    row = score_row(
        cohort=cohort,
        task_id=task_id,
        track=track,
        system=system,
        task_status="PAUSED_TASK_CONTRACT",
        ranking_eligible="false",
        independent_complete_scored_count="0",
        scored_count="0",
        unscored_count="1",
        score_coverage="0",
        original_scored_count="0",
        original_mean_score_scored_only="",
        frozen_mean_score_scored_only="null",
        frozen_mean_score_lower_bound="",
        frozen_mean_score_upper_bound="",
        pass_count="0",
        pass_rate_scored_only="",
        judge_error_count="0",
        infra_error_count="1",
        pass_at_8="",
        pass_at_8_status="PAUSED_TASK",
    )
    return row


def cost_row(cohort: str, system: str) -> dict[str, str]:
    row = {field: "" for field in EXPORTER.COST_TIME_FIELDS}
    row.update(
        {
            "cohort": cohort,
            "system": system,
            "ranking_eligible": str(cohort == "current_contract").lower(),
            "score_source": EXPORTER.FROZEN_SCORE_SOURCE,
            "attempt_count": "3",
            "spent_units_reported_count": "2",
            "spent_units_sum": "1.5",
            "spent_units_mean": "0.75",
            "gateway_spent_units_reported_count": "2",
            "gateway_spent_units_sum": "1.5",
            "gateway_spent_units_mean": "0.75",
            "token_estimated_units_reported_count": "1",
            "token_estimated_units_sum": "0.4",
            "token_estimated_units_mean": "0.4",
            "harbor_reported_cost_usd_reported_count": "1",
            "harbor_reported_cost_usd_sum": "0.3",
            "harbor_reported_cost_usd_mean": "0.3",
            "provider_total_units_snapshot_reported_count": "1",
            "provider_total_units_snapshot_latest": "12.5",
            "provider_total_units_snapshot_latest_observed_at": "2026-09-04T01:00:00+00:00",
            "run_wall_seconds_reported_count": "3",
            "run_wall_seconds_sum": "90",
            "run_wall_seconds_mean": "30",
            "harbor_wall_seconds_reported_count": "2",
            "harbor_wall_seconds_sum": "60",
            "harbor_wall_seconds_mean": "30",
            "agent_execution_seconds_reported_count": "2",
            "agent_execution_seconds_sum": "50",
            "agent_execution_seconds_mean": "25",
        }
    )
    return row


def trace_row(cohort: str, system: str) -> dict[str, str]:
    row = {field: "0" for field in EXPORTER.TRACE_FIELDS}
    row.update(
        {
            "cohort": cohort,
            "system": system,
            "ranking_eligible": str(cohort == "current_contract").lower(),
            "score_source": EXPORTER.FROZEN_SCORE_SOURCE,
            "attempt_count": "3",
            "trace_present_count": "2",
            "trace_missing_count": "1",
            "step_count_reported_count": "2",
            "total_steps": "8",
            "mean_steps": "4",
            "prompt_tokens_reported_count": "2",
            "total_prompt_tokens": "100",
            "cached_tokens_reported_count": "2",
            "total_cached_tokens": "20",
            "completion_tokens_reported_count": "2",
            "total_completion_tokens": "30",
            "total_tool_calls": "4",
            "total_completed_tool_calls": "4",
        }
    )
    return row


def create_fixture(root: Path) -> Path:
    private = root / "private"
    private.mkdir(parents=True)

    system_rows = [
        score_row(
            cohort=cohort,
            system=system,
            ranking_eligible=str(cohort == "current_contract").lower(),
        )
        for cohort in sorted(EXPORTER.EXPECTED_COHORTS)
        for system in sorted(EXPORTER.EXPECTED_SYSTEMS)
    ]
    write_csv(private / "system_summary.csv", EXPORTER.SYSTEM_FIELDS, system_rows)
    track_rows = [
        score_row(
            cohort=cohort,
            track=track,
            system=system,
            ranking_eligible=str(cohort == "current_contract").lower(),
        )
        for cohort in sorted(EXPORTER.EXPECTED_COHORTS)
        for system in sorted(EXPORTER.EXPECTED_SYSTEMS)
        for track in ("A", "B", "C")
    ]
    write_csv(
        private / "track_summary.csv",
        EXPORTER.TRACK_FIELDS,
        track_rows,
    )
    task_rows: list[dict[str, str]] = []
    for cohort in sorted(EXPORTER.EXPECTED_COHORTS):
        for task_id in sorted(EXPORTER.EXPECTED_TASK_IDS):
            track = task_id.split("-")[1]
            for system in sorted(EXPORTER.EXPECTED_SYSTEMS):
                if task_id in EXPORTER.REQUIRED_PAUSED_TASKS:
                    task_rows.append(paused_task_row(cohort, task_id, track, system))
                else:
                    task_rows.append(
                        score_row(
                            cohort=cohort,
                            task_id=task_id,
                            track=track,
                            system=system,
                            task_status=(
                                "DIAGNOSTIC_ONLY"
                                if cohort == "diagnostic_only"
                                else "HISTORICAL_BASELINE_DESCRIPTIVE"
                                if cohort == "historical_baseline"
                                else "ACTIVE"
                            ),
                            ranking_eligible=str(cohort == "current_contract").lower(),
                            independent_complete_scored_count="1",
                            pass_at_8="",
                            pass_at_8_status=(
                                "DIAGNOSTIC_ONLY"
                                if cohort == "diagnostic_only"
                                else "HISTORICAL_DESCRIPTIVE_ONLY"
                                if cohort == "historical_baseline"
                                else "N_LT_8"
                            ),
                        )
                    )
    write_csv(private / "task_system_summary.csv", EXPORTER.TASK_SYSTEM_FIELDS, task_rows)
    cost_rows = [
        cost_row(cohort, system)
        for cohort in sorted(EXPORTER.EXPECTED_COHORTS)
        for system in sorted(EXPORTER.EXPECTED_SYSTEMS)
    ]
    write_csv(private / "cost_time_summary.csv", EXPORTER.COST_TIME_FIELDS, cost_rows)
    trace_rows = [
        trace_row(cohort, system)
        for cohort in sorted(EXPORTER.EXPECTED_COHORTS)
        for system in sorted(EXPORTER.EXPECTED_SYSTEMS)
    ]
    write_csv(private / "trace_summary.csv", EXPORTER.TRACE_FIELDS, trace_rows)

    attempts: list[dict[str, str]] = []
    for status, score, passed in (("SCORED", "0.75", "true"), ("JUDGE_ERROR", "", ""), ("INFRA_ERROR", "", "")):
        row = {field: "" for field in EXPORTER.ATTEMPT_FIELDS}
        row.update(
            {
                "task_id": "P15-A-FIN-DCF-001",
                "track": "A",
                "system": "codex_gpt56sol",
                "cohort": "current_contract",
                "frozen_evaluation_status": status,
                "frozen_normalized_score": score,
                "frozen_pass": passed,
            }
        )
        attempts.append(row)
    write_csv(private / "attempts.csv", EXPORTER.ATTEMPT_FIELDS, attempts)
    (private / "criteria_long.csv").write_text("criterion_id\n", encoding="utf-8")

    output_hashes = {name: digest(private / name) for name in EXPORTER.PRIVATE_HASHED_OUTPUTS}
    provenance = {
        "schema_version": EXPORTER.PRIVATE_SCHEMA_VERSION,
        "generated_at_utc": "2026-09-04T01:15:00+00:00",
        "evidence_cutoff": "2026-09-04T05:10:00+00:00",
        "attempt_count": 3,
        "evidence_row_count": 3,
        "score_row_count": 2,
        "joined_score_count": 2,
        "evidence_only_count": 1,
        "score_only_count": 0,
        "judge_version": EXPORTER.FROZEN_JUDGE_VERSION,
        "score_source": EXPORTER.FROZEN_SCORE_SOURCE,
        "evaluator_commit": EXPORTER.FROZEN_EVALUATOR_COMMIT,
        "evaluator_data_commit": EXPORTER.FROZEN_EVALUATOR_DATA_COMMIT,
        "source_controller_head": "a" * 40,
        "systems": sorted(EXPORTER.EXPECTED_SYSTEMS),
        "task_ids": sorted(EXPORTER.EXPECTED_TASK_IDS),
        "cohorts": sorted(EXPORTER.EXPECTED_COHORTS),
        "sparse_task_system_inventory": [],
        "paused_tasks": [
            {"task_id": task_id, "reason": "TASK_CONTRACT_PAUSED", "ranking_eligible": False}
            for task_id in sorted(EXPORTER.REQUIRED_PAUSED_TASKS)
        ],
        "score_sources": {
            "original": "ORIGINAL_HARBOR_JUDGE",
            "receipt": "GUARD_FINAL_RECEIPT",
            "frozen": EXPORTER.FROZEN_SCORE_SOURCE,
        },
        "input_hashes": {
            "evidence_index.csv": "1" * 64,
            "historical_reuse_audit.csv": "4" * 64,
            "scores.csv": "2" * 64,
            "summary.json": "3" * 64,
        },
        "harbor_trace_inventory": {
            "requested_score_only_trial_count": 0,
            "trace_found_count": 0,
            "trace_missing_count": 0,
            "judge_found_count": 0,
            "judge_missing_count": 0,
            "inventory_sha256": "5" * 64,
        },
        "output_hashes": output_hashes,
        "aggregation_contract": EXPORTER.EXPECTED_AGGREGATION_CONTRACT,
        "pass_at_8_contract": EXPORTER.EXPECTED_PASS_AT_8_CONTRACT,
        "missing_value_contract": EXPORTER.EXPECTED_MISSING_VALUE_CONTRACT,
        "cost_contract": EXPORTER.EXPECTED_COST_CONTRACT,
    }
    (private / "provenance.json").write_text(
        json.dumps(provenance, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    return private


def update_provenance(private: Path, mutate) -> None:
    path = private / "provenance.json"
    value = json.loads(path.read_text(encoding="utf-8"))
    mutate(value)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def rewrite_csv(private: Path, filename: str, fields: list[str], mutate) -> None:
    path = private / filename
    with path.open(newline="", encoding="utf-8") as handle:
        rows = list(csv.DictReader(handle))
    mutate(rows)
    write_csv(path, fields, rows)
    update_provenance(
        private,
        lambda value: value["output_hashes"].__setitem__(filename, digest(path)),
    )


class ExportPostfreezeSnapshotTest(unittest.TestCase):
    def test_public_export_is_aggregate_only_and_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            private = create_fixture(root)
            output = root / "postfreeze_fixture"
            manifest = EXPORTER.export_snapshot(
                private,
                source_controller_head="a" * 40,
                expected_attempt_records=3,
                evidence_cutoff="2026-09-04T01:10:00-04:00",
                output_dir=output,
            )

            self.assertEqual(set(path.name for path in output.iterdir()), set(manifest["published_files"]))
            self.assertFalse((output / "attempts.csv").exists())
            self.assertFalse((output / "criteria_long.csv").exists())
            public_text = "\n".join(path.read_text(encoding="utf-8") for path in output.iterdir())
            for forbidden in ("/Users/", "https://", "20260904T011500Z-P15-", "req_fixture_secret"):
                self.assertNotIn(forbidden, public_text)
            for csv_path in output.glob("*.csv"):
                with csv_path.open(newline="", encoding="utf-8") as handle:
                    header = next(csv.reader(handle))
                self.assertNotIn("attempt_id", header)
                self.assertNotIn("run_id", header)
                self.assertNotIn("artifact_sha256", header)

            with (output / "task_system_summary.csv").open(newline="", encoding="utf-8") as handle:
                rows = list(csv.DictReader(handle))
            paused = [row for row in rows if row["task_status"] == "PAUSED_TASK_CONTRACT"]
            self.assertEqual({row["task_id"] for row in paused}, EXPORTER.REQUIRED_PAUSED_TASKS)
            self.assertTrue(all(row["ranking_eligible"] == "false" for row in paused))
            self.assertTrue(all(row["frozen_mean_score_scored_only"] == "" for row in paused))
            diagnostic = [row for row in rows if row["cohort"] == "diagnostic_only"]
            self.assertTrue(diagnostic)
            self.assertTrue(all(row["ranking_eligible"] == "false" for row in diagnostic))
            historical = [row for row in rows if row["cohort"] == "historical_baseline"]
            self.assertTrue(historical)
            self.assertTrue(all(row["ranking_eligible"] == "false" for row in historical))
            self.assertTrue(
                all(
                    row["task_status"] in {
                        "HISTORICAL_BASELINE_DESCRIPTIVE",
                        "PAUSED_TASK_CONTRACT",
                    }
                    for row in historical
                )
            )
            for filename, expected_hash in manifest["published_file_sha256"].items():
                self.assertEqual(digest(output / filename), expected_hash)
            private_analysis = manifest["private_analysis"]
            self.assertEqual(
                private_analysis["private_provenance_sha256"],
                digest(private / "provenance.json"),
            )
            provenance = json.loads(
                (private / "provenance.json").read_text(encoding="utf-8")
            )
            self.assertEqual(
                private_analysis["private_input_sha256"], provenance["input_hashes"]
            )
            self.assertEqual(
                private_analysis["private_output_sha256"], provenance["output_hashes"]
            )
            self.assertEqual(
                private_analysis["harbor_trace_inventory_validated"]["inventory_sha256"],
                provenance["harbor_trace_inventory"]["inventory_sha256"],
            )

            with self.assertRaises(FileExistsError):
                EXPORTER.export_snapshot(
                    private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=output,
                )

            second_private = create_fixture(root / "second")
            with (second_private / "system_summary.csv").open("a", encoding="utf-8") as handle:
                handle.write("tampered\n")
            with self.assertRaisesRegex(ValueError, "hash mismatch"):
                EXPORTER.export_snapshot(
                    second_private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=root / "must_not_exist",
                )
            self.assertFalse((root / "must_not_exist").exists())

    def test_frozen_provenance_rejections(self) -> None:
        cases = {
            "wrong data commit": lambda value: value.__setitem__("evaluator_data_commit", "c" * 40),
            "extra paused task": lambda value: value["paused_tasks"].append(
                {
                    "task_id": "P15-A-FIN-DCF-001",
                    "reason": "TASK_CONTRACT_PAUSED",
                    "ranking_eligible": False,
                }
            ),
            "wrong diagnostic contract": lambda value: value[
                "aggregation_contract"
            ].__setitem__("diagnostic_only_excluded_from_rankings", False),
            "wrong cost contract": lambda value: value["cost_contract"].__setitem__(
                "cross_basis_total_prohibited", False
            ),
        }
        for name, mutate in cases.items():
            with self.subTest(name=name), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary).resolve()
                private = create_fixture(root)
                update_provenance(private, mutate)
                with self.assertRaises(ValueError):
                    EXPORTER.export_snapshot(
                        private,
                        source_controller_head="a" * 40,
                        expected_attempt_records=3,
                        evidence_cutoff="2026-09-04T05:10:00+00:00",
                        output_dir=root / "rejected",
                    )
                self.assertFalse((root / "rejected").exists())

        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            private = create_fixture(root)
            with self.assertRaisesRegex(ValueError, "evidence_cutoff"):
                EXPORTER.export_snapshot(
                    private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:01+00:00",
                    output_dir=root / "wrong_cutoff",
                )
            with self.assertRaises(ValueError):
                EXPORTER.export_snapshot(
                    private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=True,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=root / "bool_count",
                )

    def test_rejects_private_input_drift_after_initial_hash_check(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            private = create_fixture(root)
            output = root / "must_not_publish"
            original = EXPORTER.read_and_validate_attempts

            def mutate_after_hash_check(path: Path, expected_count: int) -> None:
                original(path, expected_count)
                aggregate = private / "system_summary.csv"
                with aggregate.open(newline="", encoding="utf-8") as handle:
                    rows = list(csv.DictReader(handle))
                rows[0]["original_mean_score_scored_only"] = "0.4"
                write_csv(aggregate, EXPORTER.SYSTEM_FIELDS, rows)

            with mock.patch.object(
                EXPORTER,
                "read_and_validate_attempts",
                side_effect=mutate_after_hash_check,
            ), self.assertRaisesRegex(ValueError, "changed during export"):
                EXPORTER.export_snapshot(
                    private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=output,
                )
            self.assertFalse(output.exists())

    def test_rejects_private_input_drift_before_atomic_publish(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            private = create_fixture(root)
            output = root / "must_not_publish"
            original = EXPORTER.require_private_file_snapshot
            calls = 0

            def mutate_after_parse_check(
                private_dir: Path,
                expected: dict[str, object],
                *,
                phase: str,
            ) -> None:
                nonlocal calls
                original(private_dir, expected, phase=phase)
                calls += 1
                if calls == 1:
                    aggregate = private / "system_summary.csv"
                    with aggregate.open(newline="", encoding="utf-8") as handle:
                        rows = list(csv.DictReader(handle))
                    rows[0]["original_mean_score_scored_only"] = "0.4"
                    write_csv(aggregate, EXPORTER.SYSTEM_FIELDS, rows)

            with mock.patch.object(
                EXPORTER,
                "require_private_file_snapshot",
                side_effect=mutate_after_parse_check,
            ), self.assertRaisesRegex(ValueError, "before atomic publication"):
                EXPORTER.export_snapshot(
                    private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=output,
                )
            self.assertFalse(output.exists())
            self.assertFalse(list(root.glob(".must_not_publish.staging-*")))

    def test_rejects_symlink_paths_and_output_inside_private_input(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            private = create_fixture(root)
            private_link = root / "private-link"
            private_link.symlink_to(private, target_is_directory=True)
            with self.assertRaisesRegex(ValueError, "symlink"):
                EXPORTER.export_snapshot(
                    private_link,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=root / "private-link-output",
                )

            escaped = root / "escaped" / "release"
            dangling_output = root / "dangling-output"
            dangling_output.symlink_to(escaped, target_is_directory=True)
            with self.assertRaisesRegex(ValueError, "symlink"):
                EXPORTER.export_snapshot(
                    private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=dangling_output,
                )
            self.assertFalse(escaped.exists())

            actual_parent = root / "actual-parent"
            actual_parent.mkdir()
            linked_parent = root / "linked-parent"
            linked_parent.symlink_to(actual_parent, target_is_directory=True)
            with self.assertRaisesRegex(ValueError, "symlink"):
                EXPORTER.export_snapshot(
                    private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=linked_parent / "release",
                )
            self.assertFalse((actual_parent / "release").exists())

            nested_output = private / "public-release"
            with self.assertRaisesRegex(ValueError, "inside the private"):
                EXPORTER.export_snapshot(
                    private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=nested_output,
                )
            self.assertFalse(nested_output.exists())

    def test_atomic_publish_refuses_target_appearance(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            private = create_fixture(root)
            output = root / "raced-output"
            original = EXPORTER.atomic_rename_no_replace

            def create_competing_target(stage: Path, target: Path) -> None:
                target.mkdir()
                (target / "winner.txt").write_text("winner", encoding="utf-8")
                original(stage, target)

            with mock.patch.object(
                EXPORTER,
                "atomic_rename_no_replace",
                side_effect=create_competing_target,
            ), self.assertRaises(FileExistsError):
                EXPORTER.export_snapshot(
                    private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=output,
                )
            self.assertEqual((output / "winner.txt").read_text(encoding="utf-8"), "winner")
            self.assertFalse(list(root.glob(".raced-output.staging-*")))

    def test_atomic_publish_fails_on_unsupported_platform(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            stage = root / "stage"
            stage.mkdir()
            target = root / "target"
            with mock.patch.object(EXPORTER.sys, "platform", "unsupported"):
                with self.assertRaisesRegex(RuntimeError, "unavailable"):
                    EXPORTER.atomic_rename_no_replace(stage, target)
            self.assertTrue(stage.is_dir())
            self.assertFalse(target.exists())

    def test_attempt_and_inventory_rejections(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            private = create_fixture(root)

            def score_judge_error(rows: list[dict[str, str]]) -> None:
                row = next(item for item in rows if item["frozen_evaluation_status"] == "JUDGE_ERROR")
                row["frozen_normalized_score"] = "0"
                row["frozen_pass"] = "false"

            rewrite_csv(private, "attempts.csv", EXPORTER.ATTEMPT_FIELDS, score_judge_error)
            with self.assertRaisesRegex(ValueError, "coerces JUDGE_ERROR"):
                EXPORTER.export_snapshot(
                    private,
                    source_controller_head="a" * 40,
                    expected_attempt_records=3,
                    evidence_cutoff="2026-09-04T05:10:00+00:00",
                    output_dir=root / "judge_error_score",
                )

        mutations = {
            "missing system": (
                "system_summary.csv",
                EXPORTER.SYSTEM_FIELDS,
                lambda rows: rows.pop(),
            ),
            "missing task": (
                "task_system_summary.csv",
                EXPORTER.TASK_SYSTEM_FIELDS,
                lambda rows: rows.pop(),
            ),
            "diagnostic ranking": (
                "system_summary.csv",
                EXPORTER.SYSTEM_FIELDS,
                lambda rows: next(
                    row for row in rows if row["cohort"] == "diagnostic_only"
                ).__setitem__("ranking_eligible", "true"),
            ),
            "historical ranking": (
                "system_summary.csv",
                EXPORTER.SYSTEM_FIELDS,
                lambda rows: next(
                    row for row in rows if row["cohort"] == "historical_baseline"
                ).__setitem__("ranking_eligible", "true"),
            ),
        }
        for name, (filename, fields, mutate) in mutations.items():
            with self.subTest(name=name), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary).resolve()
                private = create_fixture(root)
                rewrite_csv(private, filename, fields, mutate)
                with self.assertRaises(ValueError):
                    EXPORTER.export_snapshot(
                        private,
                        source_controller_head="a" * 40,
                        expected_attempt_records=3,
                        evidence_cutoff="2026-09-04T05:10:00+00:00",
                        output_dir=root / "rejected",
                    )
                self.assertFalse((root / "rejected").exists())


if __name__ == "__main__":
    unittest.main()
