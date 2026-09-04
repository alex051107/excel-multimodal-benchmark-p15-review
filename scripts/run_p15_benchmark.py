#!/usr/bin/env python3
"""Create and optionally execute a bounded Harbor job for P15.

Dry-run is the default. It verifies the selected task packages and asks Harbor
to parse the generated JobConfig. ``--install-only`` checks Docker and Agent
setup without calling a model. ``--execute`` is required before any online
Agent is called.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shutil
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CONFIG = ROOT / "benchmark" / "configs" / "p15_v3_n8.json"
ENVIRONMENT_PREPARER = ROOT / "scripts" / "prepare_p15_environment.py"
HARBOR_COMPAT_DIR = ROOT / "benchmark" / "harbor_compat"
EXPECTED_HARBOR_VERSION = "0.22.0"
SYSTEM_TO_AGENT = {
    "codex_gpt56sol": "codex",
    "claude_opus5": "claude-code",
    "qwen38max": "qwen-coder",
}
REQUIRED_TASK_FILES = (
    "task.toml",
    "instruction.md",
    "rubric.json",
    "environment/Dockerfile",
    "tests/Dockerfile",
    "tests/evaluate.py",
    "tests/test.sh",
)
ENV_TEMPLATE = re.compile(r"^\$\{([A-Za-z_][A-Za-z0-9_]*)\}$")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Prepare one bounded P15 Harbor job; dry-run unless --execute is set."
    )
    parser.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    parser.add_argument("--system", choices=sorted(SYSTEM_TO_AGENT), required=True)
    scope = parser.add_mutually_exclusive_group(required=True)
    scope.add_argument("--task-id", action="append", help="May be supplied more than once.")
    scope.add_argument("--all-tasks", action="store_true")
    parser.add_argument("--attempts", type=int, default=1)
    parser.add_argument("--workers", type=int, default=1)
    parser.add_argument("--jobs-dir", type=Path, default=ROOT / "benchmark_runs")
    parser.add_argument("--harbor-bin", default="harbor")
    parser.add_argument(
        "--execute",
        action="store_true",
        help="Actually call the online Agent. Without this flag only validation runs.",
    )
    parser.add_argument(
        "--install-only",
        action="store_true",
        help="Build the task environment and verify Agent setup without calling a model.",
    )
    return parser.parse_args()


def read_json(path: Path) -> dict[str, Any]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise RuntimeError(f"配置必须是 JSON object: {path}")
    return payload


def check_harbor(harbor_bin: str) -> str:
    executable = shutil.which(harbor_bin)
    if executable is None:
        raise RuntimeError(
            "找不到 Harbor。请先安装 0.22.0：uv tool install harbor==0.22.0"
        )
    completed = subprocess.run(
        [executable, "--version"],
        text=True,
        capture_output=True,
        check=False,
        timeout=30,
    )
    version = completed.stdout.strip()
    if completed.returncode != 0 or version != EXPECTED_HARBOR_VERSION:
        raise RuntimeError(
            f"需要 Harbor {EXPECTED_HARBOR_VERSION}，当前得到 {version or 'unknown'}"
        )
    return executable


def harbor_process_env() -> dict[str, str]:
    env = os.environ.copy()
    existing = env.get("PYTHONPATH")
    env["PYTHONPATH"] = (
        str(HARBOR_COMPAT_DIR)
        if not existing
        else str(HARBOR_COMPAT_DIR) + os.pathsep + existing
    )
    return env


def harbor_python(harbor_bin: str) -> str:
    first_line = Path(harbor_bin).resolve().read_text(
        encoding="utf-8", errors="replace"
    ).splitlines()[0]
    if first_line.startswith("#!"):
        candidate = first_line[2:].strip()
        if Path(candidate).is_file():
            return candidate
    return sys.executable


def check_qwen_compat(harbor_bin: str, env: dict[str, str]) -> None:
    probe = subprocess.run(
        [
            harbor_python(harbor_bin),
            "-c",
            (
                "from harbor.agents.installed.qwen_code import QwenCode; "
                "assert getattr(QwenCode, '_p15_preinstalled_version_check', False)"
            ),
        ],
        cwd=ROOT,
        env=env,
        text=True,
        capture_output=True,
        check=False,
        timeout=30,
    )
    if probe.returncode != 0:
        raise RuntimeError(
            "Harbor Qwen compatibility shim was not loaded: " + probe.stderr[-2000:]
        )


def check_prepared_environment() -> dict[str, Any]:
    completed = subprocess.run(
        [sys.executable, str(ENVIRONMENT_PREPARER), "--check"],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
        timeout=120,
    )
    if completed.returncode != 0:
        raise RuntimeError(
            "The frozen P15 Agent image is not ready. Run "
            "`python scripts/prepare_p15_environment.py` once.\n"
            + completed.stderr[-2000:]
        )
    return json.loads(completed.stdout)


def validate_task(task_id: str, task_path: Path) -> dict[str, Any]:
    if task_path.name != task_id:
        raise RuntimeError(f"任务编号与目录不一致: {task_id} / {task_path}")
    missing = [relative for relative in REQUIRED_TASK_FILES if not (task_path / relative).is_file()]
    input_dir = task_path / "data" / "input_files"
    inputs = sorted(path for path in input_dir.glob("**/*") if path.is_file())
    if not inputs:
        missing.append("data/input_files/*")
    if missing:
        raise RuntimeError(f"{task_id} 缺少文件: {', '.join(missing)}")

    rubric = read_json(task_path / "rubric.json")
    if rubric.get("task_id") != task_id or not rubric.get("criteria"):
        raise RuntimeError(f"{task_id} 的 rubric.json 不完整")
    return {
        "task_id": task_id,
        "input_files": len(inputs),
        "criteria": len(rubric["criteria"]),
    }


def referenced_env(agent: dict[str, Any]) -> list[str]:
    variables: list[str] = []
    for value in (agent.get("env") or {}).values():
        match = ENV_TEMPLATE.fullmatch(str(value))
        if match:
            variables.append(match.group(1))
    return variables


def build_job(args: argparse.Namespace, base: dict[str, Any]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    if args.execute and args.install_only:
        raise RuntimeError("--execute and --install-only cannot be used together")
    if not 1 <= args.attempts <= 8:
        raise RuntimeError("--attempts 必须介于 1 和 8")
    if not 1 <= args.workers <= 3:
        raise RuntimeError("--workers 必须介于 1 和 3")

    tasks_by_id = {
        Path(row["path"]).name: row
        for row in base.get("tasks", [])
        if isinstance(row, dict) and isinstance(row.get("path"), str)
    }
    selected_ids = list(tasks_by_id) if args.all_tasks else list(dict.fromkeys(args.task_id or []))
    unknown = sorted(set(selected_ids) - set(tasks_by_id))
    if unknown:
        raise RuntimeError(f"配置中没有这些任务: {', '.join(unknown)}")

    checks = []
    selected_tasks = []
    for task_id in selected_ids:
        row = tasks_by_id[task_id]
        task_path = (ROOT / row["path"]).resolve()
        try:
            task_path.relative_to((ROOT / "tasks" / "pilot_v1").resolve())
        except ValueError as exc:
            raise RuntimeError(f"任务目录离开 tasks/pilot_v1: {task_path}") from exc
        checks.append(validate_task(task_id, task_path))
        selected_tasks.append(row)

    target_agent = SYSTEM_TO_AGENT[args.system]
    selected_agents = [row for row in base.get("agents", []) if row.get("name") == target_agent]
    if len(selected_agents) != 1:
        raise RuntimeError(f"系统 {args.system} 没有唯一的 Agent 配置")

    timestamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    job = {
        "job_name": f"p15_{args.system}_{len(selected_tasks)}task_{args.attempts}run_{timestamp}",
        "jobs_dir": str(args.jobs_dir.resolve()),
        "n_attempts": args.attempts,
        "n_concurrent_trials": args.workers,
        "retry": {"max_retries": 0},
        "agents": selected_agents,
        "tasks": selected_tasks,
    }
    if args.install_only:
        job["install_only"] = True
    return job, checks


def main() -> None:
    args = parse_args()
    harbor_bin = check_harbor(args.harbor_bin)
    process_env = harbor_process_env()
    check_qwen_compat(harbor_bin, process_env)
    base = read_json(args.config.resolve())
    job, checks = build_job(args, base)

    config_dir = args.jobs_dir.resolve() / "generated_configs"
    config_dir.mkdir(parents=True, exist_ok=True)
    config_path = config_dir / f"{job['job_name']}.json"
    config_path.write_text(
        json.dumps(job, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )

    parsed = subprocess.run(
        [harbor_bin, "run", "--config", str(config_path), "--print-config"],
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
        timeout=60,
        env=process_env,
    )
    if parsed.returncode != 0:
        raise RuntimeError(f"Harbor 配置检查失败: {parsed.stderr[-2000:]}")
    resolved = json.loads(parsed.stdout)
    if len(resolved.get("tasks", [])) != len(job["tasks"]):
        raise RuntimeError("Harbor 解析后的任务数量发生变化")
    if len(resolved.get("agents", [])) != 1:
        raise RuntimeError("Harbor 解析后的 Agent 数量不是 1")

    environment_receipt = None
    if args.execute or args.install_only:
        environment_receipt = check_prepared_environment()

    mode = "execute" if args.execute else "install-only" if args.install_only else "dry-run"
    summary = {
        "mode": mode,
        "harbor_version": EXPECTED_HARBOR_VERSION,
        "system": args.system,
        "model": job["agents"][0]["model_name"],
        "task_count": len(job["tasks"]),
        "attempts_per_task": args.attempts,
        "planned_trials": len(job["tasks"]) * args.attempts,
        "workers": args.workers,
        "task_package_checks": checks,
        "generated_config": str(config_path),
        "harbor_config_valid": True,
        "qwen_preinstalled_version_check": True,
    }
    if environment_receipt is not None:
        summary["environment"] = {
            "image": environment_receipt["image"],
            "image_id": environment_receipt["image_id"],
            "versions": environment_receipt["versions"],
            "contract_valid": environment_receipt["contract_valid"],
        }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if not args.execute and not args.install_only:
        return

    if args.execute:
        missing_env = [
            name for name in referenced_env(job["agents"][0]) if not os.environ.get(name)
        ]
        if missing_env:
            raise RuntimeError(
                "执行前需要设置这些环境变量: "
                + ", ".join(sorted(missing_env))
            )
    else:
        for variable in referenced_env(job["agents"][0]):
            process_env.setdefault(variable, "install-only-not-used")
    completed = subprocess.run(
        [harbor_bin, "run", "--config", str(config_path), "--yes"],
        cwd=ROOT,
        check=False,
        env=process_env,
    )
    raise SystemExit(completed.returncode)


if __name__ == "__main__":
    main()
