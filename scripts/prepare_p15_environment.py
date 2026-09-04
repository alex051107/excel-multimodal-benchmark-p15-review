#!/usr/bin/env python3
"""Build the shared P15 Agent image once, then reuse it unchanged.

The default behavior is idempotent: if the configured image exists and its
labels and installed versions match, the script exits without running a build.
It never overwrites a conflicting image tag. A changed environment must use a
new image tag and image_contract in benchmark/environment.json.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
ENVIRONMENT_CONFIG = ROOT / "benchmark" / "environment.json"
DOCKERFILE = ROOT / "benchmark" / "images" / "Dockerfile.agent-base"
BUILD_CONTEXT = DOCKERFILE.parent
DEFAULT_RECEIPT = ROOT / "benchmark_runs" / "environment_receipt.json"


def run(command: list[str], timeout: int = 120) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        cwd=ROOT,
        text=True,
        capture_output=True,
        check=False,
        timeout=timeout,
    )


def read_config() -> dict[str, Any]:
    payload = json.loads(ENVIRONMENT_CONFIG.read_text(encoding="utf-8"))
    if not isinstance(payload, dict):
        raise RuntimeError("benchmark/environment.json must contain one JSON object")
    return payload


def require_docker() -> str:
    docker = shutil.which("docker")
    if docker is None:
        raise RuntimeError("Docker is not installed or is not on PATH")
    version = run([docker, "version", "--format", "{{.Server.Version}}"], timeout=30)
    if version.returncode != 0 or not version.stdout.strip():
        raise RuntimeError("Docker Engine is not available: " + version.stderr.strip())
    return docker


def inspect_image(docker: str, image: str) -> dict[str, Any] | None:
    completed = run([docker, "image", "inspect", image], timeout=30)
    if completed.returncode != 0:
        return None
    rows = json.loads(completed.stdout)
    return rows[0] if rows else None


def probe_versions(docker: str, image: str) -> dict[str, str]:
    script = (
        "python --version 2>&1; "
        "node --version; "
        "python -c 'import openpyxl; print(openpyxl.__version__)'; "
        "codex --version; claude --version; qwen --version"
    )
    completed = run([docker, "run", "--rm", image, "sh", "-lc", script], timeout=60)
    if completed.returncode != 0:
        raise RuntimeError("Agent image version probe failed: " + completed.stderr[-2000:])
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    if len(lines) != 6:
        raise RuntimeError(f"Expected six version lines, received: {lines!r}")
    return {
        "python": lines[0].removeprefix("Python "),
        "node": lines[1].removeprefix("v"),
        "openpyxl": lines[2],
        "codex": lines[3].removeprefix("codex-cli "),
        "claude-code": lines[4].split()[0],
        "qwen-coder": lines[5],
    }


def validate_image(config: dict[str, Any], image_info: dict[str, Any], versions: dict[str, str]) -> None:
    labels = (image_info.get("Config") or {}).get("Labels") or {}
    expected = {
        "org.opencontainers.image.version": config["image_contract"],
        "p15.harbor.version": config["harbor_version"],
        "p15.openpyxl.version": config["openpyxl_version"],
        "p15.codex.version": config["agents"]["codex"]["version"],
        "p15.claude-code.version": config["agents"]["claude-code"]["version"],
        "p15.qwen-code.version": config["agents"]["qwen-coder"]["version"],
    }
    wrong_labels = {
        key: {"expected": value, "actual": labels.get(key)}
        for key, value in expected.items()
        if labels.get(key) != value
    }
    expected_versions = {
        "openpyxl": config["openpyxl_version"],
        "codex": config["agents"]["codex"]["version"],
        "claude-code": config["agents"]["claude-code"]["version"],
        "qwen-coder": config["agents"]["qwen-coder"]["version"],
    }
    wrong_versions = {
        key: {"expected": value, "actual": versions.get(key)}
        for key, value in expected_versions.items()
        if versions.get(key) != value
    }
    if not versions["python"].startswith(config["python_version_prefix"]):
        wrong_versions["python"] = {
            "expected": config["python_version_prefix"] + "x",
            "actual": versions["python"],
        }
    if versions["node"] != config["node_version"]:
        wrong_versions["node"] = {
            "expected": config["node_version"],
            "actual": versions["node"],
        }
    if wrong_labels or wrong_versions:
        raise RuntimeError(
            "The configured image tag exists but does not match the frozen contract. "
            "Do not overwrite it; bump base_image and image_contract instead.\n"
            + json.dumps(
                {"wrong_labels": wrong_labels, "wrong_versions": wrong_versions},
                ensure_ascii=False,
                indent=2,
            )
        )


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--check",
        action="store_true",
        help="Check only. Fail instead of building when the image is absent.",
    )
    parser.add_argument("--receipt", type=Path, default=DEFAULT_RECEIPT)
    args = parser.parse_args()

    config = read_config()
    docker = require_docker()
    image = str(config["base_image"])
    image_info = inspect_image(docker, image)
    action = "reused"
    if image_info is None:
        if args.check:
            raise RuntimeError(
                f"Required image {image!r} is absent. Run: "
                "python scripts/prepare_p15_environment.py"
            )
        built = subprocess.run(
            [
                docker,
                "build",
                "--file",
                str(DOCKERFILE),
                "--tag",
                image,
                str(BUILD_CONTEXT),
            ],
            cwd=ROOT,
            check=False,
        )
        if built.returncode != 0:
            raise SystemExit(built.returncode)
        image_info = inspect_image(docker, image)
        if image_info is None:
            raise RuntimeError("Docker reported success but the prepared image is absent")
        action = "built"

    versions = probe_versions(docker, image)
    validate_image(config, image_info, versions)
    receipt = {
        "checked_at": datetime.now(timezone.utc).isoformat(),
        "action": action,
        "image": image,
        "image_id": image_info["Id"],
        "repo_digests": image_info.get("RepoDigests") or [],
        "versions": versions,
        "contract_valid": True,
    }
    args.receipt.parent.mkdir(parents=True, exist_ok=True)
    args.receipt.write_text(
        json.dumps(receipt, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    print(json.dumps(receipt, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
