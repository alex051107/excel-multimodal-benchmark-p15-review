#!/usr/bin/env python3
"""Render the public-facing figures for the Judge V3 initial report."""

from __future__ import annotations

import csv
from collections import defaultdict
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager
from matplotlib.colors import LinearSegmentedColormap
from matplotlib.patches import Patch


ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / "results" / "judge_v3_initial"
OUT = ROOT / "figures" / "judge_v3_initial"

TASK_LABELS = {
    "P15-A-ENG-SIZING-001": "A  工程选泵",
    "P15-A-FIN-DCF-001": "A  现金流估值",
    "P15-A-FIN-DEBUG-001": "A  财务纠错",
    "P15-A-POLICY-EIA-001": "A  政策情景",
    "P15-A-STAT-EXPERIMENT-001": "A  配对实验",
    "P15-B-FIN-RECON-001": "B  财务对账",
    "P15-B-HEALTH-REPORT-001": "B  健康数据报告",
    "P15-B-OPS-CLEAN-JOIN-001": "B  订单清洗合并",
    "P15-B-PUBLIC-PIVOT-001": "B  原生透视表",
    "P15-B-SALES-DISCOVERY-001": "B  销售数据选择",
    "P15-C-INVOICE-001": "C  发票整理",
    "P15-C-PO-ADDENDUM-001": "C  采购变更单",
    "P15-C-QUOTE-001": "C  报价单整理",
    "P15-C-RECEIPTS-001": "C  多张票据整理",
    "P15-C-STATEMENT-001": "C  银行账单整理",
}

SYSTEM_LABELS = {
    "claude_opus5": "Claude Opus 5",
    "codex_gpt56sol": "Codex GPT-5.6 Sol",
    "qwen38max": "Qwen3.8-max",
}

BLUE = "#2878B5"
LIGHT_BLUE = "#8EC0DD"
ORANGE = "#E1812C"
GOLD = "#D9A441"
GRAY = "#A0A7AE"
DARK = "#24313A"
GRID = "#D9DEE3"
BAR = "#657F98"
HIGHLIGHT = "#B66A50"
DIFFICULTY_CMAP = LinearSegmentedColormap.from_list(
    "p15_difficulty",
    ["#5B3F78", "#8D6BA3", "#C6B4D2", "#E6DDE9", "#F4F0ED"],
)


def configure_style() -> None:
    candidates = [
        "PingFang SC",
        "Hiragino Sans GB",
        "Arial Unicode MS",
        "Noto Sans CJK SC",
        "Microsoft YaHei",
    ]
    installed = {font.name for font in font_manager.fontManager.ttflist}
    chosen = next((name for name in candidates if name in installed), "DejaVu Sans")
    plt.rcParams.update(
        {
            "font.family": "sans-serif",
            "font.sans-serif": [chosen, "DejaVu Sans"],
            "axes.unicode_minus": False,
            "figure.dpi": 150,
            "savefig.dpi": 180,
            "font.size": 11,
            "axes.titlesize": 17,
            "axes.titleweight": "bold",
            "axes.labelsize": 11,
            "xtick.labelsize": 10,
            "ytick.labelsize": 10,
        }
    )


def read_csv(name: str) -> list[dict[str, str]]:
    with (DATA / name).open(encoding="utf-8-sig", newline="") as handle:
        return list(csv.DictReader(handle))


def write_csv(name: str, rows: list[dict[str, object]], fields: list[str]) -> None:
    with (DATA / name).open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fields, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


def write_score_only_tables() -> None:
    rows = read_csv("judge_v3_candidate_full_archive_replay.csv")
    current = [
        row
        for row in rows
        if row["valid_attempt"] == "True" and row["replay_scoreable"] == "True"
    ]
    attempt_rows = [
        {
            "wave_id": row["wave_id"],
            "task_id": row["task_id"],
            "system": SYSTEM_LABELS[row["system"]],
            "score": f'{float(row["replay_score"]):.6f}',
        }
        for row in current
    ]
    write_csv(
        "v3_scores_current.csv",
        attempt_rows,
        ["wave_id", "task_id", "system", "score"],
    )

    grouped: dict[tuple[str, str], list[float]] = defaultdict(list)
    for row in current:
        grouped[(row["task_id"], row["system"])].append(float(row["replay_score"]))
    summary_rows: list[dict[str, object]] = []
    for task_id in TASK_LABELS:
        for system_id in SYSTEM_LABELS:
            scores = grouped.get((task_id, system_id), [])
            summary_rows.append(
                {
                    "task_id": task_id,
                    "task": TASK_LABELS[task_id].replace("  ", " "),
                    "system": SYSTEM_LABELS[system_id],
                    "n_in_mean": len(scores),
                    "mean_score": f"{sum(scores) / len(scores):.6f}" if scores else "",
                }
            )
    write_csv(
        "v3_scores_by_task_system.csv",
        summary_rows,
        ["task_id", "task", "system", "n_in_mean", "mean_score"],
    )


def clean_axes(ax: plt.Axes) -> None:
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.tick_params(axis="y", length=0)
    ax.grid(axis="x", color=GRID, linewidth=0.8, alpha=0.7)
    ax.set_axisbelow(True)


def render_replay_status() -> None:
    labels = ["新版规则已经算出分数", "评分程序暂时读不懂", "旧题版本不能继续评分", "还要用 Excel 打开重算"]
    values = [322, 12, 24, 10]
    colors = [BLUE, GOLD, GRAY, LIGHT_BLUE]

    fig, ax = plt.subplots(figsize=(11.5, 4.8))
    y = list(range(len(labels)))
    bars = ax.barh(y, values, height=0.58, color=colors)
    ax.set_yticks(y, labels)
    ax.invert_yaxis()
    ax.set_xlim(0, 368)
    for bar, value in zip(bars, values):
        ax.text(value + 4, bar.get_y() + bar.get_height() / 2, f"{value} 份", va="center", ha="left", color=DARK, fontweight="bold")
    ax.set_xlabel("现有 Excel 结果（共 368 份）")
    clean_axes(ax)
    fig.suptitle("新版评分规则已经给 322 份结果算出了分数", x=0.125, y=0.98, ha="left", fontsize=17, fontweight="bold")
    fig.tight_layout(rect=[0, 0, 1, 0.88])
    fig.savefig(OUT / "01_replay_status.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)


def render_task_results() -> None:
    rows = read_csv("judge_v3_candidate_task_audit_summary.csv")
    labels: list[str] = []
    means: list[float] = []
    annotations: list[str] = []
    colors: list[str] = []

    for row in rows:
        task_id = row["task_id"]
        labels.append(TASK_LABELS[task_id])
        n = int(row["valid_replay_scoreable_n"])
        if n:
            mean = float(row["valid_replay_mean"])
            annotations.append(f"平均 {mean:.2f}；{n} 份结果")
        else:
            mean = float("nan")
            annotations.append("新版题目还没重跑（旧版来源和单位有冲突）")
        means.append(mean)
        if task_id == "P15-A-POLICY-EIA-001":
            colors.append(GRAY)
        elif task_id == "P15-B-FIN-RECON-001":
            colors.append(HIGHLIGHT)
        else:
            colors.append(BAR)

    fig, ax = plt.subplots(figsize=(12.5, 9.2))
    y = list(range(len(labels)))
    widths = [0.0 if mean != mean else mean for mean in means]
    bars = ax.barh(y, widths, color=colors, height=0.62)
    ax.set_yticks(y, labels)
    ax.invert_yaxis()
    ax.set_xlim(0, 1.28)
    ax.set_xlabel("本轮已经算出分数的结果，平均分（0 至 1，分数越低题目越难）")
    ax.set_title("V3 初步重评：财务对账的平均分最低", loc="left", pad=18)

    for bar, text_label in zip(bars, annotations):
        x = max(bar.get_width() + 0.025, 0.025)
        ax.text(x, bar.get_y() + bar.get_height() / 2, text_label, va="center", ha="left", fontsize=9.5, color=DARK)

    clean_axes(ax)
    fig.tight_layout()
    fig.savefig(OUT / "02_task_results.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)


def render_old_zero_recheck() -> None:
    replay = read_csv("judge_v3_candidate_full_archive_replay.csv")
    old_zero = [
        row
        for row in replay
        if row["original_score"] == "0.0"
    ]
    groups = [("全部旧 0 分", old_zero), ("其中：本轮正常完成", [row for row in old_zero if row["valid_attempt"] == "True"])]

    series = [
        ("新版改为非零分", BLUE),
        ("评分程序还读不懂", GOLD),
        ("旧题版本不可比", ORANGE),
        ("新版仍为 0 分", GRAY),
    ]
    counts: list[list[int]] = []
    for _, rows in groups:
        nonzero = sum(row["replay_scoreable"] == "True" and float(row["replay_score"]) > 0 for row in rows)
        zero = sum(row["replay_scoreable"] == "True" and float(row["replay_score"]) == 0 for row in rows)
        task_version = sum(row["evaluator_status"] == "TASK_INVALID" for row in rows)
        unable = len(rows) - nonzero - zero - task_version
        counts.append([nonzero, unable, task_version, zero])

    fig, ax = plt.subplots(figsize=(10.8, 4.8))
    for y, ((group_label, rows), values) in enumerate(zip(groups, counts)):
        left = 0
        for (series_label, color), value in zip(series, values):
            ax.barh(y, value, left=left, height=0.5, color=color, edgecolor="white")
            if value:
                ax.text(left + value / 2, y, f"{value}", ha="center", va="center", color="white" if color != GRAY else DARK, fontweight="bold")
            left += value
        ax.text(left + 0.8, y, f"共 {len(rows)} 份", va="center", ha="left", color=DARK)

    ax.set_yticks(range(len(groups)), [group[0] for group in groups])
    ax.invert_yaxis()
    ax.set_xlim(0, max(len(rows) for _, rows in groups) * 1.08)
    ax.set_xlabel("")
    ax.legend(
        handles=[Patch(facecolor=color, label=label) for label, color in series],
        loc="lower center",
        bbox_to_anchor=(0.5, -0.42),
        ncol=4,
        frameon=False,
    )
    clean_axes(ax)
    fig.suptitle("旧评分规则留下的 98 份 0 分中，只有 3 份仍是 0 分", x=0.125, y=0.98, ha="left", fontsize=17, fontweight="bold")
    fig.tight_layout(rect=[0, 0.10, 1, 0.88])
    fig.savefig(OUT / "03_old_zero_recheck.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)


def render_old_task_results() -> None:
    rows = read_csv("old_evaluator_task_results.csv")
    labels = [TASK_LABELS[row["task_id"]] for row in rows]
    totals = [int(row["scoreable_n"]) for row in rows]
    means = [float(row["mean_score"]) for row in rows]

    fig, ax = plt.subplots(figsize=(12.5, 9.2))
    y = list(range(len(rows)))
    bars = ax.barh(y, means, height=0.62, color=BAR)
    ax.set_yticks(y, labels)
    ax.invert_yaxis()
    ax.set_xlim(0, 1.22)
    ax.set_xlabel("旧评分规则给出的平均分（0 至 1）")
    ax.set_title("旧结果中的大量低分，首先暴露了评分规则的问题", loc="left", pad=18)
    for bar, mean, n in zip(bars, means, totals):
        ax.text(max(bar.get_width() + 0.02, 0.02), bar.get_y() + bar.get_height() / 2, f"平均 {mean:.2f}；n={n}", va="center", ha="left", fontsize=9.5, color=DARK)
    clean_axes(ax)
    fig.tight_layout()
    fig.savefig(OUT / "04_old_task_results.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)


def render_task_system_heatmap() -> None:
    rows = read_csv("judge_v3_candidate_full_archive_replay.csv")
    grouped: dict[tuple[str, str], list[dict[str, str]]] = defaultdict(list)
    for row in rows:
        if row["valid_attempt"] == "True" and row["replay_scoreable"] == "True":
            grouped[(row["task_id"], row["system"])].append(row)

    task_ids = list(TASK_LABELS)
    system_ids = list(SYSTEM_LABELS)
    means = np.full((len(task_ids), len(system_ids)), np.nan)
    annotations: list[list[str]] = [["" for _ in system_ids] for _ in task_ids]

    for i, task_id in enumerate(task_ids):
        for j, system_id in enumerate(system_ids):
            cell = grouped.get((task_id, system_id), [])
            if not cell:
                annotations[i][j] = "新版题目\n还没重跑"
                continue
            scores = [float(row["replay_score"]) for row in cell]
            means[i, j] = sum(scores) / len(scores)
            annotations[i][j] = f"平均 {means[i, j]:.2f}\nn={len(cell)}"

    cmap = DIFFICULTY_CMAP.copy()
    cmap.set_bad("#E5E8EB")
    fig, ax = plt.subplots(figsize=(11.5, 10.2))
    image = ax.imshow(np.ma.masked_invalid(means), cmap=cmap, vmin=0, vmax=1, aspect="auto")

    ax.set_xticks(range(len(system_ids)), [SYSTEM_LABELS[item] for item in system_ids])
    ax.set_yticks(range(len(task_ids)), [TASK_LABELS[item] for item in task_ids])
    ax.set_xlabel("")
    ax.set_ylabel("")
    ax.set_title("V3 初步重评：同一道题在三套系统中的分数差异明显", loc="left", pad=18)
    ax.tick_params(axis="both", length=0)

    ax.set_xticks(np.arange(-0.5, len(system_ids), 1), minor=True)
    ax.set_yticks(np.arange(-0.5, len(task_ids), 1), minor=True)
    ax.grid(which="minor", color="white", linewidth=2)
    ax.tick_params(which="minor", bottom=False, left=False)
    for spine in ax.spines.values():
        spine.set_visible(False)

    for i in range(len(task_ids)):
        for j in range(len(system_ids)):
            value = means[i, j]
            color = "white" if not np.isnan(value) and value <= 0.55 else DARK
            ax.text(j, i, annotations[i][j], ha="center", va="center", color=color, fontsize=10, fontweight="bold")

    colorbar = fig.colorbar(image, ax=ax, fraction=0.035, pad=0.035)
    colorbar.set_label("平均分（0 至 1，深色表示分数更低）")
    colorbar.outline.set_visible(False)

    fig.tight_layout()
    fig.savefig(OUT / "05_task_system_heatmap.png", bbox_inches="tight", facecolor="white")
    plt.close(fig)


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    configure_style()
    write_score_only_tables()
    render_replay_status()
    render_task_results()
    render_old_zero_recheck()
    render_old_task_results()
    render_task_system_heatmap()


if __name__ == "__main__":
    main()
