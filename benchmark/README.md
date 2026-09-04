# 用 Harbor 运行 P15 Benchmark

P15 按 Harbor 0.22.0 的原生方式封装。每道题是一个独立任务包，整套实验由一个 Job 配置组合起来。

这些文件已经实际跑过。结果记在 [`VALIDATION_RECEIPT.json`](VALIDATION_RECEIPT.json)，包括共享镜像的二次复用、15 题环境构建和三个 Agent setup。我们还用未修改的 Harbor 0.22.0 检查了 Qwen，跑通了 Oracle 生成答卷到 Evaluator 评分的完整流程，并重新复算了 368 份答卷。

## 封装层级

| 层级 | 位置 | 作用 |
|---|---|---|
| 单题任务包 | [`tasks/pilot_v1/<task_id>/`](../tasks/pilot_v1/) | 题面、输入、环境、交付要求和评分程序 |
| Harbor Job | [`configs/p15_v3_n8.json`](configs/p15_v3_n8.json) | 15 道题、3 套系统、每组 8 次、串行执行、失败不自动重试 |
| 固定环境 | [`environment.json`](environment.json) 和 [`images/`](images/) | 固定 Harbor、Python、openpyxl 和三个 Agent CLI 的版本 |
| 安全入口 | [`../scripts/run_p15_benchmark.py`](../scripts/run_p15_benchmark.py) | 选择题目和系统，检查任务包，生成单次 Job；默认只检查，不调用模型 |
| 固定答卷复算 | [`../reproduction/`](../reproduction/) | 不调用模型，重新计算当前报告使用的 V3 分数 |

单题任务包采用 Harbor 的常见结构：

```text
P15-.../
├── task.toml
├── instruction.md
├── data/input_files/
├── environment/Dockerfile
├── rubric.json
├── tests/Dockerfile
├── tests/evaluate.py
└── tests/test.sh
```

Agent 必须把最终工作簿写到 `/app/output/answer.xlsx`。Harbor 会把该文件作为 artifact 保存，再在独立 verifier 环境中运行 `tests/test.sh` 和 `tests/evaluate.py`。

## 环境

- Harbor 0.22.0
- Docker Desktop 或兼容的 Docker Engine
- Python 3.11 或更高版本
- 对应在线模型的 API 凭证和服务地址

安装 Harbor：

```bash
uv tool install harbor==0.22.0
harbor --version
```

## 先准备一次共享环境

15 道题共用一个 Docker 基础镜像。第一次会下载并安装固定版本的 Codex、Claude Code、Qwen Code 和 openpyxl：

```bash
python scripts/prepare_p15_environment.py
```

脚本先检查镜像标签和实际版本。如果已经一致，它会直接显示 `"action": "reused"`，不会再执行 `docker build`，也不会重复下载 Agent。如果同名镜像存在但版本不对，脚本会停止，不会覆盖它。

当次检查到的 Docker image ID 和 CLI 版本会写入本地的 `benchmark_runs/environment_receipt.json`。该目录不上传 GitHub，避免把本机运行状态混入公开任务包。

## 先检查一题，不调用模型

下面的命令只检查任务包并让 Harbor 解析 Job 配置，不会产生在线调用：

```bash
python scripts/run_p15_benchmark.py \
  --system codex_gpt56sol \
  --task-id P15-A-FIN-DCF-001
```

成功时会显示 `"mode": "dry-run"` 和 `"harbor_config_valid": true`。

## 真正进入 Harbor 环境，但不调用模型

`--install-only` 会让 Harbor 构建单题环境并执行 Agent setup，然后结束。它用来确认 Dockerfile、共享镜像和 Agent adapter 能够一起工作，不会向在线模型发送题目。

```bash
python scripts/run_p15_benchmark.py \
  --system codex_gpt56sol \
  --task-id P15-A-FIN-DCF-001 \
  --install-only
```

Harbor 0.22.0 会识别已安装的 Codex 和 Claude Code。公开入口通过 [`harbor_compat/sitecustomize.py`](harbor_compat/sitecustomize.py) 给 Qwen Code 补上同样的精确版本检查。这个文件只在 P15 入口启动 Harbor 时生效，不会修改学长本机安装的 Harbor。

## 实际运行一题一次

先在当前 shell 设置凭证和服务地址。脚本只在运行时读取这些变量，不把真实值写进仓库：

```bash
export P15_OPENAI_API_KEY="..."
export P15_OPENAI_BASE_URL="https://your-openai-compatible-endpoint/v1"

python scripts/run_p15_benchmark.py \
  --system codex_gpt56sol \
  --task-id P15-A-FIN-DCF-001 \
  --attempts 1 \
  --execute
```

另外两套系统使用：

| 系统参数 | 环境变量 |
|---|---|
| `claude_opus5` | `P15_ANTHROPIC_API_KEY`、`P15_ANTHROPIC_BASE_URL` |
| `qwen38max` | `P15_QWEN_API_KEY`、`P15_QWEN_BASE_URL` |

输出保存在 `benchmark_runs/<job_name>/`。每次 Job 的实际配置也会放在 `benchmark_runs/generated_configs/`，便于确认当次用了哪道题、哪个系统和多少次重复。

## 运行完整的 15 × 3 × 8 实验

[`configs/p15_v3_n8.json`](configs/p15_v3_n8.json) 是完整 Harbor Job：15 道题 × 3 套系统 × 每组 8 次，共 360 次计划运行。实际执行建议分三条命令走公开入口，这样每条线都会先检查共享镜像和相应凭证：

```bash
python scripts/run_p15_benchmark.py --system codex_gpt56sol --all-tasks --attempts 8 --execute
python scripts/run_p15_benchmark.py --system claude_opus5 --all-tasks --attempts 8 --execute
python scripts/run_p15_benchmark.py --system qwen38max --all-tasks --attempts 8 --execute
```

这些命令会产生真实在线调用。正式重跑前应先分别用顶层脚本跑 1 题 × 1 次，确认模型路由、Docker、artifact 和 Evaluator 都正常，再扩大批次。

## 两种“复现”的区别

- 要精确复现本报告中的 V3 分数，运行 [`reproduction/`](../reproduction/) 中的脚本。它使用当时保存的 368 份答卷，不会重新调用模型。
- 要复现完整 Benchmark 流程，使用本页的 Harbor 入口。它会生成新答卷；模型输出有随机性，新分数不会逐份等于历史分数。

## 当前配置固定了什么

| 系统 | Agent | 模型 | Agent 版本 | 其他设置 |
|---|---|---|---|---|
| Codex GPT-5.6 Sol | `codex` | `gpt-5.6-sol` | `0.151.0` | high reasoning；关闭 web search 和 websocket transport |
| Claude Opus 5 | `claude-code` | `claude-opus-5` | `2.1.251` | 禁用 WebSearch |
| Qwen 3.8 Max | `qwen-coder` | `openai/qwen3.8-max` | `0.22.3` | OpenAI-compatible route |

并发固定为 1，自动重试固定为 0。这样运行失败不会被静默替换，结果表中的每一份答卷都对应一次明确运行。

## 两个需要单独处理的边界

1. 原生透视表题可能返回 `NATIVE_RECALC_REQUIRED`。Linux 环境不能替代 Microsoft Excel 对原生 PivotTable、PivotCache 和 PivotChart 的最终重算与确认。
2. 公开仓库中的政策情景题已经修正来源、基线和单位。重新调用 Agent 得到的是修订题面的新批次，不能与旧版政策题答卷混为同一组结果。
