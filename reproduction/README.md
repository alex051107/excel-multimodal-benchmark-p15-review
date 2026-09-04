# 从公开仓库复算 V3 结果

这个目录解决一个具体问题：审查者能否只使用 GitHub 中公开的答卷、Evaluator 和运行脚本，重新得到报告采用的 V3 分数。

## 公开了什么

| 文件 | 用途 |
|---|---|
| [`data/answer_workbooks.zip`](data/answer_workbooks.zip) | 368 份实际答卷。只保留 Excel 文件，不包含模型原始回复、账号、费用或本机运行目录 |
| [`data/workbook_manifest.csv`](data/workbook_manifest.csv) | 每份答卷对应的题目、系统、归档位置、预期状态和预期分数 |
| [`data/bundle_manifest.json`](data/bundle_manifest.json) | 答卷数量、压缩包校验值和公开前的清理记录 |
| [`../scripts/validate_all_evaluators.py`](../scripts/validate_all_evaluators.py) | 在临时目录逐题运行 15 份 Evaluator 回归测试，不改动仓库中的验证回执 |
| [`../scripts/reproduce_v3_results.py`](../scripts/reproduce_v3_results.py) | 解压答卷、调用每道题的 V3 Evaluator、逐行比较结果 |
| [`../requirements-reproduction.txt`](../requirements-reproduction.txt) | 本次复算使用的 Python 依赖版本 |

15 道题各自的题面、输入、`rubric.json`、`tests/evaluate.py`、参考工作簿和测试样例位于 [`tasks/pilot_v1/`](../tasks/pilot_v1/)。

## 一次完成复算

在仓库根目录运行：

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements-reproduction.txt
python scripts/validate_all_evaluators.py
python scripts/reproduce_v3_results.py --workers 4
```

Windows PowerShell 使用：

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements-reproduction.txt
python scripts/validate_all_evaluators.py
python scripts/reproduce_v3_results.py --workers 4
```

第一个检查成功时会显示 `"passed": 15` 和 `"all_evaluators_ok": true`；复算成功时会显示 `"reproduction_ok": true`。生成的逐份结果和汇总分别位于：

- `reproduction/output/replayed_scores.csv`
- `reproduction/output/reproduction_summary.json`

GitHub 也会通过 [Reproduce V3 scores](../.github/workflows/reproduce-v3.yml) 先运行 15 份 Evaluator 回归测试，再复算 368 份答卷。这两个检查都不读取密钥，不调用在线模型。

脚本会检查：

1. 答卷压缩包与公开清单一致；
2. 368 份答卷都能进入相应的 Evaluator；
3. 每份答卷的状态和分数与公开清单一致；
4. 当前报告使用的 286 份结果逐行一致。

其中 24 份旧版政策题不会被修订后的 Evaluator 追溯打分。复算脚本会按照本轮审查合同直接将它们记为 `TASK_INVALID`，原因是旧题面的来源、基线和单位定义互相冲突；强行生成一个数字反而无法比较。

## 数量为什么不是同一个

- 368 是归档中可供重评的 Excel 文件总数。
- 322 是 V3 当前能给出连续分数的文件数。
- 286 是当前计划批次中正常完成、且被报告逐题平均分采用的结果数。
- 其余文件保留为“需要 Excel 原生重算”“评分程序暂时无法判断”或“旧版政策题不可比较”等明确状态，不会改写成 0 分。

## 公开前做过什么清理

18 份原生透视表工作簿带有 Microsoft Excel 自动写入的本机绝对目录。公开包只把该目录字段改成 `./`，没有改动工作簿中的单元格、公式、透视表或缓存内容。`bundle_manifest.json` 记录了这次清理，复算脚本会重新验证清理后的公开包。

## 能复现到哪一步

这套材料可以复现“固定答卷经过当前 V3 Evaluator 后得到什么分数”，也是本次报告数字的直接来源。题目包中的 `task.toml`、`environment/` 和 `tests/` 也保留了 Harbor 执行接口，可用于重新生成一批新答卷。

重新调用在线模型不会逐字生成同一份 Excel：模型版本、服务端状态和采样本身都会变化。因此，新一轮 Agent 运行应被视为新的实验批次，不能拿来代替这次固定答卷的分数复算。
