# 15 道题的 Rubric 和 Evaluator 导航

审查一道题时，先读 `rubric.json`，再看 `tests/evaluate.py`。Rubric 说明业务上要检查什么，Evaluator 说明程序实际怎样查。两者不一致时，应该修改 Evaluator 或 Rubric，不应该用某个系统原来的分数决定哪一方正确。

## `rubric.json` 怎样读

| 字段 | 含义 |
|---|---|
| `review_notes.purpose_zh` | 这道题真正要检查的业务结果 |
| `review_notes.answer_location_zh` | Evaluator 怎样在工作簿里找到答案 |
| `review_notes.accepted_equivalents_zh` | 不应该被扣分的合理替代写法或布局 |
| `review_notes.not_scoreable_zh` | 哪些情况应该暂时不算分，不能直接写成 0 分 |
| `review_notes.known_limit_zh` | 当前已知的识别或验证缺口 |
| `criteria` | 实际计分项；`description` 是判断内容，`weight` 是权重 |
| `dimension` | 评分项所属的业务方面，例如对账、来源或记录完整性 |
| `method` 和 `method_params` | 程序用什么方式检查，以及使用哪个独立重算或容差 |
| `hurdle_criteria` 和 `pass_threshold` | 旧运行接口保留的二元判定字段；当前对外分析不用它们把连续分数切成两类 |

## 逐题入口和当前审查重点

| 题目 | 评分文件 | 主要检查 | 当前审查重点 |
|---|---|---|---|
| A 工程选泵 | [Rubric](../tasks/pilot_v1/P15-A-ENG-SIZING-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-A-ENG-SIZING-001/tests/evaluate.py) | 单位、水力计算、安全余量、首台合格设备 | 加入标签和数值相隔较远的布局 |
| A 现金流估值 | [Rubric](../tasks/pilot_v1/P15-A-FIN-DCF-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-A-FIN-DCF-001/tests/evaluate.py) | 预测、自由现金流、终值、股权价值和敏感性 | 用真实答卷核对标签别名和敏感性表 |
| A 财务纠错 | [Rubric](../tasks/pilot_v1/P15-A-FIN-DEBUG-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-A-FIN-DEBUG-001/tests/evaluate.py) | 最小修复、后续对账、未受影响区域 | 作为简单对照题，不单独用来判断系统上限 |
| A 政策情景 | [Rubric](../tasks/pilot_v1/P15-A-POLICY-EIA-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-A-POLICY-EIA-001/tests/evaluate.py) | 发电平衡、排放量、排放强度和单位 | 旧答卷不可比；修正版重跑前不发布分数 |
| A 配对实验 | [Rubric](../tasks/pilot_v1/P15-A-STAT-EXPERIMENT-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-A-STAT-EXPERIMENT-001/tests/evaluate.py) | 配对身份、统计量、区间、样本量和图 | 处理仍然暂时算不出分的跨表布局 |
| B 财务对账 | [Rubric](../tasks/pilot_v1/P15-B-FIN-RECON-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-B-FIN-RECON-001/tests/evaluate.py) | 汇率、匹配、异常、批准调整和差异桥 | 逐份确认当前低分是真实错误还是布局漏识别 |
| B 健康数据报告 | [Rubric](../tasks/pilot_v1/P15-B-HEALTH-REPORT-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-B-HEALTH-REPORT-001/tests/evaluate.py) | 地区映射、期间指标、比较、图表和来源 | 加入同一张表放多个业务区块的布局 |
| B 订单清洗合并 | [Rubric](../tasks/pilot_v1/P15-B-OPS-CLEAN-JOIN-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-B-OPS-CLEAN-JOIN-001/tests/evaluate.py) | 清洗、主数据连接、去重、异常和总数 | 修正连续分数与旧关键项判定的冲突 |
| B 原生透视表 | [Rubric](../tasks/pilot_v1/P15-B-PUBLIC-PIVOT-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-B-PUBLIC-PIVOT-001/tests/evaluate.py) | 原生 PivotTable、缓存、过滤、求和、图表和 KPI | 补齐数据连接与 Microsoft Excel 刷新证据 |
| B 销售数据选择 | [Rubric](../tasks/pilot_v1/P15-B-SALES-DISCOVERY-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-B-SALES-DISCOVERY-001/tests/evaluate.py) | 批准版本、时间覆盖、KPI、地区和登记表 | 防止从多张相互矛盾的表中拼分 |
| C 发票整理 | [Rubric](../tasks/pilot_v1/P15-C-INVOICE-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-C-INVOICE-001/tests/evaluate.py) | 发票身份、明细、页码、金额链和来源 | 补测页码、金额联动和重复表格变体 |
| C 采购变更单 | [Rubric](../tasks/pilot_v1/P15-C-PO-ADDENDUM-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-C-PO-ADDENDUM-001/tests/evaluate.py) | 三项变更、原行保留、修订来源和总额 | 核对总额的替代位置与公式要求 |
| C 报价单整理 | [Rubric](../tasks/pilot_v1/P15-C-QUOTE-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-C-QUOTE-001/tests/evaluate.py) | 基础范围、可选项、折扣、税和来源 | 继续补充可选项与来源表的排列变体 |
| C 多张票据整理 | [Rubric](../tasks/pilot_v1/P15-C-RECEIPTS-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-C-RECEIPTS-001/tests/evaluate.py) | 票据身份、明细、分类、总额和批次对账 | 确认检查的是答卷实际展示的对账结果 |
| C 银行账单整理 | [Rubric](../tasks/pilot_v1/P15-C-STATEMENT-001/rubric.json) · [Evaluator](../tasks/pilot_v1/P15-C-STATEMENT-001/tests/evaluate.py) | 交易、借贷方向、分类、页码和开闭额对账 | 补测分类汇总、开始行和金额联动的替代布局 |

## 审查时建议对照四类文件

1. `instruction.md` 是 Agent 真正看到的要求。Rubric 不能扣题面没有要求的事。
2. `metadata/oracle_recompute.py` 从输入和业务规则独立重算。它不应该直接拷贝参考工作簿。
3. `fixtures/equivalent/` 应该接受布局不同但业务结果正确的答案。
4. `fixtures/mutants/` 应该覆盖会改变使用者判断的真实错误。

当前 15 份 Evaluator 共检查了 168 个测试文件，程序运行结果与这些测试的预期一致。这能说明代码按当前设定运行，还不能证明每条设定都合理。最后的判断仍然需要抽取真实高分、中间分、低分和暂时无法评分的工作簿做人工复核。
