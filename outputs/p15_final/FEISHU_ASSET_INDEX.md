# 汇报配图索引

| 正文位置 | 图片 | 这张图回答的问题 | 对外数据来源 |
|---|---|---|---|
| 第三部分 | [图 1 旧版逐题结果](feishu_report_assets/03_old_task_results.png) | 为什么旧评分规则必须重审 | [`old_scores_by_task.csv`](../../results/judge_v3_initial/old_scores_by_task.csv) |
| 第三部分 | [图 2 旧 0 分重算](feishu_report_assets/04_old_zero_recheck.png) | 98 份旧 0 分在新版规则下变成了什么 | [`old_zero_recheck_summary.csv`](../../results/judge_v3_initial/old_zero_recheck_summary.csv) |
| 第五部分 | [图 3 新版逐题结果](feishu_report_assets/06_v3_task_results.png) | 各题的平均分、结果数和当前评分风险 | [`v3_scores_by_task_system.csv`](../../results/judge_v3_initial/v3_scores_by_task_system.csv) |
| 第五部分 | [图 4 题目与系统热力图](feishu_report_assets/07_v3_task_system_heatmap.png) | 同一道题在三套系统中的平均分和结果数有何差异 | [`v3_scores_by_task_system.csv`](../../results/judge_v3_initial/v3_scores_by_task_system.csv) |

所有图都用“份”表示 Excel 结果数量。热力图中的 `n` 是实际进入平均分的结果数，不一定等于原计划的 8 次。政策情景没有新版结果，在图中显示“新版题目还没重跑”，没有画成 0 分。热力图用深色表示更低的平均分，同时在每个格子中写出平均分和结果数。
