# Judge V3 Contract

## 题面明确必须结果

- 正确应用数量变更、新增 commissioning 行和 cable-tray 单价变更。
- 保留所有未列出的基础行、原 PO 行 ID 和基础 PO 历史。
- 变更记录与可读溯源信息正确，修订后排程总额正确。

## 保留的动态要求

题面明确要求修订后排程总额与公式连接。至少一个标记清楚的修订总额必须由排程数据计算，修改数量或单价时必须响应。

## V3 删除的非必要惩罚

- 删除 `Revised_Schedule!F8` 和 `PO_Header!B9` 必须同时存在两个总额公式的要求。
- 不再要求固定工作表名；使用可见表头和标签识别业务角色。
- 只要候选工作簿显示了多个修订总额，它们都必须一致响应；不强制创建重复总额。
- 变更类型可用 `Quantity`、`Insert`、`Unit price` 等简洁业务措辞，不要求固定句式。

## 新正反例

- 正例 `single_total_plain_labels`：只保留一个公式总额并使用简洁变更标签，应得 1.0/通过。
- 正例 `renamed_sheets`：全部工作表改名但可见业务表头完整，应得 1.0/通过。
- 反例 `wrong_unit_price`：修订单价与 addendum 不符，必须在 R003/R005 失败。

实际分数与通过状态见 `receipts/judge_v3_local_validation.json`。
