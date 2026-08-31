# Native PivotTable confirm sibling blocked on Windows Excel

Status: `TASK_INVALID`. Blocker: `PENDING_EXTERNAL_WINDOWS_EXCEL`.

Run `tests/confirm/windows/build_and_validate_confirm.ps1` on a Windows host with desktop Microsoft Excel. Do not create or claim a confirm reference until the native Table, PivotCache, PivotTable, filter, SUM measures, refresh configuration, cached oracle values, and PivotChart pass both COM reopen validation and the confirm Judge split.
