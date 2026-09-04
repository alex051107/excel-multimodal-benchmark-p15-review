# Judge V3 contract

## Business result that must be correct

- Preserve all eight corrected Q2 events in an Excel Table.
- Keep a genuine refreshable PivotCache and PivotTable with Region in rows, Program in columns, Quarter filtered to `2024Q2`, and SUM measures for Participants and Spend.
- Keep linked KPI results and a clustered-column PivotChart bound to the native PivotTable.

## Explicit dynamic and native requirements retained

This task explicitly requires native Excel objects, refresh, Microsoft Excel recalculation, Pivot-linked KPIs, and a native PivotChart. Those requirements remain hard acceptance conditions.

## Preferences removed from scoring

- The internal Excel Table name, PivotTable name, sheet names, fixed field indexes, exact source address, KPI cell addresses, and Pivot anchor cell.
- Dev-only `North / Outreach` wording in the confirm split; the confirm focus region and program come from its own oracle.

## Native status boundary and examples

- `NATIVE_OBJECT_CHECKED`: the final receipt-bound canonical workbook contains valid native objects and verified Excel-cached KPIs, so it receives a numeric score.
- `NATIVE_RECALC_REQUIRED`: the native structure and formulas are valid but cached KPI results are absent or stale; no numeric model score is issued until Excel recalculates and saves it.
- Positive: a package with renamed internal objects (`Q2EventsInput`, `Q2DeliveryAnalysis`) still scores 1.0.
- Negative: `wrong_aggregation.xlsx`, `wrong_chart_binding.xlsx`, and `stale_source_range.xlsx` still fail their explicit business/native requirements.
