# Judge V3 contract

## Business result that must be correct

- Align Vermont and New Hampshire to 2012–2017 and compare 2012–2014 with 2015–2017.
- Report correct baseline, current, absolute-change, and percent-change values for both states.
- Make the numerical briefing claims traceable to workbook evidence and show the correct two-state chart.
- Preserve both raw extracts and the geography map.

## Explicit implementation requirements retained

Traceability, a report, a chart, and checks are explicit. The Judge therefore requires correct supported claims, correct chart values, and reviewable checks. Traceability can be shown by formulas or by a clearly identified supporting data/metrics table. A chart may point directly to the period-metrics table or to a separate chart-data table; it is judged by the two state/value pairs, not by co-location with the chart object.

## Preferences removed from scoring

- Underscore-only sheet names, fixed cell addresses, and the reference sheet order.
- Hidden edits to two source cells as a condition for credit.
- Formula-only report claims, formula-only chart source cells, and an unstated mandatory “highest state” sentence.
- A workbook-wide “no formulas anywhere” condition for accepting a locally static, correct metrics table. Unrelated formulas elsewhere do not change those displayed metrics.
- More precision than the workbook displays; correct one-decimal rates and percentages are accepted.

## New positive and negative examples

- Positive: `Lg97ghs` uses plain-language sheet names such as `Validation Checks`, a static but correct briefing, explicit evidence, and a correct chart.
- Negative: `hardcoded_report_claim.xlsx` still fails because the Vermont change is 2.0 instead of 2.33; `stale_chart_data.xlsx` fails because the chart is wrong.
