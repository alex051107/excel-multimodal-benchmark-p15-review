# Judge V3 contract

## What the task explicitly asks for

- Preserve the ten paired observations and calculate Group 2 minus Group 1 for each subject.
- Report descriptive statistics for the paired differences.
- Perform a two-sided paired t-test, report a 95% t confidence interval, and apply the stated 0.05 decision rule.
- Report the stated normal-approximation planning result for standardized paired effect 0.8 and 80% power.
- Include one chart linked to the paired-difference calculations.

## What changes the research conclusion

The Judge checks subject identity and order, subject pairing, difference direction, mean and dispersion, the paired rather than independent test, the two-sided p-value, confidence interval, planning result, decision consistency, source preservation, and chart linkage. The split contract freezes the ten expected subject IDs; blank, changed, reordered, or duplicate IDs fail the source-pair hurdle. The paired/two-sided method and unchanged raw observations are critical.

The decision must be singular, and decision consistency (R010) is a delivery hurdle. A workbook that contains both a rejection/significance conclusion and a fail-to-reject/non-significance conclusion fails it.

The instruction does not require six physical sheets, inclusion flags, a hidden QC-exclusion experiment, fixed rows, or formula-linked inferential summaries. Those are not scored requirements.

## Accepted layouts and equivalent implementations

- One combined analysis sheet, two or three sensibly divided sheets, or the reference multi-sheet layout.
- Renamed sheets and relocated tables when the paired data, statistics, decision, planning result, and chart remain unambiguous.
- `Group`, `Treatment`, or similarly clear paired-column labels.
- Formula-linked or correctly reported statistical results. The chart must reference the paired-difference cells; formula-linked difference cells receive the dynamic-link credit.
- Equivalent two-sided paired-test formulas, including `T.DIST.2T`, a two-sided `T.TEST` on paired ranges, or a numerically correct result accompanied by an explicit paired/two-sided method statement.

Hard-coded but correct analysis is scored as analysis content, while losing the formula-link criterion instead of becoming a Judge error.
