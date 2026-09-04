# Judge V3 contract

## What the task explicitly asks for

- Diagnose one benchmark-controlled formula defect in the supplied workbook.
- Make the smallest necessary formula repair.
- Preserve assumptions, formulas, and unrelated workbook state.
- Leave the revenue, working-capital, debt, and summary schedules tied and usable.

## What changes the FP&A result

The Judge checks the corrected 2027 revenue, downstream receivables and free cash flow, debt-service coverage, formula linkage, assumption preservation, and repair locality. The root-cause formula and protected assumptions are critical.

The declared price perturbation check (R006) is also a delivery hurdle: the repaired revenue chain must respond to the price assumption, so a baseline-correct constant formula cannot pass.

## Accepted layouts and equivalent implementations

This is an in-place repair task, not a blank-workbook construction task. The supplied sheet and cell structure is therefore part of the evidence: a valid answer may use any algebraically equivalent formula in the defective cell but may not rebuild or relocate the model. Formatting-only differences are ignored.

## Calibration status

All current completed workbooks made the same correct one-cell repair. Treat this as a saturated control task for pipeline and repair-locality validation, not as evidence that separates model capability.
