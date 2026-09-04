# Judge V3 contract

## What the task explicitly asks for

- Use the existing `Source_Data`, `Assumptions`, `Forecast`, `Valuation`, and `Checks` workbook structure.
- Preserve the historical source facts.
- Produce the FY2025E-FY2029E revenue, operating, tax, D&A, capex, working-capital, and unlevered free-cash-flow forecast.
- Apply the supplied WACC and Gordon-growth terminal value, bridge enterprise value to equity value, and complete the stated sensitivity table.
- Keep calculated outputs formula-linked to assumptions.

## What changes the valuation decision

The Judge checks the full forecast, terminal value, discounting, enterprise-to-equity bridge, sensitivity coordinates, protected history, and response to growth and WACC changes. Enterprise value, equity value, and sensitivity conclusions are critical.

The growth and WACC response checks (R009 and R010) are delivery hurdles because the instruction explicitly requires calculated outputs to remain formula-linked to assumptions. Baseline-correct constant formulas do not satisfy this contract.

## Accepted layouts and equivalent implementations

The task explicitly requires the supplied workbook structure, so the five sheet roles are part of the contract. Formula text is not matched to the reference: algebraically equivalent forecasting, discounting, terminal-value, and sensitivity formulas are accepted. Formatting, style, comments, and unrelated empty cells are ignored.

Judge V3 intentionally keeps the V2 semantic calculations because the full current workbook set was scoreable and no layout false negative was found.
