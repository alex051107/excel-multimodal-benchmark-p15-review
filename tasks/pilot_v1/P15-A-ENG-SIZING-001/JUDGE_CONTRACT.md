# Judge V3 contract

## What the task explicitly asks for

- Preserve the supplied design inputs and approved three-pump catalog.
- Convert flow and pipe diameter to SI units.
- Calculate area, velocity, velocity head, friction head, total dynamic head, hydraulic power, shaft power, and the safety-adjusted minimum flow and motor rating.
- Mark each pump against flow, head, and motor constraints and recommend the first eligible pump in catalog order.
- Keep units visible and keep the calculation and selection formula-linked to the design inputs.

## What changes the engineering decision

The Judge checks the numerical hydraulic chain, the two safety-adjusted requirements, all three eligibility constraints, the selected pump, source preservation, and recalculation when flow or diameter changes. Those checks determine whether the sizing conclusion remains usable.

Because the task explicitly requires calculations and selection to remain linked to the design inputs, the design-flow recalculation check (R008) is a delivery hurdle. A formula that merely returns the baseline constant cannot pass it.

A separate `Checks` sheet, a particular row number, or the reference workbook's sheet names do not change the decision and are not required.

## Accepted layouts and equivalent implementations

- The reference five-sheet layout.
- One combined sizing sheet.
- Several sensibly named sheets, for example `Sizing`, `Catalog`, and `Selection`.
- Relocated rows and columns when labels, units, formulas, catalog values, eligibility results, and the selected pump remain unambiguous.
- Algebraically equivalent Excel formulas and either one combined eligibility result or three component checks.
- Common professional labels such as `Required motor rating`, `Motor rating (with safety)`, `Cross-sectional area`, and `Pipe cross-section area`.

Ambiguous duplicate metrics remain a Judge limitation. A recognizable but wrong or hard-coded calculation receives criterion failures rather than a layout error.
