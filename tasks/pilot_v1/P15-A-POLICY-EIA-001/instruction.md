# Electricity policy scenario model

The policy unit needs the missing scenario calculation block restored before it can review an emissions-reduction proposal. The package contains benchmark-adapted generation inputs, a policy-assumption sheet, and a workbook whose derived scenario cells have been cleared.

Rebuild the formula-linked base and policy cases. Use the supplied historical generation as the base case. Treat the difference between demand and the four listed generation categories as `Other generation`, hold that residual constant in the policy case, and let natural gas balance the remaining policy demand after coal displacement and wind/solar uplift. Calculate emissions in metric tonnes using `GWh × 1,000 × tCO2/MWh`, then calculate intensity in tCO2/MWh. Complete the decision results without changing the supplied inputs or policy assumptions.

The supplied generation values are rounded/adapted benchmark inputs informed by public EIA electricity-generation tables; they are not a verbatim EIA extract. The 4,000,000 GWh demand value is a benchmark scenario assumption. The model is an ordinary formula-linked workbook and does not require VBA. Deliver `/app/output/answer.xlsx`.
