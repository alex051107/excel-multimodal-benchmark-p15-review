---
id: R15-A-PROJECT-FINANCE-001
track: A
status: DESIGN_FROZEN
package_not_built: true
harbor_package_built: false
primary_eligible: false
split_group: R15-A-PROJECT-FINANCE-001
---

# Construction-to-operations project finance model

## Professional work and user scenario

A project-finance analyst must turn a wastewater-reuse retrofit budget, funding term sheet, draw rules, and first-year operating forecast into a monthly investment-committee model. The professional work is an equity-first funding waterfall with eligible-cost rules, retainage, interest during construction, commitment fees, the COD transition, scheduled amortization, and post-COD debt-service coverage.

## Instruction draft

“Build the committee model in the supplied workbook. Reconcile approved sources and uses, fund each construction month under the stated equity/debt waterfall, capitalize construction-period interest and commitment fees using the convention file, and switch to scheduled principal after COD. Calculate monthly ending cash, debt balance, loan-to-cost, and post-COD DSCR. Flag any month that breaches minimum cash, debt commitment, LTC, or DSCR limits. Preserve source work-package and period keys so every modeled cost and funding row is traceable. Do not replace formulas with pasted summary values.”

## Packaged-input design and authenticity boundary

- Planned inputs: construction_budget.csv, funding_terms.csv, operating_forecast.csv, milestones.csv, model_conventions.md, and starting_workbook.xlsx.
- The budget will contain 24 monthly periods, work-package IDs, eligible/ineligible flags, cash costs, commitments, and retainage.
- Funding terms will state equity cap, debt commitment, rates, fee capitalization, amortization start, minimum cash, maximum LTC, and minimum DSCR.
- The visible convention file will define monthly rates, the mid-period draw convention, retainage release, signs, and the DSCR numerator and denominator.
- All names, values, and terms will be authored synthetic fixtures. They will not be represented as a real borrower, lender, approval, or transaction.
- No input file has been generated at DESIGN_FROZEN status.

## Target workbook and professional operations

Required sheets are Assumptions, Sources_Uses, Construction_Draws, Debt_Schedule, Operating_Case, Covenants, and Checks. The workbook must map every cost once, apply equity before debt, cap draws at commitment, calculate construction interest and commitment fees on the visible bases, release retainage only at the stated milestone, switch at COD, and reconcile monthly cash, debt, LTC, and DSCR with work-package/month traceability.

## Independent oracle design

A task-local Python oracle will use Decimal arithmetic and the visible convention file. It will independently build the cost ledger, funding waterfall, debt roll-forward, interest/fee schedule, operating cash flow, and covenant series from source rows. It will emit row-keyed facts rather than cell addresses and reject duplicate keys, funding beyond commitment, unbalanced sources and uses, or a broken debt roll-forward.

## Atomic rubric design

- R001: Required sheets exist and the workbook opens as native Excel.
- R002: Every source cost row appears once under the correct work-package and month.
- R003: Eligible, equity-only, and excluded costs follow the visible rules.
- R004: Monthly equity contributions and debt draws implement the equity-first waterfall and debt cap.
- R005: Retainage is withheld and released in the correct periods without double counting.
- R006: Construction interest and commitment fees are formula-linked to the stated rate bases.
- R007: COD, post-COD interest, and scheduled principal begin in the correct periods.
- R008: Sources and uses, debt roll-forward, and ending cash reconcile monthly.
- R009: LTC and DSCR are formula-linked to modeled balances and operating cash flow.
- R010: Breach flags use supplied thresholds and identify the correct periods.
- R011: Work-package/month provenance is complete and unique.
- P001: Penalize reversed funding priority, hidden plugs, stale debt totals, concealed circularity, or wrong covenant definitions.

## Acceptable materially different equivalent

One solution may use a normalized period/work-package ledger feeding monthly summaries with structured aggregation. A materially different equivalent may use separate construction and operating blocks with explicit row waterfalls and a COD bridge. Sheet order, helper columns, and named ranges may differ; funding semantics, debt balances, covenants, and traceability must agree.

## Negative fixtures

- No-op: formatted workbook with inputs present and model outputs blank.
- Malformed: non-XLSX payload or missing required calculation sheets.
- M1: Debt is drawn before available equity is consumed.
- M2: Interest is calculated on closing debt instead of the visible mid-period convention.
- M3: Retainage is counted in the original cost and again at release.
- M4: Commitment fee is charged on drawn rather than undrawn commitment.
- M5: COD occurs one month early, shifting capitalization and principal.
- M6: DSCR divides revenue by debt service instead of using the defined cash-flow numerator.

## Two meaningful input perturbations

- P1: Increase one debt-eligible EPC cash-cost row in a debt-funded month by USD 100,000. The affected use, debt draw, interest, closing debt, remaining commitment, sources-and-uses total, and LTC must propagate; equity cap, unrelated work packages, revenue, and O&M remain invariant.
- P2: Change the annual senior-debt rate from 6.50% to 7.00%. Construction interest, capitalized debt, post-COD cash interest, ending cash, and DSCR must change; source costs, equity contributions, principal draws, revenue, and operating expenses remain invariant.

## Verifier-private CONFIRM sibling

The verifier-private CONFIRM sibling will use a different infrastructure asset, 18 periods, different work-package IDs, COD month, funding caps, rates, retainage pattern, and operating values. It preserves the visible waterfall, roll-forward, and covenant contract. Confirm truth, oracle, perturbations, and reference stay outside Agent-visible input.

## Difficulty-change policy

Allowed difficulty changes: legitimate funding tranches, a visible rate step, more work packages, a reserve-account rule, or a longer post-COD horizon.

Forbidden difficulty changes: hidden conventions, undisclosed priority, manual-iteration circularity, missing keys, arbitrary bloat, irrelevant sheets, stale caches, throttling, or model-specific traps.

## Activation rule

Activate only if a Track A PRIMARY is invalid or remains VALID_BUT_EASY after three allowed professional-complexity revisions, and only after full package construction and every reference, equivalent, negative-fixture, two-perturbation, CONFIRM, Excel, Harbor, and target-Agent gate passes.

## Build state

**package_not_built: true — no task directory, workbook, fixture, Judge, or Harbor package exists.**
