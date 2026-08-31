---
name: excel-benchmark-track-a-construction
description: Construct or repair formula-driven Track A Excel benchmark tasks covering DCF, model debugging, engineering sizing, paired statistics, and policy scenarios with auditable semantics.
---

> Validated against P15 checkpoint: `8df20f98366633a4813cf6b98ee7e78270e3bdb7`
> Last evidence refresh: `2026-08-30`

# Track A Construction

## Purpose

Build a professional formula-driven workbook task whose difficulty comes from authentic dependencies, recalculation, and decision logic. Preserve source inputs, make derived work auditable, and keep the instruction sufficient without revealing the answer or defect location.

Keep construction instructions and verifier truth repository-side. Do not place this skill in Agent-visible task inputs.

## Required inputs

- A task-specific professional brief, authoritative or benchmark-authored source data, units and conventions, and the required decision output.
- The intended workbook sheets, allowed edit region, source cells that must remain unchanged, and any chart/native-object requirement.
- An independent DEV oracle specification, a distinct CONFIRM sibling plan, and the rubric/critical-failure policy.
- The installed spreadsheet runtime with `@oai/artifact-tool` for workbook authoring, import/export, inspection, and rendering.

## Exact procedure

1. Freeze the semantic contract before authoring. Use the relevant P15 pattern below; do not combine patterns merely to add complexity.

| Pattern | Required professional chain | Two meaningful DEV perturbations |
|---|---|---|
| DCF | Revenue, EBIT, tax, NOPAT, D&A, capex, NWC, UFCF, discount/PV, terminal value, enterprise-to-equity bridge, and every sensitivity coordinate | Growth and WACC |
| DEBUG | Begin from a valid integrated model, inject one hidden root cause into the Agent input, require the smallest formula repair, and check downstream ties plus collateral locality | Price and DSO |
| ENG | Convert to SI, calculate hydraulic/velocity/friction/total head and power, evaluate flow/head/motor constraints against a complete catalog, and return the first eligible item | Design flow and internal diameter |
| STAT | Preserve paired rows, compute paired differences, two-sided paired t inference, 95% CI, `ceil(((1.96+0.84)/0.8)^2)=13`, decision text, and a chart bound to the paired calculations | One observation and one QC exclusion |
| EIA | Link generation sources, demand growth, coal displacement, renewable uplift, residual gas balance, emissions, intensity, result, and checks | Coal displacement and demand growth |

2. Author or edit the workbook with `@oai/artifact-tool`. Inspect formulas/values and render every relevant sheet before and after material edits. Do not use another library as a silent authoring substitute.
3. Keep raw/source/assumption cells typed and visible. Put derived work in formulas; do not hardcode critical outputs or cached answers.
4. Write a realistic instruction that states the professional outcome, supplied inputs, editable expectations, and output path without disclosing oracle values, CONFIRM truth, or a DEBUG defect location.
5. Split the rubric into 8–12 atomic criteria. Each high-weight criterion maps to a distinct implemented Judge check; critical semantic or source-integrity failures activate a justified veto/cap.
6. Add exactly two distinct DEV input perturbations with exact expected propagation values and protected invariants. Verify formula errors separately from unsupported-function or non-Windows recalculation limitations.
7. Create reference, materially different acceptable equivalent, no-op, malformed, and at least four task-specific semantic mutants. Mutants should exercise different professional failures and must not all collapse to one score.
8. Create a verifier-private CONFIRM sibling with a different professional/source instance and changed values while preserving the capability invariant. Keep its inputs, oracle, reference, and contract under `tests/confirm`; never revise CONFIRM in response to model failures.
9. Make the evaluator selectable by explicit `--split dev|confirm`. Account for Harbor's `tests/` Docker build context: runtime verifier assets needed inside `/tests` must exist under `tests/`, while the Agent environment receives only task inputs.

## Required outputs

- `instruction.md`, Agent-visible inputs, formula-linked `solution/reference.xlsx`, `rubric.json`, `task.toml`, and task metadata.
- A deterministic evaluator, atomic private contract, independent DEV oracle, and two exact perturbation contracts.
- A materially different acceptable equivalent, no-op, malformed workbook, and at least four task-specific mutants.
- A distinct `tests/confirm` sibling with contract, input overlay, oracle, reference, and explicit evaluator invocation.
- Local validation evidence with evidence boundaries; no unearned Windows, Harbor, agent-hardness, or human status.

## Validation checks

- Reference `=1.0`; acceptable equivalent `>=0.95`; no-op `<0.30`; malformed `=0`; every semantic mutant `<0.70` and mutant scores are not all identical.
- Repeat each deterministic fixture at least five times; require identical results. Any nondeterministic Judge component must have score standard deviation `<0.05`.
- Both DEV perturbations exactly match oracle propagation and preserve declared source/assumption invariants.
- Formula cells, chart bindings, selection logic, sensitivity coordinates, and critical-veto behavior are checked independently.
- DEV and CONFIRM use different values/source instances; CONFIRM truth is absent from `data/input_files` and `environment/input`.
- Harbor smoke proves package/Judge execution only. Windows Excel open/save/recalculate and native-object readback require a real Windows Excel receipt.

## Allowed difficulty changes

- DCF: add legitimate forecast drivers, capital-intensity logic, bridge components, or disclosed sensitivity dimensions.
- DEBUG: deepen the dependency chain while retaining one diagnosable root cause and a minimal repair.
- ENG: add a disclosed engineering constraint, unit conversion, or complete catalog attribute used in eligibility.
- STAT: add a disclosed QC rule, paired-data issue, or decision-relevant analysis that preserves paired semantics.
- EIA: add a disclosed scenario lever, balancing rule, or emissions check supported by supplied source data.
- After a real target-agent run, make at most three such revisions; otherwise use the designated reserve or record `VALID_BUT_EASY`.

## Forbidden difficulty changes

- Hidden conventions, missing keys/units, ambiguous authoritative rows, stale caches, broken file access, irrelevant volume, or decorative noise.
- Pasted critical values, intentionally opaque formulas, manual-iteration traps, token/time throttling, or target-model-specific exploits.
- DEBUG instructions that reveal the root cause, or repairs that alter assumptions/unrelated formulas.
- Independent-groups analysis for paired data, partial equipment catalogs, unbound charts, or incomplete DCF period/sensitivity checks.
- Copying this skill, oracle data, CONFIRM values, or reviewer notes into Agent-visible inputs.

## Common failure modes

- A polished workbook omits per-period DCF bridges or checks only the final valuation.
- DEBUG is solved by hardcoding a summary or causes collateral edits outside the root formula.
- ENG selection uses an incomplete visible catalog or tests only one of flow/head/motor eligibility.
- STAT uses an independent-samples test, rounds paired power planning down, or charts static cells.
- EIA carries a retired identity/source claim or breaks generation balance while totals still look plausible.
- An identity-copy workbook is labeled an acceptable equivalent; passing the same Judge is not material difference.
- Artifact-tool inspection or cached values are reported as Windows Excel evidence.

## Stop conditions

- Stop if required facts, units, mappings, or authoritative source values do not determine a reviewable result.
- Stop authoring if `@oai/artifact-tool` is unavailable; record the artifact blocker instead of switching libraries silently.
- Stop local claims at `PENDING_EXTERNAL_WINDOWS_EXCEL` when canonical Excel evidence is absent.
- Stop Harbor claims when the tests-context image cannot resolve its private runtime assets.
- Stop if CONFIRM has leaked, a semantic mutant cannot be independently specified, or difficulty requires artificial noise.

## Examples from verified P15 work

- The five Track A local receipts each recorded five stable evaluator repetitions, reference and equivalent scores of `1.0`, malformed `0`, and task-specific mutants below `0.70`; all still recorded canonical Excel as pending external Windows Excel.
- DCF local checks separated the period forecast, operating bridge, FCF bridge, discount/PV, terminal value, enterprise/equity bridge, sensitivity grid, two perturbations, and source protection.
- DEBUG checked a hidden root cause, formula locality, downstream recalculation, and collateral damage; ENG checked SI conversion, hydraulic calculations, complete catalog eligibility, and first-eligible selection.
- STAT's inspected contract uses paired inference and planned `n=13`; EIA checks generation balance, emissions/intensity, two scenario perturbations, and historical-source protection.
- All five inspected CONFIRM contracts set `dev_values_reused=false`, name changed drivers, and keep private values under `tests/confirm`.
