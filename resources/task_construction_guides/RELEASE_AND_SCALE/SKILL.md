---
name: excel-benchmark-release-and-scale
description: Assemble an evidence-bounded Excel benchmark release and blind review package without promoting unverified tasks or external claims.
---

> Validated against P15 checkpoint: `8df20f98366633a4813cf6b98ee7e78270e3bdb7`
> Last evidence refresh: `2026-08-30`

# Purpose

Turn a frozen set of repaired tasks into a reviewable release. Preserve the boundary between local validity, Harbor packaging, Agent screening, canonical Excel evidence, and human acceptance.

# Required inputs

- Frozen PRIMARY manifest and design-only RESERVE manifest.
- Current task packages, local and CONFIRM receipts, Harbor jobs, Agent attempts, Excel receipts or blockers, and source/render audits.
- Frozen release identifier, target-system configuration, budget state, and review requirements.

# Exact procedure

1. Reconcile the filesystem and manifest to the frozen IDs and track counts. Exclude stale task IDs except explicit legacy metadata.
2. Run one final local semantic gate and one final CONFIRM gate on the unchanged Judge state. Run latest-Judge Harbor Oracle/NOP/malformed smoke and index exact job paths.
3. Record canonical Windows Excel evidence only from an actual Windows Excel receipt. Keep native Pivot or other affected tasks invalid while that receipt is absent.
4. Summarize every real target-Agent attempt with score, dimension scores, duration, tokens, cost, exception, environment state, and allowed failure category. Treat auth, quota, and runtime failures separately from capability failures.
5. Build blind bundles from current instruction and public inputs only. Add output path, review form, and hashes; exclude reference, rubric, Judge, oracle, mutants, private truth, skills, and Agent results.
6. Generate criterion coverage, review manifest, checksums, task status, difficulty matrix, evaluation configuration, results summary, review guide, and release notes.
7. Verify the external review archive by listing its contents and checking it against the exclusion policy.
8. Export only the approved release allowlist into a clean repository. Create one reviewable release commit and keep development history outside that repository.

# Required outputs

- Blind bundles with manifests, coverage, checksums, and review instructions.
- Fixed task status, difficulty matrix, evaluation configuration, run results, cost records, and explicit blockers.
- Version-controlled builder skills, exactly nine design-only reserve cards, and a clean release repository or an explicit publication blocker.

# Validation checks

- Exactly 15 PRIMARY tasks: five each for A, B, and C; exactly nine RESERVE cards and no reserve packages.
- Blind-bundle content allowlist and private-asset denylist both pass.
- Checksums are computed once for final external delivery bytes and verify successfully.
- Report counts agree with task status, Harbor results, Agent attempts, and Excel receipts.
- `AGENT_HARD_CANDIDATE` is used only when the formal three-system rule is met; `ACCEPTED_HARD` requires real human evidence.
- The release tree is free of caches, inspection sidecars, credentials, internal discussions, and unrelated historical task roots.

# Allowed difficulty changes

- None inside the final release batch. Return a task to the bounded difficulty loop when a real validity or easy-task finding requires semantic revision.
- A release may downgrade a claim or task status when evidence is missing.

# Forbidden difficulty changes

- Adding documents, wrappers, dashboards, registries, web UIs, or generalized orchestration to compensate for weak task content.
- Hiding invalid tasks, environment failures, unavailable providers, cost, or quota status.
- Including construction skills or verifier-private material in blind or Agent-visible bundles.
- Mixing development history or internal review discussions into the clean release repository.

# Common failure modes

- A stale report or manifest still names a retired task ID.
- Harbor smoke is described as task difficulty.
- A preview archive or API response is treated as proof of external review or publication.
- A blind bundle contains rubric, gold, private truth, or Agent results.
- A draft report calls a task Windows-validated without a COM receipt.

# Stop conditions

Stop claim promotion when Windows Excel, human review, provider credentials, budget, repository access, or an external service is unavailable. Save the exact checkpoint, finish all independent local work, and leave the blocked status explicit.

# Current verified examples

- P15 retained `P15-B-PUBLIC-PIVOT-001` as `TASK_INVALID / PENDING_EXTERNAL_WINDOWS_EXCEL` while shipping a static, fail-closed native builder and receipt manifest.
- Latest-Judge Harbor smokes were indexed separately from difficulty attempts: valid Oracle outputs at 1.0, NOP at 0, malformed at 0, with packaging exceptions reported rather than scored.
- The nine reserve cards remained `DESIGN_FROZEN` and `package_not_built: true`.
- The current release records 85 real Agent attempts, including full one-pass coverage of all 14 valid tasks by Claude Code + Opus 4.8 and Qwen Code + Qwen3.8-max. Their earlier credential blockers are no longer current.

# Historical failure examples

- Authentication or budget failures must be recorded as environment blockers and excluded from capability statistics. Resume sampling only after the same configuration passes a fresh preflight.
